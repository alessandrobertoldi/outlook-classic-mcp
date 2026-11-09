"""
outlook-classic-mcp: a local MCP server that connects Claude (Claude desktop, Cowork) to classic Outlook for Windows.

HOW IT WORKS, IN SHORT
- The server runs on the user's PC and is launched by Claude desktop over stdio: no network port is opened.
- It talks to classic Outlook through the COM object model (pywin32 library), attaching to the Outlook session the
  user has ALREADY opened. It never starts Outlook and never knows a password: it uses the accounts Outlook has.
- Every function exposed to Claude (an MCP "tool") is an `async def` coroutine decorated with `@mcp.tool`. The real
  work happens in a small inner function `f()` that is executed on the COM thread (see ComWorker below).
- Everything that depends on the user (mailbox names, folders the server may write to, sender and recipient of the
  alerts) lives in `config.json`, never in the code. `config.example.json` explains every key.

RULES ENFORCED IN CODE (not in a prompt, so Claude cannot talk its way around them)
- No tool sends mail, except `send_alert`, whose sender and recipient are fixed in the configuration: Claude can only
  choose subject and body, never who receives it.
- No tool deletes permanently. `delete_mail` and `delete_event` move the item to the trash of its own mailbox
  (recoverable), only if `expected_subject` matches the real subject, never from the trash itself, never meetings with
  attendees. Both are disabled unless `"deletion_enabled": true`. The one exception is deleting a single occurrence of a
  recurring series, which Outlook does not route through the trash (the tool description says so).
- No tool deletes contacts, tasks, notes or folders; no tool touches rules, settings or accounts.
- Every unread mail Claude reads (or saves, or drafts a reply to, or moves) is left unread.
- `save_mail` and `save_attachments` write only inside the folders allowed by the configuration, never overwrite, and
  leave a SHA-256 fingerprint file next to every saved file (useful as a chain of custody).
- Every path coming from Claude is checked as TEXT first (it must lie under an allowed folder, no network paths, no
  alternate data streams, no device names) and only then touched on disk: a network (UNC) path can never make
  Windows open an SMB connection (which would leak the user's NTLM hash).
- Draft attachments may only come from the allowed folders ("attach_folders", default "save_folders").
- Only the mailboxes listed in config.json are reachable; any other store in the Outlook profile is invisible.
- A request that times out in Outlook is cancelled if it has not started yet; while a request is stuck, new ones are
  refused instead of piling up and running later in a burst.
- Every write operation is logged (audit.log): tool, mailbox, id, subject; never the body.

PROMPT INJECTION: the server is used by an AI agent that reads mail written by strangers. A mail may contain text that
tries to steer the agent. The rules above are enforced here, in code, precisely so that they hold even if the agent is
fooled; what the code cannot judge (for example whether a draft's recipient is legitimate) stays with the user, who
reviews and sends every draft.

PHASES (key "phase" in config.json): 1 = read-only (nothing is written, not even files), 2 = also files on disk,
read/unread flags, drafts, contacts, tasks, notes, folders, moves and deletion of mail to the trash, 3 = also calendar,
4 = also the constrained alert sending. Start at 1 and move up as trust grows.
"""

from __future__ import annotations

import asyncio
import datetime as dt
import hashlib
import html as htmlmod
import json
import logging
import logging.handlers
import ntpath
import os
import queue
import re
import sys
import threading
import time
from pathlib import Path
from typing import Any

import pythoncom
import pywintypes
import win32com.client
from mcp.server.mcpserver import MCPServer
from mcp.server.mcpserver.exceptions import ToolError
from mcp.types import ToolAnnotations

# ----------------------------------------------------------------------------
# Configuration
# ----------------------------------------------------------------------------

__version__ = "1.0.0"

# The configuration lives in config.json, next to this file (or in the file named by the environment variable
# OUTLOOK_CLASSIC_MCP_CONFIG, used by the unit tests). Without it the server stops with a clear message:
# copy config.example.json to config.json and fill it in (see README).
BASE = Path(__file__).resolve().parent
_cfg_path = Path(os.environ.get("OUTLOOK_CLASSIC_MCP_CONFIG") or (BASE / "config.json"))
if not _cfg_path.exists():
    sys.stderr.write(f"outlook-classic-mcp: {_cfg_path} is missing. Copy config.example.json to config.json and fill it in.\n")
    sys.exit(1)
CONFIG: dict[str, Any] = json.loads(_cfg_path.read_text("utf-8-sig"))  # utf-8-sig: tolerates the BOM Notepad adds


def _required(key: str) -> Any:
    """Reads a mandatory key of config.json; if it is missing, says which one instead of failing obscurely later."""
    if key not in CONFIG:
        sys.stderr.write(f"outlook-classic-mcp: config.json lacks the mandatory key \"{key}\" (see config.example.json).\n")
        sys.exit(1)
    return CONFIG[key]


# Active phase (1-4): limits which tools are available, see require_phase().
PHASE = int(CONFIG.get("phase", 1))
# Maximum wait for a single call into Outlook. When Outlook has a modal window open (password prompt, warning) it does
# not answer: an explained error is better than a hang.
TIMEOUT = int(CONFIG.get("timeout_seconds", 40))
# Maximum length of a body (mail, event, note, task) returned to Claude, to keep the conversation light.
MAX_BODY = int(CONFIG.get("max_body_chars", 20000))
# Courtesy pauses: every COM call runs on Outlook's main thread, the one that paints the window. During a long read
# (hundreds of items) Outlook would freeze for the user; a short pause every few items lets it answer clicks.
# 0 = no pause. Sensible values: 10-25 ms every 2-3 items.
COURTESY_PAUSE_MS = float(CONFIG.get("courtesy_pause_ms", 10))
PAUSE_EVERY = max(1, int(CONFIG.get("pause_every_items", 3)))
_pause_cache = {"mtime": 0.0, "checked": 0.0}


def refresh_pauses() -> None:
    """Re-reads courtesy_pause_ms and pause_every_items from config.json when the file changed (checked at most every
    5 seconds), so they can be tuned without restarting Claude desktop."""
    global COURTESY_PAUSE_MS, PAUSE_EVERY
    now = time.time()
    if now - _pause_cache["checked"] < 5:
        return
    _pause_cache["checked"] = now
    try:
        m = _cfg_path.stat().st_mtime
        if m != _pause_cache["mtime"]:
            _pause_cache["mtime"] = m
            c = json.loads(_cfg_path.read_text("utf-8-sig"))
            COURTESY_PAUSE_MS = float(c.get("courtesy_pause_ms", 10))
            PAUSE_EVERY = max(1, int(c.get("pause_every_items", 3)))
    except Exception:  # noqa: BLE001
        pass


# Mailboxes: short alias -> name of the store as Outlook shows it (Store.DisplayName, the bold name at the top of each
# account in the folder pane). Claude uses the alias ("work", "pec"...) and the server resolves the store.
# Only these stores are reachable: any other store in the profile (shared or delegated mailboxes, archives, public
# folders) is invisible to the server, even if Claude knows its exact name.
MAILBOXES: dict[str, str] = {str(k).lower(): str(v) for k, v in _required("mailboxes").items()}
DEFAULT_MAILBOX = str(_required("default_mailbox")).lower()
if DEFAULT_MAILBOX not in MAILBOXES:
    sys.stderr.write(f"outlook-classic-mcp: \"default_mailbox\" ({DEFAULT_MAILBOX}) is not one of the aliases in \"mailboxes\".\n")
    sys.exit(1)

# Deletion (move to the mailbox trash only). Switch in config.json, off when absent.
DELETION = bool(CONFIG.get("deletion_enabled", False))

# Alerts: the ONLY exception to the "no sending" rule. Sender and recipient come from the configuration and are NOT
# tool parameters: Claude cannot choose who receives them. The section is optional: when it is missing, send_alert
# refuses with a clear message (handy for whoever wants a server that never sends anything at all).
_alert = CONFIG.get("alert") or {}
ALERT_SENDER = str(_alert.get("sender", "")).strip().lower()         # address of an account present in Outlook
ALERT_RECIPIENT = str(_alert.get("recipient", "")).strip()           # the only possible recipient
ALERT_PREFIX = str(_alert.get("subject_prefix", "[Claude alert] "))  # prepended to every subject, to recognise them

# Folders the server may write files to (attachments, .msg): nowhere else, never overwriting.
SAVE_FOLDERS = [Path(p) for p in _required("save_folders")]
if not SAVE_FOLDERS:
    sys.stderr.write("outlook-classic-mcp: \"save_folders\" must contain at least one folder.\n")
    sys.exit(1)
# Folders draft attachments may be taken from (create_draft). Optional: when absent, the save_folders. Without this
# limit an injected mail could get the agent to attach any file of the PC to a draft addressed to a stranger (drafts
# sync to the mail server even if never sent).
# "attach_folders" absent or empty = create_draft accepts NO attachment at all (safe default): an explicit list is
# required, and it should not be the folder where save_mail archives messages (an injected mail could otherwise have
# archived .msg files of private threads attached to a draft for a stranger).
ATTACH_FOLDERS = [Path(p) for p in (CONFIG.get("attach_folders") or [])]
NEVER_ATTACH = {".msg", ".sha256"}  # archived mail and fingerprints are never attachable
# Maximum alerts per hour (send_alert): a guard against loops and against an injected mail flooding the user. The
# timestamps are kept in a file (LOGDIR/alerts.json), so the limit survives a restart of Claude desktop.
try:
    ALERT_MAX_PER_HOUR = max(1, int(_alert.get("max_per_hour", 10)))
except (TypeError, ValueError):
    sys.stderr.write("outlook-classic-mcp: \"alert.max_per_hour\" must be a whole number.\n")
    sys.exit(1)

# Coloured labels of appointments: the MAPI "Label" property (0x8214) with the historical Outlook 2003/2007 values
# (0 = none ... 10 = phone call). Modern Outlook no longer shows them in its own UI, but some servers still use them:
# Kerio Connect (through KOFF), for instance, maps them onto its coloured labels. On Exchange or IMAP they are written
# and read back with no visible effect: harmless.
LABELS: dict[str, int] = {
    "none": 0, "important": 1, "business": 2, "personal": 3, "vacation": 4, "must_attend": 5,
    "travel": 6, "needs_preparation": 7, "birthday": 8, "anniversary": 9, "phone_call": 10,
}
LABELS_INV = {v: k for k, v in LABELS.items()}
PROP_LABEL = "http://schemas.microsoft.com/mapi/id/{00062002-0000-0000-C000-000000000046}/82140003"

# Logs: server.log (technical, with errors) and audit.log (one JSON line per write: tool, mailbox, id, subject; never
# the content). They live in %LOCALAPPDATA%\outlook-classic-mcp, outside the repository.
LOGDIR = Path(os.environ.get("LOCALAPPDATA", str(BASE))) / "outlook-classic-mcp"
LOGDIR.mkdir(parents=True, exist_ok=True)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    handlers=[logging.handlers.RotatingFileHandler(LOGDIR / "server.log", maxBytes=2_000_000, backupCount=3,
                                                   encoding="utf-8"),
              logging.StreamHandler(sys.stderr)],
)
log = logging.getLogger("outlook-classic-mcp")
AUDIT = LOGDIR / "audit.log"


def audit(tool: str, **fields: Any) -> None:
    """Appends a line to the write log. Every tool that changes something calls it."""
    row = {"when": dt.datetime.now().isoformat(timespec="seconds"), "tool": tool, **fields}
    with AUDIT.open("a", encoding="utf-8") as f:
        f.write(json.dumps(row, ensure_ascii=False) + "\n")


# ----------------------------------------------------------------------------
# COM thread (STA): every call into Outlook goes through here, one at a time, with a timeout.
#
# Why: Outlook COM objects must be used from the thread that created them (single-threaded apartment model). The MCP
# server is asynchronous (asyncio) and may serve several requests; one dedicated thread with a queue guarantees that
# Outlook receives one call at a time, in order, and that a frozen Outlook does not freeze the server.
# ----------------------------------------------------------------------------

class ComWorker:
    def __init__(self) -> None:
        self.q: "queue.Queue[tuple[Any, dict]]" = queue.Queue()  # job queue (function, "future")
        self.app = None                                           # Outlook.Application object, attached on first call
        self.lock = threading.Lock()                              # guards busy_since and the job flags
        self.busy_since: float | None = None                      # monotonic time the running job started
        self.t = threading.Thread(target=self._loop, name="outlook-com", daemon=True)
        self.t.start()

    def _loop(self) -> None:
        """COM thread loop: initialises COM once, then runs the jobs one after the other. A job whose caller already
        gave up (timeout) is skipped: what Claude was told had failed must not happen later behind its back."""
        pythoncom.CoInitialize()
        while True:
            fn, fut = self.q.get()
            with self.lock:
                if fut["cancelled"]:
                    continue
                fut["started"] = True
                self.busy_since = time.monotonic()
            try:
                fut["r"] = fn()
            except BaseException as e:  # noqa: BLE001
                fut["e"] = e
            finally:
                with self.lock:
                    self.busy_since = None
                fut["ev"].set()  # wake up whoever waits in run()

    def run(self, fn):
        """Queues fn and waits for its result for at most TIMEOUT seconds. Called from any thread.
        - If a previous job is stuck inside Outlook for longer than TIMEOUT, the new job is refused at once instead of
          piling up behind it (it would run later, in a burst, after Claude was told it had failed).
        - On timeout, a job that has not started yet is cancelled; a job already running inside Outlook cannot be
          interrupted and may still complete: the error message says so, so the agent checks before retrying."""
        with self.lock:
            if self.busy_since is not None and time.monotonic() - self.busy_since > TIMEOUT:
                raise TimeoutError(
                    "Outlook is still busy with a previous request that timed out (probably a modal window: password "
                    "prompt, PGP passphrase, warning). Nothing new was started. Close the window in Outlook and try again."
                )
        fut: dict[str, Any] = {"ev": threading.Event(), "cancelled": False, "started": False}
        self.q.put((fn, fut))
        if not fut["ev"].wait(TIMEOUT):
            with self.lock:
                started = fut["started"]
                if not started:
                    fut["cancelled"] = True
            if started:
                raise TimeoutError(
                    f"Outlook did not answer within {TIMEOUT} s (probably a modal window: password prompt, PGP passphrase, "
                    "warning). The request had already reached Outlook and MAY STILL COMPLETE when the window is closed: "
                    "check its effect before retrying, to avoid duplicates."
                )
            raise TimeoutError(
                f"Outlook did not answer within {TIMEOUT} s (probably a modal window: password prompt, PGP passphrase, "
                "warning). The request was cancelled and will not run. Close the window in Outlook and try again."
            )
        if "e" in fut:
            raise fut["e"]
        return fut["r"]

    def outlook(self):
        """Attaches to the Outlook session that is already open (GetActiveObject). Never starts Outlook, never asks
        for credentials. GetActiveObject fails when Outlook is closed, still starting with a dialog in front, or running
        with different privileges than this process (for example Outlook "as administrator")."""
        if self.app is not None:
            try:
                _ = self.app.Version  # still alive? If Outlook was closed the object is dead: re-attach.
                return self.app
            except Exception:  # noqa: BLE001
                self.app = None
        try:
            self.app = win32com.client.GetActiveObject("Outlook.Application")
        except pythoncom.com_error as e:
            raise RuntimeError(
                "Classic Outlook does not appear to be running (or it runs as administrator). "
                f"Open Outlook and try again. Detail: {e}"
            ) from e
        return self.app


COM = ComWorker()


async def com(fn):
    """Bridge between the asynchronous MCP server and the COM thread: runs fn there and waits without blocking asyncio.
    Every error reaches Claude as a ToolError with a readable message, never as an anonymous 'Error executing tool':
    this way Claude can tell the user what went wrong (wrong mailbox, Outlook closed, invalid date...)."""
    try:
        return await asyncio.to_thread(COM.run, fn)
    except ToolError:
        raise
    except pywintypes.com_error as e:
        # excepinfo[2] holds Outlook's own error text, when there is one
        detail = e.excepinfo[2] if getattr(e, "excepinfo", None) and len(e.excepinfo) > 2 and e.excepinfo[2] else str(e)
        log.exception("COM error")
        raise ToolError(f"Outlook returned an error: {detail}") from e
    except Exception as e:  # noqa: BLE001
        log.exception("Error")
        raise ToolError(f"{type(e).__name__}: {e}") from e


# ----------------------------------------------------------------------------
# Helpers: navigating stores and folders, dates, result formatting
# ----------------------------------------------------------------------------

def require_phase(n: int) -> None:
    """Blocks a tool when the configured phase is lower than the one it needs."""
    if PHASE < n:
        raise ToolError(f"Tool not enabled: it needs phase {n}, the server runs in phase {PHASE} (config.json).")


def ns():
    """The MAPI NameSpace: the root from which stores, accounts and single items (GetItemFromID) are reached."""
    return COM.outlook().GetNamespace("MAPI")


def store_from_alias(alias: str | None):
    """From a config.json alias (or the exact store name of a configured mailbox) to the Outlook Store object.
    Empty = default mailbox. Stores that are not listed in "mailboxes" are refused, even by exact name."""
    alias = (alias or DEFAULT_MAILBOX).strip()
    if alias.lower() in MAILBOXES:
        name = MAILBOXES[alias.lower()]
    elif alias in MAILBOXES.values():
        name = alias
    else:
        raise ValueError(f"Mailbox '{alias}' is not configured. Configured mailboxes: {', '.join(MAILBOXES)}.")
    matches = [s for s in ns().Stores if s.DisplayName == name]
    if len(matches) > 1:
        raise ValueError(f"Mailbox '{alias}': {len(matches)} stores in Outlook share the name '{name}' (typical of several "
                         "data files called 'Outlook Data File'): rename them in Outlook so that the name is unique.")
    if not matches:
        raise ValueError(f"Mailbox '{alias}' ({name}) is configured but not present in Outlook.")
    return matches[0]


def folder_from_path(store, path: str | None):
    """Path relative to the store root, with '/' (or '\\') between levels, case-insensitive. Empty = store root.
    Example: 'Inbox/Customers/Smith'."""
    f = store.GetRootFolder()
    for part in [p for p in re.split(r"[\\/]", path or "") if p]:
        found = None
        for sub in f.Folders:
            if sub.Name.lower() == part.lower():
                found = sub
                break
        if found is None:
            raise ValueError(f"Folder '{part}' not found in '{f.FolderPath}'.")
        f = found
    return f


def default_folder(store, code: int, *names: str):
    """Default folder of a store (GetDefaultFolder); if the provider does not declare it, looks it up by name
    (English and localised names are tried in order). Useful codes: 6 Inbox, 16 Drafts, 9 Calendar, 10 Contacts,
    13 Tasks, 12 Notes, 3 Deleted Items."""
    try:
        return store.GetDefaultFolder(code)
    except Exception:  # noqa: BLE001
        pass
    last: Exception | None = None
    for n in names:
        try:
            return folder_from_path(store, n)
        except ValueError as ex:
            last = ex
    raise last or ValueError("Default folder not found.")


def inbox(store):
    return default_folder(store, 6, "Inbox", "Posta in arrivo", "Posteingang", "Boîte de réception", "Bandeja de entrada")


def drafts(store):
    return default_folder(store, 16, "Drafts", "Bozze", "Entwürfe", "Brouillons", "Borradores")


def calendar(store, name: str | None):
    """The store's Calendar (default folder 9) or, with 'name', another calendar reached by path. Refuses any folder
    that is not a calendar folder: an event must never end up in Drafts or in a mail folder (some servers would turn
    it into a message)."""
    cal = folder_from_path(store, name) if name else default_folder(store, 9, "Calendar", "Calendario", "Kalender", "Calendrier")
    try:
        if cal.DefaultItemType != 1:
            raise ValueError(f"'{cal.FolderPath}' is not a calendar folder: an event is created only in a Calendar, never in Drafts or in a mail folder.")
    except ValueError:
        raise
    except Exception:  # noqa: BLE001
        pass
    return cal


def check_plain_appointment(e) -> None:
    """Absolute rule: an event or memo is always a plain appointment in a Calendar, never a meeting with attendees
    (Outlook would treat it as an invitation to send or cancel) and never an item in Drafts or in mail.
    The checks are tolerant: an unreadable property is skipped, a readable wrong one is refused."""
    try:
        if int(e.MeetingStatus) != 0:
            raise ValueError("The item is a meeting with attendees: this server neither creates nor edits invitations. The user handles it in Outlook.")
    except ValueError:
        raise
    except Exception:  # noqa: BLE001
        pass
    try:
        if (e.RequiredAttendees or "").strip() or (e.OptionalAttendees or "").strip():
            raise ValueError("The item has attendees: this server neither creates nor edits invitations.")
    except ValueError:
        raise
    except Exception:  # noqa: BLE001
        pass
    try:
        parent = e.Parent
        if parent.DefaultItemType != 1:
            raise ValueError(f"The event sits in '{parent.FolderPath}', which is not a calendar: operation refused.")
    except ValueError:
        raise
    except Exception:  # noqa: BLE001
        pass


def account_for_store(store):
    """The Outlook account that delivers into this store: used to set SendUsingAccount on drafts so that the draft is
    born with the right sender. May not exist (PST stores, public folders): then None."""
    for a in ns().Accounts:
        try:
            if a.DeliveryStore is not None and a.DeliveryStore.StoreID == store.StoreID:
                return a
        except Exception:  # noqa: BLE001
            continue
    return None


def item_by_id(id_: str, mailbox: str | None):
    """From the id (EntryID) returned by the listings to the Outlook object (mail, event, contact...). Also returns
    the Store, needed by callers that must find the Drafts or the trash of the same mailbox."""
    store = store_from_alias(mailbox)
    try:
        item = ns().GetItemFromID(id_, store.StoreID)
    except pythoncom.com_error as e:
        raise ValueError(f"Item not found in mailbox '{mailbox or DEFAULT_MAILBOX}' (invalid id or item moved). {e}") from e
    # Defence in depth: the item must really belong to that (configured) store. Fail closed: if the store of the item
    # cannot be read, the item is refused.
    try:
        same = item.Parent.StoreID == store.StoreID
    except Exception as e:  # noqa: BLE001
        raise ValueError(f"Cannot verify which mailbox the item belongs to: refused. {e}") from e
    if not same:
        raise ValueError(f"The item does not belong to mailbox '{mailbox or DEFAULT_MAILBOX}'.")
    return item, store


def fmt(d) -> str | None:
    """Outlook COM date -> text 'YYYY-MM-DDTHH:MM' (local time), the format used in every answer."""
    if d is None:
        return None
    try:
        return dt.datetime(d.year, d.month, d.day, d.hour, d.minute, d.second).isoformat(timespec="minutes")
    except Exception:  # noqa: BLE001
        return str(d)


def parse_dt(s: str | None, end_of_day: bool = False) -> dt.datetime | None:
    """Text -> datetime. Accepts 'YYYY-MM-DD', 'YYYY-MM-DD HH:MM', 'YYYY-MM-DDTHH:MM'. With end_of_day a date without
    time becomes 23:59:59 (useful as the upper bound of a range)."""
    if not s:
        return None
    s = s.strip().replace("T", " ")
    for f in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d"):
        try:
            d = dt.datetime.strptime(s, f)
            if f == "%Y-%m-%d" and end_of_day:
                d = d.replace(hour=23, minute=59, second=59)
            return d
        except ValueError:
            continue
    raise ValueError(f"Unrecognised date: '{s}'. Use YYYY-MM-DD or YYYY-MM-DD HH:MM.")


def dasl_date(d: dt.datetime) -> str:
    """Date for DASL filters (Items.Restrict with '@SQL=...'): Outlook compares it in UTC and ISO format, regardless of
    the Windows language. Lesson learned: the classic Jet format ("[ReceivedTime] >= 'MM/dd/yyyy'") is parsed with the
    PC's regional settings and on a non-US Windows swaps day and month."""
    # A naive datetime passed to astimezone() is taken as local time WITH the daylight saving rule of that date
    # (not today's offset): a filter on a July date written in November is not shifted by one hour.
    return d.astimezone(dt.timezone.utc).strftime("%Y-%m-%d %H:%M")


def com_date(d: dt.datetime):
    """Date to write into an Outlook property (Start, End, DueDate...). Outlook wants local "wall clock" time; pywin32
    converts a naive datetime from local time to UTC, shifting it. Marked as UTC it reaches Outlook unchanged and is
    read as local time, which is what one expects when writing '2026-11-09 16:00'."""
    return pywintypes.Time(d.replace(tzinfo=dt.timezone.utc))


def dasl_lit(s: str) -> str:
    """Text inside a DASL filter: the quote is doubled, as in SQL."""
    return s.replace("'", "''")


def sender_smtp(item) -> str:
    """SMTP address of the sender. With Exchange, SenderEmailAddress is an unreadable X.500 address: in that case the
    real address comes from GetExchangeUser."""
    try:
        addr = item.SenderEmailAddress or ""
        if item.SenderEmailType == "EX":
            try:
                eu = item.Sender.GetExchangeUser()
                if eu is not None and eu.PrimarySmtpAddress:
                    return eu.PrimarySmtpAddress
            except Exception:  # noqa: BLE001
                pass
        return addr
    except Exception:  # noqa: BLE001
        return ""


def mail_meta(item, alias: str) -> dict[str, Any]:
    """The metadata of a mail as Claude sees it in listings: id, subject, sender, date, unread flag, attachment count,
    folder. Never the body: that is requested separately with read_mail, so listings stay light."""
    d: dict[str, Any] = {"id": item.EntryID, "mailbox": alias}
    for key, attr in (("subject", "Subject"), ("unread", "UnRead"), ("to", "To"), ("cc", "CC")):
        try:
            v = getattr(item, attr)
            d[key] = v.strip() if isinstance(v, str) else v
        except Exception:  # noqa: BLE001
            d[key] = None
    try:
        d["date"] = fmt(item.ReceivedTime)
    except Exception:  # noqa: BLE001
        d["date"] = None
    try:
        d["from"] = f"{item.SenderName} <{sender_smtp(item)}>"
    except Exception:  # noqa: BLE001
        d["from"] = None
    try:
        d["attachments"] = item.Attachments.Count
    except Exception:  # noqa: BLE001
        d["attachments"] = 0
    try:
        d["size"] = item.Size
    except Exception:  # noqa: BLE001
        pass
    try:
        d["folder"] = item.Parent.FolderPath
    except Exception:  # noqa: BLE001
        pass
    return d


def breathe(n: int) -> None:
    """Courtesy pause: called with the count of processed items; sleeps every PAUSE_EVERY items."""
    if COURTESY_PAUSE_MS > 0 and n % PAUSE_EVERY == 0:
        time.sleep(COURTESY_PAUSE_MS / 1000.0)


def iterate(items, limit: int):
    """Iterates an Items collection with GetFirst/GetNext, at most 'limit' items, with courtesy pauses.
    GetFirst/GetNext is faster than index access (items[i]) with providers that do not cache everything."""
    refresh_pauses()
    n = 0
    it = items.GetFirst()
    while it is not None and n < limit:
        yield it
        n += 1
        breathe(n)
        it = items.GetNext()


_RE_DRIVE = re.compile(r"^[A-Za-z]:[\\/]")
_DEVICE_NAMES = {"con", "prn", "aux", "nul", *(f"com{i}" for i in range(1, 10)), *(f"lpt{i}" for i in range(1, 10))}


def _norm(path: str) -> str:
    """Lexical normalisation with Windows rules, no disk access: '..' collapsed, '/' -> '\\', case folded (Windows is
    case-insensitive). ntpath is used explicitly, so the same rules apply when the unit tests run on Linux."""
    return ntpath.normcase(ntpath.normpath(path))


def lexically_inside(path: str, roots: list[Path]) -> bool:
    """True when 'path', as TEXT, is an absolute path inside one of 'roots'. Nothing on disk or on the network is
    touched, which matters on Windows: even a harmless-looking check such as exists() or resolve() on a UNC path
    opens an SMB connection and sends the user's NTLM hash to that host.
    Refused: relative paths, network and device paths (UNC, extended and device prefixes) unless a configured root is itself a
    network path, alternate data streams ('file.txt:stream'), reserved device names (CON, NUL, COM1...)."""
    if not path or "\x00" in path:
        return False
    rest = path[2:] if _RE_DRIVE.match(path) else path
    if ":" in rest:                       # alternate data stream or malformed path
        return False
    for part in re.split(r"[\\/]", rest):
        if part.split(".")[0].strip().lower() in _DEVICE_NAMES:
            return False
    n = _norm(path)
    for root in roots:
        r = _norm(str(root))
        if n == r or n.startswith(r.rstrip("\\") + "\\"):
            return True
    return False


def path_allowed(p: Path, roots: list[Path] | None = None) -> bool:
    """True when p lies inside one of the allowed folders (default: save_folders of config.json).
    Two steps: first the lexical check above (no disk access, so no SMB connection to a host chosen by a mail), then
    resolve() on the real disk, so that links or junctions inside an allowed folder cannot lead outside it."""
    roots = SAVE_FOLDERS if roots is None else roots
    if not roots or not lexically_inside(str(p), roots):
        return False
    rp = p.resolve()
    for root in roots:
        try:
            rp.relative_to(root.resolve())
            return True
        except ValueError:
            continue
    return False


_RE_BIDI = re.compile("[\u200e\u200f\u202a-\u202e\u2066-\u2069]")  # invisible direction marks (fake extensions)


def safe_file_name(name: str, fallback: str = "attachment") -> str:
    """A file name chosen by a stranger (attachment name) made safe for Windows: no folders, no forbidden characters
    (':' included, so no alternate data streams), no invisible direction marks that disguise the extension
    ('invoice\u202Efdp.exe' would display as 'invoiceexe.pdf'), no trailing dots or spaces (Windows strips them silently,
    so the saved name would differ from the one reported), no reserved device names."""
    n = Path(str(name or "").replace("\\", "/")).name
    n = _RE_BIDI.sub("", n)
    n = _RE_FORBIDDEN_NAME.sub("_", n)
    n = n.strip().rstrip(". ")
    if n.split(".")[0].strip().lower() in _DEVICE_NAMES:
        n = "_" + n
    if len(n) > 150:  # keep the extension when shortening
        stem, dot, ext = n.rpartition(".")
        n = (stem[:140] + "." + ext[:9]) if dot and stem else n[:150]
    return n or fallback


# File types that make Windows Explorer contact a remote host as soon as the folder is OPENED (icon or target on a
# UNC path -> NTLM hash sent to that host), or that run when double-clicked. Saved attachments of these types get
# ".txt" appended, so Explorer treats them as text; the original name is kept in the result and in the .sha256 file.
DANGEROUS_EXTENSIONS = {".scf", ".url", ".lnk", ".library-ms", ".website", ".searchconnector-ms", ".settingcontent-ms",
                        ".pif", ".exe", ".com", ".bat", ".cmd", ".scr", ".hta", ".js", ".jse", ".vbs", ".vbe", ".wsf",
                        ".wsh", ".ps1", ".psm1", ".msi", ".msp", ".cpl", ".reg", ".inf", ".iso", ".img", ".vhd", ".vhdx",
                        ".chm", ".application", ".appref-ms", ".xll", ".sha256"}
DANGEROUS_NAMES = {"desktop.ini", "thumbs.db", "autorun.inf"}


def neutralise_name(name: str) -> str:
    """Appends '.txt' to a file name whose type could run or make Explorer reach out (see DANGEROUS_EXTENSIONS)."""
    low = name.lower()
    if low in DANGEROUS_NAMES or Path(low).suffix in DANGEROUS_EXTENSIONS:
        return name + ".txt"
    return name


def mark_of_the_web(file: Path) -> None:
    """Writes the Zone.Identifier stream (ZoneId=3, Internet) like browsers and Outlook do for downloaded files, so
    Office Protected View and SmartScreen still apply to an attachment saved by the server. Best effort (needs NTFS)."""
    try:
        with open(str(file) + ":Zone.Identifier", "w", encoding="ascii") as fh:
            fh.write("[ZoneTransfer]\r\nZoneId=3\r\n")
    except OSError:
        pass


def free_name(folder: Path, name: str) -> Path:
    """A file name not yet used in the folder, RESERVED atomically: the empty file is created with mode 'x' (fails if
    it exists), so nothing can slip in between the check and Outlook's write; Outlook then fills that very file.
    If 'name' is taken, ' (1)', ' (2)'... is appended. A name counts as taken even if only its fingerprint
    '<name>.sha256' exists. The server never overwrites anything."""
    base = safe_file_name(name)
    stem, suffix = Path(base).stem, Path(base).suffix
    i = 0
    while True:
        dest = folder / (base if i == 0 else f"{stem} ({i}){suffix}")
        i += 1
        if dest.with_name(dest.name + ".sha256").exists():
            continue
        try:
            with dest.open("x"):
                pass
            return dest
        except FileExistsError:
            continue


# MAPI property PR_TRANSPORT_MESSAGE_HEADERS: the full Internet headers (Received, Authentication-Results,
# DKIM-Signature, Return-Path...), the ones needed to check SPF, DKIM and DMARC on a suspicious mail.
PROP_HEADERS = "http://schemas.microsoft.com/mapi/proptag/0x007D001F"
OL_MSG_UNICODE = 9  # OlSaveAsType.olMSGUnicode: .msg format with Unicode text (accents survive)
_RE_FORBIDDEN_NAME = re.compile(r'[\\/:*?"<>|\x00-\x1f]+')  # characters Windows refuses in a file name


def name_from_subject(subject: str, limit: int = 80) -> str:
    """Subject cleaned up into a Windows file name: no forbidden characters, collapsed spaces, bounded length."""
    s = _RE_FORBIDDEN_NAME.sub(" ", subject or "")
    s = re.sub(r"\s+", " ", s).strip()
    s = s[:limit].rstrip(" .")
    return s or "no subject"


def date_for_name(item) -> str:
    """YYYY-MM-DD_HHMM from the received time; for drafts (never received) from the creation time."""
    for attr in ("ReceivedTime", "CreationTime"):
        try:
            d = getattr(item, attr)
            if d is not None and 1990 < d.year < 3000:
                return f"{d.year:04d}-{d.month:02d}-{d.day:02d}_{d.hour:02d}{d.minute:02d}"
        except Exception:  # noqa: BLE001
            continue
    return dt.datetime.now().strftime("%Y-%m-%d_%H%M")


def sha256_fingerprint(file: Path, **context: Any) -> dict[str, str]:
    """Computes the SHA-256 of the file just saved and writes '<file>.sha256' next to it (chain of custody: proves the
    file has not changed since it was extracted from Outlook).
    Format: '#' comment lines with acquisition date/time and context, then the standard line
    '<digest>  <file name>', verifiable with 'sha256sum -c' (which skips comments) or by comparing Get-FileHash.
    The .sha256 file is never overwritten (mode 'x')."""
    h = hashlib.sha256()
    with file.open("rb") as fh:
        for block in iter(lambda: fh.read(1024 * 1024), b""):
            h.update(block)
    hexd = h.hexdigest()
    acquired = dt.datetime.now().astimezone().isoformat(timespec="seconds")
    lines = ["# SHA-256 fingerprint for chain of custody (outlook-classic-mcp)",
             f"# file: {file.name}",
             f"# size_bytes: {file.stat().st_size}",
             f"# acquired: {acquired}",
             "# workstation: " + os.environ.get("COMPUTERNAME", "")]
    for k, v in context.items():
        lines.append(f"# {k}: {str(v).replace(chr(10), ' ').replace(chr(13), ' ')}")
    lines.append(f"{hexd}  {file.name}")
    f_fp = file.with_name(file.name + ".sha256")
    with f_fp.open("x", encoding="utf-8", newline="\n") as fh:
        fh.write("\n".join(lines) + "\n")
    return {"sha256": hexd, "sha256_file": str(f_fp), "acquired": acquired}


def text_to_html(text: str) -> str:
    """Plain text -> minimal HTML (special characters escaped, newlines as <br>). Every draft goes through HTML, see
    create_draft for the reason."""
    return "<div style=\"font-family:Calibri,sans-serif;font-size:11pt\">" + htmlmod.escape(text).replace("\n", "<br>") + "</div>"


UTF8_CPID = 65001  # UTF-8 code page for MailItem.InternetCodepage
# References Outlook (Word) would fetch as soon as the draft is opened: a UNC or file: path makes Windows send the
# user's NTLM hash to that host. Any src/href/background/url() pointing there is refused in HTML drafts.
_RE_HTML_REF = re.compile(r"""(?:src|href|background|poster|action|data|codebase|cite|longdesc)\s*=\s*["']?\s*([^"'\s>]+)|url\(\s*["']?([^"')\s]+)""", re.I)
_RE_BAD_REF = re.compile(r"^(?:[\\/]{2}|file:|\\\\\?|[a-z]:[\\/]|mhtml:|res:|javascript:|vbscript:)", re.I)


def check_html_refs(html: str) -> None:
    """Refuses HTML that references a network, local or script URL (see _RE_HTML_REF)."""
    for m in _RE_HTML_REF.finditer(html or ""):
        ref = (m.group(1) or m.group(2) or "").strip()
        ref = re.sub(r"[\x00-\x20]", "", ref)  # control characters and spaces used to hide a scheme
        if _RE_BAD_REF.match(ref):
            raise ValueError(f"HTML refused: it references '{ref[:60]}' (network, local or script URL). Use web URLs "
                             "(https://) or plain text.")
_RE_META_CHARSET = re.compile(r"<meta[^>]*charset[^>]*>", re.I)
_RE_BODY_OPEN = re.compile(r"<body[^>]*>", re.I)


def force_utf8(item) -> None:
    """Sets the UTF-8 code page on the item BEFORE writing HTMLBody. Without it Outlook may re-encode the text with the
    original's code page (Windows-1252, say) and accented letters arrive as 'Ã ' and the like."""
    try:
        item.InternetCodepage = UTF8_CPID
    except Exception:  # noqa: BLE001
        pass


def html_with_text_on_top(existing_html: str, text_html: str) -> str:
    """Inserts text_html right after <body> of the existing HTML (the quoted original in a reply), removing the
    <meta charset> tags Outlook would use to re-encode the content. Without <body>, puts it on top."""
    base = _RE_META_CHARSET.sub("", existing_html or "")
    m = _RE_BODY_OPEN.search(base)
    if m:
        return base[:m.end()] + text_html + "<br>" + base[m.end():]
    return text_html + "<br>" + base


def event_meta(e, alias: str, cal_path: str | None = None) -> dict[str, Any]:
    """The metadata of an appointment for listings: id, subject, location, start, end, duration, recurring flag,
    coloured label, categories, reminder. The description is read separately with read_event."""
    d: dict[str, Any] = {"mailbox": alias}
    try:
        d["id"] = e.EntryID
    except Exception as ex:  # noqa: BLE001
        # A modified occurrence of a series Outlook can no longer resolve: flag it and carry on.
        d["id"] = None
        d["warning"] = "unreadable occurrence: " + (ex.excepinfo[2] if getattr(ex, "excepinfo", None) and len(ex.excepinfo) > 2 and ex.excepinfo[2] else str(ex))
    for key, attr in (("subject", "Subject"), ("location", "Location"), ("all_day", "AllDayEvent"),
                      ("duration_minutes", "Duration"), ("categories", "Categories"), ("recurring", "IsRecurring"),
                      ("reminder", "ReminderSet")):
        try:
            d[key] = getattr(e, attr)
        except Exception:  # noqa: BLE001
            d[key] = None
    try:
        d["start"] = fmt(e.Start)
        d["end"] = fmt(e.End)
    except Exception:  # noqa: BLE001
        pass
    try:
        lbl = int(e.PropertyAccessor.GetProperty(PROP_LABEL))
        d["label"] = LABELS_INV.get(lbl, str(lbl))
    except Exception:  # noqa: BLE001
        d["label"] = None
    if cal_path:
        d["calendar"] = cal_path
    return d


# ----------------------------------------------------------------------------
# Helpers for contacts, tasks, notes, folders
# ----------------------------------------------------------------------------

# OlObjectClass codes of items: used to check that an id really points to the expected kind.
CLASSES = {43: "mail", 26: "event", 40: "contact", 48: "task", 44: "note", 46: "report", 53: "meeting"}
# OlItemType codes of folders (Folder.DefaultItemType).
FOLDER_TYPES = {0: "mail", 1: "calendar", 2: "contacts", 3: "tasks", 4: "journal", 5: "notes"}
# Names the various providers give the trash, in several languages: to recognise it on IMAP too.
TRASH_NAMES = {"deleted items", "deleted messages", "trash", "bin",
               "posta eliminata", "cestino", "elementi eliminati",
               "gelöschte elemente", "gelöschte objekte", "papierkorb",
               "éléments supprimés", "corbeille",
               "elementos eliminados", "papelera",
               "verwijderde items", "prullenbak",
               "itens excluídos", "lixeira", "lixo",
               "elementy usunięte", "kosz",
               "borttaget", "papperskorgen", "slettet post", "papirkurv", "slettede elementer", "poistetut"}


def check_class(item, expected: int, what: str) -> None:
    """Refuses an id that points to a different kind of item (for example a contact passed to delete_mail)."""
    try:
        c = item.Class
    except Exception:  # noqa: BLE001
        return
    if c != expected:
        raise ValueError(f"The given id is not {what} (class {CLASSES.get(c, c)}).")


def is_trash(store, fold) -> bool:
    """True when the folder is the store's trash or sits anywhere inside it. Checked by identity (EntryID of the
    trash, walking up the parent folders), so it works in every Outlook language; the names in TRASH_NAMES are an
    extra net for providers that do not declare a trash (IMAP)."""
    trash_ids = set()
    try:
        trash_ids.add(store.GetDefaultFolder(3).EntryID)
    except Exception:  # noqa: BLE001
        try:
            trash_ids.add(trash_of(store).EntryID)
        except Exception:  # noqa: BLE001
            pass
    f = fold
    for _ in range(32):  # bounded walk up to the store root
        try:
            if f.EntryID in trash_ids:
                return True
            f = f.Parent
            if getattr(f, "Class", None) != 2:  # 2 = olFolder; above the root comes the NameSpace
                break
        except Exception:  # noqa: BLE001
            break
    try:
        parts = [p.lower() for p in fold.FolderPath.split("\\") if p]
        return any(p in TRASH_NAMES for p in parts)
    except Exception:  # noqa: BLE001
        return False


def folders_of_type(store, kind: int, depth: int = 3):
    """All folders of a store with a given DefaultItemType (2 = contacts, 3 = tasks, 5 = notes)."""
    out = []

    def walk(fold, level):
        for sub in fold.Folders:
            try:
                if sub.DefaultItemType == kind:
                    out.append(sub)
            except Exception:  # noqa: BLE001
                pass
            if level < depth:
                walk(sub, level + 1)

    walk(store.GetRootFolder(), 1)
    return out


def contact_meta(c, alias: str, full: bool = False) -> dict[str, Any]:
    """A contact card: main fields in listings, every field (addresses, notes...) with full=True.
    Empty fields are omitted to keep answers short."""
    d: dict[str, Any] = {"mailbox": alias}
    try:
        d["id"] = c.EntryID
    except Exception:  # noqa: BLE001
        d["id"] = None
    fields = [("full_name", "FullName"), ("company", "CompanyName"), ("job_title", "JobTitle"),
              ("email", "Email1Address"), ("email2", "Email2Address"), ("business_phone", "BusinessTelephoneNumber"),
              ("mobile_phone", "MobileTelephoneNumber"), ("home_phone", "HomeTelephoneNumber"), ("categories", "Categories")]
    if full:
        fields += [("first_name", "FirstName"), ("last_name", "LastName"), ("email3", "Email3Address"),
                   ("business_address", "BusinessAddress"), ("home_address", "HomeAddress"), ("website", "WebPage"),
                   ("department", "Department"), ("notes", "Body"), ("file_as", "FileAs")]
    for key, attr in fields:
        try:
            v = getattr(c, attr)
            if v not in (None, ""):
                d[key] = v.strip() if isinstance(v, str) else v
        except Exception:  # noqa: BLE001
            pass
    try:
        d["folder"] = c.Parent.FolderPath
    except Exception:  # noqa: BLE001
        pass
    return d


def task_meta(t, alias: str, full: bool = False) -> dict[str, Any]:
    """Data of a task. 'Empty' dates in Outlook are 1.1.4501: they are returned as null."""
    d: dict[str, Any] = {"mailbox": alias}
    try:
        d["id"] = t.EntryID
    except Exception:  # noqa: BLE001
        d["id"] = None
    for key, attr in (("subject", "Subject"), ("complete", "Complete"), ("percent_complete", "PercentComplete"),
                      ("priority", "Importance"), ("categories", "Categories"), ("status", "Status"),
                      ("reminder", "ReminderSet")):
        try:
            d[key] = getattr(t, attr)
        except Exception:  # noqa: BLE001
            d[key] = None
    for key, attr in (("due_date", "DueDate"), ("start_date", "StartDate"), ("completed_on", "DateCompleted"),
                      ("reminder_time", "ReminderTime")):
        try:
            v = getattr(t, attr)
            d[key] = None if (v is None or v.year >= 4500) else fmt(v)
        except Exception:  # noqa: BLE001
            d[key] = None
    if full:
        try:
            d["description"] = (t.Body or "")[:MAX_BODY]
        except Exception:  # noqa: BLE001
            pass
    try:
        d["folder"] = t.Parent.FolderPath
    except Exception:  # noqa: BLE001
        pass
    return d


def note_meta(n, alias: str, full: bool = False) -> dict[str, Any]:
    """Data of a note (Outlook's sticky notes): title (first line), categories, modification time, text when full."""
    d: dict[str, Any] = {"mailbox": alias}
    try:
        d["id"] = n.EntryID
    except Exception:  # noqa: BLE001
        d["id"] = None
    for key, attr in (("title", "Subject"), ("categories", "Categories")):
        try:
            d[key] = getattr(n, attr)
        except Exception:  # noqa: BLE001
            d[key] = None
    try:
        d["modified"] = fmt(n.LastModificationTime)
    except Exception:  # noqa: BLE001
        pass
    if full:
        try:
            d["text"] = (n.Body or "")[:MAX_BODY]
        except Exception:  # noqa: BLE001
            pass
    return d


def _apply_contact(c, first_name, last_name, company, job_title, email, email2, business_phone, mobile_phone,
                   home_phone, business_address, website, notes, categories) -> None:
    """Writes only the given fields onto the contact (None = leave alone). Shared by create_contact and update_contact."""
    for value, attr in ((first_name, "FirstName"), (last_name, "LastName"), (company, "CompanyName"), (job_title, "JobTitle"),
                        (email, "Email1Address"), (email2, "Email2Address"), (business_phone, "BusinessTelephoneNumber"),
                        (mobile_phone, "MobileTelephoneNumber"), (home_phone, "HomeTelephoneNumber"),
                        (business_address, "BusinessAddress"), (website, "WebPage"), (notes, "Body"),
                        (categories, "Categories")):
        if value is not None:
            setattr(c, attr, value)


def _apply_task(t, subject, due_date, start_date, description, priority, reminder, categories, complete,
                percent_complete) -> None:
    """Writes only the given fields onto the task (None = leave alone). reminder='' removes the reminder.
    Shared by create_task and update_task."""
    if subject is not None:
        t.Subject = subject
    if due_date is not None:
        t.DueDate = com_date(parse_dt(due_date))
    if start_date is not None:
        t.StartDate = com_date(parse_dt(start_date))
    if description is not None:
        t.Body = description
    if priority is not None:
        p = {"low": 0, "normal": 1, "high": 2}.get(str(priority).lower())
        if p is None:
            raise ValueError("Invalid priority: low, normal, high.")
        t.Importance = p
    if reminder is not None:
        if reminder == "":
            t.ReminderSet = False
        else:
            t.ReminderSet = True
            t.ReminderTime = com_date(parse_dt(reminder))
    if categories is not None:
        t.Categories = categories
    if percent_complete is not None:
        t.PercentComplete = int(percent_complete)
    if complete is not None:
        if complete:
            t.MarkComplete()
        else:
            t.Complete = False
            t.Status = 0


# ----------------------------------------------------------------------------
# MCP server
#
# `instructions` is the text Claude receives together with the tool list: it says what is there and what the rules
# are. The tools are the functions decorated with @mcp.tool; each docstring becomes the description Claude reads to
# decide when and how to use it, so it is written for Claude.
# ----------------------------------------------------------------------------

mcp = MCPServer(
    "outlook",
    instructions=(
        "Local connector for the user's classic Outlook, on their PC. "
        f"Default mailbox: alias '{DEFAULT_MAILBOX}' (store \"{MAILBOXES[DEFAULT_MAILBOX]}\"); "
        "other mailboxes by alias: " + ", ".join(MAILBOXES) + ". "
        "Covers mail, calendar, contacts, tasks, notes and folders. It does not send (except send_alert, with a fixed "
        "recipient) and does not touch rules or settings. It deletes only with delete_mail and delete_event, which move "
        "to the trash (recoverable) and must be used ONLY after the user's explicit confirmation in chat. "
        "Unread mail stays unread after reading. "
        f"Active phase: {PHASE} (1 read, 2 drafts and writes, 3 calendar, 4 alerts)."
    ),
)

# MCP annotations: tell the client whether a tool is read-only or destructive (Claude desktop uses them to decide
# when to ask the user for confirmation).
RO = ToolAnnotations(read_only_hint=True, destructive_hint=False, open_world_hint=False)
RW = ToolAnnotations(read_only_hint=False, destructive_hint=False, open_world_hint=False)
# Writes that change or overwrite existing data (updates, moves): flagged destructive so the client may ask first.
UPD = ToolAnnotations(read_only_hint=False, destructive_hint=True, idempotent_hint=False, open_world_hint=False)
# Writes that reach the outside world: a draft syncs to the mail server, an alert is sent.
OUT = ToolAnnotations(read_only_hint=False, destructive_hint=False, open_world_hint=True)

def restore_unread(item, was_unread: bool) -> None:
    """Puts back the unread flag if reading the item cleared it. Called in 'finally', so it holds on errors too."""
    try:
        if was_unread and not item.UnRead:
            item.UnRead = True
            item.Save()
    except Exception:  # noqa: BLE001
        log.warning("could not restore the unread flag on an item")


INBOX_NAMES = {"", "inbox", "posta in arrivo", "posteingang", "boîte de réception", "bandeja de entrada"}


def mail_folder(store, folder: str):
    """Inbox by any of its usual names, otherwise a folder by path."""
    return inbox(store) if folder.strip().lower() in INBOX_NAMES else folder_from_path(store, folder)


# ---------------------------------------------------------------- phase 1: reading mail

@mcp.tool(annotations=RO)
async def outlook_status() -> dict:
    """Checks that classic Outlook is open and reachable; reports version, profile, active phase and available mailboxes."""
    def f():
        app = COM.outlook()
        n = app.GetNamespace("MAPI")
        stores = [s.DisplayName for s in n.Stores]
        # Only the configured mailboxes are reported: the names of the other stores in the profile are not exposed
        # (install.ps1 prints them for whoever fills in config.json).
        return {
            "ok": True, "server_version": __version__, "outlook_version": app.Version, "profile": n.CurrentProfileName,
            "phase": PHASE, "default_mailbox": DEFAULT_MAILBOX,
            "mailboxes": {a: {"store": name, "present": name in stores} for a, name in MAILBOXES.items()},
        }
    return await com(f)


@mcp.tool(annotations=RO)
async def list_folders(mailbox: str | None = None, path: str = "", depth: int = 2) -> list[dict]:
    """Lists the folders of a mailbox (alias or store name) starting from a path (empty = root), with item counts.
    Types: mail, calendar, contacts, tasks, notes. depth is capped at 6."""
    depth = max(1, min(int(depth), 6))

    def f():
        store = store_from_alias(mailbox)
        root = folder_from_path(store, path)
        out: list[dict] = []

        def walk(fold, prefix: str, level: int):
            for sub in fold.Folders:
                breathe(len(out) + 1)
                p = f"{prefix}/{sub.Name}" if prefix else sub.Name
                row = {"path": p, "type": FOLDER_TYPES.get(sub.DefaultItemType, str(sub.DefaultItemType))}
                try:
                    row["items"] = sub.Items.Count
                    row["unread"] = sub.UnReadItemCount
                except Exception:  # noqa: BLE001
                    pass
                out.append(row)
                if level < depth:
                    walk(sub, p, level + 1)

        walk(root, path.strip("/\\"), 1)
        return out
    return await com(f)


@mcp.tool(annotations=RO)
async def list_mail(mailbox: str | None = None, folder: str = "Inbox", days: int = 7,
                    limit: int = 30, unread_only: bool = False) -> list[dict]:
    """Lists the most recent mail of a folder (default: Inbox of the default mailbox), newest first.
    Returns metadata only (id, date, sender, subject, unread flag, attachment count), never the body."""
    def f():
        store = store_from_alias(mailbox)
        fold = mail_folder(store, folder)
        since = dt.datetime.now() - dt.timedelta(days=max(1, days))
        flt = f"@SQL=\"urn:schemas:httpmail:datereceived\" >= '{dasl_date(since)}'"
        if unread_only:
            flt += " AND \"urn:schemas:httpmail:read\" = 0"
        items = fold.Items.Restrict(flt)
        items.Sort("[ReceivedTime]", True)
        return [mail_meta(it, mailbox or DEFAULT_MAILBOX) for it in iterate(items, max(1, min(limit, 200)))]
    return await com(f)


@mcp.tool(annotations=RO)
async def search_mail(text: str, mailbox: str | None = None, folder: str = "Inbox",
                      in_subject: bool = True, in_body: bool = False, sender: str | None = None,
                      from_date: str | None = None, to_date: str | None = None, limit: int = 30) -> list[dict]:
    """Searches mail in a folder by text in the subject and/or body, sender (name or address) and date range (YYYY-MM-DD).
    Body search (in_body=true) is slow and freezes Outlook: it requires from_date. Returns metadata only."""
    def f():
        store = store_from_alias(mailbox)
        fold = mail_folder(store, folder)
        cond: list[str] = []
        t = dasl_lit(text)
        parts = []
        if in_subject:
            parts.append(f"\"urn:schemas:httpmail:subject\" LIKE '%{t}%'")
        if in_body:
            parts.append(f"\"urn:schemas:httpmail:textdescription\" LIKE '%{t}%'")
        if parts:
            cond.append("(" + " OR ".join(parts) + ")")
        if sender:
            m = dasl_lit(sender)
            cond.append(f"(\"urn:schemas:httpmail:fromname\" LIKE '%{m}%' OR \"urn:schemas:httpmail:fromemail\" LIKE '%{m}%')")
        d1, d2 = parse_dt(from_date), parse_dt(to_date, end_of_day=True)
        if d1:
            cond.append(f"\"urn:schemas:httpmail:datereceived\" >= '{dasl_date(d1)}'")
        if d2:
            cond.append(f"\"urn:schemas:httpmail:datereceived\" <= '{dasl_date(d2)}'")
        if not cond:
            raise ValueError("No search criteria.")
        if in_body and not d1:
            raise ValueError("Body search requires from_date (and preferably to_date): on large folders it would freeze Outlook for seconds.")
        items = fold.Items.Restrict("@SQL=" + " AND ".join(cond))
        items.Sort("[ReceivedTime]", True)
        return [mail_meta(it, mailbox or DEFAULT_MAILBOX) for it in iterate(items, max(1, min(limit, 200)))]
    return await com(f)


@mcp.tool(annotations=RO)
async def read_mail(id: str, mailbox: str | None = None, format: str = "text", max_chars: int | None = None,
                    headers: bool = False) -> dict:
    """Reads a mail (id from list_mail/search_mail): metadata, body (format 'text' or 'html') and attachment list.
    With headers=true adds 'transport_headers': the full Internet headers (Received, Authentication-Results,
    DKIM-Signature, Return-Path...) to check SPF, DKIM and DMARC. If the mail was unread, it stays unread."""
    def f():
        item, _ = item_by_id(id, mailbox)
        was_unread = bool(item.UnRead)
        try:
            d = mail_meta(item, mailbox or DEFAULT_MAILBOX)
            lim = max(1, min(max_chars or MAX_BODY, MAX_BODY))
            body = item.HTMLBody if format == "html" else item.Body
            body = body or ""
            d["body"] = body[:lim]
            d["body_truncated"] = len(body) > lim
            atts = []
            for i in range(1, item.Attachments.Count + 1):
                a = item.Attachments.Item(i)
                atts.append({"index": i, "name": a.FileName, "size": a.Size})
            d["attachment_list"] = atts
            if headers:
                try:
                    d["transport_headers"] = item.PropertyAccessor.GetProperty(PROP_HEADERS) or ""
                except Exception:  # noqa: BLE001
                    d["transport_headers"] = ""
                if not d["transport_headers"]:
                    d["headers_note"] = ("Transport headers not available for this item (drafts, items created in Outlook, "
                                         "or not kept by the store). For forensic proof the original .eml from the server is needed.")
        finally:
            restore_unread(item, was_unread)
        d["unread"] = bool(item.UnRead)
        return d
    return await com(f)


@mcp.tool(annotations=RW)
async def save_attachments(id: str, destination_folder: str, mailbox: str | None = None,
                           indexes: list[int] | None = None) -> list[dict]:
    """Saves the attachments of a mail to disk (all of them, or those with the given indexes from read_mail) into an
    allowed folder (one of "save_folders" in config.json, or a subfolder). Never overwrites existing files; writes a
    SHA-256 fingerprint next to each file and marks it as downloaded from the Internet (Mark of the Web). File names
    are cleaned up ('name' in the result is the original, 'saved_to' the real path). If the mail was unread, it stays
    unread."""
    require_phase(2)

    def f():
        dest = Path(destination_folder)
        if not path_allowed(dest):
            raise ValueError("Folder not allowed. Allowed: " + "; ".join(str(p) for p in SAVE_FOLDERS))
        dest.mkdir(parents=True, exist_ok=True)
        item, _ = item_by_id(id, mailbox)
        was_unread = bool(item.UnRead)
        out = []
        try:
            for i in range(1, item.Attachments.Count + 1):
                if indexes and i not in indexes:
                    continue
                a = item.Attachments.Item(i)
                target = free_name(dest, neutralise_name(safe_file_name(a.FileName)))
                try:
                    a.SaveAsFile(str(target))  # fills the file reserved by free_name
                except Exception:
                    target.unlink(missing_ok=True)
                    raise
                mark_of_the_web(target)
                fp = sha256_fingerprint(target, origin="mail attachment", mailbox=mailbox or DEFAULT_MAILBOX,
                                        mail_subject=item.Subject or "", attachment=a.FileName, mail_id=id)
                out.append({"index": i, "name": a.FileName, "saved_to": str(target), "size": target.stat().st_size, **fp})
        finally:
            restore_unread(item, was_unread)
        audit("save_attachments", mailbox=mailbox or DEFAULT_MAILBOX, id=id, files=[o["saved_to"] for o in out],
              sha256=[o["sha256"] for o in out])
        return out
    return await com(f)


@mcp.tool(annotations=RW)
async def save_mail(id: str, destination_folder: str, mailbox: str | None = None) -> dict:
    """Saves the whole mail as a .msg file (Unicode): transport headers, text, HTML and attachments.
    Only into an allowed folder ("save_folders" in config.json, or a subfolder), never overwriting.
    Name: YYYY-MM-DD_HHMM_<subject>.msg with the received date. Writes <file>.sha256 next to it with the fingerprint
    and the acquisition date/time. If the mail was unread, it stays unread.
    Note: for forensic proof the .msg does NOT replace the original .eml (the DKIM signature cannot be re-verified)."""
    require_phase(2)

    def f():
        dest = Path(destination_folder)
        if not path_allowed(dest):
            raise ValueError("Folder not allowed. Allowed: " + "; ".join(str(p) for p in SAVE_FOLDERS))
        dest.mkdir(parents=True, exist_ok=True)
        item, _ = item_by_id(id, mailbox)
        if int(getattr(item, "Class", 0)) != 43:
            raise ValueError("The item is not a mail (save_mail saves mail messages only).")
        was_unread = bool(item.UnRead)
        subject = item.Subject or ""
        target = free_name(dest, f"{date_for_name(item)}_{name_from_subject(subject)}.msg")
        try:
            item.SaveAs(str(target), OL_MSG_UNICODE)  # fills the file reserved by free_name
        except Exception:
            target.unlink(missing_ok=True)  # the empty reserved file must not stay behind
            raise
        finally:
            restore_unread(item, was_unread)
        if not target.is_file() or target.stat().st_size == 0:
            target.unlink(missing_ok=True)
            raise ValueError(f"Outlook did not write the file {target}.")
        fp = sha256_fingerprint(target, origin="whole mail (.msg Unicode, MailItem.SaveAs)",
                                mailbox=mailbox or DEFAULT_MAILBOX, mail_subject=subject, mail_id=id,
                                note="the .msg does not replace the original .eml for DKIM verification")
        audit("save_mail", mailbox=mailbox or DEFAULT_MAILBOX, id=id, file=str(target), sha256=fp["sha256"])
        return {"id": id, "subject": subject, "saved_to": str(target), "size": target.stat().st_size,
                "unread": bool(item.UnRead), **fp}
    return await com(f)


@mcp.tool(annotations=RW)
async def mark_mail(id: str, read: bool, mailbox: str | None = None) -> dict:
    """Marks a mail as read (read=true) or unread (read=false)."""
    require_phase(2)

    def f():
        item, _ = item_by_id(id, mailbox)
        item.UnRead = not read
        item.Save()
        audit("mark_mail", mailbox=mailbox or DEFAULT_MAILBOX, id=id, read=read)
        return {"id": id, "unread": bool(item.UnRead)}
    return await com(f)


# ---------------------------------------------------------------- phase 2: drafts

@mcp.tool(annotations=OUT)
async def create_draft(subject: str, body: str, to: str = "", cc: str = "", mailbox: str | None = None,
                       format: str = "text", attachments: list[str] | None = None) -> dict:
    """Creates a mail draft in the Drafts folder of the mailbox (default: the default mailbox of config.json).
    Never sends: the draft stays in Outlook for the user to review, sign if they wish, and send themselves.
    Recipients separated by ';'. format: 'text' or 'html' (body already in HTML). 'attachments' = file paths on the PC,
    only inside the allowed folders ("attach_folders" in config.json, default "save_folders"): any other path is
    refused, so a mail cannot get a private file of the PC attached to a draft for a stranger."""
    require_phase(2)

    def f():
        store = store_from_alias(mailbox)
        fold = drafts(store)
        m = fold.Items.Add(0)  # olMailItem, created directly in the Drafts folder of the right mailbox
        acc = account_for_store(store)
        if acc is not None:
            m.SendUsingAccount = acc  # so the draft already carries that mailbox's sender
        m.Subject = subject
        m.To = to
        m.CC = cc
        # Even the 'text' format goes through HTML with the UTF-8 code page: writing m.Body leaves the draft in the
        # profile's default code page and accented letters may arrive corrupted.
        force_utf8(m)
        if format == "html":
            check_html_refs(body)
            m.HTMLBody = _RE_META_CHARSET.sub("", body)
        else:
            m.HTMLBody = text_to_html(body)
        # Every attachment is checked BEFORE the draft is saved: a refused path leaves no half-made draft behind.
        files = []
        for p in attachments or []:
            pp = Path(p)
            if not ATTACH_FOLDERS:
                raise ValueError("Attachments are disabled: \"attach_folders\" is not set in config.json.")
            if not path_allowed(pp, ATTACH_FOLDERS):
                raise ValueError(f"Attachment refused: {p} is outside the allowed folders ("
                                 + "; ".join(str(x) for x in ATTACH_FOLDERS) + ").")
            if pp.suffix.lower() in NEVER_ATTACH:
                raise ValueError(f"Attachment refused: files of type {pp.suffix} (archived mail, fingerprints) are never attached.")
            if not pp.is_file():
                raise ValueError(f"Attachment not found: {p}")
            files.append(pp.resolve())
        for pp in files:
            m.Attachments.Add(str(pp))
        m.Save()
        audit("create_draft", mailbox=mailbox or DEFAULT_MAILBOX, id=m.EntryID, subject=subject, to=to, cc=cc,
              attachments=[str(x) for x in files])
        return {"id": m.EntryID, "folder": fold.FolderPath, "subject": subject, "to": to, "cc": cc, "sent": False}
    return await com(f)


@mcp.tool(annotations=OUT)
async def create_reply_draft(id: str, body: str, mailbox: str | None = None, reply_all: bool = False) -> dict:
    """Creates a reply draft to a mail (id from list_mail/search_mail), with the given text above the quoted original.
    Never sends. If the original mail was unread, it stays unread."""
    require_phase(2)

    def f():
        item, store = item_by_id(id, mailbox)
        was_unread = bool(item.UnRead)
        try:
            r = _build_reply(item, store, body, reply_all)
        finally:
            restore_unread(item, was_unread)
        audit("create_reply_draft", mailbox=mailbox or DEFAULT_MAILBOX, original_id=id, id=r.EntryID)
        return {"id": r.EntryID, "folder": r.Parent.FolderPath, "subject": r.Subject, "to": r.To, "cc": r.CC, "sent": False}
    return await com(f)


def _build_reply(item, store, body: str, reply_all: bool):
    """Creates and saves the reply draft (see create_reply_draft)."""
    r = item.ReplyAll() if reply_all else item.Reply()
    # Mandatory order: FIRST read the reply's HTML with the original code page (otherwise Windows-1252 characters
    # re-read as UTF-8 become "�"), THEN force UTF-8, THEN write the new HTML.
    html_orig = r.HTMLBody or ""
    force_utf8(r)
    r.HTMLBody = html_with_text_on_top(html_orig, text_to_html(body))
    r.Save()
    # Reply() saves into the Drafts of the profile's default store: if the mail lived in another mailbox, the draft
    # is moved to that mailbox's Drafts, where the user expects it.
    # Fail closed: if the draft cannot be moved there, it is removed and the error is reported, so that no draft is
    # left in a store that is not configured.
    try:
        if r.Parent.StoreID != store.StoreID:
            r = r.Move(drafts(store))
    except Exception as e:  # noqa: BLE001
        try:
            r.Delete()
        except Exception:  # noqa: BLE001
            pass
        raise ValueError(f"The reply draft could not be placed in the Drafts of the mailbox and was discarded. {e}") from e
    return r


# ---------------------------------------------------------------- phase 3: calendar

@mcp.tool(annotations=RO)
async def list_events(from_date: str | None = None, to_date: str | None = None, mailbox: str | None = None,
                      calendar_name: str | None = None, limit: int = 100) -> list[dict]:
    """Lists the events and memos of a calendar between two dates (YYYY-MM-DD; default: today to +7 days), with
    recurrences expanded (every occurrence of a series is a row, carrying the id of the series).
    'calendar_name' empty = main Calendar; otherwise the path of another calendar (see list_folders)."""
    def f():
        store = store_from_alias(mailbox)
        cal = calendar(store, calendar_name)
        d1 = parse_dt(from_date) or dt.datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        d2 = parse_dt(to_date, end_of_day=True) or (d1 + dt.timedelta(days=7))
        # Sequence used to expand series (verified in the field): IncludeRecurrences = True, Sort on [Start], then
        # Restrict with a DASL date filter (ISO, UTC). Iterate with the enumerator, never with Count/index (Count is
        # undefined with recurrences).
        items = cal.Items
        items.IncludeRecurrences = True
        items.Sort("[Start]")
        r = items.Restrict(f"@SQL=\"urn:schemas:calendar:dtstart\" >= '{dasl_date(d1)}' AND \"urn:schemas:calendar:dtstart\" <= '{dasl_date(d2)}'")
        out = []
        lim = max(1, min(limit, 500))
        refresh_pauses()
        for e in r:
            # guard: with recurrences expanded the enumeration must never run past the end date
            try:
                if dt.datetime(e.Start.year, e.Start.month, e.Start.day, e.Start.hour, e.Start.minute) > d2:
                    break
            except Exception:  # noqa: BLE001
                pass
            out.append(event_meta(e, mailbox or DEFAULT_MAILBOX, cal.FolderPath))
            breathe(len(out))
            if len(out) >= lim:
                break
        return out
    return await com(f)


# Outlook OlRecurrenceType codes and OlDaysOfWeek mask.
RECURRENCE_TYPES = {"daily": 0, "weekly": 1, "monthly": 2, "yearly": 5}
WEEKDAYS = {"sun": 1, "mon": 2, "tue": 4, "wed": 8, "thu": 16, "fri": 32, "sat": 64}
RECURRENCE_NAMES = {0: "daily", 1: "weekly", 2: "monthly", 3: "monthly (nth weekday)", 5: "yearly", 6: "yearly (nth weekday)"}


@mcp.tool(annotations=RO)
async def read_event(id: str, mailbox: str | None = None) -> dict:
    """Reads an event or memo of the calendar (id from list_events), with its full description."""
    def f():
        e, _ = item_by_id(id, mailbox)
        d = event_meta(e, mailbox or DEFAULT_MAILBOX)
        d["description"] = (e.Body or "")[:MAX_BODY]
        try:
            d["reminder_minutes_before"] = e.ReminderMinutesBeforeStart
            d["attendees"] = e.RequiredAttendees
            if e.IsRecurring:
                rp = e.GetRecurrencePattern()
                d["recurrence"] = {"type": RECURRENCE_NAMES.get(rp.RecurrenceType, rp.RecurrenceType), "interval": rp.Interval,
                                   "start": fmt(rp.PatternStartDate), "end": None if rp.NoEndDate else fmt(rp.PatternEndDate)}
        except Exception:  # noqa: BLE001
            pass
        return d
    return await com(f)


def _apply_event(e, subject, start, end, duration_minutes, description, location, all_day,
                 label, categories, reminder_minutes) -> None:
    """Writes only the given fields onto the appointment (None = leave alone). Shared by create_event and update_event.
    Negative reminder_minutes = no reminder. The coloured label goes into the MAPI property PROP_LABEL."""
    if subject is not None:
        e.Subject = subject
    if start is not None:
        e.Start = com_date(parse_dt(start))
    if end is not None:
        e.End = com_date(parse_dt(end))
    elif duration_minutes is not None:
        e.Duration = int(duration_minutes)
    if description is not None:
        e.Body = description
    if location is not None:
        e.Location = location
    if all_day is not None:
        e.AllDayEvent = bool(all_day)
    if categories is not None:
        e.Categories = categories
    if reminder_minutes is not None:
        if int(reminder_minutes) < 0:
            e.ReminderSet = False
        else:
            e.ReminderSet = True
            e.ReminderMinutesBeforeStart = int(reminder_minutes)
    if label is not None:
        k = label.strip().lower()
        if k not in LABELS:
            raise ValueError("Invalid label. Valid: " + ", ".join(LABELS))
        e.PropertyAccessor.SetProperty(PROP_LABEL, LABELS[k])


def _apply_recurrence(e, recurrence: str, interval: int | None, weekdays: str | None,
                      recurrence_end: str | None, occurrences: int | None) -> dict:
    """Turns the appointment into a series. Must be called after Start, End and Duration are set, because the
    RecurrencePattern takes time and duration from them. Property order verified in the field: RecurrenceType, then
    DayOfWeekMask or DayOfMonth/MonthOfYear, Interval, PatternStartDate, StartTime, EndTime, Duration, and the end
    last (PatternEndDate, Occurrences or NoEndDate)."""
    k = (recurrence or "").strip().lower()
    if k not in RECURRENCE_TYPES:
        raise ValueError("Invalid recurrence: daily, weekly, monthly, yearly.")
    start = dt.datetime(e.Start.year, e.Start.month, e.Start.day, e.Start.hour, e.Start.minute)
    end = dt.datetime(e.End.year, e.End.month, e.End.day, e.End.hour, e.End.minute)
    rp = e.GetRecurrencePattern()
    rp.RecurrenceType = RECURRENCE_TYPES[k]
    if k == "weekly":
        mask = 0
        for g in [x.strip().lower()[:3] for x in (weekdays or "").replace(";", ",").split(",") if x.strip()]:
            if g not in WEEKDAYS:
                raise ValueError("Invalid weekdays: use mon, tue, wed, thu, fri, sat, sun separated by commas.")
            mask |= WEEKDAYS[g]
        if mask == 0:
            mask = [2, 4, 8, 16, 32, 64, 1][start.weekday()]  # the start day
        rp.DayOfWeekMask = mask
    elif k == "monthly":
        rp.DayOfMonth = start.day
    elif k == "yearly":
        rp.MonthOfYear = start.month
        rp.DayOfMonth = start.day
    rp.Interval = max(1, int(interval or 1))
    rp.PatternStartDate = com_date(start.replace(hour=0, minute=0, second=0))
    rp.StartTime = com_date(start)
    rp.EndTime = com_date(end)
    rp.Duration = int((end - start).total_seconds() // 60)
    if recurrence_end:
        rp.PatternEndDate = com_date(parse_dt(recurrence_end).replace(hour=0, minute=0, second=0))
    elif occurrences:
        rp.Occurrences = int(occurrences)
    else:
        rp.NoEndDate = True
    try:
        series_end = None if rp.NoEndDate else fmt(rp.PatternEndDate)
    except Exception:  # noqa: BLE001
        series_end = None
    return {"type": k, "interval": rp.Interval, "weekdays": weekdays or None,
            "end": series_end, "occurrences": (int(occurrences) if occurrences and not recurrence_end else None)}


@mcp.tool(annotations=RW)
async def create_event(subject: str, start: str, end: str | None = None, duration_minutes: int | None = 30,
                       description: str = "", location: str = "", all_day: bool = False,
                       label: str | None = None, categories: str | None = None, reminder_minutes: int | None = 15,
                       mailbox: str | None = None, calendar_name: str | None = None,
                       recurrence: str | None = None, interval: int | None = None, weekdays: str | None = None,
                       recurrence_end: str | None = None, occurrences: int | None = None) -> dict:
    """Creates an event or a memo in the calendar (default: Calendar of the default mailbox). Dates 'YYYY-MM-DD HH:MM'.
    Labels (colours, visible with servers that support them, e.g. Kerio): phone_call (yellow, for memos), business,
    important, personal, vacation, must_attend, travel, needs_preparation, birthday, anniversary, none.
    Categories separated by ';'. reminder_minutes = -1 for no reminder. Never sends invitations.
    SERIES (optional): recurrence = daily | weekly | monthly | yearly; interval = every how many days, weeks, months or
    years (default 1); weekdays for weekly, e.g. "mon,wed,fri" (default: the start day); recurrence_end = 'YYYY-MM-DD'
    or occurrences = number of times; with neither, the series never ends.
    Example: "every Friday, 5 times" = recurrence=weekly, weekdays=fri, occurrences=5."""
    require_phase(3)

    def f():
        store = store_from_alias(mailbox)
        cal = calendar(store, calendar_name)
        e = cal.Items.Add(1)  # olAppointmentItem: a plain appointment, never a meeting
        _apply_event(e, subject, start, end, duration_minutes, description, location, all_day,
                     label, categories, reminder_minutes)
        series = None
        if recurrence:
            series = _apply_recurrence(e, recurrence, interval, weekdays, recurrence_end, occurrences)
        try:
            e.MeetingStatus = 0  # olNonMeeting: no invitation, nothing to send
        except Exception:  # noqa: BLE001
            pass
        e.Save()
        try:
            check_plain_appointment(e)
            # Final check: the event must sit in the requested calendar (defence against providers that move things).
            try:
                if e.Parent.EntryID != cal.EntryID:
                    raise ValueError(f"The event was saved in '{e.Parent.FolderPath}' instead of calendar '{cal.FolderPath}': tell the user.")
            except ValueError:
                raise
            except Exception:  # noqa: BLE001
                pass
        except Exception as ex:
            # The item exists anyway: it must appear in the write log, with the reason it was refused.
            audit("create_event", mailbox=mailbox or DEFAULT_MAILBOX, calendar=cal.FolderPath, id=e.EntryID,
                  subject=subject, ok=False, error=str(ex))
            raise
        audit("create_event", mailbox=mailbox or DEFAULT_MAILBOX, calendar=cal.FolderPath, id=e.EntryID, subject=subject, series=series)
        out = event_meta(e, mailbox or DEFAULT_MAILBOX, cal.FolderPath)
        if series:
            out["series"] = series
        return out
    return await com(f)


def occurrence_of(e, occurrence_date: str):
    """Returns the single occurrence of a series that starts at the given date and time ('YYYY-MM-DD HH:MM', as the
    'start' field returned by list_events), through RecurrencePattern.GetOccurrence. Clear error when the event is not
    a series or no occurrence starts exactly then."""
    if not bool(e.IsRecurring):
        raise ValueError("The event is not a series: occurrence_date applies to recurring events only.")
    when = parse_dt(occurrence_date)
    rp = e.GetRecurrencePattern()
    try:
        return rp.GetOccurrence(com_date(when))
    except Exception as ex:  # noqa: BLE001
        raise ValueError(f"No occurrence of the series starts on {when.strftime('%Y-%m-%d at %H:%M')}: use the exact start reported by list_events. ({ex})") from ex


@mcp.tool(annotations=UPD)
async def update_event(id: str, subject: str | None = None, start: str | None = None, end: str | None = None,
                       duration_minutes: int | None = None, description: str | None = None, location: str | None = None,
                       all_day: bool | None = None, label: str | None = None, categories: str | None = None,
                       reminder_minutes: int | None = None, mailbox: str | None = None,
                       occurrence_date: str | None = None) -> dict:
    """Updates only the given fields of an event or memo (id from list_events). For a recurring event it updates the
    whole series, or a single occurrence when occurrence_date is given ('YYYY-MM-DD HH:MM', the start reported by
    list_events): that occurrence becomes an exception of the series. Never deletes."""
    require_phase(3)

    def f():
        e, _ = item_by_id(id, mailbox)
        check_class(e, 26, "an event")
        check_plain_appointment(e)
        target = occurrence_of(e, occurrence_date) if occurrence_date else e
        _apply_event(target, subject, start, end, duration_minutes, description, location, all_day,
                     label, categories, reminder_minutes)
        target.Save()
        audit("update_event", mailbox=mailbox or DEFAULT_MAILBOX, id=id, occurrence=occurrence_date)
        out = event_meta(target, mailbox or DEFAULT_MAILBOX)
        if occurrence_date:
            out["id"] = id
            out["updated_occurrence"] = occurrence_date
            out["note"] = "only this occurrence was updated (exception of the series); the id remains the series id"
        return out
    return await com(f)


# ---------------------------------------------------------------- phase 4: alerts (the only exception to "no sending")

ALERTS_FILE = LOGDIR / "alerts.json"


def reserve_alert_slot() -> None:
    """Counts the alerts of the last hour from ALERTS_FILE and records this one BEFORE it is sent. Called on the COM
    thread, which runs one job at a time: parallel calls cannot all pass the check. Persisted, so a restart of
    Claude desktop does not reset the hour."""
    now = time.time()
    try:
        times = [t for t in json.loads(ALERTS_FILE.read_text("utf-8")) if isinstance(t, (int, float)) and now - t < 3600]
    except Exception:  # noqa: BLE001
        times = []
    if len(times) >= ALERT_MAX_PER_HOUR:
        raise ValueError(f"Alert limit reached ({ALERT_MAX_PER_HOUR} per hour): nothing sent.")
    times.append(now)
    ALERTS_FILE.write_text(json.dumps(times), "utf-8")


PROP_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001F"  # PR_SMTP_ADDRESS of a resolved recipient


def recipient_smtp(rcp) -> str:
    """SMTP address of a resolved Recipient. After ResolveAll, a recipient known to Exchange (same tenant, or a GAL
    contact) carries an X500 address in .Address: the SMTP one is read from PR_SMTP_ADDRESS or the Exchange user."""
    try:
        if (rcp.AddressEntry.Type or "").upper() == "SMTP":
            return (rcp.Address or "").strip().lower()
    except Exception:  # noqa: BLE001
        pass
    for getter in (lambda: rcp.PropertyAccessor.GetProperty(PROP_SMTP_ADDRESS),
                   lambda: rcp.AddressEntry.GetExchangeUser().PrimarySmtpAddress,
                   lambda: rcp.AddressEntry.GetContact().Email1Address,
                   lambda: rcp.Address):
        try:
            v = (getter() or "").strip().lower()
            if "@" in v:
                return v
        except Exception:  # noqa: BLE001
            continue
    return ""


def build_alert(subject: str, body: str):
    """Builds the alert mail WITHOUT sending it and checks that it has exactly one recipient, the configured one.
    Shared by send_alert and by the test (which discards the item). Raises if the account or the recipient are not
    the configured ones."""
    acc = None
    for a in ns().Accounts:
        if (a.SmtpAddress or "").lower() == ALERT_SENDER:
            acc = a
            break
    if acc is None:
        raise RuntimeError(f"Sender account {ALERT_SENDER} not found in Outlook.")
    m = COM.outlook().CreateItem(0)  # olMailItem
    m.SendUsingAccount = acc
    m.To = ALERT_RECIPIENT
    m.Subject = ALERT_PREFIX + subject
    m.Body = body
    # Defence in depth: exactly one recipient, resolved, with the configured SMTP address.
    if not m.Recipients.ResolveAll() or m.Recipients.Count != 1:
        raise RuntimeError("Recipient not compliant (not resolved or not exactly one): nothing sent.")
    found = recipient_smtp(m.Recipients.Item(1))
    if found != ALERT_RECIPIENT.strip().lower():
        raise RuntimeError(f"Recipient not compliant ('{found}' instead of the configured one): nothing sent.")
    return m


@mcp.tool(annotations=OUT)
async def send_alert(subject: str, body: str) -> dict:
    """Sends an automatic alert (for example from a scheduled task) to the user themselves. Sender and recipient are
    fixed in the server configuration and are NOT parameters: this is the only sending tool and it cannot write to
    anyone else. Plain text. At most alert.max_per_hour alerts per hour (default 10)."""
    require_phase(4)
    if not ALERT_SENDER or not ALERT_RECIPIENT:
        raise ToolError('Alerts not configured: config.json lacks the "alert" section with "sender" and "recipient".')

    def f():
        reserve_alert_slot()
        m = build_alert(subject, body)
        final_subject = m.Subject  # after Send() the COM object is no longer readable: take the data first
        # Logged BEFORE sending: if the log cannot be written, nothing is sent (and a retry cannot send twice).
        audit("send_alert", subject=final_subject, to=ALERT_RECIPIENT, status="sending")
        m.Send()
        audit("send_alert", subject=final_subject, to=ALERT_RECIPIENT, status="sent")
        return {"sent": True, "from": ALERT_SENDER, "to": ALERT_RECIPIENT, "subject": final_subject}
    return await com(f)


# ---------------------------------------------------------------- contacts

@mcp.tool(annotations=RO)
async def search_contacts(text: str, mailbox: str | None = None, folder: str | None = None, limit: int = 30) -> list[dict]:
    """Searches people and companies in the address book by first name, last name, company, e-mail address or phone
    number. Without 'folder' it searches every contacts folder of the mailbox (Contacts and subfolders, other address
    books). Returns the main fields; use read_contact for the full card."""
    def f():
        store = store_from_alias(mailbox)
        folders = [folder_from_path(store, folder)] if folder else folders_of_type(store, 2)
        t = dasl_lit(text)
        # DASL filter on every useful field: one server-side pass instead of reading every contact.
        # If the provider does not support the filter (com_error), fall back to scanning the main fields.
        fields = ["urn:schemas:contacts:cn", "urn:schemas:contacts:fileas", "urn:schemas:contacts:givenName",
                  "urn:schemas:contacts:sn", "urn:schemas:contacts:o", "urn:schemas:contacts:email1",
                  "urn:schemas:contacts:email2", "urn:schemas:contacts:telephoneNumber", "urn:schemas:contacts:mobile",
                  "urn:schemas:contacts:homePhone"]
        flt = "@SQL=" + " OR ".join(f"\"{c}\" LIKE '%{t}%'" for c in fields)
        out: list[dict] = []
        lim = max(1, min(limit, 200))
        for fold in folders:
            try:
                items = fold.Items.Restrict(flt)
                for c in iterate(items, lim - len(out)):
                    if c.Class == 40:
                        out.append(contact_meta(c, mailbox or DEFAULT_MAILBOX))
            except pywintypes.com_error:
                tl = text.lower()
                for c in iterate(fold.Items, 2000):
                    try:
                        if c.Class != 40:
                            continue
                        blob = " ".join(str(getattr(c, a) or "") for a in ("FullName", "CompanyName", "Email1Address", "MobileTelephoneNumber", "BusinessTelephoneNumber"))
                        if tl in blob.lower():
                            out.append(contact_meta(c, mailbox or DEFAULT_MAILBOX))
                            if len(out) >= lim:
                                break
                    except Exception:  # noqa: BLE001
                        continue
            if len(out) >= lim:
                break
        return out
    return await com(f)


@mcp.tool(annotations=RO)
async def list_contacts(mailbox: str | None = None, folder: str | None = None, limit: int = 100) -> list[dict]:
    """Lists the contacts of an address book folder (default: Contacts; others by path, e.g. 'Contacts/<subfolder>',
    see list_folders), in alphabetical order."""
    def f():
        store = store_from_alias(mailbox)
        fold = folder_from_path(store, folder) if folder else default_folder(store, 10, "Contacts", "Contatti", "Kontakte", "Contactos")
        items = fold.Items
        items.Sort("[FileAs]")
        return [contact_meta(c, mailbox or DEFAULT_MAILBOX) for c in iterate(items, max(1, min(limit, 500))) if c.Class == 40]
    return await com(f)


@mcp.tool(annotations=RO)
async def read_contact(id: str, mailbox: str | None = None) -> dict:
    """Reads the full card of a contact (id from search_contacts/list_contacts)."""
    def f():
        c, _ = item_by_id(id, mailbox)
        check_class(c, 40, "a contact")
        return contact_meta(c, mailbox or DEFAULT_MAILBOX, full=True)
    return await com(f)


@mcp.tool(annotations=RW)
async def create_contact(first_name: str = "", last_name: str = "", company: str = "", job_title: str = "",
                         email: str = "", email2: str = "", business_phone: str = "", mobile_phone: str = "",
                         home_phone: str = "", business_address: str = "", website: str = "", notes: str = "",
                         categories: str = "", mailbox: str | None = None, folder: str | None = None) -> dict:
    """Creates a contact in the address book (default folder: Contacts of the default mailbox). At least one of
    first_name, last_name or company is required. Never deletes or edits existing contacts: use update_contact."""
    require_phase(2)

    def f():
        if not (first_name or last_name or company):
            raise ValueError("Give at least a first name, a last name or a company.")
        store = store_from_alias(mailbox)
        fold = folder_from_path(store, folder) if folder else default_folder(store, 10, "Contacts", "Contatti", "Kontakte", "Contactos")
        c = fold.Items.Add(2)  # olContactItem
        _apply_contact(c, first_name, last_name, company, job_title, email, email2, business_phone, mobile_phone,
                       home_phone, business_address, website, notes, categories)
        c.Save()
        audit("create_contact", mailbox=mailbox or DEFAULT_MAILBOX, folder=fold.FolderPath, id=c.EntryID)
        return contact_meta(c, mailbox or DEFAULT_MAILBOX, full=True)
    return await com(f)


@mcp.tool(annotations=UPD)
async def update_contact(id: str, first_name: str | None = None, last_name: str | None = None, company: str | None = None,
                         job_title: str | None = None, email: str | None = None, email2: str | None = None,
                         business_phone: str | None = None, mobile_phone: str | None = None, home_phone: str | None = None,
                         business_address: str | None = None, website: str | None = None, notes: str | None = None,
                         categories: str | None = None, mailbox: str | None = None) -> dict:
    """Updates only the given fields of an existing contact (id from search_contacts). Never deletes."""
    require_phase(2)

    def f():
        c, _ = item_by_id(id, mailbox)
        check_class(c, 40, "a contact")
        # Address changes are a classic fraud pattern ("our accounting address changed"): old and new values go to the
        # write log, so they can be checked and reverted.
        before = {"email": c.Email1Address, "email2": c.Email2Address, "business_phone": c.BusinessTelephoneNumber,
                  "mobile_phone": c.MobileTelephoneNumber}
        _apply_contact(c, first_name, last_name, company, job_title, email, email2, business_phone, mobile_phone,
                       home_phone, business_address, website, notes, categories)
        c.Save()
        changed = {k: {"old": v, "new": getattr(c, a)} for (k, v), a in zip(
            before.items(), ("Email1Address", "Email2Address", "BusinessTelephoneNumber", "MobileTelephoneNumber"))
            if v != getattr(c, a)}
        audit("update_contact", mailbox=mailbox or DEFAULT_MAILBOX, id=id, changed=changed,
              fields=[k for k, v in (("first_name", first_name), ("last_name", last_name), ("company", company),
                                     ("job_title", job_title), ("home_phone", home_phone),
                                     ("business_address", business_address), ("website", website), ("notes", notes),
                                     ("categories", categories)) if v is not None])
        return contact_meta(c, mailbox or DEFAULT_MAILBOX, full=True)
    return await com(f)


# ---------------------------------------------------------------- tasks

def tasks_folder(store, folder: str | None):
    return folder_from_path(store, folder) if folder else default_folder(store, 13, "Tasks", "Attività", "Aufgaben", "Tâches", "Tareas")


def notes_folder(store):
    return default_folder(store, 12, "Notes", "Note", "Notizen", "Notas")


@mcp.tool(annotations=RO)
async def list_tasks(mailbox: str | None = None, open_only: bool = True, limit: int = 100,
                     folder: str | None = None) -> list[dict]:
    """Lists the tasks (Tasks folder of the mailbox) by due date. open_only=false includes completed ones."""
    def f():
        store = store_from_alias(mailbox)
        fold = tasks_folder(store, folder)
        items = fold.Items
        items.Sort("[DueDate]")
        if open_only:
            items = items.Restrict("[Complete] = False")
        return [task_meta(t, mailbox or DEFAULT_MAILBOX) for t in iterate(items, max(1, min(limit, 500))) if t.Class == 48]
    return await com(f)


@mcp.tool(annotations=RO)
async def read_task(id: str, mailbox: str | None = None) -> dict:
    """Reads a full task (id from list_tasks)."""
    def f():
        t, _ = item_by_id(id, mailbox)
        check_class(t, 48, "a task")
        return task_meta(t, mailbox or DEFAULT_MAILBOX, full=True)
    return await com(f)


@mcp.tool(annotations=RW)
async def create_task(subject: str, due_date: str | None = None, start_date: str | None = None, description: str = "",
                      priority: str = "normal", reminder: str | None = None, categories: str = "",
                      mailbox: str | None = None, folder: str | None = None) -> dict:
    """Creates a task (Tasks folder of the default mailbox). Dates 'YYYY-MM-DD' or 'YYYY-MM-DD HH:MM';
    priority: low, normal, high; reminder = date and time of the reminder."""
    require_phase(2)

    def f():
        store = store_from_alias(mailbox)
        fold = tasks_folder(store, folder)
        t = fold.Items.Add(3)  # olTaskItem
        _apply_task(t, subject, due_date, start_date, description, priority, reminder, categories, None, None)
        t.Save()
        audit("create_task", mailbox=mailbox or DEFAULT_MAILBOX, folder=fold.FolderPath, id=t.EntryID, subject=subject)
        return task_meta(t, mailbox or DEFAULT_MAILBOX, full=True)
    return await com(f)


@mcp.tool(annotations=UPD)
async def update_task(id: str, subject: str | None = None, due_date: str | None = None, start_date: str | None = None,
                      description: str | None = None, priority: str | None = None, reminder: str | None = None,
                      categories: str | None = None, complete: bool | None = None, percent_complete: int | None = None,
                      mailbox: str | None = None) -> dict:
    """Updates only the given fields of a task; complete=true marks it done, false reopens it. reminder='' removes the
    reminder. Never deletes."""
    require_phase(2)

    def f():
        t, _ = item_by_id(id, mailbox)
        check_class(t, 48, "a task")
        _apply_task(t, subject, due_date, start_date, description, priority, reminder, categories, complete, percent_complete)
        t.Save()
        audit("update_task", mailbox=mailbox or DEFAULT_MAILBOX, id=id, complete=complete)
        return task_meta(t, mailbox or DEFAULT_MAILBOX, full=True)
    return await com(f)


# ---------------------------------------------------------------- notes

@mcp.tool(annotations=RO)
async def list_notes(mailbox: str | None = None, limit: int = 100) -> list[dict]:
    """Lists the notes (Notes folder of the mailbox), most recent first."""
    def f():
        store = store_from_alias(mailbox)
        fold = notes_folder(store)
        items = fold.Items
        items.Sort("[LastModificationTime]", True)
        return [note_meta(n, mailbox or DEFAULT_MAILBOX) for n in iterate(items, max(1, min(limit, 500))) if n.Class == 44]
    return await com(f)


@mcp.tool(annotations=RO)
async def read_note(id: str, mailbox: str | None = None) -> dict:
    """Reads the full text of a note (id from list_notes)."""
    def f():
        n, _ = item_by_id(id, mailbox)
        check_class(n, 44, "a note")
        return note_meta(n, mailbox or DEFAULT_MAILBOX, full=True)
    return await com(f)


@mcp.tool(annotations=RW)
async def create_note(text: str, categories: str = "", mailbox: str | None = None) -> dict:
    """Creates a note (the first line of the text becomes the title)."""
    require_phase(2)

    def f():
        store = store_from_alias(mailbox)
        fold = notes_folder(store)
        n = fold.Items.Add(5)  # olNoteItem
        n.Body = text
        if categories:
            n.Categories = categories
        n.Save()
        audit("create_note", mailbox=mailbox or DEFAULT_MAILBOX, id=n.EntryID)
        return note_meta(n, mailbox or DEFAULT_MAILBOX, full=True)
    return await com(f)


@mcp.tool(annotations=UPD)
async def update_note(id: str, text: str | None = None, categories: str | None = None, mailbox: str | None = None) -> dict:
    """Replaces the text and/or the categories of an existing note. Never deletes."""
    require_phase(2)

    def f():
        n, _ = item_by_id(id, mailbox)
        check_class(n, 44, "a note")
        if text is not None:
            n.Body = text
        if categories is not None:
            n.Categories = categories
        n.Save()
        audit("update_note", mailbox=mailbox or DEFAULT_MAILBOX, id=id)
        return note_meta(n, mailbox or DEFAULT_MAILBOX, full=True)
    return await com(f)


# ---------------------------------------------------------------- folders and moves

@mcp.tool(annotations=RW)
async def create_folder(name: str, parent_path: str = "", type: str = "mail", mailbox: str | None = None) -> dict:
    """Creates a subfolder (type: mail, calendar, contacts, tasks, notes) under 'parent_path' (empty = mailbox root).
    Never renames or deletes folders."""
    require_phase(2)

    def f():
        codes = {"mail": 6, "calendar": 9, "contacts": 10, "tasks": 13, "notes": 12}
        if type not in codes:
            raise ValueError("Invalid type: mail, calendar, contacts, tasks, notes.")
        store = store_from_alias(mailbox)
        parent = folder_from_path(store, parent_path)
        for sub in parent.Folders:
            if sub.Name.lower() == name.lower():
                raise ValueError(f"A folder '{sub.Name}' already exists in '{parent.FolderPath}'.")
        new = parent.Folders.Add(name, codes[type])
        audit("create_folder", mailbox=mailbox or DEFAULT_MAILBOX, folder=new.FolderPath)
        return {"path": new.FolderPath, "type": type}
    return await com(f)


@mcp.tool(annotations=UPD)
async def move_mail(id: str, destination_folder: str, mailbox: str | None = None,
                    destination_mailbox: str | None = None) -> dict:
    """Moves a mail to another folder (also in another mailbox). Refuses any destination that is a trash
    (Deleted Items, Trash...): this server does not delete here. The read/unread state is preserved."""
    require_phase(2)

    def f():
        item, store = item_by_id(id, mailbox)
        check_class(item, 43, "a mail")
        dest_store = store_from_alias(destination_mailbox) if destination_mailbox else store
        dest = folder_from_path(dest_store, destination_folder)
        if is_trash(dest_store, dest):
            raise ValueError("Destination refused: it is a trash folder; use delete_mail to delete.")
        if dest.DefaultItemType != 0:
            raise ValueError("The destination folder is not a mail folder.")
        was_unread = bool(item.UnRead)
        new = item.Move(dest)
        try:
            if bool(new.UnRead) != was_unread:
                new.UnRead = was_unread
                new.Save()
        except Exception:  # noqa: BLE001
            pass
        audit("move_mail", mailbox=mailbox or DEFAULT_MAILBOX, id=id, destination=dest.FolderPath, new_id=new.EntryID)
        return {"id": new.EntryID, "mailbox": destination_mailbox or mailbox or DEFAULT_MAILBOX, "folder": dest.FolderPath}
    return await com(f)


# ---------------------------------------------------------------- deletion (to the trash, recoverable)

# Tools flagged as destructive: Claude desktop may ask the user for confirmation before running them.
DEL = ToolAnnotations(read_only_hint=False, destructive_hint=True, idempotent_hint=False, open_world_hint=False)


def require_deletion() -> None:
    """Blocks the deletion tools when the switch in config.json is off (or absent)."""
    if not DELETION:
        raise ToolError('Deletion disabled: in config.json "deletion_enabled" is not true.')


def trash_of(store):
    """The mailbox trash: default folder 3 (olFolderDeletedItems) or, with providers that do not declare it (IMAP),
    a mail folder with one of the TRASH_NAMES in the first two levels."""
    try:
        return store.GetDefaultFolder(3)
    except Exception:  # noqa: BLE001
        pass

    def walk(fold, level):
        for sub in fold.Folders:
            try:
                if sub.Name.lower() in TRASH_NAMES and sub.DefaultItemType == 0:
                    return sub
            except Exception:  # noqa: BLE001
                continue
            if level < 2:
                r = walk(sub, level + 1)
                if r is not None:
                    return r
        return None

    f = walk(store.GetRootFolder(), 1)
    if f is None:
        raise ValueError("This mailbox has no recognisable trash folder: deletion refused.")
    return f


def check_subject(item, expected_subject: str) -> str:
    """Before deleting, the real subject must match the one the user confirmed: protects against a wrong id or an
    item moved in the meantime."""
    real = (item.Subject or "").strip()
    if real != (expected_subject or "").strip():
        raise ValueError(f"Deletion refused: the real subject is \"{real}\", different from the confirmed \"{expected_subject}\".")
    return real


@mcp.tool(annotations=DEL)
async def delete_mail(id: str, expected_subject: str, mailbox: str | None = None) -> dict:
    """Moves a mail to the trash of its mailbox (recoverable). USE ONLY after the user has confirmed that precise mail
    in chat. 'expected_subject' must match the real subject, otherwise it refuses.
    Refuses mail already in the trash (no permanent deletion)."""
    require_phase(2)
    require_deletion()

    def f():
        item, store = item_by_id(id, mailbox)
        check_class(item, 43, "a mail")
        subj = check_subject(item, expected_subject)
        if is_trash(store, item.Parent):
            raise ValueError("The mail is already in the trash: permanent deletion is not allowed.")
        dest = trash_of(store)
        origin = item.Parent.FolderPath
        new = item.Move(dest)
        audit("delete_mail", mailbox=mailbox or DEFAULT_MAILBOX, id=id, subject=subj, from_folder=origin, to_folder=dest.FolderPath)
        return {"deleted": True, "subject": subj, "from_folder": origin, "trash": dest.FolderPath, "id_in_trash": new.EntryID}
    return await com(f)


@mcp.tool(annotations=DEL)
async def delete_event(id: str, expected_subject: str, whole_series: bool = False, mailbox: str | None = None,
                       occurrence_date: str | None = None) -> dict:
    """Moves a calendar event or memo to the trash of its mailbox (recoverable). USE ONLY after the user has confirmed
    that precise event in chat. 'expected_subject' must match the real subject.
    For a recurring event: whole_series=true deletes the whole series (to the trash); occurrence_date
    ('YYYY-MM-DD HH:MM', the start reported by list_events) deletes ONLY that occurrence. WARNING: a single occurrence
    does not go through the trash (it becomes an exception of the series) and can only be recovered by recreating it:
    tell the user before asking for confirmation. Refuses meetings with attendees (deleting them would send
    cancellations) and events already in the trash."""
    require_phase(3)
    require_deletion()

    def f():
        e, store = item_by_id(id, mailbox)
        check_class(e, 26, "an event")
        subj = check_subject(e, expected_subject)
        if whole_series and occurrence_date:
            raise ValueError("Give either whole_series=true or occurrence_date, not both.")
        try:
            status = int(e.MeetingStatus)
        except Exception:  # noqa: BLE001
            status = 0
        if status != 0:
            raise ValueError("It is a meeting with attendees: deleting it could send cancellations. The user deletes it by hand in Outlook.")
        if is_trash(store, e.Parent):
            raise ValueError("The event is already in the trash: permanent deletion is not allowed.")
        if occurrence_date:
            occ = occurrence_of(e, occurrence_date)
            occ_start = fmt(occ.Start)
            occ.Delete()
            audit("delete_event", mailbox=mailbox or DEFAULT_MAILBOX, id=id, subject=subj, occurrence=occ_start,
                  whole_series=False, note="single occurrence, no trash")
            return {"deleted": True, "subject": subj, "occurrence": occ_start, "whole_series": False,
                    "note": "only this occurrence was deleted: the series continues; it does not go through the trash and can only be recovered by recreating it"}
        if bool(e.IsRecurring) and not whole_series:
            raise ValueError("Recurring event: give whole_series=true (the whole series disappears) or occurrence_date (only that occurrence).")
        start = fmt(e.Start)
        dest = trash_of(store)
        origin = e.Parent.FolderPath
        # Delete(), not Move(): the same operation as the Delete key in Outlook. Verified in the field that with some
        # providers (Kerio Connect via KOFF) an appointment moved with Move() into Deleted Items, which is a mail
        # folder, is converted into a message with a .msg attachment and broken accents; with Delete() it remains a
        # restorable calendar item.
        e.Delete()
        # Find the event in the trash to return its new id (handy to restore or verify it).
        id_trash = None
        try:
            r = dest.Items.Restrict("@SQL=\"urn:schemas:httpmail:subject\" = '" + dasl_lit(subj) + "'")
            for it in iterate(r, 50):
                # same subject AND same start: an older event with the same subject is not the one just deleted
                if it.Class == 26 and fmt(it.Start) == start:
                    id_trash = it.EntryID
                    break
        except Exception:  # noqa: BLE001
            pass
        audit("delete_event", mailbox=mailbox or DEFAULT_MAILBOX, id=id, subject=subj, start=start,
              whole_series=bool(whole_series), from_folder=origin, to_folder=dest.FolderPath, id_in_trash=id_trash)
        return {"deleted": True, "subject": subj, "start": start, "whole_series": bool(whole_series),
                "from_folder": origin, "trash": dest.FolderPath, "id_in_trash": id_trash,
                "note": "moved to the trash with Delete, like the Delete key in Outlook: it remains a restorable calendar item"}
    return await com(f)


if __name__ == "__main__":
    # Started by Claude desktop over stdio: the server reads MCP requests from standard input and answers on standard
    # output. That is why logs go to a file and to stderr, never to stdout.
    log.info("Starting outlook-classic-mcp %s, phase %s, timeout %s s, deletion %s", __version__, PHASE, TIMEOUT, DELETION)
    mcp.run(transport="stdio")
