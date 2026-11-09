# test_unit.py - unit tests that need NO Windows, NO Outlook and NO network: they run anywhere (CI included).
#
# HOW:  python -m unittest -v test_unit            (from the repository folder; needs only the "mcp" package)
#
# WHAT: the Windows COM modules (pythoncom, pywintypes, win32com) are replaced by stubs and server.py is imported with a
# temporary config.json, so the pure parts of the server are tested for real: path allow-list (traversal, UNC, device
# paths, alternate data streams, reserved names), file-name cleaning (direction marks, trailing dots, device names),
# atomic file-name reservation, HTML reference filter for drafts, dangerous-attachment neutralisation, date parsing and
# DASL formatting, the COM worker (timeouts, cancellation, busy guard), the alert rate limit, the mailbox allow-list
# and the tool inventory (count, annotations, phases). test_accents.py remains the end-to-end test on a real Outlook.

import asyncio
import datetime as dt
import importlib
import json
import os
import sys
import tempfile
import threading
import time
import types
import unittest
from pathlib import Path

HERE = Path(__file__).resolve().parent


def _stub_windows_modules() -> None:
    """Minimal stand-ins for the pywin32 modules, enough to import server.py outside Windows."""
    if sys.platform == "win32":
        return
    pythoncom = types.ModuleType("pythoncom")

    class com_error(Exception):
        pass

    pythoncom.com_error = com_error
    pythoncom.CoInitialize = lambda: None
    pywintypes = types.ModuleType("pywintypes")
    pywintypes.com_error = com_error
    pywintypes.Time = lambda d: d
    win32com = types.ModuleType("win32com")
    client = types.ModuleType("win32com.client")

    def _no_outlook(_name):
        raise com_error("no Outlook in the unit tests")

    client.GetActiveObject = _no_outlook
    win32com.client = client
    sys.modules.update({"pythoncom": pythoncom, "pywintypes": pywintypes, "win32com": win32com, "win32com.client": client})


_TMP = Path(tempfile.mkdtemp(prefix="outlook_classic_mcp_unit_"))
SAVE_ROOT = _TMP / "allowed" / "Claude outputs"
ATTACH_ROOT = _TMP / "allowed" / "to-send"
for _p in (SAVE_ROOT, ATTACH_ROOT):
    _p.mkdir(parents=True)
CONFIG = {
    "phase": 4, "timeout_seconds": 1,
    "mailboxes": {"work": "Work Mailbox", "archive": "Outlook Data File"}, "default_mailbox": "work",
    "save_folders": [str(SAVE_ROOT)], "attach_folders": [str(ATTACH_ROOT)],
    "deletion_enabled": True,
    "alert": {"sender": "me@example.com", "recipient": "alerts@example.com", "max_per_hour": 3},
}
_cfg = _TMP / "config.json"
_cfg.write_text(json.dumps(CONFIG), "utf-8")
os.environ["OUTLOOK_CLASSIC_MCP_CONFIG"] = str(_cfg)
os.environ["LOCALAPPDATA"] = str(_TMP / "localappdata")   # logs of the test go to the temp folder
_stub_windows_modules()
sys.path.insert(0, str(HERE))
server = importlib.import_module("server")

WIN = sys.platform == "win32"


def wpath(s: str) -> str:
    """Windows-style test paths are rewritten under the temp folder on POSIX, keeping the shape of the attack."""
    return s


class PathAllowList(unittest.TestCase):
    roots = [Path(r"F:\Claude-Cowork"), Path(r"C:\Users\x\Claude outputs")]

    def test_inside(self):
        for p in (r"F:\Claude-Cowork\A\b", r"f:/claude-cowork/a", r"F:\Claude-Cowork", r"C:\Users\x\Claude outputs\p"):
            self.assertTrue(server.lexically_inside(p, self.roots), p)

    def test_refused(self):
        for p in (r"F:\Claude-Cowork-evil\x", r"F:\Claude-Cowork\..\Windows", r"F:\Claude-Cowork\a\..\..\x",
                  r"\\192.0.2.1\share\x", r"//192.0.2.1/share", r"\\?\F:\Claude-Cowork\x", r"\\?\UNC\h\s",
                  r"\\.\F:\Claude-Cowork\x", r"F:\Claude-Cowork\x.txt:stream", r"F:\Claude-Cowork\CON\x",
                  r"F:\Claude-Cowork\nul.txt", r"Claude-Cowork\x", r"C:\Users\x\Desktop", "", "F:\\Claude-Cowork\\a\x00b",
                  r"C:Claude-Cowork\x", r"F:\Claude-Cowork \x"):
            self.assertFalse(server.lexically_inside(p, self.roots), p)

    def test_unc_root_only_when_configured(self):
        unc_root = [Path(r"\\nas\share\claude")]
        self.assertTrue(server.lexically_inside(r"\\nas\share\claude\x", unc_root))
        self.assertFalse(server.lexically_inside(r"\\evil\share\claude", unc_root))

    def test_empty_roots_allow_nothing(self):
        self.assertFalse(server.path_allowed(SAVE_ROOT / "x", []))

    def test_path_allowed_real_roots(self):
        self.assertTrue(server.path_allowed(SAVE_ROOT / "sub"))
        self.assertFalse(server.path_allowed(_TMP / "elsewhere"))
        self.assertTrue(server.path_allowed(ATTACH_ROOT / "a.pdf", server.ATTACH_FOLDERS))
        self.assertFalse(server.path_allowed(SAVE_ROOT / "a.pdf", server.ATTACH_FOLDERS))


class FileNames(unittest.TestCase):
    def test_safe_file_name(self):
        cases = {"report.pdf:x": "report.pdf_x", "invoice\u202efdp.exe": "invoicefdp.exe", "a/../../b.txt": "b.txt",
                 "..\\..\\evil.dll": "evil.dll", "trail. . ": "trail", "CON.txt": "_CON.txt", "": "attachment",
                 "..": "attachment", "normal \u00e8.pdf": "normal \u00e8.pdf"}
        for raw, expected in cases.items():
            self.assertEqual(server.safe_file_name(raw), expected, repr(raw))
        long = server.safe_file_name("x" * 300 + ".pdf")
        self.assertTrue(long.endswith(".pdf") and len(long) <= 150)

    def test_neutralise_name(self):
        for n in ("fattura.scf", "Fattura.URL", "a.lnk", "desktop.ini", "Desktop.INI", "x.pdf.sha256", "run.exe", "x.library-ms"):
            self.assertTrue(server.neutralise_name(n).endswith(".txt"), n)
        for n in ("doc.pdf", "img.png", "a.docm", "report.xlsx"):
            self.assertEqual(server.neutralise_name(n), n)

    def test_name_from_subject(self):
        self.assertEqual(server.name_from_subject('R: Pratica 1 / gi\u00e0 "pi\u00f9" <x>: ok?  ...'), "R Pratica 1 gi\u00e0 pi\u00f9 x ok")
        self.assertEqual(server.name_from_subject(""), "no subject")

    def test_free_name_is_atomic_and_never_reuses(self):
        d = _TMP / "free"
        d.mkdir()
        a = server.free_name(d, "x.txt")
        self.assertTrue(a.exists() and a.stat().st_size == 0, "the name must be reserved on disk")
        b = server.free_name(d, "x.txt")
        (d / "x (2).txt.sha256").touch()
        c = server.free_name(d, "x.txt")
        self.assertEqual([a.name, b.name, c.name], ["x.txt", "x (1).txt", "x (3).txt"])

    def test_fingerprint_file(self):
        d = _TMP / "fp"
        d.mkdir()
        f = d / "a.bin"
        f.write_bytes(b"ciao")
        fp = server.sha256_fingerprint(f, mail_subject="line1\nline2")
        self.assertEqual(fp["sha256"], "b133a0c0e9bee3be20163d2ad31d6248db292aa6dcb1ee087a2aa50e0fc75ae2")
        lines = Path(fp["sha256_file"]).read_text("utf-8").splitlines()
        self.assertEqual(lines[-1], f"{fp['sha256']}  a.bin")
        self.assertTrue(all(l.startswith("#") for l in lines[:-1]), "context lines are comments on one line each")
        with self.assertRaises(FileExistsError):
            server.sha256_fingerprint(f)


class HtmlDrafts(unittest.TestCase):
    def test_bad_references_refused(self):
        for h in ('<img src="\\\\evil\\x.png">', "<img src='//evil/x'>", '<a href="file:///C:/x">',
                  '<div style="background:url(\\\\evil\\a)">', '<img src="  \\\\evil\\x">', '<IMG SRC = "C:\\x\\y.png">',
                  '<img src="fi\x00le:///x">', '<a href="javascript:alert(1)">', '<img src="\\\\?\\UNC\\evil\\x">'):
            with self.assertRaises(ValueError, msg=h):
                server.check_html_refs(h)

    def test_good_html_accepted(self):
        for h in ('<p>ciao &agrave;</p><a href="https://example.com/x">x</a><img src="https://cdn/x.png">',
                  '<img src="cid:logo">', '<a href="mailto:a@b.it">', '<p>path in text: \\\\server\\share</p>'):
            server.check_html_refs(h)

    def test_text_to_html_escapes(self):
        self.assertIn("&lt;b&gt;", server.text_to_html("<b>\n"))
        self.assertIn("<br>", server.text_to_html("a\nb"))


class Dates(unittest.TestCase):
    def test_parse(self):
        self.assertEqual(server.parse_dt("2026-07-01 09:30"), dt.datetime(2026, 7, 1, 9, 30))
        self.assertEqual(server.parse_dt("2026-07-01", end_of_day=True), dt.datetime(2026, 7, 1, 23, 59, 59))
        self.assertIsNone(server.parse_dt(""))
        with self.assertRaises(ValueError):
            server.parse_dt("01/07/2026")

    def test_dasl_date_uses_the_rule_of_that_date(self):
        # Whatever the local zone, the conversion must equal the per-date astimezone(), not today's offset.
        d = dt.datetime(2026, 7, 1, 12, 0)
        self.assertEqual(server.dasl_date(d), d.astimezone(dt.timezone.utc).strftime("%Y-%m-%d %H:%M"))

    def test_dasl_lit(self):
        self.assertEqual(server.dasl_lit("O'Brien"), "O''Brien")


class ComWorkerBehaviour(unittest.TestCase):
    def test_timeout_cancels_queued_and_refuses_while_busy(self):
        w = server.ComWorker()
        ran, msgs = [], {}

        def blocker():
            try:
                w.run(lambda: time.sleep(3))
            except TimeoutError as e:
                msgs["blocker"] = str(e)

        t = threading.Thread(target=blocker)
        t.start()
        time.sleep(0.2)
        with self.assertRaises(TimeoutError) as c1:
            w.run(lambda: ran.append("late"))
        self.assertIn("cancelled", str(c1.exception))
        time.sleep(0.9)
        with self.assertRaises(TimeoutError) as c2:
            w.run(lambda: ran.append("busy"))
        self.assertIn("still busy", str(c2.exception))
        t.join()
        time.sleep(1.8)
        self.assertEqual(ran, [], "cancelled requests must never run afterwards")
        self.assertIn("MAY STILL COMPLETE", msgs["blocker"])
        self.assertEqual(w.run(lambda: 42), 42)


class AlertLimit(unittest.TestCase):
    def test_rate_limit_persisted(self):
        server.ALERTS_FILE = _TMP / "alerts.json"
        for _ in range(CONFIG["alert"]["max_per_hour"]):
            server.reserve_alert_slot()
        with self.assertRaises(ValueError):
            server.reserve_alert_slot()
        self.assertEqual(len(json.loads(server.ALERTS_FILE.read_text("utf-8"))), 3)


class Config(unittest.TestCase):
    def test_mailboxes_and_folders(self):
        self.assertEqual(server.DEFAULT_MAILBOX, "work")
        self.assertEqual(set(server.MAILBOXES), {"work", "archive"})
        self.assertEqual(server.ATTACH_FOLDERS, [ATTACH_ROOT])
        self.assertEqual(server.ALERT_MAX_PER_HOUR, 3)

    def test_store_lookup_refuses_unconfigured(self):
        # store_from_alias must refuse BEFORE touching Outlook (no COM in the tests)
        with self.assertRaises(ValueError) as c:
            server.store_from_alias("Shared Mailbox")
        self.assertIn("not configured", str(c.exception))


class ToolInventory(unittest.TestCase):
    def test_tools(self):
        tools = asyncio.run(server.mcp.list_tools())
        names = {t.name for t in tools}
        self.assertEqual(len(tools), 32, sorted(names))
        for n in ("outlook_status", "list_mail", "read_mail", "save_mail", "save_attachments", "create_draft",
                  "create_reply_draft", "send_alert", "delete_mail", "delete_event", "move_mail", "update_contact"):
            self.assertIn(n, names)
        ann = {t.name: t.annotations for t in tools}
        self.assertTrue(ann["read_mail"].read_only_hint)
        self.assertTrue(ann["delete_mail"].destructive_hint and ann["delete_event"].destructive_hint)
        self.assertTrue(ann["update_contact"].destructive_hint and ann["move_mail"].destructive_hint)
        self.assertTrue(ann["send_alert"].open_world_hint and ann["create_draft"].open_world_hint)
        self.assertFalse(ann["list_mail"].open_world_hint)

    def test_phase_gate(self):
        real = server.PHASE
        server.PHASE = 1
        try:
            for n in (2, 3, 4):
                with self.assertRaises(server.ToolError):
                    server.require_phase(n)
            server.require_phase(1)
        finally:
            server.PHASE = real


if __name__ == "__main__":
    unittest.main(verbosity=2)
