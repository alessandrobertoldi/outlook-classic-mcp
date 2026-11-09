# outlook-classic-mcp

A local MCP server that connects Claude (Claude desktop and Cowork) to **classic Outlook for Windows**: mail, calendar, contacts, tasks, notes and folders, through Outlook's COM object model. It uses the Outlook session already open on the PC: no passwords, no external server, no network port.

*Versione italiana: [README.it.md](README.it.md).*

## What it does

Once installed, in Claude desktop or Cowork you can say things like "any unread mail from the Smith account?", "put a phone call in my calendar on Tuesday at 10", "draft a reply to this last mail", "save the attachments of that invoice", "what is on my agenda this week", and Claude does it in your Outlook, with your accounts, without leaving the PC.

Tools exposed (32):

| Area | Tools |
|---|---|
| Status and folders | `outlook_status`, `list_folders`, `create_folder` |
| Mail | `list_mail`, `search_mail`, `read_mail` (optionally with transport headers for SPF, DKIM, DMARC), `save_attachments`, `save_mail` (.msg with SHA-256 fingerprint), `mark_mail`, `move_mail` |
| Drafts | `create_draft`, `create_reply_draft` |
| Calendar | `list_events`, `read_event`, `create_event` (recurring series too), `update_event` (whole series or a single occurrence) |
| Contacts | `search_contacts`, `list_contacts`, `read_contact`, `create_contact`, `update_contact` |
| Tasks | `list_tasks`, `read_task`, `create_task`, `update_task` |
| Notes | `list_notes`, `read_note`, `create_note`, `update_note` |
| Alerts | `send_alert` (the only sending tool, fixed recipient) |
| Deletion | `delete_mail`, `delete_event` (trash only, with subject confirmation) |

## Safety rules, enforced in code

The server is meant to be driven by an AI, so its limits are not recommendations in a prompt but checks in the code, which Claude cannot talk its way around:

- **It does not send mail.** Mail is born as a draft in the Drafts folder: you review it, sign it if you wish, and send it yourself. The one exception is `send_alert`, which writes only from the sender to the recipient fixed in `config.json` (typically from one address of yours to another): Claude chooses subject and body, never who receives it.
- **It never deletes permanently.** `delete_mail` and `delete_event` move items to the mailbox trash, only if the subject passed matches the real one, never from the trash itself, never meetings with attendees. They are off unless `"deletion_enabled": true`. Contacts, tasks, notes and folders cannot be deleted at all. One technical exception: deleting a single occurrence of a recurring series, which Outlook does not route through the trash (the tool says so in its description, so Claude tells the user before asking for confirmation).
- **Unread mail stays unread**, even after Claude has read it, saved its attachments, drafted a reply or moved it.
- **It writes files only into the folders you listed** in `save_folders`, never overwriting, and leaves a SHA-256 fingerprint with date and time next to every file (chain of custody).
- **No credentials.** It attaches to the Outlook already running; when Outlook is closed it answers with an error, it does not start it and does not ask for a password.
- **It does not touch rules, settings or accounts.**
- **Every write is logged** (`%LOCALAPPDATA%\outlook-classic-mcp\audit.log`: tool, mailbox, id, subject; never the content).
- **Progressive phases.** With `"phase": 1` the server only reads; 2 adds drafts and non-destructive writes; 3 the calendar; 4 the alerts. Start at 1 and move up as trust grows.

## Requirements

- Windows 10 or 11 with **classic Outlook** (the Outlook of Microsoft 365 or Office 2016 and later, process `OUTLOOK.EXE`). The "new Outlook" (`olk.exe`) exposes no COM interface and does not work.
- Outlook running as a normal user (not "as administrator": processes with different privileges cannot see each other).
- **Python 3.11 or newer** from [python.org](https://www.python.org/downloads/windows/), or [uv](https://docs.astral.sh/uv/).
- **Claude desktop** for Windows, with Cowork if you want to use it from there.

It works with any account configured in Outlook: Exchange and Microsoft 365, IMAP, POP, .pst data files, and also servers such as Kerio Connect through its KOFF connector (the environment it was born in and is verified in every day). Outlook's language does not matter: default folders are reached by their internal type, with name lookups in English, Italian, German, French and Spanish as a fallback.

## Installation

1. Download the repository (Code > Download ZIP, or `git clone`) into a stable folder, for example `C:\Users\<your user>\Claude\outlook-classic-mcp`. Do not move it after installing: its path goes into Claude desktop's configuration.

2. Copy `config.example.json` to `config.json` and fill in the values in `< >` (every key is explained below). Above all you need the exact names of your Outlook stores: read them in Outlook's folder pane (the bold name at the top of each account), or let the installer print them in the next step.

3. Open Outlook. Then, in **Windows PowerShell** (normal user), from the server folder:

   ```powershell
   powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\install.ps1"
   ```

   The script creates an isolated Python environment (`.venv`), installs `mcp` and `pywin32`, tests the connection to Outlook, test-starts the server and registers the `outlook` entry in Claude desktop's `claude_desktop_config.json`, keeping any other MCP server you already have.

4. Quit Claude desktop **completely** (tray icon > Quit) and open it again. The server appears among the local connectors. Try: "check the Outlook status".

5. Recommended: run the automatic test, which exercises every writing tool with accented letters and cleans up after itself:

   ```powershell
   .\.venv\Scripts\python.exe .\test_accents.py
   ```

   It must end with `RESULT: ALL OK`. (It needs `"phase": 3` or `4`; with `"deletion_enabled": false` the deletion checks are skipped.)

To update: replace the files, run `install.ps1` again if the dependencies changed, and restart Claude desktop completely. The server code is loaded when the app starts.

## Configuration (`config.json`)

| Key | Required | Meaning |
|---|---|---|
| `phase` | no (1) | 1 read only, 2 also drafts, contacts, tasks, notes, folders, moves and mail deletion, 3 also calendar, 4 also `send_alert` |
| `timeout_seconds` | no (40) | maximum wait for one call into Outlook; beyond it, an explained error ("a modal window is probably open in Outlook") |
| `max_body_chars` | no (20000) | maximum length of texts returned in one read |
| `mailboxes` | **yes** | short alias → exact store name in Outlook; stores not listed stay invisible |
| `default_mailbox` | **yes** | alias used when Claude names none |
| `save_folders` | **yes** | folders (and subfolders) the server may write files to |
| `deletion_enabled` | no (false) | enables `delete_mail` and `delete_event` |
| `alert.sender`, `alert.recipient`, `alert.subject_prefix` | no | sender (an Outlook account) and only recipient of `send_alert`; without the section, sending is always refused |
| `courtesy_pause_ms`, `pause_every_items` | no (10, 3) | pauses during long reads so that the Outlook window keeps breathing; re-read on the fly without a restart |

`config.json` is personal and excluded from the repository through `.gitignore`: do not publish it.

## How it works with Claude

- Claude sees every tool with its description (the docstring in `server.py`) and the mailbox aliases in the server instructions. Listings return metadata only (id, subject, sender, date); content is read separately, so conversations stay light.
- Ids are Outlook `EntryID`s: Claude receives them from listings and passes them to the read and update tools. A moved item changes id.
- Dates are written `YYYY-MM-DD` or `YYYY-MM-DD HH:MM`, local time.
- For recurring series, `list_events` returns one row per occurrence carrying the series id; `update_event` and `delete_event` act on the whole series, or on one occurrence when they receive `occurrence_date` (the exact start reported by the listing).
- The coloured labels of appointments (`phone_call`, `business`, `important`...) are Outlook's old "Label" property: some servers (Kerio Connect, for instance) show them as colours, others ignore them harmlessly.

A tip for Cowork users with scheduled tasks: keep Outlook open, because the server does not start it; when a task must warn you about something, `send_alert` is made for that.

## Technical lessons (Outlook COM from Python)

Collected during development, useful to whoever wants to change the code.

- **One COM thread.** Outlook objects must be used from the thread that created them (STA). The server is asynchronous, so every call goes through a dedicated thread with a queue and a timeout (`ComWorker`). A frozen Outlook (modal window) produces an explained error, not a hung server.
- **Date filters in DASL with ISO dates in UTC** (`@SQL="urn:schemas:httpmail:datereceived" >= '2026-11-09 15:00'`). The classic Jet format (`[ReceivedTime] >= 'MM/dd/yyyy'`) is parsed with the Windows regional settings and on a non-US system swaps day and month.
- **Written dates marked as UTC.** pywin32 converts a naive `datetime` from local time to UTC, shifting it; passed with `tzinfo=UTC` it reaches Outlook unchanged and is read as local time (`com_date`).
- **Recurrences:** `Items.Sort("[Start]")`, `IncludeRecurrences = True`, then `Restrict` with the DASL filter, and iterate with the enumerator (never `Count` or index). A series is built with `GetRecurrencePattern()` after `Start`, `End` and `Duration` are set, in the order RecurrenceType, DayOfWeekMask or DayOfMonth/MonthOfYear, Interval, PatternStartDate, StartTime, EndTime, Duration, and the end last. A single occurrence is reached with `GetOccurrence(start_date)`.
- **Accents in drafts:** set `InternetCodepage = 65001` before writing `HTMLBody`, and always write HTML (even for "plain text"): with `Body` the draft stays in the profile's code page and accents may get corrupted. In replies, read the original's HTML **before** changing the code page, remove the `<meta charset>` tags and insert the new text right after `<body>`.
- **Delete an event with `Delete()`, not with `Move()` to the trash.** Deleted Items is a mail folder: some providers convert the moved appointment into a message with a .msg attachment (and broken accents). `Delete()` is what the Delete key does and leaves a restorable event.
- **After `Send()` the COM object is no longer readable:** take the data to return before.
- **Courtesy pauses.** Every call runs on Outlook's main thread: during a read of hundreds of items the UI stays frozen for the user. A pause of a few milliseconds every few items reduces the annoyance a lot; it does not remove it, because two Outlook instances on the same profile do not exist.
- **Talking errors.** Every exception is converted into a `ToolError` with Outlook's own message (`excepinfo[2]`): Claude receives "Mailbox 'x' not found" instead of "Error executing tool", and can explain it to the user.

## Known limits

- The PC must be on, the user logged into Windows and classic Outlook open, not elevated.
- Encrypted mail (PGP, S/MIME) arrives encrypted: only Outlook decrypts it, in its own window.
- Body search (`search_mail` with `in_body=true`) is slow on folders with tens of thousands of messages: that is why it requires a date range.
- Deleting a single occurrence of a series does not go through the trash (Outlook behaviour).
- Outlook and Claude desktop must run with the same privileges (both as a normal user).

## Repository layout

| File | What it does |
|---|---|
| `server.py` | the MCP server: every tool and every constraint, commented throughout |
| `config.example.json` | configuration template with an explanation of every key |
| `install.ps1` | creates `.venv`, installs the dependencies, tests COM, registers the server in Claude desktop |
| `test_accents.py` | automatic test of the writing tools with accented letters |
| `README.it.md` | this document in Italian |
| `.gitignore` | excludes `config.json`, `.venv`, logs and backup copies |

## License

MIT: see `LICENSE`. The software is provided as is; before giving an AI access to your mailbox, start from phase 1 and read `audit.log` during the first days.
