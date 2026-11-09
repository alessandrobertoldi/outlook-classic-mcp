# test_accents.py - automatic test of outlook-classic-mcp, with accented letters.
#
# WHEN: after installing and after every change to server.py, before restarting Claude desktop. It runs against the
# code on disk (it imports server.py directly), so Claude desktop need not be open; Outlook must be open.
# HOW (from the server folder, in PowerShell):
#   .\.venv\Scripts\python.exe .\test_accents.py            runs the test and cleans up
#   .\.venv\Scripts\python.exe .\test_accents.py --keep     leaves the created items in place, to look at them in Outlook
#
# WHAT IT DOES: creates drafts (text, HTML, reply to the latest received mail, with attachment), saves a mail as .msg
# and its attachment with SHA-256 fingerprints, reads transport headers, creates an event and deletes it (checking that
# a calendar item remains in the trash), creates a weekly series of 3 and updates and deletes single occurrences, then
# the whole series, creates a note and a contact. Every text is read back through the server's own tools and must come
# back identical, without the typical signs of broken accents ("Ã", "â€", "Â", "�"). Every item carries
# "ACCENT TEST" in its title. At the end it DELETES ONLY the items it created itself (recognised by id), moving them to
# the trash. It sends nothing: send_alert is not part of the test.
#
# REQUIREMENTS: config.json filled in, classic Outlook open as a normal user, at least one mail received in the last 7
# days in the default mailbox, "phase": 3 or 4 (the test writes drafts and events) and "deletion_enabled": true for the
# deletion checks (otherwise they are skipped and reported as such). Test files go into a subfolder
# "accent_test_<date>" of the FIRST folder in "save_folders", created and then removed by the test.
# RESULT: the last line must be "RESULT: ALL OK".

import asyncio
import datetime as dt
import hashlib
import shutil
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import server  # noqa: E402

TEXT = "Accent test: già servirà così perché più città è caffè àèéìòù ÀÈÉÌÒÙ ñ ü ß € the typographic apostrophe ’ and “curly quotes”."
SIGNS = ("Ã", "â€", "Â", "�")
KEEP = "--keep" in sys.argv
created: list[tuple[str, str]] = []  # (description, entry_id)
results: list[tuple[str, bool, str]] = []
# Dummy recipient of the drafts (they are never sent) and test folder, inside the first allowed folder.
TEST_RECIPIENT = "accent.test@example.com"
TEST_FOLDER = server.SAVE_FOLDERS[0] / dt.datetime.now().strftime("accent_test_%Y%m%d_%H%M%S")


def sha(p: Path) -> str:
    return hashlib.sha256(p.read_bytes()).hexdigest()


def check_fingerprint(name: str, file: str, digest: str, digest_file: str) -> None:
    f, fd = Path(file), Path(digest_file)
    ok = f.is_file() and fd.is_file() and sha(f) == digest
    if ok:
        lines = fd.read_text("utf-8").splitlines()
        ok = lines[-1] == f"{digest}  {f.name}" and any(l.startswith("# acquired: ") for l in lines)
    results.append((name, ok, "" if ok else f"file={file} digest={digest} sha256_file={digest_file}"))


async def draft_state(eid: str) -> tuple[int, str]:
    """Code page and HTML Outlook really stored in the draft (reading the text back is not enough: a text draft once
    looked fine through read_mail and still had corrupted accents inside Outlook)."""
    def f():
        it = server.ns().GetItemFromID(eid)
        return int(it.InternetCodepage), it.HTMLBody or ""
    return await server.com(f)


async def check_saved_draft(name: str, eid: str) -> None:
    cp, html = await draft_state(eid)
    ok = cp == server.UTF8_CPID and not any(sg in html for sg in SIGNS) and "citt" in html
    results.append((name + ": UTF-8 and HTML without signs", ok, "" if ok else f"InternetCodepage={cp} html={html[:200]}"))


def check(name: str, got: str, expected: str = TEXT) -> None:
    ok = expected in (got or "") and not any(sg in (got or "") for sg in SIGNS)
    detail = "" if ok else (got or "")[:300].replace("\r", " ").replace("\n", " ")
    results.append((name, ok, detail))


async def main() -> None:
    status = await server.outlook_status()
    print("Outlook", status["outlook_version"], "phase", status["phase"])

    # 1. text draft
    b = await server.create_draft(subject="ACCENT TEST text draft", body=TEXT, to=TEST_RECIPIENT)
    created.append(("text draft", b["id"]))
    r = await server.read_mail(b["id"], format="text")
    check("create_draft (text)", r["body"])
    await check_saved_draft("create_draft (text)", b["id"])

    # 2. HTML draft
    b = await server.create_draft(subject="ACCENT TEST html draft", body=server.text_to_html(TEXT), to=TEST_RECIPIENT, format="html")
    created.append(("html draft", b["id"]))
    r = await server.read_mail(b["id"], format="text")
    check("create_draft (html)", r["body"])
    await check_saved_draft("create_draft (html)", b["id"])

    # 3. reply draft to the latest mail in the Inbox
    mails = await server.list_mail(days=7, limit=1)
    if mails:
        orig = mails[0]
        orig_text = (await server.read_mail(orig["id"], format="text"))["body"]
        b = await server.create_reply_draft(orig["id"], body=TEXT)
        created.append(("reply draft", b["id"]))
        r = await server.read_mail(b["id"], format="text")
        check("create_reply_draft (new text)", r["body"])
        # the quoted original must not contain encoding signs (unless the original already had them)
        signs_orig = any(sg in orig_text for sg in SIGNS)
        signs_reply = any(sg in (r["body"] or "") for sg in SIGNS)
        results.append(("create_reply_draft (quoted original)", signs_orig or not signs_reply,
                        "" if (signs_orig or not signs_reply) else (r["body"] or "")[:300]))
    else:
        results.append(("create_reply_draft", False, "no mail in the last 7 days"))

    # 3b. save_mail and save_attachments with fingerprints
    TEST_FOLDER.mkdir(parents=True, exist_ok=False)
    attachment = TEST_FOLDER / "source" / "attachment già più.txt"
    attachment.parent.mkdir()
    attachment.write_text(TEXT, "utf-8")
    subj = "ACCENT TEST save_mail già più: perché? <ok>"
    b = await server.create_draft(subject=subj, body=TEXT, to=TEST_RECIPIENT, attachments=[str(attachment)])
    created.append(("draft with attachment", b["id"]))
    await server.mark_mail(b["id"], read=False)  # make it unread: save_mail and save_attachments must leave it so
    dest = TEST_FOLDER / "saved"
    s1 = await server.save_mail(b["id"], str(dest))
    p1 = Path(s1["saved_to"])
    results.append(("save_mail (.msg file with cleaned name and accents)",
                    p1.suffix == ".msg" and "ACCENT TEST save_mail già più perché ok" in p1.name and p1.stat().st_size > 1000, p1.name))
    check_fingerprint("save_mail (SHA-256 fingerprint and .sha256 file)", s1["saved_to"], s1["sha256"], s1["sha256_file"])
    s2 = await server.save_mail(b["id"], str(dest))
    results.append(("save_mail (never overwrite: second distinct file)",
                    s2["saved_to"] != s1["saved_to"] and sha(p1) == s1["sha256"], s2["saved_to"]))
    r = await server.read_mail(b["id"])
    results.append(("save_mail (unread mail stays unread)", bool(r["unread"]) and bool(s1["unread"]), f"unread={r['unread']}"))
    a1 = await server.save_attachments(b["id"], str(dest))
    ok_a = len(a1) == 1 and Path(a1[0]["saved_to"]).read_text("utf-8") == TEXT
    results.append(("save_attachments (identical attachment)", ok_a, str(a1)[:200]))
    if a1:
        check_fingerprint("save_attachments (SHA-256 fingerprint and .sha256 file)", a1[0]["saved_to"], a1[0]["sha256"], a1[0]["sha256_file"])
    r = await server.read_mail(b["id"])
    results.append(("save_attachments (unread mail stays unread)", bool(r["unread"]), f"unread={r['unread']}"))
    # a folder OUTSIDE the allowed ones must be refused (the Windows temp folder, unless it is allowed)
    outside = Path(tempfile.gettempdir()) / "accent_test_outside"
    if server.path_allowed(outside):
        results.append(("save_mail (folder not allowed refused)", True, "skipped: the temp folder is among the allowed ones"))
    else:
        try:
            await server.save_mail(b["id"], str(outside))
            results.append(("save_mail (folder not allowed refused)", False, f"saved into {outside}"))
        except Exception:  # noqa: BLE001
            results.append(("save_mail (folder not allowed refused)", True, ""))

    # 3c. transport headers and save_mail on a real received mail (unread, if there is one)
    real = await server.list_mail(days=7, limit=1, unread_only=True) or await server.list_mail(days=7, limit=1)
    if real:
        v = real[0]
        r = await server.read_mail(v["id"], headers=True, max_chars=200)
        hdr = r.get("transport_headers") or ""
        results.append(("read_mail headers=true (Received present)", "received:" in hdr.lower(), r.get("headers_note", "") or hdr[:200]))
        results.append(("read_mail headers=true (unread mail stays unread)", r["unread"] == v["unread"], f"before={v['unread']} after={r['unread']}"))
        sv = await server.save_mail(v["id"], str(dest))
        after = await server.read_mail(v["id"], max_chars=10)
        check_fingerprint("save_mail on received mail (fingerprint)", sv["saved_to"], sv["sha256"], sv["sha256_file"])
        results.append(("save_mail on received mail (read/unread state unchanged"
                        + (", it was UNREAD)" if v["unread"] else ", it was already read: partial check)"),
                        after["unread"] == v["unread"], f"before={v['unread']} after={after['unread']}"))
    else:
        results.append(("headers and save_mail on received mail", False, "no mail in the last 7 days"))

    # 4. event
    start = (dt.datetime.now() + dt.timedelta(days=1)).replace(hour=15, minute=0, second=0, microsecond=0)
    e = await server.create_event(subject="ACCENT TEST event già così", start=start.strftime("%Y-%m-%d %H:%M"),
                                  duration_minutes=15, description=TEXT, label="phone_call", reminder_minutes=-1)
    ev = await server.read_event(e["id"])
    check("create_event (description)", ev["description"])
    check("create_event (subject)", ev["subject"], "ACCENT TEST event già così")
    # deletion with delete_event: an EVENT (class 26) with intact accents must remain in the trash
    if server.DELETION:
        res = await server.delete_event(e["id"], expected_subject="ACCENT TEST event già così")
        idt = res.get("id_in_trash")
        if idt:
            in_trash = await server.read_event(idt)
            check("delete_event (event in trash, description)", in_trash.get("description", ""))
            check("delete_event (event in trash, subject)", in_trash.get("subject", ""), "ACCENT TEST event già così")
            results.append(("delete_event (still an event, not a message)", True, ""))
        else:
            results.append(("delete_event (event found in trash)", False, "no calendar item with that subject in the trash: maybe converted to a message"))
    else:
        created.append(("event", e["id"]))
        results.append(("delete_event", True, "skipped: deletion_enabled=false"))

    # 4b. recurring series: weekly, 3 times, then deletion of the whole series
    series_start = (dt.datetime.now() + dt.timedelta(days=2)).replace(hour=16, minute=0, second=0, microsecond=0)
    sr = await server.create_event(subject="ACCENT TEST series già così", start=series_start.strftime("%Y-%m-%d %H:%M"),
                                   duration_minutes=15, description=TEXT, label="phone_call", reminder_minutes=-1,
                                   recurrence="weekly", occurrences=3)
    results.append(("create_event (series: recurring)", bool(sr.get("recurring")), str(sr)[:200]))
    window = dict(from_date=series_start.strftime("%Y-%m-%d"), to_date=(series_start + dt.timedelta(days=30)).strftime("%Y-%m-%d"), limit=300)
    listing = await server.list_events(**window)
    occ = [x for x in listing if (x.get("subject") or "") == "ACCENT TEST series già così"]
    results.append(("create_event (series: 3 occurrences in the calendar)", len(occ) == 3, f"found {len(occ)} occurrences"))
    if occ:
        check("create_event (series: subject with accents)", occ[0]["subject"], "ACCENT TEST series già così")
    # 4c. single occurrence: update the second, delete the third, the others stay
    if len(occ) == 3:
        ordered = sorted(occ, key=lambda x: x["start"])
        second, third = ordered[1]["start"], ordered[2]["start"]
        try:
            m2 = await server.update_event(sr["id"], location="Room già booked", occurrence_date=second)
            results.append(("update_event (single occurrence)", m2.get("updated_occurrence") == second, str(m2)[:200]))
        except Exception as ex:  # noqa: BLE001
            results.append(("update_event (single occurrence)", False, str(ex)[:200]))
        if server.DELETION:
            try:
                d3 = await server.delete_event(sr["id"], expected_subject="ACCENT TEST series già così", occurrence_date=third)
                results.append(("delete_event (single occurrence)", d3.get("occurrence") == third, str(d3)[:200]))
            except Exception as ex:  # noqa: BLE001
                results.append(("delete_event (single occurrence)", False, str(ex)[:200]))
            after1 = await server.list_events(**window)
            left1 = sorted([x for x in after1 if (x.get("subject") or "") == "ACCENT TEST series già così"], key=lambda x: x["start"])
            results.append(("delete_event (single occurrence: 2 of 3 remain)", len(left1) == 2, f"remaining {len(left1)}"))
            loc2 = left1[1].get("location") if len(left1) == 2 else None
            results.append(("update_event (updated occurrence kept, location with accent)", loc2 == "Room già booked", f"location={loc2}"))
    else:
        results.append(("single occurrences", False, "skipped: the series did not have 3 occurrences"))
    if server.DELETION:
        await server.delete_event(sr["id"], expected_subject="ACCENT TEST series già così", whole_series=True)
        after = await server.list_events(**window)
        left = [x for x in after if (x.get("subject") or "") == "ACCENT TEST series già così"]
        results.append(("delete_event (whole series gone from the calendar)", len(left) == 0, f"remaining {len(left)}"))
    else:
        created.append(("series", sr["id"]))

    # 5. note
    n = await server.create_note(text="ACCENT TEST note\n" + TEXT)
    created.append(("note", n["id"]))
    nt = await server.read_note(n["id"])
    check("create_note", nt["text"])

    # 6. contact
    c = await server.create_contact(first_name="Accent", last_name="Test già così", company="Società è più", notes=TEXT,
                                    email="accent.test@example.com")
    created.append(("contact", c["id"]))
    ct = await server.read_contact(c["id"])
    check("create_contact (notes)", ct.get("notes", ""))
    check("create_contact (company)", ct.get("company", ""), "Società è più")


def cleanup() -> None:
    if KEEP:
        print("Test items left in Outlook (--keep):", ", ".join(f"{d}" for d, _ in created))
        print("Test files left in:", TEST_FOLDER)
        return
    if TEST_FOLDER.is_dir() and TEST_FOLDER.name.startswith("accent_test_"):
        shutil.rmtree(TEST_FOLDER)  # only the folder created by this test
        print("Test file folder removed:", TEST_FOLDER)

    def f():
        ns = server.ns()
        n = 0
        for desc, eid in created:
            try:
                it = ns.GetItemFromID(eid)
                subj = getattr(it, "Subject", "") or getattr(it, "FullName", "")
                if "ACCENT TEST" in str(subj) or "Accent Test" in str(subj):
                    it.Delete()  # goes to the trash, recoverable
                    n += 1
            except Exception as ex:  # noqa: BLE001
                print(f"  cleanup {desc}: failed ({ex})")
        return n
    print(f"Test items moved to the trash: {server.COM.run(f)}")


if __name__ == "__main__":
    try:
        asyncio.run(main())
    finally:
        cleanup()
    all_ok = True
    for name, ok, detail in results:
        all_ok &= ok
        print(("OK   " if ok else "FAIL ") + name + ("" if ok else f"  -> {detail}"))
    print("RESULT: ALL OK" if all_ok else "RESULT: PROBLEMS FOUND")
    sys.exit(0 if all_ok else 1)
