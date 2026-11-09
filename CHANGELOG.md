# Changelog

All notable changes to this project are documented here. The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/);
versions follow [Semantic Versioning](https://semver.org/).

## [1.0.0] - 2026-11-10

First tagged release, after two independent security reviews.

### Security
- Paths from the agent are checked as text before any disk or network access: UNC, `\\?\`, device paths, alternate
  data streams, reserved names and traversal are refused without Windows ever contacting a host.
- Draft attachments only from `attach_folders` (explicit; absent = no attachments); `.msg` and `.sha256` never attached.
- HTML drafts that reference a network, local or script URL are refused.
- Saved attachments: cleaned names (direction marks, trailing dots, reserved names), dangerous types neutralised with
  `.txt`, Mark of the Web, atomic file-name reservation (never overwrite, no race).
- Phase 1 is read-only in fact: saving files and changing the read flag need phase 2.
- Only the mailboxes listed in `config.json` are reachable, even by exact store name; other store names are not exposed;
  stores with duplicate names are refused.
- COM worker: a request that times out before starting is cancelled; new requests are refused while Outlook is stuck.
- `send_alert`: recipient checked on the resolved SMTP address, hourly limit persisted on disk, audit before sending.
- Unread flag restored in `finally`; contact e-mail and phone changes logged with old and new values.
- `install.ps1`: UTF-8 everywhere, pinned `requirements.txt`, asks before replacing an existing `outlook` entry.
- `SECURITY.md` and a "Security model" section in the README.

### Added
- `save_mail` (.msg with transport headers and SHA-256 fingerprint), `read_mail(headers=true)`, fingerprints for
  `save_attachments`, recurring series and single occurrences in the calendar tools.
- `test_unit.py` (runs anywhere, no Outlook) and a GitHub Actions workflow; `test_accents.py` with 56 end-to-end checks.
- `server_version` in `outlook_status`; `OUTLOOK_CLASSIC_MCP_CONFIG` to point the server at another config file.

### Fixed
- DST handling in DASL date filters; `list_events` no longer returns one row past the end date.

## [0.1.0] - 2026-11-09

First public version (English tool names), Italian README.
