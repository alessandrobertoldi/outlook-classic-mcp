# Security policy

`outlook-classic-mcp` gives an AI agent access to a mailbox, so security reports are taken seriously and handled first.

## Reporting a vulnerability

Please **do not open a public issue**. Use GitHub's private reporting instead: on the repository page, open the
**Security** tab and choose **Report a vulnerability**. Only the maintainer sees the report.

Useful details: the version or commit, Windows and Outlook version, the mail provider (Exchange, IMAP, Kerio...),
the steps or the mail content that triggers the problem, and what an attacker gains.

You can expect an acknowledgement within 5 working days and a first assessment within 15. Fixes are released as soon
as they are ready; reporters are credited in the release notes unless they prefer otherwise.

## Scope

In scope, for example:

- a mail (or any content the agent reads) that makes the server send mail to anyone other than the configured alert
  recipient, delete permanently, write outside the allowed folders, attach files from outside `attach_folders`,
  reach a mailbox not listed in `config.json`, or make Windows contact a host chosen by the attacker (through a path
  passed to the server, a saved file that Explorer would resolve, or an HTML draft that Outlook would render);
- bypasses of the phase, deletion or subject-confirmation guards;
- secrets or mail content leaking into the logs;
- problems in `install.ps1` that damage the Claude desktop configuration or weaken the system.

Out of scope: the behaviour of Claude itself, of Outlook or of the mail provider, and attacks that require
administrator access to the PC or control of `config.json` (whoever can edit it already controls the server).

## Supported versions

Only the latest tagged release (see `CHANGELOG.md`) receives fixes.
