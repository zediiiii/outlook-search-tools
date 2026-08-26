# Email Context

Pull every email exchanged with one person over the last N days out of Outlook, into a
single markdown file with full text and dates — meant to be pasted into Claude for a
"catch me up" summary before a meeting.

Uses Outlook COM automation against your local desktop profile. No API keys, no Azure app
registration, no IT ticket. Read-only — it never sends, moves, or deletes anything.

## Usage

Everything involving that person in the last 180 days:

```bash
powershell -ExecutionPolicy Bypass -File "Get-EmailContext.ps1" -Person "jdoe@example.org" -Days 180
```

Output lands in `context\<person>_<date>.md` unless you pass `-OutFile`.

This deliberately **includes group and distribution mail** — messages a third party sent
where you and they were both recipients. For meeting prep most of the actual project state
lives in those, so they're in by default. `-DirectOnly` narrows to pure one-on-one if you
ever want it, but it's the wrong default for catching up.

## Options

| Flag | Meaning |
|---|---|
| `-Person` | Email address, partial address, or display name. Accepts several: `-Person 'jdoe@example.org','Jane Doe'` |
| `-Days` | Lookback window. Default 180. |
| `-Since` | Explicit start date instead of `-Days`, e.g. `-Since 2026-01-01` |
| `-OutFile` | Where to write it. |
| `-DirectOnly` | Only messages they sent, or that you sent to them. |
| `-TrimQuoted` | Strip quoted reply history. Much smaller file, some text lost. |
| `-IncludeDeleted` | Also scan Deleted Items, Junk, Drafts. |
| `-AllStores` | Also scan other mailboxes / PSTs / online archives on the profile. |

## What it does

Walks every mail folder in your mailbox (not just Inbox and Sent — it picks up anything
you've filed into subfolders), date-filters server-side, then keeps messages where the
person appears as sender, To, or CC. Exchange `/o=...` internal addresses are resolved to
real SMTP addresses. Duplicates filed in two folders are collapsed.

Output is chronological, oldest first: a header with counts and date range, an index table
of every message, then the complete untruncated body of each one with sender, recipients,
attachment names, and source folder.

## Notes

- Requires the **classic** Outlook desktop app. The new "Outlook for Windows" store app
  has no COM interface — if you ever get switched to it, this stops working.
- Outlook doesn't have to be open; the script starts it in the background.
- If nothing matches: try a shorter token (just the surname), a bigger `-Days`, or add
  `-AllStores` in case the mail lives in an online archive.
- Roughly 4KB of output ≈ 1k tokens. A 40k-token file is fine to hand to Claude directly;
  if a run comes back much larger, add `-TrimQuoted`.

## Privacy

The files this script writes contain **real correspondence** — full message bodies plus the
names and addresses of everyone on each thread. `context/` is gitignored for that reason.
If you fork this or change the output path, make sure the new path is ignored too. Nothing
is ever uploaded anywhere by the script itself; it reads your local mailbox and writes a
local file, and that's all.
