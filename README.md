# Email Context

Two small PowerShell tools for getting at what's buried in your Outlook mailbox:

- **`Get-EmailContext.ps1`** — pulls every email exchanged with one person over the last
  N days into a single markdown file with full text and dates, for pasting into Claude as
  a "catch me up" before a meeting.
- **`Find-Email.ps1`** — finds the exact message in a long reply chain where your search
  words were actually written, and gives you a search string that pulls it up in new Outlook.

Both use Outlook COM automation against your local desktop profile. No API keys, no Azure
app registration, no IT ticket. Read-only — they never send, move, or delete anything.

## Install (any Windows PC)

1. Download the latest **`FADOE-vX.Y.Z.zip`** from
   [Releases](https://github.com/zediiiii/outlook-context/releases/latest).
2. Extract it, open the extracted folder, and double-click **`Install.cmd`**.
3. Open **FADOE** from the desktop shortcut or the Start menu.

It installs to `%LOCALAPPDATA%\Programs\FADOE` for your Windows account only — no admin
rights needed. To upgrade, download the newer zip and run `Install.cmd` again. To remove it:

```bash
powershell -ExecutionPolicy Bypass -File "%LOCALAPPDATA%\Programs\FADOE\Uninstall-FADOE.ps1"
```

**Requirements:** Windows 10 or 11 with the **classic** Outlook desktop app installed and
signed in to your mailbox once. You can keep using new Outlook day to day — FADOE just reads
your mail through classic Outlook in the background.

## FADOE: one window for both (Find A Damn Outlook Email)

`FADOE.ps1` opens a proper window (follows Windows light/dark mode) with two tabs:

- **Find an email** — type what you remember, press Enter. Results on the left, the whole
  message on the right with your words highlighted and links clickable. Buttons to copy
  the new Outlook search string, or **open just that one message** in its own window —
  which sidesteps conversation grouping entirely, no Outlook settings changed.
  Double-click a result to copy its search.
- **Catch up on a person** — enter their address and how many days back. Every message is
  listed and readable in the window, and the whole document is put on your clipboard to
  paste into Claude. Buttons to copy it again, open the file, or show it in its folder.

`Ctrl+1` / `Ctrl+2` switch tabs. Searches run in the background, so the window stays
responsive and shows which folder it's on.

Running from a clone of this repo instead of the installer? Create the desktop and Start
menu shortcuts once, then just double-click:

```bash
powershell -ExecutionPolicy Bypass -File "FADOE.ps1" -InstallShortcut
```

`-Theme light` or `-Theme dark` overrides the Windows setting. The two scripts below
still work on their own from the command line.

**Why "open just this message" exists:** with conversation grouping on, new Outlook
decides which message of a thread to show, and no search operator can change that. The
search string narrows things to the right thread and day; opening the single message
skips the thread altogether. That window comes from classic Outlook, because new Outlook
can't be scripted.

## Get-EmailContext: catch up on one person

Everything involving that person in the last 180 days:

```bash
powershell -ExecutionPolicy Bypass -File "Get-EmailContext.ps1" -Person "jdoe@example.org" -Days 180
```

Output lands in `context\<person>_<date>.md` unless you pass `-OutFile`.

This deliberately **includes group and distribution mail** — messages a third party sent
where you and they were both recipients. For meeting prep most of the actual project state
lives in those, so they're in by default. `-DirectOnly` narrows to pure one-on-one if you
ever want it, but it's the wrong default for catching up.

### Options

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

### What it does

Walks every mail folder in your mailbox (not just Inbox and Sent — it picks up anything
you've filed into subfolders), date-filters server-side, then keeps messages where the
person appears as sender, To, or CC. Exchange `/o=...` internal addresses are resolved to
real SMTP addresses. Duplicates filed in two folders are collapsed.

Output is chronological, oldest first: a header with counts and date range, an index table
of every message, then the complete untruncated body of each one with sender, recipients,
attachment names, and source folder.

If nothing matches: try a shorter token (just the surname), a bigger `-Days`, or add
`-AllStores` in case the mail lives in an online archive. Roughly 4KB of output ≈ 1k tokens;
a 40k-token file is fine to hand to Claude directly, and `-TrimQuoted` shrinks bigger ones.

## Find-Email: find the exact message in a thread

Outlook's search matches every reply in a conversation, because each reply quotes the
messages before it — so the one you want gets buried. This matches only the text each
sender actually wrote (quoted history ignored) and returns one result per conversation:
the message where your words appear.

```bash
powershell -ExecutionPolicy Bypass -File "Find-Email.ps1" 'kickoff video'
```

Each result shows the exact date and time, sender, subject, folder, which message it is in
the thread (e.g. *message 1 of 48, oldest first*), the matching passage highlighted, any
links in it, and a search string for new Outlook:

```
from:jdoe@example.org AND subject:"Budget planning" AND sent:08/04/2026
```

The top result's search string is copied to the clipboard — paste it into new Outlook's
search box. With conversations grouped, open the thread and look for the time shown.

| Flag | Meaning |
|---|---|
| `Search` | Words that must all appear. Quote a phrase: `'budget "site visit"'` |
| `-From` | Only senders whose name or address contains this text. |
| `-Days` | Lookback window. Default 365. |
| `-Top` | Conversations to show. Default 10. |
| `-NoClipboard` | Don't copy the top search string. |

Ranking: words the sender wrote beat words that only appear in the subject; all your words
in the subject line ranks highest; words close together (a phrase) beat scattered mentions;
newsletters and mass mailings sink below real correspondence; then newest first.

## Notes

- Requires the **classic** Outlook desktop app to be *installed* with your mail profile set
  up. You don't have to use it or even open it — you can do all your email in new Outlook.
  New Outlook can't be scripted, so these tools read your mailbox through classic Outlook
  in the background. Don't uninstall classic Outlook.
- Classic Outlook keeps about **one year** of mail on your PC by default, so these tools
  can't see older mail. New Outlook's own search still covers everything on the server.
- Shared mailboxes on your profile are searched too. If a result lives in one, click a
  folder in that mailbox in new Outlook before pasting the search string.

## Making a release

Bump `$FadoeVersion` at the top of `FADOE.ps1`, commit, then:

```bash
powershell -ExecutionPolicy Bypass -File "tools\Build-Release.ps1"
```

That writes `dist\FADOE-v<version>.zip`; attach it to a GitHub release tagged `v<version>`.

## Privacy

The files this script writes contain **real correspondence** — full message bodies plus the
names and addresses of everyone on each thread. `context/` is gitignored for that reason.
If you fork this or change the output path, make sure the new path is ignored too. Nothing
is ever uploaded anywhere by the script itself; it reads your local mailbox and writes a
local file, and that's all.
