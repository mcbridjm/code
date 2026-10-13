# outlook_cli

PowerShell tools for reviewing and editing Outlook and Teams meetings. Teams meetings are Outlook calendar items, so everything works through classic desktop Outlook (COM) — no app registration or Graph permissions needed.

## Requirements

- Windows with **classic Outlook** running and signed in. New Outlook is not supported (no COM).
- Windows PowerShell 5.1 or PowerShell 7.
- If scripts are blocked by execution policy, run with:
  ```powershell
  powershell -ExecutionPolicy Bypass -File .\outlook_cli\outlook_cli.ps1 help
  ```

## Files

| File | Purpose |
|---|---|
| `outlook_cli.ps1` | Command-line entry point (subcommands below) |
| `OutlookCli.psm1` | Shared Outlook functions used by the CLI and GUI |
| `Edit-MeetingAttendees.ps1` | WinForms attendee editor (launched by `outlook_cli.ps1 edit`) |

## Commands

```powershell
.\outlook_cli\outlook_cli.ps1 <command> [query] [options]
```

### search

Find meeting subjects containing any of the words, best matches first.

```powershell
.\outlook_cli\outlook_cli.ps1 search "EAWM Backlog Grooming"
```

Shows each subject with words matched, occurrence count, next and last dates. Useful when you don't remember the exact meeting name.

### attendees

List attendees of the next occurrence on or after `-Date` (default today). If no exact subject match, shows the closest subjects instead.

```powershell
.\outlook_cli\outlook_cli.ps1 attendees "EAWM Backlog Refinement"
.\outlook_cli\outlook_cli.ps1 attendees "EAWM Backlog Refinement" -Date 2026-10-26 -Csv .\attendees.csv
```

### update

Change attendees and/or the series end date. By default it saves the change and opens the meeting in Outlook so you can click **Send Update** (and choose who gets notified).

```powershell
# Preview first - nothing is saved
.\outlook_cli\outlook_cli.ps1 update "EAWM Backlog Refinement" -Series `
    -Optional Ken.Nelson@va.gov,Adam.Schaefer@us.ibm.com `
    -Required Matthew.Fernicola@va.gov `
    -Remove "Ravi Chamarti" `
    -EndDate 2026-12-13 -Preview

# Same command without -Preview to apply
```

| Option | Meaning |
|---|---|
| `-Required a,b` | Add as Required, or move existing attendees to Required |
| `-Optional a,b` | Add as Optional, or move existing attendees to Optional |
| `-Remove a,b` | Remove attendees |
| `-Series` | Apply to the whole series instead of the `-Date` occurrence |
| `-EndDate <date>` | Set the series end date (implies `-Series`) |
| `-Date <date>` | Which occurrence to target (default: next on/after today) |
| `-Preview` | Show what would change without saving |
| `-Send` | Send to **all** attendees immediately instead of opening in Outlook |

Notes:
- **Emails match exactly**; they are not resolved, so an external address (e.g. `@us.ibm.com`) won't match the same person's internal entry. **Names** are resolved via the address book.
- **Moved/cancelled occurrence check:** changing a series' end date through COM wipes individually moved or cancelled occurrences. `update` records them first, restores any that were lost (time, subject, location), and applies the attendee changes to them too.
- If the end date is already the requested date, the series isn't re-saved.
- Only meetings you organize can be updated.

### edit

Opens the attendee editor GUI.

```powershell
.\outlook_cli\outlook_cli.ps1 edit
```

### help

```powershell
.\outlook_cli\outlook_cli.ps1 help
```

### Common options

| Option | Default | Meaning |
|---|---|---|
| `-PastDays <n>` | 365 | Days back to search |
| `-FutureDays <n>` | 365 | Days ahead to search |
| `-Like` | off | Match subject by substring instead of exact |

## Attendee editor (GUI)

1. **Find meetings:** *Use Outlook selection* (selected/open item), *Search* by subject and date range, or *Find by link* (paste a Teams join link).
2. Choose **This occurrence** or **Whole series** for recurring meetings.
3. Queue changes (they apply to every checked meeting):
   - Type names/emails and *Queue add* as Required or Optional.
   - *Queue DL* adds a distribution list (tick *Expand DL* to add its members).
   - *Import CSV...* (format below).
   - Select attendees in the list and *Queue remove* / *Queue -> Required* / *Queue -> Optional*.
   - *Queue copy of all to checked* copies the highlighted meeting's attendees.
4. **Preview** shows the diff per meeting; **Apply** makes the changes and opens each meeting in Outlook to send (max 10 at a time; change with `-BatchSize`). *Auto-send* sends to all attendees without review.
5. **Export CSV...** saves attendees of checked (or all listed) meetings.

### CSV import format

```csv
Email,Type,Action
jane.doe@va.gov,Optional,Add
john.smith@va.gov,,Remove
pat.lee@va.gov,Required,Move
```

- `Type`: `Required` (default) or `Optional`. `Action`: `Add` (default), `Remove`, or `Move`.
- Any invalid row cancels the whole import.

## Known limitations

- Outlook may show an "a program is trying to access…" prompt when reading meeting bodies (Teams link search) or addresses, if antivirus status isn't recognized.
- Restoring a moved occurrence restores time, subject, and location only — not other one-off edits to it.
- Date filtering uses the system locale date format.
- When sending a series update after an end-date change, send to **all attendees** so everyone gets the new end date.
- **Teams channel meetings** (created in Teams with a channel selected) are organized by the team's Microsoft 365 group, not you. Your calendar holds only an invitee copy, and the group's calendar isn't reachable via COM, so `update`/GUI can't edit them. They are detected automatically (their Teams join link points to a channel thread, `@thread.skype`/`@thread.tacv2`): `update` stops with *"This is a Teams channel meeting, edit it in Teams"* and the GUI shows *Teams channel (edit in Teams)*. Edit attendees in Teams (Calendar → open meeting → Edit) instead.

## Adding a feature

1. Put reusable Outlook logic in `OutlookCli.psm1` and add it to `Export-ModuleMember`.
2. In `outlook_cli.ps1`: add the command to the `ValidateSet`, add a `switch` block, and update `Show-Help`.
3. Update this README.

## Changelog

- **2026-10-09** — Initial GUI editor; CLI with `search`, `attendees`, `edit`; GUI refactored onto shared module; added `update` command with series end date and moved-occurrence check; README added.
- **2026-10-09** — Detect Teams channel meetings (`Test-TeamsChannelMeeting`, `Get-OutlookEditBlocker`) and report "edit it in Teams" in CLI and GUI.
