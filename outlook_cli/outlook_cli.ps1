#Requires -Version 5.1
<#
.SYNOPSIS
    Command-line tools for Outlook/Teams meetings (classic Outlook via COM).
.EXAMPLE
    .\outlook_cli.ps1 search "EAWM Backlog Grooming"
.EXAMPLE
    .\outlook_cli.ps1 attendees "EAWM Backlog Refinement" -Date 2026-10-26 -Csv .\attendees.csv
.EXAMPLE
    .\outlook_cli.ps1 edit
.EXAMPLE
    .\outlook_cli.ps1 update "EAWM Backlog Refinement" -Series -Optional Ken.Nelson@va.gov -Remove "Ravi Chamarti" -EndDate 2026-12-13 -Preview
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [ValidateSet('search', 'attendees', 'update', 'edit', 'help')]
    [string]$Command = 'help',

    [Parameter(Position = 1)]
    [string]$Query,

    [int]$PastDays = 365,
    [int]$FutureDays = 365,
    [datetime]$Date = (Get-Date).Date,
    [switch]$Like,
    [string]$Csv,

    [string[]]$Required,
    [string[]]$Optional,
    [string[]]$Remove,
    [Nullable[datetime]]$EndDate,
    [switch]$Series,
    [switch]$Preview,
    [switch]$Send
)

Import-Module (Join-Path $PSScriptRoot 'OutlookCli.psm1') -Force

$from = (Get-Date).Date.AddDays(-$PastDays)
$to = (Get-Date).Date.AddDays($FutureDays)

function Show-Help {
    @'
outlook_cli <command> [query] [options]

Commands:
  search <words>        Find meeting subjects containing any of the words (best matches first)
  attendees <subject>   List attendees of the next occurrence on/after -Date
  update <subject>      Change attendees and/or series end date (saves, then opens in Outlook to send)
  edit                  Open the attendee editor GUI
  help                  Show this help

Options:
  -PastDays <n>         Days back to search (default 365)
  -FutureDays <n>       Days ahead to search (default 365)
  -Date <date>          Occurrence date for 'attendees'/'update' (default today)
  -Like                 Match subject by substring instead of exact
  -Csv <path>           'attendees': also export to CSV

Update options:
  -Required <a,b>       Add as Required, or move existing attendees to Required
  -Optional <a,b>       Add as Optional, or move existing attendees to Optional
  -Remove <a,b>         Remove attendees
                        (emails match exactly; names are resolved via the address book)
  -Series               Apply to the whole series instead of the -Date occurrence
  -EndDate <date>       Set the series end date (implies -Series); moved/cancelled
                        occurrences are checked and restored automatically
  -Preview              Show what would change without saving
  -Send                 Send the update to ALL attendees instead of opening it in Outlook
'@
}

function Assert-Query {
    if (-not $Query) { throw "'$Command' needs a query. Run: outlook_cli.ps1 help" }
}

function New-AttendeeChange([string]$Action, [string]$Text, [string]$Type) {
    $Text = $Text.Trim()
    # Exact emails skip resolution so an external address can't resolve to the person's internal entry.
    if ($Text -match '@') { return [pscustomobject]@{ Action = $Action; Name = ''; Email = $Text; Type = $Type } }
    $res = Resolve-OutlookAddress $Text
    if ($res) { return [pscustomobject]@{ Action = $Action; Name = ''; Email = $res.Email; Type = $Type } }
    if ($Action -eq 'Add') { throw "Could not resolve '$Text' (not found or ambiguous)." }
    [pscustomobject]@{ Action = $Action; Name = $Text; Email = ''; Type = $Type }
}

function Split-List([string[]]$Values) {
    $Values -split '[;,]' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
}

switch ($Command) {
    'help' { Show-Help }

    'search' {
        Assert-Query
        $results = @(Find-OutlookMeetingSubject -Query $Query -From $from -To $to)
        "Found $($results.Count) subject(s) matching '$Query':"
        $results | Format-Table -AutoSize | Out-String -Width 250
    }

    'attendees' {
        Assert-Query
        $hits = @(Get-OutlookMeeting -Subject $Query -Like:$Like -From $from -To $to)
        if (-not $hits) {
            "No meeting named '$Query'. Closest subjects:"
            Find-OutlookMeetingSubject -Query $Query -From $from -To $to | Select-Object -First 10 | Format-Table -AutoSize | Out-String -Width 250
            return
        }
        "Occurrences: $($hits.Count)"
        $hits | Select-Object -First 10 | ForEach-Object {
            "  {0:g}  {1}  (Recurring: {2}, Organizer: {3})" -f $_.Start, $_.Subject, ($_.RecurrenceState -ne 0), $_.Organizer
        }
        $m = $hits | Where-Object { $_.Start -ge $Date } | Select-Object -First 1
        if (-not $m) { $m = $hits[-1] }
        $list = @(Get-OutlookAttendee -Item $m | Sort-Object Type, Name)
        "`nAttendees for $($m.Subject) on $($m.Start.ToString('g')):  " +
            "$(@($list | Where-Object Type -eq 'Required').Count) required, $(@($list | Where-Object Type -eq 'Optional').Count) optional"
        $list | Format-Table -AutoSize | Out-String -Width 250
        if ($Csv) {
            $list | Select-Object @{ n = 'Subject'; e = { $m.Subject } }, @{ n = 'Start'; e = { $m.Start } }, Name, Email, Type, Response |
                Export-Csv -Path $Csv -NoTypeInformation
            "Exported to $Csv"
        }
    }

    'update' {
        Assert-Query
        if (-not ($Required -or $Optional -or $Remove -or $EndDate)) { throw 'Nothing to update. Use -Required, -Optional, -Remove or -EndDate.' }
        $hits = @(Get-OutlookMeeting -Subject $Query -Like:$Like -From $from -To $to)
        if (-not $hits) { throw "No meeting named '$Query'. Try: outlook_cli.ps1 search `"$Query`"" }
        $subjects = @($hits.Subject | Sort-Object -Unique)
        if ($subjects.Count -gt 1) { throw "Several subjects match; be more specific:`n  $($subjects -join "`n  ")" }
        $occ = $hits | Where-Object { $_.Start -ge $Date } | Select-Object -First 1
        if (-not $occ) { $occ = $hits[-1] }

        $isRecurring = $occ.RecurrenceState -ne 0
        if ($EndDate -and -not $isRecurring) { throw "'$Query' is not recurring; -EndDate needs a series." }
        $useSeries = $isRecurring -and ($Series -or $EndDate)
        $target = if ($useSeries) { $occ.GetRecurrencePattern().Parent } else { $occ }
        $blocker = Get-OutlookEditBlocker $target
        if ($blocker) { throw "'$Query': $blocker" }

        $changes = @(
            foreach ($a in Split-List $Remove) { New-AttendeeChange 'Remove' $a '' }
            foreach ($a in Split-List $Required) { New-AttendeeChange 'Add' $a 'Required' }
            foreach ($a in Split-List $Optional) { New-AttendeeChange 'Add' $a 'Optional' }
        )

        $scope = if ($useSeries) { 'series' } else { "occurrence $($occ.Start.ToString('ddd MM/dd/yyyy h:mm tt'))" }
        "Updating $scope of '$($target.Subject)'$(if ($Preview) { '  [PREVIEW - nothing saved]' })"
        $log = { param($m) Write-Host $m }
        if ($useSeries) {
            $changed = Update-OutlookSeries -Master $target -Changes $changes -EndDate $EndDate -DryRun:$Preview -Log $log
        }
        else {
            $changed = Update-OutlookAttendee -Item $target -Changes $changes -DryRun:$Preview -Log $log
            if ($changed -and -not $Preview) { $target.Save() }
        }
        if (-not $changed) { 'No changes needed.'; return }
        if ($Preview) { return }

        # Saving occurrences can leave the master object stale; reopen it.
        if ($useSeries) { $target = (Connect-Outlook).GetItemFromID($target.EntryID) }
        if ($Send) {
            $target.Send()
            'Update sent to all attendees.'
        }
        else {
            $target.Display()
            'Saved (not sent). Opened in Outlook - review and click Send Update.'
        }
    }

    'edit' {
        $gui = Join-Path $PSScriptRoot 'Edit-MeetingAttendees.ps1'
        if ([Threading.Thread]::CurrentThread.ApartmentState -eq 'STA') { & $gui }
        else { Start-Process powershell.exe -ArgumentList '-STA', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$gui`"" }
    }
}
