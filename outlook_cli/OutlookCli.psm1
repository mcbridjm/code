# Shared Outlook (classic, COM) helpers for outlook_cli. Add new reusable functions here.

$script:Outlook = $null
$script:Ns = $null

$olRequired = 1; $olOptional = 2; $olResource = 3
$olAppointment = 26; $olMeetingRequest = 53
$olMeeting = 1; $olMeetingReceived = 3
$olResponseOrganized = 1
$ResponseNames = @{ 0 = 'None'; 1 = 'Organizer'; 2 = 'Tentative'; 3 = 'Accepted'; 4 = 'Declined'; 5 = 'No response' }
$RecipientTypes = @{ 1 = 'Required'; 2 = 'Optional'; 3 = 'Resource' }

function Connect-Outlook {
    if (-not $script:Ns) {
        try {
            $script:Outlook = New-Object -ComObject Outlook.Application
            $script:Ns = $script:Outlook.GetNamespace('MAPI')
        }
        catch { throw "Could not connect to classic Outlook (new Outlook has no COM): $($_.Exception.Message)" }
    }
    $script:Ns
}

function Get-SmtpAddress($Recipient) {
    try {
        $ae = $Recipient.AddressEntry
        if ($ae -and $ae.Type -eq 'EX') {
            $xu = $ae.GetExchangeUser()
            if ($xu) { return $xu.PrimarySmtpAddress }
            $xdl = $ae.GetExchangeDistributionList()
            if ($xdl) { return $xdl.PrimarySmtpAddress }
        }
        if ($ae) { return $ae.Address }
    }
    catch { }
    $Recipient.Address
}

# Expands recurring meetings into individual occurrences within the range.
function Get-CalendarItem {
    param(
        [datetime]$From = (Get-Date).Date.AddDays(-30),
        [datetime]$To = (Get-Date).Date.AddDays(90)
    )
    $items = (Connect-Outlook).GetDefaultFolder(9).Items
    # Sort must precede IncludeRecurrences for occurrences to expand correctly.
    $items.Sort('[Start]')
    $items.IncludeRecurrences = $true
    $found = $items.Restrict(("[Start] >= '{0}' AND [Start] < '{1}'" -f $From.ToString('g'), $To.ToString('g')))
    $i = $found.GetFirst()
    while ($i) {
        if ($i.Class -eq 26) { $i }
        $i = $found.GetNext()
    }
}

function Find-OutlookMeetingSubject {
    param(
        [Parameter(Mandatory)][string]$Query,
        [datetime]$From = (Get-Date).Date.AddDays(-365),
        [datetime]$To = (Get-Date).Date.AddDays(365)
    )
    $words = @($Query -split '\s+' | Where-Object { $_ })
    $now = Get-Date
    Get-CalendarItem -From $From -To $To |
        ForEach-Object {
            $subject = $_.Subject
            $score = @($words | Where-Object { $subject -like "*$_*" }).Count
            if ($score) { [pscustomobject]@{ Subject = $subject; Score = $score; Start = $_.Start } }
        } |
        Group-Object Subject |
        ForEach-Object {
            $starts = $_.Group.Start | Sort-Object
            [pscustomobject]@{
                Subject     = $_.Name
                WordsMatched = "$($_.Group[0].Score)/$($words.Count)"
                Occurrences = $_.Count
                Next        = $starts | Where-Object { $_ -ge $now } | Select-Object -First 1
                Last        = $starts[-1]
                Score       = $_.Group[0].Score
            }
        } |
        Sort-Object @{ Expression = 'Score'; Descending = $true }, Subject |
        Select-Object Subject, WordsMatched, Occurrences, Next, Last
}

function Get-OutlookMeeting {
    param(
        [Parameter(Mandatory)][string]$Subject,
        [switch]$Like,
        [datetime]$From = (Get-Date).Date.AddDays(-30),
        [datetime]$To = (Get-Date).Date.AddDays(90)
    )
    Get-CalendarItem -From $From -To $To | Where-Object {
        if ($Like) { $_.Subject -like "*$Subject*" } else { $_.Subject -eq $Subject }
    }
}

function Get-OutlookAttendee {
    param([Parameter(Mandatory)]$Item)
    $recips = $Item.Recipients
    for ($n = 1; $n -le $recips.Count; $n++) {
        $r = $recips.Item($n)
        [pscustomobject]@{
            Name     = $r.Name
            Email    = (Get-SmtpAddress $r)
            Type     = $RecipientTypes[[int]$r.Type]
            Response = $ResponseNames[[int]$r.MeetingResponseStatus]
        }
    }
}

function Test-OutlookMeeting($Item) {
    $Item.Class -eq $olAppointment -and $Item.MeetingStatus -in $olMeeting, $olMeetingReceived
}

function Test-OutlookOrganizer($Item) {
    $Item.MeetingStatus -eq $olMeeting
}

# Channel meetings join via the channel thread (@thread.skype/@thread.tacv2), not a 19:meeting_ thread.
function Test-TeamsChannelMeeting($Item) {
    $decoded = $Item.Body
    if (-not $decoded) { return $false }
    do { $prev = $decoded; $decoded = [uri]::UnescapeDataString($decoded) } while ($decoded -ne $prev)
    $decoded -match '19:[^@\s/]+@thread\.(skype|tacv2)'
}

# Returns $null if you can edit the meeting, otherwise the reason you can't.
function Get-OutlookEditBlocker($Item) {
    if (Test-OutlookOrganizer $Item) { return $null }
    if (Test-TeamsChannelMeeting $Item) { return "This is a Teams channel meeting (organizer: $($Item.Organizer)), edit it in Teams." }
    "You are not the organizer (organizer: $($Item.Organizer))."
}

# Meetings open in an Outlook window, else those selected in the active explorer.
function Get-OutlookSelectedMeeting {
    [void](Connect-Outlook)
    $items = @()
    $insp = $script:Outlook.ActiveInspector()
    if ($insp) { $items += $insp.CurrentItem }
    else {
        $exp = $script:Outlook.ActiveExplorer()
        if ($exp) { foreach ($s in $exp.Selection) { $items += $s } }
    }
    foreach ($i in $items) {
        if ($i.Class -eq $olMeetingRequest) { $i = $i.GetAssociatedAppointment($false) }
        if ($i -and (Test-OutlookMeeting $i)) { $i }
    }
}

function Resolve-OutlookAddress([string]$Text) {
    $r = (Connect-Outlook).CreateRecipient($Text.Trim())
    if (-not $r.Resolve()) { return $null }
    [pscustomobject]@{ Name = $r.Name; Email = (Get-SmtpAddress $r) }
}

function Get-OutlookDistributionListMember([string]$Text) {
    $r = (Connect-Outlook).CreateRecipient($Text.Trim())
    if (-not $r.Resolve()) { throw "Could not resolve '$Text'." }
    $ae = $r.AddressEntry
    $xdl = $ae.GetExchangeDistributionList()
    $members = if ($xdl) { $xdl.GetExchangeDistributionListMembers() } else { $ae.Members }
    if (-not $members) { throw "'$($r.Name)' is not a distribution list or has no members." }
    foreach ($m in $members) {
        $email = $m.Address
        try { $xu = $m.GetExchangeUser(); if ($xu) { $email = $xu.PrimarySmtpAddress } } catch { }
        [pscustomobject]@{ Name = $m.Name; Email = $email }
    }
}

# Returns the text to look for in a meeting body to identify a Teams join link.
function Get-TeamsMeetingToken([string]$Link) {
    $decoded = $Link.Trim()
    do { $prev = $decoded; $decoded = [uri]::UnescapeDataString($decoded) } while ($decoded -ne $prev)
    if ($decoded -match '19:(meeting_[^@\s/]+)@thread\.v2') { return $Matches[1] }
    if ($decoded -match '/meet/(\d+)') { return $Matches[1] }
    return $null
}

function Find-OutlookRecipientIndex($Item, $Change) {
    $recips = $Item.Recipients
    for ($i = 1; $i -le $recips.Count; $i++) {
        $r = $recips.Item($i)
        if ($r.Type -eq $olResource) { continue }
        if (($Change.Email -and (Get-SmtpAddress $r) -eq $Change.Email) -or ($Change.Name -and $r.Name -eq $Change.Name)) { return $i }
    }
    return 0
}

# Applies Add/Remove/Move changes (Action, Name, Email, Type) to an item without saving or sending.
# Returns $true if the item was (or with -DryRun would be) modified.
function Update-OutlookAttendee {
    param(
        [Parameter(Mandatory)]$Item,
        [Parameter(Mandatory)]$Changes,
        [switch]$DryRun,
        [scriptblock]$Log = { param($m) Write-Host $m }
    )
    $changed = $false
    foreach ($c in $Changes) {
        $label = "$($c.Name) <$($c.Email)>"
        $idx = Find-OutlookRecipientIndex $Item $c
        $r = if ($idx) { $Item.Recipients.Item($idx) } else { $null }
        $newType = if ($c.Type -eq 'Optional') { $olOptional } else { $olRequired }

        if ($r -and $r.MeetingResponseStatus -eq $olResponseOrganized) {
            $null = & $Log "   skip organizer: $label"
            continue
        }
        switch ($c.Action) {
            'Add' {
                if ($r -and $r.Type -eq $newType) { $null = & $Log "   = already $($c.Type): $label"; break }
                if ($r) {
                    if (-not $DryRun) { $r.Type = $newType }
                    $null = & $Log "   ~ $label -> $($c.Type)"
                }
                else {
                    if (-not $DryRun) {
                        $new = $Item.Recipients.Add($c.Email)
                        $new.Type = $newType
                        if (-not $new.Resolve()) { $new.Delete(); $null = & $Log "   ! could not resolve $label"; break }
                    }
                    $null = & $Log "   + $label ($($c.Type))"
                }
                $changed = $true
            }
            'Remove' {
                if (-not $r) { $null = & $Log "   ? not found: $label"; break }
                if (-not $DryRun) { $Item.Recipients.Remove($idx) }
                $null = & $Log "   - $label"
                $changed = $true
            }
            'Move' {
                if (-not $r) { $null = & $Log "   ? not found: $label"; break }
                if ($r.Type -eq $newType) { $null = & $Log "   = already $($c.Type): $label"; break }
                if (-not $DryRun) { $r.Type = $newType }
                $null = & $Log "   ~ $label -> $($c.Type)"
                $changed = $true
            }
        }
    }
    return $changed
}

function Get-ExceptionSnapshot($Pattern) {
    foreach ($e in $Pattern.Exceptions) {
        $orig = $e.OriginalDate
        if ($orig.TimeOfDay -eq [timespan]::Zero) { $orig = $orig.Date + $Pattern.StartTime.TimeOfDay }
        if ($e.Deleted) { [pscustomobject]@{ Original = $orig; Deleted = $true } }
        else {
            $a = $e.AppointmentItem
            [pscustomobject]@{ Original = $orig; Deleted = $false; Start = $a.Start; End = $a.End; Subject = $a.Subject; Location = $a.Location }
        }
    }
}

# Applies attendee changes and/or a new end date to a recurring series master, saves (does not send),
# restores moved/deleted occurrences that COM drops when the pattern changes, and syncs attendee
# changes onto exception occurrences. Returns $true if anything was (or with -DryRun would be) changed.
function Update-OutlookSeries {
    param(
        [Parameter(Mandatory)]$Master,
        $Changes = @(),
        [Nullable[datetime]]$EndDate,
        [switch]$DryRun,
        [scriptblock]$Log = { param($m) Write-Host $m }
    )
    $p = $Master.GetRecurrencePattern()
    $snapshot = @(Get-ExceptionSnapshot $p)
    $changed = $false

    if ($Changes) {
        $null = & $Log '== Series attendees'
        if (Update-OutlookAttendee -Item $Master -Changes $Changes -DryRun:$DryRun -Log $Log) { $changed = $true }
    }

    if ($EndDate) {
        $end = ([datetime]$EndDate).Date
        if ($end -lt $p.PatternStartDate.Date) { throw "End date $($end.ToShortDateString()) is before the series start $($p.PatternStartDate.ToShortDateString())." }
        $old = if ($p.NoEndDate) { 'no end' } else { $p.PatternEndDate.ToShortDateString() }
        if (-not $p.NoEndDate -and $p.PatternEndDate.Date -eq $end) {
            $null = & $Log "== Series end already $old"
        }
        else {
            $null = & $Log "== Series end: $old -> $($end.ToShortDateString())"
            foreach ($s in $snapshot | Where-Object { $_.Original.Date -gt $end }) {
                $null = & $Log "   exception on $($s.Original.ToShortDateString()) falls after the new end and will be dropped"
            }
            if (-not $DryRun) { $p.PatternEndDate = $end }
            $changed = $true
        }
    }

    if ($snapshot) {
        $null = & $Log "== Moved/cancelled occurrences to check: $($snapshot.Count)"
        foreach ($s in $snapshot) {
            $what = if ($s.Deleted) { 'cancelled' } else { "moved to $($s.Start.ToString('ddd MM/dd/yyyy h:mm tt'))" }
            $null = & $Log "   $($s.Original.ToString('ddd MM/dd/yyyy h:mm tt')) $what"
        }
    }
    if ($DryRun -or -not $changed) { return $changed }

    $Master.Save()

    $p = $Master.GetRecurrencePattern()
    $lastDate = if ($p.NoEndDate) { [datetime]::MaxValue } else { $p.PatternEndDate.Date }
    $current = @(Get-ExceptionSnapshot $p)
    foreach ($s in $snapshot | Where-Object { $_.Original.Date -le $lastDate }) {
        if ($current | Where-Object { $_.Original -eq $s.Original }) { continue }
        $o = $Master.GetRecurrencePattern().GetOccurrence($s.Original)
        if ($s.Deleted) {
            $o.Delete()
            $null = & $Log "   restored cancellation of $($s.Original.ToShortDateString())"
        }
        else {
            $o.Start = $s.Start
            $o.End = $s.End
            $o.Subject = $s.Subject
            $o.Location = $s.Location
            $o.Save()
            $null = & $Log "   restored move $($s.Original.ToShortDateString()) -> $($s.Start.ToString('g'))"
        }
    }

    if ($Changes) {
        $count = $Master.GetRecurrencePattern().Exceptions.Count
        for ($i = 1; $i -le $count; $i++) {
            # Re-fetch each time: saving an exception invalidates previously fetched exception objects.
            $e = $Master.GetRecurrencePattern().Exceptions.Item($i)
            if ($e.Deleted) { continue }
            $item = $e.AppointmentItem
            $null = & $Log "== Exception $($item.Start.ToString('ddd MM/dd/yyyy h:mm tt')) attendees"
            if (Update-OutlookAttendee -Item $item -Changes $Changes -Log $Log) { $item.Save() }
        }
    }
    return $true
}

Export-ModuleMember -Function Connect-Outlook, Get-SmtpAddress, Get-CalendarItem, Find-OutlookMeetingSubject, Get-OutlookMeeting, Get-OutlookAttendee,
    Test-OutlookMeeting, Test-OutlookOrganizer, Test-TeamsChannelMeeting, Get-OutlookEditBlocker, Get-OutlookSelectedMeeting, Resolve-OutlookAddress, Get-OutlookDistributionListMember,
    Get-TeamsMeetingToken, Find-OutlookRecipientIndex, Update-OutlookAttendee, Update-OutlookSeries
