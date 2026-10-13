#Requires -Version 5.1
<#
.SYNOPSIS
    GUI for editing required/optional attendees on Outlook and Teams meetings via classic Outlook COM.
.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Edit-MeetingAttendees.ps1
.NOTES
    Requires classic Outlook (new Outlook has no COM). CSV import columns: Email, Type (Required|Optional), Action (Add|Remove|Move).
#>
[CmdletBinding()]
param(
    [int]$BatchSize = 10
)

Set-StrictMode -Version 2
Add-Type -AssemblyName System.Windows.Forms, System.Drawing

if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    throw 'Run this script in Windows PowerShell 5.1 or with -STA.'
}

Import-Module (Join-Path $PSScriptRoot 'OutlookCli.psm1') -Force

try { [void](Connect-Outlook) }
catch {
    [void][Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Edit Meeting Attendees', 'OK', 'Error')
    exit 1
}

#region GUI-specific meeting helpers

function New-MeetingInfo($Item) {
    [pscustomobject]@{
        Item        = $Item
        Subject     = $Item.Subject
        Start       = $Item.Start
        IsRecurring = ($Item.RecurrenceState -ne 0)
        IsOrganizer = (Test-OutlookOrganizer $Item)
        IsChannel   = (-not (Test-OutlookOrganizer $Item) -and (Test-TeamsChannelMeeting $Item))
    }
}

function Find-Meetings([datetime]$From, [datetime]$To, [string]$Subject) {
    Get-CalendarItem -From $From -To $To |
        Where-Object { (Test-OutlookMeeting $_) -and (-not $Subject -or $_.Subject -like "*$Subject*") } |
        ForEach-Object { New-MeetingInfo $_ }
}

function Get-TargetItem($Meeting, [bool]$Series) {
    if ($Series -and $Meeting.IsRecurring) { return $Meeting.Item.GetRecurrencePattern().Parent }
    return $Meeting.Item
}

function Get-Attendees($Item) {
    Get-OutlookAttendee -Item $Item | Where-Object Type -ne 'Resource' | ForEach-Object {
        $isOrg = ($_.Response -eq 'Organizer')
        [pscustomobject]@{
            Name        = $_.Name
            Email       = $_.Email
            Type        = $(if ($isOrg) { 'Organizer' } else { $_.Type })
            Response    = $_.Response
            IsOrganizer = $isOrg
        }
    }
}

#endregion

#region GUI

function New-Grid([int]$X, [int]$Y, [int]$W, [int]$H) {
    $g = New-Object Windows.Forms.DataGridView
    $g.Location = New-Object Drawing.Point($X, $Y)
    $g.Size = New-Object Drawing.Size($W, $H)
    $g.AllowUserToAddRows = $false
    $g.AllowUserToDeleteRows = $false
    $g.RowHeadersVisible = $false
    $g.SelectionMode = 'FullRowSelect'
    $g.AutoSizeColumnsMode = 'Fill'
    $g.ReadOnly = $true
    $form.Controls.Add($g)
    $g
}

function New-Control([string]$Type, [string]$Text, [int]$X, [int]$Y, [int]$W, [int]$H, $Parent = $form) {
    $c = New-Object "Windows.Forms.$Type"
    $c.Text = $Text
    $c.Location = New-Object Drawing.Point($X, $Y)
    if ($W) { $c.Size = New-Object Drawing.Size($W, $H) } else { $c.AutoSize = $true }
    $Parent.Controls.Add($c)
    $c
}

$form = New-Object Windows.Forms.Form
$form.Text = 'Edit Meeting Attendees (Outlook / Teams)'
$form.Size = New-Object Drawing.Size(1200, 860)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedSingle'
$form.MaximizeBox = $false

# Find meetings
$grpFind = New-Control GroupBox 'Find meetings' 10 10 1165 85
$btnSelected = New-Control Button 'Use Outlook selection' 10 20 150 25 $grpFind
New-Control Label 'Subject contains:' 170 24 0 0 $grpFind | Out-Null
$txtSubject = New-Control TextBox '' 270 21 200 22 $grpFind
New-Control Label 'From' 480 24 0 0 $grpFind | Out-Null
$dtFrom = New-Control DateTimePicker '' 515 21 110 22 $grpFind
New-Control Label 'To' 635 24 0 0 $grpFind | Out-Null
$dtTo = New-Control DateTimePicker '' 660 21 110 22 $grpFind
$dtFrom.Format = 'Short'; $dtTo.Format = 'Short'
$dtFrom.Value = (Get-Date).Date.AddDays(-30)
$dtTo.Value = (Get-Date).Date.AddDays(90)
$btnSearch = New-Control Button 'Search' 785 20 95 25 $grpFind
New-Control Label 'Teams join link:' 10 56 0 0 $grpFind | Out-Null
$txtTeams = New-Control TextBox '' 110 53 660 22 $grpFind
$btnTeams = New-Control Button 'Find by link' 785 51 95 25 $grpFind
$rbOccurrence = New-Control RadioButton 'This occurrence' 900 20 130 22 $grpFind
$rbSeries = New-Control RadioButton 'Whole series' 900 48 130 22 $grpFind
$rbOccurrence.Checked = $true
$btnExport = New-Control Button 'Export CSV...' 1040 20 110 25 $grpFind

# Meetings
$gridMeetings = New-Grid 10 105 1165 180
$gridMeetings.ReadOnly = $false
$colApply = New-Object Windows.Forms.DataGridViewCheckBoxColumn
$colApply.HeaderText = 'Apply'; $colApply.FillWeight = 10
[void]$gridMeetings.Columns.Add($colApply)
foreach ($h in 'Subject', 'Start', 'Recurring', 'Organizer') {
    $i = $gridMeetings.Columns.Add($h, $h)
    $gridMeetings.Columns[$i].ReadOnly = $true
}
$gridMeetings.Columns['Subject'].FillWeight = 60
$gridMeetings.Add_CurrentCellDirtyStateChanged({ if ($gridMeetings.IsCurrentCellDirty) { [void]$gridMeetings.CommitEdit('CurrentCellChange') } })

# Attendees of highlighted meeting
New-Control Label 'Attendees of highlighted meeting' 10 295 0 0 | Out-Null
$gridAttendees = New-Grid 10 315 640 220
foreach ($h in 'Name', 'Email', 'Type', 'Response') { [void]$gridAttendees.Columns.Add($h, $h) }
$btnQueueRemove = New-Control Button 'Queue remove' 10 540 120 25
$btnQueueReq = New-Control Button 'Queue -> Required' 135 540 125 25
$btnQueueOpt = New-Control Button 'Queue -> Optional' 265 540 125 25
$btnCopyFrom = New-Control Button 'Queue copy of all to checked' 395 540 180 25

# Add attendees
$grpAdd = New-Control GroupBox 'Add attendees (names or emails separated by ; , or new lines)' 665 295 510 105
$txtAdd = New-Control TextBox '' 10 20 490 42 $grpAdd
$txtAdd.Multiline = $true; $txtAdd.ScrollBars = 'Vertical'
$cboType = New-Control ComboBox '' 10 70 95 22 $grpAdd
$cboType.DropDownStyle = 'DropDownList'
[void]$cboType.Items.AddRange(@('Required', 'Optional')); $cboType.SelectedIndex = 0
$btnAdd = New-Control Button 'Queue add' 112 69 80 25 $grpAdd
$chkExpand = New-Control CheckBox 'Expand DL' 200 72 85 20 $grpAdd
$btnAddDL = New-Control Button 'Queue DL' 288 69 90 25 $grpAdd
$btnCsv = New-Control Button 'Import CSV...' 385 69 115 25 $grpAdd

# Pending changes
New-Control Label 'Pending changes (applied to every checked meeting)' 665 405 0 0 | Out-Null
$gridChanges = New-Grid 665 425 510 110
foreach ($h in 'Action', 'Name', 'Email', 'Type') { [void]$gridChanges.Columns.Add($h, $h) }
$btnRemoveChange = New-Control Button 'Remove change' 665 540 120 25
$btnClearChanges = New-Control Button 'Clear' 790 540 80 25

# Run
$chkAutoSend = New-Control CheckBox 'Auto-send update to ALL attendees (skip review in Outlook)' 10 578 0 0
$btnPreview = New-Control Button 'Preview' 905 573 130 28
$btnApply = New-Control Button 'Apply' 1045 573 130 28
$txtLog = New-Control TextBox '' 10 610 1165 200
$txtLog.Multiline = $true; $txtLog.ScrollBars = 'Vertical'; $txtLog.ReadOnly = $true
$txtLog.Font = New-Object Drawing.Font('Consolas', 9)

function Write-Log([string]$Message) {
    $txtLog.AppendText("$(Get-Date -Format HH:mm:ss)  $Message`r`n")
}

function Invoke-Safe([scriptblock]$Action) {
    $form.Cursor = 'WaitCursor'
    try { & $Action }
    catch {
        Write-Log "ERROR: $($_.Exception.Message)"
        [void][Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Error', 'OK', 'Error')
    }
    finally { $form.Cursor = 'Default' }
}

function Set-Meetings($Meetings, [bool]$CheckOrganized) {
    $gridMeetings.Rows.Clear()
    $gridAttendees.Rows.Clear()
    foreach ($m in @($Meetings)) {
        $idx = $gridMeetings.Rows.Add(
            ($CheckOrganized -and $m.IsOrganizer),
            $m.Subject,
            $m.Start.ToString('g'),
            $(if ($m.IsRecurring) { 'Yes' } else { '' }),
            $(if ($m.IsOrganizer) { 'You' } elseif ($m.IsChannel) { 'Teams channel (edit in Teams)' } else { 'Other (read-only)' }))
        $gridMeetings.Rows[$idx].Tag = $m
    }
    Write-Log "Found $(@($Meetings).Count) meeting(s)."
    Show-Attendees
}

function Get-CheckedRows { foreach ($row in $gridMeetings.Rows) { if ($row.Cells[0].Value -eq $true) { $row } } }

function Get-FocusedMeeting {
    if ($gridMeetings.CurrentRow) { return $gridMeetings.CurrentRow.Tag }
    return $null
}

function Show-Attendees {
    $gridAttendees.Rows.Clear()
    $m = Get-FocusedMeeting
    if (-not $m) { return }
    foreach ($a in Get-Attendees (Get-TargetItem $m $rbSeries.Checked)) {
        $idx = $gridAttendees.Rows.Add($a.Name, $a.Email, $a.Type, $a.Response)
        $gridAttendees.Rows[$idx].Tag = $a
    }
}

function Add-PendingChange([string]$Action, [string]$Name, [string]$Email, [string]$Type) {
    $change = [pscustomobject]@{ Action = $Action; Name = $Name; Email = $Email; Type = $Type }
    $idx = $gridChanges.Rows.Add($Action, $Name, $Email, $(if ($Action -eq 'Remove') { '' } else { $Type }))
    $gridChanges.Rows[$idx].Tag = $change
}

function Get-PendingChanges { foreach ($row in $gridChanges.Rows) { $row.Tag } }

function Add-QueuedFromAttendeeSelection([string]$Action, [string]$Type) {
    foreach ($row in $gridAttendees.SelectedRows) {
        $a = $row.Tag
        if ($a.IsOrganizer) { Write-Log "Skipped organizer $($a.Name)."; continue }
        Add-PendingChange $Action $a.Name $a.Email $Type
    }
}

function Invoke-Run([bool]$DryRun) {
    $changes = @(Get-PendingChanges)
    if (-not $changes) { Write-Log 'No pending changes.'; return }
    $rows = @(Get-CheckedRows)
    if (-not $rows) { Write-Log 'No meetings checked.'; return }
    if (-not $DryRun -and $chkAutoSend.Checked) {
        $answer = [Windows.Forms.MessageBox]::Show("Send updates to ALL attendees of $($rows.Count) meeting(s)?", 'Confirm auto-send', 'YesNo', 'Warning')
        if ($answer -ne 'Yes') { return }
    }
    Write-Log $(if ($DryRun) { '--- Preview (no changes saved) ---' } else { '--- Apply ---' })
    $opened = 0
    foreach ($row in $rows) {
        $m = $row.Tag
        if ($m.IsChannel) { Write-Log "Skipped: '$($m.Subject)' is a Teams channel meeting, edit it in Teams."; continue }
        if (-not $m.IsOrganizer) { Write-Log "Skipped (not organizer): $($m.Subject)"; continue }
        if (-not $DryRun -and -not $chkAutoSend.Checked -and $opened -ge $BatchSize) {
            Write-Log "Batch limit of $BatchSize reached. Send the open meetings, then click Apply again for the rest."
            break
        }
        $item = Get-TargetItem $m $rbSeries.Checked
        $scope = if ($m.IsRecurring) { $(if ($rbSeries.Checked) { ' [series]' } else { ' [occurrence]' }) } else { '' }
        Write-Log "== $($item.Subject)  $($item.Start.ToString('g'))$scope"
        $changed = Update-OutlookAttendee -Item $item -Changes $changes -DryRun:$DryRun -Log { param($msg) Write-Log $msg }
        if ($DryRun) { continue }
        if ($changed) {
            if ($chkAutoSend.Checked) {
                $item.Send()
                Write-Log '   sent update to all attendees'
            }
            else {
                $item.Display()
                $opened++
                Write-Log '   opened in Outlook - click Send Update to choose who gets notified'
            }
        }
        $row.Cells[0].Value = $false
    }
    if (-not $DryRun) { Show-Attendees }
}

$gridMeetings.Add_SelectionChanged({ Invoke-Safe { Show-Attendees } })
$rbSeries.Add_CheckedChanged({ Invoke-Safe { Show-Attendees } })

$btnSelected.Add_Click({ Invoke-Safe {
    $found = @(Get-OutlookSelectedMeeting | ForEach-Object { New-MeetingInfo $_ })
    if (-not $found) { Write-Log 'No meeting selected or open in Outlook.'; return }
    Set-Meetings $found $true
} })

$btnSearch.Add_Click({ Invoke-Safe {
    Set-Meetings @(Find-Meetings $dtFrom.Value.Date $dtTo.Value.Date.AddDays(1) $txtSubject.Text.Trim()) $false
} })

$btnTeams.Add_Click({ Invoke-Safe {
    $token = Get-TeamsMeetingToken $txtTeams.Text
    if (-not $token) { Write-Log 'Could not find a Teams meeting ID in that link.'; return }
    $found = @(Find-Meetings $dtFrom.Value.Date $dtTo.Value.Date.AddDays(1) '' | Where-Object { $_.Item.Body -and $_.Item.Body.Contains($token) })
    if (-not $found) { Write-Log 'No meeting in the date range contains that Teams link.'; return }
    Set-Meetings $found $true
} })

$btnExport.Add_Click({ Invoke-Safe {
    $rows = @(Get-CheckedRows)
    if (-not $rows) { $rows = @($gridMeetings.Rows) }
    if (-not $rows) { Write-Log 'No meetings to export.'; return }
    $dlg = New-Object Windows.Forms.SaveFileDialog
    $dlg.Filter = 'CSV files (*.csv)|*.csv'
    $dlg.FileName = "MeetingAttendees_$(Get-Date -Format yyyyMMdd_HHmm).csv"
    if ($dlg.ShowDialog() -ne 'OK') { return }
    $data = foreach ($row in $rows) {
        $item = Get-TargetItem $row.Tag $rbSeries.Checked
        foreach ($a in Get-Attendees $item) {
            [pscustomobject]@{ Subject = $item.Subject; Start = $item.Start; Name = $a.Name; Email = $a.Email; Type = $a.Type; Response = $a.Response }
        }
    }
    $data | Export-Csv -Path $dlg.FileName -NoTypeInformation
    Write-Log "Exported $(@($data).Count) row(s) to $($dlg.FileName)"
} })

$btnQueueRemove.Add_Click({ Invoke-Safe { Add-QueuedFromAttendeeSelection 'Remove' '' } })
$btnQueueReq.Add_Click({ Invoke-Safe { Add-QueuedFromAttendeeSelection 'Move' 'Required' } })
$btnQueueOpt.Add_Click({ Invoke-Safe { Add-QueuedFromAttendeeSelection 'Move' 'Optional' } })

$btnCopyFrom.Add_Click({ Invoke-Safe {
    $m = Get-FocusedMeeting
    if (-not $m) { Write-Log 'Highlight a source meeting first.'; return }
    foreach ($a in Get-Attendees (Get-TargetItem $m $rbSeries.Checked)) {
        if (-not $a.IsOrganizer) { Add-PendingChange 'Add' $a.Name $a.Email $a.Type }
    }
    Write-Log "Queued attendees from '$($m.Subject)'. Check the target meetings, then Preview/Apply."
} })

$btnAdd.Add_Click({ Invoke-Safe {
    foreach ($entry in ($txtAdd.Text -split '[;,\r\n]+' | Where-Object { $_.Trim() })) {
        $res = Resolve-OutlookAddress $entry
        if ($res) { Add-PendingChange 'Add' $res.Name $res.Email $cboType.SelectedItem }
        else { Write-Log "Could not resolve '$($entry.Trim())' (not found or ambiguous)." }
    }
    $txtAdd.Clear()
} })

$btnAddDL.Add_Click({ Invoke-Safe {
    foreach ($entry in ($txtAdd.Text -split '[;,\r\n]+' | Where-Object { $_.Trim() })) {
        if ($chkExpand.Checked) {
            $members = @(Get-OutlookDistributionListMember $entry)
            foreach ($mem in $members) { Add-PendingChange 'Add' $mem.Name $mem.Email $cboType.SelectedItem }
            Write-Log "Queued $($members.Count) member(s) of '$($entry.Trim())'."
        }
        else {
            $res = Resolve-OutlookAddress $entry
            if ($res) { Add-PendingChange 'Add' $res.Name $res.Email $cboType.SelectedItem }
            else { Write-Log "Could not resolve '$($entry.Trim())'." }
        }
    }
    $txtAdd.Clear()
} })

$btnCsv.Add_Click({ Invoke-Safe {
    $dlg = New-Object Windows.Forms.OpenFileDialog
    $dlg.Filter = 'CSV files (*.csv)|*.csv'
    if ($dlg.ShowDialog() -ne 'OK') { return }
    $rows = @(Import-Csv -Path $dlg.FileName)
    if (-not $rows) { Write-Log 'CSV is empty.'; return }
    if ($rows[0].PSObject.Properties.Name -notcontains 'Email') { throw 'CSV must have an Email column (optional: Type, Action).' }
    $hasType = $rows[0].PSObject.Properties.Name -contains 'Type'
    $hasAction = $rows[0].PSObject.Properties.Name -contains 'Action'
    $errors = @(); $queued = @(); $line = 1
    foreach ($row in $rows) {
        $line++
        $action = if ($hasAction -and $row.Action) { $row.Action.Trim() } else { 'Add' }
        $type = if ($hasType -and $row.Type) { $row.Type.Trim() } else { 'Required' }
        if (-not $row.Email -or -not $row.Email.Trim()) { $errors += "Line ${line}: Email is empty."; continue }
        $action = @('Add', 'Remove', 'Move') | Where-Object { $_ -eq $action }
        $type = @('Required', 'Optional') | Where-Object { $_ -eq $type }
        if (-not $action) { $errors += "Line ${line}: Action must be Add, Remove or Move."; continue }
        if (-not $type) { $errors += "Line ${line}: Type must be Required or Optional."; continue }
        $res = Resolve-OutlookAddress $row.Email
        if (-not $res) {
            if ($action -eq 'Add') { $errors += "Line ${line}: could not resolve '$($row.Email)'."; continue }
            $res = [pscustomobject]@{ Name = ''; Email = $row.Email.Trim() }
        }
        $queued += [pscustomobject]@{ Action = $action; Name = $res.Name; Email = $res.Email; Type = $type }
    }
    if ($errors) {
        $errors | ForEach-Object { Write-Log "CSV $_" }
        throw "CSV has $($errors.Count) error(s); nothing was queued. See log."
    }
    foreach ($q in $queued) { Add-PendingChange $q.Action $q.Name $q.Email $q.Type }
    Write-Log "Queued $($queued.Count) change(s) from CSV."
} })

$btnRemoveChange.Add_Click({ foreach ($row in @($gridChanges.SelectedRows)) { $gridChanges.Rows.Remove($row) } })
$btnClearChanges.Add_Click({ $gridChanges.Rows.Clear() })
$btnPreview.Add_Click({ Invoke-Safe { Invoke-Run $true } })
$btnApply.Add_Click({ Invoke-Safe { Invoke-Run $false } })

#endregion

[void]$form.ShowDialog()
$form.Dispose()
