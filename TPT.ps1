# ============================================================================
# Team Points Tracker - Real Time Points Tracker for Team/Tech
# Created and developed by PIA - members: Pogi, LT, Rona
#
# Do not edit the code without informing PIA
# PIA site: https://eagleviewcloud.sharepoint.com/sites/QualityTeam/SitePages/PIA.aspx
#
# Version: 1.10.7
# ============================================================================

<#
.SYNOPSIS
    Team Points Tracker (TPT) - real-time points for a team / tech (multi-user)
    Reads \\<machine>\c$\Eagleview\Logs\Twister.log on the machines of the techs in scope,
    lists the report IDs each tech worked on (for a date or date range) and their Points
    from the Operations Center API.

.NOTES
    * Run with an account that has admin-share (c$) read access to the machines.
    * Run:  powershell -STA -ExecutionPolicy Bypass -File .\TPT.ps1
            (optional: -RosterPath C:\path\to\TechRoster.csv)
    * Tech ID   = folder name before "TCT"  (7657LGTCT18 -> 7657LG)
    * Report ID = 2nd "_" token of the .evm filename (P3_73665878_IE_[...] -> 73665878)
    * Only numeric Report IDs of 6+ digits are counted.
    * Points    = sum of "points" for the user's taskStateId(s) from the Operations Center API
                  (uses your Windows login). NOTE: points are the user's total on that report,
                  not only the points earned on the searched date. Within one search a report is
                  counted once per user, so use a date range (not several single-day searches)
                  to avoid counting the same report twice.
    * Techs / machines come from TechRoster.csv (same folder as this script).
      Columns: FullName, TechId, User, Team, Machines, Active
        - Machines : separate with ";"  (e.g. MNL-BRKLY-787;MNL-BRKLY-791)
        - Active   : Y / N  (N, No, 0 or False hides the tech; blank = Y)
      Edit the CSV, then click "Reload roster" - no need to restart the tool.

    UI
    * Team / User choose the scope (who and which machines get scanned).
      "All users" = everyone in the selected team.
    * Date = from .. to (inclusive). Changing the first date moves the second one along.
    * Summary table  : one row per tech in scope.  Click row(s) to filter the detail table
                       (Ctrl/Shift-click for several).  Ctrl-click a selected row (or right-click >
                       "Show all details") clears the filter.
    * Detail table   : one row per user + report (machines/times merged if used on several PCs).
    * Right-click any table: copy the cell value or the highlighted row(s).
    * Points that cannot be fetched (no records, user not on the report, API error) are left blank.
    * Your last Team / User choice is remembered per Windows user in
      %APPDATA%\PointsMonitoring\settings.json (first run: All teams / All users).
#>

[CmdletBinding()]
param([string]$RosterPath)

$AppVersion = '1.10.7'
$PiaUrl     = 'https://eagleviewcloud.sharepoint.com/sites/QualityTeam/SitePages/PIA.aspx'

$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
if (-not $RosterPath) { $RosterPath = Join-Path $ScriptDir 'TechRoster.csv' }

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()
[Net.ServicePointManager]::DefaultConnectionLimit = 32

$LogRelativePath     = 'c$\Eagleview\Logs\Twister.log'
$ApiBase             = 'https://api.cmh.platform-prod2.evinternal.net/operations-center/api/Report'
$MaxParallelMachines = 16     # simultaneous SMB log reads
$MaxParallelApi      = 8      # simultaneous report lookups
$MaxRangeDays        = 31
$inv                 = [System.Globalization.CultureInfo]::InvariantCulture

# ----------------------------------------------------------------------------
# Tech roster  (TechRoster.csv: FullName,TechId,User,Team,Machines,Active)
# ----------------------------------------------------------------------------
function Import-Roster {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) { throw "Roster file not found:`r`n$Path" }

    # open with sharing so it still loads while the CSV is open in Excel
    $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    try {
        $sr   = New-Object System.IO.StreamReader($fs, [System.Text.Encoding]::UTF8, $true)
        $text = $sr.ReadToEnd()
    }
    finally { $fs.Close() }

    $rows = @($text | ConvertFrom-Csv)
    if ($rows.Count -eq 0) { throw "Roster file is empty:`r`n$Path" }

    $need    = 'TechId','User','Machines'
    $have    = $rows[0].PSObject.Properties.Name
    $missing = @($need | Where-Object { $have -notcontains $_ })
    if ($missing.Count) { throw "Roster is missing column(s): $($missing -join ', ')`r`nExpected: FullName, TechId, User, Team, Machines, Active" }

    $list = New-Object System.Collections.Generic.List[object]
    $seen = @{}
    $warn = New-Object System.Collections.Generic.List[string]
    $line = 1
    foreach ($r in $rows) {
        $line++
        $techId = ([string]$r.TechId).Trim().ToUpper()
        if (-not $techId -or $techId.StartsWith('#')) { continue }           # blank / commented row

        $active = if ($have -contains 'Active') { ([string]$r.Active).Trim() } else { '' }
        if ($active -match '^(n|no|0|false|inactive)$') { continue }          # switched off

        if ($seen.ContainsKey($techId)) { $warn.Add("line ${line}: duplicate TechId $techId skipped"); continue }
        $user = ([string]$r.User).Trim()
        if (-not $user) { $warn.Add("line ${line}: $techId has no User - skipped"); continue }
        $seen[$techId] = $true

        $full = if ($have -contains 'FullName') { ([string]$r.FullName).Trim() } else { '' }
        $team = if ($have -contains 'Team')     { ([string]$r.Team).Trim() }     else { '' }
        if (-not $full) { $full = $user }
        if (-not $team) { $team = 'Unassigned' }

        $machines = @(([string]$r.Machines) -split '[;,|\s]+' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        if ($machines.Count -eq 0) { $warn.Add("line ${line}: $techId has no machines") }

        $list.Add([pscustomobject]@{
            FullName    = $full
            TechId      = $techId
            User        = $user
            Team        = $team
            Machines    = ($machines -join ', ')
            MachineList = $machines
        })
    }
    if ($list.Count -eq 0) { throw "No active techs found in:`r`n$Path" }

    [pscustomobject]@{ Techs = $list.ToArray(); Warnings = $warn.ToArray() }
}

$RosterWarnings = @()
try {
    $loaded         = Import-Roster -Path $RosterPath
    $script:Techs   = @($loaded.Techs)
    $RosterWarnings = @($loaded.Warnings)
}
catch {
    [void][System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Cannot load roster', 'OK', 'Error')
    return
}

# ----------------------------------------------------------------------------
# Log search - runs in its own runspace (one per machine).
# One pass over the log looks for ALL techs in scope and ALL dates in the range.
# ----------------------------------------------------------------------------
$SearchScript = {
    param([string]$Machine, [string[]]$TechIds, [string[]]$DateStrs, [string]$LogRelativePath)

    $sw        = [System.Diagnostics.Stopwatch]::StartNew()
    $found     = @{}
    $firstDate = $null
    try {
        $path = "\\$Machine\$LogRelativePath"

        # quick SMB (445) check so an offline machine fails in ~3s instead of ~40s
        $tcp = New-Object System.Net.Sockets.TcpClient
        try {
            $iar = $tcp.BeginConnect($Machine, 445, $null, $null)
            if (-not $iar.AsyncWaitHandle.WaitOne(3000)) { throw "Cannot reach $Machine (port 445 not responding)" }
            $tcp.EndConnect($iar)
        }
        finally { $tcp.Close() }
        if (-not (Test-Path -LiteralPath $path)) { throw "Cannot reach $path" }

        $opts = [System.Text.RegularExpressions.RegexOptions]'IgnoreCase, Compiled'
        $alt  = ($TechIds | ForEach-Object { [regex]::Escape($_) }) -join '|'

        # "DEBUG [Name] - 2026-10-07 10:38:42.818 - ..."
        $rxStamp = [regex]::new('^\w+\s+\[[^\]]*\]\s+-\s+(?<date>\d{4}-\d{2}-\d{2})\s+(?<time>\d{2}:\d{2}:\d{2})', $opts)
        # "\7657LGTCT18\InProgress\<file>.evm"  (one or more sub-folders; a folder name can't contain ".evm")
        $rxFile  = [regex]::new('\\(?<tech>' + $alt + ')TCT(?<n>\d+)\\(?:(?:(?!\.evm)[^\\])+\\)+(?<file>[^\\]+?)\.evm', $opts)

        # cheap pre-filter: longest common prefix of the requested dates (whole date for a single day)
        $dateSet = [System.Collections.Generic.HashSet[string]]::new([string[]]$DateStrs)
        $prefix  = $DateStrs[0]
        foreach ($d in $DateStrs) {
            while (-not $d.StartsWith($prefix, [StringComparison]::Ordinal)) { $prefix = $prefix.Substring(0, $prefix.Length - 1) }
        }

        $fs = [System.IO.File]::Open($path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        $sr = New-Object System.IO.StreamReader($fs)
        try {
            while ($null -ne ($line = $sr.ReadLine())) {
                # remember the first date in the log (to detect rotated / truncated logs)
                if ($null -eq $firstDate) {
                    $m0 = $rxStamp.Match($line)
                    if ($m0.Success) { $firstDate = $m0.Groups['date'].Value }
                }
                if ($line.IndexOf($prefix, [StringComparison]::Ordinal) -lt 0) { continue }
                if ($line.IndexOf('TCT',   [StringComparison]::Ordinal) -lt 0) { continue }

                $s = $rxStamp.Match($line)
                if (-not $s.Success -or -not $dateSet.Contains($s.Groups['date'].Value)) { continue }

                $f = $rxFile.Match($line)
                if (-not $f.Success) { continue }

                $file  = $f.Groups['file'].Value
                $parts = $file.Split('_')
                if ($parts.Count -lt 2 -or $parts[1] -notmatch '^\d{6,}$') { continue }

                $rid   = $parts[1]
                $tech  = $f.Groups['tech'].Value.ToUpper()
                $stamp = $s.Groups['date'].Value + ' ' + $s.Groups['time'].Value
                $key   = "$tech|$rid"

                if ($found.ContainsKey($key)) {
                    $e = $found[$key]
                    if ([string]::CompareOrdinal($stamp, $e.First) -lt 0) { $e.First = $stamp }
                    if ([string]::CompareOrdinal($stamp, $e.Last)  -gt 0) { $e.Last  = $stamp }
                } else {
                    $found[$key] = [pscustomobject]@{
                        Tech     = $tech
                        ReportId = $rid
                        Machine  = $Machine
                        First    = $stamp
                        Last     = $stamp
                        File     = $file
                    }
                }
            }
        }
        finally { $sr.Close(); $fs.Close() }

        return [pscustomobject]@{ Machine = $Machine; Results = @($found.Values); FirstLogDate = $firstDate; Error = $null; Seconds = $sw.Elapsed.TotalSeconds }
    }
    catch {
        return [pscustomobject]@{ Machine = $Machine; Results = @(); FirstLogDate = $firstDate; Error = $_.Exception.Message; Seconds = $sw.Elapsed.TotalSeconds }
    }
}

# ----------------------------------------------------------------------------
# Points lookup - runs in its own runspace, one per report ID.
# The two API calls are made once per report; points are then computed for every
# user found on that report.
#   1) /Report/{id}/TaskMeasurementPointUsers -> taskStateId(s) per user
#   2) /Report/{id}/TaskMeasurementPoint      -> sum "points" where taskStateId matches
# ----------------------------------------------------------------------------
$PointsScript = {
    param([string]$ReportId, [string[]]$Users, [string]$ApiBase)

    $log     = New-Object System.Collections.Generic.List[string]
    $results = New-Object System.Collections.Generic.List[object]
    $sw      = [System.Diagnostics.Stopwatch]::StartNew()
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    # GET json and return a flat array. Retries once on timeouts / 5xx (not on 4xx).
    function Get-Items([string]$Url) {
        for ($attempt = 1; $attempt -le 2; $attempt++) {
            $suffix = if ($attempt -gt 1) { ' (retry)' } else { '' }
            $log.Add("    GET $Url$suffix")
            try {
                $raw   = Invoke-RestMethod -Uri $Url -UseDefaultCredentials -TimeoutSec 60
                $items = @($raw | ForEach-Object { $_ })
                $log.Add("      -> OK, $($items.Count) item(s)")
                return ,$items
            }
            catch {
                $code = 0
                try { $code = [int]$_.Exception.Response.StatusCode } catch { }
                $log.Add("      -> FAILED $code : $($_.Exception.Message)")
                if ($attempt -ge 2 -or ($code -ge 400 -and $code -lt 500)) { throw }
                Start-Sleep -Milliseconds 700
            }
        }
    }

    $log.Add("  Report $ReportId (user(s): $($Users -join ', '))")
    try {
        $apiUsers = Get-Items "$ApiBase/$ReportId/TaskMeasurementPointUsers"
        foreach ($u in $apiUsers) { $log.Add("        user: $($u.userName)  taskStateId=$($u.taskStateId)") }

        $idsByUser = @{}
        $anyMatch  = $false
        foreach ($user in $Users) {
            $ids = @($apiUsers | Where-Object { $_.userName -ieq $user } | ForEach-Object { $_.taskStateId })
            $idsByUser[$user] = $ids
            if ($ids.Count -gt 0) { $anyMatch = $true; $log.Add("      [$user] matched taskStateId(s): $($ids -join ', ')") }
            else { $log.Add("      [$user] not found on report") }
        }

        $pts = @()
        if ($anyMatch) { $pts = Get-Items "$ApiBase/$ReportId/TaskMeasurementPoint" }

        foreach ($user in $Users) {
            $ids = $idsByUser[$user]
            if ($ids.Count -eq 0) {
                $results.Add([pscustomobject]@{ User = $user; Points = $null; Error = "user '$user' not found on report" })
                continue
            }
            $sum = 0.0
            foreach ($p in $pts) {
                $hit = $ids -contains $p.taskStateId
                if ($hit) {
                    try { $sum += [double]$p.points } catch { $log.Add("        WARNING: bad points value '$($p.points)'") }
                }
                $mark = if ($hit) { "[MATCH $user]" } else { '' }
                $log.Add(("        {0,-20} {1} | points={2} | taskStateId={3}" -f $mark, $p.description, $p.points, $p.taskStateId))
            }
            $log.Add(("      => total points for {0} = {1}" -f $user, $sum))
            $results.Add([pscustomobject]@{ User = $user; Points = $sum; Error = $null })
        }
        $log.Add("      ({0} ms)" -f [int]$sw.Elapsed.TotalMilliseconds)
    }
    catch {
        $log.Add("      => ERROR: $($_.Exception.Message)")
        $results.Clear()
        foreach ($user in $Users) { $results.Add([pscustomobject]@{ User = $user; Points = $null; Error = $_.Exception.Message }) }
    }
    [pscustomobject]@{ ReportId = $ReportId; Results = $results.ToArray(); Log = $log.ToArray() }
}

function Format-Points($v) { ([double]$v).ToString('0.0#', $inv) }
function Format-Stamp([string]$s, [bool]$multiDay) { if ($multiDay) { $s.Substring(5) } else { $s.Substring(11) } }

# ----------------------------------------------------------------------------
# UI
# ----------------------------------------------------------------------------
# Dark Navy dashboard theme.  UI-only changes live in this section.
$C = @{
    Window      = [System.Drawing.Color]::FromArgb(15, 23, 42)       # #0F172A
    Panel       = [System.Drawing.Color]::FromArgb(23, 32, 51)       # #172033
    Panel2      = [System.Drawing.Color]::FromArgb(30, 41, 59)       # #1E293B
    Input       = [System.Drawing.Color]::FromArgb(37, 50, 71)       # #253247
    Grid        = [System.Drawing.Color]::FromArgb(17, 24, 39)       # #111827
    GridAlt     = [System.Drawing.Color]::FromArgb(24, 33, 51)
    Primary     = [System.Drawing.Color]::FromArgb(37, 99, 235)      # #2563EB
    PrimaryHov  = [System.Drawing.Color]::FromArgb(29, 78, 216)      # #1D4ED8
    Success     = [System.Drawing.Color]::FromArgb(34, 197, 94)      # #22C55E
    Warning     = [System.Drawing.Color]::FromArgb(245, 158, 11)     # #F59E0B
    Danger      = [System.Drawing.Color]::FromArgb(239, 68, 68)      # #EF4444
    Text        = [System.Drawing.Color]::FromArgb(248, 250, 252)    # #F8FAFC
    Muted       = [System.Drawing.Color]::FromArgb(148, 163, 184)    # #94A3B8
    Border      = [System.Drawing.Color]::FromArgb(51, 65, 85)       # #334155
    Selection   = [System.Drawing.Color]::FromArgb(37, 99, 235)
}

$form = New-Object System.Windows.Forms.Form
$form.Text          = 'Team Points Tracker'
$form.ClientSize    = New-Object System.Drawing.Size(1360, 860)
$form.MinimumSize   = New-Object System.Drawing.Size(1180, 700)
$form.StartPosition = 'CenterScreen'
$form.Font          = New-Object System.Drawing.Font('Segoe UI', 9)
$form.BackColor     = $C.Window
$form.ForeColor     = $C.Text

$tip = New-Object System.Windows.Forms.ToolTip
$tip.AutoPopDelay = 20000
$tip.InitialDelay = 350
$tip.ReshowDelay = 100

function Style-Button($b, [string]$kind='secondary') {
    $b.FlatStyle = 'Flat'
    $b.FlatAppearance.BorderSize = 0
    $b.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 9)
    $b.Cursor = [System.Windows.Forms.Cursors]::Hand
    switch ($kind) {
        'primary' {
            $b.BackColor = $C.Primary; $b.ForeColor = $C.Text
            $b.FlatAppearance.MouseOverBackColor = $C.PrimaryHov
            $b.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(30, 64, 175)
        }
        'danger' {
            $b.BackColor = $C.Danger; $b.ForeColor = $C.Text
            $b.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(220, 38, 38)
            $b.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(185, 28, 28)
        }
        'ghost' {
            $b.BackColor = $C.Input; $b.ForeColor = $C.Text
            $b.FlatAppearance.BorderSize = 1
            $b.FlatAppearance.BorderColor = $C.Border
            $b.FlatAppearance.MouseOverBackColor = $C.Panel2
        }
        default {
            $b.BackColor = $C.Panel2; $b.ForeColor = $C.Text
            $b.FlatAppearance.BorderSize = 1
            $b.FlatAppearance.BorderColor = $C.Border
            $b.FlatAppearance.MouseOverBackColor = $C.Input
        }
    }
}

function New-TextLabel([string]$text, [int]$size=9, [System.Drawing.FontStyle]$style='Regular', $color=$null) {
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $text
    $l.AutoSize = $true
    $l.Font = New-Object System.Drawing.Font('Segoe UI', $size, $style)
    $l.ForeColor = if ($null -ne $color) { $color } else { $C.Text }
    $l.BackColor = [System.Drawing.Color]::Transparent
    $l
}

# Main dashboard: [header] [filters] [results] [status] [log]
$main = New-Object System.Windows.Forms.TableLayoutPanel
$main.Dock = 'Fill'
$main.ColumnCount = 1
$main.RowCount = 5
$main.BackColor = $C.Window
[void]$main.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$main.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 76)))
[void]$main.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 150)))
[void]$main.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$main.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 42)))
[void]$main.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 166)))
$form.Controls.Add($main)

# ---- header -----------------------------------------------------------------
$header = New-Object System.Windows.Forms.TableLayoutPanel
$header.Dock = 'Fill'
$header.BackColor = $C.Panel
$header.Padding = New-Object System.Windows.Forms.Padding(24, 10, 22, 10)
$header.ColumnCount = 2
$header.RowCount = 1
[void]$header.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 78)))
[void]$header.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 22)))
$main.Controls.Add($header, 0, 0)

$headerLeft = New-Object System.Windows.Forms.TableLayoutPanel
$headerLeft.Dock = 'Fill'; $headerLeft.ColumnCount = 1; $headerLeft.RowCount = 2
[void]$headerLeft.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,100)))
[void]$headerLeft.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,34)))
[void]$headerLeft.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent,100)))
$header.Controls.Add($headerLeft,0,0)

$lblTitle = New-TextLabel 'TEAM POINTS TRACKER' 18 ([System.Drawing.FontStyle]::Bold)
$lblTitle.Dock = 'Fill'; $lblTitle.TextAlign = 'MiddleLeft'; $headerLeft.Controls.Add($lblTitle,0,0)
$lblSubtitle = New-TextLabel 'Real-time technician and report monitoring' 8.5 ([System.Drawing.FontStyle]::Regular) $C.Muted
$lblSubtitle.Dock = 'Fill'; $lblSubtitle.TextAlign = 'MiddleLeft'; $lblSubtitle.Padding = New-Object System.Windows.Forms.Padding(1,0,0,0); $headerLeft.Controls.Add($lblSubtitle,0,1)

$headerRight = New-Object System.Windows.Forms.TableLayoutPanel
$headerRight.Dock = 'Fill'; $headerRight.ColumnCount = 2; $headerRight.RowCount = 1
[void]$headerRight.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute,130)))
[void]$headerRight.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,100)))
$header.Controls.Add($headerRight,1,0)

$lblReady = New-Object System.Windows.Forms.Label
$lblReady.Text = 'READY'; $lblReady.Dock='Fill'; $lblReady.TextAlign='MiddleCenter'
$lblReady.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 9)
$lblReady.ForeColor = $C.Success; $lblReady.BackColor = $C.Panel2
$lblReady.Margin = New-Object System.Windows.Forms.Padding(8, 8, 6, 8)
$headerRight.Controls.Add($lblReady,0,0)
$lblVersion = New-TextLabel "v$AppVersion" 8.5 ([System.Drawing.FontStyle]::Regular) $C.Muted
$lblVersion.Dock='Fill'; $lblVersion.TextAlign='MiddleRight'; $lblVersion.Padding=New-Object System.Windows.Forms.Padding(0,0,2,0); $headerRight.Controls.Add($lblVersion,1,0)

# ---- filters ----------------------------------------------------------------
$filterHost = New-Object System.Windows.Forms.Panel
$filterHost.Dock = 'Fill'; $filterHost.BackColor = $C.Window
$filterHost.Padding = New-Object System.Windows.Forms.Padding(18, 8, 18, 8)
$main.Controls.Add($filterHost,0,1)

$filters = New-Object System.Windows.Forms.TableLayoutPanel
$filters.Dock='Fill'; $filters.BackColor=$C.Panel; $filters.Padding=New-Object System.Windows.Forms.Padding(16,7,16,7)
$filters.ColumnCount=8; $filters.RowCount=3
$colPct=@(8,20,8,20,8,18,9,9)
for($i=0;$i -lt 8;$i++){[void]$filters.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,$colPct[$i])))}
[void]$filters.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,44)))
[void]$filters.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,42)))
[void]$filters.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent,100)))
$filterHost.Controls.Add($filters)

$lblTeam = New-TextLabel 'TEAM' 8 ([System.Drawing.FontStyle]::Bold) $C.Muted; $lblTeam.Dock='Fill'; $lblTeam.TextAlign='MiddleLeft'; $filters.Controls.Add($lblTeam,0,0)
$cboTeam = New-Object System.Windows.Forms.ComboBox
$cboTeam.DropDownStyle='DropDownList'; $cboTeam.Dock='Fill'; $cboTeam.FlatStyle='Flat'; $cboTeam.BackColor=$C.Input; $cboTeam.ForeColor=$C.Text; $cboTeam.Font=New-Object System.Drawing.Font('Segoe UI',9)
$tip.SetToolTip($cboTeam,'Choose the team to include in the scan')
$filters.Controls.Add($cboTeam,1,0); $filters.SetColumnSpan($cboTeam,1)
$lblUser = New-TextLabel 'USER' 8 ([System.Drawing.FontStyle]::Bold) $C.Muted; $lblUser.Dock='Fill'; $lblUser.TextAlign='MiddleLeft'; $filters.Controls.Add($lblUser,2,0)
$cboUser = New-Object System.Windows.Forms.ComboBox
$cboUser.DropDownStyle='DropDownList'; $cboUser.Dock='Fill'; $cboUser.DropDownWidth=420; $cboUser.FlatStyle='Flat'; $cboUser.BackColor=$C.Input; $cboUser.ForeColor=$C.Text; $cboUser.Font=New-Object System.Drawing.Font('Segoe UI',9)
$tip.SetToolTip($cboUser,'Choose a user or keep All users selected')
$filters.Controls.Add($cboUser,3,0)

$lblFrom = New-TextLabel 'FROM' 8 ([System.Drawing.FontStyle]::Bold) $C.Muted; $lblFrom.Dock='Fill'; $lblFrom.TextAlign='MiddleLeft'; $filters.Controls.Add($lblFrom,4,0)
$dtpFrom=New-Object System.Windows.Forms.DateTimePicker
$dtpFrom.Format='Short'; $dtpFrom.Dock='Fill'; $dtpFrom.Value=Get-Date; $dtpFrom.Font=New-Object System.Drawing.Font('Segoe UI',9); $dtpFrom.BackColor=$C.Input; $dtpFrom.ForeColor=$C.Text
$filters.Controls.Add($dtpFrom,5,0)
$lblTo = New-TextLabel 'TO' 8 ([System.Drawing.FontStyle]::Bold) $C.Muted; $lblTo.Dock='Fill'; $lblTo.TextAlign='MiddleLeft'; $filters.Controls.Add($lblTo,6,0)
$dtpTo=New-Object System.Windows.Forms.DateTimePicker
$dtpTo.Format='Short'; $dtpTo.Dock='Fill'; $dtpTo.Value=Get-Date; $dtpTo.Font=New-Object System.Drawing.Font('Segoe UI',9); $dtpTo.BackColor=$C.Input; $dtpTo.ForeColor=$C.Text
$filters.Controls.Add($dtpTo,7,0)
$tip.SetToolTip($dtpFrom,'First day (inclusive). Use a range for shifts that cross midnight.')
$tip.SetToolTip($dtpTo,'Last day (inclusive)')

$lblQuick = New-TextLabel 'QUICK RANGE' 8 ([System.Drawing.FontStyle]::Bold) $C.Muted; $lblQuick.Dock='Fill'; $lblQuick.TextAlign='MiddleLeft'; $filters.Controls.Add($lblQuick,0,1)
$quickPanel=New-Object System.Windows.Forms.FlowLayoutPanel; $quickPanel.Dock='Fill'; $quickPanel.BackColor=[System.Drawing.Color]::Transparent; $quickPanel.WrapContents=$false; $quickPanel.Padding=New-Object System.Windows.Forms.Padding(0,4,0,0)
$filters.Controls.Add($quickPanel,1,1)
$btnToday=New-Object System.Windows.Forms.Button; $btnToday.Text='Today'; $btnToday.Size=New-Object System.Drawing.Size(72,28); Style-Button $btnToday 'secondary'; $quickPanel.Controls.Add($btnToday)
$btnYest=New-Object System.Windows.Forms.Button; $btnYest.Text='Yesterday'; $btnYest.Size=New-Object System.Drawing.Size(88,28); $btnYest.Margin=New-Object System.Windows.Forms.Padding(8,0,0,0); Style-Button $btnYest 'secondary'; $quickPanel.Controls.Add($btnYest)

$lblScopeTitle=New-TextLabel 'SCOPE' 8 ([System.Drawing.FontStyle]::Bold) $C.Muted; $lblScopeTitle.Dock='Fill'; $lblScopeTitle.TextAlign='MiddleLeft'; $filters.Controls.Add($lblScopeTitle,2,1)
$lblScope=New-TextLabel '25 technicians  |  18 machines' 9 ([System.Drawing.FontStyle]::Bold) $C.Text; $lblScope.Dock='Fill'; $lblScope.TextAlign='MiddleLeft'; $filters.Controls.Add($lblScope,3,1); $filters.SetColumnSpan($lblScope,3)

$btnRoster=New-Object System.Windows.Forms.Button; $btnRoster.Text='Open Roster'; $btnRoster.Dock='Fill'; $btnRoster.Margin=New-Object System.Windows.Forms.Padding(8,2,4,2); Style-Button $btnRoster 'ghost'; $filters.Controls.Add($btnRoster,6,1)
$btnReload=New-Object System.Windows.Forms.Button; $btnReload.Text='Reload'; $btnReload.Dock='Fill'; $btnReload.Margin=New-Object System.Windows.Forms.Padding(4,2,0,2); Style-Button $btnReload 'ghost'; $filters.Controls.Add($btnReload,7,1)

$lblHint = New-TextLabel 'Select filters, then search the selected machines and date range.' 8.5 ([System.Drawing.FontStyle]::Regular) $C.Muted
$lblHint.Dock='Fill'; $lblHint.TextAlign='MiddleLeft'; $filters.Controls.Add($lblHint,0,2); $filters.SetColumnSpan($lblHint,5)
$btnRun=New-Object System.Windows.Forms.Button; $btnRun.Text='SEARCH FILES'; $btnRun.Dock='Fill'; $btnRun.Margin=New-Object System.Windows.Forms.Padding(8,4,0,4); Style-Button $btnRun 'primary'; $filters.Controls.Add($btnRun,5,2); $filters.SetColumnSpan($btnRun,3)
$form.AcceptButton=$btnRun

# ---- grids ------------------------------------------------------------------
function New-Grid($defs) {
    $g=New-Object System.Windows.Forms.DataGridView
    $g.Dock='Fill'; $g.ReadOnly=$true; $g.AllowUserToAddRows=$false; $g.AllowUserToDeleteRows=$false
    $g.AllowUserToResizeRows=$false; $g.SelectionMode='FullRowSelect'; $g.MultiSelect=$true
    $g.AutoSizeColumnsMode='Fill'; $g.RowHeadersVisible=$false; $g.BackgroundColor=$C.Grid; $g.BorderStyle='None'
    $g.GridColor=$C.Border; $g.EnableHeadersVisualStyles=$false; $g.AutoGenerateColumns=$false
    $g.ColumnHeadersHeight=32; $g.ColumnHeadersBorderStyle='Single'; $g.RowTemplate.Height=27
    $g.ColumnHeadersDefaultCellStyle.BackColor=$C.Panel2; $g.ColumnHeadersDefaultCellStyle.ForeColor=$C.Muted
    $g.ColumnHeadersDefaultCellStyle.Font=New-Object System.Drawing.Font('Segoe UI Semibold',8.5)
    $g.ColumnHeadersDefaultCellStyle.SelectionBackColor=$C.Panel2; $g.ColumnHeadersDefaultCellStyle.SelectionForeColor=$C.Text
    $g.DefaultCellStyle.BackColor=$C.Grid; $g.DefaultCellStyle.ForeColor=$C.Text
    $g.DefaultCellStyle.SelectionBackColor=$C.Selection; $g.DefaultCellStyle.SelectionForeColor=$C.Text
    $g.DefaultCellStyle.Padding=New-Object System.Windows.Forms.Padding(8,0,8,0)
    $g.AlternatingRowsDefaultCellStyle.BackColor=$C.GridAlt; $g.AlternatingRowsDefaultCellStyle.ForeColor=$C.Text
    $g.DefaultCellStyle.NullValue=''
    foreach($d in $defs){
        $c=New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $c.Name=$d.Name; $c.HeaderText=if($d.Header){$d.Header}else{$d.Name}; $c.DataPropertyName=$d.Name; $c.FillWeight=$d.Weight
        $align=[System.Windows.Forms.DataGridViewContentAlignment]$d.Align; $c.DefaultCellStyle.Alignment=$align
        if($d.Format){$c.DefaultCellStyle.Format=$d.Format; $c.DefaultCellStyle.FormatProvider=$inv}
        if($d.Indent){$c.DefaultCellStyle.Padding=New-Object System.Windows.Forms.Padding($d.Indent,0,8,0)}
        [void]$g.Columns.Add($c); $g.Columns[$d.Name].HeaderCell.Style.Alignment=$align
    }
    $g
}

# summary data (Reports/Points are numeric so they sort properly; State/Flag drive the display)
$dtS=New-Object System.Data.DataTable
[void]$dtS.Columns.Add('Team',[string]); [void]$dtS.Columns.Add('TechId',[string]); [void]$dtS.Columns.Add('User',[string])
[void]$dtS.Columns.Add('Reports',[int]); [void]$dtS.Columns.Add('Points',[double]); [void]$dtS.Columns.Add('State',[string]); [void]$dtS.Columns.Add('Flag',[string])

# detail data
$dtD=New-Object System.Data.DataTable
[void]$dtD.Columns.Add('User',[string]); [void]$dtD.Columns.Add('Report ID',[string]); [void]$dtD.Columns.Add('Points',[double])
[void]$dtD.Columns.Add('Machine',[string]); [void]$dtD.Columns.Add('File',[string]); [void]$dtD.Columns.Add('First Seen',[string]); [void]$dtD.Columns.Add('Last Seen',[string]); [void]$dtD.Columns.Add('State',[string])

$gridS=New-Grid @(
    @{Name='Team';Align='MiddleLeft';Weight=105;Indent=8},
    @{Name='TechId';Header='Tech ID';Align='MiddleCenter';Weight=80},
    @{Name='User';Align='MiddleLeft';Weight=175},
    @{Name='Reports';Align='MiddleCenter';Weight=75},
    @{Name='Points';Align='MiddleCenter';Weight=90;Format='0.0#'}
)
$gridD=New-Grid @(
    @{Name='User';Align='MiddleLeft';Weight=125}, @{Name='Report ID';Align='MiddleCenter';Weight=80},
    @{Name='Points';Align='MiddleCenter';Weight=65;Format='0.0#'}, @{Name='Machine';Align='MiddleLeft';Weight=130},
    @{Name='File';Align='MiddleLeft';Weight=250}, @{Name='First Seen';Align='MiddleCenter';Weight=90}, @{Name='Last Seen';Align='MiddleCenter';Weight=90}
)

function New-GridPanel($titleLabel,$button,$grid) {
    $t=New-Object System.Windows.Forms.TableLayoutPanel; $t.Dock='Fill'; $t.BackColor=$C.Panel; $t.ColumnCount=2; $t.RowCount=2
    [void]$t.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,100)))
    [void]$t.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute,120)))
    [void]$t.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,38)))
    [void]$t.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent,100)))
    $titleLabel.Dock='Fill'; $titleLabel.TextAlign='MiddleLeft'; $titleLabel.Padding=New-Object System.Windows.Forms.Padding(14,0,0,0)
    $titleLabel.Font=New-Object System.Drawing.Font('Segoe UI Semibold',10)
    $button.Dock='Fill'; $button.Margin=New-Object System.Windows.Forms.Padding(4,6,10,5); Style-Button $button 'ghost'
    $t.Controls.Add($titleLabel,0,0); $t.Controls.Add($button,1,0); $t.Controls.Add($grid,0,1); $t.SetColumnSpan($grid,2)
    $t
}

$btnCsvS=New-Object System.Windows.Forms.Button; $btnCsvS.Text='Export CSV'; $btnCsvS.Size=New-Object System.Drawing.Size(110,26); $tip.SetToolTip($btnCsvS,'Export the summary table to CSV')
$btnCsvD=New-Object System.Windows.Forms.Button; $btnCsvD.Text='Export CSV'; $btnCsvD.Size=New-Object System.Drawing.Size(110,26); $tip.SetToolTip($btnCsvD,'Export the detail rows currently shown')

$gbS=New-Object System.Windows.Forms.Panel; $gbS.Dock='Fill'; $gbS.BackColor=$C.Panel; $gbS.Padding=New-Object System.Windows.Forms.Padding(1,0,1,1)
$lblSummaryHeader=New-TextLabel 'SUMMARY  |  0 TECHS  |  0 REPORTS  |  0 POINTS' 10 ([System.Drawing.FontStyle]::Bold)
$gbS.Controls.Add((New-GridPanel $lblSummaryHeader $btnCsvS $gridS))

$gbD=New-Object System.Windows.Forms.Panel; $gbD.Dock='Fill'; $gbD.BackColor=$C.Panel; $gbD.Padding=New-Object System.Windows.Forms.Padding(1,0,1,1)
$lblDetailsHeader=New-TextLabel 'DETAILS  |  0 REPORT ROWS' 10 ([System.Drawing.FontStyle]::Bold)
$gbD.Controls.Add((New-GridPanel $lblDetailsHeader $btnCsvD $gridD))

$split=New-Object System.Windows.Forms.SplitContainer; $split.Dock='Fill'; $split.Orientation='Horizontal'; $split.BackColor=$C.Border
$split.Panel1.Padding=New-Object System.Windows.Forms.Padding(18,0,18,6); $split.Panel2.Padding=New-Object System.Windows.Forms.Padding(18,6,18,0)
$split.Panel1.Controls.Add($gbS); $split.Panel2.Controls.Add($gbD); $main.Controls.Add($split,0,2)
$form.Add_Shown({
    try {
        $avail = $split.Height - $split.SplitterWidth
        if ($avail -gt 330) { $split.Panel1MinSize=130; $split.Panel2MinSize=200 }
        $d = [int]($avail*0.39)
        $d = [Math]::Max($split.Panel1MinSize, [Math]::Min($d, $avail - $split.Panel2MinSize))
        $split.SplitterDistance = $d
    } catch {}
})

# ---- status bar --------------------------------------------------------------
$status = New-Object System.Windows.Forms.TableLayoutPanel
$status.Dock='Fill'; $status.BackColor=$C.Panel2; $status.ColumnCount=4; $status.RowCount=1; $status.Padding=New-Object System.Windows.Forms.Padding(18,0,18,0)
[void]$status.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,48)))
[void]$status.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,17)))
[void]$status.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,17)))
[void]$status.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,18)))
$main.Controls.Add($status,0,3)

$lblSummary=New-TextLabel 'Ready.' 8.7 ([System.Drawing.FontStyle]::Bold) $C.Muted; $lblSummary.Dock='Fill'; $lblSummary.TextAlign='MiddleLeft'; $lblSummary.AutoEllipsis=$true
$status.Controls.Add($lblSummary,0,0)
$lblLogState=New-TextLabel 'LOGS  |  0 / 0' 8.5 ([System.Drawing.FontStyle]::Bold) $C.Muted; $lblLogState.Dock='Fill'; $lblLogState.TextAlign='MiddleCenter'; $status.Controls.Add($lblLogState,1,0)
$lblPointState=New-TextLabel 'POINTS  |  0 / 0' 8.5 ([System.Drawing.FontStyle]::Bold) $C.Muted; $lblPointState.Dock='Fill'; $lblPointState.TextAlign='MiddleCenter'; $status.Controls.Add($lblPointState,2,0)
$btnLog=New-Object System.Windows.Forms.Button; $btnLog.Text='Show Log'; $btnLog.Dock='Fill'; $btnLog.Margin=New-Object System.Windows.Forms.Padding(6,7,0,7); Style-Button $btnLog 'ghost'; $status.Controls.Add($btnLog,3,0)

# ---- bottom log --------------------------------------------------------------
$logHost=New-Object System.Windows.Forms.Panel; $logHost.Dock='Fill'; $logHost.BackColor=$C.Window; $logHost.Padding=New-Object System.Windows.Forms.Padding(18,6,18,0); $main.Controls.Add($logHost,0,4)
$txtLog=New-Object System.Windows.Forms.TextBox; $txtLog.Dock='Fill'; $txtLog.Multiline=$true; $txtLog.ReadOnly=$true; $txtLog.ScrollBars='Vertical'; $txtLog.Font=New-Object System.Drawing.Font('Consolas',8.5); $txtLog.BackColor=$C.Grid; $txtLog.ForeColor=$C.Muted; $txtLog.BorderStyle='FixedSingle'; $txtLog.Visible=$false
$logHost.Controls.Add($txtLog)

$btnLog.Add_Click({
    $txtLog.Visible = -not $txtLog.Visible
    $btnLog.Text = if ($txtLog.Visible) { 'Hide Log' } else { 'Show Log' }
    $main.RowStyles[4].Height = if ($txtLog.Visible) { 150 } else { 0 }
})
$main.RowStyles[4].Height = 0

function Write-Status([string]$msg) {
    $line='[{0}] {1}' -f (Get-Date -Format 'HH:mm:ss'),$msg
    $txtLog.AppendText($line+"`r`n")
}

# ---- state ---------------------------------------------------------------------
$script:Loading     = $false
$script:Populating  = $false
$script:Running     = $false
$script:UserItems   = @()
$script:SyncDate    = $false
$script:ToFollows   = $true

$script:Pool = $null; $script:Jobs = @()
$script:PPool = $null; $script:PJobs = @()
$script:All = $null
$script:Scope = @()
$script:TechUser = @{}
$script:Dates = @()
$script:DateLabel = ''
$script:Stopwatch = $null
$script:ReportUsers = @{}
$script:TechReports = @{}
$script:DetailRows = @{}
$script:SummaryRows = @{}
$script:PointsMap = @{}
$script:Failed  = New-Object 'System.Collections.Generic.HashSet[string]'
$script:Partial = New-Object 'System.Collections.Generic.HashSet[string]'
$script:NaCount = 0
$script:PtsDone = 0
$script:PtsTotal = 0
$script:SumBase = ''; $script:SumProgress = ''; $script:SumTotal = ''

# ---- display helpers -------------------------------------------------------------
function Update-SummaryLabel {
    $parts = @()
    if ($script:SumBase)     { $parts += $script:SumBase }
    if ($script:SumProgress) { $parts += $script:SumProgress }
    if ($script:SumTotal)    { $parts += $script:SumTotal }
    $txt = $parts -join '   |   '
    if (-not $txt) { $txt = 'Ready.' }
    $lblSummary.ForeColor = if ($script:Running) { $C.Text } else { $C.Muted }
    $lblSummary.Text = $txt
    $tip.SetToolTip($lblSummary, $txt)

    $logDone = if ($script:Jobs) { @($script:Jobs | Where-Object { $_.Done }).Count } else { 0 }
    $logTotal = if ($script:Jobs) { @($script:Jobs).Count } else { 0 }
    $lblLogState.Text = "LOGS  |  $logDone / $logTotal"
    $lblLogState.ForeColor = if ($logTotal -gt 0 -and $logDone -eq $logTotal) { $C.Success } elseif ($script:Running) { $C.Primary } else { $C.Muted }

    $pDone = [int]$script:PtsDone; $pTotal = [int]$script:PtsTotal
    $lblPointState.Text = "POINTS  |  $pDone / $pTotal"
    $lblPointState.ForeColor = if ($pTotal -gt 0 -and $pDone -eq $pTotal) { $C.Success } elseif ($script:Running) { $C.Primary } else { $C.Muted }
    $lblReady.Text = if ($script:Running) { 'RUNNING' } else { 'READY' }
    $lblReady.ForeColor = if ($script:Running) { $C.Primary } else { $C.Success }
}

function Update-Titles {
    $reports=0; $pts=0.0
    foreach($r in $dtS.Rows){
        if($r['Reports'] -isnot [DBNull]){$reports += [int]$r['Reports']}
        if($r['Points'] -isnot [DBNull]){$pts += [double]$r['Points']}
    }
    $lblSummaryHeader.Text = 'SUMMARY  |  {0} TECHS  |  {1} REPORTS  |  {2} POINTS' -f $dtS.Rows.Count,$reports,(Format-Points $pts)
    $vis=$dtD.DefaultView.Count; $tot=$dtD.Rows.Count
    if($dtD.DefaultView.RowFilter){$lblDetailsHeader.Text="DETAILS  |  SHOWING $vis OF $tot  |  FILTERED"} else {$lblDetailsHeader.Text="DETAILS  |  $tot REPORT ROWS"}
}

function Apply-DetailFilter {
    if ($script:Populating) { return }
    $users = @($gridS.SelectedRows | ForEach-Object { $_.DataBoundItem.Row['User'] } | Select-Object -Unique)
    $f = ''
    if ($users.Count -gt 0) {
        $f = '[User] IN (' + (($users | ForEach-Object { "'" + ($_ -replace "'", "''") + "'" }) -join ',') + ')'
    }
    if ($dtD.DefaultView.RowFilter -ne $f) { $dtD.DefaultView.RowFilter = $f }
    Update-Titles
}

# colour / text rules for the two grids
$gridS.Add_CellFormatting({
    param($sender, $e)
    if ($e.RowIndex -lt 0) { return }
    $drv = $sender.Rows[$e.RowIndex].DataBoundItem
    if ($null -eq $drv) { return }
    $row  = $drv.Row
    $name = $sender.Columns[$e.ColumnIndex].Name
    if ($name -eq 'Reports' -or $name -eq 'Points') {
        if ($row[$name] -is [DBNull]) { $e.Value = [string]$row['State']; $e.FormattingApplied = $true }
    }
    if ($row['Reports'] -isnot [DBNull] -and [int]$row['Reports'] -eq 0) {
        $e.CellStyle.ForeColor = [System.Drawing.Color]::Gray
    }
})
$gridD.Add_CellFormatting({
    param($sender, $e)
    if ($e.RowIndex -lt 0) { return }
    $drv = $sender.Rows[$e.RowIndex].DataBoundItem
    if ($null -eq $drv) { return }
    $row = $drv.Row
    if ($sender.Columns[$e.ColumnIndex].Name -eq 'Points' -and $row['Points'] -is [DBNull]) {
        $e.Value = [string]$row['State']; $e.FormattingApplied = $true
    }
})
# after (re)binding, the grid selects row 1 by itself - clear it so the detail table starts unfiltered
$gridS.Add_DataBindingComplete({
    param($sender, $e)
    $sender.ClearSelection()
    try { $sender.CurrentCell = $null } catch { }
})
$gridS.Add_SelectionChanged({ Apply-DetailFilter })

# ---- right-click menu (both tables): copy cell / highlighted rows ----------------------------------
$script:CtxRow = -1
$script:CtxCol = -1
function Set-ClipboardText([string]$text) {
    if ([string]::IsNullOrEmpty($text)) { return }
    try { [System.Windows.Forms.Clipboard]::SetText($text) } catch { }
}

$ctxMenu = New-Object System.Windows.Forms.ContextMenuStrip
$miCell  = $ctxMenu.Items.Add('Copy cell value')
$miRows  = $ctxMenu.Items.Add('Copy highlighted row(s)')
$miSep   = New-Object System.Windows.Forms.ToolStripSeparator
[void]$ctxMenu.Items.Add($miSep)
$miAll   = $ctxMenu.Items.Add('Show all details')
$gridS.ContextMenuStrip = $ctxMenu
$gridD.ContextMenuStrip = $ctxMenu

$gridMouseDown = {
    param($sender, $e)
    if ($e.Button -ne [System.Windows.Forms.MouseButtons]::Right) { return }
    $script:CtxRow = -1; $script:CtxCol = -1
    $h = $sender.HitTest($e.X, $e.Y)
    if ($h.Type -eq 'Cell') {
        $script:CtxRow = $h.RowIndex; $script:CtxCol = $h.ColumnIndex
        # keep an existing multi-row highlight if the clicked row is part of it
        if (-not $sender.Rows[$h.RowIndex].Selected) {
            $sender.ClearSelection()
            $sender.Rows[$h.RowIndex].Selected = $true
        }
    }
}
$gridS.Add_MouseDown($gridMouseDown)
$gridD.Add_MouseDown($gridMouseDown)

$ctxMenu.Add_Opening({
    param($sender, $e)
    $g = $sender.SourceControl
    $isSummary = ($g -eq $gridS)
    $miCell.Enabled = ($script:CtxRow -ge 0)
    $miRows.Enabled = ($null -ne $g -and $g.SelectedRows.Count -gt 0)
    $miSep.Visible  = $isSummary
    $miAll.Visible  = $isSummary
})
$miCell.Add_Click({
    $g = $ctxMenu.SourceControl
    if ($g -and $script:CtxRow -ge 0 -and $script:CtxRow -lt $g.Rows.Count -and $script:CtxCol -ge 0) {
        Set-ClipboardText ([string]$g.Rows[$script:CtxRow].Cells[$script:CtxCol].FormattedValue)
    }
})
$miRows.Add_Click({
    $g = $ctxMenu.SourceControl
    if (-not $g) { return }
    $lines = foreach ($r in ($g.SelectedRows | Sort-Object Index)) {
        (@($g.Columns | ForEach-Object { [string]$r.Cells[$_.Index].FormattedValue }) -join "`t")
    }
    Set-ClipboardText (@($lines) -join "`r`n")
})
$miAll.Add_Click({
    $gridS.ClearSelection()
    try { $gridS.CurrentCell = $null } catch { }
    Apply-DetailFilter
})

# ---- Team / User scope ---------------------------------------------------------
function Get-ScopeTechs {
    $i = $cboUser.SelectedIndex
    if ($i -gt 0 -and $i -le $script:UserItems.Count) { return @($script:UserItems[$i - 1]) }
    return @($script:UserItems)
}

$SettingsPath = Join-Path ([Environment]::GetFolderPath('ApplicationData')) 'PointsMonitoring\settings.json'

function Get-Settings {
    try { if (Test-Path -LiteralPath $SettingsPath) { return (Get-Content -LiteralPath $SettingsPath -Raw | ConvertFrom-Json) } } catch { }
    [pscustomobject]@{ Team = ''; TechId = '' }
}

function Save-Settings {
    try {
        $team = [string]$cboTeam.SelectedItem
        if ($team -eq 'All teams') { $team = '' }
        $i   = $cboUser.SelectedIndex
        $tid = if ($i -gt 0 -and $i -le $script:UserItems.Count) { $script:UserItems[$i - 1].TechId } else { '' }
        $dir = Split-Path -Parent $SettingsPath
        if (-not (Test-Path -LiteralPath $dir)) { [void](New-Item -ItemType Directory -Path $dir -Force) }
        [pscustomobject]@{ Team = $team; TechId = $tid } | ConvertTo-Json | Set-Content -LiteralPath $SettingsPath -Encoding UTF8
    } catch { }
}

function Update-Scope {
    $sel      = @(Get-ScopeTechs)
    $machines = @($sel | ForEach-Object { $_.MachineList } | Select-Object -Unique)
    $lblScope.Text = "{0} technicians  |  {1} machines" -f $sel.Count, $machines.Count
    $tip.SetToolTip($lblScope, ($machines -join ', '))
}

function Update-UserList {
    $prevId = $null
    $i = $cboUser.SelectedIndex
    if ($i -gt 0 -and $i -le $script:UserItems.Count) { $prevId = $script:UserItems[$i - 1].TechId }

    $team = [string]$cboTeam.SelectedItem
    $script:UserItems = @($script:Techs | Where-Object { $team -eq 'All teams' -or $_.Team -eq $team } | Sort-Object FullName)

    $script:Loading = $true
    $cboUser.BeginUpdate()
    $cboUser.Items.Clear()
    [void]$cboUser.Items.Add(("All users ({0})" -f $script:UserItems.Count))
    $sel = 0
    for ($k = 0; $k -lt $script:UserItems.Count; $k++) {
        $t = $script:UserItems[$k]
        [void]$cboUser.Items.Add(("{0}  ({1})" -f $t.FullName, $t.TechId))
        if ($prevId -and $t.TechId -eq $prevId) { $sel = $k + 1 }
    }
    $cboUser.SelectedIndex = $sel
    $cboUser.EndUpdate()
    $script:Loading = $false
    Update-Scope
}

function Update-TeamList {
    $prev = [string]$cboTeam.SelectedItem
    $script:Loading = $true
    $cboTeam.Items.Clear()
    [void]$cboTeam.Items.Add('All teams')
    $script:Techs | Select-Object -ExpandProperty Team -Unique | Sort-Object | ForEach-Object { [void]$cboTeam.Items.Add($_) }
    if ($prev -and $cboTeam.Items.Contains($prev)) { $cboTeam.SelectedItem = $prev } else { $cboTeam.SelectedIndex = 0 }
    $script:Loading = $false
}

$cboTeam.Add_SelectedIndexChanged({ if (-not $script:Loading) { Update-UserList } })
$cboUser.Add_SelectedIndexChanged({ if (-not $script:Loading) { Update-Scope } })

# ---- dates ------------------------------------------------------------------------------
$dtpFrom.Add_ValueChanged({
    if ($script:SyncDate) { return }
    if ($script:ToFollows -or $dtpTo.Value.Date -lt $dtpFrom.Value.Date) {
        $script:SyncDate = $true; $dtpTo.Value = $dtpFrom.Value; $script:SyncDate = $false
    }
})
$dtpTo.Add_ValueChanged({
    if ($script:SyncDate) { return }
    $script:ToFollows = ($dtpTo.Value.Date -eq $dtpFrom.Value.Date)
    if ($dtpTo.Value.Date -lt $dtpFrom.Value.Date) {
        $script:SyncDate = $true; $dtpFrom.Value = $dtpTo.Value; $script:SyncDate = $false
        $script:ToFollows = $true
    }
})
function Set-BothDates([datetime]$d) {
    $script:SyncDate = $true
    $dtpFrom.Value = $d; $dtpTo.Value = $d
    $script:SyncDate = $false
    $script:ToFollows = $true
}
$btnToday.Add_Click({ Set-BothDates (Get-Date) })
$btnYest.Add_Click({  Set-BothDates ((Get-Date).AddDays(-1)) })

# ---- roster buttons ----------------------------------------------------------------------
$btnRoster.Add_Click({
    try { Start-Process -FilePath $RosterPath }
    catch { [void][System.Windows.Forms.MessageBox]::Show("Cannot open:`r`n$RosterPath`r`n$($_.Exception.Message)", 'Open roster', 'OK', 'Warning') }
})
$btnReload.Add_Click({
    try { $loaded = Import-Roster -Path $RosterPath }
    catch {
        [void][System.Windows.Forms.MessageBox]::Show($_.Exception.Message + "`r`n`r`n(keeping the previous roster)", 'Cannot reload roster', 'OK', 'Error')
        return
    }
    $script:Techs = @($loaded.Techs)
    Update-TeamList
    Update-UserList
    Write-Status ("Roster reloaded: {0} active tech(s) from {1}" -f $script:Techs.Count, $RosterPath)
    foreach ($w in $loaded.Warnings) { Write-Status "  roster: $w" }
})

# ---- search -----------------------------------------------------------------------------------
$timer  = New-Object System.Windows.Forms.Timer
$timer.Interval = 250
$ptimer = New-Object System.Windows.Forms.Timer
$ptimer.Interval = 250

function Reset-SearchUi {
    $script:Running = $false
    $btnRun.Text = 'SEARCH FILES'
    Style-Button $btnRun 'primary'
    $btnRun.Enabled = $true
}

function Stop-Jobs {
    foreach ($j in (@($script:Jobs) + @($script:PJobs))) {
        try { if (-not $j.Done) { [void]$j.PS.BeginStop($null, $null) } } catch { }
    }
}

function Stop-Search {
    $timer.Stop(); $ptimer.Stop()
    Stop-Jobs
    $script:Jobs = @(); $script:PJobs = @()
    Write-Status 'Search cancelled.'
    $script:SumProgress = 'CANCELLED - results may be incomplete'
    Update-SummaryLabel
    Reset-SearchUi
}

function Complete-Points {
    $script:PPool.Close(); $script:PPool.Dispose()

    Write-Status 'Summary by user:'
    foreach ($t in ($script:Scope | Sort-Object Team, User)) {
        $row = $script:SummaryRows[$t.TechId]
        if ($null -eq $row -or $row['Reports'] -is [DBNull]) { continue }
        $cnt = [int]$row['Reports']
        if ($cnt -eq 0) { continue }

        $pts = 0.0; $na = 0
        foreach ($rid in $script:TechReports[$t.TechId]) {
            $k = "{0}|{1}" -f $t.User.ToLower(), $rid
            if ($script:PointsMap.ContainsKey($k)) { $pts += [double]$script:PointsMap[$k] } else { $na++ }
        }
        if ($na -ge $cnt) {
            $row['Points'] = [DBNull]::Value; $row['State'] = ''
        } else {
            $row['Points'] = $pts; $row['State'] = ''
        }
        Write-Status ("  {0,-24} {1,3} report(s)   {2} point(s)" -f $t.User, $cnt, (Format-Points $pts))
    }

    $total = 0.0
    foreach ($v in $script:PointsMap.Values) { $total += [double]$v }
    $script:SumProgress = ''
    $script:SumTotal    = "Total points: $(Format-Points $total)"
    Update-Titles
    Update-SummaryLabel
    $gridS.Refresh(); $gridD.Refresh()
    Write-Status ("Done in {0:N1}s total." -f $script:Stopwatch.Elapsed.TotalSeconds)
    Reset-SearchUi
}

function Complete-Search {
    if ($script:Pool) { try { $script:Pool.Close(); $script:Pool.Dispose() } catch { } }

    # one row per tech+report (a report used on several PCs is merged -> points are never double counted)
    $merged = @{}
    foreach ($r in $script:All) {
        $k = "{0}|{1}" -f $r.Tech, $r.ReportId
        if (-not $merged.ContainsKey($k)) {
            $merged[$k] = [pscustomobject]@{
                Tech = $r.Tech; ReportId = $r.ReportId
                Machines = (New-Object System.Collections.Generic.List[string])
                Files    = (New-Object System.Collections.Generic.List[string])
                First = $r.First; Last = $r.Last
            }
        }
        $m = $merged[$k]
        if (-not $m.Machines.Contains($r.Machine)) { $m.Machines.Add($r.Machine) }
        if (-not $m.Files.Contains($r.File))       { $m.Files.Add($r.File) }
        if ([string]::CompareOrdinal($r.First, $m.First) -lt 0) { $m.First = $r.First }
        if ([string]::CompareOrdinal($r.Last,  $m.Last)  -gt 0) { $m.Last  = $r.Last }
    }
    $rows = @($merged.Values | Sort-Object @{ Expression = { $script:TechUser[$_.Tech] } }, First, ReportId)
    $multiDay = $script:Dates.Count -gt 1

    $script:Populating = $true
    $gridS.DataSource = $null; $gridD.DataSource = $null
    $dtD.DefaultView.RowFilter = ''
    $dtS.Clear(); $dtD.Clear()

    $script:DetailRows  = @{}
    $script:SummaryRows = @{}
    $script:ReportUsers = @{}
    $script:TechReports = @{}

    foreach ($m in $rows) {
        $user = $script:TechUser[$m.Tech]
        $row = $dtD.NewRow()
        $row['User']       = $user
        $row['Report ID']  = $m.ReportId
        $row['Points']     = [DBNull]::Value
        $row['State']      = '...'
        $row['Machine']    = ($m.Machines -join ', ')
        $row['File']       = ($m.Files -join ' | ')
        $row['First Seen'] = Format-Stamp $m.First $multiDay
        $row['Last Seen']  = Format-Stamp $m.Last  $multiDay
        $dtD.Rows.Add($row)
        $script:DetailRows[("{0}|{1}" -f $user.ToLower(), $m.ReportId)] = $row

        if (-not $script:ReportUsers.ContainsKey($m.ReportId)) { $script:ReportUsers[$m.ReportId] = New-Object System.Collections.Generic.List[string] }
        if (-not $script:ReportUsers[$m.ReportId].Contains($user)) { $script:ReportUsers[$m.ReportId].Add($user) }
        if (-not $script:TechReports.ContainsKey($m.Tech)) { $script:TechReports[$m.Tech] = New-Object System.Collections.Generic.List[string] }
        $script:TechReports[$m.Tech].Add($m.ReportId)
    }

    # summary: every tech in scope (so "no reports" and "could not read logs" are distinguishable)
    $withReports = 0
    foreach ($t in ($script:Scope | Sort-Object Team, User)) {
        $cnt = if ($script:TechReports.ContainsKey($t.TechId)) { $script:TechReports[$t.TechId].Count } else { 0 }
        $okMachines = @($t.MachineList | Where-Object { -not $script:Failed.Contains($_) }).Count

        $row = $dtS.NewRow()
        $row['Team'] = $t.Team; $row['TechId'] = $t.TechId; $row['User'] = $t.User
        $row['Flag'] = ''
        if ($t.MachineList.Count -eq 0 -or $okMachines -eq 0) {
            $row['Reports'] = [DBNull]::Value; $row['Points'] = [DBNull]::Value; $row['State'] = ''
        }
        elseif ($cnt -eq 0) {
            $row['Reports'] = 0; $row['Points'] = 0.0; $row['State'] = ''
        }
        else {
            $withReports++
            $row['Reports'] = $cnt; $row['Points'] = [DBNull]::Value; $row['State'] = '...'
        }
        $dtS.Rows.Add($row)
        $script:SummaryRows[$t.TechId] = $row
    }

    $gridS.DataSource = $dtS
    $gridD.DataSource = $dtD
    $script:Populating = $false
    $gridS.ClearSelection()
    Update-Titles

    $ids = @($script:ReportUsers.Keys)
    $script:SumBase = "{0}: {1} tech(s) with reports, {2} unique report(s), {3} machine(s) scanned" -f $script:DateLabel, $withReports, $ids.Count, $script:Jobs.Count
    Write-Status ("Logs read in {0:N1}s." -f $script:Stopwatch.Elapsed.TotalSeconds)

    if ($ids.Count -eq 0) {
        $script:SumTotal = 'Total points: 0'
        Update-SummaryLabel
        Reset-SearchUi
        return
    }

    $script:PtsDone = 0; $script:PtsTotal = $ids.Count
    $script:SumProgress = "Fetching points 0/$($ids.Count) ..."
    Update-SummaryLabel
    Write-Status "Fetching points for $($ids.Count) report(s) ..."

    $script:PointsMap = @{}
    $script:PPool = [runspacefactory]::CreateRunspacePool(1, [Math]::Min($MaxParallelApi, $ids.Count))
    $script:PPool.Open()
    $script:PJobs = @()
    foreach ($id in $ids) {
        $ps = [powershell]::Create()
        $ps.RunspacePool = $script:PPool
        [void]$ps.AddScript($PointsScript).AddArgument($id).AddArgument([string[]]$script:ReportUsers[$id].ToArray()).AddArgument($ApiBase)
        $script:PJobs += [pscustomobject]@{ ReportId = $id; PS = $ps; Handle = $ps.BeginInvoke(); Done = $false }
    }
    $ptimer.Start()
}

$timer.Add_Tick({
    $timer.Stop()
    $pending = 0
    foreach ($j in $script:Jobs) {
        if ($j.Done) { continue }
        if (-not $j.Handle.IsCompleted) { $pending++; continue }

        $j.Done = $true
        $out = $null
        try { $out = @($j.PS.EndInvoke($j.Handle))[-1] } catch { }
        if ($null -eq $out) {
            $msg = if ($j.PS.Streams.Error.Count) { $j.PS.Streams.Error[0].ToString() } else { 'No result returned' }
            [void]$script:Failed.Add($j.Machine)
        }
        elseif ($out.Error) {
            [void]$script:Failed.Add($out.Machine)
        }
        else {
            Write-Status ("  {0}: {1} report(s) found in {2:N1}s" -f $out.Machine, @($out.Results).Count, $out.Seconds)
            foreach ($r in $out.Results) { $script:All.Add($r) }
            if ($out.FirstLogDate -and [string]::CompareOrdinal($out.FirstLogDate, $script:Dates[0]) -gt 0) {
                [void]$script:Partial.Add($out.Machine)
            }
        }
        $j.PS.Dispose()
    }
    $script:SumProgress = "Reading logs: {0}/{1} machine(s) done" -f ($script:Jobs.Count - $pending), $script:Jobs.Count
    Update-SummaryLabel
    if ($pending -eq 0) { Complete-Search } else { $timer.Start() }
})

$ptimer.Add_Tick({
    $ptimer.Stop()
    $pending = 0
    foreach ($j in $script:PJobs) {
        if ($j.Done) { continue }
        if (-not $j.Handle.IsCompleted) { $pending++; continue }

        $j.Done = $true
        $script:PtsDone++
        $out = $null
        try { $out = @($j.PS.EndInvoke($j.Handle))[-1] } catch { }
        $j.PS.Dispose()

        $results = @()
        if ($out -and $out.Results) { $results = @($out.Results) }
        else {
            $msg = 'No result returned'
            Write-Status "  Report $($j.ReportId): points ERROR - $msg"
            $results = @($script:ReportUsers[$j.ReportId] | ForEach-Object { [pscustomobject]@{ User = $_; Points = $null; Error = $msg } })
        }

        foreach ($res in $results) {
            $row = $script:DetailRows[("{0}|{1}" -f $res.User.ToLower(), $j.ReportId)]
            if ($res.Error) {
                $script:NaCount++
                $why = if ($res.Error -match '\(404\)|not found on report') { 'No task measurement point records found.' } else { "points ERROR - $($res.Error)" }
                Write-Status "  Report $($j.ReportId) [$($res.User)]: $why"
                if ($row) { $row['State'] = '' }
            } else {
                $script:PointsMap[("{0}|{1}" -f $res.User.ToLower(), $j.ReportId)] = $res.Points
                if ($row) { $row['Points'] = $res.Points; $row['State'] = '' }
                Write-Status "  Report $($j.ReportId) [$($res.User)]: $(Format-Points $res.Points) point(s)"
            }
        }
    }
    $script:SumProgress = "Fetching points {0}/{1} ..." -f $script:PtsDone, $script:PtsTotal
    Update-SummaryLabel
    if ($pending -eq 0) { Complete-Points } else { $ptimer.Start() }
})

$btnRun.Add_Click({
    if ($script:Running) { Stop-Search; return }
    Save-Settings

    $sel = @(Get-ScopeTechs)
    if ($sel.Count -eq 0) { [void][System.Windows.Forms.MessageBox]::Show('No techs match the Team / User selection.', 'Nothing to search'); return }
    $machines = @($sel | ForEach-Object { $_.MachineList } | Select-Object -Unique)
    if ($machines.Count -eq 0) { [void][System.Windows.Forms.MessageBox]::Show('The selected tech(s) have no machines in the roster.', 'Nothing to search'); return }

    $from = $dtpFrom.Value.Date
    $to   = $dtpTo.Value.Date
    if ($to -lt $from) { $tmp = $from; $from = $to; $to = $tmp }
    $days = ($to - $from).Days + 1
    if ($days -gt $MaxRangeDays) { [void][System.Windows.Forms.MessageBox]::Show("Date range is limited to $MaxRangeDays days.", 'Date range'); return }

    $script:Running = $true
    $btnRun.Text = 'CANCEL SEARCH'
    Style-Button $btnRun 'danger'

    $script:Populating = $true
    $gridS.DataSource = $null; $gridD.DataSource = $null
    $dtD.DefaultView.RowFilter = ''
    $dtS.Clear(); $dtD.Clear()
    $gridS.DataSource = $dtS; $gridD.DataSource = $dtD
    $script:Populating = $false
    Update-Titles
    $txtLog.Clear()

    $script:All       = New-Object System.Collections.Generic.List[object]
    $script:Scope     = $sel
    $script:Dates     = @(0..($days - 1) | ForEach-Object { $from.AddDays($_).ToString('yyyy-MM-dd', $inv) })
    $script:DateLabel = if ($days -eq 1) { $script:Dates[0] } else { "{0} .. {1}" -f $script:Dates[0], $script:Dates[-1] }
    $script:Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $script:TechUser  = @{}
    foreach ($t in $sel) { $script:TechUser[$t.TechId] = $t.User }
    $techIds = [string[]]@($sel | ForEach-Object { $_.TechId })
    $script:Failed.Clear(); $script:Partial.Clear(); $script:NaCount = 0
    $script:SumBase = ''; $script:SumProgress = 'Starting ...'; $script:SumTotal = ''
    Update-SummaryLabel

    $script:Pool = [runspacefactory]::CreateRunspacePool(1, [Math]::Min($MaxParallelMachines, $machines.Count))
    $script:Pool.Open()
    $script:Jobs = @()

    Write-Status "Searching $($techIds.Count) tech(s) on $($machines.Count) machine(s) ($($script:DateLabel)) ..."
    foreach ($m in $machines) {
        $ps = [powershell]::Create()
        $ps.RunspacePool = $script:Pool
        [void]$ps.AddScript($SearchScript).AddArgument($m).AddArgument($techIds).AddArgument([string[]]$script:Dates).AddArgument($LogRelativePath)
        $script:Jobs += [pscustomobject]@{ Machine = $m; PS = $ps; Handle = $ps.BeginInvoke(); Done = $false }
    }
    $timer.Start()
})

# ---- export ----------------------------------------------------------------------------------------
function Get-RangeText { if ($dtpFrom.Value.Date -eq $dtpTo.Value.Date) { $dtpFrom.Value.ToString('yyyyMMdd', $inv) } else { "{0}-{1}" -f $dtpFrom.Value.ToString('yyyyMMdd', $inv), $dtpTo.Value.ToString('yyyyMMdd', $inv) } }
function Get-Num($v) { if ($v -is [DBNull]) { '' } else { $v } }

$btnCsvS.Add_Click({
    if ($dtS.Rows.Count -eq 0) { return }
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter   = 'CSV file (*.csv)|*.csv'
    $dlg.FileName = "TwisterSummary_{0}.csv" -f (Get-RangeText)
    if ($dlg.ShowDialog() -eq 'OK') {
        $dtS.Rows | ForEach-Object {
            [pscustomobject]@{
                DateFrom = $script:Dates[0]; DateTo = $script:Dates[-1]
                Team = $_['Team']; TechId = $_['TechId']; User = $_['User']
                Reports = (Get-Num $_['Reports']); Points = (Get-Num $_['Points'])
            }
        } | Export-Csv -Path $dlg.FileName -NoTypeInformation
        Write-Status "Exported summary to $($dlg.FileName)"
    }
})

$btnCsvD.Add_Click({
    if ($dtD.DefaultView.Count -eq 0) { return }
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter   = 'CSV file (*.csv)|*.csv'
    $dlg.FileName = "TwisterDetails_{0}.csv" -f (Get-RangeText)
    if ($dlg.ShowDialog() -eq 'OK') {
        $dtD.DefaultView | ForEach-Object {
            $r = $_.Row
            [pscustomobject]@{
                DateFrom = $script:Dates[0]; DateTo = $script:Dates[-1]
                User = $r['User']; ReportId = $r['Report ID']
                Points = (Get-Num $r['Points'])
                Machine = $r['Machine']; File = $r['File']; FirstSeen = $r['First Seen']; LastSeen = $r['Last Seen']
            }
        } | Export-Csv -Path $dlg.FileName -NoTypeInformation
        Write-Status "Exported details to $($dlg.FileName)"
    }
})

$form.Add_FormClosing({
    Save-Settings
    $timer.Stop(); $ptimer.Stop()
    Stop-Jobs
})

# ---- go -------------------------------------------------------------------------------------------------
$gridS.DataSource = $dtS
$gridD.DataSource = $dtD
Update-TeamList
Update-UserList
# restore the last Team / User (first run: All teams / All users)
$pref = Get-Settings
if ($pref.Team -and $cboTeam.Items.Contains([string]$pref.Team)) {
    $script:Loading = $true; $cboTeam.SelectedItem = [string]$pref.Team; $script:Loading = $false
    Update-UserList
}
if ($pref.TechId) {
    for ($k = 0; $k -lt $script:UserItems.Count; $k++) {
        if ($script:UserItems[$k].TechId -eq [string]$pref.TechId) { $cboUser.SelectedIndex = $k + 1; break }
    }
}
Update-Titles
Write-Status ("Roster: {0} active tech(s) from {1}" -f $script:Techs.Count, $RosterPath)
foreach ($w in $RosterWarnings) { Write-Status "  roster: $w" }
[void]$form.ShowDialog()
