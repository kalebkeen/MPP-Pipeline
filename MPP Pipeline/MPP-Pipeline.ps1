# =============================================================================
# MPP-Pipeline.ps1
# Unified GUI tool: Filter → Preview → Export MPP files to XML
#
# USAGE:
#   PowerShell -ExecutionPolicy Bypass -File "MPP-Pipeline.ps1"
#
# REQUIREMENTS:
#   - Windows PowerShell 5.1 or PowerShell 7+
#   - Microsoft Project installed (for COM-based XML export)
# =============================================================================

Set-StrictMode -Version Latest
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Collections

[System.Windows.Forms.Application]::EnableVisualStyles()

# =============================================================================
# GLOBAL STATE
# =============================================================================
$script:FilterRules      = [System.Collections.Generic.List[hashtable]]::new()
$script:SelectedFiles    = [System.Collections.Generic.List[hashtable]]::new()
$script:ExportCancelled  = $false
$script:SavedRuleSets    = [System.Collections.Generic.Dictionary[string, object]]::new()

# =============================================================================
# HELPER: Date parsing (from Sort-MPP.ps1)
# =============================================================================
function Get-SortDate {
    param([string]$BaseName, $File)
    $sep2 = '[-./\\]'
    $pat2 = "([0-9]{1,4})$sep2([0-9]{1,2})$sep2([0-9]{1,4})"
    if ($BaseName -match $pat2) {
        $a = [int]$Matches[1]; $b = [int]$Matches[2]; $c = [int]$Matches[3]
        if ($Matches[1].Length -eq 4) {
            $year = $a; $month = $b; $day = $c
        } else {
            $month = $a; $day = $b; $year = $c
            if ($year -lt 100) { $year += 2000 }
        }
        try {
            return (Get-Date -Year $year -Month $month -Day $day -Hour 0 -Minute 0 -Second 0)
        } catch {
            return $File.LastWriteTime
        }
    }
    return $File.LastWriteTime
}

# =============================================================================
# HELPER: Score a file against the rule list
# Returns [int] best priority rank that matched, or 9999 if nothing matched
# AND-group logic: within an AND-group, ALL keywords must match
# =============================================================================
function Get-FileScore {
    param(
        [string]$FullPath,
        [string]$FileName,
        [System.Collections.Generic.List[hashtable]]$Rules
    )

    if ($Rules.Count -eq 0) { return 1 }   # no rules = everything rank 1

    $matchScope = { param($rule, $path, $name)
        $target = switch ($rule.Scope) {
            'Filename'  { $name }
            'Full Path' { $path }
            default     { "$path $name" }
        }
        return ($target -imatch [regex]::Escape($rule.Keyword))
    }

    # Group rules by GroupID (AND groups share the same GroupID)
    $groups = @{}
    foreach ($r in $Rules) {
        $gid = $r.GroupID
        if (-not $groups.ContainsKey($gid)) { $groups[$gid] = @() }
        $groups[$gid] += $r
    }

    $bestRank = 9999
    foreach ($gid in $groups.Keys) {
        $groupRules = $groups[$gid]
        $logic      = $groupRules[0].Logic   # all rules in a group share same Logic

        if ($logic -eq 'AND') {
            # All rules in this group must match
            $allMatch = $true
            foreach ($r in $groupRules) {
                if (-not (& $matchScope $r $FullPath $FileName)) { $allMatch = $false; break }
            }
            if ($allMatch) {
                $minRank = ($groupRules | Measure-Object -Property Priority -Minimum).Minimum
                if ($minRank -lt $bestRank) { $bestRank = $minRank }
            }
        } else {
            # OR: any rule in this group matching is enough
            foreach ($r in $groupRules) {
                if (& $matchScope $r $FullPath $FileName) {
                    if ($r.Priority -lt $bestRank) { $bestRank = $r.Priority }
                }
            }
        }
    }
    return $bestRank
}

# =============================================================================
# HELPER: Run the scan/sort and populate $script:SelectedFiles
# =============================================================================
function Invoke-Scan {
    param(
        [string]$SourcePath,
        [int]$MaxPerFolder,
        [System.Collections.Generic.List[hashtable]]$Rules,
        [string]$ExcludeFoldersCsv
    )

    $script:SelectedFiles.Clear()

    # Defensive recast - Rules may arrive unwrapped from pipeline
    if ($null -eq $Rules) {
        $Rules = [System.Collections.Generic.List[hashtable]]::new()
    } elseif ($Rules -isnot [System.Collections.Generic.List[hashtable]]) {
        $tmp = [System.Collections.Generic.List[hashtable]]::new()
        foreach ($r in @($Rules)) { $tmp.Add($r) }
        $Rules = $tmp
    }

    $excludeList = @()
    if ($ExcludeFoldersCsv.Trim() -ne '') {
        $excludeList = $ExcludeFoldersCsv -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }
    }

    # Manual paced walk instead of one big Get-ChildItem -Recurse burst:
    #  - excluded folders are pruned BEFORE descending, so the share never
    #    traverses them at all
    #  - a short pause between directories keeps the metadata requests from
    #    monopolizing the shared drive
    $allFiles    = [System.Collections.Generic.List[System.IO.FileInfo]]::new()
    $dirQueue    = [System.Collections.Generic.Queue[string]]::new()
    $dirQueue.Enqueue($SourcePath)
    $dirsScanned = 0
    while ($dirQueue.Count -gt 0) {
        $dir = $dirQueue.Dequeue()

        $skipDir = $false
        foreach ($ex in $excludeList) {
            if ($dir -imatch [regex]::Escape($ex)) { $skipDir = $true; break }
        }
        if ($skipDir) { continue }

        try {
            foreach ($fp in [System.IO.Directory]::EnumerateFiles($dir, '*.mpp')) {
                $allFiles.Add([System.IO.FileInfo]::new($fp))
            }
            foreach ($dp in [System.IO.Directory]::EnumerateDirectories($dir)) {
                $dirQueue.Enqueue($dp)
            }
        } catch { }   # unreadable folder (permissions, path length) — skip it

        $dirsScanned++
        if (($dirsScanned % 5) -eq 0) {
            $lblPreviewSummary.Text = "Scanning… ($dirsScanned folders, $($allFiles.Count) MPP files)"
            [System.Windows.Forms.Application]::DoEvents()
        }
        Start-Sleep -Milliseconds 150
    }

    # Apply folder exclusions (second pass catches filename-level matches;
    # whole excluded folders were already pruned during the walk above)
    $workingFiles = foreach ($f in $allFiles) {
        $excluded = $false
        foreach ($ex in $excludeList) {
            if ($f.FullName -imatch [regex]::Escape($ex)) { $excluded = $true; break }
        }
        if (-not $excluded) { $f }
    }

    if ($null -eq $workingFiles) { return }

    # Group by immediate parent folder
    $groups = @{}
    foreach ($file in $workingFiles) {
        $key = $file.Directory.FullName
        if (-not $groups.ContainsKey($key)) { $groups[$key] = [System.Collections.Generic.List[object]]::new() }
        $groups[$key].Add($file)
    }

    foreach ($folderKey in ($groups.Keys | Sort-Object)) {
        $filesInFolder = $groups[$folderKey]

        # Score each file
        $records = foreach ($file in $filesInFolder) {
            $base  = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
            $score = Get-FileScore -FullPath $file.FullName -FileName $file.Name -Rules $Rules
            [PSCustomObject]@{
                File      = $file
                Base      = $base
                Score     = $score
                SortDate  = Get-SortDate $base $file
            }
        }

        # Filter out unmatched (score 9999) only when rules exist
        if ($Rules.Count -gt 0) {
            $matched = $records | Where-Object { $_.Score -lt 9999 }
            $matchedArr = @($matched)
            if ($null -eq $matched -or $matchedArr.Count -eq 0) { continue }
            $matched = $matchedArr
            $records = $matched
        }

        # Pick top N per folder: best score first, then newest date
        $winners = $records |
            Sort-Object Score, @{Expression='SortDate'; Descending=$true} |
            Select-Object -First $MaxPerFolder

        foreach ($w in @($winners)) {
            $script:SelectedFiles.Add(@{
                Include   = $true
                FilePath  = $w.File.FullName
                FileName  = $w.File.Name
                Folder    = $folderKey
                Score     = $w.Score
                SortDate    = $w.SortDate.ToString('yyyy-MM-dd')
                StatusDate  = ''
                DateSource  = ''
            })
        }
    }
}

# =============================================================================
# HELPER: Read Status Date from MPP via COM
# =============================================================================
function Get-MppStatusDate {
    param([string]$FilePath, $MspApp)
    try {
        $MspApp.DisplayAlerts = $false
        $MspApp.FileOpen($FilePath, $true) | Out-Null
        Start-Sleep -Milliseconds 800
        $sd = $MspApp.ActiveProject.StatusDate
        $MspApp.FileClose(0) | Out-Null
        if ($sd -and $sd -ne '1/1/1984') {
            return ([datetime]$sd).ToString('yyyy-MM-dd')
        }
        return ''
    } catch {
        try { $MspApp.FileClose(0) | Out-Null } catch {}
        return ''
    }
}

# =============================================================================
# HELPER: Inject StatusDate attribute into XML file
# =============================================================================
function Set-XmlStatusDate {
    param([string]$XmlPath, [string]$StatusDate)
    try {
        if (-not (Test-Path $XmlPath)) { return }
        [xml]$doc = Get-Content -LiteralPath $XmlPath -Encoding UTF8
        $ns  = 'http://schemas.microsoft.com/project'
        $nsm = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
        $nsm.AddNamespace('p', $ns)

        $node = $doc.SelectSingleNode('//p:StatusDate', $nsm)
        if ($null -ne $node) {
            $node.InnerText = "${StatusDate}T00:00:00"
        } else {
            $proj    = $doc.SelectSingleNode('//p:Project', $nsm)
            if ($null -ne $proj) {
                $newNode = $doc.CreateElement('StatusDate', $ns)
                $newNode.InnerText = "${StatusDate}T00:00:00"
                $proj.AppendChild($newNode) | Out-Null
            }
        }
        $doc.Save($XmlPath)
    } catch {}
}

# =============================================================================
# HELPER: Local staging — MS Project must never read/write over the network.
# Staged copies are keyed by source path + size + timestamp, so a file staged
# during Preview (status-date read) is re-used during Export without a second
# network copy.
# =============================================================================
$script:StagingDir = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'MPP Pipeline\staging'

function Clear-StagingDir {
    try {
        if (Test-Path -LiteralPath $script:StagingDir) {
            Remove-Item -LiteralPath $script:StagingDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    } catch { }
}

function Get-StagedCopy {
    param([string]$SourcePath)

    if (-not (Test-Path -LiteralPath $script:StagingDir)) {
        New-Item -ItemType Directory -Path $script:StagingDir -Force | Out-Null
    }

    $fi  = [System.IO.FileInfo]::new($SourcePath)
    $sig = '{0}|{1}|{2}' -f $SourcePath.ToLowerInvariant(), $fi.LastWriteTimeUtc.Ticks, $fi.Length
    $md5 = [System.Security.Cryptography.MD5]::Create()
    try {
        $hash = ([System.BitConverter]::ToString($md5.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($sig))) -replace '-','').Substring(0,12)
    } finally { $md5.Dispose() }

    $staged = Join-Path $script:StagingDir ('{0}_{1}.mpp' -f [System.IO.Path]::GetFileNameWithoutExtension($SourcePath), $hash)
    if (Test-Path -LiteralPath $staged) {
        return @{ Path = $staged; Copied = $false }
    }
    Copy-Item -LiteralPath $SourcePath -Destination $staged -Force
    return @{ Path = $staged; Copied = $true }
}

function Invoke-PacedPause {
    # Sleeps N seconds while keeping the UI responsive; the point is to leave
    # the shared drive idle between network operations so other users' requests
    # get through.
    param([int]$Seconds, [bool]$RespectCancel = $true)
    if ($Seconds -le 0) { return }
    $until = (Get-Date).AddSeconds($Seconds)
    while ((Get-Date) -lt $until) {
        if ($RespectCancel -and $script:ExportCancelled) { return }
        Start-Sleep -Milliseconds 200
        [System.Windows.Forms.Application]::DoEvents()
    }
}

# =============================================================================
# HELPER: Full shell folder picker (shows network drives + mapped drives)
# =============================================================================
function Show-FolderPicker {
    param(
        [string]$Title      = 'Select Folder',
        [string]$StartPath  = ''
    )
    # Flags: 0x0001 = only folders, 0x0040 = show edit box, 0x0010 = include network
    # Using Shell.Application BrowseForFolder roots at Desktop showing everything
    try {
        $shell  = New-Object -ComObject Shell.Application
        # BrowseForFolder(hwnd, title, flags, rootFolder)
        # rootFolder 0 = Desktop (shows all including network)
        $folder = $shell.BrowseForFolder(0, $Title, 0x0041, 0)
        if ($null -ne $folder) {
            # Self() gives the FolderItem; Path gives the real path
            $item = $folder.Self()
            return $item.Path
        }
    } catch {
        # Fallback to standard FolderBrowserDialog if COM fails
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description  = $Title
        if ($StartPath -ne '') { $dlg.SelectedPath = $StartPath }
        if ($dlg.ShowDialog() -eq 'OK') { return $dlg.SelectedPath }
    }
    return $null
}

# =============================================================================
# BUILD THE MAIN FORM
# =============================================================================
$form = New-Object System.Windows.Forms.Form
$form.Text            = 'MPP Pipeline  —  Filter · Preview · Export'
$form.Size            = New-Object System.Drawing.Size(1020, 800)
$form.StartPosition   = 'CenterScreen'
$form.MinimumSize     = New-Object System.Drawing.Size(900, 700)
$form.Font            = New-Object System.Drawing.Font('Segoe UI', 9)
$form.BackColor       = [System.Drawing.Color]::FromArgb(245, 245, 248)

# ── Title bar ──────────────────────────────────────────────────────────────
# ── Tab control ────────────────────────────────────────
$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Font    = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$tabs.Padding = New-Object System.Drawing.Point(14, 4)
# Add to form FIRST, then set Dock - this is the correct order for PS2EXE
$form.Controls.Add($tabs)
$tabs.Dock = [System.Windows.Forms.DockStyle]::Fill


# helper: standard GroupBox
function New-GroupBox {
    param($Text, $Left, $Top, $Width, $Height)
    $gb = New-Object System.Windows.Forms.GroupBox
    $gb.Text     = $Text
    $gb.Left     = $Left; $gb.Top = $Top
    $gb.Width    = $Width; $gb.Height = $Height
    $gb.Font     = New-Object System.Drawing.Font('Segoe UI', 9)
    return $gb
}

# helper: standard Label
function New-Label {
    param($Text, $Left, $Top, $Width = 120, $Height = 20)
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $Text; $l.Left = $Left; $l.Top = $Top
    $l.Width = $Width; $l.Height = $Height
    $l.TextAlign = 'MiddleLeft'
    return $l
}

# helper: standard TextBox
function New-Textbox {
    param($Left, $Top, $Width, $Text = '')
    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Left = $Left; $tb.Top = $Top; $tb.Width = $Width
    $tb.Text = $Text
    return $tb
}

# helper: standard Button
function New-Btn {
    param($Text, $Left, $Top, $Width = 90, $Height = 28, $Color = $null)
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $Text; $b.Left = $Left; $b.Top = $Top
    $b.Width = $Width; $b.Height = $Height
    $b.FlatStyle = 'Flat'
    if ($null -ne $Color) {
        $b.BackColor = $Color
        $b.ForeColor = [System.Drawing.Color]::White
    }
    $b.FlatAppearance.BorderSize = 1
    return $b
}

# =============================================================================
# TAB 1 — SETUP
# =============================================================================
$tabSetup        = New-Object System.Windows.Forms.TabPage
$tabSetup.Text   = '  1 · Setup  '
$tabSetup.BackColor = [System.Drawing.Color]::FromArgb(245,245,248)
$tabs.TabPages.Add($tabSetup)

# ── Source folder ─────────────────────────────────────────────────────────
$gbSource = New-GroupBox 'Source Folder  (root folder with .MPP files)' 12 16 980 72
$tabSetup.Controls.Add($gbSource)

$tbSource = New-Textbox 12 28 828 ''
$gbSource.Controls.Add($tbSource)

$btnBrowseSource = New-Btn 'Browse…' 850 26 110 28
$gbSource.Controls.Add($btnBrowseSource)
$btnBrowseSource.Add_Click({
    $picked = Show-FolderPicker -Title 'Select the root folder containing your MPP files' -StartPath $tbSource.Text
    if ($null -ne $picked -and $picked -ne '') { $tbSource.Text = $picked }
})

# ── Output folder ─────────────────────────────────────────────────────────
$gbDest = New-GroupBox 'Output Folder  (XML files written here)' 12 98 980 72
$tabSetup.Controls.Add($gbDest)

$tbDest = New-Textbox 12 28 828 ''
$gbDest.Controls.Add($tbDest)

$btnBrowseDest = New-Btn 'Browse…' 850 26 110 28
$gbDest.Controls.Add($btnBrowseDest)
$btnBrowseDest.Add_Click({
    $picked = Show-FolderPicker -Title 'Select the output folder for XML files' -StartPath $tbDest.Text
    if ($null -ne $picked -and $picked -ne '') { $tbDest.Text = $picked }
})

# ── Output folder name ────────────────────────────────────────────────────
$gbFolderName = New-GroupBox 'Output Folder Name  (a dated subfolder will be created inside the Output Folder)' 12 180 980 80
$tabSetup.Controls.Add($gbFolderName)

$lblFolderName = New-Label 'Folder name prefix:' 12 30 140
$gbFolderName.Controls.Add($lblFolderName)

$tbFolderName = New-Textbox 158 28 300 'PipelineOutput'
$gbFolderName.Controls.Add($tbFolderName)

$lblFolderNamePreview = New-Object System.Windows.Forms.Label
$lblFolderNamePreview.Left = 468; $lblFolderNamePreview.Top = 30
$lblFolderNamePreview.Width = 490; $lblFolderNamePreview.Height = 20
$lblFolderNamePreview.ForeColor = [System.Drawing.Color]::FromArgb(80,80,100)
$lblFolderNamePreview.Font = New-Object System.Drawing.Font('Segoe UI', 8.5)
$lblFolderNamePreview.Text = "Preview: PipelineOutput $(Get-Date -Format 'yyyy-MM-dd')"
$gbFolderName.Controls.Add($lblFolderNamePreview)

$tbFolderName.Add_TextChanged({
    $prefix = $tbFolderName.Text.Trim()
    if ($prefix -eq '') { $prefix = 'PipelineOutput' }
    $lblFolderNamePreview.Text = "Preview: $prefix $(Get-Date -Format 'yyyy-MM-dd')"
})

# ── Per-folder limit ──────────────────────────────────────────────────────
$gbLimit = New-GroupBox 'Per-Folder Selection Limit' 12 270 500 72
$tabSetup.Controls.Add($gbLimit)

$lblLimit = New-Label 'Max files per folder:' 12 30 150
$gbLimit.Controls.Add($lblLimit)

$nudLimit = New-Object System.Windows.Forms.NumericUpDown
$nudLimit.Left = 168; $nudLimit.Top = 28
$nudLimit.Width = 70; $nudLimit.Height = 24
$nudLimit.Minimum = 1; $nudLimit.Maximum = 50; $nudLimit.Value = 1
$gbLimit.Controls.Add($nudLimit)

$lblLimitNote = New-Label '(1 = pick the single best file per folder)' 248 30 280
$lblLimitNote.ForeColor = [System.Drawing.Color]::Gray
$gbLimit.Controls.Add($lblLimitNote)

# ── Excluded folder keywords ───────────────────────────────────────────────
$gbExclude = New-GroupBox 'Exclude Folder Keywords  (comma-separated path segments to skip)' 12 352 980 72
$tabSetup.Controls.Add($gbExclude)

$tbExclude = New-Textbox 12 28 950 'Comparison'
$gbExclude.Controls.Add($tbExclude)

# ── Next button ───────────────────────────────────────────────────────────
$btnToFilter = New-Btn '  Next: Filter Rules  ▶' 720 450 212 36 ([System.Drawing.Color]::FromArgb(31,78,120))
$btnToFilter.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
$tabSetup.Controls.Add($btnToFilter)
$btnToFilter.Add_Click({
    if ($tbSource.Text.Trim() -eq '' -or -not (Test-Path $tbSource.Text.Trim())) {
        [System.Windows.Forms.MessageBox]::Show('Please enter a valid Source Folder path.','Validation',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    if ($tbDest.Text.Trim() -eq '') {
        [System.Windows.Forms.MessageBox]::Show('Please enter an Output Folder path.','Validation',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    $tabs.SelectedIndex = 1
})

# =============================================================================
# TAB 2 — FILTER RULES
# =============================================================================
$tabFilter       = New-Object System.Windows.Forms.TabPage
$tabFilter.Text  = '  2 · Filter Rules  '
$tabFilter.BackColor = [System.Drawing.Color]::FromArgb(245,245,248)
$tabs.TabPages.Add($tabFilter)

# ── Instructions label ────────────────────────────────────────────────────
$lblRuleInstr = New-Object System.Windows.Forms.Label
$lblRuleInstr.Text = "Add keyword rules below.  Priority = lower number wins.  " +
                     "AND-group rules ALL must match.  OR rules match independently.  " +
                     "Files with no match are excluded (unless rule list is empty = include all)."
$lblRuleInstr.Left = 12; $lblRuleInstr.Top = 12
$lblRuleInstr.Width = 980; $lblRuleInstr.Height = 56
$lblRuleInstr.AutoSize = $false
$lblRuleInstr.MaximumSize = New-Object System.Drawing.Size(980, 0)
$lblRuleInstr.ForeColor = [System.Drawing.Color]::FromArgb(80,80,100)
$tabFilter.Controls.Add($lblRuleInstr)

# ── Add-rule form ─────────────────────────────────────────────────────────
$gbAddRule = New-GroupBox 'Add New Rule' 12 52 970 110
$tabFilter.Controls.Add($gbAddRule)

# Row 1: Keyword | Priority | Logic | AND Group ID | Scope
$gbAddRule.Controls.Add((New-Label 'Keyword:' 10 28 62))
$tbRuleKeyword = New-Textbox 72 26 200
$gbAddRule.Controls.Add($tbRuleKeyword)

$gbAddRule.Controls.Add((New-Label 'Priority:' 282 28 56))
$nudRulePriority = New-Object System.Windows.Forms.NumericUpDown
$nudRulePriority.Left = 338; $nudRulePriority.Top = 26
$nudRulePriority.Width = 56; $nudRulePriority.Minimum = 1; $nudRulePriority.Maximum = 999; $nudRulePriority.Value = 10
$gbAddRule.Controls.Add($nudRulePriority)

$gbAddRule.Controls.Add((New-Label 'Logic:' 404 28 46))
$cboLogic = New-Object System.Windows.Forms.ComboBox
$cboLogic.Left = 450; $cboLogic.Top = 26; $cboLogic.Width = 66
$cboLogic.DropDownStyle = 'DropDownList'
$cboLogic.Items.AddRange(@('OR','AND')) | Out-Null
$cboLogic.SelectedIndex = 0
$gbAddRule.Controls.Add($cboLogic)

$gbAddRule.Controls.Add((New-Label 'AND Grp:' 526 28 62))
$nudGroupID = New-Object System.Windows.Forms.NumericUpDown
$nudGroupID.Left = 588; $nudGroupID.Top = 26
$nudGroupID.Width = 52; $nudGroupID.Minimum = 1; $nudGroupID.Maximum = 99; $nudGroupID.Value = 1
$gbAddRule.Controls.Add($nudGroupID)

$gbAddRule.Controls.Add((New-Label 'Scope:' 650 28 48))
$cboScope = New-Object System.Windows.Forms.ComboBox
$cboScope.Left = 698; $cboScope.Top = 26; $cboScope.Width = 100
$cboScope.DropDownStyle = 'DropDownList'
$cboScope.Items.AddRange(@('Filename','Full Path','Both')) | Out-Null
$cboScope.SelectedIndex = 0
$gbAddRule.Controls.Add($cboScope)

# Add Rule button - clearly to the right with no overlap
$btnAddRule = New-Btn '+ Add Rule' 812 22 140 30 ([System.Drawing.Color]::FromArgb(31,78,120))
$gbAddRule.Controls.Add($btnAddRule)

# Row 2: note label
$lblGroupNote = New-Label 'AND Group ID only matters when Logic = AND. Rules with the same Group ID must ALL match.' 10 62 940 20
$lblGroupNote.ForeColor = [System.Drawing.Color]::Gray
$gbAddRule.Controls.Add($lblGroupNote)

# ── Rules grid ────────────────────────────────────────────────────────────
$gbRules = New-GroupBox 'Current Rules  (select row + use buttons to edit)' 12 172 970 308
$tabFilter.Controls.Add($gbRules)

$dgvRules = New-Object System.Windows.Forms.DataGridView
$dgvRules.Left = 8; $dgvRules.Top = 22
$dgvRules.Width = 870; $dgvRules.Height = 276
$dgvRules.AllowUserToAddRows    = $false
$dgvRules.AllowUserToDeleteRows = $false
$dgvRules.MultiSelect           = $false
$dgvRules.SelectionMode         = 'FullRowSelect'
$dgvRules.RowHeadersVisible     = $false
$dgvRules.AutoSizeColumnsMode   = 'Fill'
$dgvRules.BackgroundColor       = [System.Drawing.Color]::White
$dgvRules.BorderStyle           = 'None'
$dgvRules.ReadOnly              = $true

foreach ($col in @(
    @{Name='Priority'; Header='Priority'; FillW=60},
    @{Name='Keyword';  Header='Keyword';  FillW=200},
    @{Name='Logic';    Header='Logic';    FillW=50},
    @{Name='GroupID';  Header='AND Grp';  FillW=55},
    @{Name='Scope';    Header='Scope';    FillW=80}
)) {
    $c = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $c.Name = $col.Name; $c.HeaderText = $col.Header
    $c.FillWeight = $col.FillW
    $dgvRules.Columns.Add($c) | Out-Null
}

$gbRules.Controls.Add($dgvRules)

# Side buttons for rules
$btnRuleUp     = New-Btn '▲ Up'      886 30  76 28
$btnRuleDown   = New-Btn '▼ Down'    886 64  76 28
$btnRuleDelete = New-Btn '✕ Delete'  886 98  76 28 ([System.Drawing.Color]::FromArgb(180,30,30))
$btnRuleClear  = New-Btn 'Clear All' 886 140 76 28
$gbRules.Controls.Add($btnRuleUp)
$gbRules.Controls.Add($btnRuleDown)
$gbRules.Controls.Add($btnRuleDelete)
$gbRules.Controls.Add($btnRuleClear)

function Refresh-RulesGrid {
    $dgvRules.Rows.Clear()
    foreach ($r in $script:FilterRules) {
        $dgvRules.Rows.Add($r.Priority, $r.Keyword, $r.Logic, $r.GroupID, $r.Scope) | Out-Null
    }
}

$btnAddRule.Add_Click({
    if ($tbRuleKeyword.Text.Trim() -eq '') {
        [System.Windows.Forms.MessageBox]::Show('Please enter a keyword.','Add Rule',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    $script:FilterRules.Add(@{
        Keyword  = $tbRuleKeyword.Text.Trim()
        Priority = [int]$nudRulePriority.Value
        Logic    = $cboLogic.SelectedItem.ToString()
        GroupID  = [int]$nudGroupID.Value
        Scope    = $cboScope.SelectedItem.ToString()
    })
    # Re-sort by priority
    $sorted = $script:FilterRules | Sort-Object { $_['Priority'] }
    $script:FilterRules.Clear()
    foreach ($s in $sorted) { $script:FilterRules.Add($s) }
    Refresh-RulesGrid
    $tbRuleKeyword.Text = ''
    $nudRulePriority.Value = [math]::Min(999, $nudRulePriority.Value + 10)
})

$btnRuleDelete.Add_Click({
    if ($dgvRules.SelectedRows.Count -eq 0) { return }
    $idx = $dgvRules.SelectedRows[0].Index
    $script:FilterRules.RemoveAt($idx)
    Refresh-RulesGrid
})

$btnRuleClear.Add_Click({
    if ([System.Windows.Forms.MessageBox]::Show('Remove all rules?','Confirm',[System.Windows.Forms.MessageBoxButtons]::YesNo) -eq 'Yes') {
        $script:FilterRules.Clear()
        Refresh-RulesGrid
    }
})

$btnRuleUp.Add_Click({
    if ($dgvRules.SelectedRows.Count -eq 0) { return }
    $idx = $dgvRules.SelectedRows[0].Index
    if ($idx -eq 0) { return }
    $tmp = $script:FilterRules[$idx - 1]
    $script:FilterRules[$idx - 1] = $script:FilterRules[$idx]
    $script:FilterRules[$idx] = $tmp
    Refresh-RulesGrid
    $dgvRules.Rows[$idx - 1].Selected = $true
})

$btnRuleDown.Add_Click({
    if ($dgvRules.SelectedRows.Count -eq 0) { return }
    $idx = $dgvRules.SelectedRows[0].Index
    if ($idx -ge $script:FilterRules.Count - 1) { return }
    $tmp = $script:FilterRules[$idx + 1]
    $script:FilterRules[$idx + 1] = $script:FilterRules[$idx]
    $script:FilterRules[$idx] = $tmp
    Refresh-RulesGrid
    $dgvRules.Rows[$idx + 1].Selected = $true
})

# ── Scan button ───────────────────────────────────────────────────────────
$btnScan = New-Btn '  Scan & Preview  ▶' 720 562 232 38 ([System.Drawing.Color]::FromArgb(31,78,120))
$btnScan.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
$tabFilter.Controls.Add($btnScan)

# ── Save / Load rule sets ─────────────────────────────────────────────────
# ── In-app rule set save/load panel ──────────────────────────────────────
$gbRuleSets = New-GroupBox 'Saved Rule Sets  (in-app, session only)' 12 490 970 64
$tabFilter.Controls.Add($gbRuleSets)

$lblRSName = New-Label 'Name:' 8 28 40
$gbRuleSets.Controls.Add($lblRSName)

$tbRSName = New-Textbox 50 26 180 ''
$gbRuleSets.Controls.Add($tbRSName)

$btnSaveRules = New-Btn 'Save Current' 240 24 110 28 ([System.Drawing.Color]::FromArgb(31,78,120))
$gbRuleSets.Controls.Add($btnSaveRules)

$cboRuleSets = New-Object System.Windows.Forms.ComboBox
$cboRuleSets.Left = 360; $cboRuleSets.Top = 26; $cboRuleSets.Width = 280
$cboRuleSets.DropDownStyle = 'DropDownList'
$gbRuleSets.Controls.Add($cboRuleSets)

$btnLoadRules = New-Btn 'Load Selected' 650 24 110 28 ([System.Drawing.Color]::FromArgb(70,100,140))
$gbRuleSets.Controls.Add($btnLoadRules)

$btnDeleteRS = New-Btn 'Delete' 770 24 80 28 ([System.Drawing.Color]::FromArgb(180,30,30))
$gbRuleSets.Controls.Add($btnDeleteRS)

$lblRuleHint = New-Label 'Tip: Empty rule list = include ALL .mpp files found.' 12 568 680 20
$lblRuleHint.ForeColor = [System.Drawing.Color]::FromArgb(100,100,140)
$tabFilter.Controls.Add($lblRuleHint)

function Refresh-RuleSetsDropdown {
    $cboRuleSets.Items.Clear()
    foreach ($k in ($script:SavedRuleSets.Keys | Sort-Object)) {
        $cboRuleSets.Items.Add($k) | Out-Null
    }
    if ($cboRuleSets.Items.Count -gt 0) { $cboRuleSets.SelectedIndex = 0 }
}

# =============================================================================
# TAB 3 — PREVIEW & EDIT
# =============================================================================
$tabPreview       = New-Object System.Windows.Forms.TabPage
$tabPreview.Text  = '  3 · Preview & Edit  '
$tabPreview.BackColor = [System.Drawing.Color]::FromArgb(245,245,248)
$tabs.TabPages.Add($tabPreview)

$lblPreviewSummary = New-Label '' 12 12 760 24
$lblPreviewSummary.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
$tabPreview.Controls.Add($lblPreviewSummary)

$btnRescan = New-Btn '↺ Re-Scan' 800 8 120 28 ([System.Drawing.Color]::FromArgb(100,100,140))
$tabPreview.Controls.Add($btnRescan)

# ── Preview grid ──────────────────────────────────────────────────────────
$dgvPreview = New-Object System.Windows.Forms.DataGridView
$dgvPreview.Left = 12; $dgvPreview.Top = 44
$dgvPreview.Width = 930; $dgvPreview.Height = 540
$dgvPreview.AllowUserToAddRows    = $false
$dgvPreview.AllowUserToDeleteRows = $false
$dgvPreview.MultiSelect           = $true
$dgvPreview.SelectionMode         = 'FullRowSelect'
$dgvPreview.RowHeadersVisible     = $false
$dgvPreview.AutoSizeColumnsMode   = 'Fill'
$dgvPreview.BackgroundColor       = [System.Drawing.Color]::White
$dgvPreview.BorderStyle           = 'None'
$dgvPreview.EditMode              = 'EditOnKeystrokeOrF2'

$colInclude = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$colInclude.Name = 'Include'; $colInclude.HeaderText = 'Export?'; $colInclude.FillWeight = 50
$dgvPreview.Columns.Add($colInclude) | Out-Null

foreach ($col in @(
    @{Name='FileName';  Header='File Name';    FillW=220},
    @{Name='SortDate';  Header='File Date';    FillW=80},
    @{Name='Score';     Header='Priority';     FillW=55},
    @{Name='Folder';    Header='Source Folder';FillW=400},
    @{Name='DateSource'; Header='Date Source'; FillW=80},
    @{Name='StatusDate'; Header='Status Date (editable)'; FillW=120}
)) {
    $c = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $c.Name = $col.Name; $c.HeaderText = $col.Header
    $c.FillWeight = $col.FillW
    # StatusDate is editable; DateSource shows where it came from (read-only)
    if ($col.Name -notin @('StatusDate')) { $c.ReadOnly = $true }
    $dgvPreview.Columns.Add($c) | Out-Null
}

$tabPreview.Controls.Add($dgvPreview)

# Select-all / deselect-all
$pnlPreviewBtns = New-Object System.Windows.Forms.Panel
$pnlPreviewBtns.Left = 12; $pnlPreviewBtns.Top = 592
$pnlPreviewBtns.Width = 930; $pnlPreviewBtns.Height = 32
$tabPreview.Controls.Add($pnlPreviewBtns)

$btnSelAll   = New-Btn '✔ Select All'   0 2 110 28
$btnSelNone  = New-Btn '✗ Deselect All' 118 2 110 28
$btnToExport = New-Btn '  Next: Export  ▶' 760 2 180 28 ([System.Drawing.Color]::FromArgb(31,78,120))
$btnToExport.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$pnlPreviewBtns.Controls.AddRange(@($btnSelAll, $btnSelNone, $btnToExport))

$btnSelAll.Add_Click({
    foreach ($row in $dgvPreview.Rows) { $row.Cells['Include'].Value = $true }
})
$btnSelNone.Add_Click({
    foreach ($row in $dgvPreview.Rows) { $row.Cells['Include'].Value = $false }
})
$btnToExport.Add_Click({ $tabs.SelectedIndex = 3 })

# =============================================================================
# HELPER: Pre-read status dates from all selected MPP files via COM
# =============================================================================
# HELPER: Parse a date from a filename using the same flexible parser as Sort
# Returns [datetime] or $null
# =============================================================================
function Get-FilenameDate {
    param([string]$BaseName)
    # Use [0-9] instead of \d to avoid PS2EXE regex escape issues
    # Separators: dash, dot, forward-slash, backslash
    $sep = '[-./\\]'
    $pat = "([0-9]{1,4})$sep([0-9]{1,2})$sep([0-9]{1,4})"
    if ($BaseName -match $pat) {
        $a = [int]$Matches[1]; $b = [int]$Matches[2]; $c = [int]$Matches[3]
        if ($Matches[1].Length -eq 4) {
            $year = $a; $month = $b; $day = $c
        } else {
            $month = $a; $day = $b; $year = $c
            if ($year -lt 100) { $year += 2000 }
        }
        # Sanity check: year 2000-2040, month 1-12, day 1-31
        if ($year -ge 2000 -and $year -le 2040 -and
            $month -ge 1   -and $month -le 12  -and
            $day   -ge 1   -and $day   -le 31) {
            try { return (Get-Date -Year $year -Month $month -Day $day -Hour 0 -Minute 0 -Second 0) }
            catch { }
        }
    }
    return $null
}

# =============================================================================
# HELPER: Pre-read status dates using 3-tier fallback
#   Tier 1: MS Project StatusDate field (Project Information)
#   Tier 2: Date parsed from filename (flexible parser, sanity-checked)
#   Tier 3: File Last Modified date (with warning flag)
# =============================================================================
function Read-StatusDates {
    if ($script:SelectedFiles.Count -eq 0) { return }

    $lblPreviewSummary.Text = 'Reading status dates from MPP files via MS Project…'
    [System.Windows.Forms.Application]::DoEvents()

    $msp = $null
    try {
        for ($i = 1; $i -le 10; $i++) {
            try {
                $msp = New-Object -ComObject 'MSProject.Application'
                Start-Sleep -Seconds 3
                $msp.Visible       = $false
                $msp.DisplayAlerts = $false
                break
            } catch {
                $msp = $null
                Start-Sleep -Seconds 2
            }
        }
        if ($null -eq $msp) {
            $lblPreviewSummary.Text = 'Could not launch MS Project to read status dates.'
            return
        }

        $total = $script:SelectedFiles.Count
        $done  = 0

        foreach ($entry in $script:SelectedFiles) {
            $done++
            $lblPreviewSummary.Text = "Reading status dates… ($done / $total)"
            [System.Windows.Forms.Application]::DoEvents()

            $resolvedDate   = $null
            $resolvedSource = ''

            # ── Tier 1: MS Project StatusDate field ───────────────────────────
            $tier1Copied = $false
            try {
                # Stage the file locally so MS Project never reads over the
                # network; the staged copy is re-used by the Export stage.
                $openPath = $entry.FilePath
                try {
                    $stage       = Get-StagedCopy -SourcePath $entry.FilePath
                    $openPath    = $stage.Path
                    $tier1Copied = $stage.Copied
                } catch { }   # staging failed (source locked?) — open from source

                $msp.DisplayAlerts = $false
                $msp.FileOpen($openPath, $true) | Out-Null
                Start-Sleep -Milliseconds 600
                $sdRaw = $msp.ActiveProject.StatusDate
                $msp.FileClose(0) | Out-Null

                if ($null -ne $sdRaw) {
                    $sdStr = "$sdRaw".Trim()
                    if ($sdStr -ne '' -and $sdStr -ne 'NA') {
                        try {
                            $sdDt = [datetime]$sdRaw
                            if ($sdDt.Year -ge 2000 -and $sdDt.Year -le 2040) {
                                $resolvedDate   = $sdDt
                                $resolvedSource = 'MPP Field'
                            }
                        } catch { }
                    }
                }
            } catch {
                try { $msp.FileClose(0) | Out-Null } catch { }
            }

            # Rest the shared drive between network copies (skipped when the
            # file was already staged, since the share wasn't touched)
            if ($tier1Copied -and $done -lt $total) {
                Invoke-PacedPause -Seconds ([int]$nudPauseSec.Value) -RespectCancel $false
            }

            # ── Tier 2: Date from filename ────────────────────────────────────
            if ($null -eq $resolvedDate) {
                $baseName = [System.IO.Path]::GetFileNameWithoutExtension($entry.FileName)
                $fnDate   = Get-FilenameDate -BaseName $baseName
                if ($null -ne $fnDate) {
                    $resolvedDate   = $fnDate
                    $resolvedSource = 'Filename'
                }
            }

            # ── Tier 3: Last Modified date (fallback) ─────────────────────────
            if ($null -eq $resolvedDate) {
                try {
                    $fileInfo       = Get-Item -LiteralPath $entry.FilePath -ErrorAction Stop
                    $resolvedDate   = $fileInfo.LastWriteTime
                    $resolvedSource = 'Modified*'   # asterisk = unreliable, warn user
                } catch { }
            }

            # ── Store result ──────────────────────────────────────────────────
            if ($null -ne $resolvedDate) {
                $entry.StatusDate  = $resolvedDate.ToString('yyyy-MM-dd')
                $entry.DateSource  = $resolvedSource
            } else {
                $entry.StatusDate  = ''
                $entry.DateSource  = 'None'
            }
        }
    } finally {
        if ($null -ne $msp) {
            try {
                $msp.Quit()
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($msp) | Out-Null
            } catch { }
        }
    }
}

function Refresh-PreviewGrid {
    $dgvPreview.Rows.Clear()
    $count = 0
    foreach ($f in $script:SelectedFiles) {
        $scoreDisplay = if ($f.Score -eq 9999) { 'none' } else { $f.Score }
        $ds = if ($f.DateSource) { $f.DateSource } else { '' }
        $rowIdx = $dgvPreview.Rows.Add($f.Include, $f.FileName, $f.SortDate, $scoreDisplay, $f.Folder, $ds, $f.StatusDate)
        # Color the DateSource cell to flag unreliable Modified* dates
        if ($ds -eq 'Modified*') {
            $dgvPreview.Rows[$rowIdx].Cells['DateSource'].Style.ForeColor = [System.Drawing.Color]::FromArgb(180,80,0)
            $dgvPreview.Rows[$rowIdx].Cells['DateSource'].Style.Font = New-Object System.Drawing.Font('Segoe UI', 8.5, [System.Drawing.FontStyle]::Bold)
        } elseif ($ds -eq 'MPP Field') {
            $dgvPreview.Rows[$rowIdx].Cells['DateSource'].Style.ForeColor = [System.Drawing.Color]::FromArgb(20,120,40)
        } elseif ($ds -eq 'Filename') {
            $dgvPreview.Rows[$rowIdx].Cells['DateSource'].Style.ForeColor = [System.Drawing.Color]::FromArgb(31,78,120)
        }
        $count++
    }
    $total   = $script:SelectedFiles.Count
    $checked = ($script:SelectedFiles | Where-Object { $_.Include }).Count
    $lblPreviewSummary.Text = "Found $total file(s) selected by filter rules.  $checked marked for export."
}

# Sync checkbox changes back to $script:SelectedFiles
$dgvPreview.Add_CellValueChanged({
    param($s, $e)
    if ($e.ColumnIndex -eq 0 -and $e.RowIndex -ge 0) {
        $script:SelectedFiles[$e.RowIndex].Include = [bool]$dgvPreview.Rows[$e.RowIndex].Cells['Include'].Value
        $checked = ($script:SelectedFiles | Where-Object { $_.Include }).Count
        $lblPreviewSummary.Text = "Found $($script:SelectedFiles.Count) file(s) selected by filter rules.  $checked marked for export."
    }
    if ($e.ColumnIndex -ge 0 -and $e.RowIndex -ge 0) {
        $colName = $dgvPreview.Columns[$e.ColumnIndex].Name
        if ($colName -eq 'StatusDate') {
            $newVal = [string]$dgvPreview.Rows[$e.RowIndex].Cells['StatusDate'].Value
            $script:SelectedFiles[$e.RowIndex].StatusDate = $newVal
            # Mark as manually overridden
            if ($newVal.Trim() -ne '') {
                $script:SelectedFiles[$e.RowIndex].DateSource = 'Manual'
                $dgvPreview.Rows[$e.RowIndex].Cells['DateSource'].Value = 'Manual'
                $dgvPreview.Rows[$e.RowIndex].Cells['DateSource'].Style.ForeColor = [System.Drawing.Color]::FromArgb(100,0,140)
                $dgvPreview.Rows[$e.RowIndex].Cells['DateSource'].Style.Font = New-Object System.Drawing.Font('Segoe UI', 8.5, [System.Drawing.FontStyle]::Italic)
            }
        }
    }
})

$dgvPreview.Add_CurrentCellDirtyStateChanged({
    if ($dgvPreview.IsCurrentCellDirty) { $dgvPreview.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit) }
})

# ── Rule set save/load handlers (in-app, no file system) ─────────────────
$btnSaveRules.Add_Click({
    if ($script:FilterRules.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('No rules to save. Add at least one rule first.','Save Rules',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    $name = $tbRSName.Text.Trim()
    if ($name -eq '') {
        [System.Windows.Forms.MessageBox]::Show('Please enter a name for this rule set.','Save Rules',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    # Deep-copy the current rules into the dictionary
    $copy = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($r in $script:FilterRules) {
        $copy.Add(@{
            Priority = $r.Priority
            Keyword  = $r.Keyword
            Logic    = $r.Logic
            GroupID  = $r.GroupID
            Scope    = $r.Scope
        })
    }
    $script:SavedRuleSets[$name] = $copy
    Refresh-RuleSetsDropdown
    # Select the one we just saved
    $idx = $cboRuleSets.Items.IndexOf($name)
    if ($idx -ge 0) { $cboRuleSets.SelectedIndex = $idx }
    [System.Windows.Forms.MessageBox]::Show("Saved $($copy.Count) rule(s) as '$name'.",'Saved',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
})

$btnLoadRules.Add_Click({
    if ($cboRuleSets.Items.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('No saved rule sets yet. Add rules and click Save Current first.','Load Rules',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    $name = $cboRuleSets.SelectedItem.ToString()
    $saved = $script:SavedRuleSets[$name]
    $ans = [System.Windows.Forms.MessageBox]::Show(
        "Load '$name' ($($saved.Count) rule(s))?`nThis will replace your current rules.",
        'Load Rules',
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($ans -eq 'Yes') {
        $script:FilterRules.Clear()
        foreach ($r in $saved) {
            $script:FilterRules.Add(@{
                Priority = $r.Priority
                Keyword  = $r.Keyword
                Logic    = $r.Logic
                GroupID  = $r.GroupID
                Scope    = $r.Scope
            })
        }
        Refresh-RulesGrid
    }
})

$btnDeleteRS.Add_Click({
    if ($cboRuleSets.Items.Count -eq 0) { return }
    $name = $cboRuleSets.SelectedItem.ToString()
    $ans = [System.Windows.Forms.MessageBox]::Show(
        "Delete rule set '$name'?",
        'Delete',
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning)
    if ($ans -eq 'Yes') {
        $script:SavedRuleSets.Remove($name) | Out-Null
        Refresh-RuleSetsDropdown
    }
})

# Scan button wires into preview

$btnScan.Add_Click({
    $src = $tbSource.Text.Trim()
    if ($src -eq '' -or -not (Test-Path $src)) {
        [System.Windows.Forms.MessageBox]::Show('Please set a valid Source Folder on the Setup tab.','Validation',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    $lblPreviewSummary.Text = 'Scanning…'
    [System.Windows.Forms.Application]::DoEvents()

    Invoke-Scan -SourcePath $src `
                -MaxPerFolder ([int]$nudLimit.Value) `
                -Rules $script:FilterRules `
                -ExcludeFoldersCsv $tbExclude.Text

    Read-StatusDates
    Refresh-PreviewGrid
    $tabs.SelectedIndex = 2
})

$btnRescan.Add_Click({
    $src = $tbSource.Text.Trim()
    if ($src -eq '' -or -not (Test-Path $src)) {
        [System.Windows.Forms.MessageBox]::Show('Please set a valid Source Folder on the Setup tab.','Validation',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    $lblPreviewSummary.Text = 'Scanning…'
    [System.Windows.Forms.Application]::DoEvents()
    Invoke-Scan -SourcePath $src `
                -MaxPerFolder ([int]$nudLimit.Value) `
                -Rules $script:FilterRules `
                -ExcludeFoldersCsv $tbExclude.Text
    Read-StatusDates
    Refresh-PreviewGrid
})

# =============================================================================
# TAB 4 — EXPORT
# =============================================================================
$tabExport       = New-Object System.Windows.Forms.TabPage
$tabExport.Text  = '  4 · Export  '
$tabExport.BackColor = [System.Drawing.Color]::FromArgb(245,245,248)
$tabs.TabPages.Add($tabExport)

# ── Export options ────────────────────────────────────────────────────────
$gbExportOpts = New-GroupBox 'Export Options' 12 12 920 148
$tabExport.Controls.Add($gbExportOpts)

$chkStatusDate = New-Object System.Windows.Forms.CheckBox
$chkStatusDate.Text = 'Embed Status Date in exported XML'
$chkStatusDate.Left = 12; $chkStatusDate.Top = 28
$chkStatusDate.Width = 280; $chkStatusDate.Checked = $true
$chkStatusDate.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$gbExportOpts.Controls.Add($chkStatusDate)

$lblStatusDateSrc = New-Label 'Source:' 300 30 50
$gbExportOpts.Controls.Add($lblStatusDateSrc)

$cboStatusDateSrc = New-Object System.Windows.Forms.ComboBox
$cboStatusDateSrc.Left = 352; $cboStatusDateSrc.Top = 28; $cboStatusDateSrc.Width = 200
$cboStatusDateSrc.DropDownStyle = 'DropDownList'
$cboStatusDateSrc.Items.AddRange(@('Read from MPP (auto)','Manual override for all files')) | Out-Null
$cboStatusDateSrc.SelectedIndex = 0
$gbExportOpts.Controls.Add($cboStatusDateSrc)

$lblManualDate = New-Label 'Override date (yyyy-MM-dd):' 560 30 190
$gbExportOpts.Controls.Add($lblManualDate)

$tbManualDate = New-Textbox 752 28 200 ''
$tbManualDate.Enabled = $false
$gbExportOpts.Controls.Add($tbManualDate)

$lblStatusDateNote = New-Label 'You can also override per-file in the Preview tab (Status Date Override column).' 12 62 880 20
$lblStatusDateNote.ForeColor = [System.Drawing.Color]::Gray
$gbExportOpts.Controls.Add($lblStatusDateNote)

$lblStatusDateNote2 = New-Label 'Per-file override takes precedence over the global source setting above.' 12 82 880 20
$lblStatusDateNote2.ForeColor = [System.Drawing.Color]::FromArgb(150,80,0)
$gbExportOpts.Controls.Add($lblStatusDateNote2)

# ── Batch size ──────────────────────────────────────────────────────────
$lblBatchSize = New-Label 'MS Project batch size:' 12 108 150
$lblBatchSize.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$gbExportOpts.Controls.Add($lblBatchSize)

$nudBatchSize = New-Object System.Windows.Forms.NumericUpDown
$nudBatchSize.Left = 162; $nudBatchSize.Top = 106
$nudBatchSize.Width = 70; $nudBatchSize.Height = 24
$nudBatchSize.Minimum = 10; $nudBatchSize.Maximum = 500; $nudBatchSize.Value = 100
$gbExportOpts.Controls.Add($nudBatchSize)

$lblPauseSec = New-Label 'Pause between files (sec):' 250 108 170
$lblPauseSec.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$gbExportOpts.Controls.Add($lblPauseSec)

$nudPauseSec = New-Object System.Windows.Forms.NumericUpDown
$nudPauseSec.Left = 424; $nudPauseSec.Top = 106
$nudPauseSec.Width = 60; $nudPauseSec.Height = 24
$nudPauseSec.Minimum = 0; $nudPauseSec.Maximum = 60; $nudPauseSec.Value = 3
$gbExportOpts.Controls.Add($nudPauseSec)

$lblBatchNote = New-Label 'Batch restarts prevent memory leaks. The pause rests the shared drive between file copies.' 500 110 410 20
$lblBatchNote.ForeColor = [System.Drawing.Color]::Gray
$lblBatchNote.Font = New-Object System.Drawing.Font('Segoe UI', 8.5)
$gbExportOpts.Controls.Add($lblBatchNote)

$chkStatusDate.Add_CheckedChanged({
    $en = $chkStatusDate.Checked
    $cboStatusDateSrc.Enabled = $en
    $tbManualDate.Enabled     = ($en -and $cboStatusDateSrc.SelectedIndex -eq 1)
    $lblStatusDateSrc.Enabled = $en
    $lblManualDate.Enabled    = $en
})

$cboStatusDateSrc.Add_SelectedIndexChanged({
    $tbManualDate.Enabled = ($chkStatusDate.Checked -and $cboStatusDateSrc.SelectedIndex -eq 1)
})

# ── Progress ──────────────────────────────────────────────────────────────
$gbProgress = New-GroupBox 'Export Progress' 12 168 920 180
$tabExport.Controls.Add($gbProgress)

$pgBar = New-Object System.Windows.Forms.ProgressBar
$pgBar.Left = 12; $pgBar.Top = 28; $pgBar.Width = 890; $pgBar.Height = 24
$pgBar.Style = 'Continuous'
$gbProgress.Controls.Add($pgBar)

$pnlStats = New-Object System.Windows.Forms.Panel
$pnlStats.Left = 12; $pnlStats.Top = 60; $pnlStats.Width = 890; $pnlStats.Height = 28
$gbProgress.Controls.Add($pnlStats)

$lblProcessed = New-Label 'Processed: 0 / 0' 0 4 220
$lblProcessed.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$lblElapsed   = New-Label 'Elapsed: 0s' 230 4 160
$lblETA       = New-Label 'ETA: —' 400 4 200
$lblErrors    = New-Label 'Errors: 0' 610 4 160
$pnlStats.Controls.AddRange(@($lblProcessed, $lblElapsed, $lblETA, $lblErrors))

$pnlExportBtns = New-Object System.Windows.Forms.Panel
$pnlExportBtns.Left = 12; $pnlExportBtns.Top = 96; $pnlExportBtns.Width = 890; $pnlExportBtns.Height = 36
$gbProgress.Controls.Add($pnlExportBtns)

$btnStartExport  = New-Btn '▶  Start Export' 0 4 160 30 ([System.Drawing.Color]::FromArgb(20,120,40))
$btnStartExport.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
$btnCancelExport = New-Btn '■  Cancel'        170 4 100 30 ([System.Drawing.Color]::FromArgb(160,30,30))
$btnCancelExport.Enabled = $false
$pnlExportBtns.Controls.AddRange(@($btnStartExport, $btnCancelExport))

# ── Log ───────────────────────────────────────────────────────────────────
$gbLog = New-GroupBox 'Export Log' 12 328 920 330
$tabExport.Controls.Add($gbLog)

$rtbLog = New-Object System.Windows.Forms.RichTextBox
$rtbLog.Left = 8; $rtbLog.Top = 22
$rtbLog.Width = 900; $rtbLog.Height = 298
$rtbLog.ReadOnly   = $true
$rtbLog.BackColor  = [System.Drawing.Color]::FromArgb(20,20,30)
$rtbLog.ForeColor  = [System.Drawing.Color]::FromArgb(200,220,200)
$rtbLog.Font       = New-Object System.Drawing.Font('Consolas', 8.5)
$rtbLog.ScrollBars = 'Vertical'
$gbLog.Controls.Add($rtbLog)

function Write-Log {
    param($Message, $Color = $null)
    $ts = (Get-Date).ToString('HH:mm:ss')
    $rtbLog.SelectionStart  = $rtbLog.TextLength
    $rtbLog.SelectionLength = 0
    if ($null -ne $Color) {
        $rtbLog.SelectionColor = $Color
    } else {
        $rtbLog.SelectionColor = [System.Drawing.Color]::FromArgb(200,220,200)
    }
    $rtbLog.AppendText("[$ts] $Message`n")
    $rtbLog.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

# ── Export logic ──────────────────────────────────────────────────────────
$btnStartExport.Add_Click({
    $destPath     = $tbDest.Text.Trim()
    $folderPrefix = $tbFolderName.Text.Trim()
    if ($folderPrefix -eq '') { $folderPrefix = 'PipelineOutput' }
    $folderPrefix = $folderPrefix -replace '[\\/:*?"<>|]', '_'   # sanitise invalid chars
    $runFolder    = "$folderPrefix $(Get-Date -Format 'yyyy-MM-dd')"
    $destPath     = Join-Path $destPath $runFolder
    if (-not (Test-Path $destPath)) {
        New-Item -ItemType Directory -Path $destPath -Force | Out-Null
    }
    if ($destPath -eq '') {
        [System.Windows.Forms.MessageBox]::Show('Please set the Output Folder on the Setup tab.','Validation',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    $toExport = @($script:SelectedFiles | Where-Object { $_.Include })
    if ($toExport.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('No files are checked for export. Go to the Preview tab and check at least one file.','Nothing to Export',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    $useStatusDate = $chkStatusDate.Checked
    $sdSource      = $cboStatusDateSrc.SelectedIndex   # 0=auto, 1=manual
    $sdManual      = $tbManualDate.Text.Trim()

    if (-not (Test-Path $destPath)) {
        New-Item -ItemType Directory -Path $destPath -Force | Out-Null
    }

    $script:ExportCancelled = $false
    $btnStartExport.Enabled  = $false
    $btnCancelExport.Enabled = $true
    $pgBar.Maximum   = $toExport.Count
    $pgBar.Value     = 0
    $errorCount      = 0
    $processedCount  = 0
    $rtbLog.Clear()

    Write-Log "Starting export of $($toExport.Count) file(s) → $destPath" ([System.Drawing.Color]::FromArgb(150,220,255))

    # ── Helper: launch MS Project COM ────────────────────────────────────
    function Start-MSProject {
        $app = $null
        for ($i = 1; $i -le 10; $i++) {
            try {
                $app = New-Object -ComObject 'MSProject.Application'
                Start-Sleep -Seconds 3
                $app.Visible       = $false
                $app.DisplayAlerts = $false
                Write-Log "MS Project ready (attempt $i)." ([System.Drawing.Color]::FromArgb(100,255,100))
                return $app
            } catch {
                Write-Log "Waiting for MS Project… (attempt $i/10)" ([System.Drawing.Color]::FromArgb(255,200,80))
                $app = $null
                Start-Sleep -Seconds 3
            }
        }
        return $null
    }

    function Stop-MSProject { param($app)
        try {
            $app.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            Start-Sleep -Seconds 2
            Write-Log 'MS Project closed cleanly.' ([System.Drawing.Color]::FromArgb(150,150,150))
        } catch {
            Write-Log 'Warning: MS Project may not have closed cleanly.' ([System.Drawing.Color]::FromArgb(255,160,40))
        }
    }

    $batchSize     = [int]$nudBatchSize.Value
    $pauseSec      = [int]$nudPauseSec.Value
    $startTime     = Get-Date
    $batchNum      = 0
    $fileIndex     = 0

    # Staging dir must exist even if an individual Get-StagedCopy fails —
    # the XML is always written locally first, then moved to the destination.
    if (-not (Test-Path -LiteralPath $script:StagingDir)) {
        New-Item -ItemType Directory -Path $script:StagingDir -Force | Out-Null
    }

    Write-Log "Batch size: $batchSize files per MS Project session" ([System.Drawing.Color]::FromArgb(150,220,255))
    Write-Log "Pause between files: ${pauseSec}s  |  Staging: $($script:StagingDir)" ([System.Drawing.Color]::FromArgb(150,220,255))

    Write-Log 'Launching MS Project COM…'
    $msp = Start-MSProject
    if ($null -eq $msp) {
        Write-Log 'ERROR: MS Project failed to initialize.' ([System.Drawing.Color]::FromArgb(255,80,80))
        $btnStartExport.Enabled  = $true
        $btnCancelExport.Enabled = $false
        return
    }
    $batchNum++
    $batchCount = 0
    Write-Log "Batch 1 started." ([System.Drawing.Color]::FromArgb(150,220,255))

    foreach ($entry in $toExport) {
        if ($script:ExportCancelled) {
            Write-Log 'Export cancelled by user.' ([System.Drawing.Color]::FromArgb(255,160,40))
            break
        }

        # ── Restart MS Project if batch size reached ──────────────────────
        if ($batchCount -gt 0 -and ($batchCount % $batchSize) -eq 0) {
            Write-Log "Batch $batchNum complete ($batchSize files). Restarting MS Project…" ([System.Drawing.Color]::FromArgb(255,200,80))
            Stop-MSProject $msp
            $msp = Start-MSProject
            if ($null -eq $msp) {
                Write-Log 'ERROR: MS Project failed to restart.' ([System.Drawing.Color]::FromArgb(255,80,80))
                break
            }
            $batchNum++
            Write-Log "Batch $batchNum started." ([System.Drawing.Color]::FromArgb(150,220,255))
        }

        $srcFile   = $entry.FilePath
        $baseName  = [System.IO.Path]::GetFileNameWithoutExtension($entry.FileName)
        # All XML files go flat into the single dated run folder
        $outFile  = Join-Path $destPath "$baseName.xml"
        # MS Project reads/writes only local paths; the XML is moved to the
        # destination after the file is closed.
        $localXml = Join-Path $script:StagingDir "$baseName.xml"

        Write-Log "Exporting: $($entry.FileName)"

        # Stage the MPP locally — one sequential read of the share instead of
        # MS Project's random-access I/O over the network
        $openPath = $srcFile
        try {
            $stage    = Get-StagedCopy -SourcePath $srcFile
            $openPath = $stage.Path
            if ($stage.Copied) {
                Write-Log "  Staged locally." ([System.Drawing.Color]::FromArgb(150,150,150))
            } else {
                Write-Log "  Re-using copy staged during Preview." ([System.Drawing.Color]::FromArgb(150,150,150))
            }
        } catch {
            Write-Log "  Could not stage locally ($($_.Exception.Message)) — opening from source." ([System.Drawing.Color]::FromArgb(255,200,80))
        }

        $maxRetries = 5
        $attempt    = 0
        $success    = $false

        while ($attempt -lt $maxRetries -and -not $success -and -not $script:ExportCancelled) {
            $attempt++
            try {
                $msp.DisplayAlerts = $false
                $msp.FileOpen($openPath, $true) | Out-Null
                Start-Sleep -Milliseconds 900

                # Read status date from MPP if needed
                $effectiveSD = ''
                if ($useStatusDate) {
                    if ($entry.StatusDate -and $entry.StatusDate.Trim() -ne '') {
                        # Per-file override from preview grid takes priority
                        $effectiveSD = $entry.StatusDate.Trim()
                        Write-Log "  Status date (per-file override): $effectiveSD" ([System.Drawing.Color]::FromArgb(180,180,255))
                    } elseif ($sdSource -eq 1 -and $sdManual -ne '') {
                        $effectiveSD = $sdManual
                        Write-Log "  Status date (manual global): $effectiveSD" ([System.Drawing.Color]::FromArgb(180,180,255))
                    } elseif ($sdSource -eq 0) {
                        # Read from MPP - handle all forms of "not set"
                        try {
                            $sdRaw = $msp.ActiveProject.StatusDate
                            $sdOk  = $false
                            if ($null -ne $sdRaw) {
                                $sdStr = "$sdRaw".Trim()
                                # MS Project returns "NA" or empty when not set
                                if ($sdStr -ne '' -and $sdStr -ne 'NA') {
                                    try {
                                        $sdDt = [datetime]$sdRaw
                                        # The magic "not set" sentinel is Jan 1 1984 regardless of locale
                                        if ($sdDt.Year -ne 1984) {
                                            $effectiveSD = $sdDt.ToString('yyyy-MM-dd')
                                            $sdOk = $true
                                            Write-Log "  Status date (from MPP): $effectiveSD" ([System.Drawing.Color]::FromArgb(180,180,255))
                                        }
                                    } catch { }
                                }
                            }
                            if (-not $sdOk) {
                                Write-Log "  Status date: not set in this MPP file." ([System.Drawing.Color]::Gray)
                            }
                        } catch {
                            Write-Log "  Could not read status date: $($_.Exception.Message)" ([System.Drawing.Color]::Gray)
                        }
                    }
                }

                $msp.FileSaveAs($localXml) | Out-Null
                Start-Sleep -Milliseconds 500
                $msp.FileClose(0) | Out-Null

                if (Test-Path $localXml) {
                    # Inject status date into XML (still local — no network I/O)
                    if ($useStatusDate -and $effectiveSD -ne '') {
                        Set-XmlStatusDate -XmlPath $localXml -StatusDate $effectiveSD
                        # Verify the date actually landed in the XML
                        try {
                            [xml]$verifyDoc = Get-Content -LiteralPath $localXml -Encoding UTF8
                            $nsm2 = New-Object System.Xml.XmlNamespaceManager($verifyDoc.NameTable)
                            $nsm2.AddNamespace('p','http://schemas.microsoft.com/project')
                            $verifyNode = $verifyDoc.SelectSingleNode('//p:StatusDate', $nsm2)
                            if ($verifyNode -and $verifyNode.InnerText -ne '') {
                                Write-Log "  ✔ Status date verified in XML: $($verifyNode.InnerText)" ([System.Drawing.Color]::FromArgb(100,255,180))
                            } else {
                                Write-Log "  ⚠ Status date not found in XML after injection." ([System.Drawing.Color]::FromArgb(255,160,40))
                            }
                        } catch {
                            Write-Log "  Could not verify status date in XML." ([System.Drawing.Color]::Gray)
                        }
                    }
                    # One sequential write to the destination
                    Move-Item -LiteralPath $localXml -Destination $outFile -Force
                    $success = $true
                    Write-Log "  → $outFile" ([System.Drawing.Color]::FromArgb(100,255,100))
                } else {
                    throw "Output file not created: $localXml"
                }

            } catch {
                $msg = $_.Exception.Message
                if ($msg -match 'RPC_E_CALL_REJECTED' -or $msg -match '0x80010001') {
                    Write-Log "  [Retry $attempt/$maxRetries] MS Project busy…" ([System.Drawing.Color]::FromArgb(255,200,80))
                    try { $msp.FileClose(0) | Out-Null } catch {}
                    Start-Sleep -Seconds (3 * $attempt)
                } else {
                    Write-Log "  ERROR: $msg" ([System.Drawing.Color]::FromArgb(255,80,80))
                    $errorCount++
                    try { $msp.FileClose(0) | Out-Null } catch {}
                    break
                }
            }
        }

        if (-not $success) {
            if ($attempt -ge $maxRetries) {
                Write-Log "  FAILED after $maxRetries retries." ([System.Drawing.Color]::FromArgb(255,80,80))
            }
            $errorCount++
        }

        $processedCount++
        $batchCount++
        $pgBar.Value = $processedCount

        # Update stats
        $elapsed = (Get-Date) - $startTime
        $rate    = if ($processedCount -gt 0) { $elapsed.TotalSeconds / $processedCount } else { 0 }
        $remain  = [math]::Max(0, ($toExport.Count - $processedCount) * $rate)
        $etaStr  = if ($remain -gt 60) { '{0}m {1}s' -f [int]($remain / 60), [int]($remain % 60) } else { '{0}s' -f [int]$remain }

        $lblProcessed.Text = "Processed: $processedCount / $($toExport.Count)"
        $lblElapsed.Text   = 'Elapsed: {0}m {1}s' -f [int]$elapsed.TotalMinutes, ($elapsed.Seconds)
        $lblETA.Text       = "ETA: $etaStr"
        $lblErrors.Text    = "Errors: $errorCount"
        [System.Windows.Forms.Application]::DoEvents()

        # Rest the shared drive between files
        if ($pauseSec -gt 0 -and $processedCount -lt $toExport.Count) {
            Invoke-PacedPause -Seconds $pauseSec
        }
    }

    # ── Final cleanup ─────────────────────────────────────────────────────
    if ($null -ne $msp) { Stop-MSProject $msp }

    $elapsed = (Get-Date) - $startTime
    Write-Log ('─' * 60)
    Write-Log ('DONE.  Exported: {0}  |  Errors: {1}  |  Total time: {2}m {3}s' -f
        $processedCount, $errorCount,
        [int]$elapsed.TotalMinutes, ($elapsed.Seconds)) ([System.Drawing.Color]::FromArgb(150,220,255))

    $btnStartExport.Enabled  = $true
    $btnCancelExport.Enabled = $false
    $pgBar.Value = $pgBar.Maximum

    [System.Windows.Forms.MessageBox]::Show(
        "Export complete.`nProcessed: $processedCount`nErrors: $errorCount`nOutput: $destPath",
        'Export Complete',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information)
})

$btnCancelExport.Add_Click({
    $script:ExportCancelled = $true
    Write-Log 'Cancel requested — finishing current file…' ([System.Drawing.Color]::FromArgb(255,160,40))
    $btnCancelExport.Enabled = $false
})

# =============================================================================
# LAUNCH
# =============================================================================
Clear-StagingDir   # remove staged copies left over from a previous session

$form.Add_Shown({
    $tabs.Dock = [System.Windows.Forms.DockStyle]::Fill
    $form.Refresh()
    $form.Activate()
})
$form.Add_FormClosed({ Clear-StagingDir })
[System.Windows.Forms.Application]::Run($form)
