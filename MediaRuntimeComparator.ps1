# Load required .NET assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Get script directory
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$global:cfgPath = Join-Path $scriptDir "MediaRuntimeComparator.cfg"

# Default config values
$defaultConfig = @{
    FFprobePath      = ""
    LastMediaFolder  = $scriptDir
    ScriptFolder     = $scriptDir
    ToleranceSeconds = 0
    WindowHeight     = 600
    WindowWidth      = 1000
}

# Load or create config
$global:cfg = @{}
if (Test-Path $global:cfgPath) {
    Get-Content $global:cfgPath | ForEach-Object {
        if ($_ -match '^(.+?)=(.*)$') {
            $global:cfg[$matches[1]] = $matches[2]
        }
    }
}
foreach ($key in $defaultConfig.Keys) {
    if (-not $global:cfg.ContainsKey($key)) {
        $global:cfg[$key] = $defaultConfig[$key]
    }
}
$global:cfg.GetEnumerator() | Sort-Object Name | ForEach-Object {
    "$($_.Key)=$($_.Value)" | Out-File $global:cfgPath -Encoding UTF8 -Append:$false
}

# Globals
$global:mediaData = @()
$global:stopScan = $false

# Create main form
$mainForm = New-Object Windows.Forms.Form
$mainForm.Text = "Media Runtime Comparator for ffprobe (FFMPEG tools)"
$mainForm.Size = New-Object Drawing.Size([int]$global:cfg.WindowWidth, [int]$global:cfg.WindowHeight)
$mainForm.StartPosition = "CenterScreen"
$mainForm.FormBorderStyle = 'Sizable'

# Create ListView
$global:listView = New-Object Windows.Forms.ListView
$global:listView.View = 'Details'
$global:listView.FullRowSelect = $true
$global:listView.GridLines = $true
$global:listView.Dock = 'Fill'
[void]$global:listView.Columns.Add("File",400)
[void]$global:listView.Columns.Add("Duration",100)
[void]$global:listView.Columns.Add("Seconds",100)
[void]$global:listView.Columns.Add("Size (KB)",100)

# Create ProgressBar
$global:progressBar = New-Object Windows.Forms.ProgressBar
$global:progressBar.Dock = 'Bottom'
$global:progressBar.Height = 18

# Create StatusBar with 3 panels
$global:statusBar = New-Object Windows.Forms.StatusBar
$global:statusBar.SizingGrip = $false
$global:statusBar.ShowPanels = $true

$global:statusPanelLeft = New-Object Windows.Forms.StatusBarPanel
$global:statusPanelLeft.Width = 300
$global:statusPanelLeft.AutoSize = 'None'
$global:statusPanelLeft.Text = "[STATUS] Ready."

$global:statusPanelMiddle = New-Object Windows.Forms.StatusBarPanel
$global:statusPanelMiddle.Width = 300
$global:statusPanelMiddle.AutoSize = 'None'
$global:statusPanelMiddle.Text = ""

$global:statusPanelRight = New-Object Windows.Forms.StatusBarPanel
$global:statusPanelRight.Alignment = 'Right'
$global:statusPanelRight.AutoSize = 'Spring'
$global:statusPanelRight.Text = "Click Scan to begin. Use Stop to interrupt."

$global:statusBar.Panels.AddRange(@($global:statusPanelLeft, $global:statusPanelMiddle, $global:statusPanelRight))

# Utility: Write status
function Write-Status {
    param([string]$msg,[System.Drawing.Color]$color=$null)
    $global:statusPanelLeft.Text = "[STATUS] $msg"
    if ($color) { $global:statusBar.ForeColor = $color }
    [System.Windows.Forms.Application]::DoEvents()
    Write-Host $msg
}

# Utility: Update instructions
function Update-Instructions {
    param([string]$msg)
    $global:statusPanelRight.Text = $msg
    [System.Windows.Forms.Application]::DoEvents()
}

# Utility: Update parse error summary
function Update-ParseSummary {
    param([int]$errors, [int]$total)
    $percent = if ($total -eq 0) { 0 } else { [math]::Round(($errors / $total) * 100, 1) }
    $global:statusPanelMiddle.Text = "ffprobe parsing errors: $errors of $total files ($percent%)"
    [System.Windows.Forms.Application]::DoEvents()
}

# Utility: Convert seconds to HH:MM:SS
function Convert-Time([double]$seconds) {
    $ts=[TimeSpan]::FromSeconds($seconds)
    "{0:D2}:{1:D2}:{2:D2}" -f [int]$ts.TotalHours,$ts.Minutes,$ts.Seconds
}

# Utility: Validate ffprobe path
function Check-FFprobe {
    $path=$global:ffprobeTextBox.Text
    if ([string]::IsNullOrWhiteSpace($path)) {
        $global:ffprobePath=""
        $global:ffprobeTextBox.BackColor=[Drawing.Color]::LightCoral
        Write-Status "FFprobe path is empty." ([Drawing.Color]::Red)
        return $false
    }
    if (Test-Path $path -PathType Leaf) {
        $global:ffprobePath=$path
        $global:ffprobeTextBox.BackColor=[Drawing.Color]::LightGreen
        Write-Status "FFprobe ready." ([Drawing.Color]::Green)
        return $true
    } else {
        $global:ffprobePath=""
        $global:ffprobeTextBox.BackColor=[Drawing.Color]::LightCoral
        Write-Status "FFprobe not found." ([Drawing.Color]::Red)
        return $false
    }
}

# Utility: Highlight duplicates
function Highlight-Duplicates {
    param([double]$Tolerance)
    for ($i = 0; $i -lt $global:mediaData.Count; $i++) {
        $ref = $global:mediaData[$i]
        $matches = $global:mediaData | Where-Object {
            ($_ -ne $ref) -and ([math]::Abs($_.DurationSeconds - $ref.DurationSeconds) -le $Tolerance)
        }
        if ($matches) {
            foreach ($item in $global:listView.Items) {
                if ($item.Text -eq $ref.FileName) {
                    $item.BackColor = [Drawing.Color]::Yellow
                    break
                }
            }
        }
    }
}

function Invoke-MediaScan {
    $global:stopScan = $false
    if (-not (Check-FFprobe)) { return }

    $folderPath = $global:folderTextBox.Text
    if (-not (Test-Path $folderPath)) {
        Write-Status "Folder not found: $folderPath" ([Drawing.Color]::Red)
        return
    }

    # Clamp tolerance input
    $tolerance = 0
    if (-not [double]::TryParse($global:toleranceTextBox.Text, [ref]$tolerance)) {
        $tolerance = 0
    }
    $tolerance = [math]::Min([math]::Max($tolerance, 0), 60)
    $global:toleranceTextBox.Text = $tolerance.ToString()

    $global:listView.Items.Clear()
    $global:mediaData = @()

    $files = Get-ChildItem $folderPath -Recurse -File |
        Where-Object { $_.Extension -match '^(?i)\.(mp4|mkv|avi|mov|wmv|flv|webm|ts|m2ts|m4v|mpg|mpeg|3gp|3g2|ogg|ogv|ogm|vob|divx|rm|rmvb|asf|f4v|mxf|mts|m2v|mp2|mp3|aac|wav|flac|alac|wma|m4a|opus|aiff|au|ac3|dts|amr|caf)$' }

    if (-not $files) {
        Write-Status "No media files found." ([Drawing.Color]::Red)
        Update-ParseSummary 0 0
        return
    }

    $global:progressBar.Maximum = $files.Count
    $global:progressBar.Value = 0

    Write-Status "Scanning $($files.Count) files in $folderPath" ([Drawing.Color]::Blue)

    $parseErrors = 0
    $totalFiles = $files.Count

    foreach ($f in $files) {
        if ($global:stopScan) {
            Write-Status "Scan stopped by user." ([Drawing.Color]::OrangeRed)
            Update-ParseSummary $parseErrors $totalFiles
            break
        }

        try {
            $out = (& $global:ffprobePath -v error -show_entries format=duration `
                     -of default=noprint_wrappers=1:nokey=1 "$($f.FullName)" 2>&1).Trim()
            if ($out -match '^\d+(\.\d+)?$') {
                $dur = [double]::Parse($out, [Globalization.CultureInfo]::InvariantCulture)
                $item = [PSCustomObject]@{
                    FileName          = $f.Name
                    FullName          = $f.FullName
                    DurationSeconds   = $dur
                    DurationFormatted = (Convert-Time $dur)
                    SizeKB            = [math]::Round($f.Length / 1KB)
                }
                $global:mediaData += $item

                $lvi = New-Object Windows.Forms.ListViewItem($item.FileName)
                $lvi.SubItems.Add($item.DurationFormatted)
                $lvi.SubItems.Add(("{0:F6}" -f $item.DurationSeconds))
                $lvi.SubItems.Add(("{0:N0}" -f $item.SizeKB))
                $global:listView.Items.Add($lvi)
            } else {
                $parseErrors++
                Write-Host "ERROR: $($f.Name) ffprobe parse error: $out" -ForegroundColor Red
            }
        } catch {
            $parseErrors++
            Write-Host "EXCEPTION: $($f.Name) $($_.Exception.Message)" -ForegroundColor Red
        }

        $global:progressBar.Increment(1)
        Update-ParseSummary $parseErrors $totalFiles
        [System.Windows.Forms.Application]::DoEvents()
    }

    $global:progressBar.Value = 0
    Highlight-Duplicates -Tolerance $tolerance

    Write-Status "Scan complete. $($global:mediaData.Count) files." ([Drawing.Color]::Green)
    Update-ParseSummary $parseErrors $totalFiles
    Update-Instructions "Scan complete. Right-click to delete. Double-click to open."
}

# Layout container
$layout = New-Object Windows.Forms.TableLayoutPanel
$layout.Dock = 'Fill'
$layout.RowCount = 2
$layout.ColumnCount = 1
$layout.RowStyles.Add((New-Object Windows.Forms.RowStyle('Percent',100)))
$layout.RowStyles.Add((New-Object Windows.Forms.RowStyle('Absolute',80)))
$mainForm.Controls.Add($layout)

# Add ListView to layout
$layout.Controls.Add($global:listView,0,0)

# Bottom panel for controls
$bottomControls = New-Object Windows.Forms.Panel
$bottomControls.Dock = 'Fill'
$layout.Controls.Add($bottomControls,0,1)

# Folder selection
$lblFolder = New-Object Windows.Forms.Label
$lblFolder.Text = "Folder:"; $lblFolder.Location = New-Object Drawing.Point(10,12); $lblFolder.AutoSize = $true
$global:folderTextBox = New-Object Windows.Forms.TextBox
$global:folderTextBox.Width = 350; $global:folderTextBox.Location = New-Object Drawing.Point(70,10)
$global:folderTextBox.Text = $global:cfg.LastMediaFolder
$btnBrowse = New-Object Windows.Forms.Button
$btnBrowse.Text = "Browse"; $btnBrowse.Width = 70; $btnBrowse.Location = New-Object Drawing.Point(430,8)
$btnBrowse.Add_Click({
    $dlg = New-Object Windows.Forms.FolderBrowserDialog
    $currentPath = $global:folderTextBox.Text
    if (-not [string]::IsNullOrWhiteSpace($currentPath) -and (Test-Path $currentPath)) {
        $dlg.SelectedPath = $currentPath
    }
    if ($dlg.ShowDialog() -eq 'OK') {
        $global:folderTextBox.Text = $dlg.SelectedPath
    }
})

# Scan button
$btnScan = New-Object Windows.Forms.Button
$btnScan.Text = "Scan"; $btnScan.Width = 60; $btnScan.Location = New-Object Drawing.Point(510,8)
$btnScan.Add_Click({
    $btnScan.Enabled = $false
    $btnStop.Enabled = $true
    Update-Instructions "Scanning... You can press Stop to cancel."
    Invoke-MediaScan
    $btnScan.Enabled = $true
    $btnStop.Enabled = $false
})

# Stop button
$btnStop = New-Object Windows.Forms.Button
$btnStop.Text = "Stop"; $btnStop.Width = 60; $btnStop.Location = New-Object Drawing.Point(580,8)
$btnStop.Enabled = $false
$btnStop.Add_Click({
    $global:stopScan = $true
    Write-Status "Stopping scan..." ([Drawing.Color]::OrangeRed)
    Update-Instructions "Scan interrupted. You may adjust settings and retry."
})

# FFprobe path
$lblFF = New-Object Windows.Forms.Label
$lblFF.Text = "FFprobe:"; $lblFF.Location = New-Object Drawing.Point(10,42); $lblFF.AutoSize = $true
$global:ffprobeTextBox = New-Object Windows.Forms.TextBox
$global:ffprobeTextBox.Width = 350; $global:ffprobeTextBox.Location = New-Object Drawing.Point(70,40)

# Restore FFprobe path from config or fallback
$ffprobeFromCfg = $global:cfg.FFprobePath
$ffprobeInScript = Join-Path $scriptDir "ffprobe.exe"
if (-not [string]::IsNullOrWhiteSpace($ffprobeFromCfg) -and (Test-Path $ffprobeFromCfg)) {
    $global:ffprobeTextBox.Text = $ffprobeFromCfg
} elseif (Test-Path $ffprobeInScript) {
    $global:ffprobeTextBox.Text = $ffprobeInScript
} else {
    $global:ffprobeTextBox.Text = ""
}

$btnFF = New-Object Windows.Forms.Button
$btnFF.Text = "Browse"; $btnFF.Width = 70; $btnFF.Location = New-Object Drawing.Point(430,38)
$btnFF.Add_Click({
    $dlg = New-Object Windows.Forms.OpenFileDialog
    $dlg.Filter = "ffprobe.exe|ffprobe.exe|All files|*.*"
    if ($dlg.ShowDialog() -eq 'OK') {
        $global:ffprobeTextBox.Text = $dlg.FileName
        Check-FFprobe | Out-Null
    }
})

# Tolerance input
$lblTol = New-Object Windows.Forms.Label
$lblTol.Text = "Tolerance (s) [0–60]:"; $lblTol.AutoSize = $true; $lblTol.Location = New-Object Drawing.Point(510,42)
$global:toleranceTextBox = New-Object Windows.Forms.TextBox
$global:toleranceTextBox.Width = 60; $global:toleranceTextBox.Location = New-Object Drawing.Point(630,40)
$global:toleranceTextBox.Text = $global:cfg.ToleranceSeconds

# Add controls to bottom panel
$bottomControls.Controls.AddRange(@(
    $lblFolder,$global:folderTextBox,$btnBrowse,$btnScan,$btnStop,
    $lblFF,$global:ffprobeTextBox,$btnFF,$lblTol,$global:toleranceTextBox
))

# Add status bar and progress bar
$mainForm.Controls.Add($global:progressBar)
$mainForm.Controls.Add($global:statusBar)

# Context menu for deleting files
$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
$deleteItem = $contextMenu.Items.Add("Delete")
$deleteItem.Add_Click({
    $lv = $contextMenu.SourceControl
    if ($lv.SelectedItems.Count -gt 0) {
        $sel = $lv.SelectedItems[0].Text
        $entry = $global:mediaData | Where-Object FileName -eq $sel
        if ($entry) {
            $res = [Windows.Forms.MessageBox]::Show("Delete $($entry.FullName)?","Confirm Delete",
                [Windows.Forms.MessageBoxButtons]::YesNo,[Windows.Forms.MessageBoxIcon]::Warning)
            if ($res -eq [Windows.Forms.DialogResult]::Yes) {
                try {
                    Remove-Item -LiteralPath $entry.FullName -Force
                    $lv.Items.Remove($lv.SelectedItems[0])
                    $global:mediaData = $global:mediaData | Where-Object { $_.FullName -ne $entry.FullName }
                    Write-Status "Deleted $($entry.FileName)" ([Drawing.Color]::Green)
                } catch {
                    Write-Status "Delete failed: $_" ([Drawing.Color]::Red)
                }
            }
        }
    }
})
$global:listView.ContextMenuStrip = $contextMenu

# Sorting logic
class ListViewItemComparer : System.Collections.IComparer {
    [int]$col
    [string]$order
    ListViewItemComparer([int]$c, [string]$o) { $this.col=$c; $this.order=$o }
    [int]Compare($x, $y) {
        $sx = $x.SubItems[$this.col].Text
        $sy = $y.SubItems[$this.col].Text
        $nx = 0; $ny = 0
        $isNumX = [double]::TryParse($sx, [ref]$nx)
        $isNumY = [double]::TryParse($sy, [ref]$ny)
        if ($isNumX -and $isNumY) { $cmp = $nx.CompareTo($ny) }
        else { $cmp = [string]::Compare($sx, $sy, $true) }
        if ($this.order -eq 'Descending') { return -$cmp }
        return $cmp
    }
}
$global:listSort = @{ Column = -1; Order = 'Ascending' }
$global:listView.Add_ColumnClick({
    param($sender,$e)
    if ($e.Column -eq $global:listSort.Column) {
        $global:listSort.Order = if ($global:listSort.Order -eq 'Ascending') {'Descending'} else {'Ascending'}
    } else {
        $global:listSort.Column = $e.Column
        $global:listSort.Order = 'Ascending'
    }
    $sender.ListViewItemSorter = [ListViewItemComparer]::new($global:listSort.Column, $global:listSort.Order)
    $sender.Sort()
})

# Double-click to open file
$global:listView.Add_DoubleClick({
    if ($global:listView.SelectedItems.Count -gt 0) {
        $sel  = $global:listView.SelectedItems[0].Text
        $path = ($global:mediaData | Where-Object FileName -eq $sel).FullName
        if ($path) { Start-Process $path }
    }
})

# Save config on exit
$mainForm.Add_FormClosing({
    $newCfg = @{
        FFprobePath      = $global:ffprobeTextBox.Text
        LastMediaFolder  = $global:folderTextBox.Text
        ScriptFolder     = $scriptDir
        ToleranceSeconds = if ([string]::IsNullOrWhiteSpace($global:toleranceTextBox.Text)) { 0 } else { $global:toleranceTextBox.Text }
        WindowHeight     = $mainForm.Height
        WindowWidth      = $mainForm.Width
    }

    # Overwrite config file with sorted keys
    $sorted = $newCfg.GetEnumerator() | Sort-Object Name
    $sorted | ForEach-Object {
        "$($_.Key)=$($_.Value)" | Out-File $global:cfgPath -Encoding UTF8 -Append:$false
    }
})

# Launch the GUI
[System.Windows.Forms.Application]::Run($mainForm)