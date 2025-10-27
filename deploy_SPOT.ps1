<# Deploy-SPOT.ps1
Creates the SPOT app folder in Documents, writes SPOT.ps1, and drops a desktop launcher.
#>

$docsPath    = [Environment]::GetFolderPath('MyDocuments')
$desktopPath = [Environment]::GetFolderPath('Desktop')
$spotDir     = Join-Path $docsPath 'SQL Performance Observation Tool'
$spotPath    = Join-Path $spotDir 'SPOT.ps1'
$launcher    = Join-Path $desktopPath 'Launch-SPOT.ps1'

# Ensure target folder exists
New-Item -ItemType Directory -Path $spotDir -Force | Out-Null

# -----------------------------
# Write SPOT.ps1
# -----------------------------
$spotContent = @'
#Requires -Version 5.1
<#
.SYNOPSIS
SQL Performance Observation Tool (SPOT)

.DESCRIPTION
A dashboard for collecting and exploring SQL Server performance snapshots
into utility.SPOT tables via SPOT.CaptureSnapshot.

Features:
- Captures live samples (WhoIsActive, waits, blocking, AG health, etc) on an interval.
- Saves those samples in SQL (utility.SPOT.* tables).
- Lets you browse a chosen snapshot: per-sample view, grid per category.
- Dark theme + consistent styling.
- Sample Timeline slider with per-sample navigation.
- HealthChecks de-duplicated per sample.
- Blocking info called out directly in the WhoIsActive grid.
- Execution plan viewer on demand per row.
- Session Manager: persists multiple connection profiles to disk (Documents\SPOT\Sessions.xml),
  supports create/update/delete, and auto-loads + tests last used session.

.AUTHOR
    Jake Morgan
    Blackcat Data Solutions
    October 2025
#>

Add-Type -AssemblyName System.Windows.Forms
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    [System.Windows.Forms.MessageBox]::Show(
        "Restarting SPOT in STA mode for the GUI...",
        "SPOT",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null
    Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -STA -File `"$PSCommandPath`"" -WindowStyle Hidden
    exit
}
if ($PSVersionTable.PSVersion -lt [Version]"5.1") {
    [System.Windows.Forms.MessageBox]::Show(
        "PowerShell 5.1 or higher is required. Current: $($PSVersionTable.PSVersion)",
        "SPOT prerequisite",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
    exit 1
}

Add-Type -AssemblyName System.Data
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms

# =======================
# SESSION STORAGE SETUP
# =======================
$script:SpotRootFolder   = Join-Path $env:USERPROFILE "Documents\SPOT"
$script:SessionsFilePath = Join-Path $script:SpotRootFolder "Sessions.xml"

# We will *always* hold sessions as a mutable ArrayList
$script:Sessions = New-Object System.Collections.ArrayList
$script:CurrentSessionName = $null  # track which session is "current" in UI

function Ensure-SessionsStorage {
    if (-not (Test-Path $script:SpotRootFolder)) {
        New-Item -ItemType Directory -Path $script:SpotRootFolder -Force | Out-Null
    }
    if (-not (Test-Path $script:SessionsFilePath)) {
        $blank = [pscustomobject]@{
            LastSession = ""
            Sessions    = @()
        }
        $blank | Export-Clixml -Path $script:SessionsFilePath
    }
}

function Convert-ToArrayList {
    param($items)

    $list = New-Object System.Collections.ArrayList
    if ($null -eq $items) { return $list }

    if ($items -is [System.Collections.IEnumerable] -and
        -not ($items -is [string])) {

        foreach ($i in $items) { [void]$list.Add($i) }
    }
    else {
        # single object case
        [void]$list.Add($items)
    }

    return $list
}

function Load-AllSessions {
    Ensure-SessionsStorage
    try {
        $data = Import-Clixml -Path $script:SessionsFilePath

        # rebuild ArrayList of sessions
        $script:Sessions = New-Object System.Collections.ArrayList
        if ($data.Sessions) {
            $tmpList = Convert-ToArrayList $data.Sessions
            foreach ($s in $tmpList) {
                # normalise
                $null = $script:Sessions.Add(
                    [pscustomobject]@{
                        Name     = $s.Name
                        Server   = $s.Server
                        Username = $s.Username
                        Password = $s.Password
                    }
                )
            }
        }

        $script:CurrentSessionName = $data.LastSession
    }
    catch {
        # corrupt / missing etc.
        $script:Sessions = New-Object System.Collections.ArrayList
        $script:CurrentSessionName = $null
    }
}

function Save-AllSessions {
    Ensure-SessionsStorage

    # serialisable plain array rather than ArrayList
    $plainSessions = @()
    foreach ($s in $script:Sessions) {
        $plainSessions += [pscustomobject]@{
            Name     = $s.Name
            Server   = $s.Server
            Username = $s.Username
            Password = $s.Password
        }
    }

    $outObj = [pscustomobject]@{
        LastSession = $script:CurrentSessionName
        Sessions    = $plainSessions
    }
    $outObj | Export-Clixml -Path $script:SessionsFilePath
}

function Get-SessionByName {
    param([string]$Name)
    foreach ($sess in $script:Sessions) {
        if ($sess.Name -eq $Name) { return $sess }
    }
    return $null
}

function Ensure-SessionsArrayList {
    # if someone somehow replaced it with a PSCustomObject or $null,
    # re-wrap it as ArrayList
    if (-not ($script:Sessions -is [System.Collections.ArrayList])) {
        $script:Sessions = Convert-ToArrayList $script:Sessions
    }
}

function Upsert-Session {
    param(
        [string]$Name,
        [string]$Server,
        [string]$Username,
        [securestring]$Password
    )

    Ensure-SessionsArrayList

    $existing = Get-SessionByName -Name $Name
    if ($existing) {
        $existing.Server   = $Server
        $existing.Username = $Username
        $existing.Password = $Password
    } else {
        $null = $script:Sessions.Add(
            [pscustomobject]@{
                Name     = $Name
                Server   = $Server
                Username = $Username
                Password = $Password
            }
        )
    }

    $script:CurrentSessionName = $Name
    Save-AllSessions
}

function Remove-SessionByName {
    param([string]$Name)

    if ([string]::IsNullOrWhiteSpace($Name)) { return }

    Ensure-SessionsArrayList

    # rebuild without the deleted one
    $newList = New-Object System.Collections.ArrayList
    foreach ($sess in $script:Sessions) {
        if ($sess.Name -ne $Name) {
            $null = $newList.Add($sess)
        }
    }
    $script:Sessions = $newList

    if ($script:CurrentSessionName -eq $Name) {
        $script:CurrentSessionName = $null
    }

    Save-AllSessions
}

function Set-UiFromSession {
    param([pscustomobject]$session)

    if (-not $session) { return }

    $serverTextBox.Text    = $session.Server
    $usernameTextBox.Text  = $session.Username
    $passwordTextBox.Tag   = $session.Password
    $passwordTextBox.Text  = '••••••••••'

    $connectionStatusLabel.Text = "Status: Not Connected"
    $connectionStatusLabel.ForeColor = $fgSecondary
    $captureButton.Enabled = $false
}

function Load-LastSessionToUi {
    if ([string]::IsNullOrWhiteSpace($script:CurrentSessionName)) { return }
    $sess = Get-SessionByName -Name $script:CurrentSessionName
    if ($sess) {
        Log-Message "Loaded last session '$($sess.Name)'."
        Set-UiFromSession -session $sess
        $currentSessionValue.Text = $sess.Name
        if ($sessionComboBox.Visible) {
            $sessionComboBox.SelectedItem = $sess.Name
        }
        Invoke-ConnectionTest
    }
}

function Load-SavedConnection {
    # legacy import from $PSScriptRoot\config.xml
    $legacyFile = Join-Path $PSScriptRoot 'config.xml'
    if (-not (Test-Path $legacyFile)) { return }

    try {
        $credential = Import-CliXml -Path $legacyFile
        $user, $server = $credential.UserName.Split('@')

        Log-Message "Importing legacy config.xml as session 'Legacy'."
        Upsert-Session -Name "Legacy" `
            -Server $server `
            -Username $user `
            -Password $credential.Password
    }
    catch {
        Log-Message "Could not import legacy config.xml. File may be corrupt."
    }
}

# =======================
# THEME
# =======================
$bgMain        = [System.Drawing.ColorTranslator]::FromHtml("#1e1e1e")
$bgPanel       = [System.Drawing.ColorTranslator]::FromHtml("#252526")
$bgPanelBorder = [System.Drawing.ColorTranslator]::FromHtml("#3f3f46")

$fgPrimary     = [System.Drawing.ColorTranslator]::FromHtml("#d4d4d4")
$fgSecondary   = [System.Drawing.ColorTranslator]::FromHtml("#9ca3af")
$fgAccent      = [System.Drawing.ColorTranslator]::FromHtml("#ffffff")
$fgWarn        = [System.Drawing.ColorTranslator]::FromHtml("#ffd700")

$accentBlue    = [System.Drawing.ColorTranslator]::FromHtml("#0e639c")
$btnGray       = [System.Drawing.ColorTranslator]::FromHtml("#3a3d41")

$gridHeaderBg  = [System.Drawing.ColorTranslator]::FromHtml("#2d2d30")
$gridHeaderFg  = $fgAccent
$gridRowBg     = [System.Drawing.ColorTranslator]::FromHtml("#1e1e1e")
$gridRowSelBg  = [System.Drawing.ColorTranslator]::FromHtml("#094771")
$gridRowSelFg  = $fgAccent
$gridLines     = [System.Drawing.ColorTranslator]::FromHtml("#3f3f46")
$blockingRed   = [System.Drawing.ColorTranslator]::FromHtml("#b00020")

$consoleBack   = [System.Drawing.ColorTranslator]::FromHtml("#1a1a1a")
$consoleFore   = [System.Drawing.ColorTranslator]::FromHtml("#d4d4d4")

$timelineActiveText = $fgWarn

$fontRegular = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
$fontBold    = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

# runtime global state
$script:loadedSnapshotData   = $null
$script:sortState            = @{}
$script:dataGrids            = @{}
$script:timelineLabels       = @()
$script:isUpdatingTimeline   = $false

# =======================
# MAIN FORM
# =======================
$mainForm = New-Object System.Windows.Forms.Form
$mainForm.Text = "SQL Performance Observation Tool"
$mainForm.Size = New-Object System.Drawing.Size(1200, 800)
$mainForm.MinimumSize = New-Object System.Drawing.Size(800, 600)
$mainForm.StartPosition = "CenterScreen"
$mainForm.FormBorderStyle = "Sizable"
$mainForm.BackColor = $bgMain
$mainForm.ForeColor = $fgPrimary
$mainForm.Font = $fontRegular

# =======================
# DARK GROUPBOX CREATOR
# =======================
function New-DarkGroupBox {
    param(
        [string]$text,
        [System.Drawing.Point]$location,
        [System.Drawing.Size]$size,
        [string]$anchor = "Top, Left"
    )
    $gb = New-Object System.Windows.Forms.GroupBox
    $gb.Text = $text
    $gb.Location = $location
    $gb.Size = $size
    $gb.Anchor = $anchor
    $gb.ForeColor = $fgPrimary
    $gb.BackColor = $bgPanel
    return $gb
}

# =======================
# CONNECTION DETAILS BOX
# =======================
$connectionGroupBox = New-DarkGroupBox -text "Connection Details" `
    -location ([System.Drawing.Point]::new(20,20)) `
    -size     ([System.Drawing.Size]::new(300,270)) `
    -anchor   "Top, Left"
$mainForm.Controls.Add($connectionGroupBox)

$sessionSelectLabel = New-Object System.Windows.Forms.Label
$sessionSelectLabel.Text = "Session:"
$sessionSelectLabel.Location = New-Object System.Drawing.Point(10, 30)
$sessionSelectLabel.AutoSize = $true
$sessionSelectLabel.ForeColor = $fgPrimary

$sessionComboBox = New-Object System.Windows.Forms.ComboBox
$sessionComboBox.Location = New-Object System.Drawing.Point(120, 27)
$sessionComboBox.Size = New-Object System.Drawing.Size(160,20)
$sessionComboBox.DropDownStyle = 'DropDownList'
$sessionComboBox.BackColor = $gridRowBg
$sessionComboBox.ForeColor = $fgPrimary
$sessionComboBox.FlatStyle = 'Flat'
$sessionComboBox.Visible = $true       # keep visible even if empty now
$sessionSelectLabel.Visible = $true    # ditto

$serverLabel = New-Object System.Windows.Forms.Label
$serverLabel.Text = "Server/Instance:"
$serverLabel.Location = New-Object System.Drawing.Point(10, 60)
$serverLabel.AutoSize = $true
$serverLabel.ForeColor = $fgPrimary

$serverTextBox = New-Object System.Windows.Forms.TextBox
$serverTextBox.Location = New-Object System.Drawing.Point(120, 57)
$serverTextBox.Size = New-Object System.Drawing.Size(160, 20)
$serverTextBox.BackColor = $gridRowBg
$serverTextBox.ForeColor = $fgPrimary
$serverTextBox.BorderStyle = 'FixedSingle'

$usernameLabel = New-Object System.Windows.Forms.Label
$usernameLabel.Text = "Username:"
$usernameLabel.Location = New-Object System.Drawing.Point(10, 90)
$usernameLabel.AutoSize = $true
$usernameLabel.ForeColor = $fgPrimary

$usernameTextBox = New-Object System.Windows.Forms.TextBox
$usernameTextBox.Location = New-Object System.Drawing.Point(120, 87)
$usernameTextBox.Size = New-Object System.Drawing.Size(160, 20)
$usernameTextBox.BackColor = $gridRowBg
$usernameTextBox.ForeColor = $fgPrimary
$usernameTextBox.BorderStyle = 'FixedSingle'

$passwordLabel = New-Object System.Windows.Forms.Label
$passwordLabel.Text = "Password:"
$passwordLabel.Location = New-Object System.Drawing.Point(10, 120)
$passwordLabel.AutoSize = $true
$passwordLabel.ForeColor = $fgPrimary

$passwordTextBox = New-Object System.Windows.Forms.TextBox
$passwordTextBox.Location = New-Object System.Drawing.Point(120, 117)
$passwordTextBox.Size = New-Object System.Drawing.Size(160, 20)
$passwordTextBox.UseSystemPasswordChar = $true
$passwordTextBox.BackColor = $gridRowBg
$passwordTextBox.ForeColor = $fgPrimary
$passwordTextBox.BorderStyle = 'FixedSingle'

$testConnectionButton = New-Object System.Windows.Forms.Button
$testConnectionButton.Text = "Test"
$testConnectionButton.Location = New-Object System.Drawing.Point(10, 160)
$testConnectionButton.Size = New-Object System.Drawing.Size(80, 30)
$testConnectionButton.BackColor = $btnGray
$testConnectionButton.ForeColor = $fgPrimary
$testConnectionButton.FlatStyle = "Flat"

$saveConnectionButton = New-Object System.Windows.Forms.Button
$saveConnectionButton.Text = "Save"
$saveConnectionButton.Location = New-Object System.Drawing.Point(100, 160)
$saveConnectionButton.Size = New-Object System.Drawing.Size(80, 30)
$saveConnectionButton.BackColor = $btnGray
$saveConnectionButton.ForeColor = $fgPrimary
$saveConnectionButton.FlatStyle = "Flat"

$deleteSessionButton = New-Object System.Windows.Forms.Button
$deleteSessionButton.Text = "🗑 Delete"
$deleteSessionButton.Location = New-Object System.Drawing.Point(190, 160)
$deleteSessionButton.Size = New-Object System.Drawing.Size(80, 30)
$deleteSessionButton.BackColor = $btnGray
$deleteSessionButton.ForeColor = $blockingRed # $fgPrimary
$deleteSessionButton.FlatStyle = "Flat"
$deleteSessionButton.Font = New-Object System.Drawing.Font("Segoe UI Emoji", 9, [System.Drawing.FontStyle]::Regular)

$connectionStatusLabel = New-Object System.Windows.Forms.Label
$connectionStatusLabel.Text = "Status: Not Connected"
$connectionStatusLabel.Location = New-Object System.Drawing.Point(10, 200)
$connectionStatusLabel.AutoSize = $true
$connectionStatusLabel.ForeColor = $fgSecondary


$currentSessionLabel = New-Object System.Windows.Forms.Label
$currentSessionLabel.Text = "Current Session:"
$currentSessionLabel.Location = New-Object System.Drawing.Point(10, 225)
$currentSessionLabel.AutoSize = $true
$currentSessionLabel.ForeColor = $fgPrimary

$currentSessionValue = New-Object System.Windows.Forms.Label
$currentSessionValue.Text = "(none)"
$currentSessionValue.Location = New-Object System.Drawing.Point(120, 225)
$currentSessionValue.AutoSize = $true
$currentSessionValue.ForeColor = $fgPrimary

$utilityDbLabel = New-Object System.Windows.Forms.Label
#$utilityDbLabel.Text = "Target DB:"
$utilityDbLabel.Location = New-Object System.Drawing.Point(10, 245)
$utilityDbLabel.AutoSize = $true
$utilityDbLabel.ForeColor = $fgPrimary

$utilityDbValue = New-Object System.Windows.Forms.Label
#$utilityDbValue.Text = "utility"
$utilityDbValue.Location = New-Object System.Drawing.Point(120, 245)
$utilityDbValue.AutoSize = $true
$utilityDbValue.ForeColor = $fgPrimary

$connectionGroupBox.Controls.AddRange(@(
    $sessionSelectLabel, $sessionComboBox,
    $serverLabel, $serverTextBox,
    $usernameLabel, $usernameTextBox,
    $passwordLabel, $passwordTextBox,
    $testConnectionButton, $saveConnectionButton, $deleteSessionButton,
    $connectionStatusLabel, $currentSessionLabel, $currentSessionValue,
    $utilityDbLabel, $utilityDbValue
))

# =======================
# SNAPSHOT CONFIG BOX
# =======================
$snapshotGroupBox = New-DarkGroupBox -text "Snapshot Configuration" `
    -location ([System.Drawing.Point]::new(20,300)) `
    -size     ([System.Drawing.Size]::new(300,230)) `
    -anchor   "Top, Left"
$mainForm.Controls.Add($snapshotGroupBox)

$sampleRateLabel = New-Object System.Windows.Forms.Label
$sampleRateLabel.Text = "Sample Interval (s):"
$sampleRateLabel.Location = New-Object System.Drawing.Point(10, 30)
$sampleRateLabel.AutoSize = $true
$sampleRateLabel.ForeColor = $fgPrimary

$sampleRateValueLabel = New-Object System.Windows.Forms.Label
$sampleRateValueLabel.Location = New-Object System.Drawing.Point(260, 30)
$sampleRateValueLabel.AutoSize = $true
$sampleRateValueLabel.ForeColor = $fgAccent

$sampleRateSlider = New-Object System.Windows.Forms.TrackBar
$sampleRateSlider.Location = New-Object System.Drawing.Point(120, 25)
$sampleRateSlider.Size = New-Object System.Drawing.Size(140, 45)
$sampleRateSlider.Minimum = 5
$sampleRateSlider.Maximum = 12
$sampleRateSlider.Value = 5
$sampleRateSlider.TickStyle = 'BottomRight'
$sampleRateSlider.BackColor = $bgPanel
$sampleRateValueLabel.Text = $sampleRateSlider.Value

$numSamplesLabel = New-Object System.Windows.Forms.Label
$numSamplesLabel.Text = "Number of Samples:"
$numSamplesLabel.Location = New-Object System.Drawing.Point(10, 80)
$numSamplesLabel.AutoSize = $true
$numSamplesLabel.ForeColor = $fgPrimary

$numSamplesValueLabel = New-Object System.Windows.Forms.Label
$numSamplesValueLabel.Location = New-Object System.Drawing.Point(260, 80)
$numSamplesValueLabel.AutoSize = $true
$numSamplesValueLabel.ForeColor = $fgAccent

$numSamplesSlider = New-Object System.Windows.Forms.TrackBar
$numSamplesSlider.Location = New-Object System.Drawing.Point(120, 75)
$numSamplesSlider.Size = New-Object System.Drawing.Size(140, 45)
$numSamplesSlider.Minimum = 1
$numSamplesSlider.Maximum = 20
$numSamplesSlider.Value = 3
$numSamplesSlider.TickStyle = 'BottomRight'
$numSamplesSlider.BackColor = $bgPanel
$numSamplesValueLabel.Text = $numSamplesSlider.Value

$captureButton = New-Object System.Windows.Forms.Button
$captureButton.Text = "Start Capture"
$captureButton.Location = New-Object System.Drawing.Point(10, 120)
$captureButton.Size = New-Object System.Drawing.Size(280, 40)
$captureButton.BackColor = $accentBlue
$captureButton.ForeColor = $fgAccent
$captureButton.FlatStyle = "Flat"
$captureButton.Enabled = $false

$loadSnapshotButton = New-Object System.Windows.Forms.Button
$loadSnapshotButton.Text = "Load Snapshot from DB"
$loadSnapshotButton.Location = New-Object System.Drawing.Point(10, 165)
$loadSnapshotButton.Size = New-Object System.Drawing.Size(280, 25)
$loadSnapshotButton.BackColor = $btnGray
$loadSnapshotButton.ForeColor = $fgPrimary
$loadSnapshotButton.FlatStyle = "Flat"

$snapshotGroupBox.Controls.AddRange(@(
    $sampleRateLabel, $sampleRateValueLabel, $sampleRateSlider,
    $numSamplesLabel, $numSamplesValueLabel, $numSamplesSlider,
    $captureButton, $loadSnapshotButton
))

# =======================
# CONSOLE LOG
# =======================
$consoleGroupBox = New-DarkGroupBox -text "Console Log" `
    -location ([System.Drawing.Point]::new(340,20)) `
    -size     ([System.Drawing.Size]::new(820,140)) `
    -anchor   "Top, Left, Right"
$mainForm.Controls.Add($consoleGroupBox)

$consoleLogTextBox = New-Object System.Windows.Forms.TextBox
$consoleLogTextBox.Location = New-Object System.Drawing.Point(10, 20)
$consoleLogTextBox.Size = New-Object System.Drawing.Size(800, 110)
$consoleLogTextBox.Multiline = $true
$consoleLogTextBox.ScrollBars = "Vertical"
$consoleLogTextBox.ReadOnly = $true
$consoleLogTextBox.BackColor = $consoleBack
$consoleLogTextBox.ForeColor = $consoleFore
$consoleLogTextBox.BorderStyle = 'FixedSingle'
$consoleLogTextBox.Anchor = "Top, Bottom, Left, Right"
$consoleGroupBox.Controls.Add($consoleLogTextBox)

# =======================
# SAMPLE TIMELINE
# =======================
$timelineGroupBox = New-DarkGroupBox -text "Sample Timeline:" `
    -location ([System.Drawing.Point]::new(340,170)) `
    -size     ([System.Drawing.Size]::new(820,150)) `
    -anchor   "Top, Left, Right"
$timelineGroupBox.Visible = $false
$mainForm.Controls.Add($timelineGroupBox)

$timelineTrackBar = New-Object System.Windows.Forms.TrackBar
$timelineTrackBar.Location     = New-Object System.Drawing.Point(10, 40)
$timelineTrackBar.Size         = New-Object System.Drawing.Size(800, 45)
$timelineTrackBar.Minimum      = 0
$timelineTrackBar.Maximum      = 0
$timelineTrackBar.TickFrequency = 1
$timelineTrackBar.SmallChange   = 1
$timelineTrackBar.LargeChange   = 1
$timelineTrackBar.TickStyle     = 'BottomRight'
$timelineTrackBar.BackColor     = $bgPanel
$timelineTrackBar.Enabled       = $false
$timelineTrackBar.AutoSize      = $false
$timelineTrackBar.Height        = 45
$timelineTrackBar.Anchor        = "Top, Left, Right"
$timelineGroupBox.Controls.Add($timelineTrackBar)

# =======================
# RESULTS TABCONTROL
# =======================
$resultsTabControl = New-Object System.Windows.Forms.TabControl
$resultsTabControl.Location = New-Object System.Drawing.Point(340,330)
$resultsTabControl.Size = New-Object System.Drawing.Size(820, 430)
$resultsTabControl.Visible = $false
$resultsTabControl.Anchor = "Top, Bottom, Left, Right"
$resultsTabControl.BackColor = $bgPanel
$resultsTabControl.ForeColor = $fgPrimary
$mainForm.Controls.Add($resultsTabControl)

# =======================
# CONTEXT MENU FOR GRIDS
# =======================
$gridContextMenu = New-Object System.Windows.Forms.ContextMenuStrip
$menuCopyRow     = New-Object System.Windows.Forms.ToolStripMenuItem("Copy Row to Clipboard")
$menuViewPlan    = New-Object System.Windows.Forms.ToolStripMenuItem("View Execution Plan")
[void]$gridContextMenu.Items.Add($menuCopyRow)
[void]$gridContextMenu.Items.Add($menuViewPlan)

# =======================
# HELPERS
# =======================
function Log-Message {
    param([string]$Message)
    $consoleLogTextBox.AppendText("$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $Message`r`n")
    $consoleLogTextBox.Update()
}

function Get-PasswordFromTextBox {
    if ($passwordTextBox.Tag -is [System.Security.SecureString]) {
        $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($passwordTextBox.Tag)
        $password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
        return $password
    } else {
        return $passwordTextBox.Text
    }
}

function Apply-DarkModeToGrid {
    param([System.Windows.Forms.DataGridView]$grid)

    $grid.EnableHeadersVisualStyles = $false
    $grid.ColumnHeadersBorderStyle = 'None'
    $grid.RowHeadersVisible = $false

    $headerStyle = New-Object System.Windows.Forms.DataGridViewCellStyle
    $headerStyle.BackColor = $gridHeaderBg
    $headerStyle.ForeColor = $gridHeaderFg
    $headerStyle.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $grid.ColumnHeadersDefaultCellStyle = $headerStyle

    $cellStyle = New-Object System.Windows.Forms.DataGridViewCellStyle
    $cellStyle.BackColor = $gridRowBg
    $cellStyle.ForeColor = $fgPrimary
    $cellStyle.SelectionBackColor = $gridRowSelBg
    $cellStyle.SelectionForeColor = $gridRowSelFg
    $grid.DefaultCellStyle = $cellStyle

    $grid.BackgroundColor = $gridRowBg
    $grid.GridColor = $gridLines
    $grid.BorderStyle = 'None'

    $grid.ContextMenuStrip = $gridContextMenu
    $grid.ReadOnly = $true
    $grid.SelectionMode = "FullRowSelect"
    $grid.MultiSelect = $false
    $grid.Font = $fontRegular
}

function Highlight-WhoIsActiveRows {
    param([System.Windows.Forms.DataGridView]$grid)

    $redStyle = New-Object System.Windows.Forms.DataGridViewCellStyle
    $redStyle.BackColor = $blockingRed
    $redStyle.ForeColor = $fgAccent
    $redStyle.SelectionBackColor = $blockingRed
    $redStyle.SelectionForeColor = $fgAccent

    $longWaitThresholdMs = 1000

    foreach ($row in $grid.Rows) {
        $data = $row.DataBoundItem
        if ($null -eq $data) { continue }

        $isProblem = $false

        if ($data.'Blocking Info' -and $data.'Blocking Info'.Trim() -ne "") {
            $isProblem = $true
        }

        if (-not $isProblem -and $data.wait_info -and $data.wait_info -notmatch 'WAITFOR') {
            if ($data.wait_info -match '\((\d+)ms\)') {
                $waitMs = [int]$Matches[1]
                if ($waitMs -gt $longWaitThresholdMs) {
                    $isProblem = $true
                }
            }
        }

        if ($isProblem) {
            $row.DefaultCellStyle = $redStyle
        }
    }
}

function Format-Datetime2Literal {
    param([datetime]$dt)
    "'{0}'" -f $dt.ToString('yyyy-MM-ddTHH:mm:ss.fff')
}

function Test-HasValue {
    param($v)
    if ($null -eq $v) { return $false }
    if ($v -is [System.DBNull]) { return $false }
    if ($v -is [int] -and $v -eq 0) { return $false }
    if ($v -is [string]) {
        if ($v.Trim() -eq "") { return $false }
        if ($v -eq "0") { return $false }
    }
    return $true
}

# =======================
# TAB + GRID BUILD
# =======================
function On-GridDataBindingComplete {
    param($sender, $e)
    $grid = $sender
    foreach ($column in $grid.Columns) {
        switch ($column.Name) {
            'sql_text'          { $column.AutoSizeMode = 'None'; $column.Width = 260; break }
            'waiting_query'     { $column.AutoSizeMode = 'None'; $column.Width = 260; break }
            'blocking_query'    { $column.AutoSizeMode = 'None'; $column.Width = 260; break }
            'Purpose'           { $column.AutoSizeMode = 'None'; $column.Width = 380; break }
            'Check Description' { $column.AutoSizeMode = 'None'; $column.Width = 260; break }
            default             { $column.AutoSizeMode = 'AllCells' }
        }
    }
    if ($grid.Columns["query_plan"]) {
        $grid.Columns["query_plan"].Visible = $false
    }
}

$On_ColumnHeaderClick = {
    param($sender, $e)
    $grid = $sender

    foreach ($column in $grid.Columns) {
        if ($column.Index -ne $e.ColumnIndex) {
            $column.HeaderCell.SortGlyphDirection = 'None'
        }
    }

    $sortProperty = $grid.Columns[$e.ColumnIndex].DataPropertyName
    if (-not $sortProperty) { return }

    $currentDirection = if (
        $script:sortState[$grid.Name] -and
        $script:sortState[$grid.Name].Property -eq $sortProperty
    ) {
        $script:sortState[$grid.Name].Direction
    } else {
        'Ascending'
    }

    $newDirection = if ($currentDirection -eq 'Ascending') { 'Descending' } else { 'Ascending' }

    $data = $grid.DataSource | Sort-Object -Property $sortProperty -Descending:($newDirection -eq 'Descending')
    $grid.DataSource = $null
    $grid.DataSource = [System.Collections.ArrayList]$data

    $grid.Columns[$e.ColumnIndex].HeaderCell.SortGlyphDirection = $newDirection
    $script:sortState[$grid.Name] = @{
        Property  = $sortProperty
        Direction = $newDirection
    }
}

function Build-UiTabsOnce {
    $resultsTabControl.TabPages.Clear()
    $script:dataGrids.Clear()

    $dataTypes = @("WhoIsActive", "HealthChecks", "AGHealth", "WaitStats", "Blocking")
    foreach ($type in $dataTypes) {
        $tabPage = New-Object System.Windows.Forms.TabPage
        $tabPage.Text = $type
        $tabPage.BackColor = $bgPanel
        $tabPage.ForeColor = $fgPrimary
        $tabPage.Font = $fontRegular

        $grid = New-Object System.Windows.Forms.DataGridView
        $grid.Name = $type
        $grid.Dock = "Fill"
        $grid.AutoSizeColumnsMode = "None"
        $grid.AllowUserToResizeColumns = $true

        Apply-DarkModeToGrid -grid $grid

        $grid.add_ColumnHeaderMouseClick($On_ColumnHeaderClick)
        $grid.add_DataBindingComplete({ param($s,$e) On-GridDataBindingComplete -sender $s -e $e })

        if ($type -eq 'WhoIsActive') {
            $grid.add_DataBindingComplete({
                param($s,$e)
                if ($this.Columns["Blocking Info"]) {
                    $this.Columns["Blocking Info"].DisplayIndex = 0
                }
                if ($this.Columns["query_plan"]) {
                    $this.Columns["query_plan"].Visible = $false
                }
                Highlight-WhoIsActiveRows -grid $this
            })
        }

        $tabPage.Controls.Add($grid)
        $resultsTabControl.TabPages.Add($tabPage)
        $script:dataGrids[$type] = $grid
    }
}

function Bind-SampleToGrids {
    param($selectedSample)

    foreach ($key in $script:dataGrids.Keys) {
        $grid = $script:dataGrids[$key]
        $data = $selectedSample.$key

        $bindableData = New-Object System.Collections.ArrayList
        if ($data) {
            if ($data -is [array]) {
                $bindableData.AddRange($data)
            } else {
                $bindableData.Add($data)
            }
        }

        if ($key -eq 'WaitStats' -and $data) {
            $bindableData = [System.Collections.ArrayList]($data | Sort-Object -Property 'wait_time_ms' -Descending)
        }

        $grid.DataSource = $null
        $grid.DataSource = $bindableData
    }
}

# =======================
# TIMELINE LABEL HANDLING
# =======================
$script:timelineLabels = @()

function Build-TimelineLabels {
    param(
        [System.Collections.IList]$samples
    )

    foreach ($lbl in $script:timelineLabels) {
        if ($lbl -and -not $lbl.IsDisposed) {
            $timelineGroupBox.Controls.Remove($lbl)
            $lbl.Dispose()
        }
    }
    $script:timelineLabels = @()

    if (-not $samples -or $samples.Count -eq 0) { return }

    $trackLeft   = $timelineTrackBar.Left
    $trackTop    = $timelineTrackBar.Top
    $trackWidth  = $timelineTrackBar.Width
    $knobFudge   = [int]16
    $labelY      = $trackTop + $timelineTrackBar.Height + 5

    $count = $samples.Count
    for ($i=0; $i -lt $count; $i++) {
        $sampleNum = $samples[$i].SampleNumber.ToString()

        if ($count -gt 1) {
            $pct = $i / [double]($count - 1)
        } else {
            $pct = 0.0
        }
        $xOffset = [int]( ($trackWidth - $knobFudge) * $pct )
        $labelX  = $trackLeft + $xOffset

        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text      = $sampleNum
        $lbl.AutoSize  = $true
        $lbl.ForeColor = $fgPrimary
        $lbl.BackColor = $bgPanel
        $lbl.Font      = $fontRegular
        $lbl.Cursor    = [System.Windows.Forms.Cursors]::Hand
        $lbl.Tag       = $i
        $lbl.Location  = New-Object System.Drawing.Point($labelX,$labelY)

        $lbl.Add_Click({
            param($sender,$args)
            $idx = [int]$sender.Tag
            $script:isUpdatingTimeline = $true
            $timelineTrackBar.Value = $idx
            $script:isUpdatingTimeline = $false
            Set-TimelineIndex -index $idx
        })

        $timelineGroupBox.Controls.Add($lbl)
        $script:timelineLabels += $lbl
    }
}

function Update-TimelineLabelHighlight {
    param([int]$index)

    for ($i=0; $i -lt $script:timelineLabels.Count; $i++) {
        $lbl = $script:timelineLabels[$i]
        if ($null -eq $lbl -or $lbl.IsDisposed) { continue }

        if ($i -eq $index) {
            $lbl.ForeColor = $timelineActiveText
            $lbl.Font      = $fontBold
        } else {
            $lbl.ForeColor = $fgPrimary
            $lbl.Font      = $fontRegular
        }
    }
}

function Set-TimelineIndex {
    param([int]$index)

    if ($script:loadedSnapshotData -and
        $script:loadedSnapshotData.Samples -and
        $index -ge 0 -and
        $index -lt $script:loadedSnapshotData.Samples.Count) {

        $selectedSample = $script:loadedSnapshotData.Samples[$index]
        $sampleNumber   = $selectedSample.SampleNumber
        $stampDisplay   = [datetime]$selectedSample.Timestamp | Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

        Log-Message "Loading Sample $sampleNumber ($stampDisplay)..."

        $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            Bind-SampleToGrids -selectedSample $selectedSample
            Update-TimelineLabelHighlight -index $index

            $timelineGroupBox.Text = "Sample Timeline: $stampDisplay"
        }
        finally {
            Log-Message "Loaded Sample $sampleNumber."
            $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    }
}

# =======================
# SNAPSHOT LOADING
# =======================
function Get-SqlData {
    param([string]$ConnectionString, [string]$Query)

    $connection = New-Object System.Data.SqlClient.SqlConnection($ConnectionString)
    $command    = New-Object System.Data.SqlClient.SqlCommand($Query, $connection)
    $command.CommandTimeout = 180
    $adapter  = New-Object System.Data.SqlClient.SqlDataAdapter($command)
    $dataset  = New-Object System.Data.DataSet
    try {
        $connection.Open()
        [void]$adapter.Fill($dataset)
    } catch {
        throw $_
    } finally {
        if ($connection.State -eq 'Open') { $connection.Close() }
    }

    $table = $null
    if ($dataset.Tables.Count -gt 0) {
        $table = ($dataset.Tables | Where-Object { $_.Columns.Count -gt 0 } | Select-Object -Last 1)
    }
    if (-not $table) { return @() }

    $results = New-Object System.Collections.ArrayList
    foreach ($row in $table.Rows) {
        $obj = New-Object -TypeName PSObject
        foreach ($col in $table.Columns) {
            $obj | Add-Member -MemberType NoteProperty -Name $col.ColumnName -Value $row[$col]
        }
        [void]$results.Add($obj)
    }
    return $results
}

function Load-SnapshotById {
    param([string]$ConnectionString, [datetime]$SnapshotId)

    try {
        Log-Message "Loading snapshot $($SnapshotId.ToString('yyyy-MM-dd HH:mm:ss.fff')) from DB..."
        $lit = Format-Datetime2Literal $SnapshotId

        $samples = Get-SqlData -ConnectionString $ConnectionString -Query @"
SELECT SampleId, SampleNumber, CaptureTimestamp
FROM SPOT.Samples
WHERE SnapshotId = $lit
ORDER BY SampleNumber;
"@

        if (-not $samples -or $samples.Count -eq 0) {
            throw "No samples found for SnapshotId $SnapshotId"
        }

        $script:loadedSnapshotData = @{
            Samples = New-Object System.Collections.ArrayList
        }

        Build-UiTabsOnce

        foreach ($s in $samples) {
            $sid = $s.SampleId

            $sampleObj = @{
                SampleNumber = $s.SampleNumber
                Timestamp    = $s.CaptureTimestamp
            }

            $whois = Get-SqlData -ConnectionString $ConnectionString -Query @"
SELECT [dd hh:mm:ss.mss],[session_id],[sql_text],[login_name],[wait_info],[CPU],[tempdb_allocations],[tempdb_current],
       [blocking_session_id],[reads],[writes],[physical_reads],[used_memory],[status],[open_tran_count],[percent_complete],
       [host_name],[database_name],[program_name],[start_time],[login_time],[request_id],[query_plan],[collection_time]
FROM SPOT.WhoIsActive
WHERE SnapshotId = $lit AND SampleId = $sid;
"@

            if ($whois) {
                $leadBlockers = @{}
                foreach ($row in $whois) {
                    $blkBy = $row.blocking_session_id
                    if (Test-HasValue $blkBy) {
                        $parent = $blkBy.ToString()
                        if (-not $leadBlockers.ContainsKey($parent)) {
                            $leadBlockers[$parent] = @()
                        }
                        $leadBlockers[$parent] += $row.session_id.ToString()
                    }
                }

                foreach ($row in $whois) {
                    $sidStr  = $row.session_id.ToString()
                    $blkBy   = $row.blocking_session_id
                    $info    = ""

                    $iAmBlocked = Test-HasValue $blkBy
                    $iAmLead    = $leadBlockers.ContainsKey($sidStr)

                    $blockerIsAlsoBlocked = $false
                    if ($iAmBlocked) {
                        $parentSid = $blkBy.ToString()
                        $parentRow = $whois | Where-Object { $_.session_id.ToString() -eq $parentSid } | Select-Object -First 1
                        if ($parentRow) {
                            $parentBlocker = $parentRow.blocking_session_id
                            if (Test-HasValue $parentBlocker) {
                                $blockerIsAlsoBlocked = $true
                            }
                        }
                    }

                    $iAmAlsoBlocked = $false
                    if ($iAmLead) {
                        $meParentBlocker = $row.blocking_session_id
                        if (Test-HasValue $meParentBlocker) {
                            $iAmAlsoBlocked = $true
                        }
                    }

                    if ($iAmLead) {
                        $info = "LEAD"
                        if ($iAmAlsoBlocked) {
                            $info += " (CHAIN)"
                        }
                    }
                    elseif ($iAmBlocked) {
                        $info = "BLOCKED BY $blkBy"
                        if ($blockerIsAlsoBlocked) {
                            $info += " (CHAIN)"
                        }
                    }
                    else {
                        $info = ""
                    }

                    Add-Member -InputObject $row -MemberType NoteProperty -Name 'Blocking Info' -Value $info -Force
                }
            }
            $sampleObj.WhoIsActive = $whois

            $hcRaw = Get-SqlData -ConnectionString $ConnectionString -Query @"
SELECT [Check Description],[Purpose],[Current Value]
FROM SPOT.HealthChecks
WHERE SnapshotId = $lit AND SampleId = $sid
ORDER BY [Check Description];
"@

            $hcGrouped = @()
            if ($hcRaw) {
                $hcGrouped = $hcRaw |
                    Group-Object -Property 'Check Description','Purpose' |
                    ForEach-Object {
                        $last = $_.Group[-1]
                        [pscustomobject]@{
                            'Check Description' = $last.'Check Description'
                            'Purpose'           = $last.'Purpose'
                            'Current Value'     = $last.'Current Value'
                        }
                    }
            }
            $sampleObj.HealthChecks = $hcGrouped

            $agRaw = Get-SqlData -ConnectionString $ConnectionString -Query @"
SELECT replica_server_name, ag_name, database_name, synchronization_state_desc,
       is_suspended, log_send_queue_size, redo_queue_size
FROM SPOT.AGHealth
WHERE SnapshotId = $lit AND SampleId = $sid;
"@
            if ($agRaw) {
                $agRaw = $agRaw | Select-Object replica_server_name,ag_name,database_name,
                                            synchronization_state_desc,is_suspended,
                                            log_send_queue_size,redo_queue_size -Unique
            }
            $sampleObj.AGHealth = $agRaw

            $wsRaw = Get-SqlData -ConnectionString $ConnectionString -Query @"
SELECT wait_type, waiting_tasks_count, wait_time_ms, max_wait_time_ms, signal_wait_time_ms
FROM SPOT.WaitStats
WHERE SnapshotId = $lit AND SampleId = $sid;
"@
            if ($wsRaw) {
                $wsRaw = $wsRaw | Select-Object wait_type,waiting_tasks_count,wait_time_ms,
                                          max_wait_time_ms,signal_wait_time_ms -Unique
            }
            $sampleObj.WaitStats = $wsRaw

            $blkRaw = Get-SqlData -ConnectionString $ConnectionString -Query @"
SELECT waiting_session_id, blocking_session_id, wait_duration_ms, waiting_query, blocking_query
FROM SPOT.Blocking
WHERE SnapshotId = $lit AND SampleId = $sid;
"@
            if ($blkRaw) {
                $blkRaw = $blkRaw | Select-Object waiting_session_id,blocking_session_id,
                                        wait_duration_ms,waiting_query,blocking_query -Unique
            }
            $sampleObj.Blocking = $blkRaw

            $null = $script:loadedSnapshotData.Samples.Add($sampleObj)
        }

        $timelineGroupBox.Visible = $true
        $resultsTabControl.Visible = $true

        $count = $script:loadedSnapshotData.Samples.Count
        $timelineTrackBar.Minimum = 0
        $timelineTrackBar.Maximum = [math]::Max(0, $count - 1)
        $timelineTrackBar.TickFrequency = 1
        $timelineTrackBar.SmallChange  = 1
        $timelineTrackBar.LargeChange  = 1
        $timelineTrackBar.Enabled = $true

        Build-TimelineLabels -samples $script:loadedSnapshotData.Samples

        $script:isUpdatingTimeline = $true
        $timelineTrackBar.Value = 0
        $script:isUpdatingTimeline = $false
        Set-TimelineIndex -index 0

        Log-Message "Snapshot loaded successfully."
    }
    catch {
        Log-Message "ERROR: Failed to load snapshot from DB. $_"
    }
}

# =======================
# CONTEXT MENU HANDLERS
# =======================
function Get-SelectedRowObjectFromGrid {
    param([System.Windows.Forms.DataGridView]$grid)
    if ($grid.SelectedRows.Count -gt 0) {
        return $grid.SelectedRows[0].DataBoundItem
    } elseif ($grid.CurrentRow -and $grid.CurrentRow.Index -ge 0) {
        return $grid.CurrentRow.DataBoundItem
    }
    return $null
}

function Copy-RowToClipboard {
    param([System.Windows.Forms.DataGridView]$grid)

    $rowObj = Get-SelectedRowObjectFromGrid -grid $grid
    if (-not $rowObj) { return }

    $sbHeader = New-Object System.Text.StringBuilder
    $sbValues = New-Object System.Text.StringBuilder

    foreach ($col in $grid.Columns) {
        if ($col.Visible) {
            [void]$sbHeader.Append($col.HeaderText)
            [void]$sbHeader.Append("`t")
            $val = ""
            try { $val = ($rowObj."$($col.DataPropertyName)") } catch {}
            [void]$sbValues.Append($val)
            [void]$sbValues.Append("`t")
        }
    }

    $text = ($sbHeader.ToString().Trim("`t") + "`r`n" + $sbValues.ToString().Trim("`t"))
    [System.Windows.Forms.Clipboard]::SetText($text)
    Log-Message "Row copied to clipboard."
}

function Show-ExecutionPlan {
    param([System.Windows.Forms.DataGridView]$grid)

    $rowObj = Get-SelectedRowObjectFromGrid -grid $grid
    if (-not $rowObj) { return }

    $planXml = $null
    try { $planXml = $rowObj.query_plan } catch { $planXml = $null }

    if (-not $planXml -or [string]::IsNullOrWhiteSpace([string]$planXml)) {
        [System.Windows.Forms.MessageBox]::Show(
            "No execution plan available for this row.",
            "SPOT",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
        return
    }

    $tempFile = [System.IO.Path]::Combine(
        [System.IO.Path]::GetTempPath(),
        ("SPOT_Plan_{0}.sqlplan" -f ([Guid]::NewGuid().ToString("N")))
    )

    try {
        [System.IO.File]::WriteAllText($tempFile, [System.Text.Encoding]::UTF8.GetString([System.Text.Encoding]::UTF8.GetBytes([string]$planXml)), [System.Text.Encoding]::UTF8)
        Start-Process $tempFile
        Log-Message "Execution plan opened: $tempFile"
    }
    catch {
        Log-Message "ERROR: Could not open execution plan. $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show(
            "Could not open execution plan viewer on this machine.",
            "SPOT",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
}

$menuCopyRow.Add_Click({
    $src = $gridContextMenu.SourceControl
    if ($src -is [System.Windows.Forms.DataGridView]) {
        Copy-RowToClipboard -grid $src
    }
})

$menuViewPlan.Add_Click({
    $src = $gridContextMenu.SourceControl
    if ($src -is [System.Windows.Forms.DataGridView]) {
        Show-ExecutionPlan -grid $src
    }
})

# =======================
# CONNECTION / CAPTURE / SESSION BUTTON HELPERS
# =======================
function Invoke-ConnectionTest {
    $server = $serverTextBox.Text
    $username = $usernameTextBox.Text
    $password = Get-PasswordFromTextBox
    $connectionString = "Server=$server;Database=master;User ID=$username;Password=$password;Connection Timeout=5;TrustServerCertificate=True"
    $connection = New-Object System.Data.SqlClient.SqlConnection($connectionString)
    try {
        Log-Message "Testing connection to $server..."
        $connection.Open()
        $connectionStatusLabel.Text = "Status: Connection Successful"
        $connectionStatusLabel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#32CD32")
        $captureButton.Enabled = $true
        Log-Message "Connection to $server successful."
    }
    catch {
        $connectionStatusLabel.Text = "Status: Connection Failed"
        $connectionStatusLabel.ForeColor = [System.Drawing.ColorTranslator]::fromHtml("#FF4500")
        $captureButton.Enabled = $false
        Log-Message "ERROR: Failed to connect to '$server'. Exception: $($_.Exception.ToString())"
    }
    finally {
        if ($connection.State -eq 'Open') { $connection.Close() }
    }
}

# =======================
# BUTTON EVENTS
# =======================
$testConnectionButton.Add_Click({
    Invoke-ConnectionTest
})

function Prompt-ForSessionName {
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Save Session As"
    $dlg.Size = New-Object System.Drawing.Size(400,180)
    $dlg.StartPosition = "CenterParent"
    $dlg.BackColor = $bgMain
    $dlg.ForeColor = $fgPrimary
    $dlg.Font = $fontRegular
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "Session Name:"
    $lbl.AutoSize = $true
    $lbl.Location = New-Object System.Drawing.Point(15,20)
    $lbl.ForeColor = $fgPrimary

    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Location = New-Object System.Drawing.Point(120,18)
    $tb.Size = New-Object System.Drawing.Size(250,20)
    $tb.BackColor = $gridRowBg
    $tb.ForeColor = $fgPrimary
    $tb.BorderStyle = 'FixedSingle'

    if (-not [string]::IsNullOrWhiteSpace($script:CurrentSessionName)) {
        $tb.Text = $script:CurrentSessionName
    }

    $okBtn = New-Object System.Windows.Forms.Button
    $okBtn.Text = 'Save'
    $okBtn.Width = 80
    $okBtn.Location = New-Object System.Drawing.Point(210,70)
    $okBtn.BackColor = $btnGray
    $okBtn.ForeColor = $fgPrimary
    $okBtn.FlatStyle = 'Flat'
    $okBtn.Font = $fontRegular

    $cancelBtn = New-Object System.Windows.Forms.Button
    $cancelBtn.Text = 'Cancel'
    $cancelBtn.Width = 80
    $cancelBtn.Location = New-Object System.Drawing.Point(300,70)
    $cancelBtn.BackColor = $btnGray
    $cancelBtn.ForeColor = $fgPrimary
    $cancelBtn.FlatStyle = 'Flat'
    $cancelBtn.Font = $fontRegular

    $dlg.Controls.AddRange(@($lbl,$tb,$okBtn,$cancelBtn))

    $okBtn.Add_Click({
        if (-not [string]::IsNullOrWhiteSpace($tb.Text)) {
            $dlg.Tag = $tb.Text
            $dlg.DialogResult = 'OK'
            $dlg.Close()
        }
    })
    $cancelBtn.Add_Click({
        $dlg.DialogResult = 'Cancel'
        $dlg.Close()
    })

    $res = $dlg.ShowDialog($mainForm)
    if ($res -eq 'OK') {
        return $dlg.Tag
    }
    return $null
}

$saveConnectionButton.Add_Click({
    $sessName = Prompt-ForSessionName
    if ([string]::IsNullOrWhiteSpace($sessName)) {
        Log-Message "Save Session cancelled."
        return
    }

    $securePwd = ($passwordTextBox.Text | ConvertTo-SecureString -AsPlainText -Force)
    Upsert-Session -Name $sessName `
        -Server $serverTextBox.Text `
        -Username $usernameTextBox.Text `
        -Password $securePwd

    $script:CurrentSessionName = $sessName
    $currentSessionValue.Text = $sessName
    Log-Message "Session '$sessName' saved to $script:SessionsFilePath"

    Populate-SessionDropdown
    $sessionComboBox.SelectedItem = $sessName
})

$deleteSessionButton.Add_Click({
    $nameToDelete = $script:CurrentSessionName
    if ([string]::IsNullOrWhiteSpace($nameToDelete)) {
        Log-Message "No active session to delete."
        return
    }

    $res = [System.Windows.Forms.MessageBox]::Show(
        "Delete session '$nameToDelete'?",
        "Confirm Delete",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    if ($res -ne [System.Windows.Forms.DialogResult]::Yes) {
        Log-Message "Delete cancelled."
        return
    }

    Remove-SessionByName -Name $nameToDelete
    Log-Message "Session '$nameToDelete' deleted."

    # Reset UI state
    $serverTextBox.Text   = ""
    $usernameTextBox.Text = ""
    $passwordTextBox.Text = ""
    $passwordTextBox.Tag  = $null
    $currentSessionValue.Text = "(none)"
    $script:CurrentSessionName = $null

    $connectionStatusLabel.Text = "Status: Not Connected"
    $connectionStatusLabel.ForeColor = $fgSecondary
    $captureButton.Enabled = $false

    Populate-SessionDropdown

    # If any session remains, load first and test it
    if ($sessionComboBox.Items.Count -gt 0) {
        $sessionComboBox.SelectedIndex = 0
        $firstName = $sessionComboBox.SelectedItem
        $firstSession = Get-SessionByName -Name $firstName
        if ($firstSession) {
            $script:CurrentSessionName = $firstSession.Name
            $currentSessionValue.Text  = $firstSession.Name
            Set-UiFromSession -session $firstSession
            Save-AllSessions
            Log-Message "Session '$($firstSession.Name)' loaded after delete."
            Invoke-ConnectionTest
        }
    }
})

$sampleRateSlider.Add_ValueChanged({
    $sampleRateValueLabel.Text = $sampleRateSlider.Value
})
$numSamplesSlider.Add_ValueChanged({
    $numSamplesValueLabel.Text = $numSamplesSlider.Value
})

$captureButton.Add_Click({
    $captureButton.Enabled = $false
    $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        $connStr = "Server=$($serverTextBox.Text);Database=utility;User ID=$($usernameTextBox.Text);Password=$(Get-PasswordFromTextBox);TrustServerCertificate=True"
        $sampleRate = $sampleRateSlider.Value
        $numSamples = $numSamplesSlider.Value
        $getPlans = 1

        #Log-Message "Executing SPOT.CaptureSnapshot (Samples=$numSamples, Interval=$sampleRate s, GetPlans=$getPlans)"
        Log-Message ("Executing SPOT - estimated runtime {0} seconds" -f ($numSamples * $sampleRate))

        $summary = Get-SqlData -ConnectionString $connStr -Query @"
EXEC SPOT.CaptureSnapshot
    @SampleCount       = $numSamples,
    @SampleTimeSeconds = $sampleRate,
    @GetPlans          = 1;
"@
        if (-not $summary -or $summary.Count -eq 0) {
            throw "Capture returned no summary output"
        }

        $snapId = [datetime]$summary[0].SnapshotId
        Log-Message "Capture complete. SnapshotId: $($snapId.ToString('yyyy-MM-dd HH:mm:ss.fff'))"
        Load-SnapshotById -ConnectionString $connStr -SnapshotId $snapId
    }
    catch {
        Log-Message "CAPTURE FAILED: $($_.Exception.ToString())"
    }
    finally {
        $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default
        $captureButton.Enabled = $true
    }
})

$loadSnapshotButton.Add_Click({
    $connStr = "Server=$($serverTextBox.Text);Database=utility;User ID=$($usernameTextBox.Text);Password=$(Get-PasswordFromTextBox);TrustServerCertificate=True"
    Show-SnapshotPicker -ConnectionString $connStr
})

# =======================
# SNAPSHOT PICKER DIALOG
# =======================
function Show-SnapshotPicker {
    param([string]$ConnectionString)

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Choose Snapshot"
    $dlg.Size = New-Object System.Drawing.Size(760, 460)
    $dlg.StartPosition = "CenterParent"
    $dlg.BackColor = $bgMain
    $dlg.ForeColor = $fgPrimary
    $dlg.Font = $fontRegular

    $lv = New-Object System.Windows.Forms.ListView
    $lv.View = 'Details'
    $lv.FullRowSelect = $true
    $lv.HideSelection = $false
    $lv.GridLines = $true
    $lv.Dock = 'Top'
    $lv.Height = 340
    $lv.BackColor = $gridRowBg
    $lv.ForeColor = $fgPrimary
    $lv.Font = $fontRegular

    [void]$lv.Columns.Add("SnapshotId",          170)
    [void]$lv.Columns.Add("StartedAtUtc",        150)
    [void]$lv.Columns.Add("ServerName",          130)
    [void]$lv.Columns.Add("SampleCount",          90)
    [void]$lv.Columns.Add("SampleTimeSeconds",   120)
    [void]$lv.Columns.Add("GetPlans",             70)

    $statusLbl = New-Object System.Windows.Forms.Label
    $statusLbl.Text = "Querying..."
    $statusLbl.AutoSize = $true
    $statusLbl.Location = New-Object System.Drawing.Point(10, 345)
    $statusLbl.ForeColor = $fgPrimary
    $statusLbl.Font = $fontRegular

    $refreshBtn = New-Object System.Windows.Forms.Button
    $refreshBtn.Text = 'Refresh'
    $refreshBtn.Width = 100
    $refreshBtn.Location = New-Object System.Drawing.Point(410, 380)
    $refreshBtn.BackColor = $btnGray
    $refreshBtn.ForeColor = $fgPrimary
    $refreshBtn.FlatStyle = 'Flat'
    $refreshBtn.Font = $fontRegular

    $okBtn = New-Object System.Windows.Forms.Button
    $okBtn.Text = 'Load'
    $okBtn.Width = 100
    $okBtn.Location = New-Object System.Drawing.Point(520, 380)
    $okBtn.BackColor = $btnGray
    $okBtn.ForeColor = $fgPrimary
    $okBtn.FlatStyle = 'Flat'
    $okBtn.Font = $fontRegular

    $cancelBtn = New-Object System.Windows.Forms.Button
    $cancelBtn.Text = 'Cancel'
    $cancelBtn.Width = 100
    $cancelBtn.Location = New-Object System.Drawing.Point(630, 380)
    $cancelBtn.BackColor = $btnGray
    $cancelBtn.ForeColor = $fgPrimary
    $cancelBtn.FlatStyle = 'Flat'
    $cancelBtn.Font = $fontRegular

    $dlg.Controls.AddRange(@($lv, $statusLbl, $refreshBtn, $okBtn, $cancelBtn))

    $query = @"
SELECT TOP 200
       SnapshotId,
       StartedAtUtc,
       ServerName,
       SampleCount,
       SampleTimeSeconds,
       GetPlans
FROM SPOT.Snapshots
ORDER BY StartedAtUtc DESC;
"@

    $load = {
        try {
            $statusLbl.Text = "Querying..."
            $lv.BeginUpdate()
            $lv.Items.Clear()

            $rows = Get-SqlData -ConnectionString $ConnectionString -Query $query
            if (-not $rows -or $rows.Count -eq 0) {
                $statusLbl.Text = "No snapshots found."
            } else {
                foreach ($r in $rows) {
                    $snapStr  = ([datetime]$r.SnapshotId).ToString('yyyy-MM-dd HH:mm:ss.fff')
                    $startStr = ([datetime]$r.StartedAtUtc).ToString('yyyy-MM-dd HH:mm:ss.fff')
                    $srv      = [string]$r.ServerName
                    $cnt      = [string]$r.SampleCount
                    $secs     = [string]$r.SampleTimeSeconds
                    $gp       = if ([int]$r.GetPlans -eq 1) { '1' } else { '0' }

                    $item = New-Object System.Windows.Forms.ListViewItem($snapStr)
                    [void]$item.SubItems.Add($startStr)
                    [void]$item.SubItems.Add($srv)
                    [void]$item.SubItems.Add($cnt)
                    [void]$item.SubItems.Add($secs)
                    [void]$item.SubItems.Add($gp)

                    $item.Tag = [datetime]$r.SnapshotId
                    [void]$lv.Items.Add($item)
                }
                $statusLbl.Text = ("Loaded {0} snapshots" -f $rows.Count)
            }
        }
        catch {
            $statusLbl.Text = "Error querying snapshots (see console)"
            Log-Message "ERROR: Could not query SPOT.Snapshots. $($_.Exception.Message)"
        }
        finally {
            $lv.EndUpdate()
        }
    }

    & $load
    $refreshBtn.Add_Click({ & $load })

    $okBtn.Add_Click({
        if ($lv.SelectedItems.Count -gt 0) {
            $dlg.DialogResult = 'OK'
            $dlg.Close()
        }
    })

    $cancelBtn.Add_Click({
        $dlg.DialogResult = 'Cancel'
        $dlg.Close()
    })

    if ($dlg.ShowDialog($mainForm) -eq 'OK' -and $lv.SelectedItems.Count -gt 0) {
        $snapId = [datetime]$lv.SelectedItems[0].Tag
        $connStr = "Server=$($serverTextBox.Text);Database=utility;User ID=$($usernameTextBox.Text);Password=$(Get-PasswordFromTextBox);TrustServerCertificate=True"
        Load-SnapshotById -ConnectionString $connStr -SnapshotId $snapId
    }
}

# =======================
# TIMELINE SLIDER EVENT
# =======================
$timelineTrackBar.Add_ValueChanged({
    if ($script:isUpdatingTimeline) { return }
    Set-TimelineIndex -index $timelineTrackBar.Value
})

# =======================
# SESSION COMBO EVENTS
# =======================
$sessionComboBox.add_SelectedIndexChanged({
    $chosenName = $sessionComboBox.SelectedItem
    if ([string]::IsNullOrWhiteSpace($chosenName)) { return }

    $chosenSession = Get-SessionByName -Name $chosenName
    if ($chosenSession) {
        $script:CurrentSessionName = $chosenSession.Name
        $currentSessionValue.Text  = $chosenSession.Name
        Set-UiFromSession -session $chosenSession
        Save-AllSessions
        Log-Message "Session '$($chosenSession.Name)' loaded from selector."
        Invoke-ConnectionTest
    }
})

# =======================
# SESSION DROPDOWN POPULATION
# =======================
function Populate-SessionDropdown {
    Ensure-SessionsArrayList

    $sessionComboBox.Items.Clear()

    foreach ($s in $script:Sessions) {
        [void]$sessionComboBox.Items.Add($s.Name)
    }

    if ($sessionComboBox.Items.Count -gt 0) {
        if (-not [string]::IsNullOrWhiteSpace($script:CurrentSessionName) -and
            $sessionComboBox.Items.Contains($script:CurrentSessionName)) {

            $sessionComboBox.SelectedItem = $script:CurrentSessionName
        }
        else {
            $sessionComboBox.SelectedIndex = 0
        }
    }
    else {
        # no sessions: leave dropdown blank & enabled for visual consistency
        $sessionComboBox.Text = ""
    }
}

# =======================
# FORM LIFECYCLE
# =======================
$mainForm.Add_Shown({
    Load-AllSessions
    Load-SavedConnection
    Populate-SessionDropdown

    if (-not [string]::IsNullOrWhiteSpace($script:CurrentSessionName)) {
        $currentSessionValue.Text = $script:CurrentSessionName
    } else {
        $currentSessionValue.Text = "(none)"
    }

    Load-LastSessionToUi
})

[void]$mainForm.ShowDialog()
$mainForm.Dispose()
'@

Set-Content -Path $spotPath -Value $spotContent -Encoding UTF8

# -----------------------------
# Write Desktop launcher
# -----------------------------
$launcherContent = @'
# Launch-SPOT.ps1

$scriptPath = Join-Path $env:USERPROFILE "Documents\SQL Performance Observation Tool\SPOT.ps1"

if (Test-Path $scriptPath) {
    Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""
} else {
    Write-Host "SPOT script not found at: $scriptPath" -ForegroundColor Red
}
'@

Set-Content -Path $launcher -Value $launcherContent -Encoding UTF8

Write-Host "Deployed SPOT to: $spotPath" -ForegroundColor Green
Write-Host "Desktop launcher created: $launcher" -ForegroundColor Green
