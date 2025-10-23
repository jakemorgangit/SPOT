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
A PowerShell GUI for capturing SQL Server performance snapshots.
.DESCRIPTION
Utility for capturing and reviewing SQL Server performance metrics using a snapshot sampling mechanism.
.NOTES
Author: Jake Morgan - Blackcat Data Solutions Limited 2025
Date: October 2025
Prerequisites:
- PowerShell 5.1 or higher
- SQL Server 2012 or higher
- Adam Machanic's sp_whoisactive stored procedure installed in a "utility" database on the target server.
- The user account requires VIEW SERVER STATE permission and EXECUTE permission on sp_whoisactive.
#>

# --- Small start-up feature ---
# Ensure WinForms-friendly apartment state and show a clear prereq message if PS is too old.
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
if ($PSVersionTable.PSVersion.Major -lt 5 -or ($PSVersionTable.PSVersion.Major -eq 5 -and $PSVersionTable.PSVersion.Minor -lt 1)) {
    [System.Windows.Forms.MessageBox]::Show(
        "PowerShell 5.1 or higher is required. Current: $($PSVersionTable.PSVersion)",
        "SPOT prerequisite",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
    exit 1
}

#region Assembly Loading
Add-Type -AssemblyName System.Data
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
#endregion

#region GUI Definition & Theming

# --- Main Form ---
$mainForm = New-Object System.Windows.Forms.Form
$mainForm.Text = "SQL Performance Observation Tool"
$mainForm.Size = New-Object System.Drawing.Size(1200, 800)
$mainForm.MinimumSize = New-Object System.Drawing.Size(800, 600)
$mainForm.StartPosition = "CenterScreen"
$mainForm.FormBorderStyle = "Sizable"
$mainForm.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#2D2D30")
$mainForm.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#F1F1F1")

# --- UI Colour Palette ---
$darkGray     = [System.Drawing.ColorTranslator]::FromHtml("#3E3E42")
$darkerGray   = [System.Drawing.ColorTranslator]::FromHtml("#2D2D30")
$lightGray    = [System.Drawing.ColorTranslator]::FromHtml("#555555")
$gridLineGray = [System.Drawing.ColorTranslator]::FromHtml("#4A4A4A")
$whiteText    = [System.Drawing.ColorTranslator]::FromHtml("#F1F1F1")
$blueAccent   = [System.Drawing.ColorTranslator]::FromHtml("#007ACC")
$redHighlight = [System.Drawing.ColorTranslator]::FromHtml("#7A2A2A")
$consoleBack  = [System.Drawing.ColorTranslator]::FromHtml("#1E1E1E")
$consoleFore  = [System.Drawing.ColorTranslator]::FromHtml("#D4D4D4")

#region Connection GroupBox
$connectionGroupBox = New-Object System.Windows.Forms.GroupBox; $connectionGroupBox.Text = "Connection Details"; $connectionGroupBox.Size = New-Object System.Drawing.Size(300, 220); $connectionGroupBox.Location = New-Object System.Drawing.Point(20, 20); $connectionGroupBox.ForeColor = $whiteText; $connectionGroupBox.Anchor = "Top, Left"; $mainForm.Controls.Add($connectionGroupBox)
$serverLabel = New-Object System.Windows.Forms.Label; $serverLabel.Text = "Server/Instance:"; $serverLabel.Location = New-Object System.Drawing.Point(10, 30); $serverLabel.AutoSize = $true
$serverTextBox = New-Object System.Windows.Forms.TextBox; $serverTextBox.Location = New-Object System.Drawing.Point(120, 27); $serverTextBox.Size = New-Object System.Drawing.Size(160, 20); $serverTextBox.BackColor = $darkGray; $serverTextBox.ForeColor = $whiteText
$usernameLabel = New-Object System.Windows.Forms.Label; $usernameLabel.Text = "Username:"; $usernameLabel.Location = New-Object System.Drawing.Point(10, 60); $usernameLabel.AutoSize = $true
$usernameTextBox = New-Object System.Windows.Forms.TextBox; $usernameTextBox.Location = New-Object System.Drawing.Point(120, 57); $usernameTextBox.Size = New-Object System.Drawing.Size(160, 20); $usernameTextBox.BackColor = $darkGray; $usernameTextBox.ForeColor = $whiteText
$passwordLabel = New-Object System.Windows.Forms.Label; $passwordLabel.Text = "Password:"; $passwordLabel.Location = New-Object System.Drawing.Point(10, 90); $passwordLabel.AutoSize = $true
$passwordTextBox = New-Object System.Windows.Forms.TextBox; $passwordTextBox.Location = New-Object System.Drawing.Point(120, 87); $passwordTextBox.Size = New-Object System.Drawing.Size(160, 20); $passwordTextBox.UseSystemPasswordChar = $true; $passwordTextBox.BackColor = $darkGray; $passwordTextBox.ForeColor = $whiteText
$testConnectionButton = New-Object System.Windows.Forms.Button; $testConnectionButton.Text = "Test Connection"; $testConnectionButton.Location = New-Object System.Drawing.Point(10, 130); $testConnectionButton.Size = New-Object System.Drawing.Size(130, 30); $testConnectionButton.BackColor = $lightGray; $testConnectionButton.FlatStyle = "Flat"
$saveConnectionButton = New-Object System.Windows.Forms.Button; $saveConnectionButton.Text = "Save Connection"; $saveConnectionButton.Location = New-Object System.Drawing.Point(150, 130); $saveConnectionButton.Size = New-Object System.Drawing.Size(130, 30); $saveConnectionButton.BackColor = $lightGray; $saveConnectionButton.FlatStyle = "Flat"
$connectionStatusLabel = New-Object System.Windows.Forms.Label; $connectionStatusLabel.Text = "Status: Not Connected"; $connectionStatusLabel.Location = New-Object System.Drawing.Point(10, 175); $connectionStatusLabel.AutoSize = $true
$connectionGroupBox.Controls.AddRange(@($serverLabel, $serverTextBox, $usernameLabel, $usernameTextBox, $passwordLabel, $passwordTextBox, $testConnectionButton, $saveConnectionButton, $connectionStatusLabel))
#endregion

#region Snapshot Configuration GroupBox
$snapshotGroupBox = New-Object System.Windows.Forms.GroupBox; $snapshotGroupBox.Text = "Snapshot Configuration"; $snapshotGroupBox.Size = New-Object System.Drawing.Size(300, 200); $snapshotGroupBox.Location = New-Object System.Drawing.Point(20, 260); $snapshotGroupBox.ForeColor = $whiteText; $snapshotGroupBox.Anchor = "Top, Left"; $mainForm.Controls.Add($snapshotGroupBox)
$sampleRateLabel = New-Object System.Windows.Forms.Label; $sampleRateLabel.Text = "Sample Rate (s):"; $sampleRateLabel.Location = New-Object System.Drawing.Point(10, 30); $sampleRateLabel.AutoSize = $true
$sampleRateValueLabel = New-Object System.Windows.Forms.Label; $sampleRateValueLabel.Location = New-Object System.Drawing.Point(260, 30)
$sampleRateSlider = New-Object System.Windows.Forms.TrackBar; $sampleRateSlider.Location = New-Object System.Drawing.Point(120, 25); $sampleRateSlider.Size = New-Object System.Drawing.Size(140, 45); $sampleRateSlider.Minimum = 1; $sampleRateSlider.Maximum = 12; $sampleRateSlider.Value = 5; $sampleRateValueLabel.Text = $sampleRateSlider.Value
$numSamplesLabel = New-Object System.Windows.Forms.Label; $numSamplesLabel.Text = "Number of Samples:"; $numSamplesLabel.Location = New-Object System.Drawing.Point(10, 80); $numSamplesLabel.AutoSize = $true
$numSamplesValueLabel = New-Object System.Windows.Forms.Label; $numSamplesValueLabel.Location = New-Object System.Drawing.Point(260, 80)
$numSamplesSlider = New-Object System.Windows.Forms.TrackBar; $numSamplesSlider.Location = New-Object System.Drawing.Point(120, 75); $numSamplesSlider.Size = New-Object System.Drawing.Size(140, 45); $numSamplesSlider.Minimum = 1; $numSamplesSlider.Maximum = 12; $numSamplesSlider.Value = 3; $numSamplesValueLabel.Text = $numSamplesSlider.Value
$captureButton = New-Object System.Windows.Forms.Button; $captureButton.Text = "Start Capture"; $captureButton.Location = New-Object System.Drawing.Point(10, 140); $captureButton.Size = New-Object System.Drawing.Size(280, 40); $captureButton.BackColor = $blueAccent; $captureButton.FlatStyle = "Flat"; $captureButton.Enabled = $false
$snapshotGroupBox.Controls.AddRange(@($sampleRateLabel, $sampleRateValueLabel, $sampleRateSlider, $numSamplesLabel, $numSamplesValueLabel, $numSamplesSlider, $captureButton))
#endregion

#region Console Log GroupBox
$consoleGroupBox = New-Object System.Windows.Forms.GroupBox; $consoleGroupBox.Text = "Console Log"; $consoleGroupBox.Size = New-Object System.Drawing.Size(820, 220); $consoleGroupBox.Location = New-Object System.Drawing.Point(340, 20); $consoleGroupBox.ForeColor = $whiteText; $consoleGroupBox.Anchor = "Top, Left, Right"; $mainForm.Controls.Add($consoleGroupBox)
$consoleLogTextBox = New-Object System.Windows.Forms.TextBox; $consoleLogTextBox.Location = New-Object System.Drawing.Point(10, 20); $consoleLogTextBox.Size = New-Object System.Drawing.Size(800, 190); $consoleLogTextBox.Multiline = $true; $consoleLogTextBox.ScrollBars = "Vertical"; $consoleLogTextBox.ReadOnly = $true; $consoleLogTextBox.BackColor = $consoleBack; $consoleLogTextBox.ForeColor = $consoleFore; $consoleLogTextBox.Anchor = "Top, Bottom, Left, Right"; $consoleGroupBox.Controls.Add($consoleLogTextBox)
#endregion

#region Results Area
$sampleSelectorGroupBox = New-Object System.Windows.Forms.GroupBox; $sampleSelectorGroupBox.Text = "Select Sample"; $sampleSelectorGroupBox.Size = New-Object System.Drawing.Size(820, 80); $sampleSelectorGroupBox.Location = New-Object System.Drawing.Point(340, 260); $sampleSelectorGroupBox.ForeColor = $whiteText; $sampleSelectorGroupBox.Visible = $false; $sampleSelectorGroupBox.Anchor = "Top, Left, Right"; $mainForm.Controls.Add($sampleSelectorGroupBox)
$sampleFlowPanel = New-Object System.Windows.Forms.FlowLayoutPanel; $sampleFlowPanel.Dock = "Fill"; $sampleSelectorGroupBox.Controls.Add($sampleFlowPanel)
$resultsTabControl = New-Object System.Windows.Forms.TabControl; $resultsTabControl.Location = New-Object System.Drawing.Point(340, 350); $resultsTabControl.Size = New-Object System.Drawing.Size(820, 410); $resultsTabControl.Visible = $false; $resultsTabControl.Anchor = "Top, Bottom, Left, Right"; $mainForm.Controls.Add($resultsTabControl)
$loadSnapshotButton = New-Object System.Windows.Forms.Button; $loadSnapshotButton.Text = "Load Snapshot Data"; $loadSnapshotButton.Location = New-Object System.Drawing.Point(20, 480); $loadSnapshotButton.Size = New-Object System.Drawing.Size(300, 40); $loadSnapshotButton.BackColor = $lightGray; $loadSnapshotButton.FlatStyle = "Flat"; $loadSnapshotButton.Anchor = "Top, Left"; $mainForm.Controls.Add($loadSnapshotButton)
#endregion
#endregion

#region Core Functions
function Log-Message { param([string]$Message)
    $consoleLogTextBox.AppendText("$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $Message`r`n")
    $consoleLogTextBox.Update()
}

function Get-SqlData {
    param([string]$ConnectionString, [string]$Query)
    $connection = New-Object System.Data.SqlClient.SqlConnection($ConnectionString)
    $command = New-Object System.Data.SqlClient.SqlCommand($Query, $connection)
    $command.CommandTimeout = 120
    $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($command)
    $dataset = New-Object System.Data.DataSet
    try {
        $connection.Open()
        $adapter.Fill($dataset) | Out-Null
    }
    catch { throw $_ }
    finally { if ($connection.State -eq 'Open') { $connection.Close() } }
    if ($dataset.Tables.Count -eq 0) { return @() }
    $results = New-Object System.Collections.ArrayList
    foreach ($row in $dataset.Tables[0].Rows) {
        $obj = New-Object -TypeName PSObject
        foreach ($col in $row.Table.Columns) {
            $obj | Add-Member -MemberType NoteProperty -Name $col.ColumnName -Value $row[$col]
        }
        $results.Add($obj) | Out-Null
    }
    return $results
}

function Load-SavedConnection {
    $configFile = "$PSScriptRoot\config.xml"
    if (Test-Path $configFile) {
        try {
            $credential = Import-CliXml -Path $configFile
            $user, $server = $credential.UserName.Split('@')
            $serverTextBox.Text = $server
            $usernameTextBox.Text = $user
            $passwordTextBox.Tag = $credential.Password
            $passwordTextBox.Text = '••••••••••'
            Log-Message "Saved connection details loaded."
        }
        catch { Log-Message "Could not load saved connection details. File may be corrupt." }
    }
}

function Get-PasswordFromTextBox {
    if ($passwordTextBox.Tag -is [System.Security.SecureString]) {
        $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($passwordTextBox.Tag)
        $password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
        return $password
    }
    else { return $passwordTextBox.Text }
}

function Apply-DarkModeToGrid {
    param([System.Windows.Forms.DataGridView]$grid)
    $grid.EnableHeadersVisualStyles = $false
    $grid.ColumnHeadersBorderStyle = 'None'
    $grid.RowHeadersVisible = $false

    $headerStyle = New-Object System.Windows.Forms.DataGridViewCellStyle
    $headerStyle.BackColor = $darkerGray
    $headerStyle.ForeColor = $whiteText
    $headerStyle.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $grid.ColumnHeadersDefaultCellStyle = $headerStyle

    $cellStyle = New-Object System.Windows.Forms.DataGridViewCellStyle
    $cellStyle.BackColor = $darkGray
    $cellStyle.ForeColor = $whiteText
    $cellStyle.SelectionBackColor = $blueAccent
    $cellStyle.SelectionForeColor = $whiteText
    $grid.DefaultCellStyle = $cellStyle

    $grid.BackgroundColor = $darkGray
    $grid.GridColor = $gridLineGray
    $grid.BorderStyle = 'None'
}

function Highlight-WhoIsActiveRows {
    param([System.Windows.Forms.DataGridView]$grid)
    $longWaitThresholdMs = 1000
    $redStyle = New-Object System.Windows.Forms.DataGridViewCellStyle
    $redStyle.BackColor = $redHighlight
    $redStyle.ForeColor = $whiteText
    $redStyle.SelectionBackColor = $redHighlight
    $redStyle.SelectionForeColor = $whiteText

    foreach ($row in $grid.Rows) {
        $data = $row.DataBoundItem
        $isProblem = $false
        if ($data.'Blocking Info') { $isProblem = $true }
        elseif ($data.wait_info -and $data.wait_info -notmatch 'WAITFOR') {
            if ($data.wait_info -match '\((\d+)ms\)') {
                if ([int]$Matches[1] -gt $longWaitThresholdMs) { $isProblem = $true }
            }
        }
        if ($isProblem) { $row.DefaultCellStyle = $redStyle }
    }
}

function Load-SnapshotData {
    param([string]$FilePath)
    try {
        Log-Message "Loading snapshot data from $FilePath..."
        $script:loadedSnapshotData = Get-Content -Path $FilePath -Raw | ConvertFrom-Json
        $sampleFlowPanel.Controls.Clear()
        $resultsTabControl.TabPages.Clear()
        $script:dataGrids.Clear()
        $script:sortState.Clear()

        $dataTypes = @("WhoIsActive", "HealthChecks", "AGHealth", "WaitStats", "Blocking")
        foreach ($type in $dataTypes) {
            $tabPage = New-Object System.Windows.Forms.TabPage; $tabPage.Text = $type
            $grid = New-Object System.Windows.Forms.DataGridView -Property @{ Name = $type; Dock = "Fill"; AutoSizeColumnsMode = "None"; AllowUserToResizeColumns = $true }
            Apply-DarkModeToGrid -grid $grid
            $grid.add_ColumnHeaderMouseClick($On_ColumnHeaderClick)
            $grid.add_DataBindingComplete({ param($s, $e) On_GridDataBindingComplete -sender $s -e $e })
            if ($type -eq 'WhoIsActive') {
                $grid.add_DataBindingComplete({
                    param($s, $e)
                    if ($this.Columns["Blocking Info"]) { $this.Columns["Blocking Info"].DisplayIndex = 0 }
                    Highlight-WhoIsActiveRows -grid $this
                })
            }
            $tabPage.Controls.Add($grid)
            $resultsTabControl.TabPages.Add($tabPage)
            $script:dataGrids[$type] = $grid
        }

        for ($i = 0; $i -lt $script:loadedSnapshotData.Samples.Count; $i++) {
            $sample = $script:loadedSnapshotData.Samples[$i]
            $rb = New-Object System.Windows.Forms.RadioButton
            $rb.Text = "Sample $($i+1) - $([datetime]$sample.Timestamp | Get-Date -Format 'HH:mm:ss')"
            $rb.Tag = $i
            $rb.AutoSize = $true
            $rb.add_CheckedChanged($radio_CheckedChanged)
            $sampleFlowPanel.Controls.Add($rb)
        }

        $sampleSelectorGroupBox.Visible = $true
        $resultsTabControl.Visible = $true
        if ($sampleFlowPanel.Controls.Count -gt 0) { $sampleFlowPanel.Controls[0].Checked = $true }
        Log-Message "Snapshot data loaded successfully."
    }
    catch {
        Log-Message "ERROR: Failed to load or parse snapshot file '$FilePath'. Exception: $($_.Exception.ToString())"
    }
}
#endregion

#region Event Handlers
function On_GridDataBindingComplete {
    param($sender, $e)
    $grid = $sender
    foreach ($column in $grid.Columns) {
        switch ($column.Name) {
            'sql_text'          { $column.AutoSizeMode = 'None'; $column.Width = 200; break }
            'waiting_query'     { $column.AutoSizeMode = 'None'; $column.Width = 200; break }
            'blocking_query'    { $column.AutoSizeMode = 'None'; $column.Width = 200; break }
            'Purpose'           { $column.AutoSizeMode = 'None'; $column.Width = 350; break }
            'Check Description' { $column.AutoSizeMode = 'None'; $column.Width = 250; break }
            default             { $column.AutoSizeMode = 'AllCells' }
        }
    }
}

$testConnectionButton.Add_Click({
    $server = $serverTextBox.Text; $username = $usernameTextBox.Text; $password = Get-PasswordFromTextBox
    $connectionString = "Server=$server;Database=master;User ID=$username;Password=$password;Connection Timeout=5;TrustServerCertificate=True"
    $connection = New-Object System.Data.SqlClient.SqlConnection($connectionString)
    try {
        Log-Message "Testing connection to $server..."
        $connection.Open()
        $connectionStatusLabel.Text = "Status: Connection Successful"; $connectionStatusLabel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#32CD32")
        $captureButton.Enabled = $true
        Log-Message "Connection to $server successful."
    }
    catch {
        $connectionStatusLabel.Text = "Status: Connection Failed"; $connectionStatusLabel.ForeColor = [System.Drawing.ColorTranslator]::fromHtml("#FF4500")
        $captureButton.Enabled = $false
        Log-Message "ERROR: Failed to connect to '$server'. Exception: $($_.Exception.ToString())"
    }
    finally { if ($connection.State -eq 'Open') { $connection.Close() } }
})

$saveConnectionButton.Add_Click({
    $configFile = "$PSScriptRoot\config.xml"
    $credential = New-Object System.Management.Automation.PSCredential("$($usernameTextBox.Text)@$($serverTextBox.Text)", ($passwordTextBox.Text | ConvertTo-SecureString -AsPlainText -Force))
    $credential | Export-CliXml -Path $configFile
    $passwordTextBox.Tag = $credential.Password
    Log-Message "Connection details saved to $configFile"
})

$sampleRateSlider.Add_ValueChanged({ $sampleRateValueLabel.Text = $sampleRateSlider.Value })
$numSamplesSlider.Add_ValueChanged({ $numSamplesValueLabel.Text = $numSamplesSlider.Value })

$captureButton.Add_Click({
    $captureButton.Enabled = $false
    $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        $connectionString = "Server=$($serverTextBox.Text);Database=utility;User ID=$($usernameTextBox.Text);Password=$(Get-PasswordFromTextBox);TrustServerCertificate=True"
        $sampleRate = $sampleRateSlider.Value; $numSamples = $numSamplesSlider.Value
        Log-Message "Starting snapshot capture. The application will be unresponsive during this time."

        $snapshotData = @{ CaptureTimestamp = Get-Date -Format 'o'; ServerName = $serverTextBox.Text; Samples = New-Object System.Collections.ArrayList }

        $healthCheckQuery = @"
SET NOCOUNT ON; SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED;
DECLARE @HealthChecks TABLE ([Check Description] NVARCHAR(255), [Purpose] NVARCHAR(1000), [Current Value] NVARCHAR(MAX));
DECLARE @AvgSqlCpuUtilization INT, @AvgTotalCpuUtilization INT;
;WITH RingBuffer AS (SELECT record.value('(./Record/SchedulerMonitorEvent/SystemHealth/SystemIdle)[1]', 'int') AS SystemIdle, record.value('(./Record/SchedulerMonitorEvent/SystemHealth/ProcessUtilization)[1]', 'int') AS ProcessUtilization FROM (SELECT CONVERT(XML, record) AS record FROM sys.dm_os_ring_buffers WHERE ring_buffer_type = N'RING_BUFFER_SCHEDULER_MONITOR') AS t)
SELECT @AvgSqlCpuUtilization = AVG(ProcessUtilization), @AvgTotalCpuUtilization = 100 - AVG(SystemIdle) FROM RingBuffer;
INSERT INTO @HealthChecks ([Check Description], [Purpose], [Current Value]) VALUES ('CPU Utilization (SQL Server Process)', 'Shows the percentage of total CPU time being used by the SQL Server process itself. High values indicate CPU pressure from queries.', ISNULL(CAST(@AvgSqlCpuUtilization AS NVARCHAR(10)) + '%', 'Data not available')), ('CPU Utilization (Total - Including OS)', 'Shows the total CPU usage on the server. If this is high but SQL CPU is low, another process is consuming CPU resources.', ISNULL(CAST(@AvgTotalCpuUtilization AS NVARCHAR(10)) + '%', 'Data not available'));
INSERT INTO @HealthChecks ([Check Description], [Purpose], [Current Value]) SELECT 'Runnable Schedulers (Signal Waits)', 'A count of tasks ready to run but waiting for CPU. Consistently high values (e.g., > 5-10) indicate significant CPU pressure.', CAST(SUM(runnable_tasks_count) AS NVARCHAR(10)) FROM sys.dm_os_schedulers WHERE scheduler_id < 255;
INSERT INTO @HealthChecks ([Check Description], [Purpose], [Current Value]) SELECT 'Memory Grants Pending', 'Number of queries waiting for a memory grant to execute. Any value > 0 is a clear sign of memory pressure.', CAST(COUNT(*) AS NVARCHAR(10)) FROM sys.dm_exec_query_memory_grants WHERE grant_time IS NULL;
DECLARE @FinalPLEValue NVARCHAR(MAX), @FinalBCHRValue NVARCHAR(MAX);
BEGIN TRY
    DECLARE @PLEValue BIGINT; SELECT @PLEValue = cntr_value FROM sys.dm_os_performance_counters WHERE [object_name] LIKE '%Buffer Manager%' AND counter_name = 'Page life expectancy'; SET @FinalPLEValue = ISNULL(CAST(@PLEValue AS NVARCHAR(20)) + ' seconds', 'N/A');
    DECLARE @BCHRValue DECIMAL(10,2); SELECT @BCHRValue = CAST((a.cntr_value * 1.0 / b.cntr_value) * 100 AS DECIMAL(10,2)) FROM sys.dm_os_performance_counters AS a JOIN sys.dm_os_performance_counters AS b ON a.object_name = b.object_name WHERE a.counter_name = 'Buffer cache hit ratio' AND b.counter_name = 'Buffer cache hit ratio base' AND a.object_name LIKE '%Buffer Manager%'; SET @FinalBCHRValue = ISNULL(CAST(@BCHRValue AS NVARCHAR(20)) + '%', 'N/A');
END TRY BEGIN CATCH IF @FinalPLEValue IS NULL SET @FinalPLEValue = 'Calculation Error'; IF @FinalBCHRValue IS NULL SET @FinalBCHRValue = 'Calculation Error'; END CATCH;
INSERT INTO @HealthChecks ([Check Description], [Purpose], [Current Value]) VALUES ('Page Life Expectancy (PLE)', 'Seconds a data page will stay in cache. A low value indicates memory pressure.', @FinalPLEValue);
INSERT INTO @HealthChecks ([Check Description], [Purpose], [Current Value]) VALUES ('Buffer Cache Hit Ratio', 'Percentage of pages found in memory without having to be read from disk. Should be > 95% for OLTP systems. Low values indicate insufficient memory.', @FinalBCHRValue);
INSERT INTO @HealthChecks ([Check Description], [Purpose], [Current Value]) SELECT 'Pending I/O Requests', 'The number of I/O requests currently waiting to be completed by the storage subsystem. A consistently high value indicates an I/O bottleneck.', CAST(COUNT(*) AS NVARCHAR(10)) FROM sys.dm_io_pending_io_requests;
INSERT INTO @HealthChecks ([Check Description], [Purpose], [Current Value]) SELECT 'TempDB Page Latch Contention', 'Checks for tasks waiting on key allocation pages in TempDB. Any value > 0 indicates active TempDB contention, often solved by adding more TempDB files.', CAST(COUNT(*) AS NVARCHAR(10)) FROM sys.dm_os_waiting_tasks WHERE wait_type LIKE 'PAGELATCH%' AND resource_description LIKE '2:%';
IF OBJECT_ID('tempdb..#LogSpace') IS NOT NULL DROP TABLE #LogSpace; CREATE TABLE #LogSpace ([Database Name] NVARCHAR(128), [Log Size (MB)] DECIMAL(18, 5), [Log Space Used (%)] DECIMAL(18, 5), [Status] INT); INSERT INTO #LogSpace EXEC ('DBCC SQLPERF(LOGSPACE);');
INSERT INTO @HealthChecks ([Check Description], [Purpose], [Current Value]) SELECT TOP 1 'Highest Transaction Log Usage', 'The database with the most-full transaction log. A log > 90% is at risk of growing or halting all data modification in that database.', 'DB: ' + [Database Name] + ' is ' + CAST([Log Space Used (%)] AS NVARCHAR(20)) + '% full.' FROM #LogSpace WHERE [Database Name] NOT IN ('master', 'model', 'msdb', 'tempdb', 'mssqlsystemresource','utility','monitoring') ORDER BY [Log Space Used (%)] DESC;
DROP TABLE #LogSpace;
INSERT INTO @HealthChecks ([Check Description], [Purpose], [Current Value]) SELECT 'Oldest Active Transaction (Elapsed Time)', 'The elapsed time in seconds of the longest-running open user transaction. Long transactions cause blocking and prevent transaction log truncation.', ISNULL(CAST(DATEDIFF(second, MIN(t.transaction_begin_time), GETDATE()) AS NVARCHAR(20)) + ' seconds', 'N/A') FROM sys.dm_tran_active_transactions AS t INNER JOIN sys.dm_tran_session_transactions AS s ON t.transaction_id = s.transaction_id WHERE s.session_id > 50;
SELECT [Check Description], [Purpose], [Current Value] FROM @HealthChecks;
"@

        for ($i = 1; $i -le $numSamples; $i++) {
            Log-Message "Capturing sample $i of $numSamples..."
            $whoIsActiveQuery = "EXEC sp_whoisactive @output_column_list = '[dd hh:mm:ss.mss][session_id][sql_text][login_name][wait_info][CPU][tempdb_allocations][tempdb_current][blocking_session_id][reads][writes][physical_reads][used_memory][status][open_tran_count][percent_complete][host_name][database_name][program_name][start_time][login_time][request_id][collection_time]'"
            $agHealthQuery = "SELECT ar.replica_server_name, ag.name AS ag_name, DB_NAME(drs.database_id) AS database_name, drs.synchronization_state_desc, drs.is_suspended, drs.log_send_queue_size, drs.redo_queue_size FROM sys.dm_hadr_database_replica_states drs JOIN sys.availability_replicas ar ON drs.replica_id = ar.replica_id JOIN sys.availability_groups ag ON drs.group_id = ag.group_id;"
            $waitStatsQuery = "SELECT * FROM sys.dm_os_wait_stats WHERE wait_type NOT IN ('BROKER_TASK_STOP','BROKER_EVENTHANDLER','BROKER_RECEIVE_WAITFOR','CHECKPOINT_QUEUE','CLR_AUTO_EVENT','DBMIRROR_DBM_EVENT','FT_IFTS_SCHEDULER_IDLE_WAIT','HADR_CLUSAPI_CALL','HADR_TIMER_TASK','LAZYWRITER_SLEEP','LOGMGR_QUEUE','REQUEST_FOR_DEADLOCK_SEARCH','SLEEP_TASK','SQLTRACE_BUFFER_FLUSH','WAITFOR','XE_DISPATCHER_WAIT','XE_TIMER_EVENT')"
            $blockingQuery = "SELECT wt.session_id AS waiting_session_id, wt.blocking_session_id, wt.wait_duration_ms, (SELECT [text] FROM sys.dm_exec_requests r CROSS APPLY sys.dm_exec_sql_text(r.sql_handle) WHERE r.session_id = wt.session_id) AS waiting_query, (SELECT [text] FROM sys.dm_exec_requests r CROSS APPLY sys.dm_exec_sql_text(r.sql_handle) WHERE r.session_id = wt.blocking_session_id) AS blocking_query FROM sys.dm_os_waiting_tasks wt WHERE wt.blocking_session_id IS NOT NULL AND wt.blocking_session_id <> 0;"

            $sample = @{
                SampleNumber = $i; Timestamp = Get-Date -Format 'o'
                WhoIsActive = Get-SqlData -ConnectionString $connectionString -Query $whoIsActiveQuery
                HealthChecks = Get-SqlData -ConnectionString $connectionString -Query $healthCheckQuery
                AGHealth = Get-SqlData -ConnectionString $connectionString -Query $agHealthQuery
                WaitStats = Get-SqlData -ConnectionString $connectionString -Query $waitStatsQuery
                Blocking = Get-SqlData -ConnectionString $connectionString -Query $blockingQuery
            }
            $snapshotData.Samples.Add($sample) | Out-Null
            if ($i -lt $numSamples) { Start-Sleep -Seconds $sampleRate }
        }

        $fileName = "Snapshot-$((Get-Date -Format 'yyyyMMddHHmmss')).json"
        $fullPath = Join-Path -Path $PSScriptRoot -ChildPath $fileName
        $snapshotData | ConvertTo-Json -Depth 10 | Out-File -FilePath $fullPath -Encoding UTF8
        Log-Message "Snapshot capture complete. Data saved to: $fullPath"

        Load-SnapshotData -FilePath $fullPath
    }
    catch { Log-Message "CAPTURE FAILED: $($_.Exception.ToString())" }
    finally {
        $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default
        $captureButton.Enabled = $true
    }
})

$script:loadedSnapshotData = $null
$script:dataGrids = @{}
$script:sortState = @{}

$radio_CheckedChanged = {
    param($sender, $e)
    if ($sender.Checked) {
        $sampleIndex = $sender.Tag
        $selectedSample = $script:loadedSnapshotData.Samples[$sampleIndex]
        $sampleTimestamp = [datetime]$selectedSample.Timestamp | Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $sampleNumber = $selectedSample.SampleNumber

        Log-Message "Loading Sample $sampleNumber (Timestamp: $sampleTimestamp)..."
        $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            $whoIsActiveData = $selectedSample.WhoIsActive
            if ($whoIsActiveData) {
                $blockerIDs = $whoIsActiveData | Where-Object { $_.blocking_session_id } | Select-Object -ExpandProperty blocking_session_id -Unique
                $blockers = New-Object System.Collections.Generic.HashSet[string]; $blockerIDs | ForEach-Object { $blockers.Add($_.ToString()) | Out-Null }
                $chainedBlockers = $whoIsActiveData | Where-Object { $blockers.Contains($_.session_id.ToString()) -and $_.blocking_session_id } | Select-Object -ExpandProperty session_id -Unique
                $chainedBlockerSet = New-Object System.Collections.Generic.HashSet[string]; $chainedBlockers | ForEach-Object { $chainedBlockerSet.Add($_.ToString()) | Out-Null }
                $whoIsActiveData | ForEach-Object {
                    $blockingInfo = ""; $sessionIdStr = $_.session_id.ToString()
                    if ($blockers.Contains($sessionIdStr)) { $blockingInfo = "LEAD" }
                    if ($chainedBlockerSet.Contains($sessionIdStr)) { $blockingInfo += " (CHAIN)" }
                    if ($_.blocking_session_id) {
                        $blockingInfo = "BLOCKED BY $($_.blocking_session_id)"
                        if ($chainedBlockerSet.Contains($_.blocking_session_id.ToString())) { $blockingInfo += " (CHAIN)" }
                    }
                    Add-Member -InputObject $_ -MemberType NoteProperty -Name 'Blocking Info' -Value $blockingInfo -Force
                }
            }

            foreach ($key in $script:dataGrids.Keys) {
                $grid = $script:dataGrids[$key]; $data = $selectedSample.$key
                $bindableData = New-Object System.Collections.ArrayList
                if ($data) { if ($data -is [array]) { $bindableData.AddRange($data) } else { $bindableData.Add($data) } }
                if ($key -eq 'WaitStats') { $bindableData = [System.Collections.ArrayList]($data | Sort-Object -Property 'wait_time_ms' -Descending) }
                $grid.DataSource = $null; $grid.DataSource = $bindableData
            }
        }
        finally {
            Log-Message "Loading of Sample $sampleNumber complete."
            $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    }
}

$On_ColumnHeaderClick = {
    param($sender, $e)
    $grid = $sender
    foreach ($column in $grid.Columns) { if ($column.Index -ne $e.ColumnIndex) { $column.HeaderCell.SortGlyphDirection = 'None' } }
    $sortProperty = $grid.Columns[$e.ColumnIndex].DataPropertyName; if (-not $sortProperty) { return }
    $currentDirection = if ($script:sortState[$grid.Name] -and $script:sortState[$grid.Name].Property -eq $sortProperty) { $script:sortState[$grid.Name].Direction } else { 'Ascending' }
    $newDirection = if ($currentDirection -eq 'Ascending') { 'Descending' } else { 'Ascending' }
    $data = $grid.DataSource | Sort-Object -Property $sortProperty -Descending:($newDirection -eq 'Descending')
    $grid.DataSource = $null; $grid.DataSource = [System.Collections.ArrayList]$data
    $grid.Columns[$e.ColumnIndex].HeaderCell.SortGlyphDirection = $newDirection
    $script:sortState[$grid.Name] = @{ Property = $sortProperty; Direction = $newDirection }
}

$loadSnapshotButton.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.InitialDirectory = $PSScriptRoot
    $openFileDialog.Filter = "JSON Files (*.json)|*.json"
    $openFileDialog.Title = "Select a Snapshot File"
    if ($openFileDialog.ShowDialog() -eq "OK") {
        Load-SnapshotData -FilePath $openFileDialog.FileName
    }
})
#endregion

#region Form Load and Show
$mainForm.Add_Shown({ Load-SavedConnection })
[void]$mainForm.ShowDialog()
$mainForm.Dispose()
#endregion
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
