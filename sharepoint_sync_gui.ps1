Add-Type -AssemblyName System.Windows.Forms, System.Drawing

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$SyncScript = Join-Path $ScriptDir "sharepoint_sync.ps1"

if (-not (Test-Path $SyncScript)) {
    [System.Windows.Forms.MessageBox]::Show(
        "Unable to find sharepoint_sync.ps1 in the same directory as this GUI.",
        "Missing Script",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
    exit 1
}

$Form = New-Object System.Windows.Forms.Form
$Form.Text = "SharePoint Sync Helper"
$Form.Size = New-Object System.Drawing.Size 720, 520
$Form.StartPosition = "CenterScreen"

$Label = New-Object System.Windows.Forms.Label
$Label.Text = "Configuration file"
$Label.AutoSize = $true
$Label.Location = New-Object System.Drawing.Point 12, 18
$Form.Controls.Add($Label)

$ConfigPathBox = New-Object System.Windows.Forms.TextBox
$ConfigPathBox.Text = Join-Path $ScriptDir "sharepoint_sync_config.json"
$ConfigPathBox.Size = New-Object System.Drawing.Size 540, 24
$ConfigPathBox.Location = New-Object System.Drawing.Point 12, 40
$Form.Controls.Add($ConfigPathBox)

$BrowseButton = New-Object System.Windows.Forms.Button
$BrowseButton.Text = "Browse..."
$BrowseButton.Size = New-Object System.Drawing.Size 100, 26
$BrowseButton.Location = New-Object System.Drawing.Point 564, 38
$BrowseButton.Add_Click({
    $Dialog = New-Object System.Windows.Forms.OpenFileDialog
    $Dialog.Filter = "JSON config|*.json|All files|*.*"
    $Dialog.FileName = $ConfigPathBox.Text
    if ($Dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $ConfigPathBox.Text = $Dialog.FileName
    }
})
$Form.Controls.Add($BrowseButton)

$DryRunCheck = New-Object System.Windows.Forms.CheckBox
$DryRunCheck.Text = "Dry run (show actions without running OneDrive)"
$DryRunCheck.AutoSize = $true
$DryRunCheck.Location = New-Object System.Drawing.Point 12, 80
$Form.Controls.Add($DryRunCheck)

$VerboseCheck = New-Object System.Windows.Forms.CheckBox
$VerboseCheck.Text = "Verbose logging"
$VerboseCheck.AutoSize = $true
$VerboseCheck.Location = New-Object System.Drawing.Point 12, 108
$Form.Controls.Add($VerboseCheck)

$RunButton = New-Object System.Windows.Forms.Button
$RunButton.Text = "Run sync"
$RunButton.Size = New-Object System.Drawing.Size 150, 30
$RunButton.Location = New-Object System.Drawing.Point 12, 140

$StatusLabel = New-Object System.Windows.Forms.Label
$StatusLabel.Text = "Ready"
$StatusLabel.AutoSize = $true
$StatusLabel.Location = New-Object System.Drawing.Point 180, 148
$Form.Controls.Add($StatusLabel)

$OutputBox = New-Object System.Windows.Forms.TextBox
$OutputBox.Multiline = $true
$OutputBox.ReadOnly = $true
$OutputBox.ScrollBars = "Vertical"
$OutputBox.Location = New-Object System.Drawing.Point 12, 190
$OutputBox.Size = New-Object System.Drawing.Size 684, 260
$OutputBox.Anchor = [System.Windows.Forms.AnchorStyles] "Top, Bottom, Left, Right"
$Form.Controls.Add($OutputBox)

function Append-Log {
    param([string]$Message)

    if (-not $Message) { return }

    $UpdateBlock = {
        $OutputBox.AppendText("$Message`r`n")
        $OutputBox.SelectionStart = $OutputBox.Text.Length
        $OutputBox.ScrollToCaret()
    }

    if ($OutputBox.InvokeRequired) {
        $OutputBox.Invoke([System.Action]$UpdateBlock)
    } else {
        & $UpdateBlock
    }
}

$CloseButton = New-Object System.Windows.Forms.Button
$CloseButton.Text = "Close"
$CloseButton.Size = New-Object System.Drawing.Size 80, 30
$CloseButton.Location = New-Object System.Drawing.Point 616, 460
$CloseButton.Add_Click({ $Form.Close() })
$Form.Controls.Add($CloseButton)

$RunButton.Add_Click({
    $ConfigPath = $ConfigPathBox.Text.Trim()

    if (-not (Test-Path $ConfigPath)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please pick a valid configuration file before running the sync.",
            "Missing Configuration",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    $Arguments = @(
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        "`"$SyncScript`"",
        "-Config",
        "`"$ConfigPath`""
    )
    if ($DryRunCheck.Checked) {
        $Arguments += "-DryRun"
    }
    if ($VerboseCheck.Checked) {
        $Arguments += "-Verbose"
    }

    $ProcessInfo = New-Object System.Diagnostics.ProcessStartInfo
    $ProcessInfo.FileName = "powershell"
    $ProcessInfo.Arguments = ($Arguments -join " ")
    $ProcessInfo.RedirectStandardOutput = $true
    $ProcessInfo.RedirectStandardError = $true
    $ProcessInfo.UseShellExecute = $false
    $ProcessInfo.CreateNoWindow = $true

    $Process = New-Object System.Diagnostics.Process
    $Process.StartInfo = $ProcessInfo
    $Process.EnableRaisingEvents = $true

    $RunButton.Enabled = $false
    $StatusLabel.Text = "Running..."
    $OutputBox.Clear()
    Append-Log "Starting sharepoint_sync.ps1..."

    $Process.add_OutputDataReceived({
        if ($_.Data) {
            Append-Log $_.Data
        }
    })
    $Process.add_ErrorDataReceived({
        if ($_.Data) {
            Append-Log $_.Data
        }
    })

    $Process.add_Exited({
        if ($OutputBox.InvokeRequired) {
            $OutputBox.Invoke([System.Action]{
                $StatusLabel.Text = "Idle"
                $RunButton.Enabled = $true
            })
        } else {
            $StatusLabel.Text = "Idle"
            $RunButton.Enabled = $true
        }
        Append-Log "Process finished with exit code $($Process.ExitCode)."
    })

    $Process.Start() | Out-Null
    $Process.BeginOutputReadLine()
    $Process.BeginErrorReadLine()
})

$Form.Controls.Add($RunButton)

[System.Windows.Forms.Application]::Run($Form)
