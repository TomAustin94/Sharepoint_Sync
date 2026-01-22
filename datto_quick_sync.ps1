<#
.SYNOPSIS
    Wrapper for running the SharePoint sync script via Datto RMM quick job/component.

.DESCRIPTION
    Either consumes an existing `.json` configuration file or builds one from supplied
    variables, forwards the desired flags to `sharepoint_sync.ps1`, captures stdout/stderr,
    and exits with the same code so Datto reports success/failure.

.PARAMETER ConfigPath
    Path to the JSON configuration file to pass to `sharepoint_sync.ps1`. Defaults to
    `sharepoint_sync_config.json` in the same directory.

.PARAMETER TenantId, ClientId, ClientSecret, SyncRoot, SyncCommandTemplate
    Required values when the quick job should synthesize a config file instead of reading an existing one.

.PARAMETER Scopes, StatusCommand, CacheFile
    Optional overrides when supplying inline variables (otherwise use the defaults from `sharepoint_sync_config.json`).

.PARAMETER DryRun
    Forward the dry-run flag (logs actions without invoking the OneDrive client).

.PARAMETER Verbose
    Forward the verbose logging flag to the core script.
>
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot "sharepoint_sync_config.json"),
    [string]$TenantId,
    [string]$ClientId,
    [string]$ClientSecret,
    [string[]]$Scopes = @("https://graph.microsoft.com/.default"),
    [string]$SyncRoot,
    [string]$SyncCommandTemplate,
    [string]$StatusCommand,
    [string]$CacheFile = ".\\sharepoint_sync_cache.json",
    [switch]$DryRun,
    [switch]$Verbose
)

function HasValue($Value) { -not [string]::IsNullOrWhiteSpace($Value) }

$InlineConfigPath = $null
$UseInlineConfig = HasValue $TenantId `
    -and HasValue $ClientId `
    -and HasValue $ClientSecret `
    -and HasValue $SyncRoot `
    -and HasValue $SyncCommandTemplate

Write-Output "Datto quick job starting at $(Get-Date -Format 's')"

if ($UseInlineConfig) {
    $InlineConfigPath = Join-Path $PSScriptRoot ("sharepoint_sync_inline_{0}.json" -f ([Guid]::NewGuid()))
    $ConfigBuilder = @{
        tenant_id             = $TenantId
        client_id             = $ClientId
        client_secret         = $ClientSecret
        sync_root             = $SyncRoot
        sync_command_template = $SyncCommandTemplate
        scopes                = $Scopes
    }
    if (HasValue $StatusCommand) { $ConfigBuilder.status_command = $StatusCommand }
    if (HasValue $CacheFile)     { $ConfigBuilder.cache_file     = $CacheFile }

    $ConfigBuilder | ConvertTo-Json -Depth 4 | Set-Content -Path $InlineConfigPath
    $ConfigPath = $InlineConfigPath
    Write-Output "Built inline configuration at $InlineConfigPath from supplied variables."
} else {
    Write-Output "Using configuration file: $ConfigPath"
}

$SyncScript = Join-Path $PSScriptRoot "sharepoint_sync.ps1"
if (-not (Test-Path $SyncScript)) {
    Write-Error "sharepoint_sync.ps1 not found in $PSScriptRoot. Ensure the script lives alongside this wrapper."
    if ($InlineConfigPath) { Remove-Item -Path $InlineConfigPath -Force }
    exit 1
}

if (-not (Test-Path $ConfigPath)) {
    Write-Error "Configuration path '$ConfigPath' could not be resolved."
    if ($InlineConfigPath) { Remove-Item -Path $InlineConfigPath -Force }
    exit 2
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
if ($DryRun)  { $Arguments += "-DryRun" }
if ($Verbose) { $Arguments += "-Verbose" }

$ProcessInfo = New-Object System.Diagnostics.ProcessStartInfo
$ProcessInfo.FileName = "powershell"
$ProcessInfo.Arguments = ($Arguments -join " ")
$ProcessInfo.RedirectStandardOutput = $true
$ProcessInfo.RedirectStandardError = $true
$ProcessInfo.UseShellExecute = $false
$ProcessInfo.CreateNoWindow = $true

$Process = New-Object System.Diagnostics.Process
$Process.StartInfo = $ProcessInfo
$Process.Start() | Out-Null

$StdOut = $Process.StandardOutput.ReadToEnd()
$StdErr = $Process.StandardError.ReadToEnd()
$Process.WaitForExit()

if ($StdOut) {
    Write-Output $StdOut.Trim()
}
if ($StdErr) {
    Write-Error $StdErr.Trim()
}

Write-Output "Datto quick job completed with exit code $($Process.ExitCode)"

if ($InlineConfigPath -and (Test-Path $InlineConfigPath)) {
    Remove-Item -Path $InlineConfigPath -Force
}

exit $Process.ExitCode
