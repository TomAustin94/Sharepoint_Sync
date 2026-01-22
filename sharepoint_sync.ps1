<#
.SYNOPSIS
    Sync SharePoint Online document libraries to OneDrive using Microsoft Graph.

.DESCRIPTION
    Authenticates with an Azure AD application (client credentials), enumerates
    SharePoint Online sites the signed-in user has joined, skips sites already cached
    or reported by the OneDrive status command, and kicks off the configured
    OneDrive sync template for the remaining sites while keeping a cache.

.PARAMETER Config
    Path to the JSON configuration file describing credentials, sync command,
    and other settings.

.PARAMETER DryRun
    Outputs the actions that would run without invoking the OneDrive client.

.PARAMETER Verbose
    Enables verbose logging via Write-Verbose.
>
param(
    [string]$Config = ".\sharepoint_sync_config.json",
    [switch]$DryRun,
    [switch]$Verbose
)

function Write-LogLine {
    param(
        [string]$Message,
        [ValidateSet("INFO","SUCCESS","WARN","ERROR")]$Level = "INFO"
    )

    $Color = switch ($Level) {
        "SUCCESS" { "Green" }
        "WARN"    { "Yellow" }
        "ERROR"   { "Red" }
        default   { "Cyan" }
    }

    Write-Host "[$Level] $Message" -ForegroundColor $Color
}

function Validate-Config {
    param(
        $ConfigData,
        [string]$Path
    )

    $RequiredKeys = @(
        "tenant_id",
        "client_id",
        "client_secret",
        "sync_command_template",
        "sync_root"
    )

    $Missing = foreach ($Key in $RequiredKeys) {
        $Value = $ConfigData.$Key
        if (-not $Value -or [string]::IsNullOrWhiteSpace("$Value")) {
            $Key
        }
    }

    if ($Missing) {
        $List = $Missing -join ", "
        throw "Configuration file '$Path' is missing the required properties: $List."
    }
}

function Get-JsonFile {
    param([string]$Path, $Default)
    if (Test-Path $Path) {
        return Get-Content -Raw -Path $Path | ConvertFrom-Json
    }
    return $Default
}

function Save-JsonFile {
    param([string]$Path, $Data)
    $Directory = Split-Path -Parent $Path
    if (-not (Test-Path $Directory)) {
        New-Item -ItemType Directory -Path $Directory -Force | Out-Null
    }
    $Data | ConvertTo-Json -Depth 4 | Set-Content -Path $Path
}

function Get-AccessToken {
    param($TenantId, $ClientId, $ClientSecret, $Scopes)
    $Body = @{
        client_id     = $ClientId
        scope         = $Scopes -join " "
        client_secret = $ClientSecret
        grant_type    = "client_credentials"
    }

    $TokenUri = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $Response = Invoke-RestMethod -Method Post -Uri $TokenUri -Body $Body -ContentType "application/x-www-form-urlencoded"
    return $Response.access_token
}

function Get-JoinedSites {
    param($AccessToken)
    $Headers = @{ Authorization = "Bearer $AccessToken" }
    $Uri = "https://graph.microsoft.com/v1.0/me/joinedSites?`$select=id,displayName,webUrl,siteCollection&`$top=999"

    $Sites = @()
    do {
        $Response = Invoke-RestMethod -Uri $Uri -Headers $Headers
        $Sites += $Response.value
        $Uri = $Response.'@odata.nextLink'
    } while ($Uri)

    return $Sites
}

function Get-SyncedUrls {
    param($StatusCommand)
    if ([string]::IsNullOrWhiteSpace($StatusCommand)) { return @() }
    try {
        $Output = Invoke-Expression $StatusCommand 2>&1
        $Pattern = "https?://\S+"
        return [System.Text.RegularExpressions.Regex]::Matches($Output, $Pattern) | ForEach-Object { $_.Value } | Select-Object -Unique
    } catch {
        Write-Warning "Unable to execute status command: $_"
        return @()
    }
}

function Sanitize-FolderName {
    param([string]$Value)
    $Sanitized = $Value -replace '[<>:"/\\|?*]', '_'
    return if ($Sanitized.Trim().Length) { $Sanitized.Trim() } else { "sharepoint_site" }
}

if (-not (Test-Path $Config)) {
    Write-LogLine "Configuration file '$Config' not found. Please create it or pass -Config with a valid path." "ERROR"
    exit 1
}

$ConfigData = Get-JsonFile -Path $Config -Default @{ }
$ConfigData = $ConfigData ?? @{ }
try {
    Validate-Config -ConfigData $ConfigData -Path $Config
} catch {
    Write-LogLine $_.Exception.Message "ERROR"
    exit 1
}

Write-LogLine "Loaded configuration and validated required fields." "SUCCESS"
$CachePath = $ConfigData.cache_file ?? ".\sharepoint_sync_cache.json"
$Cache = Get-JsonFile -Path $CachePath -Default @{ syncedSites = @() }
$Cache.syncedSites = @($Cache.syncedSites)
Write-LogLine "Using cache file at $CachePath" "INFO"

$StatusCommand = $ConfigData.status_command
$SyncedUrls = Get-SyncedUrls -StatusCommand $StatusCommand
if ($StatusCommand) {
    Write-LogLine "Status command executed; detected $($SyncedUrls.Count) currently syncing URL(s)." "INFO"
} else {
    Write-LogLine "No status command configured; all joined sites will be considered for new syncs." "WARN"
}

if ($DryRun) {
    Write-LogLine "Dry run enabled; OneDrive sync commands will only be logged." "WARN"
}

$Token = Get-AccessToken -TenantId $ConfigData.tenant_id `
    -ClientId $ConfigData.client_id `
    -ClientSecret $ConfigData.client_secret `
    -Scopes ($ConfigData.scopes ?? @("https://graph.microsoft.com/.default"))

$Sites = Get-JoinedSites -AccessToken $Token
Write-LogLine "Discovered $($Sites.Count) joined site(s)." "INFO"
$Added = 0

$SyncRootDir = $ConfigData.sync_root
if (-not (Test-Path $SyncRootDir)) {
    New-Item -ItemType Directory -Path $SyncRootDir -Force | Out-Null
}
$SyncRoot = (Get-Item -Path $SyncRootDir).FullName
Write-LogLine "Sync root directory prepared at $SyncRoot" "INFO"

foreach ($Site in $Sites) {
    $WebUrl = $Site.webUrl
    $CachedSite = $Cache.syncedSites | Where-Object { $_.siteId -eq $Site.id }
    if ($CachedSite -or ($SyncedUrls -contains $WebUrl)) {
        Write-Verbose "Skipping already tracked site $WebUrl"
        continue
    }

    $FolderName = Sanitize-FolderName -Value ($Site.displayName ?? $Site.id)
    $LocalPath = Join-Path $SyncRoot $FolderName
    $CommandTemplate = $ConfigData.sync_command_template
    $Command = $CommandTemplate.Replace("{site_web_url}", $WebUrl).Replace("{local_folder}", $LocalPath)

    Write-Output "Syncing $WebUrl to $LocalPath"
    Write-Verbose "Command: $Command"

    if (-not $DryRun) {
        if (-not (Test-Path $LocalPath)) {
            New-Item -ItemType Directory -Path $LocalPath -Force | Out-Null
        }
        try {
            Invoke-Expression $Command
        } catch {
            Write-Warning "Failed to start OneDrive sync for $WebUrl: $_"
            continue
        }
    }

    $Cache.syncedSites += @{
        siteId      = $Site.id
        webUrl      = $WebUrl
        displayName = $Site.displayName
        localPath   = $LocalPath
    }

    $Added++
}

if (-not $DryRun) {
    Save-JsonFile -Path $CachePath -Data $Cache
}

if ($Added -gt 0) {
    Write-LogLine "Queued $Added new sync job(s)." "SUCCESS"
} else {
    Write-LogLine "No new sync jobs were required." "INFO"
}

Write-Output "Processed $($Sites.Count) site(s); queued $Added new sync(s)."
