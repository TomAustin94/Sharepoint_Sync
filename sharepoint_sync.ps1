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
    Write-Error "Configuration file $Config not found."
    exit 1
}

$ConfigData = Get-JsonFile -Path $Config -Default @{ }
$CachePath = $ConfigData.cache_file ?? ".\sharepoint_sync_cache.json"
$Cache = Get-JsonFile -Path $CachePath -Default @{ syncedSites = @() }
$Cache.syncedSites = @($Cache.syncedSites)

$StatusCommand = $ConfigData.status_command
$SyncedUrls = Get-SyncedUrls -StatusCommand $StatusCommand

$Token = Get-AccessToken -TenantId $ConfigData.tenant_id `
    -ClientId $ConfigData.client_id `
    -ClientSecret $ConfigData.client_secret `
    -Scopes ($ConfigData.scopes ?? @("https://graph.microsoft.com/.default"))

$Sites = Get-JoinedSites -AccessToken $Token
$Added = 0

$SyncRootDir = $ConfigData.sync_root
if (-not (Test-Path $SyncRootDir)) {
    New-Item -ItemType Directory -Path $SyncRootDir -Force | Out-Null
}
$SyncRoot = (Get-Item -Path $SyncRootDir).FullName

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

Write-Output "Processed $($Sites.Count) site(s); queued $Added new sync(s)."
