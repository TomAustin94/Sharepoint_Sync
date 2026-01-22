    \(•_•)/  "Whattup my Zimba"

 # SharePoint Online → OneDrive Sync

This PowerShell helper authenticates with Microsoft Graph for SharePoint Online, lists every SharePoint Online site the target user has joined, compares those sites against OneDrive’s current sync state/cache, and issues the configured OneDrive client command for any newly discovered libraries so duplicates never get registered.

## Requirements

- Windows 10+ or PowerShell 7+ environment with `Invoke-RestMethod`/`Invoke-Expression`.
- Registered Azure AD application with `client_id`, `client_secret`, and `tenant_id` that has delegated Graph scopes such as `Sites.Read.All`/`Files.ReadWrite.All`.
- OneDrive client installed and the desired sync command template available (`onedrive.exe /sharepoint:<url>` etc.).
  
## Configuration (`sharepoint_sync_config.json`)

```json
{
  "tenant_id": "<GUID>",
  "client_id": "<GUID>",
  "client_secret": "<secret>",
  "scopes": ["https://graph.microsoft.com/.default"],
  "sync_root": "C:\\Users\\Me\\OneDrive - Company\\SharePoint",
  "sync_command_template": "onedrive.exe /sync {local_folder} /sharepoint:{site_web_url}",
  "status_command": "onedrive-cli.exe status",
  "cache_file": ".\\sharepoint_sync_cache.json"
}
```

- `sync_root`: local folder that becomes the root for every library sync.
- `sync_command_template`: supports `{site_web_url}` and `{local_folder}` placeholders.
- `status_command`: optional CLI/XML readout to detect already-synced URLs (used for deduplication).
- `cache_file`: persists previously synced site IDs so re-runs skip known targets.

## Usage

```powershell
.\sharepoint_sync.ps1 -Config ".\sharepoint_sync_config.json" [-DryRun] [-Verbose]
```

- `-DryRun`: prints commands without touching OneDrive.
- `-Verbose`: surfaces diagnostics to help troubleshoot Graph calls or sync failures.
- Output includes how many sites were scanned and how many were queued for sync.

## Flow and Duplication Guard

1. Authenticate via client credentials.
2. Page through `/me/joinedSites` to enumerate accessible SharePoint document libraries.
3. Query the optional status command plus the JSON cache to see if a site is already synced.
4. Run the OneDrive command template for any uncached URL, create the local folder if missing, and append the site entry to the cache.
5. When not in dry run, the cache file saves the new state so future runs skip duplicates automatically.

## Next Steps

1. Fill in the configuration JSON with real credentials and command templates.
2. Run `.\sharepoint_sync.ps1 -DryRun` to verify outputs.
3. Schedule the script (Task Scheduler, Scheduled Jobs) once verified so the local sync root stays aligned with available SharePoint sites.
