\(•_•)/  "Whattup my Zimba"

 # SharePoint Online → OneDrive Sync

This collection of PowerShell helpers authenticates via Microsoft Graph, enumerates the SharePoint Online sites a user has joined, and launches your configured OneDrive sync command without duplicating existing targets. Use the CLI entry point (`sharepoint_sync.ps1`), the WinForms GUI (`sharepoint_sync_gui.ps1`), or the Datto RMM quick-job wrapper (`datto_quick_sync.ps1`) depending on the audience.

## Contents

- `sharepoint_sync.ps1`: core helper that reads configuration, checks existing syncs, and starts new ones.
- `sharepoint_sync_gui.ps1`: lightweight GUI for end users to pick config files, toggle dry run/verbose, and watch a live log.
- `datto_quick_sync.ps1`: wrapper for component quick jobs; can consume a config file or build one from Datto variables before invoking the core helper.

## Requirements

- Windows 10/11 or PowerShell 7+ host with `Invoke-RestMethod`, `Invoke-Expression`, and access to the OneDrive sync client (`onedrive.exe`, `onedrive-cli`, etc.).
- Registered Azure AD application (`tenant_id`, `client_id`, `client_secret`) granted Graph scopes such as `Sites.Read.All` / `Files.ReadWrite.All`.
- Optional `status_command` that emits currently syncing URLs so duplicates in-flight are skipped.
- OneDrive sync command template that supports `{site_web_url}` and `{local_folder}` placeholders.

## Configuration (`sharepoint_sync_config.json`)

Place this JSON file next to the scripts or generate it dynamically via the Datto wrapper.

```json
{
  "tenant_id": "<GUID>",
  "client_id": "<GUID>",
  "client_secret": "xxxxxxxxxxxxxxxxxxxxxxxxxxxx",
  "scopes": ["https://graph.microsoft.com/.default"],
  "sync_root": "C:\\Users\\Me\\OneDrive - Company\\SharePoint",
  "sync_command_template": "onedrive.exe /sync {local_folder} /sharepoint:{site_web_url}",
  "status_command": "onedrive-cli.exe status",
  "cache_file": ".\\sharepoint_sync_cache.json"
}
```

- `sync_root`: root folder for every synced library; the script creates child folders under this path.
- `sync_command_template`: used verbatim after replacing `{site_web_url}` and `{local_folder}`.
- `status_command`: optional external command (e.g., `onedrive-cli.exe status`) returning URLs to skip duplicates.
- `cache_file`: tracks previously synced site IDs so reruns don’t repeat work.

## CLI Usage

```powershell
.\sharepoint_sync.ps1 -Config ".\sharepoint_sync_config.json" [-DryRun] [-Verbose]
```

- `-DryRun` prints each planned sync without invoking OneDrive.
- `-Verbose` surfaces Graph queries, HTTP responses, and command execution details.
- Output now clearly logs the config/cache paths, status-command findings, discovered sites, and whether new sync jobs were queued.

## GUI Usage

1. Double-click `sharepoint_sync_gui.ps1` or run it from PowerShell in the scripts folder.
2. Browse for the `sharepoint_sync_config.json` JSON file if it lives elsewhere.
3. Toggle **Dry run** or **Verbose logging** before hitting **Run sync**.
4. Monitor the status label and log box for live stdout/stderr from the core helper.

The GUI wraps `sharepoint_sync.ps1` and runs it in the background process, streaming every line into the textarea so non-technical users can trigger syncs without opening PowerShell manually.

## Datto RMM Quick Job

Use `datto_quick_sync.ps1` when deploying as a component or quick job:

### Option A: config file

```powershell
.\datto_quick_sync.ps1 -ConfigPath "C:\path\to\sharepoint_sync_config.json" [-DryRun] [-Verbose]
```

### Option B: inline variables (preferred for Datto placeholders)

```powershell
.\datto_quick_sync.ps1 `
    -TenantId "<GUID>" `
    -ClientId "<GUID>" `
    -ClientSecret "<secret>" `
    -SyncRoot "C:\Users\Me\OneDrive - Company\SharePoint" `
    -SyncCommandTemplate "onedrive.exe /sync {local_folder} /sharepoint:{site_web_url}" `
    -StatusCommand "onedrive-cli.exe status" `
    -CacheFile ".\sharepoint_sync_cache.json"
```

- Validates that `sharepoint_sync.ps1` is alongside the wrapper before launching.
- Builds a temporary config when required variables are provided, invokes the helper with `-NoProfile -ExecutionPolicy Bypass`, streams stdout/stderr, and exits with the same code so Datto reports accurate success/failure.
- Removes the inline config file after completion and logs start/end times plus the active configuration path for troubleshooting.

## Workflow & Idempotency

1. Authenticate with Microsoft Graph via the configured client credentials.
2. Enumerate `/me/joinedSites` (paged) to discover accessible SharePoint document libraries.
3. Merge the optional `status_command` output with the persisted cache to know which URLs OneDrive already tracks.
4. For every uncached/untracked site, replace placeholders in the command template, create the local folder if it doesn’t exist, and invoke the OneDrive sync command.
5. When not in dry run, the cache file is saved after processing so reruns skip already synced targets automatically.

## Next Steps

1. Provide real Azure credentials, Directory IDs, and OneDrive command templates in `sharepoint_sync_config.json` or via the Datto quick job parameters.
2. Test with `-DryRun` (CLI, GUI, or Datto) to confirm output before letting OneDrive modify your sync state.
3. Schedule the script via Task Scheduler, Scheduled Jobs, or as a recurring Datto quick job so the sync root automatically reflects the SharePoint sites the user joins.
