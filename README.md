# SharePoint → Foundry IQ knowledge source sync

An Azure Function (C#, .NET 6, isolated worker) that runs every 15 minutes. It mirrors the files in a
SharePoint document library into the file knowledge source behind your Foundry knowledge base.

| SharePoint state            | What happens in the knowledge source                              |
|-----------------------------|-------------------------------------------------------------------|
| New file                    | Uploaded                                                          |
| File content changed        | **Replaced in place** (same `fileId`, re-extracted and re-embedded) |
| Unchanged                   | Skipped. Compared by the SharePoint `cTag`, which is stored as file metadata |
| Deleted / renamed / moved   | Old copy deleted (only files this function uploaded)              |
| Duplicate names in KS       | Newest kept, extras deleted                                       |

Files are named by their path inside the library (`HR/Policies/handbook.pdf`), so two files with the
same name in different folders don't collide.

## How the pieces fit

A Foundry IQ knowledge base lives on an **Azure AI Search** service. Your knowledge base references one or more
*knowledge sources*. This function writes to a knowledge source of `kind: "file"` through the Search REST API
(`2026-08-01-preview`). An upload is synchronous: when the call returns, the file is already extracted,
chunked, embedded and searchable.

```
SharePoint library ──Graph (app + client secret)──▶ Function ──Search REST (api-key)──▶ file knowledge source ──▶ knowledge base ──▶ agents
```

## Setup

### 1. App registration (SharePoint side)
The app needs **application** permission on Microsoft Graph, with admin consent:
- `Sites.Selected` (recommended), plus a `read` grant on just this site, **or**
- `Sites.Read.All` (reads every site in the tenant)

### 2. Search side
- **Key:** the Azure AI Search service's **admin key**. That service is listed under the knowledge base's
  connection in Foundry. A query key is read-only, and the Foundry project/resource key won't authenticate here.
- **Knowledge source:** check that your knowledge base uses a knowledge source with `"kind": "file"`. If it
  doesn't, create one with `setup/create-file-knowledge-source.http` and add it to the knowledge base in the Foundry portal.

### 3. Configure
Copy `local.settings.sample.json` to `local.settings.json` and fill it in. In Azure, set the same keys as app
settings, with the two secrets as Key Vault references:
```
SharePoint__ClientSecret = @Microsoft.KeyVault(SecretUri=https://<vault>.vault.azure.net/secrets/sp-client-secret/)
Search__ApiKey           = @Microsoft.KeyVault(SecretUri=https://<vault>.vault.azure.net/secrets/search-admin-key/)
```

| Setting | Notes |
|---|---|
| `SharePoint__Hostname` | `contoso.sharepoint.com` |
| `SharePoint__SitePath` | `sites/YourSite` (empty = root site) |
| `SharePoint__LibraryName` | `Documents` is the default "Shared Documents" library. The URL name works too. |
| `SharePoint__FolderPath` | Optional subfolder. Empty = whole library. |
| `Search__KnowledgeSourceName` | The file knowledge source name |
| `Sync__DeleteOrphans` | `true` mirrors deletions. Set `false` to only add/replace. |
| `Sync__MaxConcurrency` | Parallel uploads (default 4). The service returns 429 if this is too high. |
| `Sync__MaxFileSizeMB` | 100 (use 50 on Free/Basic search tiers) |

### 4. Run / deploy
```bash
dotnet build
func start                      # needs Azurite running for AzureWebJobsStorage
func azure functionapp publish <app-name>
```
To trigger a run immediately while testing:
`POST http://localhost:7071/admin/functions/SharePointToKnowledgeSourceSync` with body `{}`.

**Hosting:** .NET 6 is out of support (Nov 2024), and Flex Consumption doesn't offer it. Use a Consumption,
Premium or Dedicated (App Service) plan, with the stack set to .NET 6 (isolated).
`host.json` sets a 10-minute timeout, the maximum on Consumption. The first sync of a large library may need more than one run. That's fine because
each run picks up where the last one stopped.

## Moving to .NET 8/10 later
The code is plain C# 10, so upgrading means changing only `TargetFramework` and the package versions in the
`.csproj`. Optionally switch `Program.cs` to `FunctionsApplication.CreateBuilder` (Worker 2.x).

## Limits to be aware of (file knowledge source, 2026-08-01-preview)
- **200 files** per knowledge source and **100 MB** per file (50 MB on Free/Basic). The function logs
  a warning when the library exceeds this. Beyond that, switch to a blob knowledge source.
- Supported types: PDF, Word, PowerPoint, Excel, JSON, text/Markdown/HTML/CSV. Other types are skipped.
  Images need `standard` extraction (Content Understanding) and are skipped under `minimal`.
- Each upload/update can take up to 180 s of processing. A file that takes longer returns 504 and is retried on the next run.
- **This is a preview API**: no SLA, and request shapes may change before GA.

## Behavior notes
- Change detection uses the SharePoint `cTag`, so metadata-only edits (e.g. a column value) don't
  trigger a re-upload. Only content changes do.
- If SharePoint returns zero eligible files, deletions are skipped. That protects the knowledge base
  if the app's permissions get revoked or the folder path is wrong.
- Any per-file failure is logged, the rest of the run continues, and the invocation is marked failed
  so your App Insights alerts fire. The file is retried next run.
