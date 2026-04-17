# SharePoint-Sites-Audit

## Why this tool?

The SharePoint admin center shows you aggregate numbers — 4,200 sites, 18,000 external shares — but not *which specific sites* are the problem. Finding the 12 publicly-shared sites hiding in a 4,000-site tenant, or the Teams missing sensitivity labels, takes hours of clicking.

This tool iterates every site, OneDrive, M365 group, and Team, runs 14 API-detectable security checks against each one, and produces a self-contained HTML report with a ranked list of entities to fix. Scan defaults to the top 100 sites by storage; full-tenant scan is opt-in.

## What it checks

Per-entity, per-API. Fourteen checks covering four entity types:

**Sites (8 checks)**
- Publicly accessible sites (Anyone-link sharing enabled)
- Sites allowing anonymous access
- Excessive external users on a site
- Site sharing more permissive than tenant baseline
- Inactive sites (no content changes > 365 days)
- Single-owner or no-owner sites
- Site missing a sensitivity label
- Site access granted via direct users rather than groups

**OneDrive (2 checks)**
- OneDrive accounts with excessive external sharing
- OneDrive sharing more permissive than tenant baseline

**Groups / Teams (3 checks)**
- M365 group missing sensitivity label
- Team missing sensitivity label
- Group with guest members and no sensitivity label

**Document libraries (1 check)**
- Default sensitivity label not configured on library

## Requirements

- PowerShell 7.x (Windows, macOS, or Linux)
- `PnP.PowerShell` module (auto-installs if missing) — cross-platform, modern auth
- `Microsoft.Graph` module (auto-installs if missing)
- **Role**: SharePoint Administrator minimum
- **Graph scopes**: `Group.Read.All`, `Directory.Read.All`, `InformationProtectionPolicy.Read`

### First-time setup (ONCE per tenant)

PnP.PowerShell 3.x no longer ships with a shared Microsoft app — every tenant registers its own. The tool does this for you on first run:

1. Launch `pwsh ./SPO-Manager.ps1`
2. Pick menu option **0** ("First-time setup")
3. A browser window opens — sign in as Global Administrator and approve consent
4. Tool saves the generated ClientId to `tenants/<your-domain>/config.json`

After that, options 1-4 work without further prompts. The registered app is **delegated-only** — no client secret, no certificate. Scopes requested: `Group.Read.All`, `Directory.Read.All`, `InformationProtectionPolicy.Read`, `SharePoint AllSites.FullControl`.

If your org doesn't allow non-GA admins to register apps, ask a Global Admin to run option 0 once; the ClientId can then be reused by any delegated user with SharePoint Admin role.

## How it works

The pipeline runs four stages, each writing JSON to `tenants/<domain>/data/`:

1. **Export-Data** — connects to SPO + Graph, pulls site list, OneDrive accounts, groups, teams, assigned labels
2. **Analyze-Sites** — runs the 8 per-site checks; writes `site-findings.json`
3. **Analyze-OneDrive** — runs the 2 per-OneDrive checks; writes `onedrive-findings.json`
4. **Analyze-GroupsTeams** — runs the 4 per-group/team/library checks; writes `group-findings.json`
5. **Analyze-KeyInsights** — computes per-entity scores + tenant roll-up; writes `key-insights.json`
6. **Generate-Report** — produces HTML

Finally, the HTML report includes:
- Tenant posture score (weighted average by entity storage size)
- KPI row (total entities, entities with findings, High/Medium/Low counts)
- Entity table (DataTables — sort/filter/search) with expandable rows showing per-entity findings
- Filter by entity type: Site / OneDrive / Group / Team
- Direct deep-links to each entity in the SharePoint admin center

## Output

- `tenants/<domain>/data/` — JSON exports (one file per stage)
- `tenants/<domain>/reports/SP_Sites_Audit_<timestamp>.html` — the final report

## Run modes

The interactive menu offers:

1. **Export data only** — fetch and store, no analysis
2. **Analyze existing data** — re-run analysis against a prior export
3. **Full pipeline (default: Sample mode)** — top 100 sites by storage + all OneDrives/groups/teams
4. **Full pipeline (Full scan)** — every site. Can take 20-40 minutes on 10k+ site tenants.

## Usage

```powershell
pwsh ./SPO-Manager.ps1
```

## Safety

- **Audit-only.** No remediation. No changes made to your tenant.
- **Read-only Graph scopes.** No `.ReadWrite.` permissions requested.
- **No telemetry.** Data stays on your machine.
- **Tenant data gitignored.** The `tenants/` folder is excluded from version control by default.

## Honest limitations

- **SPO module is Windows-first.** On macOS it works but has device-login quirks. Graph fallbacks preferred where available.
- **Tenant-level config checks out of scope.** For SharePoint tenant-wide settings (custom scripting, default link type, DLP, retention) use the SharePoint admin center or a future dedicated tool.
- **Manual-only items dropped.** Anything not detectable via API (e.g., 3rd-party backup, Purview DLP policies) is intentionally not in this tool's scope.
