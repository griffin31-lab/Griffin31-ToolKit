# CA-Policy-Analyzer

## Why this tool?

Conditional Access is the front door to your tenant — but most admins can't tell at a glance whether their policies actually cover everyone, whether exclusions are safe, or whether they're missing modern controls like phishing-resistant MFA, token protection, or risk-based policies.

This tool exports your full CA configuration, scores every policy, flags tenant-wide gaps against Microsoft's 2026 best practices, and produces a single self-contained HTML report an admin can scan in under a minute.

Validated against Microsoft Learn (April 2026) — covers the May 13 2026 enforcement change, legacy MFA retirement, and modern controls (CAE, auth context, token protection, AI agent policies, workload identities).

## Requirements

- PowerShell 7.x (Windows or macOS)
- Microsoft.Graph module (auto-installs if missing — pinned to a consistent version to avoid submodule conflicts)
- Permissions: `Policy.Read.All`, `Directory.Read.All`, `Group.Read.All`, `Application.Read.All`

## How it works

The pipeline runs six stages, each writing JSON to `tenants/<domain>/data/`:

1. **Export-Data** — pulls policies, users, guests, groups (incl. nested memberships), service principals, named locations, directory roles
2. **Detect-BreakGlass** — identifies accounts excluded from all enabled policies (your emergency access)
3. **Analyze-NestedGroups** — resolves nested group membership so exclusions are counted correctly
4. **Analyze-PolicyGaps** — scores every policy 0-100 with deductions for weak MFA, broad exclusions, single-user/single-app targeting, platform gaps, location bypass, no-op and stale policies
5. **Analyze-MissingControls** — checks your tenant against 23 modern CA controls (Grant / Session / Scenario), each tagged with required license
6. **Analyze-KeyInsights** — runs 15 priority-sorted checks (lockout risk, break-glass, legacy auth, weak MFA, admin protection, guest coverage, May-2026 enforcement, overlap, stale, 240-policy limit…) and computes a tenant posture score mapped to Microsoft's 3-phase deployment model

Finally, **Generate-Report** builds a single HTML file with sticky sidebar, posture score hero, phase coverage bars, priority-sorted insight cards, missing-controls grouped by license, full policy table (DataTables — sort / filter / search), with direct deep-links to each policy in the Entra portal.

## License badges

Not every recommended control ships with every SKU. Each missing control is tagged so you know what's free on P1 vs. what needs a higher tier:

- **P1** — Entra ID P1 (most controls)
- **P2** — Entra ID P2 (risk-based sign-in, risk-based user, AI agent policies)
- **Purview** — Microsoft Purview (insider risk signals)
- **WID** — Workload Identities license (service principal CA)

## Output

- `tenants/<domain>/data/` — JSON exports and analysis output (one file per stage)
- `tenants/<domain>/reports/CA_Gap_Analysis_<timestamp>.html` — the final report

## Usage

```powershell
pwsh ./CA-Manager.ps1
```

Menu options: (1) Export only, (2) Analyze existing export, (3) Full pipeline (export + analyze + report).

The report is self-contained HTML — open in any browser, share as a file. (CDN assets for icons/fonts/tables load from unpkg/jsdelivr, so for fully-offline viewing keep a browser cached copy.)
