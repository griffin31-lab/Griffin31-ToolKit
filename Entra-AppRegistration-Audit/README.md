# Entra-AppRegistration-Audit

> **One scan for every app registration _and_ enterprise app** — API permission risk (High/Med/Low), credential expiry, staleness, enabled/disabled state, and owners — in a single Excel + interactive HTML report with gated cleanup actions.

<sub>[← Back to Griffin31 ToolKit](../) · Cross-platform (Windows · macOS · Linux) · PowerShell 7</sub>

---

## What you get

- **App registrations _and_ enterprise apps** — app registrations are audited by the permissions they **request**; enterprise apps (service principals) are audited by the permissions they **actually hold consent to** — including third-party / SaaS / gallery apps that have **no app registration** in your tenant, plus their delegated grants and app-role assignments.
- **Permission risk per app** — every API permission classified **High / Medium / Low by what it can actually do**, not by how many it has. App's overall risk = its single highest-risk permission.
- **Credential health** — expired and expiring certs/secrets (red ≤7 days, amber ≤14), for both app registrations and enterprise apps (SAML signing certs, secrets)
- **Staleness** — last sign-in across **all flows** (user-delegated *and* app-only / daemon, as client or resource) → Active / Stale / Never used. A **Last Flow** column shows which type it last used. Managed identities show **N/A (not tracked)** — their token use largely isn't captured by this report, so they're never falsely flagged "Never used".
- **Enabled / disabled state** — tracked correctly as **two independent states** (Microsoft treats them separately): the **app registration**'s own activation (`isDisabled` — the portal "State: Activated/Deactivated") and the **enterprise app / service principal** sign-in (`accountEnabled` — "Enabled for users to sign-in"). App-registration views show **App Reg** + **SP Sign-in** side by side; enterprise views show the SP state. All **filterable** (Excel AutoFilter + HTML dropdown)
- **Owners resolved** + tenant-owned vs external
- **Workload identity protection candidates** — flags which apps you can target with Conditional Access for workload identities (single-tenant only) vs ID Protection only (multi-tenant), per Microsoft's eligibility rules
- **KPI dashboard** and **direct Entra portal links** on every row
- **Gated actions** — remove expired creds, disable SP, or delete app (typed-`YES` / `DELETE`)

## Quick start

```powershell
pwsh ./entra-appregistration-audit.ps1
```

Prompts for: admin UPN, expiry threshold (30/60/90/custom), stale threshold (90 default). After connecting, it shows **live counts per object type** and lets you multi-select what to audit:

| # | Type | Where it lives in Entra |
| --- | --- | --- |
| 1 | **App registrations** | Entra > **App registrations** (apps built in your tenant) |
| 2 | **Enterprise apps — your org** | Entra > **Enterprise applications** (single-tenant apps your org created) |
| 3 | **Enterprise apps — third-party** | External apps users consented to (**top consent-phishing risk**) |
| 4 | **Enterprise apps — Microsoft** | Microsoft first-party apps (usually noise; slower — a Graph call each) |
| 5 | **Managed identities** | Azure managed identities |

Enter comma-separated choices (e.g. `1,3`). Default `1,2,3`.

## Why this tool?

Credential-expiry and staleness checks answer "are the keys expiring?" and "is the app still used?" — but neither answers the most important question: **what can this app actually access?** A stale app with `Directory.ReadWrite.All` is far more dangerous than one with `User.Read`. This tool adds that missing layer — a real permission-risk classification — and folds credential and staleness checks into the same single scan.

## The permission risk model

Risk is based on **what the permission grants**, with two rules that make it accurate:

| Level | Examples |
| --- | --- |
| 🔴 **High** | `RoleManagement.ReadWrite.Directory`, `Directory.ReadWrite.All`, `Application.ReadWrite.All`, `AppRoleAssignment.ReadWrite.All`, `User.ReadWrite.All`, `Mail.ReadWrite`, `Mail.Send`, `Files.ReadWrite.All`, `Sites.FullControl.All`, `full_access_as_app` |
| 🟡 **Medium** | `Directory.Read.All`, `User.Read.All`, `Group.Read.All`, `Mail.Read`, `Files.Read.All`, `AuditLog.Read.All`, `Policy.Read.All` |
| 🟢 **Low** | `User.Read`, `openid`, `profile`, `email`, `offline_access`, `People.Read` |

- **Application (app-only) > Delegated.** A tenant-wide read (`Mail.Read`, `Directory.Read.All`, `Files.Read.All`, …) granted as an **application** permission runs with no user context and exposes the whole tenant — so it's **escalated to High**.
- **Heuristic fallback.** Permissions not in the explicit lists are still scored: anything matching `*.ReadWrite.All`, `*FullControl*`, `*ReadWrite.Directory`, or `*Manage.All` → High; `*.Read.All` → Medium.

Permission display names are resolved live from each resource API's service principal (`appRoles` + `oauth2PermissionScopes`), so GUIDs become readable names like `Graph: Directory.ReadWrite.All (Application)`.

## Workload identity protection candidates

Each app is marked for workload-identity protection eligibility, per Microsoft's rules:

| Mark | Meaning |
| --- | --- |
| **CA + ID Protection** | Single-tenant SP registered in your tenant — can be targeted by **Conditional Access for workload identities** (block by IP/risk) **and** ID Protection risk detection. Needs Workload ID Premium. |
| **ID Protection only** | Multi-tenant / external SP — eligible for **ID Protection** risk detection, but **not** CA workload-identity policies. |
| **No (Microsoft app)** / **No SP** | Microsoft first-party apps and managed identities are out of scope for both. |

The flag appears as a column on the **Permission Risk** sheet and as a count on the **Summary** sheet, so you can see at a glance which high-risk apps are also CA-protectable.

## Requirements

- PowerShell 7.x (Windows or macOS)
- `Microsoft.Graph` module — auto-installs if missing
- `ImportExcel` module — auto-installs if missing
- **Entra ID P1 or P2** for the sign-in activity report (without it, staleness shows as `Unknown` but permission + credential analysis still works)
- Permissions: `Application.Read.All`, `Directory.Read.All`, `AuditLog.Read.All` (audit). An audit run **never** asks for write access. `Application.ReadWrite.All` is requested **only if** you choose an action (remove / disable / delete) — decline and it stays audit-only with no error.

## How it works

1. **Setup** — admin UPN, expiry + stale thresholds
2. **Connect** — Microsoft Graph sign-in via browser
3. **Fetch** — app registrations, service principals, `servicePrincipalSignInActivities`, `oauth2PermissionGrants` (delegated consent)
4. **Scope** — pick object types to audit, with live counts per type (app regs / enterprise apps by owner / managed identities)
5. **Resolve permissions** — pull each resource API's roles/scopes to turn GUIDs into names
6. **Analyze app registrations** — risk per requested permission, credential buckets, staleness
7. **Analyze enterprise apps** — actual granted permissions (delegated grants + per-SP app-role assignments), SP credentials, enabled state, staleness
8. **Resolve owners**
9. **Review & decide** — export only / remove expired creds / disable SP / delete app
10. **Export** — Excel + interactive HTML, with portal hyperlinks

## Output

`AppRegistration_Audit_<timestamp>.xlsx` **and** `AppRegistration_Audit_<timestamp>.html` on your Desktop:

- **Summary** — KPI dashboard + permission-risk breakdown
- **Permission Risk** — every app registration with permissions: overall risk, High/Med/Low counts, **Enabled** state, sensitive permission list, owner, link
- **Enterprise Apps** — every enterprise app (service principal) in scope, including permission-less / SSO / disabled ones: overall risk, High/Med/Low counts, **App Reg?** (has a local registration or not), type, **Enabled** state, activity, category, granted permission list
- **Expired Creds** / **Expiring Creds** — color-coded by urgency, with a **Source** column (app reg vs enterprise) and **Enabled** state
- **Stale & Unused** — stale / never-used / orphaned apps with last sign-in
- **Action Log** — appears only when an action runs

Every data sheet has **AutoFilter** on, so you can filter by Enabled/Disabled, risk, source, etc. directly. The HTML report is a self-contained interactive view (tabs, search, risk chips, product + enabled/disabled dropdowns).

> **Note:** cleanup actions (remove creds / disable / delete) still operate on **app registrations** only — the enterprise-app view is audit-only.

## Safety

- **Audit-only by default.**
- **Remove / disable** require typing `YES`. **Delete** requires `YES` then `DELETE`.
- **Disable before delete** — set `accountEnabled=false`, wait, then delete if nothing broke.
- **Scope is opt-in** — default audits app registrations + your-org and third-party enterprise apps; Microsoft apps and managed identities are only scanned if you pick them.

## Related tools

- [Entra-StaleDevices-Cleanup](../Entra-StaleDevices-Cleanup/) — same audit → disable → delete pattern for devices
- [EXO-AppPermissions-Manager](../EXO-AppPermissions-Manager/) — scope Exchange Online app permissions to mailboxes
