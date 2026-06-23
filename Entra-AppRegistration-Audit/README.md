# Entra-AppRegistration-Audit

> **One scan for every app registration** — API permission risk (High/Med/Low), credential expiry, staleness, and owners — in a single Excel report with gated cleanup actions.

<sub>[← Back to Griffin31 ToolKit](../) · Cross-platform (Windows · macOS · Linux) · PowerShell 7</sub>

---

## What you get

- **Permission risk per app** — every API permission classified **High / Medium / Low by what it can actually do**, not by how many it has. App's overall risk = its single highest-risk permission.
- **Credential health** — expired and expiring certs/secrets (red ≤7 days, amber ≤14)
- **Staleness** — last sign-in activity → Active / Stale / Never used / No service principal
- **Owners resolved** + tenant-owned vs external
- **Workload identity protection candidates** — flags which apps you can target with Conditional Access for workload identities (single-tenant only) vs ID Protection only (multi-tenant), per Microsoft's eligibility rules
- **KPI dashboard** and **direct Entra portal links** on every row
- **Gated actions** — remove expired creds, disable SP, or delete app (typed-`YES` / `DELETE`)

## Quick start

```powershell
pwsh ./entra-appregistration-audit.ps1
```

Prompts for: admin UPN, expiry threshold (30/60/90/custom), stale threshold (90 default), include Microsoft first-party apps (excluded by default).

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

1. **Setup** — admin UPN, expiry + stale thresholds, first-party include/exclude
2. **Connect** — Microsoft Graph sign-in via browser
3. **Fetch** — app registrations, service principals, `servicePrincipalSignInActivities`
4. **Resolve permissions** — pull each resource API's roles/scopes to turn GUIDs into names
5. **Analyze** — risk per permission, credential buckets, staleness, owners
6. **Review & decide** — export only / remove expired creds / disable SP / delete app
7. **Export** — Excel with portal hyperlinks

## Output

`AppRegistration_Audit_<timestamp>.xlsx` on your Desktop:

- **Summary** — KPI dashboard + permission-risk breakdown
- **Permission Risk** — every app with permissions: overall risk, High/Med/Low counts, sensitive permission list, owner, link
- **Expired Creds** / **Expiring Creds** — color-coded by urgency
- **Stale & Unused** — stale / never-used / orphaned apps with last sign-in
- **Action Log** — appears only when an action runs

## Safety

- **Audit-only by default.**
- **Remove / disable** require typing `YES`. **Delete** requires `YES` then `DELETE`.
- **Disable before delete** — set `accountEnabled=false`, wait, then delete if nothing broke.
- **Microsoft first-party apps excluded by default.**

## Related tools

- [Entra-StaleDevices-Cleanup](../Entra-StaleDevices-Cleanup/) — same audit → disable → delete pattern for devices
- [EXO-AppPermissions-Manager](../EXO-AppPermissions-Manager/) — scope Exchange Online app permissions to mailboxes
