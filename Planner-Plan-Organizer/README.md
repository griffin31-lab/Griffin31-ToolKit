# Planner-Plan-Organizer

> **Organize a Microsoft Planner plan** — sort buckets A-Z, merge duplicate buckets, clean up empty or stale buckets, and export a local JSON backup.

<sub>[← Back to Griffin31 ToolKit](../)</sub>

---

## What you get

- **Interactive menu** — pick a plan, then run actions one at a time
- **Sort buckets A-Z** — server-side Planner re-order using the `orderHint` anchor
- **Delete empty buckets** — after preview + typed `DELETE` confirmation
- **Find duplicate buckets** — normalizes names (case + punctuation) and lets you pick which bucket to keep; moves tasks into it, then optionally deletes the now-empty siblings
- **Stale bucket cleanup** — delete buckets whose last task activity is older than 90 / 180 / 365 days or a custom cutoff
- **Local JSON backup** — full snapshot of plan, buckets, and tasks to your home directory (strongly recommended before any destructive action)
- **Every destructive action** requires typed confirmation (`YES`, `DELETE`, or `MERGE`)

## Quick start

```powershell
pwsh ./planner-organize.ps1
```

Prompts for: admin UPN, then browser sign-in, then pick the target plan.

## Why this tool?

Planner plans accumulate clutter fast — duplicate buckets from different people (`Marketing`, `marketing`, `MARKETING `), empty placeholders, and stale buckets from closed initiatives. The web UI has no bulk sort, no duplicate detection, and no safe merge. This tool does all three against the Graph API, with previews and typed confirmations, so you can tidy a plan in a couple of minutes instead of dragging buckets around by hand.

## Requirements

- PowerShell 7.x (Windows or macOS)
- `Microsoft.Graph` module — auto-installs if missing
- Permissions (delegated): `Tasks.ReadWrite`, `Group.ReadWrite.All`, `User.Read`, `GroupMember.Read.All`
- Account must be a **member of the Microsoft 365 Group that owns the plan** — Planner only grants access via group membership

## How it works

1. **Setup** — admin UPN
2. **Connect** — Microsoft Graph sign-in via browser
3. **Pick plan** — search by name (scans your accessible groups), paste the Planner URL, or enter a plan ID
4. **Load** — fetches plan, buckets, and tasks
5. **Menu** — run any action; each one previews affected items and asks for typed confirmation before writing

### Menu

```
[1] Show overview (buckets + task counts)
[2] Sort buckets A-Z
[3] Delete empty buckets
[4] Find duplicate BUCKETS (merge tasks + delete spare)
[5] List tasks with dates (created / due / completed)
[6] Export backup (JSON of plan, buckets, tasks)
[7] Cleanup stale buckets (delete by last activity date)
[R] Refresh data
[Q] Quit
```

## Output

- **Console only** for most actions (sort / delete / merge preview + result).
- **Backup** — `planner-backup_<plan-title>_<timestamp>.json` in your home directory. Contains the full plan, all buckets, and all tasks, including titles, descriptions, assignees, and dates. **Treat it like sensitive data** — it is never uploaded anywhere.

## Safety

- **Read-only by default.** Every destructive menu item previews affected items first.
- **Typed confirmations** — `YES` to sort, `DELETE` to delete buckets, `MERGE` to merge duplicates.
- **Backup option built in.** The script nags you to run `[6] Export backup` before any destructive action.
- **Planner API limitation** — tasks and buckets don't expose `lastModifiedDateTime`. "Stale" is derived from the most recent `createdDateTime` or `completedDateTime` of the tasks inside each bucket, so a bucket full of very old tasks that nobody has touched may still look "active" if a new task was created recently.

## Privacy

- Nothing is sent anywhere except Microsoft Graph under your delegated token.
- No telemetry, no logging to disk, no network calls outside `graph.microsoft.com`.
- Backup files are local only — the filename and path are printed on export so you know exactly what was written.

## Related tools

- [SharePoint-Sites-Audit](../SharePoint-Sites-Audit/) — same "audit, then optionally act" pattern for SPO / OneDrive / Teams / M365 Groups
- [Entra-StaleApps-Cleanup](../Entra-StaleApps-Cleanup/) — same pattern for stale app registrations
