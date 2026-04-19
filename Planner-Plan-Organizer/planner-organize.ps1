# ==============================
# Microsoft Planner - Plan Organizer
# Sort buckets A-Z, merge duplicates, delete empty/stale buckets,
# and export a local JSON backup of any Planner plan.
#
# Interactive delegated sign-in via Entra ID (browser).
# Single-file script. Supports PowerShell 7.x on Windows and macOS.
# ==============================

$ErrorActionPreference = "Stop"

# ── PS7 check ──
if ($PSVersionTable.PSVersion.Major -lt 7) {
  Write-Host ""
  Write-Host "  [!] This script requires PowerShell 7 or later." -ForegroundColor Red
  if ($IsWindows -or $env:OS -match "Windows") {
    Write-Host "  Install:  https://aka.ms/install-powershell" -ForegroundColor Yellow
  } else {
    Write-Host "  Install:  brew install powershell/tap/powershell" -ForegroundColor Yellow
  }
  return
}

# ── Microsoft.Graph module ──
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
  Write-Host ""
  Write-Host "  [!] Microsoft.Graph module is not installed." -ForegroundColor Red
  $c = Read-Host "  Install now? (Y/n)"
  if ($c -eq 'n' -or $c -eq 'N') { return }
  Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
}

$banner = @"

  ============================================================
     Microsoft Planner - Organize Plan
     Sort / Cleanup / Merge Duplicates / Backup
  ============================================================

"@
Write-Host $banner -ForegroundColor Cyan

# ── Step 1: Admin UPN ──
Write-Host "  This script connects via delegated Graph permissions." -ForegroundColor Gray
Write-Host "  You will be prompted to sign in with a browser." -ForegroundColor Gray
Write-Host ""

$AdminUPN = Read-Host "  Enter admin UPN (e.g. admin@contoso.onmicrosoft.com)"
if (-not $AdminUPN -or $AdminUPN.Trim().Length -eq 0) {
  Write-Host "  [!] No UPN provided. Exiting." -ForegroundColor Red
  return
}
$AdminUPN = $AdminUPN.Trim()
if ($AdminUPN -notmatch '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9][a-zA-Z0-9.-]{0,252}\.[a-zA-Z]{2,63}$') {
  Write-Host "  [!] '$AdminUPN' doesn't look like a valid UPN." -ForegroundColor Red
  return
}
$upnDomain = ($AdminUPN -split '@')[-1]
Write-Host "  Tenant domain: " -NoNewline; Write-Host $upnDomain -ForegroundColor Yellow

# ── Step 2: Connect (before plan search - we need to query accessible plans) ──
Write-Host ""
Write-Host "  Connecting to Microsoft Graph (browser sign-in)..." -ForegroundColor Cyan
$scopes = @("Tasks.ReadWrite", "Group.ReadWrite.All", "User.Read", "GroupMember.Read.All")
try {
  Connect-MgGraph -Scopes $scopes -NoWelcome
  $ctx = Get-MgContext
  if (-not $ctx) { Write-Host "  [!] No Graph context." -ForegroundColor Red; return }
  Write-Host "  Connected as: " -NoNewline; Write-Host $ctx.Account -ForegroundColor Green
  Write-Host "  Tenant:       " -NoNewline; Write-Host $ctx.TenantId -ForegroundColor Green
  $TenantId = $ctx.TenantId
  if ($ctx.Account -and $ctx.Account -ne $AdminUPN) {
    Write-Host ""
    Write-Host "  [!] Signed in as $($ctx.Account), but you entered $AdminUPN." -ForegroundColor Yellow
    $ok = Read-Host "  Continue anyway? (y/N)"
    if ($ok -ne 'y' -and $ok -ne 'Y') { Disconnect-MgGraph | Out-Null; return }
  }
} catch {
  Write-Host "  [!] Failed to connect: $($_.Exception.Message)" -ForegroundColor Red
  return
}

# ── Helpers (required early for plan search) ──
function Invoke-GraphGetEarly {
  param([string]$Uri)
  $list = [System.Collections.Generic.List[object]]::new()
  $next = $Uri
  while ($next) {
    $resp = Invoke-MgGraphRequest -Method GET -Uri $next -ErrorAction Stop
    if ($resp -is [hashtable] -and $resp.ContainsKey('value')) {
      if ($resp.value) { foreach ($v in $resp.value) { [void]$list.Add($v) } }
      $next = $resp.'@odata.nextLink'
    } else {
      return $resp
    }
  }
  return $list.ToArray()
}

# ── Step 3: Select plan ──
Write-Host ""
Write-Host "  Select target plan:" -ForegroundColor White
Write-Host "    [1] Search by plan name (recommended)" -ForegroundColor Gray
Write-Host "    [2] Paste Planner URL" -ForegroundColor Gray
Write-Host "    [3] Enter plan ID directly" -ForegroundColor Gray
$method = Read-Host "  Choose (1-3) [default: 1]"

$PlanId = $null

if ($method -eq '2') {
  $url = (Read-Host "  Paste Planner URL").Trim()
  if ($url -match '/plan/([A-Za-z0-9_\-]+)') { $PlanId = $Matches[1] }
  if (-not $PlanId) { Write-Host "  [!] Could not parse plan ID from URL." -ForegroundColor Red; Disconnect-MgGraph | Out-Null; return }
} elseif ($method -eq '3') {
  $PlanId = (Read-Host "  Enter plan ID").Trim()
  if (-not $PlanId) { Write-Host "  [!] No plan ID." -ForegroundColor Red; Disconnect-MgGraph | Out-Null; return }
} else {
  $query = (Read-Host "  Enter plan name (full or partial, case-insensitive)").Trim()
  if (-not $query) { Write-Host "  [!] No search term." -ForegroundColor Red; Disconnect-MgGraph | Out-Null; return }

  Write-Host ""
  Write-Host "  Searching accessible plans..." -ForegroundColor Cyan

  # Gather plans from user's groups
  $allPlans = [System.Collections.Generic.List[object]]::new()
  $seen = @{}

  try {
    $groups = @(Invoke-GraphGetEarly "https://graph.microsoft.com/v1.0/me/memberOf/microsoft.graph.group?`$select=id,displayName&`$top=999")
    Write-Host "    - scanning $($groups.Count) group(s)..." -ForegroundColor DarkGray
    $gNum = 0
    foreach ($g in $groups) {
      $gNum++
      Write-Progress -Activity "Scanning groups" -Status "$gNum / $($groups.Count): $($g.displayName)" -PercentComplete (($gNum / [Math]::Max(1,$groups.Count)) * 100)
      try {
        $plans = @(Invoke-GraphGetEarly "https://graph.microsoft.com/v1.0/groups/$($g.id)/planner/plans")
        foreach ($p in $plans) {
          if (-not $seen.ContainsKey($p.id)) {
            $seen[$p.id] = $true
            [void]$allPlans.Add([pscustomobject]@{
              Id        = $p.id
              Title     = $p.title
              GroupName = $g.displayName
              GroupId   = $g.id
            })
          }
        }
      } catch {
        # group may not have planner - ignore
      }
    }
    Write-Progress -Activity "Scanning groups" -Completed
  } catch {
    Write-Host "  [!] Could not list groups: $($_.Exception.Message)" -ForegroundColor Red
    Disconnect-MgGraph | Out-Null; return
  }

  $found = @($allPlans | Where-Object { $_.Title -and $_.Title.ToLower().Contains($query.ToLower()) })
  if ($found.Count -eq 0) {
    Write-Host "  [!] No plans found matching '$query'." -ForegroundColor Red
    Write-Host "      Scanned $($allPlans.Count) plans across $($groups.Count) groups." -ForegroundColor DarkGray
    Disconnect-MgGraph | Out-Null; return
  }

  Write-Host ""
  Write-Host "  Found $($found.Count) matching plan(s):" -ForegroundColor Green
  for ($k = 0; $k -lt $found.Count; $k++) {
    $m = $found[$k]
    Write-Host ("    [{0,2}] {1}   (group: {2})" -f ($k+1), $m.Title, $m.GroupName) -ForegroundColor White
  }

  Write-Host ""
  $sel = Read-Host "  Select plan number (1-$($found.Count))"
  if (-not ($sel -match '^\d+$') -or [int]$sel -lt 1 -or [int]$sel -gt $found.Count) {
    Write-Host "  [!] Invalid selection." -ForegroundColor Red
    Disconnect-MgGraph | Out-Null; return
  }
  $chosen = $found[[int]$sel - 1]
  $PlanId = $chosen.Id
  Write-Host ""
  Write-Host "  Selected: " -NoNewline; Write-Host $chosen.Title -ForegroundColor Green
  Write-Host "  Group:    " -NoNewline; Write-Host $chosen.GroupName -ForegroundColor Green
}

Write-Host ""
Write-Host "  Plan ID:   " -NoNewline; Write-Host $PlanId   -ForegroundColor Yellow
Write-Host "  Tenant ID: " -NoNewline; Write-Host $TenantId -ForegroundColor Yellow

# ── Helpers ──
function Invoke-GraphGet {
  param([string]$Uri)
  $next = $Uri
  $list = [System.Collections.Generic.List[object]]::new()
  while ($next) {
    $resp = Invoke-MgGraphRequest -Method GET -Uri $next -ErrorAction Stop
    if ($resp -is [hashtable] -and $resp.ContainsKey('value')) {
      if ($resp.value) {
        foreach ($v in $resp.value) { [void]$list.Add($v) }
      }
      $next = $resp.'@odata.nextLink'
    } else {
      return $resp
    }
  }
  return $list.ToArray()
}

function Get-ResourceEtag {
  param([string]$Uri)
  $r = Invoke-MgGraphRequest -Method GET -Uri $Uri -ErrorAction Stop
  return $r.'@odata.etag'
}

function Load-PlanData {
  Write-Host "    - plan..." -ForegroundColor DarkGray
  $script:Plan = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/planner/plans/$PlanId" -ErrorAction Stop
  Write-Host "    - buckets..." -ForegroundColor DarkGray
  $script:Buckets = @(Invoke-GraphGet "https://graph.microsoft.com/v1.0/planner/plans/$PlanId/buckets")
  Write-Host "    - tasks..." -ForegroundColor DarkGray
  $script:Tasks = @(Invoke-GraphGet "https://graph.microsoft.com/v1.0/planner/plans/$PlanId/tasks")

  $script:CountByBucket = @{}
  foreach ($b in $script:Buckets) {
    if ($b.id) { $script:CountByBucket[$b.id] = 0 }
  }
  foreach ($t in $script:Tasks) {
    if ($t.bucketId -and $script:CountByBucket.ContainsKey($t.bucketId)) {
      $script:CountByBucket[$t.bucketId]++
    }
  }
  $script:BucketName = @{}
  foreach ($b in $script:Buckets) {
    if ($b.id) { $script:BucketName[$b.id] = $b.name }
  }
}

function Fmt-Date {
  param($d)
  if (-not $d) { return "-" }
  try { return ([datetime]$d).ToString("yyyy-MM-dd") } catch { return "$d" }
}

# ── Load once ──
Write-Host ""
Write-Host "  Loading plan data..." -ForegroundColor Cyan
try {
  Load-PlanData
} catch {
  Write-Host ""
  Write-Host "  [!] Failed to load plan data." -ForegroundColor Red
  Write-Host "      $($_.Exception.Message)" -ForegroundColor Yellow
  if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
    Write-Host "      $($_.ErrorDetails.Message)" -ForegroundColor DarkYellow
  }
  Write-Host ""
  Write-Host "  Common causes:" -ForegroundColor Gray
  Write-Host "    - Your account is not a member of the group that owns this plan." -ForegroundColor Gray
  Write-Host "    - The consent did not include Tasks.ReadWrite / Group.ReadWrite.All." -ForegroundColor Gray
  Write-Host "    - The plan ID in the URL is wrong or the plan was deleted." -ForegroundColor Gray
  Disconnect-MgGraph | Out-Null
  return
}
Write-Host "  Plan:    " -NoNewline; Write-Host $script:Plan.title -ForegroundColor Green
Write-Host "  Buckets: " -NoNewline; Write-Host $script:Buckets.Count -ForegroundColor Green
Write-Host "  Tasks:   " -NoNewline; Write-Host $script:Tasks.Count -ForegroundColor Green

# ── Actions ──

function Get-BucketCount {
  param($BucketId)
  if ($BucketId -and $script:CountByBucket.ContainsKey($BucketId)) {
    return [int]$script:CountByBucket[$BucketId]
  }
  return 0
}

function Get-BucketName {
  param($BucketId)
  if ($BucketId -and $script:BucketName.ContainsKey($BucketId)) {
    $n = $script:BucketName[$BucketId]
    if ($n) { return $n }
  }
  return "(unknown)"
}

function Show-Overview {
  Write-Host ""
  if ($script:Buckets.Count -eq 0) {
    Write-Host "  No buckets in this plan." -ForegroundColor Yellow
    return
  }
  Write-Host "  Overview:" -ForegroundColor White
  $sorted = @($script:Buckets | Sort-Object { if ($_.name) { $_.name.ToLower() } else { "" } })
  foreach ($b in $sorted) {
    $c = Get-BucketCount $b.id
    $color = if ($c -eq 0) { "Red" } else { "Gray" }
    $nm = if ($b.name) { $b.name } else { "(no name)" }
    Write-Host ("    - {0,-40}  {1} task(s)" -f $nm, $c) -ForegroundColor $color
  }
  Write-Host ""
  $emptyCount = @($script:Buckets | Where-Object { (Get-BucketCount $_.id) -eq 0 }).Count
  Write-Host "  Empty buckets: $emptyCount" -ForegroundColor White
}

function Sort-Buckets {
  if ($script:Buckets.Count -lt 2) {
    Write-Host ""
    Write-Host "  Only $($script:Buckets.Count) bucket(s) - nothing to sort." -ForegroundColor Green
    return
  }
  # Use ordinal string comparison to match Planner's server-side sort exactly.
  $arr = [object[]]($script:Buckets)
  [array]::Sort($arr, [Comparison[object]]{ param($a,$b) [string]::CompareOrdinal([string]$a.orderHint, [string]$b.orderHint) })
  $current = @($arr)
  $sorted  = @($script:Buckets | Sort-Object { if ($_.name) { $_.name.ToLower() } else { "" } })

  Write-Host ""
  Write-Host "  First 5 buckets by orderHint (what Planner actually shows at top):" -ForegroundColor DarkGray
  for ($k = 0; $k -lt [Math]::Min(5, $current.Count); $k++) {
    Write-Host ("    {0,3}. {1,-30}  hint: {2}" -f ($k+1), $current[$k].name, $current[$k].orderHint) -ForegroundColor DarkGray
  }

  Write-Host ""
  Write-Host "  Preview - current order  ->  new A-Z order:" -ForegroundColor White
  Write-Host ("  {0,-3} {1,-38}  {2,-3} {3}" -f "#","Current","#","New (A-Z)") -ForegroundColor DarkGray
  Write-Host ("  {0,-3} {1,-38}  {2,-3} {3}" -f "--","-------","--","---------") -ForegroundColor DarkGray
  for ($i = 0; $i -lt $sorted.Count; $i++) {
    $cur = if ($i -lt $current.Count) { $current[$i].name } else { "" }
    $new = $sorted[$i].name
    $moved = ($cur -ne $new)
    $color = if ($moved) { "Yellow" } else { "Gray" }
    Write-Host ("  {0,-3} {1,-38}  {2,-3} {3}" -f ($i+1), $cur, ($i+1), $new) -ForegroundColor $color
  }
  $moves = 0
  for ($i = 0; $i -lt $sorted.Count; $i++) {
    if ($i -lt $current.Count -and $current[$i].name -ne $sorted[$i].name) { $moves++ }
  }
  Write-Host ""
  Write-Host "  $moves bucket(s) will change position (total $($sorted.Count))." -ForegroundColor White
  if ($moves -eq 0) {
    Write-Host "  Already sorted according to orderHint values above." -ForegroundColor Green
    Write-Host ""
    $f = Read-Host "  Force re-apply A-Z anyway? (y/N)"
    if ($f -ne 'y' -and $f -ne 'Y') { return }
  } else {
    Write-Host ""
    $c = Read-Host "  Type 'YES' to apply new order"
    if ($c -ne 'YES') { Write-Host "  Cancelled." -ForegroundColor Yellow; return }
  }

  Write-Host ""
  Write-Host "  Applying order..." -ForegroundColor Cyan
  Write-Host "  (Using Planner's ' !' anchor - successive moves preserve processing order.)" -ForegroundColor DarkGray
  # Process in alphabetical order. Setting orderHint to " !" assigns each bucket
  # a hint that sorts AFTER previous " !" calls, so processing order = final order.
  $sortedAsc = @($script:Buckets | Sort-Object { if ($_.name) { $_.name.ToLower() } else { "" } })
  $n = $sortedAsc.Count
  $i = 0
  $ok = 0; $fail = 0
  foreach ($b in $sortedAsc) {
    $i++
    $finalPos = $i
    $hint = " !"
    try {
      $etag = Get-ResourceEtag "https://graph.microsoft.com/v1.0/planner/buckets/$($b.id)"
      $body = @{ orderHint = $hint } | ConvertTo-Json -Compress
      Invoke-MgGraphRequest -Method PATCH `
        -Uri "https://graph.microsoft.com/v1.0/planner/buckets/$($b.id)" `
        -Headers @{ "If-Match" = $etag } `
        -ContentType "application/json" `
        -Body $body -ErrorAction Stop
      Write-Host ("    [{0,3}] {1}" -f $finalPos, $b.name) -ForegroundColor Green
      $ok++
    } catch {
      $msg = $_.Exception.Message
      if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $msg = $_.ErrorDetails.Message }
      Write-Host ("    [fail] {0} - {1}" -f $b.name, $msg) -ForegroundColor Red
      $fail++
    }
  }
  Write-Host ""
  Write-Host "  Sort complete. OK: $ok   Failed: $fail" -ForegroundColor White
  Load-PlanData
}

function Remove-EmptyBuckets {
  $empty = @($script:Buckets | Where-Object { (Get-BucketCount $_.id) -eq 0 })
  Write-Host ""
  if ($empty.Count -eq 0) {
    Write-Host "  No empty buckets." -ForegroundColor Green
    return
  }
  Write-Host "  Empty buckets to delete:" -ForegroundColor White
  foreach ($b in $empty) { Write-Host "    - $($b.name)" -ForegroundColor Red }

  Write-Host ""
  $c = Read-Host "  Type 'DELETE' to remove them"
  if ($c -ne 'DELETE') { Write-Host "  Cancelled." -ForegroundColor Yellow; return }

  foreach ($b in $empty) {
    try {
      $etag = Get-ResourceEtag "https://graph.microsoft.com/v1.0/planner/buckets/$($b.id)"
      Invoke-MgGraphRequest -Method DELETE `
        -Uri "https://graph.microsoft.com/v1.0/planner/buckets/$($b.id)" `
        -Headers @{ "If-Match" = $etag } -ErrorAction Stop
      Write-Host "    [deleted] $($b.name)" -ForegroundColor Green
    } catch {
      Write-Host "    [fail]    $($b.name) - $($_.Exception.Message)" -ForegroundColor Red
    }
  }
  Load-PlanData
}

function Normalize-Name {
  param([string]$s)
  if (-not $s) { return "" }
  return ($s -replace '[^a-zA-Z0-9]', '').ToLower()
}

function Move-TaskToBucket {
  param([string]$TaskId, [string]$TargetBucketId)
  $etag = Get-ResourceEtag "https://graph.microsoft.com/v1.0/planner/tasks/$TaskId"
  $body = @{ bucketId = $TargetBucketId; orderHint = " !" } | ConvertTo-Json -Compress
  Invoke-MgGraphRequest -Method PATCH `
    -Uri "https://graph.microsoft.com/v1.0/planner/tasks/$TaskId" `
    -Headers @{ "If-Match" = $etag } `
    -ContentType "application/json" `
    -Body $body -ErrorAction Stop
}

function Find-DuplicateBuckets {
  if ($script:Buckets.Count -lt 2) {
    Write-Host ""
    Write-Host "  Need at least 2 buckets." -ForegroundColor Green
    return
  }
  Write-Host ""
  Write-Host "  Finding duplicate/similar bucket names..." -ForegroundColor Cyan
  Write-Host "  Matching rule: lowercase + strip all non-alphanumeric." -ForegroundColor DarkGray
  Write-Host "  (Examples that match: 'Alpha' = 'ALPHA', 'Project X' = 'projectx')" -ForegroundColor DarkGray

  $groups = @($script:Buckets |
    Where-Object { $_.name -and $_.name.Trim().Length -gt 0 } |
    Group-Object { Normalize-Name $_.name } |
    Where-Object { $_.Count -gt 1 })

  if ($groups.Count -eq 0) {
    Write-Host ""
    Write-Host "  No duplicate buckets found." -ForegroundColor Green
    return
  }

  Write-Host ""
  Write-Host "  Found $($groups.Count) duplicate group(s):" -ForegroundColor Yellow
  $gNum = 0
  foreach ($g in $groups) {
    $gNum++
    $totalTasks = 0
    foreach ($m in $g.Group) { $totalTasks += (Get-BucketCount $m.id) }
    Write-Host ""
    Write-Host ("  Group $gNum - normalized '$($g.Name)'  ($($g.Count) buckets, $totalTasks task(s) total)") -ForegroundColor White
    foreach ($m in $g.Group) {
      $c = Get-BucketCount $m.id
      Write-Host ("      - {0,-40}  {1} task(s)" -f $m.name, $c) -ForegroundColor Gray
    }
  }

  Write-Host ""
  Write-Host "  IMPORTANT: consider running [6] Export backup before merging." -ForegroundColor Yellow
  Write-Host "  You will be asked per group which bucket to keep (or skip)." -ForegroundColor Gray
  Write-Host ""
  $go = Read-Host "  Continue to merge flow? (y/N)"
  if ($go -ne 'y' -and $go -ne 'Y') { Write-Host "  Cancelled." -ForegroundColor Yellow; return }

  $gIdx = 0
  foreach ($g in $groups) {
    $gIdx++
    Write-Host ""
    Write-Host "  ---- Group $gIdx of $($groups.Count) ----" -ForegroundColor Cyan
    Write-Host "  Buckets sharing normalized name '$($g.Name)':" -ForegroundColor White
    $members = @($g.Group)
    for ($k = 0; $k -lt $members.Count; $k++) {
      $b = $members[$k]
      $c = Get-BucketCount $b.id
      Write-Host ("    [{0}]  {1,-40}  {2} task(s)" -f ($k+1), $b.name, $c) -ForegroundColor Gray
    }
    Write-Host ""
    Write-Host "  Choose the bucket to KEEP (tasks from the others will move here)." -ForegroundColor White
    $pick = Read-Host "  Target bucket number (1-$($members.Count)), or S to skip this group"
    if ($pick -eq 's' -or $pick -eq 'S') { Write-Host "  Skipped." -ForegroundColor Yellow; continue }
    if (-not ($pick -match '^\d+$') -or [int]$pick -lt 1 -or [int]$pick -gt $members.Count) {
      Write-Host "  Invalid selection. Skipping." -ForegroundColor Yellow
      continue
    }
    $targetIdx = [int]$pick - 1
    $target = $members[$targetIdx]
    $sources = @($members | Where-Object { $_.id -ne $target.id })

    $totalTasksToMove = 0
    foreach ($s in $sources) { $totalTasksToMove += (Get-BucketCount $s.id) }
    Write-Host ""
    Write-Host "  Will move $totalTasksToMove task(s) into: $($target.name)" -ForegroundColor White
    Write-Host "  From:" -ForegroundColor White
    foreach ($s in $sources) {
      Write-Host ("    - {0}  ({1} task(s))" -f $s.name, (Get-BucketCount $s.id)) -ForegroundColor Gray
    }
    $conf = Read-Host "  Type 'MERGE' to proceed"
    if ($conf -ne 'MERGE') { Write-Host "  Skipped." -ForegroundColor Yellow; continue }

    $moved = 0; $failMove = 0
    foreach ($s in $sources) {
      $tasksInSource = @($script:Tasks | Where-Object { $_.bucketId -eq $s.id })
      foreach ($t in $tasksInSource) {
        try {
          Move-TaskToBucket -TaskId $t.id -TargetBucketId $target.id
          $moved++
        } catch {
          $failMove++
          $msg = $_.Exception.Message
          if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $msg = $_.ErrorDetails.Message }
          Write-Host "    [fail] move task '$($t.title)' - $msg" -ForegroundColor Red
        }
      }
    }
    Write-Host "    Moved: $moved   Failed: $failMove" -ForegroundColor Green

    Write-Host ""
    $delOpt = Read-Host "  Delete the now-empty source bucket(s)? (y/N)"
    if ($delOpt -eq 'y' -or $delOpt -eq 'Y') {
      foreach ($s in $sources) {
        try {
          $etag = Get-ResourceEtag "https://graph.microsoft.com/v1.0/planner/buckets/$($s.id)"
          Invoke-MgGraphRequest -Method DELETE `
            -Uri "https://graph.microsoft.com/v1.0/planner/buckets/$($s.id)" `
            -Headers @{ "If-Match" = $etag } -ErrorAction Stop
          Write-Host "    [deleted] $($s.name)" -ForegroundColor Green
        } catch {
          Write-Host "    [fail]    $($s.name) - $($_.Exception.Message)" -ForegroundColor Red
        }
      }
    }
    # refresh between groups so counts stay accurate
    Load-PlanData
  }
  Write-Host ""
  Write-Host "  Done." -ForegroundColor White
}

function Get-BucketLastActivity {
  param([string]$BucketId)
  $tasks = @($script:Tasks | Where-Object { $_.bucketId -eq $BucketId })
  if ($tasks.Count -eq 0) { return $null }
  $maxDate = $null
  foreach ($t in $tasks) {
    foreach ($d in @($t.createdDateTime, $t.completedDateTime)) {
      if ($d) {
        try {
          $dt = [datetime]$d
          if (-not $maxDate -or $dt -gt $maxDate) { $maxDate = $dt }
        } catch {}
      }
    }
  }
  return $maxDate
}

function Remove-StaleBuckets {
  if ($script:Buckets.Count -eq 0) {
    Write-Host ""
    Write-Host "  No buckets in this plan." -ForegroundColor Green
    return
  }
  Write-Host ""
  Write-Host "  Cleanup stale buckets (by last task activity)." -ForegroundColor Cyan
  Write-Host "  Note: Planner has no 'last modified' timestamp on buckets. Using the most" -ForegroundColor DarkGray
  Write-Host "  recent createdDateTime / completedDateTime of tasks inside each bucket." -ForegroundColor DarkGray
  Write-Host ""
  Write-Host "    [1]  Buckets with no activity in last 90 days" -ForegroundColor White
  Write-Host "    [2]  180 days" -ForegroundColor White
  Write-Host "    [3]  365 days" -ForegroundColor White
  Write-Host "    [4]  Custom (number of days)" -ForegroundColor White
  Write-Host "    [5]  Custom (specific cutoff date YYYY-MM-DD)" -ForegroundColor White
  $pick = Read-Host "  Choose (1-5)"

  $cutoff = $null
  switch ($pick) {
    '1' { $cutoff = (Get-Date).AddDays(-90) }
    '2' { $cutoff = (Get-Date).AddDays(-180) }
    '3' { $cutoff = (Get-Date).AddDays(-365) }
    '4' {
      $d = Read-Host "  Number of days"
      if (-not ($d -match '^\d+$') -or [int]$d -le 0) { Write-Host "  Invalid." -ForegroundColor Red; return }
      $cutoff = (Get-Date).AddDays(-[int]$d)
    }
    '5' {
      $s = Read-Host "  Cutoff date (YYYY-MM-DD) - buckets with activity before this are stale"
      try { $cutoff = [datetime]::ParseExact($s, "yyyy-MM-dd", $null) }
      catch { Write-Host "  Invalid date format." -ForegroundColor Red; return }
    }
    default { Write-Host "  Invalid choice." -ForegroundColor Red; return }
  }
  Write-Host ""
  Write-Host "  Cutoff: $($cutoff.ToString('yyyy-MM-dd'))" -ForegroundColor White
  Write-Host "  (Buckets with last activity strictly before this date are stale.)" -ForegroundColor DarkGray

  $stale = [System.Collections.Generic.List[object]]::new()
  foreach ($b in $script:Buckets) {
    $last = Get-BucketLastActivity -BucketId $b.id
    $isStale = $false
    if (-not $last) { $isStale = $true }
    elseif ($last -lt $cutoff) { $isStale = $true }
    if ($isStale) {
      [void]$stale.Add([pscustomobject]@{
        Bucket   = $b
        LastActivity = $last
        TaskCount = Get-BucketCount $b.id
      })
    }
  }

  if ($stale.Count -eq 0) {
    Write-Host ""
    Write-Host "  No stale buckets found." -ForegroundColor Green
    return
  }

  Write-Host ""
  Write-Host "  Found $($stale.Count) stale bucket(s):" -ForegroundColor Yellow
  Write-Host ("  {0,-40}  {1,-12}  {2}" -f "Bucket","Last activity","Tasks") -ForegroundColor DarkGray
  Write-Host ("  {0,-40}  {1,-12}  {2}" -f "------","-------------","-----") -ForegroundColor DarkGray
  $sorted = @($stale | Sort-Object { if ($_.LastActivity) { $_.LastActivity } else { [datetime]::MinValue } })
  foreach ($s in $sorted) {
    $dateStr = if ($s.LastActivity) { $s.LastActivity.ToString("yyyy-MM-dd") } else { "(no tasks)" }
    $color = if ($s.TaskCount -eq 0) { "Red" } else { "Yellow" }
    Write-Host ("  {0,-40}  {1,-12}  {2}" -f $s.Bucket.name, $dateStr, $s.TaskCount) -ForegroundColor $color
  }

  $withTasks = @($stale | Where-Object { $_.TaskCount -gt 0 })
  $noTasks   = @($stale | Where-Object { $_.TaskCount -eq 0 })

  Write-Host ""
  Write-Host "  $($noTasks.Count) empty, $($withTasks.Count) with tasks." -ForegroundColor White
  Write-Host ""
  Write-Host "  IMPORTANT: tasks inside these buckets will be permanently deleted." -ForegroundColor Red
  Write-Host "  Strongly recommend running [6] Export backup first." -ForegroundColor Red
  Write-Host ""
  Write-Host "    [E]mpty only - safely delete only the $($noTasks.Count) empty bucket(s)" -ForegroundColor White
  Write-Host "    [A]ll - delete tasks + buckets (all $($stale.Count))" -ForegroundColor White
  Write-Host "    [C]ancel" -ForegroundColor Gray
  $act = Read-Host "  Choose"

  $targets = @()
  switch ($act.ToUpper()) {
    'E' { $targets = $noTasks }
    'A' { $targets = $stale }
    default { Write-Host "  Cancelled." -ForegroundColor Yellow; return }
  }
  if ($targets.Count -eq 0) { Write-Host "  Nothing to delete." -ForegroundColor Green; return }

  Write-Host ""
  Write-Host "  About to delete $($targets.Count) bucket(s)." -ForegroundColor Red
  $conf = Read-Host "  Type 'DELETE' to confirm"
  if ($conf -ne 'DELETE') { Write-Host "  Cancelled." -ForegroundColor Yellow; return }

  $okB = 0; $failB = 0; $okT = 0; $failT = 0
  foreach ($s in $targets) {
    $b = $s.Bucket
    if ($s.TaskCount -gt 0) {
      $tasksIn = @($script:Tasks | Where-Object { $_.bucketId -eq $b.id })
      foreach ($t in $tasksIn) {
        try {
          $etagT = Get-ResourceEtag "https://graph.microsoft.com/v1.0/planner/tasks/$($t.id)"
          Invoke-MgGraphRequest -Method DELETE `
            -Uri "https://graph.microsoft.com/v1.0/planner/tasks/$($t.id)" `
            -Headers @{ "If-Match" = $etagT } -ErrorAction Stop
          $okT++
        } catch {
          Write-Host "    [fail] delete task '$($t.title)' - $($_.Exception.Message)" -ForegroundColor Red
          $failT++
        }
      }
    }
    try {
      $etagB = Get-ResourceEtag "https://graph.microsoft.com/v1.0/planner/buckets/$($b.id)"
      Invoke-MgGraphRequest -Method DELETE `
        -Uri "https://graph.microsoft.com/v1.0/planner/buckets/$($b.id)" `
        -Headers @{ "If-Match" = $etagB } -ErrorAction Stop
      Write-Host "    [deleted] $($b.name)" -ForegroundColor Green
      $okB++
    } catch {
      Write-Host "    [fail]    $($b.name) - $($_.Exception.Message)" -ForegroundColor Red
      $failB++
    }
  }
  Write-Host ""
  Write-Host "  Buckets - deleted: $okB   failed: $failB" -ForegroundColor White
  Write-Host "  Tasks   - deleted: $okT   failed: $failT" -ForegroundColor White
  Load-PlanData
}

function Export-Backup {
  $safeTitle = ($script:Plan.title -replace '[^a-zA-Z0-9\-_]', '_')
  $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
  $path = Join-Path $HOME "planner-backup_${safeTitle}_${stamp}.json"

  $payload = [ordered]@{
    exportedAt = (Get-Date).ToString("o")
    tenantId   = $TenantId
    planId     = $PlanId
    plan       = $script:Plan
    buckets    = $script:Buckets
    tasks      = $script:Tasks
  }

  try {
    $payload | ConvertTo-Json -Depth 10 | Set-Content -Path $path -Encoding UTF8
    Write-Host ""
    Write-Host "  Backup saved:" -ForegroundColor Green
    Write-Host "    $path" -ForegroundColor White
    Write-Host ("    $($script:Buckets.Count) bucket(s), $($script:Tasks.Count) task(s)") -ForegroundColor Gray
  } catch {
    Write-Host ""
    Write-Host "  [!] Backup failed: $($_.Exception.Message)" -ForegroundColor Red
  }
}

function List-TasksWithDates {
  if ($script:Tasks.Count -eq 0) {
    Write-Host ""
    Write-Host "  No tasks in this plan." -ForegroundColor Green
    return
  }
  Write-Host ""
  Write-Host "  Note: Planner API does NOT expose last-accessed or last-modified." -ForegroundColor DarkYellow
  Write-Host "        Showing created + completed (and due) dates instead." -ForegroundColor DarkYellow
  Write-Host ""
  Write-Host "    [1] Oldest created first" -ForegroundColor Gray
  Write-Host "    [2] Newest created first" -ForegroundColor Gray
  Write-Host "    [3] Open tasks only (not completed)" -ForegroundColor Gray
  $choice = Read-Host "  Choice (1-3)"

  $list = if ($choice -eq '2') {
    @($script:Tasks | Sort-Object createdDateTime -Descending)
  } elseif ($choice -eq '3') {
    @($script:Tasks | Where-Object { $_.percentComplete -lt 100 } | Sort-Object createdDateTime)
  } else {
    @($script:Tasks | Sort-Object createdDateTime)
  }

  if ($list.Count -eq 0) {
    Write-Host "  No tasks match the filter." -ForegroundColor Green
    return
  }

  Write-Host ""
  Write-Host ("  {0,-10}  {1,-10}  {2,-10}  {3,-20}  {4}" -f "Created","Due","Done","Bucket","Title") -ForegroundColor White
  Write-Host ("  {0,-10}  {1,-10}  {2,-10}  {3,-20}  {4}" -f "-------","---","----","------","-----") -ForegroundColor DarkGray
  foreach ($t in $list) {
    $bn = Get-BucketName $t.bucketId
    if ($bn.Length -gt 20) { $bn = $bn.Substring(0,17) + "..." }
    $title = if ($t.title) { $t.title } else { "(no title)" }
    if ($title.Length -gt 60) { $title = $title.Substring(0,57) + "..." }
    Write-Host ("  {0,-10}  {1,-10}  {2,-10}  {3,-20}  {4}" -f `
      (Fmt-Date $t.createdDateTime), (Fmt-Date $t.dueDateTime), (Fmt-Date $t.completedDateTime), $bn, $title) -ForegroundColor Gray
  }
  Write-Host ""
  Write-Host "  Total: $($list.Count) task(s)" -ForegroundColor White
}

# ── Menu loop ──
while ($true) {
  Write-Host ""
  Write-Host "  ------------------------------------------------------------" -ForegroundColor DarkGray
  Write-Host "  Plan: $($script:Plan.title)   (buckets: $($script:Buckets.Count), tasks: $($script:Tasks.Count))" -ForegroundColor White
  Write-Host "  ------------------------------------------------------------" -ForegroundColor DarkGray
  Write-Host "    [1] Show overview (buckets + task counts)" -ForegroundColor White
  Write-Host "    [2] Sort buckets A-Z" -ForegroundColor White
  Write-Host "    [3] Delete empty buckets" -ForegroundColor White
  Write-Host "    [4] Find duplicate BUCKETS (merge tasks + delete spare)" -ForegroundColor White
  Write-Host "    [5] List tasks with dates (created / due / completed)" -ForegroundColor White
  Write-Host "    [6] Export backup (JSON of plan, buckets, tasks)" -ForegroundColor White
  Write-Host "    [7] Cleanup stale buckets (delete by last activity date)" -ForegroundColor White
  Write-Host "    [R] Refresh data" -ForegroundColor Gray
  Write-Host "    [Q] Quit" -ForegroundColor Gray
  Write-Host ""
  $pick = Read-Host "  Choose"

  switch ($pick.ToUpper()) {
    '1' { Show-Overview }
    '2' { Sort-Buckets }
    '3' { Remove-EmptyBuckets }
    '4' { Find-DuplicateBuckets }
    '5' { List-TasksWithDates }
    '6' { Export-Backup }
    '7' { Remove-StaleBuckets }
    'R' { Write-Host "  Refreshing..." -ForegroundColor Cyan; Load-PlanData }
    'Q' { break }
    default { Write-Host "  Unknown choice." -ForegroundColor Yellow }
  }
  if ($pick.ToUpper() -eq 'Q') { break }
}

Write-Host ""
Write-Host "  Signing out." -ForegroundColor Cyan
Disconnect-MgGraph | Out-Null
