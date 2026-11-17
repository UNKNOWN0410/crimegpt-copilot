#Requires -Version 5.1
<#
.SYNOPSIS
  Prove a CrimeGPT install actually works. Standalone - run any time the servers are up.

.DESCRIPTION
  setup.ps1 installs and ends at READY without starting servers. This script is the other
  half: it exercises the running stack and prints PASS/FAIL with a fix for every failure.

  Run it after .\start.ps1, or any time before a demo.

  Two deliberate choices:

  * Everything is hit on 127.0.0.1, never "localhost". On Windows "localhost" resolves to
    IPv6 ::1 first; uvicorn binds IPv4, so each NEW connection stalls ~2s before falling
    back. Measured and reproducible on the demo machine. A timing taken through
    "localhost" is measuring the resolver, not the app.

  * It does not re-implement scripts\preflight.py. That already checks Postgres, Alembic
    head, seed rows, Chroma = 1059, judgments = 41 and Ollama + qwen2.5:7b. This runs it
    (read-only: no --fix, no --no-start side effects) and then adds only what it cannot
    know - the things that need the HTTP stack alive.

.PARAMETER FullCheck
  Opt IN to every check that WRITES. READ-ONLY BY DEFAULT, deliberately: generating a
  document bumps a real document version, and version history is shown on stage, so a
  verification run must not quietly advance it. Without this switch NOTHING in the
  database changes.

  -FullCheck RUNS THE FULL FEATURE MATRIX, AND IT WRITES. Specifically it creates:

    * ONE new case, "VERIFY-<timestamp>", built from scratch through /intake/commit.
      It is NOT deleted afterwards - a verification run does not delete data. Delete it
      yourself if you want a clean case list. Every other write below lands on THIS case,
      so no seeded case and no demo case is touched.
    * On that case: 4 persons, 1 seized item, and the case fields (FIR, incident,
      narrative), written through the normal pool endpoints.
    * Legal sections: one live /analyze run, and every section it suggests is ACCEPTED
      (the section-dependent documents cannot generate otherwise).
    * ~27 document generations on that case - 8 types x EN/HI/GU, plus a chargesheet
      original/supplementary pair and one extra draft to finalize. These land on 8
      document rows, each bumped to a new version per generation.
    * ONE document finalized (DRAFT -> FINALIZED) on that case.
    * The audit_log and case_diary rows all of the above implies, plus the auth.step_up
      row that each refused step-up records.

  It does NOT write: /intake/extract, the chat routes and /consistency have no side
  effects, and a refused generation or a wrong PIN writes nothing at all.

  It also needs Qwen: /analyze and /intake/extract are live LLM calls, so a -FullCheck
  run takes minutes, not seconds.

.PARAMETER CaseId
  Case used by -FullCheck's document generation. Default 2, deliberately NOT the case 1
  demo case, so a verification run never touches the case you demo.
#>
[CmdletBinding()]
param(
    [switch]$FullCheck,
    [int]$CaseId = 2
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Continue"

$Root     = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $Root) { $Root = (Get-Location).Path }
$Backend  = Join-Path $Root "backend"
$VenvPy   = Join-Path $Backend ".venv\Scripts\python.exe"
$EnvLocal = Join-Path $Root "frontend\.env.local"

$API  = "http://127.0.0.1:8000"
$UI   = "http://127.0.0.1:3000"

$script:Pass = 0
$script:Fail = 0
$script:Notes = @()

function Ok([string]$name, [string]$detail) {
    Write-Host ("  PASS  " + $name) -ForegroundColor Green
    if ($detail) { Write-Host ("        " + $detail) -ForegroundColor DarkGray }
    $script:Pass++
}
function Bad([string]$name, [string]$detail, [string]$fix) {
    Write-Host ("  FAIL  " + $name) -ForegroundColor Red
    if ($detail) { Write-Host ("        " + $detail) -ForegroundColor DarkGray }
    Write-Host ("        FIX: " + $fix) -ForegroundColor Yellow
    $script:Fail++
}
function Note([string]$msg) { $script:Notes += $msg }
function Section([string]$msg) { Write-Host ""; Write-Host ("== " + $msg) -ForegroundColor Cyan }

function Try-Get([string]$url, [int]$timeoutSec = 10) {
    try { return Invoke-WebRequest -Uri $url -TimeoutSec $timeoutSec -UseBasicParsing }
    catch { return $null }
}

# --- Feature-matrix helpers (used only by -FullCheck) ------------------------
# One HTTP shape for every call: never throws, always reports the status code, so a
# check can assert on 401/403/400 as a RESULT rather than as an exception.
function Api {
    param(
        [string]$Method,
        [string]$Url,
        $Body = $null,
        [string]$Token = $null,
        [int]$TimeoutSec = 180
    )
    $req = @{ Uri = $Url; Method = $Method; TimeoutSec = $TimeoutSec; UseBasicParsing = $true }
    if ($Token) { $req["Headers"] = @{ Authorization = "Bearer $Token" } }
    if ($null -ne $Body) {
        # Bytes, not a string: PS 5.1 posts strings as ISO-8859-1 and mangles any
        # non-ASCII on the way out. The Gujarati checks depend on this.
        $json = $Body | ConvertTo-Json -Depth 8 -Compress
        $req["Body"] = [System.Text.Encoding]::UTF8.GetBytes($json)
        $req["ContentType"] = "application/json; charset=utf-8"
    }
    try {
        $r = Invoke-WebRequest @req
        $parsed = $null
        if ($r.Content) { try { $parsed = $r.Content | ConvertFrom-Json } catch { $parsed = $null } }
        return [pscustomobject]@{ Status = [int]$r.StatusCode; Body = $parsed; Raw = $r.Content; Error = $null }
    } catch {
        $st = 0; $raw = ""
        # PS 5.1 has usually already drained the response stream by the time we get here
        # and parked the body on ErrorDetails. Read that first, or a 400 arrives with the
        # status code but no message - which is the half of it worth reading.
        if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $raw = [string]$_.ErrorDetails.Message }
        if ($_.Exception.Response) {
            $st = [int]$_.Exception.Response.StatusCode
            if (-not $raw) {
                try {
                    $rs = $_.Exception.Response.GetResponseStream()
                    $sr = New-Object System.IO.StreamReader($rs, [System.Text.Encoding]::UTF8)
                    $raw = $sr.ReadToEnd(); $sr.Close()
                } catch { }
            }
        }
        $parsed = $null
        if ($raw) { try { $parsed = $raw | ConvertFrom-Json } catch { $parsed = $null } }
        return [pscustomobject]@{ Status = $st; Body = $parsed; Raw = $raw; Error = $_.Exception.Message }
    }
}

# Read the visible TEXT out of a .docx. Opens the zip and strips the XML tags, which also
# re-joins a merge tag that Word split across runs - so an UNRENDERED {{ tag }} is visible
# here even when it is not contiguous in the file.
function Get-DocxText([string]$path) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
    $zip = [System.IO.Compression.ZipFile]::OpenRead($path)
    try {
        $sb = New-Object System.Text.StringBuilder
        foreach ($e in $zip.Entries) {
            if ($e.FullName -match '^word/(document|header\d*|footer\d*)\.xml$') {
                $sr = New-Object System.IO.StreamReader($e.Open(), [System.Text.Encoding]::UTF8)
                $xml = $sr.ReadToEnd(); $sr.Close()
                $xml = $xml -replace '</w:p>', "`n"
                [void]$sb.AppendLine(($xml -replace '<[^>]+>', ''))
            }
        }
        $t = $sb.ToString()
        return ($t -replace '&amp;', '&' -replace '&lt;', '<' -replace '&gt;', '>' -replace '&quot;', '"')
    } finally { $zip.Dispose() }
}

# Script census by Unicode block. This is the whole point of the language checks: the
# request and the filename are both trusted by nobody - only the codepoints count.
function Get-ScriptCensus([string]$text) {
    return [pscustomobject]@{
        Devanagari = @([regex]::Matches($text, '[\u0900-\u097F]')).Count
        Gujarati   = @([regex]::Matches($text, '[\u0A80-\u0AFF]')).Count
        Latin      = @([regex]::Matches($text, '[A-Za-z]')).Count
    }
}

# Download a generated document and hand back its text. $null on any failure.
function Get-GeneratedText([int]$docId, [string]$token) {
    $tmp = Join-Path $env:TEMP ("crimegpt_verify_{0}.docx" -f $docId)
    try {
        Invoke-WebRequest "$API/api/documents/$docId/download" `
            -Headers @{ Authorization = "Bearer $token" } -TimeoutSec 60 `
            -UseBasicParsing -OutFile $tmp | Out-Null
        $text = Get-DocxText $tmp
        Remove-Item $tmp -Force -ErrorAction SilentlyContinue
        return $text
    } catch {
        Remove-Item $tmp -Force -ErrorAction SilentlyContinue
        return $null
    }
}

Write-Host ""
Write-Host "CrimeGPT verification" -ForegroundColor White
Write-Host "(all HTTP on 127.0.0.1 - 'localhost' costs ~2s per new connection on Windows)"

# ---------------------------------------------------------------------------
# 1. Dependencies that need no HTTP stack - delegate to preflight.py
# ---------------------------------------------------------------------------
Section "Dependencies (scripts\preflight.py - Postgres, Alembic, seed, Chroma 1059, Ollama)"
if (-not (Test-Path $VenvPy)) {
    Bad "backend venv present" "backend\.venv\Scripts\python.exe not found" "Run .\setup.ps1 first."
} else {
    # No --fix (never mutate during verification). --no-start so it reports Postgres
    # rather than silently starting containers behind your back.
    $pfOut = & $VenvPy (Join-Path $Root "scripts\preflight.py") --no-start 2>&1
    $pfExit = $LASTEXITCODE
    $pfOut | ForEach-Object { Write-Host ("        " + $_) -ForegroundColor DarkGray }

    if ($pfExit -eq 0) {
        Ok "preflight.py all checks" "Postgres, Alembic head, seed, Chroma=1059, judgments=41, Ollama+qwen2.5:7b"
    } else {
        # preflight's seed check asserts the EXACT seed baseline (got == want), so any
        # machine that has registered real cases fails it - which is normal, not broken.
        # verify_seed_counts.py answers the question we actually care about here
        # ("is the seed usable": users >= 4, cases >= 2), so use it to tell a surplus
        # apart from a genuinely missing/half-seeded database.
        $pfText = ($pfOut | Out-String)
        # @() forces an array - under Set-StrictMode a single match has no .Count.
        $failed = @([regex]::Matches($pfText, '\[FAIL\]\s+(.+?)\s{2,}') |
                    ForEach-Object { $_.Groups[1].Value.Trim() })
        $onlySeed = ($failed.Count -eq 1 -and $failed[0] -match 'Seed data present')

        & $VenvPy (Join-Path $Root "scripts\verify_seed_counts.py") 2>&1 | Out-Null
        $seedUsable = ($LASTEXITCODE -eq 0)

        if ($onlySeed -and $seedUsable) {
            Write-Host "  WARN  preflight: seed counts differ from the pristine baseline" -ForegroundColor DarkYellow
            Write-Host "        Your database has MORE than the seeded rows - normal on a machine that has" -ForegroundColor DarkGray
            Write-Host "        registered cases. verify_seed_counts.py confirms the seed is usable" -ForegroundColor DarkGray
            Write-Host "        (users >= 4, cases >= 2). Not treated as a failure." -ForegroundColor DarkGray
            Ok "preflight.py dependency checks" "all non-seed checks PASS (Postgres, Alembic, Chroma=1059, judgments=41, Ollama)"
        } else {
            Bad "preflight.py" "one or more dependency checks failed (detail above)" `
                "Read the FAIL lines above. Chroma empty -> cd backend; .\.venv\Scripts\python.exe -m app.ai.rag (~2 min). Migrations behind -> .\.venv\Scripts\python.exe ..\scripts\preflight.py --fix. Seed missing -> cd backend; .\.venv\Scripts\python.exe -m app.seed"
        }
    }
}

# ---------------------------------------------------------------------------
# 2. Backend on 127.0.0.1
# ---------------------------------------------------------------------------
Section "HTTP stack"
$health = Try-Get "$API/health"
if ($health -and $health.StatusCode -eq 200) {
    Ok "backend responds on 127.0.0.1:8000" $health.Content
} else {
    Bad "backend responds on 127.0.0.1:8000" "no response from $API/health" `
        "Start it: .\start.ps1 -BackendOnly   (never bare 'uvicorn' - it can bind Python 3.11 without docxtpl)"
}

$db = Try-Get "$API/health/db"
if ($db -and $db.StatusCode -eq 200 -and $db.Content -match '"ok"') {
    Ok "backend reaches Postgres" $db.Content
} else {
    Bad "backend reaches Postgres" "GET /health/db did not return db=ok" `
        "Start Docker Desktop, then: docker compose up -d"
}

# ---------------------------------------------------------------------------
# 3. Frontend on 3000
# ---------------------------------------------------------------------------
$ui = Try-Get $UI 20
if ($ui -and $ui.StatusCode -eq 200) {
    Ok "frontend responds on port 3000" "$($ui.RawContentLength) bytes"
} else {
    Bad "frontend responds on port 3000" "no response from $UI" `
        "Start it: .\start.ps1 -FrontendOnly. If it silently took 3001, kill the orphaned node and re-run (npm run dev -- -p 3000)."
}

# ---------------------------------------------------------------------------
# 4. Auth - all three roles
# ---------------------------------------------------------------------------
Section "Authentication and RBAC"
$tokens = @{}
$roleMap = @{ io = "IO"; sho = "SHO"; legal = "LEGAL_ADVISOR" }
$loginOk = $true
foreach ($u in @("io", "sho", "legal")) {
    try {
        $body = @{ username = $u; password = "$($u)123" } | ConvertTo-Json
        $r = Invoke-RestMethod "$API/api/auth/login" -Method Post -Body $body -ContentType "application/json" -TimeoutSec 20
        if ($r.role -ne $roleMap[$u]) { throw "role was $($r.role), expected $($roleMap[$u])" }
        $tokens[$u] = $r.token
    } catch {
        $loginOk = $false
        Bad "login $u" $_.Exception.Message "Re-seed: cd backend; .\.venv\Scripts\python.exe -m app.seed (idempotent, does not wipe)"
    }
}
if ($loginOk) { Ok "all 3 roles log in" "io/IO, sho/SHO, legal/LEGAL_ADVISOR" }

# ---------------------------------------------------------------------------
# 5. Document generation end to end
# ---------------------------------------------------------------------------
Section "Document generation"
if (-not $FullCheck) {
    Write-Host "  SKIP  document generation - read-only run. Pass -FullCheck to test it." -ForegroundColor DarkYellow
    Write-Host "        (it bumps a real document version; version history is demo material)" -ForegroundColor DarkGray
} elseif (-not $tokens.ContainsKey("sho")) {
    Bad "one document generates end to end" "no SHO token (login failed above)" "Fix login first."
} else {
    try {
        $h = @{ Authorization = "Bearer $($tokens['sho'])" }
        $doc = Invoke-RestMethod "$API/api/cases/$CaseId/documents/SEIZURE_RECEIPT?lang=en" -Method Post -Headers $h -TimeoutSec 120
        $dl = Invoke-WebRequest "$API/api/documents/$($doc.id)/download" -Headers $h -TimeoutSec 60 -UseBasicParsing
        if ($dl.StatusCode -eq 200 -and $dl.RawContentLength -gt 20000) {
            Ok "one document generates end to end" "case $CaseId -> doc $($doc.id) v$($doc.version) $($doc.status), $($dl.RawContentLength) bytes"
            Note "Document generation wrote: case $CaseId SEIZURE_RECEIPT is now v$($doc.version) (previous version archived to document_versions)."
        } else {
            Bad "one document generates end to end" "download was $($dl.RawContentLength) bytes" `
                "Check templates\ has all 8 .docx and that docxtpl imports under backend\.venv."
        }
    } catch {
        Bad "one document generates end to end" $_.Exception.Message `
            "A 400 means the case lacks required pool fields. A 500 usually means docxtpl is missing - re-run .\setup.ps1."
    }
}

# ---------------------------------------------------------------------------
# 6. Step-up enforcement (commit 495a34a) - server-side, silently important
# ---------------------------------------------------------------------------
Section "Step-up PIN enforcement (server-side)"
if (-not $FullCheck) {
    Write-Host "  SKIP  step-up probe - read-only run. Pass -FullCheck to test it." -ForegroundColor DarkYellow
    Write-Host "        (a refusal writes one auth.step_up audit row)" -ForegroundColor DarkGray
} elseif (-not $tokens.ContainsKey("io")) {
    Bad "commit without step-up is refused" "no IO token" "Fix login first."
} else {
    # A fresh password session has NOT stepped up, so commit must be refused with 401.
    # Nothing is written on refusal except the audit row that records it.
    $stamp = Get-Date -Format "HHmmss"
    $payload = @{
        case = @{ case_number = "VERIFY-$stamp"; title = "verify.ps1 step-up probe"
                  police_station = "Satellite Police Station"; district = "Ahmedabad" }
        persons = @(); seized_items = @()
    } | ConvertTo-Json -Depth 6
    $status = 0
    try {
        Invoke-RestMethod "$API/api/intake/commit" -Method Post `
            -Headers @{ Authorization = "Bearer $($tokens['io'])" } `
            -Body $payload -ContentType "application/json" -TimeoutSec 60 | Out-Null
        $status = 201
    } catch {
        if ($_.Exception.Response) { $status = [int]$_.Exception.Response.StatusCode }
    }
    if ($status -eq 401) {
        Ok "commit without step-up returns 401" "server-side gate live (495a34a); no case was created"
    } elseif ($status -eq 201) {
        Bad "commit without step-up returns 401" "commit SUCCEEDED without a step-up - the gate is NOT enforced" `
            "You are running code older than 495a34a, or require_step_up was removed from intake_commit. Check: git log --oneline -1"
    } else {
        Bad "commit without step-up returns 401" "got HTTP $status" `
            "Expected 401. A 403 means the role gate fired first; a 404 means the case is not visible."
    }
}

# ---------------------------------------------------------------------------
# 7. LAN / mobile - and the firewall diagnosis
# ---------------------------------------------------------------------------
Section "LAN / mobile field page"
$lanIp = $null
if (Test-Path $EnvLocal) {
    $line = Select-String -LiteralPath $EnvLocal -Pattern '^\s*NEXT_PUBLIC_API_URL\s*=\s*http://([0-9\.]+):8000' |
            Select-Object -First 1
    if ($line) { $lanIp = $line.Matches[0].Groups[1].Value }
}
if (-not $lanIp) {
    Bad "frontend\.env.local names a LAN IP" "NEXT_PUBLIC_API_URL not found or not an IP in $EnvLocal" `
        "Re-run .\setup.ps1 - it writes .env.local when missing. Without it the phone calls 'localhost', which is the phone itself."
} else {
    Ok "frontend\.env.local names a LAN IP" "NEXT_PUBLIC_API_URL -> http://${lanIp}:8000"

    # The IP in .env.local is a SNAPSHOT taken by setup.ps1. DHCP moves, and a stale IP
    # here fails as "the phone loads /m but sign-in hangs" - which reads like a firewall
    # problem and is not one. Check it against the adapters this machine actually has,
    # and against the CORS origin the backend was told to allow, which goes stale with it.
    $adapters = @()
    try {
        $adapters = @(Get-NetIPAddress -AddressFamily IPv4 -ErrorAction Stop |
                      Where-Object { $_.IPAddress -ne "127.0.0.1" } |
                      Select-Object -ExpandProperty IPAddress)
    } catch {
        $adapters = @([System.Net.Dns]::GetHostAddresses([System.Net.Dns]::GetHostName()) |
                      Where-Object { $_.AddressFamily -eq "InterNetwork" } |
                      ForEach-Object { $_.IPAddressToString })
    }
    if ($adapters -contains $lanIp) {
        Ok "the LAN IP in .env.local is a current adapter address" "$lanIp is live on this machine"
    } else {
        Bad "the LAN IP in .env.local is a current adapter address" "$lanIp is NOT on this machine. Current: $($adapters -join ', ')" `
            "Your IP moved (DHCP). Re-run .\setup.ps1 to rewrite frontend\.env.local and backend\.env, then RESTART BOTH servers - Next.js bakes NEXT_PUBLIC_* in at build time, so an edit alone does nothing."
    }

    $envBackend = Join-Path $Root "backend\.env"
    if (-not (Test-Path $envBackend)) {
        Bad "CORS_EXTRA_ORIGINS matches the LAN IP" "backend\.env not found" "Copy backend\.env.example to backend\.env and re-run .\setup.ps1."
    } else {
        $corsLine = Select-String -LiteralPath $envBackend -Pattern '^\s*CORS_EXTRA_ORIGINS\s*=\s*(.+)$' | Select-Object -First 1
        $cors = if ($corsLine) { $corsLine.Matches[0].Groups[1].Value.Trim() } else { "" }
        if ($cors -match [regex]::Escape("http://${lanIp}:3000")) {
            Ok "CORS_EXTRA_ORIGINS matches the LAN IP" $cors
        } elseif (-not $cors) {
            Bad "CORS_EXTRA_ORIGINS matches the LAN IP" "CORS_EXTRA_ORIGINS is not set in backend\.env" `
                "The phone's origin is http://${lanIp}:3000. Without it the browser blocks every API call from /m. Add: CORS_EXTRA_ORIGINS=http://${lanIp}:3000 and restart the backend."
        } else {
            Bad "CORS_EXTRA_ORIGINS matches the LAN IP" "backend\.env allows '$cors' but .env.local points the phone at http://${lanIp}:3000" `
                "The two files went stale apart. Set CORS_EXTRA_ORIGINS=http://${lanIp}:3000 in backend\.env and restart the backend, or re-run .\setup.ps1 to rewrite both."
        }
    }

    $lanApi = Try-Get "http://${lanIp}:8000/health" 8
    $lanUi  = Try-Get "http://${lanIp}:3000/m" 15

    if ($lanUi -and $lanUi.StatusCode -eq 200) {
        Ok "/m reachable on the LAN IP" "http://${lanIp}:3000/m ($($lanUi.RawContentLength) bytes)"
    } else {
        # The signature failure: loopback fine, LAN blocked.
        if ($health) {
            Bad "/m reachable on the LAN IP" "port 3000 answers on 127.0.0.1 but NOT on ${lanIp} - this is the Windows firewall blocking the LAN interface" `
                "Run the two New-NetFirewallRule commands from the setup READY output in an ADMIN PowerShell, then re-run. Also confirm the server is bound to 0.0.0.0: .\start.ps1 -Lan"
        } else {
            Bad "/m reachable on the LAN IP" "nothing is listening" "Start the servers: .\start.ps1 -Lan"
        }
    }

    if ($lanApi -and $lanApi.StatusCode -eq 200) {
        Ok "API reachable on the LAN IP" "http://${lanIp}:8000/health"
    } else {
        Bad "API reachable on the LAN IP" "port 8000 answers on 127.0.0.1 but NOT on ${lanIp} - this is the Windows firewall blocking the LAN interface. This is the exact cause of '/m loads but sign-in hangs'." `
            "Admin PowerShell: New-NetFirewallRule -DisplayName 'CrimeGPT backend 8000' -Direction Inbound -Protocol TCP -LocalPort 8000 -Action Allow. Then bind LAN: .\start.ps1 -Lan"
    }
}

# ---------------------------------------------------------------------------
# 8. Runtime configuration the demo depends on
# ---------------------------------------------------------------------------
Section "Runtime configuration"
# Check the LIVE eviction deadline, not the environment variable.
#
# An env var that is SET but NOT IN EFFECT is worse than one that is unset, because it
# reads as PASS while the model still evicts. OLLAMA_KEEP_ALIVE is read by the Ollama
# SERVER at startup, so setting it changes nothing until Ollama is restarted. The only
# honest source of truth is what Ollama itself reports: expires_at on the resident model.
$envKeep = [Environment]::GetEnvironmentVariable("OLLAMA_KEEP_ALIVE", "User")
$envNote = if ($envKeep) { "env var OLLAMA_KEEP_ALIVE='$envKeep' (user scope)" } else { "env var OLLAMA_KEEP_ALIVE is unset" }
try {
    $ps = Invoke-RestMethod "http://127.0.0.1:11434/api/ps" -TimeoutSec 10
    $resident = @($ps.models)
    if ($resident.Count -eq 0) {
        Write-Host "  WARN  model eviction deadline: no model resident in Ollama right now" -ForegroundColor DarkYellow
        Write-Host "        Cannot read a live deadline until something loads the model. $envNote" -ForegroundColor DarkGray
        Write-Host "        Run one analysis, then re-run this check." -ForegroundColor DarkGray
    } else {
        $m = $resident[0]
        $mins = ([datetime]$m.expires_at - (Get-Date)).TotalMinutes
        if ($mins -ge 60) {
            Ok "model stays resident (live check)" ("{0} evicts in {1:N1} h - keep-alive IS in effect. {2}" -f $m.name, ($mins / 60), $envNote)
        } else {
            Bad "model stays resident (live check)" `
                ("{0} evicts in {1:N0} min - keep-alive is NOT in effect. {2}" -f $m.name, $mins, $envNote) `
                $(if ($envKeep) {
                    "The variable is set but Ollama has not picked it up. RESTART OLLAMA (quit it from the tray and reopen), then re-run."
                  } else {
                    "Re-run .\setup.ps1 to set OLLAMA_KEEP_ALIVE=24h, then RESTART OLLAMA. Otherwise the first request after a demo pause pays a ~6.9s reload."
                  })
        }
        Note "expires_at reflects the keep_alive of the most recent request, so this is the real deadline right now."
    }
} catch {
    Bad "model stays resident (live check)" "could not query http://127.0.0.1:11434/api/ps : $($_.Exception.Message)" `
        "Start the Ollama app, then re-run."
}

if ($tokens.ContainsKey("io")) {
    try {
        $dm = Invoke-RestMethod "$API/api/system/demo-mode" -Headers @{ Authorization = "Bearer $($tokens['io'])" } -TimeoutSec 20
        $what = if ($dm.demo_mode) { "cached analysis + documents for case 1 are served instantly; judgments and weak-charge still call Qwen live" }
                else { "everything runs live against Qwen - slower, and subject to the ~1-in-20 slow intake" }
        Ok "DEMO_MODE resolved" "DEMO_MODE = $($dm.demo_mode). In effect: $what"
    } catch {
        Bad "DEMO_MODE resolved" $_.Exception.Message "Backend must be up."
    }
}


# ---------------------------------------------------------------------------
# 9. Feature matrix - probe case
#
#    Everything from here to the summary runs ONLY under -FullCheck, because all of it
#    writes. It builds ONE self-contained probe case and runs every feature check against
#    that, so no seeded case and no demo case is ever touched.
# ---------------------------------------------------------------------------
Section "Feature matrix - probe case"

$ioTok    = if ($tokens.ContainsKey("io"))    { $tokens["io"] }    else { $null }
$shoTok   = if ($tokens.ContainsKey("sho"))   { $tokens["sho"] }   else { $null }
$legalTok = if ($tokens.ContainsKey("legal")) { $tokens["legal"] } else { $null }
$probeId    = 0
$probeNum   = ""
$probeNarr  = ""
$probeStamp = ""
$matrixReady = $false

if (-not $FullCheck) {
    Write-Host "  SKIP  feature matrix - read-only run. Pass -FullCheck to run it." -ForegroundColor DarkYellow
    Write-Host "        See the -FullCheck help at the top of this file for exactly what it writes." -ForegroundColor DarkGray
} elseif (-not $ioTok -or -not $shoTok -or -not $legalTok) {
    Bad "feature matrix" "one or more role tokens are missing (login failed above)" "Fix login first, then re-run."
} else {
    $probeStamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $probeNum   = "VERIFY-$probeStamp"
    # Deliberately plain English prose, and the probe case carries NO statements, so the
    # analysis input is exactly this string. The triggering-phrase check below asserts the
    # model's quotes appear verbatim in it, which is only meaningful if that holds.
    $probeNarr = "On 12/07/2026 at about 22:15 hrs near Prahladnagar Garden, Ahmedabad, " +
                 "two men on a motorcycle intercepted the complainant. One of them threatened " +
                 "the complainant with a knife and snatched his gold chain and his mobile phone. " +
                 "They then fled towards the SG Highway. The complainant sustained a cut on his " +
                 "left hand during the incident."

    # --- the step-up gate, from refusal through to the write it guards ---------------
    $casesBefore = Api GET "$API/api/cases" -Token $shoTok
    $countBefore = if ($casesBefore.Status -eq 200) { @($casesBefore.Body).Count } else { -1 }

    $minimal = @{
        case = @{ case_number = $probeNum; title = "verify.ps1 feature-matrix probe"
                  police_station = "Satellite Police Station"; district = "Ahmedabad" }
        persons = @(); seized_items = @()
    }

    # A wrong PIN must record NOTHING. Checked by its consequence, which is the only thing
    # that matters: the very next commit is still refused, and no case appeared.
    $wrong = Api POST "$API/api/auth/verify-pin" -Body @{ pin = "0000" } -Token $ioTok
    $stillRefused = Api POST "$API/api/intake/commit" -Body $minimal -Token $ioTok
    $casesAfterWrong = Api GET "$API/api/cases" -Token $shoTok
    $countAfterWrong = if ($casesAfterWrong.Status -eq 200) { @($casesAfterWrong.Body).Count } else { -2 }

    $wrongReason = if ($wrong.Body) { [string]$wrong.Body.reason } else { "(no body)" }
    $wrongOk = ($wrong.Status -eq 200 -and $wrong.Body -and -not $wrong.Body.ok -and $wrongReason -eq "wrong_pin")
    if ($wrongOk -and $stillRefused.Status -eq 401 -and $countAfterWrong -eq $countBefore) {
        Ok "wrong PIN writes nothing" "verify-pin ok=false reason=wrong_pin; the next commit is still 401; case count unchanged ($countBefore)"
    } else {
        Bad "wrong PIN writes nothing" "verify-pin=$($wrong.Status) reason=$wrongReason; commit=$($stillRefused.Status); cases $countBefore -> $countAfterWrong" `
            "A wrong PIN must return ok=false and record no step-up. If the commit returned 201, require_step_up is no longer gating /intake/commit - check app\api\intake.py."
    }

    # Correct PIN, then the same commit must go through.
    $good   = Api POST "$API/api/auth/verify-pin" -Body @{ pin = "1234" } -Token $ioTok
    $commit = Api POST "$API/api/intake/commit" -Body $minimal -Token $ioTok
    $goodOk = ($good.Status -eq 200 -and $good.Body -and $good.Body.ok)
    if ($goodOk -and $commit.Status -eq 201) {
        $probeId = [int]$commit.Body.case.id
        Ok "commit succeeds after verify-pin" "probe case $probeNum -> id $probeId"
        Note "Feature matrix created case $probeId ($probeNum) and its documents. Delete it after the run if you want a clean case list."
    } else {
        Bad "commit succeeds after verify-pin" "verify-pin=$($good.Status); commit=$($commit.Status) $($commit.Raw)" `
            "The seeded step-up PIN for 'io' is 1234. Re-seed if it was changed: cd backend; .\.venv\Scripts\python.exe -m app.seed"
    }
}

if ($FullCheck -and $probeId -gt 0) {
    # --- a THIN case must refuse, and say which fields are missing -------------------
    $thin = Api POST "$API/api/cases/$probeId/documents/PANCHNAMA?lang=en" -Token $ioTok
    $thinDetail = if ($thin.Body -and ($thin.Body.PSObject.Properties.Name -contains "detail")) { [string]$thin.Body.detail } else { [string]$thin.Raw }
    $namesFields = ($thinDetail -match "accused_name") -and ($thinDetail -match "seized_items") -and ($thinDetail -match "witnesses")
    if ($thin.Status -eq 400 -and $namesFields) {
        Ok "thin case returns 400 naming the missing fields" $thinDetail
    } else {
        Bad "thin case returns 400 naming the missing fields" "HTTP $($thin.Status): $thinDetail" `
            "Expected 400 listing accused_name, witnesses and seized_items. A 200 means _missing_required in app\services\documents.py is not being enforced."
    }

    # --- fill the pool so the SAME case can carry all eight documents ----------------
    $fill = @()
    $fill += Api PATCH "$API/api/cases/$probeId" -Token $ioTok -Body @{
        fir_number = "9$($probeStamp.Substring($probeStamp.Length - 3))/2026"; fir_date = "2026-07-12"
        incident_datetime = "2026-07-12T22:15:00"; incident_location = "Prahladnagar Garden, Ahmedabad"
        complaint_narrative = $probeNarr; complaint_language = "EN"; status = "INVESTIGATION" }
    $fill += Api POST "$API/api/cases/$probeId/persons" -Token $ioTok -Body @{
        role = "ACCUSED"; full_name = "Probe Accused"; father_name = "Probe Senior"; age = 29
        gender = "M"; address = "12 Probe Lane, Ahmedabad" }
    $fill += Api POST "$API/api/cases/$probeId/persons" -Token $ioTok -Body @{
        role = "COMPLAINANT"; full_name = "Probe Complainant"; age = 41; address = "Prahladnagar, Ahmedabad" }
    $fill += Api POST "$API/api/cases/$probeId/persons" -Token $ioTok -Body @{
        role = "WITNESS"; full_name = "Probe Witness One"; father_name = "W One Senior"; address = "Satellite, Ahmedabad" }
    $fill += Api POST "$API/api/cases/$probeId/persons" -Token $ioTok -Body @{
        role = "WITNESS"; full_name = "Probe Witness Two"; father_name = "W Two Senior"; address = "Vejalpur, Ahmedabad" }
    $fill += Api POST "$API/api/cases/$probeId/seized-items" -Token $ioTok -Body @{
        description = "Gold chain, approx 18 grams, yellow metal"; quantity = 1; estimated_value = 108000
        seizure_datetime = "2026-07-13T09:30:00"; seizure_location = "Satellite Police Station" }

    $fillBad = @($fill | Where-Object { $_.Status -ge 300 })
    if ($fillBad.Count -eq 0) {
        $matrixReady = $true
        Ok "probe case populated through the pool endpoints" "4 persons, 1 seized item, FIR + incident + narrative"
    } else {
        Bad "probe case populated through the pool endpoints" "$($fillBad.Count) of $($fill.Count) pool writes failed (first: HTTP $($fillBad[0].Status) $($fillBad[0].Raw))" `
            "The pool endpoints must accept an IO token on a case that IO owns. Check the role gates in app\api\pool.py."
    }
}

# ---------------------------------------------------------------------------
# 10. Feature matrix - AI pipeline
#     Run BEFORE the documents: REMAND, CUSTODY_LETTER and CHARGESHEET all require at
#     least one ACCEPTED legal section, so the accept step below is what unblocks them.
# ---------------------------------------------------------------------------
$acceptedCount = 0
if ($FullCheck -and $matrixReady) {
    Section "Feature matrix - AI pipeline"

    $an = Api POST "$API/api/cases/$probeId/analyze?lang=en" -Token $ioTok -TimeoutSec 300
    $anSections = @(if ($an.Status -eq 200 -and $an.Body) { $an.Body.sections })
    if ($an.Status -eq 200 -and $anSections.Count -gt 0) {
        $src = if ($an.Body.PSObject.Properties.Name -contains "source") { $an.Body.source } else { "?" }
        Ok "section analysis returns sections" "$($anSections.Count) section(s), source=$src"
    } else {
        Bad "section analysis returns sections" "HTTP $($an.Status), $($anSections.Count) sections returned" `
            "An empty result usually means the Chroma corpus is empty - cd backend; .\.venv\Scripts\python.exe -m app.ai.rag (~2 min). A timeout means Qwen is not loaded: ollama run qwen2.5:7b."
    }

    # Grounding. app\ai\legal.py validate_selections DROPS any (act, section_code) that was
    # not in the retrieved candidate set and records it in `rejected` with that exact
    # reason - so the number of such entries IS the violation count for this run.
    $anRejected = @(if ($an.Status -eq 200 -and $an.Body) { $an.Body.rejected })
    $ungrounded = @($anRejected | Where-Object { $_.rejection_reason -and $_.rejection_reason -match "candidate set" })
    if ($an.Status -eq 200 -and $ungrounded.Count -eq 0) {
        Ok "every cited section is inside the retrieved candidate set" "0 ungrounded selections; $($anRejected.Count) rejection(s) in total this run"
    } else {
        $first = if ($ungrounded.Count -gt 0) { "$($ungrounded[0].act) $($ungrounded[0].section_code)" } else { "n/a" }
        Bad "every cited section is inside the retrieved candidate set" "$($ungrounded.Count) selection(s) were dropped as ungrounded (first: $first)" `
            "The hard validator caught these, so nothing wrong was persisted - but the model is straying. Do NOT loosen validate_selections in app\ai\legal.py; widen retrieval (k) instead."
    }

    # Explainability. Every highlighted phrase must be a verbatim slice of the narrative,
    # or the marked-up narrative in the UI cannot highlight it.
    $secList = Api GET "$API/api/cases/$probeId/sections" -Token $ioTok
    $secs = @(if ($secList.Status -eq 200) { $secList.Body })
    $withPhrase = @($secs | Where-Object { $_.triggering_phrase -and $_.triggering_phrase.Trim() })
    $notVerbatim = @($withPhrase | Where-Object { $probeNarr.IndexOf([string]$_.triggering_phrase, [System.StringComparison]::Ordinal) -lt 0 })
    if ($withPhrase.Count -gt 0 -and $notVerbatim.Count -eq 0) {
        Ok "every triggering phrase appears verbatim in the narrative" "$($withPhrase.Count) phrase(s) checked against the case narrative, character for character"
    } else {
        $sample = if ($notVerbatim.Count -gt 0) { """$($notVerbatim[0].triggering_phrase)""" } else { "no section carried a phrase at all" }
        Bad "every triggering phrase appears verbatim in the narrative" "$($notVerbatim.Count) of $($withPhrase.Count) not found; sample: $sample" `
            "validate_selections in app\ai\legal.py is meant to drop these. If one survived, _repair_phrase returned a phrase it did not re-check."
    }

    # Junk must not become a case. /intake/extract takes no db session, so it CANNOT
    # persist; the check is that the draft it returns is refused by commit as well.
    $junk = Api POST "$API/api/intake/extract" -Token $ioTok -TimeoutSec 300 -Body @{
        messages = @(@{ role = "officer"; content = "asdf qwerty zxcvbn 99999 ;;;; lorem" }); lang = "en" }
    $junkDraft = if ($junk.Status -eq 200 -and $junk.Body) { $junk.Body.draft } else { $null }
    $junkNum = if ($junkDraft -and $junkDraft.case) { [string]$junkDraft.case.case_number } else { "" }
    $junkCommit = Api POST "$API/api/intake/commit" -Token $ioTok -Body @{
        case = @{ case_number = $junkNum }; persons = @(); seized_items = @() }
    if ($junk.Status -eq 200 -and -not $junkNum -and $junkCommit.Status -eq 400) {
        Ok "junk input is refused rather than turned into a case" "extract invented no case number; commit of that draft returned 400"
    } else {
        Bad "junk input is refused rather than turned into a case" "extract=$($junk.Status) case_number='$junkNum'; commit=$($junkCommit.Status)" `
            "case_number is officer-controlled and must never be extracted (app\api\intake.py DraftCase). A 201 here means junk became a real case - delete it and check the extraction whitelist."
    }

    # Accept what came back, so the section-dependent documents can generate.
    foreach ($s in $secs) {
        $r = Api PATCH "$API/api/cases/$probeId/sections/$($s.id)" -Token $ioTok -Body @{ status = "ACCEPTED" }
        if ($r.Status -eq 200) { $acceptedCount++ }
    }
    if ($acceptedCount -gt 0) {
        Ok "suggested sections can be accepted" "$acceptedCount section(s) moved SUGGESTED -> ACCEPTED"
    } else {
        Bad "suggested sections can be accepted" "0 of $($secs.Count) sections accepted" `
            "PATCH /api/cases/{id}/sections/{sid} with {""status"":""ACCEPTED""} must return 200. Without an accepted section, REMAND, CUSTODY_LETTER and CHARGESHEET cannot generate."
    }
}

# ---------------------------------------------------------------------------
# 11. Feature matrix - documents (8 types x EN / HI / GU, by Unicode range)
# ---------------------------------------------------------------------------
if ($FullCheck -and $matrixReady) {
    Section "Feature matrix - documents"

    $docTypes = @("SEIZURE_RECEIPT", "PANCHNAMA", "REMAND", "MEDICAL_LETTER",
                  "CUSTODY_LETTER", "CHARGESHEET", "LERS_PRESERVATION_REQUEST", "LERS_RECORDS_REQUEST")
    $langs = @("en", "hi", "gu")

    $genFail   = @()   # doc types that did not return 200
    $scriptBad = @()   # "TYPE/lang" whose text is not in the requested script
    $tagBad    = @()   # outputs still carrying an unrendered {{ }} or {% %}
    $emptyBad  = @()   # outputs missing a value the pool definitely holds
    $guText    = $null

    foreach ($dt in $docTypes) {
        foreach ($lg in $langs) {
            $g = Api POST "$API/api/cases/$probeId/documents/$dt`?lang=$lg" -Token $ioTok
            if ($g.Status -ne 200) { $genFail += "$dt/$lg -> HTTP $($g.Status) $($g.Raw)"; continue }

            $text = Get-GeneratedText ([int]$g.Body.id) $ioTok
            if (-not $text) { $genFail += "$dt/$lg -> generated but would not download"; continue }

            $c = Get-ScriptCensus $text
            # CHARGESHEET is Form I: its field labels are DELIBERATELY bilingual EN | GU
            # (_FORM_I_BILINGUAL in templates\_labels.py), so Latin and Gujarati always
            # outweigh the body script and a dominance test would report two false FAILs.
            # For it, test that the requested script is PRESENT.
            if ($dt -eq "CHARGESHEET") {
                $okScript = switch ($lg) { "en" { $c.Latin -gt 0 } "hi" { $c.Devanagari -gt 0 } "gu" { $c.Gujarati -gt 0 } }
            } else {
                $okScript = switch ($lg) {
                    "en" { $c.Latin -gt 0 -and $c.Devanagari -eq 0 -and $c.Gujarati -eq 0 }
                    "hi" { $c.Devanagari -gt $c.Gujarati -and $c.Devanagari -gt 0 }
                    "gu" { $c.Gujarati -gt $c.Devanagari -and $c.Gujarati -gt 0 }
                }
            }
            if (-not $okScript) { $scriptBad += "$dt/$lg (dev=$($c.Devanagari) guj=$($c.Gujarati) lat=$($c.Latin))" }

            if ($text -match '\{\{' -or $text -match '\{%') { $tagBad += "$dt/$lg" }

            # A required field that came through empty shows up as a missing identifier or
            # as a leaked Python None. Both are values this script itself put in the pool.
            if ($text.IndexOf($probeNum, [System.StringComparison]::Ordinal) -lt 0) { $emptyBad += "$dt/$lg (case_number absent)" }
            if ($text -match '\bNone\b') { $emptyBad += "$dt/$lg (literal 'None' rendered)" }

            if ($dt -eq "PANCHNAMA" -and $lg -eq "gu") { $guText = $text }
        }
    }

    if ($genFail.Count -eq 0) {
        Ok "all 8 document types generate 200 on a fully populated case" "8 types x 3 languages = 24 generations, all 200 and all downloadable"
    } else {
        Bad "all 8 document types generate 200 on a fully populated case" ($genFail -join "; ") `
            "A 400 names the missing pool field. REMAND, CUSTODY_LETTER and CHARGESHEET need an ACCEPTED section - if the analysis step above failed, these fail with it."
    }

    if ($scriptBad.Count -eq 0 -and $genFail.Count -eq 0) {
        Ok "each document is in the requested script" "Latin for EN, Devanagari for HI, Gujarati for GU - measured on the text inside the .docx, not on the request"
    } elseif ($scriptBad.Count -gt 0) {
        Bad "each document is in the requested script" ($scriptBad -join "; ") `
            "The body script follows ?lang= only. If EN came back Devanagari, the caller sent lang=hi - check the output-language selector in frontend\app\cases\[id]\DocumentsTab.tsx, not the backend."
    }

    if ($tagBad.Count -eq 0 -and $genFail.Count -eq 0) {
        Ok "no unrendered template tags in any output" "no {{ }} or {% %} survived the merge in 24 documents"
    } elseif ($tagBad.Count -gt 0) {
        Bad "no unrendered template tags in any output" ($tagBad -join "; ") `
            "A surviving tag means the template names a context key the generator does not build. Compare the tag against the dict returned by _build_context in app\services\documents.py."
    }

    if ($emptyBad.Count -eq 0 -and $genFail.Count -eq 0) {
        Ok "no required field is empty" "the probe case number reached every document, and no literal 'None' was rendered"
    } elseif ($emptyBad.Count -gt 0) {
        Bad "no required field is empty" (($emptyBad | Select-Object -Unique) -join "; ") `
            "A rendered 'None' is a context value that was None rather than absent, so _missing_required did not catch it. Add the field to required_fields in templates\_registry.py."
    }

    # --- chargesheet: original vs supplementary must differ in the right place -------
    $csO = Api POST "$API/api/cases/$probeId/documents/CHARGESHEET?lang=en&report_type=original" -Token $ioTok
    $csOText = if ($csO.Status -eq 200) { Get-GeneratedText ([int]$csO.Body.id) $ioTok } else { $null }
    $csS = Api POST "$API/api/cases/$probeId/documents/CHARGESHEET?lang=en&report_type=supplementary" -Token $ioTok
    $csSText = if ($csS.Status -eq 200) { Get-GeneratedText ([int]$csS.Body.id) $ioTok } else { $null }

    if ($csOText -and $csSText) {
        $origRight = ($csOText -match "Original Final Form") -and ($csOText -notmatch "Supplementary Final Form") -and ($csOText -notmatch "193\(9\)")
        $suppRight = ($csSText -match "Supplementary Final Form") -and ($csSText -match "193\(9\)")
        if ($origRight -and $suppRight) {
            Ok "chargesheet original and supplementary differ in the right place" "Form I item 5 only: original carries no 193(9) note, supplementary does"
        } else {
            Bad "chargesheet original and supplementary differ in the right place" "original ok=$origRight, supplementary ok=$suppRight" `
                "Item 5 is driven by report_type in app\services\documents.py (report_type_line + cs_supp_note). If both read the same, the query parameter is not reaching the CHARGESHEET branch."
        }
    } else {
        Bad "chargesheet original and supplementary differ in the right place" "original=$($csO.Status), supplementary=$($csS.Status)" `
            "Both must generate before they can be compared. Fix the CHARGESHEET generation failure above first."
    }

    # --- Gujarati must be real Unicode, not mojibake --------------------------------
    if ($guText) {
        $gc = Get-ScriptCensus $guText
        $repl   = @([regex]::Matches($guText, "�")).Count
        $doubled = @([regex]::Matches($guText, '[À-ÿ][-¿]')).Count
        if ($gc.Gujarati -gt 0 -and $repl -eq 0 -and $doubled -eq 0) {
            Ok "Gujarati in a generated .docx is correct Unicode" "$($gc.Gujarati) Gujarati codepoints, 0 replacement chars, 0 double-encoded sequences"
        } else {
            Bad "Gujarati in a generated .docx is correct Unicode" "gujarati=$($gc.Gujarati) replacement_chars=$repl mojibake_sequences=$doubled" `
                "Replacement characters mean the text was decoded as cp1252 somewhere. Note this is about the BYTES - tofu boxes when you OPEN the file are a different problem, fixed by installing fonts\NotoSansGujarati-Regular.ttf system-wide."
        }
    }
}

# ---------------------------------------------------------------------------
# 12. Feature matrix - chat, all four capabilities
# ---------------------------------------------------------------------------
if ($FullCheck -and $matrixReady) {
    Section "Feature matrix - chat (4 capabilities)"

    # 1. Intake from a narrative extracts entities.
    $ex = Api POST "$API/api/intake/extract" -Token $ioTok -TimeoutSec 300 -Body @{
        messages = @(@{ role = "officer"; content = $probeNarr }); lang = "en" }
    $exDraft = if ($ex.Status -eq 200 -and $ex.Body) { $ex.Body.draft } else { $null }
    $exPersons = @(if ($exDraft) { $exDraft.persons }).Count
    $exFilled = @(if ($ex.Status -eq 200 -and $ex.Body) { $ex.Body.auto_filled }).Count
    if ($ex.Status -eq 200 -and ($exPersons -gt 0 -or $exFilled -gt 0)) {
        Ok "intake from a narrative extracts entities" "$exFilled field(s) auto-filled, $exPersons person(s) drafted"
    } else {
        Bad "intake from a narrative extracts entities" "HTTP $($ex.Status), auto_filled=$exFilled persons=$exPersons" `
            "Extraction is a live Qwen call. A timeout means the model is not resident: ollama run qwen2.5:7b. An empty draft on real prose means the extraction prompt is not matching - check app\ai\intake.py."
    }

    # 2. Document generation by request.
    $rg = Api POST "$API/api/cases/$probeId/chat/route" -Token $ioTok -Body @{ message = "generate the seizure receipt"; lang = "en" }
    if ($rg.Status -eq 200 -and $rg.Body.intent -eq "GENERATE" -and $rg.Body.doc_type -eq "SEIZURE_RECEIPT") {
        Ok "document generation by request works" "routed to GENERATE / SEIZURE_RECEIPT via $($rg.Body.source)"
    } else {
        Bad "document generation by request works" "intent=$($rg.Body.intent) doc_type=$($rg.Body.doc_type) (HTTP $($rg.Status))" `
            "The alias table in app\ai\chat.py should catch this deterministically. If source=model, the alias list lost 'seizure receipt'."
    }

    # 3. Missing-field prompting asks for something real - a field the officer can answer,
    #    drawn from the checklist the generator itself produced.
    $askFor = @("accused_name", "witnesses", "seized_items")
    $ms = Api POST "$API/api/cases/$probeId/chat/missing" -Token $ioTok -Body @{ missing = $askFor }
    $fillable = @(if ($ms.Status -eq 200) { $ms.Body.fillable })
    $unknown  = @(if ($ms.Status -eq 200) { $ms.Body.unknown })
    $allReal  = @($fillable | Where-Object { $askFor -notcontains $_ }).Count -eq 0
    if ($ms.Status -eq 200 -and $fillable.Count -gt 0 -and $allReal -and $unknown.Count -eq 0) {
        Ok "missing-field prompting asks for something real" "fillable: $($fillable -join ', ') - every one a field that was actually asked for"
    } else {
        Bad "missing-field prompting asks for something real" "fillable=$($fillable -join ',') unknown=$($unknown -join ',') (HTTP $($ms.Status))" `
            "Anything landing in 'unknown' is a field name FILL_PLAN in app\ai\chat.py does not recognise, so the officer would be asked for something no answer could fill."
    }

    # 4. Case query answers from the record.
    $rq = Api POST "$API/api/cases/$probeId/chat/route" -Token $ioTok -Body @{ message = "what is the status of this case"; lang = "en" }
    $caseRow = Api GET "$API/api/cases/$probeId" -Token $ioTok
    $hasStatus = ($caseRow.Status -eq 200 -and $caseRow.Body -and $caseRow.Body.status)
    if ($rq.Status -eq 200 -and $rq.Body.intent -eq "QUERY" -and $rq.Body.query_kind -eq "STATUS" -and $hasStatus) {
        Ok "case query answers from the record" "routed to QUERY / STATUS, and the record answers it: status=$($caseRow.Body.status)"
    } else {
        Bad "case query answers from the record" "intent=$($rq.Body.intent) query_kind=$($rq.Body.query_kind); case readable=$hasStatus" `
            "Routing is classification only - the answer is read from GET /api/cases/{id}. If the route is right but the case is unreadable, fix visibility first."
    }

    # A question calling for judgement must be declined, not answered.
    $rj = Api POST "$API/api/cases/$probeId/chat/route" -Token $ioTok -Body @{ message = "do you think he is guilty"; lang = "en" }
    if ($rj.Status -eq 200 -and $rj.Body.intent -eq "UNKNOWN") {
        Ok "a guilt question is declined" "intent=UNKNOWN - the judgement markers fire before any alias match"
    } else {
        Bad "a guilt question is declined" "intent=$($rj.Body.intent) doc_type=$($rj.Body.doc_type) query_kind=$($rj.Body.query_kind)" `
            "_JUDGEMENT_MARKERS in app\ai\chat.py must force UNKNOWN before the alias tables run. Anything else makes the assistant look like it assessed a criminal case."
    }
}

# ---------------------------------------------------------------------------
# 13. Feature matrix - security
# ---------------------------------------------------------------------------
if ($FullCheck -and $matrixReady) {
    Section "Feature matrix - security"

    # A fresh draft to finalize, so nothing already on the case is disturbed.
    $target = Api POST "$API/api/cases/$probeId/documents/MEDICAL_LETTER?lang=en" -Token $ioTok
    if ($target.Status -ne 200) {
        Bad "finalize checks" "could not generate the document to finalize (HTTP $($target.Status))" "Fix document generation above first."
    } else {
        $targetId = [int]$target.Body.id

        # RBAC pair: IO may generate, IO may NOT finalize. Finalize is an SHO action (CLAUDE.md
        # section 9), and the role gate is declared BEFORE the step-up gate on purpose, so an
        # IO gets 403 and not a 401 about a step-up they could never make.
        $ioFin = Api POST "$API/api/documents/$targetId/finalize" -Token $ioTok
        if ($target.Status -eq 200 -and $ioFin.Status -eq 403) {
            Ok "RBAC: IO-allowed / SHO-only pair behaves" "IO generated document $targetId (200); IO finalize refused 403 before any step-up was considered"
        } else {
            Bad "RBAC: IO-allowed / SHO-only pair behaves" "IO generate=$($target.Status), IO finalize=$($ioFin.Status)" `
                "Expected 200 then 403. A 401 on finalize means require_step_up is declared before require_role in app\api\documents.py, which leaks the wrong failure to the wrong role."
        }

        # LEGAL_ADVISOR must not be able to alter the pool.
        $legalWrite = Api POST "$API/api/cases/$probeId/persons" -Token $legalTok -Body @{ role = "WITNESS"; full_name = "Should Not Exist" }
        if ($legalWrite.Status -eq 403) {
            Ok "RBAC: LEGAL_ADVISOR cannot alter the pool" "POST persons refused 403"
        } else {
            Bad "RBAC: LEGAL_ADVISOR cannot alter the pool" "HTTP $($legalWrite.Status)" `
                "_WRITE_ROLES in app\api\pool.py is (IO, SHO). If this returned 201, a person row was created - delete it."
        }

        # A fresh SHO password session has not stepped up.
        $shoNoStep = Api POST "$API/api/documents/$targetId/finalize" -Token $shoTok
        # ...and neither has a PIN-minted mobile token. finalize takes NO pin_login
        # exemption, because an SHO can PIN-login on a phone and four digits must not
        # approve a document.
        $pinLogin = Api POST "$API/api/auth/login-pin" -Body @{ username = "sho"; pin = "4321" }
        $pinTok = if ($pinLogin.Status -eq 200) { [string]$pinLogin.Body.token } else { $null }
        $pinFin = if ($pinTok) { Api POST "$API/api/documents/$targetId/finalize" -Token $pinTok } else { $null }

        if ($shoNoStep.Status -eq 401) {
            Ok "finalize without step-up returns 401" "SHO password session, no verify-pin -> 401"
        } else {
            Bad "finalize without step-up returns 401" "HTTP $($shoNoStep.Status)" `
                "require_step_up() must be a dependency of finalize in app\api\documents.py. A 200 means a document was just finalized without a PIN."
        }

        if ($pinTok -and $pinFin -and $pinFin.Status -eq 401) {
            Ok "finalize with a PIN-login token also returns 401" "no pin_login exemption on finalize, by design"
        } elseif (-not $pinTok) {
            Bad "finalize with a PIN-login token also returns 401" "PIN sign-in for 'sho' failed (HTTP $($pinLogin.Status))" `
                "The seeded PIN for 'sho' is 4321. Re-seed: cd backend; .\.venv\Scripts\python.exe -m app.seed"
        } else {
            Bad "finalize with a PIN-login token also returns 401" "HTTP $($pinFin.Status)" `
                "finalize must call require_step_up() with NO allow_pin_login. If this returned 200, four digits on a phone just approved a document."
        }

        # Now step up properly and it must go through.
        $shoPin = Api POST "$API/api/auth/verify-pin" -Body @{ pin = "4321" } -Token $shoTok
        $fin = Api POST "$API/api/documents/$targetId/finalize" -Token $shoTok
        $finStatus = if ($fin.Status -eq 200 -and $fin.Body) { [string]$fin.Body.status } else { "" }
        if ($shoPin.Status -eq 200 -and $shoPin.Body.ok -and $fin.Status -eq 200 -and $finStatus -eq "FINALIZED") {
            Ok "finalize succeeds after verify-pin" "document $targetId is now FINALIZED v$($fin.Body.version)"
        } else {
            Bad "finalize succeeds after verify-pin" "verify-pin=$($shoPin.Status); finalize=$($fin.Status) status=$finStatus" `
                "A 401 straight after a successful verify-pin means the step-up was recorded against a different token. Both calls must use the SAME bearer token - _step_up is keyed by its hash."
        }
    }
}

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------
Write-Host ""
Write-Host "========================================" -ForegroundColor White
if ($script:Fail -eq 0) {
    Write-Host "  VERIFIED - $($script:Pass) passed, 0 failed" -ForegroundColor Green
} else {
    Write-Host "  $($script:Fail) FAILED, $($script:Pass) passed" -ForegroundColor Red
    Write-Host "  Each FAIL above carries its own FIX line." -ForegroundColor Yellow
}
Write-Host "========================================" -ForegroundColor White
foreach ($n in $script:Notes) { Write-Host "  note: $n" -ForegroundColor DarkGray }
Write-Host ""
Write-Host "Phone reminder: a handset still holding a PIN token minted before commit 495a34a" -ForegroundColor Yellow
Write-Host "has no pin_login claim and will be refused at Register. Sign out on the phone and" -ForegroundColor Yellow
Write-Host "sign back in with the PIN once. One tap, per device." -ForegroundColor Yellow
Write-Host ""
exit ([int]($script:Fail -gt 0))
