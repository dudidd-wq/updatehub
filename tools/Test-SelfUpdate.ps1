<#
    UpdateHub — tools\Test-SelfUpdate.ps1
    End-to-end test of the self-update with Windows PowerShell 5.1 on this PC — WITHOUT touching the installed copy.

    Everything happens in a new folder under C:\Windows\Temp that only administrators can write to: a stand-in for
    "Program Files\UpdateHub", a stand-in data folder, a local web server on 127.0.0.1 that plays the role of GitHub,
    and a throwaway signing key that exists only in memory for the length of the test. The real engine code
    (engine.ps1 next to this folder) does the work; every case checks the status, the row, the reason and the files.

    Run from an elevated Windows PowerShell (the local web server needs it):
        powershell -NoProfile -ExecutionPolicy Bypass -File .\tools\Test-SelfUpdate.ps1
    Result: printed, and saved as tools\Test-SelfUpdate.log.
#>
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$logPath = Join-Path $PSScriptRoot 'Test-SelfUpdate.log'
$isAdmin = (New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) { Write-Host 'Run this from an elevated PowerShell (Run as administrator).' -ForegroundColor Red; exit 1 }
if ($PSVersionTable.PSVersion.Major -ne 5) { Write-Host "This test is for Windows PowerShell 5.1 (this is $($PSVersionTable.PSVersion))." -ForegroundColor Red; exit 1 }

# Test root: admin-only, so no unelevated process can swap the code this elevated test loads.
$root = Join-Path (Join-Path ([Environment]::GetFolderPath('Windows')) 'Temp') ('uh-selftest-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
$sec = New-Object System.Security.AccessControl.DirectorySecurity
$sec.SetAccessRuleProtection($true, $false)
foreach ($sid in 'S-1-5-32-544', 'S-1-5-18') {
    $sec.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule((New-Object System.Security.Principal.SecurityIdentifier($sid)),
        'FullControl', [System.Security.AccessControl.InheritanceFlags]'ContainerInherit, ObjectInherit', 'None', 'Allow')))
}
[void][IO.Directory]::CreateDirectory($root, $sec)
$inst = Join-Path $root 'install'; $remote = Join-Path $root 'remote'; $data = Join-Path $root 'data'
New-Item -ItemType Directory -Force -Path $inst, $remote, $data | Out-Null
$names = @('engine.ps1', 'server.ps1', 'updater.html', 'README.txt', 'run.bat')
foreach ($f in $names + 'settings.json') { Copy-Item -LiteralPath (Join-Path $repo $f) -Destination $inst }
$orig = @{}; foreach ($f in $names) { $orig[$f] = [IO.File]::ReadAllBytes((Join-Path $inst $f)) }
$settingsOrig = [IO.File]::ReadAllBytes((Join-Path $inst 'settings.json'))

# ── the engine's own functions (dot-sourced: it returns before its main part) ──────────────────────────────────────
# The engine sets $ErrorActionPreference = 'Continue' for itself; the test keeps that, so its functions behave as in a run.
. (Join-Path $inst 'engine.ps1') -Mode Scan
# Point EVERYTHING at the test folder before any engine function runs.
$DataDir = $data
$script:AppDir = $inst; $script:InstallDir = $inst
$script:AclError = ''; $script:SettingsError = ''; $script:Settings = $null
$cur = $script:AppVersion
if (-not $script:InstallDir.StartsWith($root) -or -not $DataDir.StartsWith($root) -or -not $script:AppDir.StartsWith($root)) { throw 'safety check: test paths are not in the test folder' }

$rsa = $null; $web = $null; $rs = $null; $sync = [hashtable]::Synchronized(@{})
$results = New-Object System.Collections.ArrayList
try {
    # ── throwaway key (in memory only) ────────────────────────────────────────────────────────────────────────────
    $rsa = New-Object System.Security.Cryptography.RSACryptoServiceProvider(3072)
    $rsa.PersistKeyInCsp = $false
    $pub = $rsa.ExportParameters($false)
    $script:ReleaseKeyN = [Convert]::ToBase64String($pub.Modulus)
    $script:ReleaseKeyE = [Convert]::ToBase64String($pub.Exponent)

    # ── local "GitHub" ────────────────────────────────────────────────────────────────────────────────────────────
    function Get-FreePort { $t = New-Object Net.Sockets.TcpListener([Net.IPAddress]::Loopback, 0); $t.Start(); $p = $t.LocalEndpoint.Port; $t.Stop(); return $p }
    $port = Get-FreePort; do { $deadPort = Get-FreePort } while ($deadPort -eq $port)
    $rs = [runspacefactory]::CreateRunspace(); $rs.Open()
    $web = [powershell]::Create(); $web.Runspace = $rs
    [void]$web.AddScript({
        param($Prefix, $Root, $Sync)
        $l = New-Object Net.HttpListener; $l.Prefixes.Add($Prefix); $l.Start(); $Sync.L = $l
        while ($l.IsListening) {
            try { $c = $l.GetContext() } catch { break }
            $p = Join-Path $Root ($c.Request.Url.AbsolutePath.TrimStart('/').Replace('/', '\'))
            if ([IO.File]::Exists($p)) { $b = [IO.File]::ReadAllBytes($p); $c.Response.StatusCode = 200; $c.Response.ContentLength64 = $b.Length; $c.Response.OutputStream.Write($b, 0, $b.Length) }
            else { $c.Response.StatusCode = 404 }
            $c.Response.Close()
        }
    }).AddArgument("http://127.0.0.1:$port/").AddArgument($remote).AddArgument($sync)
    $webHandle = $web.BeginInvoke()
    for ($i = 0; $i -lt 50 -and -not $sync.L; $i++) { Start-Sleep -Milliseconds 100 }
    if (-not $sync.L) { throw 'local web server did not start' }
    $live = "http://127.0.0.1:$port/"; $script:ReleaseBase = $live

    # ── release builder ───────────────────────────────────────────────────────────────────────────────────────────
    $utf8 = New-Object Text.UTF8Encoding($false)
    function New-Release([string]$Ver, [string]$Notes = 'test', [scriptblock]$BreakEngine) {
        $dir = Join-Path $remote ('v' + $Ver); New-Item -ItemType Directory -Force -Path $dir | Out-Null
        $list = @(); $bytes = @{}
        foreach ($f in $names) {
            $t = $utf8.GetString($orig[$f])                     # a BOM stays in the text as U+FEFF and is written back
            $t = $t.Replace('$script:AppVersion = ''' + $cur + '''', '$script:AppVersion = ''' + $Ver + '''')
            $t = $t.Replace('$ServerVersion = ''' + $cur + '''', '$ServerVersion = ''' + $Ver + '''')
            $t = $t.Replace('<meta name="updatehub-version" content="' + $cur + '">', '<meta name="updatehub-version" content="' + $Ver + '">')
            if ($f -eq 'engine.ps1' -and $BreakEngine) { $t = & $BreakEngine $t }
            $b = $utf8.GetBytes($t); $bytes[$f] = $b
            [IO.File]::WriteAllBytes((Join-Path $dir $f), $b)
            $list += [ordered]@{ name = $f; size = $b.Length; sha256 = (Get-Sha256Hex $b) }
        }
        $m = [ordered]@{ app = 'UpdateHub'; version = $Ver; tag = 'v' + $Ver; released = '2026-01-01'; notes = $Notes; files = $list }
        $body = $utf8.GetBytes(($m | ConvertTo-Json -Depth 5))
        $sig = [Convert]::ToBase64String($rsa.SignData($body, [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1))
        return @{ Signed = [byte[]]($utf8.GetBytes($sig + "`n") + $body); Files = $bytes; Dir = $dir }
    }
    function Publish([byte[]]$Signed) {
        New-Item -ItemType Directory -Force -Path (Join-Path $remote 'main') | Out-Null
        [IO.File]::WriteAllBytes((Join-Path $remote 'main\release.signed'), $Signed)
    }
    function Set-Mode([string]$M) { [IO.File]::WriteAllText((Join-Path $DataDir 'selfupdate-config.json'), ('{"mode":"' + $M + '"}')) }
    function Test-Install($Set, [string[]]$Absent = @()) {
        foreach ($f in $names) {
            $p = Join-Path $inst $f
            if ($Absent -contains $f) { if ([IO.File]::Exists($p)) { return $false }; continue }
            if (-not [IO.File]::Exists($p)) { return $false }
            if ((Get-Sha256Hex ([IO.File]::ReadAllBytes($p))) -ne (Get-Sha256Hex $Set[$f])) { return $false }
        }
        return $true
    }
    $statusPath = Join-Path $DataDir 'selfupdate.json'
    function Get-Events {
        # The engine keeps events.jsonl open for writing: read with FileShare.ReadWrite.
        $p = Join-Path $script:RunDir 'events.jsonl'; if (-not [IO.File]::Exists($p)) { return @() }
        $fs = New-Object IO.FileStream($p, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
        $sr = New-Object IO.StreamReader($fs, (New-Object Text.UTF8Encoding($false)))
        try { return @($sr.ReadToEnd() -split "`n" | Where-Object { $_.Trim() }) } finally { $sr.Dispose() }
    }

    Initialize-Run
    # One case: run the provider exactly as a real run does, then check status file (rewritten), row, reason, files.
    function Invoke-Case([string]$Name, [string]$RunMode, [string]$Expect, [string]$ItemExpect, [string]$ReasonLike, [scriptblock]$Check,
                         [string]$ProvExpect, [switch]$NoStatusWrite) {
        $script:Mode = $RunMode
        New-Prov 'selfupdate' 'UpdateHub עצמו' 1
        $before = [datetime]::MinValue; if ([IO.File]::Exists($statusPath)) { $before = [IO.File]::GetLastWriteTimeUtc($statusPath) }
        $evStart = (Get-Events).Count
        Start-Sleep -Milliseconds 30
        $err = ''
        try { Invoke-Provider 'selfupdate' { param($p) Run-SelfUpdate $p } } catch { $err = $_.Exception.Message }
        $script:CaseEvents = @(Get-Events | Select-Object -Skip $evStart)
        $st = $null; try { $st = [IO.File]::ReadAllText($statusPath) | ConvertFrom-Json } catch { }
        $wrote = [IO.File]::Exists($statusPath) -and [IO.File]::GetLastWriteTimeUtc($statusPath) -gt $before
        $prov = $script:Prov['selfupdate']
        $item = @($prov.items) | Select-Object -First 1
        $why = @()
        if ($NoStatusWrite) { if ($wrote) { $why += 'status file was rewritten' } }
        else {
            if (-not $wrote) { $why += 'status file not rewritten' }
            if (-not $st -or $st.status -ne $Expect) { $why += "status=$(if ($st) { $st.status } else { '-' })" }
        }
        if ($ItemExpect -and (-not $item -or $item.status -ne $ItemExpect)) { $why += "row=$(if ($item) { $item.status } else { '-' })" }
        if ($ReasonLike -and (-not $item -or $item.reason -notlike "*$ReasonLike*")) { $why += "reason='$(if ($item) { $item.reason })'" }
        if ($ProvExpect -and $prov.status -ne $ProvExpect) { $why += "provider=$($prov.status)" }
        if ($Check) { try { if (-not [bool](& $Check)) { $why += 'file/backup check failed' } } catch { $why += 'check: ' + $_.Exception.Message } }
        if ($err) { $why += "error: $err" }
        $pass = ($why.Count -eq 0)
        $line = '{0}  {1,-52} {2}' -f $(if ($pass) { 'PASS' } else { 'FAIL' }), $Name, $(if ($pass) { "[$Expect$(if ($item) { ' / ' + $item.status })]" } else { $why -join '; ' })
        [void]$results.Add($line)
        Write-Host $line -ForegroundColor $(if ($pass) { 'Green' } else { 'Red' })
    }
    $swapped = { @($script:CaseEvents | ForEach-Object { try { ($_ | ConvertFrom-Json).m } catch { } } | Where-Object { $_ -eq 'הוחלף: engine.ps1' }).Count -gt 0 }
    $noTempBackup = { -not [IO.Directory]::Exists((Join-Path $DataDir 'selfupdate\backup-new')) }

    # ── cases ────────────────────────────────────────────────────────────────────────────────────────────────────
    $rCur = New-Release $cur
    Set-Mode 'auto'

    $script:ReleaseBase = "http://127.0.0.1:$deadPort/"
    Invoke-Case 'GitHub unreachable -> quiet note' 'Update' 'unreachable' 'info' 'לא ניתן לבדוק' { Test-Install $orig }
    $script:ReleaseBase = $live

    Publish ($utf8.GetBytes('<html><body>Blocked by your network administrator</body></html>'))
    Invoke-Case 'proxy block page -> quiet note, not an alarm' 'Update' 'unreachable' 'info' 'שאינו קובץ גרסה' { Test-Install $orig }

    $bad = [byte[]]$rCur.Signed.Clone(); $bad[$bad.Length - 5] = $bad[$bad.Length - 5] -bxor 1; Publish $bad
    Invoke-Case 'manifest altered -> signature rejected' 'Update' 'failed' 'failed' 'החתימה' { Test-Install $orig }

    Publish $rCur.Signed
    Invoke-Case 'same version -> current' 'Update' 'current' 'current' '' { Test-Install $orig }

    $rOther = New-Release $cur 'different notes, same version'; Publish $rOther.Signed
    Invoke-Case 'same version, other signed manifest -> refused' 'Update' 'failed' 'failed' 'שונה מזה שכבר אושר' { Test-Install $orig }

    Remove-Item -LiteralPath (Join-Path $inst 'README.txt') -Force
    Publish $rCur.Signed
    Invoke-Case 'missing installed file -> repaired' 'Update' 'repaired' 'updated' 'README.txt' { Test-Install $orig }

    Set-Mode 'notify'
    $r90 = New-Release '9.0'; Publish $r90.Signed
    Invoke-Case 'newer, notify -> available only' 'Update' 'available' 'available' 'גרסה 9.0' { Test-Install $orig }
    Invoke-Case 'newer, scan -> available only' 'Scan' 'available' 'available' 'גרסה 9.0' { Test-Install $orig }

    $r80 = New-Release '8.0'; Publish $r80.Signed
    Invoke-Case 'older than newest seen -> stale' 'Update' 'stale' 'info' 'ישנה' { Test-Install $orig }

    Set-Mode 'auto'
    $r91 = New-Release '9.1'; [IO.File]::WriteAllBytes((Join-Path $r91.Dir 'server.ps1'), $utf8.GetBytes('tampered')); Publish $r91.Signed
    Invoke-Case 'file at tag differs from signed hash -> refused' 'Update' 'failed' 'failed' 'טביעת האצבע' { (Test-Install $orig) -and -not (& $swapped) }

    $r92 = New-Release '9.2'; Remove-Item -LiteralPath (Join-Path $r92.Dir 'updater.html'); Publish $r92.Signed
    Invoke-Case 'file at tag missing (404) -> unavailable' 'Update' 'unavailable' 'info' 'updater.html' { (Test-Install $orig) -and -not (& $swapped) }

    $r93 = New-Release '9.3' 'broken engine' { param($t) $t.Replace("`$ErrorActionPreference = 'Continue'", "`$ErrorActionPreference = 'Continue'`r`nthrow 'selftest: broken engine'") }
    Publish $r93.Signed
    Invoke-Case 'new engine fails to load -> swapped, then restored' 'Update' 'failed' 'failed' 'לא נטענה' {
        (Test-Install $orig) -and (& $swapped) -and (& $noTempBackup) -and -not [IO.Directory]::Exists((Join-Path $DataDir 'selfupdate\backup')) }

    $r94 = New-Release '9.4'; Publish $r94.Signed
    $lock = [IO.File]::Open((Join-Path $inst 'server.ps1'), 'Open', 'Read', 'Read')     # blocks the swap of server.ps1
    try {
        Invoke-Case 'swap blocked mid-way -> swapped, then restored' 'Update' 'failed' 'failed' 'החלפת הקבצים נכשלה' {
            (Test-Install $orig) -and (& $swapped) -and (& $noTempBackup) }
    } finally { $lock.Dispose() }

    Remove-Item -LiteralPath (Join-Path $inst 'run.bat') -Force                         # a file the old version does not have
    $r95 = New-Release '9.5' 'broken engine' { param($t) $t.Replace("`$ErrorActionPreference = 'Continue'", "`$ErrorActionPreference = 'Continue'`r`nthrow 'selftest: broken engine'") }
    Publish $r95.Signed
    Invoke-Case 'failed upgrade removes a file it added' 'Update' 'failed' 'failed' 'לא נטענה' { (Test-Install $orig @('run.bat')) -and (& $swapped) }
    [IO.File]::WriteAllBytes((Join-Path $inst 'run.bat'), $orig['run.bat'])

    $r96 = New-Release '9.6' 'good release'; Publish $r96.Signed
    [IO.File]::WriteAllText((Join-Path $inst 'settings.json'), '{ "providers": { "selfupdate": false }, broken')
    $script:SettingsError = 'test: invalid JSON'
    Invoke-Case 'broken settings.json with "selfupdate": false -> skipped' 'Update' '' '' '' { Test-Install $orig } 'skipped' -NoStatusWrite
    [IO.File]::WriteAllText((Join-Path $inst 'settings.json'), '{ broken')
    Invoke-Case 'broken settings.json -> check only, no install' 'Update' 'available' 'available' 'גרסה 9.6' { Test-Install $orig }
    [IO.File]::WriteAllBytes((Join-Path $inst 'settings.json'), $settingsOrig); $script:SettingsError = ''

    Invoke-Case 'newer, auto -> installed, previous kept as backup' 'Update' 'updated' 'updated' 'עודכן ל-9.6' {
        (Test-Install $r96.Files) -and (& $noTempBackup) -and
        ((Get-Sha256Hex ([IO.File]::ReadAllBytes((Join-Path $DataDir 'selfupdate\backup\engine.ps1')))) -eq (Get-Sha256Hex $orig['engine.ps1'])) }

    Set-Mode 'notify'
    $r97 = New-Release '9.7' 'installed by the button'; Publish $r97.Signed
    Invoke-Case 'notify + "install now" button -> installed' 'SelfUpdate' 'updated' 'updated' 'עודכן ל-9.7' { Test-Install $r97.Files }

    $o = [IO.File]::ReadAllText($statusPath) | ConvertFrom-Json
    $o.failingSince = (Get-Date).AddDays(-31).ToString('s'); [IO.File]::WriteAllText($statusPath, ($o | ConvertTo-Json -Compress))
    $script:ReleaseBase = "http://127.0.0.1:$deadPort/"
    Invoke-Case 'unreachable for 31 days -> visible failure' 'Update' 'failed' 'failed' '31 ימים' { Test-Install $r97.Files }
    $script:ReleaseBase = $live

    Set-Mode 'off'
    Invoke-Case 'mode off -> no check' 'Update' 'off' '' '' { Test-Install $r97.Files } 'skipped'
}
catch {
    [void]$results.Add('FAIL  test aborted: ' + $_.Exception.Message + ' @ ' + $_.InvocationInfo.PositionMessage)
    Write-Host $results[-1] -ForegroundColor Red
}
finally {
    try { if ($sync.L) { $sync.L.Stop() } } catch { }
    try { if ($web) { [void]$web.EndInvoke($webHandle); $web.Dispose() } } catch { }
    try { if ($rs) { $rs.Close() } } catch { }
    if ($rsa) { $rsa.Dispose() }
}

$failed = @($results | Where-Object { $_ -like 'FAIL*' }).Count
$head = "UpdateHub self-update test · engine $cur · PowerShell $($PSVersionTable.PSVersion) · $(Get-Date -Format s) · $($results.Count - $failed)/$($results.Count) passed"
$out = @($head) + $results
if ($failed) { $out += "Test folder kept for inspection: $root  (engine log: $($script:RunDir)\events.jsonl)" }
else { Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue }
[IO.File]::WriteAllLines($logPath, [string[]]$out, (New-Object Text.UTF8Encoding($true)))
Write-Host ''; Write-Host $head -ForegroundColor $(if ($failed) { 'Red' } else { 'Green' })
Write-Host "Saved: $logPath"
