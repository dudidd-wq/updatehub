<#
    UpdateHub — server.ps1
    Local dashboard server for the UpdateHub engine. Windows PowerShell 5.1, no add-ons.
    Listens on 127.0.0.1 only. Always runs elevated, and always from C:\Program Files\UpdateHub:
    started from anywhere else it acts as the installer and relaunches the installed copy.

    GET  /                        dashboard (updater.html) — contains no secret
    GET  /api/state               live state + engine/schedule status          ┐
    GET  /api/events?run=&after=  event log lines after a line number          │ every /api/* call needs
    GET  /api/history             summaries of recent runs                     │ the per-session key in the
    POST /api/run?mode=update|scan                                              │ X-UpdateHub header
    POST /api/cancel                                                            │
    POST /api/reboot                                                            │
    POST /api/schedule?enable=1|0&day=Sunday&time=03:00                        ┘

    server.ps1 -Uninstall   removes the task(s), the shortcuts, Program Files\UpdateHub and ProgramData\UpdateHub.
#>
param([int]$Port = 0, [switch]$Uninstall)

$ErrorActionPreference = 'Continue'
$AppDir = $PSScriptRoot

# ── preconditions ─────────────────────────────────────────────────────────────────────────────────
if ($ExecutionContext.SessionState.LanguageMode -ne 'FullLanguage') {
    Write-Host 'UpdateHub cannot run: PowerShell is in Constrained Language Mode (WDAC / AppLocker / Smart App Control policy).' -ForegroundColor Red
    Read-Host 'Press Enter to close' | Out-Null; exit 3
}
# Paths from OS APIs, not environment variables (a user-level variable could redirect this elevated process).
$Sys32      = [Environment]::SystemDirectory
$PsExe      = Join-Path $Sys32 'WindowsPowerShell\v1.0\powershell.exe'
$Shutdown   = Join-Path $Sys32 'shutdown.exe'
$DataDir    = Join-Path ([Environment]::GetFolderPath('CommonApplicationData')) 'UpdateHub'
$InstallDir = Join-Path ([Environment]::GetFolderPath('ProgramFiles')) 'UpdateHub'
$EnginePath = Join-Path $InstallDir 'engine.ps1'
$UserData   = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'UpdateHub'
# Load modules (ScheduledTasks etc.) only from Program Files / System32, never from the user-writable Documents folder.
$pfRoot = [Environment]::GetFolderPath('ProgramFiles')
$env:PSModulePath = (@($env:PSModulePath -split ';' | Where-Object {
    $_ -and ($_.StartsWith($pfRoot, [StringComparison]::OrdinalIgnoreCase) -or $_.StartsWith($Sys32, [StringComparison]::OrdinalIgnoreCase)) }) -join ';')
$TaskName   = 'UpdateHub'
$AppFiles   = @('engine.ps1', 'settings.json', 'server.ps1', 'updater.html', 'run.bat', 'README.txt')
$script:Notice = ''

$isAdmin = (New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    # Started from the shortcut without its "run as administrator" flag, or from a plain prompt: ask for elevation.
    $a = '-NoProfile -ExecutionPolicy Bypass -File "' + $PSCommandPath + '"'; if ($Uninstall) { $a += ' -Uninstall' }
    Start-Process -FilePath $PsExe -ArgumentList $a -Verb RunAs
    exit
}

# ── helpers shared by install / uninstall / serving ──────────────────────────────────────────────
function New-LockedDir([string]$Path) {
    # Administrators + SYSTEM only. Born locked, so no standard user can pre-create entries or junctions in it.
    $sec = New-Object System.Security.AccessControl.DirectorySecurity
    $sec.SetAccessRuleProtection($true, $false)
    $inh = [System.Security.AccessControl.InheritanceFlags]'ContainerInherit, ObjectInherit'
    foreach ($sid in 'S-1-5-32-544', 'S-1-5-18') {
        $sec.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule((New-Object System.Security.Principal.SecurityIdentifier($sid)), 'FullControl', $inh, 'None', 'Allow')))
    }
    try { [void][IO.Directory]::CreateDirectory($Path, $sec) } catch { New-Item -ItemType Directory -Force -Path $Path | Out-Null }
}

# $true when no broad group (Everyone, Authenticated Users, Users, INTERACTIVE, Guests) may create, change or delete
# the folder's contents or the app files in it.
function Test-FolderSafe([string]$Dir) {
    $broad = @('S-1-1-0', 'S-1-5-11', 'S-1-5-32-545', 'S-1-5-4', 'S-1-5-32-546')
    $mask  = 0x2 -bor 0x4 -bor 0x40 -bor 0x10000 -bor 0x40000 -bor 0x80000 -bor 0x10000000 -bor 0x40000000
    $targets = @($Dir) + @($AppFiles | ForEach-Object { Join-Path $Dir $_ } | Where-Object { Test-Path -LiteralPath $_ })
    foreach ($t in $targets) {
        $acl = Get-Acl -LiteralPath $t -ErrorAction Stop
        foreach ($ace in $acl.GetAccessRules($true, $true, [System.Security.Principal.SecurityIdentifier])) {
            if ($ace.AccessControlType -ne 'Allow') { continue }
            if ($t -eq $Dir -and ($ace.PropagationFlags -band [System.Security.AccessControl.PropagationFlags]::InheritOnly)) { continue }
            if ($broad -contains $ace.IdentityReference.Value -and ([int64]$ace.FileSystemRights -band $mask)) { return $false }
        }
    }
    return $true
}

function Test-EngineRunning {
    # The engine holds the named mutex for its whole life: if it exists, an engine is alive. Atomic, no PID games.
    if ($script:EngineProc -and -not $script:EngineProc.HasExited) { return $true }   # started by us, mutex not created yet
    $m = $null
    if ([Threading.Mutex]::TryOpenExisting('Global\UpdateHub.Engine', [ref]$m)) { $m.Dispose(); return $true }
    return $false
}

function Get-Shortcuts {
    return @{
        Desktop   = Join-Path ([Environment]::GetFolderPath('Desktop'))  'UpdateHub.lnk'
        Menu      = Join-Path ([Environment]::GetFolderPath('Programs')) 'UpdateHub.lnk'
        Uninstall = Join-Path ([Environment]::GetFolderPath('Programs')) 'UpdateHub - הסרה.lnk'
    }
}
function New-AdminShortcut([string]$Lnk, [string]$Arguments) {
    # Points straight at System32 powershell.exe (absolute path resolved here, by the OS API) with the
    # "Run as administrator" flag: no batch file, no PATH, no environment variable on the way to elevation.
    $sh = New-Object -ComObject WScript.Shell
    $s = $sh.CreateShortcut($Lnk)
    $s.TargetPath = $PsExe
    $s.Arguments = $Arguments
    $s.WorkingDirectory = $InstallDir
    $s.IconLocation = $PsExe + ',0'
    $s.Description = 'UpdateHub — עדכון מלא של המחשב'
    $s.Save()
    $b = [IO.File]::ReadAllBytes($Lnk); $b[0x15] = $b[0x15] -bor 0x20; [IO.File]::WriteAllBytes($Lnk, $b)   # SLDF_RUNAS_USER
}

# ── uninstall ─────────────────────────────────────────────────────────────────────────────────────
if ($Uninstall) {
    if (Test-EngineRunning) { Write-Host 'A run is in progress - try again after it finishes.' -ForegroundColor Yellow; Read-Host 'Press Enter' | Out-Null; exit 1 }
    # A dashboard server that is still open would keep InstallDir in use: end it first (only our own server.ps1).
    $srvPath = Join-Path $InstallDir 'server.ps1'
    Get-CimInstance Win32_Process -Filter "Name = 'powershell.exe'" -ErrorAction SilentlyContinue |
        Where-Object { $_.ProcessId -ne $PID -and $_.CommandLine -and $_.CommandLine.IndexOf($srvPath, [StringComparison]::OrdinalIgnoreCase) -ge 0 } |
        ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
    Unregister-ScheduledTask -TaskName $TaskName -TaskPath '\' -Confirm:$false -ErrorAction SilentlyContinue
    Get-ScheduledTask -TaskPath '\UpdateHub\' -ErrorAction SilentlyContinue | Unregister-ScheduledTask -Confirm:$false -ErrorAction SilentlyContinue
    try { $svc = New-Object -ComObject Schedule.Service; $svc.Connect(); $svc.GetFolder('\').DeleteFolder('UpdateHub', 0) } catch { }   # the now-empty task folder
    $sc = Get-Shortcuts; foreach ($k in $sc.Keys) { Remove-Item -LiteralPath $sc[$k] -Force -ErrorAction SilentlyContinue }
    foreach ($d in $DataDir, $UserData) { Remove-Item -LiteralPath $d -Recurse -Force -ErrorAction SilentlyContinue }
    Write-Host ''
    Write-Host '  UpdateHub removed: weekly task, shortcuts, ProgramData\UpdateHub, LocalAppData\UpdateHub.' -ForegroundColor Cyan
    Write-Host "  $InstallDir is removed when this window closes." -ForegroundColor Cyan
    Write-Host '  Left in place on purpose (shared with other tools): the Microsoft.WinGet.Client module, the NuGet provider,' -ForegroundColor DarkGray
    Write-Host '  and the Microsoft Update registration (Windows Update > Advanced > "Receive updates for other Microsoft products").' -ForegroundColor DarkGray
    Read-Host '  Press Enter to finish' | Out-Null
    # This script lives in InstallDir: delete the folder from a separate process after this one exits.
    Set-Location -LiteralPath $Sys32   # a folder that is some process's current directory cannot be removed
    Start-Process -FilePath (Join-Path $Sys32 'cmd.exe') -WindowStyle Hidden -WorkingDirectory $Sys32 `
        -ArgumentList ('/d /s /c ""' + (Join-Path $Sys32 'PING.EXE') + '" -n 4 127.0.0.1 >nul & rmdir /s /q "' + $InstallDir + '""')
    exit 0
}

# ── install: the server only ever SERVES from Program Files ───────────────────────────────────────
# Started from the extracted folder (run.bat), this process is the installer: it copies the app into the admin-only
# Program Files folder, creates the "run as administrator" shortcuts, and hands over to the installed copy.
#   * first install          → copy (the same one-time trust as running any downloaded installer)
#   * later, from a folder no other account can write to (e.g. inside your profile) → copy = update UpdateHub
#   * later, from a folder other accounts CAN write to (e.g. C:\Tools inherits "Authenticated Users: Modify")
#                            → nothing is copied; the installed copy is started as it is
$sameDir = ([IO.Path]::GetFullPath($AppDir).TrimEnd('\') -ieq [IO.Path]::GetFullPath($InstallDir).TrimEnd('\'))
if (-not $sameDir) {
    Get-ChildItem -LiteralPath $AppDir -File -ErrorAction SilentlyContinue | Unblock-File -ErrorAction SilentlyContinue
    $installed = Test-Path -LiteralPath $EnginePath
    $copy = $false
    try { $copy = (-not $installed) -or (Test-FolderSafe $AppDir) } catch { $copy = -not $installed }
    $engineBusy = $false
    try { $m = $null; if ([Threading.Mutex]::TryOpenExisting('Global\UpdateHub.Engine', [ref]$m)) { $m.Dispose(); $engineBusy = $true } } catch { }
    if ($copy -and $engineBusy -and $installed) { Write-Host '  A run is in progress - UpdateHub will be updated next time.' -ForegroundColor Yellow; $copy = $false }
    if ($copy) {
        try {
            if (-not (Test-Path -LiteralPath $InstallDir)) { New-Item -ItemType Directory -Force -Path $InstallDir | Out-Null }   # inherits Program Files ACL
            foreach ($f in $AppFiles) {
                $src = Join-Path $AppDir $f
                if (Test-Path -LiteralPath $src) { Copy-Item -LiteralPath $src -Destination (Join-Path $InstallDir $f) -Force -ErrorAction Stop }
            }
            Write-Host "  Installed to $InstallDir" -ForegroundColor Cyan
        } catch { Write-Host ("Cannot write $InstallDir : " + $_.Exception.Message) -ForegroundColor Red; Read-Host 'Press Enter' | Out-Null; exit 1 }
    } elseif ($installed) {
        Write-Host "  $AppDir can be changed by other accounts on this PC - nothing was copied from it." -ForegroundColor Yellow
        Write-Host '  Starting the installed copy. Use the UpdateHub shortcut on the desktop / Start menu from now on.' -ForegroundColor Yellow
    }
    try {
        $sc = Get-Shortcuts; $srv = '-NoProfile -ExecutionPolicy Bypass -File "' + (Join-Path $InstallDir 'server.ps1') + '"'
        New-AdminShortcut $sc.Desktop $srv
        New-AdminShortcut $sc.Menu $srv
        New-AdminShortcut $sc.Uninstall ($srv + ' -Uninstall')
    } catch { Write-Host ('  Shortcuts were not created: ' + $_.Exception.Message) -ForegroundColor Yellow }
    Start-Process -FilePath $PsExe -WorkingDirectory $InstallDir -ArgumentList ('-NoProfile -ExecutionPolicy Bypass -File "' + (Join-Path $InstallDir 'server.ps1') + '"')
    exit 0
}

# ── from here on: the installed server ───────────────────────────────────────────────────────────
if (-not (Test-Path -LiteralPath $DataDir)) { New-LockedDir $DataDir }
if ($Port -eq 0) {
    $Port = 8765
    try { $s = Get-Content -Raw -Encoding UTF8 (Join-Path $AppDir 'settings.json') | ConvertFrom-Json; if ($s.port) { $Port = [int]$s.port } } catch { }
}
$Base = "http://127.0.0.1:$Port/"
$AllowedHosts = @("127.0.0.1:$Port", "localhost:$Port")
$BootPath = Join-Path $DataDir 'boot.key'                # DataDir is Administrators/SYSTEM only

# Two keys:
#   * boot key    — single use. It travels in the URL the browser is opened with (so it shows up in the browser's
#                   command line and history), and dies the moment the page exchanges it.
#   * session key — handed to the page by POST /api/session in exchange for the boot key; it never appears in a
#                   URL, a command line or a file. Every other /api/* call requires it.
# Another standard account can reach 127.0.0.1 but cannot read DataDir, the owner's browser or the owner's processes.
function New-Key { $b = New-Object byte[] 24; [Security.Cryptography.RandomNumberGenerator]::Create().GetBytes($b); return ([Convert]::ToBase64String($b)).Replace('+', '-').Replace('/', '_').TrimEnd('=') }
function Write-BootKey([string]$Key) {
    # Own protected ACL (Administrators/SYSTEM), independent of DataDir's ACL at this moment.
    $sec = New-Object System.Security.AccessControl.FileSecurity
    $sec.SetAccessRuleProtection($true, $false)
    foreach ($sid in 'S-1-5-32-544', 'S-1-5-18') {
        $sec.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule((New-Object System.Security.Principal.SecurityIdentifier($sid)), 'FullControl', 'Allow')))
    }
    Remove-Item -LiteralPath $BootPath -Force -ErrorAction SilentlyContinue
    $fs = [IO.File]::Create($BootPath, 64, [IO.FileOptions]::None, $sec)
    try { $b = [Text.Encoding]::ASCII.GetBytes($Key); $fs.Write($b, 0, $b.Length) } finally { $fs.Dispose() }
}

function Test-PortIsOurs([switch]$Starting) {
    try {
        $ls = @(Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction Stop)
        $foreign = @($ls | Where-Object { $_.OwningProcess -ne 4 }).Count -gt 0                           # 4 = System (http.sys)
        # Our own start: refuse only on positive evidence of a second listener (no false alarm if the query is empty).
        if ($Starting) { return (-not $foreign) }
        if ($ls.Count -eq 0 -or $foreign) { return $false }
        $srv = Join-Path $InstallDir 'server.ps1'
        $mine = @(Get-CimInstance Win32_Process -Filter "Name = 'powershell.exe'" -ErrorAction Stop |
                  Where-Object { $_.ProcessId -ne $PID -and $_.CommandLine -and $_.CommandLine.IndexOf($srv, [StringComparison]::OrdinalIgnoreCase) -ge 0 })
        return ($mine.Count -gt 0)
    } catch { return [bool]$Starting }   # query failed: our own start proceeds, the busy branch opens nothing
}

$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add($Base)
try { $listener.Start() }
catch {
    # Only if the port really belongs to OUR running server (served by http.sys, PID 4, and an installed server.ps1
    # process alive) is it handed a fresh single-use boot key and a tab opened. Anything else — another program, or a
    # plain socket another account bound first to impersonate the dashboard — gets nothing: no key, no browser.
    if (Test-PortIsOurs) {
        Write-Host "UpdateHub is already open. Opening the browser." -ForegroundColor Yellow
        $bk = New-Key; try { Write-BootKey $bk } catch { }
        Start-Process ($Base + '?k=' + $bk)
    } else {
        Write-Host "Port $Port is used by another program - UpdateHub was NOT opened." -ForegroundColor Red
        Write-Host "Choose another port in $InstallDir\settings.json (""port"") and start UpdateHub again." -ForegroundColor Red
        Read-Host 'Press Enter to close' | Out-Null
    }
    exit
}
# A raw socket on 127.0.0.1 next to http.sys would receive the browser's connections: refuse to run beside one.
if (-not (Test-PortIsOurs -Starting)) {
    Write-Host "Another program also listens on port $Port - UpdateHub stopped. Choose another port in settings.json." -ForegroundColor Red
    $listener.Stop(); Read-Host 'Press Enter to close' | Out-Null; exit 1
}
Remove-Item -LiteralPath (Join-Path $DataDir 'session.token') -Force -ErrorAction SilentlyContinue   # v1.2 key file, no longer used

$Token = New-Key
$boot = New-Key; Write-BootKey $boot
$OpenUrl = $Base + '?k=' + $boot

# Console text is English on purpose: the legacy console has no bidi support, so Hebrew renders reversed.
$Host.UI.RawUI.WindowTitle = 'UpdateHub'
Write-Host ''
Write-Host '  UpdateHub is running' -ForegroundColor Cyan
Write-Host "  Dashboard: $Base (opened in your browser)" -ForegroundColor Cyan
Write-Host '  Closing this window closes the dashboard only - a running update continues to the end.' -ForegroundColor DarkGray
Write-Host ''

function Send-Bytes($Ctx, [int]$Code, [string]$Type, [byte[]]$Bytes) {
    $r = $Ctx.Response
    $r.StatusCode = $Code
    $r.ContentType = $Type
    $r.Headers['Cache-Control'] = 'no-store'
    $r.Headers['X-Content-Type-Options'] = 'nosniff'
    $r.Headers['Referrer-Policy'] = 'no-referrer'
    $r.Headers['X-Frame-Options'] = 'DENY'
    $r.ContentLength64 = $Bytes.Length
    $r.OutputStream.Write($Bytes, 0, $Bytes.Length)
    $r.OutputStream.Close()
}
function Send-Json($Ctx, [string]$Json, [int]$Code = 200) { Send-Bytes $Ctx $Code 'application/json; charset=utf-8' ([Text.Encoding]::UTF8.GetBytes($Json)) }
function ConvertTo-JsonString([string]$S) { return ($S | ConvertTo-Json -Compress) }

function Read-Shared([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    for ($i = 0; $i -lt 3; $i++) {
        try {
            $fs = New-Object IO.FileStream($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]'ReadWrite, Delete')
            $sr = New-Object IO.StreamReader($fs, (New-Object Text.UTF8Encoding($false)))
            try { return $sr.ReadToEnd() } finally { $sr.Dispose() }
        } catch { Start-Sleep -Milliseconds 60 }
    }
    return $null
}

$script:SchedCache = $null; $script:SchedAt = [datetime]::MinValue
function Get-ScheduleJson([switch]$Fresh) {
    if (-not $Fresh -and $script:SchedCache -and ((Get-Date) - $script:SchedAt).TotalSeconds -lt 15) { return $script:SchedCache }
    $o = [ordered]@{ enabled = $false; day = 'Sunday'; time = '03:00'; next = ''; last = ''; lastResult = '' }
    $spec = Read-Shared (Join-Path $DataDir 'schedule.json')
    if ($spec) { try { $j = $spec | ConvertFrom-Json; $o.day = [string]$j.day; $o.time = [string]$j.time } catch { } }
    $t = Get-ScheduledTask -TaskName $TaskName -TaskPath '\' -ErrorAction SilentlyContinue
    if ($t) {
        $o.enabled = ([string]$t.State -ne 'Disabled')
        $i = Get-ScheduledTaskInfo -TaskName $TaskName -TaskPath '\' -ErrorAction SilentlyContinue
        if ($i) {
            if ($i.NextRunTime) { $o.next = $i.NextRunTime.ToString('s') }
            if ($i.LastRunTime -and $i.LastRunTime.Year -gt 2000) { $o.last = $i.LastRunTime.ToString('s'); $o.lastResult = ('0x{0:X}' -f $i.LastTaskResult) }
        }
    }
    $script:SchedCache = ($o | ConvertTo-Json -Compress); $script:SchedAt = Get-Date
    return $script:SchedCache
}

function Set-Schedule([bool]$Enable, [string]$Day, [string]$Time) {
    if (-not $Enable) {
        Unregister-ScheduledTask -TaskName $TaskName -TaskPath '\' -Confirm:$false -ErrorAction SilentlyContinue
        return
    }
    if ($Day -notin 'Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday') { $Day = 'Sunday' }
    if ($Time -notmatch '^([01]\d|2[0-3]):[0-5]\d$') { $Time = '03:00' }
    $action    = New-ScheduledTaskAction -Execute $PsExe -Argument ('-NoProfile -NonInteractive -ExecutionPolicy Bypass -WindowStyle Hidden -File "' + $EnginePath + '" -Mode Update -Scheduled')
    $trigger   = New-ScheduledTaskTrigger -Weekly -DaysOfWeek $Day -At ([datetime]::ParseExact($Time, 'HH:mm', $null))
    $user      = [Security.Principal.WindowsIdentity]::GetCurrent().Name
    # Runs as the signed-in user, elevated; the engine itself drops to the normal user token for per-user tools.
    $principal = New-ScheduledTaskPrincipal -UserId $user -LogonType Interactive -RunLevel Highest
    $settings  = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
                    -ExecutionTimeLimit (New-TimeSpan -Hours 12) -MultipleInstances IgnoreNew
    Register-ScheduledTask -TaskName $TaskName -TaskPath '\' -Action $action -Trigger $trigger -Principal $principal -Settings $settings `
        -Description 'UpdateHub — עדכון אוטומטי מלא של המחשב' -Force -ErrorAction Stop | Out-Null
    $spec = [ordered]@{ day = $Day; time = $Time } | ConvertTo-Json -Compress
    [IO.File]::WriteAllText((Join-Path $DataDir 'schedule.json'), $spec, (New-Object Text.UTF8Encoding($false)))
}

$script:EngineProc = $null
function Start-Engine([string]$Mode) {
    if (Test-EngineRunning) { return $false }
    # -NonInteractive: any unexpected prompt fails fast instead of waiting forever in a hidden window.
    # The engine also takes a named mutex as its first act, so even a race here can never run two engines.
    $script:EngineProc = Start-Process -FilePath $PsExe -WindowStyle Hidden -PassThru `
        -ArgumentList ('-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "' + $EnginePath + '" -Mode ' + $Mode)
    return $true
}

function Get-StateJson {
    $state = Read-Shared (Join-Path $DataDir 'state.json')
    if (-not $state -or -not $state.TrimStart().StartsWith('{')) { $state = 'null' }
    $running = 'false'; if (Test-EngineRunning) { $running = 'true' }
    return ('{"engineRunning":' + $running + ',"notice":' + (ConvertTo-JsonString $script:Notice) + ',"schedule":' + (Get-ScheduleJson) + ',"state":' + $state + '}')
}

# events.jsonl is read incrementally: only the new bytes since the last poll, cut at the last complete line, so a
# multi-byte character or a half-written line is never decoded early.
$script:Ev = @{ Run = ''; Pos = 0L; Rest = (New-Object byte[] 0); Lines = (New-Object System.Collections.Generic.List[string]) }
function Update-EvCache([string]$Run) {
    if ($script:Ev.Run -ne $Run) { $script:Ev = @{ Run = $Run; Pos = 0L; Rest = (New-Object byte[] 0); Lines = (New-Object System.Collections.Generic.List[string]) } }
    $path = Join-Path $DataDir ("runs\$Run\events.jsonl")
    if (-not (Test-Path -LiteralPath $path)) { return }
    try {
        $fs = New-Object IO.FileStream($path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]'ReadWrite, Delete')
        try {
            if ($fs.Length -lt $script:Ev.Pos) { $script:Ev.Pos = 0L; $script:Ev.Rest = New-Object byte[] 0; $script:Ev.Lines.Clear() }
            $n = [int]($fs.Length - $script:Ev.Pos); if ($n -le 0) { return }
            [void]$fs.Seek($script:Ev.Pos, [IO.SeekOrigin]::Begin)
            $buf = New-Object byte[] ($script:Ev.Rest.Length + $n)
            [Array]::Copy($script:Ev.Rest, $buf, $script:Ev.Rest.Length)
            $got = $fs.Read($buf, $script:Ev.Rest.Length, $n)
            $script:Ev.Pos += $got
            $len = $script:Ev.Rest.Length + $got
            $cut = [Array]::LastIndexOf($buf, [byte]10, $len - 1)
            if ($cut -lt 0) { $script:Ev.Rest = New-Object byte[] $len; [Array]::Copy($buf, $script:Ev.Rest, $len); return }
            $text = [Text.Encoding]::UTF8.GetString($buf, 0, $cut)
            $restLen = $len - $cut - 1; $script:Ev.Rest = New-Object byte[] $restLen
            if ($restLen -gt 0) { [Array]::Copy($buf, $cut + 1, $script:Ev.Rest, 0, $restLen) }
            foreach ($l in $text.Split("`n")) { $script:Ev.Lines.Add($l.TrimEnd("`r")) }
        } finally { $fs.Dispose() }
    } catch { }
}
function Get-EventsJson([string]$Run, [int]$After) {
    if ($Run -notmatch '^\d{8}-\d{6}$') { return '{"next":0,"events":[]}' }
    Update-EvCache $Run
    $complete = $script:Ev.Lines.Count
    if ($After -lt 0) { $After = 0 }
    if ($After -ge $complete) { return ('{"next":' + $After + ',"events":[]}') }
    $end = [Math]::Min($complete, $After + 4000)
    $slice = New-Object System.Collections.Generic.List[string]
    for ($i = $After; $i -lt $end; $i++) { $l = $script:Ev.Lines[$i]; if ($l.StartsWith('{')) { $slice.Add($l) } }
    return ('{"next":' + $end + ',"events":[' + ($slice -join ',') + ']}')
}

function Get-HistoryJson {
    $dirs = @(Get-ChildItem (Join-Path $DataDir 'runs') -Directory -ErrorAction SilentlyContinue | Sort-Object Name -Descending | Select-Object -First 15)
    $items = @()
    foreach ($d in $dirs) {
        $s = Read-Shared (Join-Path $d.FullName 'summary.json')
        if ($s -and $s.TrimStart().StartsWith('{')) { $items += $s }
    }
    return ('[' + ($items -join ',') + ']')
}

function Test-KeyEquals([string]$Given, [string]$Key = $Token) {
    if (-not $Given -or -not $Key -or $Given.Length -ne $Key.Length) { return $false }
    $d = 0; for ($i = 0; $i -lt $Key.Length; $i++) { $d = $d -bor ([int][char]$Given[$i] -bxor [int][char]$Key[$i]) }   # constant time
    return ($d -eq 0)
}
function Test-Trusted($Req, [string]$Path) {
    # Host allow-list blocks DNS rebinding; Origin check blocks cross-site pages; the per-session key blocks every
    # other local process and account. The page itself (/) holds no secret and is served without the key.
    if ($AllowedHosts -notcontains $Req.UserHostName) { return $false }
    $origin = $Req.Headers['Origin']
    if ($origin -and ($AllowedHosts | ForEach-Object { "http://$_" }) -notcontains $origin) { return $false }
    if ($Path -eq '/api/session') {
        # Only the current single-use boot key (read from its admin-only file) may be exchanged.
        if ($Req.HttpMethod -ne 'POST') { return $false }
        $bk = ''
        try {
            # A boot key lives at most 2 minutes: one left behind in browser history by an aborted launch dies quickly.
            if ((Get-Item -LiteralPath $BootPath -ErrorAction Stop).LastWriteTimeUtc -lt [DateTime]::UtcNow.AddMinutes(-2)) {
                Remove-Item -LiteralPath $BootPath -Force -ErrorAction SilentlyContinue
            } else { $bk = ([IO.File]::ReadAllText($BootPath)).Trim() }
        } catch { }
        return (Test-KeyEquals ([string]$Req.Headers['X-UpdateHub']) $bk)
    }
    if ($Path.StartsWith('/api/')) { return (Test-KeyEquals ([string]$Req.Headers['X-UpdateHub'])) }
    return ($Req.HttpMethod -eq 'GET')
}

function Invoke-Request($Ctx) {
    $req  = $Ctx.Request
    $path = $req.Url.AbsolutePath.ToLowerInvariant()
    if (-not (Test-Trusted $req $path)) { Send-Json $Ctx '{"error":"forbidden","auth":true}' 403; return }

    if ($req.HttpMethod -eq 'GET') {
        switch ($path) {
            { $_ -in '/', '/index.html' } {
                $html = [IO.File]::ReadAllBytes((Join-Path $AppDir 'updater.html'))
                Send-Bytes $Ctx 200 'text/html; charset=utf-8' $html; return
            }
            '/api/state'   { Send-Json $Ctx (Get-StateJson); return }
            '/api/events'  { $a = 0; [void][int]::TryParse([string]$req.QueryString['after'], [ref]$a)
                             Send-Json $Ctx (Get-EventsJson ([string]$req.QueryString['run']) $a); return }
            '/api/history' { Send-Json $Ctx (Get-HistoryJson); return }
            '/favicon.ico' { Send-Bytes $Ctx 204 'image/x-icon' (New-Object byte[] 0); return }
        }
    }
    elseif ($req.HttpMethod -eq 'POST') {
        switch ($path) {
            '/api/session' {
                Remove-Item -LiteralPath $BootPath -Force -ErrorAction SilentlyContinue        # single use
                Send-Json $Ctx ('{"key":' + (ConvertTo-JsonString $Token) + '}'); return
            }
            '/api/run' {
                $mode = 'Update'; if ($req.QueryString['mode'] -eq 'scan') { $mode = 'Scan' }
                if (Start-Engine $mode) { Write-Host ("[{0}] run started: {1}" -f (Get-Date -Format 'HH:mm:ss'), $mode); Send-Json $Ctx '{"ok":true}' }
                else { Send-Json $Ctx '{"ok":false,"error":"ריצה כבר פעילה"}' 409 }
                return
            }
            '/api/cancel' {
                Set-Content -LiteralPath (Join-Path $DataDir 'cancel.flag') -Value (Get-Date -Format 's')
                Send-Json $Ctx '{"ok":true}'; return
            }
            '/api/reboot' {
                if (Test-EngineRunning) { Send-Json $Ctx '{"ok":false,"error":"ריצה פעילה — אתחול רק אחרי סיומה"}' 409; return }
                & $Shutdown /r /t 15 /c "UpdateHub: אתחול לסיום העדכונים" | Out-Null
                Send-Json $Ctx '{"ok":true}'; return
            }
            '/api/schedule' {
                try {
                    Set-Schedule ($req.QueryString['enable'] -eq '1') ([string]$req.QueryString['day']) ([string]$req.QueryString['time'])
                    Send-Json $Ctx ('{"ok":true,"schedule":' + (Get-ScheduleJson -Fresh) + '}')
                } catch { Send-Json $Ctx ('{"ok":false,"error":' + (ConvertTo-JsonString $_.Exception.Message) + '}') 500 }
                return
            }
        }
    }
    Send-Json $Ctx '{"error":"not found"}' 404
}

# A task registered by an older version may still point at a user folder: re-register it on the protected copy.
$oldTask = Get-ScheduledTask -TaskName $TaskName -TaskPath '\' -ErrorAction SilentlyContinue
if ($oldTask -and ([string]$oldTask.Actions[0].Arguments).IndexOf($EnginePath, [StringComparison]::OrdinalIgnoreCase) -lt 0) {
    $d = 'Sunday'; $tm = '03:00'
    $spec = Read-Shared (Join-Path $DataDir 'schedule.json')
    if ($spec) { try { $j = $spec | ConvertFrom-Json; if ($j.day) { $d = [string]$j.day }; if ($j.time) { $tm = [string]$j.time } } catch { } }
    try { Set-Schedule $true $d $tm; Write-Host '  Weekly task re-pointed to the protected copy.' -ForegroundColor Cyan }
    catch { Write-Host ('  Could not update the weekly task: ' + $_.Exception.Message) -ForegroundColor Yellow }
}

# ── main loop (async wait so Ctrl+C still works) ──────────────────────────────────
Start-Process $OpenUrl
try {
    while ($listener.IsListening) {
        $task = $listener.GetContextAsync()
        while (-not $task.AsyncWaitHandle.WaitOne(500)) { }
        $ctx = $task.GetAwaiter().GetResult()
        try { Invoke-Request $ctx }
        catch {
            Write-Host ("error: " + $_.Exception.Message) -ForegroundColor Red
            try { Send-Json $ctx ('{"error":' + (ConvertTo-JsonString $_.Exception.Message) + '}') 500 } catch { }
        }
    }
} finally {
    Remove-Item -LiteralPath $BootPath -Force -ErrorAction SilentlyContinue
    $listener.Stop(); $listener.Close()
}
