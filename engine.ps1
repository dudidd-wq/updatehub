<#
    UpdateHub — engine.ps1
    Vendor-neutral, full-machine updater for Windows 10/11 (Windows PowerShell 5.1, no add-ons).

    Sources (no vendor tools required):
      * Windows Update / Microsoft Update   — WUA COM API directly (incl. optional/BrowseOnly updates)
      * Microsoft Update Catalog by HWID    — drivers + UEFI capsule firmware (BIOS/EC/TBT on OEM machines)
      * NVIDIA driver API                   — GeForce/RTX (catalog lags behind NVIDIA)
      * Office CDN version manifest         — Microsoft 365 Click-to-Run, verified against the channel's latest build
      * winget (Microsoft.WinGet.Client)    — every installed app with a winget source (incl. msstore)
      * Store, PowerShell modules, pip, VS Code extensions, WSL distros, Chocolatey/Scoop/npm, Defender

    Every install is VERIFIED afterwards (version re-read). "Updated" is only reported if the version changed.

    Live state  : %ProgramData%\UpdateHub\state.json
    Event log   : %ProgramData%\UpdateHub\runs\<runId>\events.jsonl   (level 1=summary 2=detail 3=verbose)
    Read by     : server.ps1 + updater.html

    Usage: powershell.exe -NoProfile -ExecutionPolicy Bypass -File engine.ps1 -Mode Update|Scan [-Scheduled]
#>
[CmdletBinding()]
param(
    [ValidateSet('Scan', 'Update')] [string]$Mode = 'Update',
    [string]$DataDir = (Join-Path ([Environment]::GetFolderPath('CommonApplicationData')) 'UpdateHub'),
    [switch]$Scheduled
)

$ErrorActionPreference = 'Continue'
if ($ExecutionContext.SessionState.LanguageMode -ne 'FullLanguage') {
    Write-Host 'UpdateHub cannot run: PowerShell is in Constrained Language Mode (WDAC / AppLocker / Smart App Control).'; exit 3
}
$ProgressPreference    = 'SilentlyContinue'   # Invoke-WebRequest progress bar slows PS 5.1 downloads ~10x
try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch { }
try { [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false) } catch { }
$env:WSL_UTF8 = '1'                            # wsl.exe prints UTF-16 unless told otherwise

$script:AppDir     = $PSScriptRoot
# System paths come from the OS APIs, never from environment variables: user-level variables (ComSpec, SystemRoot,
# ProgramFiles...) override system ones and would let an unelevated process redirect what this elevated engine runs.
$script:Sys32      = [Environment]::SystemDirectory
$script:Cmd        = Join-Path $script:Sys32 'cmd.exe'
$script:SysDrive   = [IO.Path]::GetPathRoot($script:Sys32).TrimEnd('\')
# Module auto-loading: the per-user module folder (Documents\WindowsPowerShell\Modules) comes FIRST in PSModulePath and
# any unelevated process can drop a "PowerShellGet" or "ScheduledTasks" there. This elevated process (and every child
# it starts) loads modules only from Program Files and System32.
$pfRoot = [Environment]::GetFolderPath('ProgramFiles')
$env:PSModulePath = (@($env:PSModulePath -split ';' | Where-Object {
    $_ -and ($_.StartsWith($pfRoot, [StringComparison]::OrdinalIgnoreCase) -or $_.StartsWith($script:Sys32, [StringComparison]::OrdinalIgnoreCase)) }) -join ';')
$script:UA         = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) UpdateHub/1.5'
$script:MuService  = '7971f918-a847-4430-9279-4a52d1efe18d'   # Microsoft Update service ID
$script:CatBase    = 'https://www.catalog.update.microsoft.com'

# ════════════════════════════════════════════════════════════════════════════
#  Settings
# ════════════════════════════════════════════════════════════════════════════
$script:Defaults = @{
    officeForceClose   = $false
    featureUpgrades    = $false
    previewUpdates     = $true
    firmware           = $true
    firmwareRequireAC  = $true
    firmwareMinBattery = 40
    restorePoint       = $true
    autoRebootScheduled= $false
    excludeWinget      = @()
    excludeDevices     = ''
    excludeWuTitle     = ''
    helpUpdateDays     = 7
    keepRuns           = 20
    minFreeGB          = 10
    allowMetered       = $false
}
$script:Settings = $null; $script:SettingsError = ''; $script:AclError = ''; $script:SkipRun = ''; $script:TmpDir = ''
function Import-Settings {
    $path = Join-Path $script:AppDir 'settings.json'
    if (Test-Path $path) {
        try { $script:Settings = Get-Content -Raw -Path $path -Encoding UTF8 | ConvertFrom-Json -ErrorAction Stop }
        catch { $script:Settings = $null; $script:SettingsError = $_.Exception.Message }
    }
}
function Get-Setting([string]$Name) {
    if ($script:Settings -and ($script:Settings.PSObject.Properties.Name -contains $Name)) { return $script:Settings.$Name }
    return $script:Defaults[$Name]
}
function Test-ProviderEnabled([string]$Id) {
    if ($script:Settings -and $script:Settings.PSObject.Properties.Name -contains 'providers') {
        $pv = $script:Settings.providers
        if ($pv -and ($pv.PSObject.Properties.Name -contains $Id)) { return [bool]$pv.$Id }
    }
    return $true
}

# ════════════════════════════════════════════════════════════════════════════
#  State + event log
# ════════════════════════════════════════════════════════════════════════════
$script:Prov     = @{}
$script:Seq      = 0
$script:LastSave = [datetime]::MinValue
$script:EvWriter = $null

function Initialize-Run {
    $script:RunId  = Get-Date -Format 'yyyyMMdd-HHmmss'
    $script:RunDir = Join-Path $DataDir ("runs\" + $script:RunId)
    $script:WorkDir = Join-Path $DataDir 'work'
    New-Item -ItemType Directory -Force -Path $script:RunDir, $script:WorkDir | Out-Null
    $script:StatePath  = Join-Path $DataDir 'state.json'
    $script:CancelFlag = Join-Path $DataDir 'cancel.flag'
    # Empty stdin for every child process: an unexpected prompt gets EOF instead of hanging a hidden run.
    $script:NullIn = Join-Path $script:RunDir 'stdin.empty'
    [IO.File]::WriteAllText($script:NullIn, '')
    $fs = New-Object System.IO.FileStream((Join-Path $script:RunDir 'events.jsonl'), [IO.FileMode]::Append, [IO.FileAccess]::Write, [IO.FileShare]::ReadWrite)
    $script:EvWriter = New-Object System.IO.StreamWriter($fs, (New-Object System.Text.UTF8Encoding($false)))
    $script:EvWriter.AutoFlush = $true
    $script:State = [ordered]@{
        app            = 'UpdateHub'
        version        = '1.5'
        runId          = $script:RunId
        mode           = $Mode
        scheduled      = [bool]$Scheduled
        status         = 'running'
        startedAt      = (Get-Date).ToString('s')
        finishedAt     = ''
        progress       = 0
        current        = ''
        rebootRequired = $false
        error          = ''
        machine        = [ordered]@{}
        counts         = [ordered]@{}
        providers      = (New-Object System.Collections.ArrayList)
    }
}

# Level: 1 = summary, 2 = detail, 3 = verbose.  Kind: info|ok|warn|err|step|cmd|out
function Write-Ev([int]$Level, [string]$Prov, [string]$Msg, [string]$Kind = 'info') {
    $script:Seq++
    $o = [ordered]@{ s = $script:Seq; t = (Get-Date).ToString('HH:mm:ss'); l = $Level; p = $Prov; k = $Kind; m = $Msg }
    if ($script:EvWriter) { try { $script:EvWriter.WriteLine(($o | ConvertTo-Json -Compress)) } catch { } }
    if ($Level -le 2) { Write-Host ("[{0}] {1,-13} {2}" -f $o.t, $Prov, $Msg) }
}

$script:FinalStates = @('current', 'updated', 'reboot', 'failed', 'skipped', 'manual', 'info')
function Update-Progress {
    $final = $script:FinalStates
    if ($Mode -eq 'Scan') { $final = $final + 'available' }
    $sum = 0.0; $wsum = 0.0
    $counts = [ordered]@{}
    foreach ($p in $script:State.providers) {
        $w = [double]$p.weight; $wsum += $w
        switch ($p.status) {
            'running' {
                $items = @($p.items); $n = $items.Count
                if ($n -gt 0) {
                    $fin = @($items | Where-Object { $final -contains $_.status }).Count
                    $sum += $w * (0.1 + 0.85 * ($fin / $n))
                } else { $sum += $w * 0.1 }
            }
            'pending' { }
            default   { $sum += $w }
        }
        foreach ($it in $p.items) {
            if ($counts.Contains($it.status)) { $counts[$it.status]++ } else { $counts[$it.status] = 1 }
        }
    }
    if ($wsum -gt 0) { $script:State.progress = [math]::Round(100.0 * $sum / $wsum, 1) }
    $script:State.counts = $counts
}

$script:StateDirty = $false
function Save-State([switch]$Force) {
    $now = Get-Date
    # Throttled saves are never lost: they mark the state dirty and the next Flush-State (called by every
    # wait loop, download and external process) or forced save writes them.
    if (-not $Force -and ($now - $script:LastSave).TotalMilliseconds -lt 400) { $script:StateDirty = $true; return }
    $script:LastSave = $now
    $script:StateDirty = $false
    Update-Progress
    $json = $script:State | ConvertTo-Json -Depth 8 -Compress
    $tmp = $script:StatePath + '.tmp'
    $enc = New-Object System.Text.UTF8Encoding($false)
    for ($i = 0; $i -lt 5; $i++) {
        try {
            [IO.File]::WriteAllText($tmp, $json, $enc)
            if (Test-Path $script:StatePath) { [IO.File]::Replace($tmp, $script:StatePath, $null) }
            else { [IO.File]::Move($tmp, $script:StatePath) }
            return
        } catch { Start-Sleep -Milliseconds 120 }
    }
    try { [IO.File]::WriteAllText($script:StatePath, $json, $enc) } catch { }   # last resort: non-atomic write
}
function Flush-State { if ($script:StateDirty) { Save-State -Force } }

function New-Prov([string]$Id, [string]$Title, [int]$Weight = 1) {
    $p = [ordered]@{ id = $Id; title = $Title; weight = $Weight; status = 'pending'; summary = ''; items = (New-Object System.Collections.ArrayList) }
    [void]$script:State.providers.Add($p)
    $script:Prov[$Id] = $p
}
function Set-Prov($P, [string]$Status, [string]$Summary) {
    $P.status = $Status
    if ($PSBoundParameters.ContainsKey('Summary')) { $P.summary = $Summary }
    Save-State -Force
}
function Add-Entry($P, [string]$Name, [string]$Source, [string]$Installed, [string]$Available = '', [string]$Status = 'checking') {
    $it = [ordered]@{ name = $Name; source = $Source; installed = $Installed; available = $Available; status = $Status; reason = '' }
    [void]$P.items.Add($it)
    Save-State
    return $it
}
function Set-Entry($It, [string]$Status, [string]$Reason, [string]$Available) {
    $It.status = $Status
    if ($PSBoundParameters.ContainsKey('Reason'))    { $It.reason = $Reason }
    if ($PSBoundParameters.ContainsKey('Available')) { $It.available = $Available }
    if ($Status -eq 'reboot') { $script:State.rebootRequired = $true }
    # Transitions into a long phase must be visible immediately — the next save may be minutes away.
    if ($Status -in 'downloading', 'installing') { Save-State -Force } else { Save-State }
}
function Set-Current([string]$Text) { $script:State.current = $Text; Save-State -Force }

function Get-ProvSummary($P) {
    $names = [ordered]@{ updated = 'עודכנו'; reboot = 'ממתינים לאתחול'; available = 'זמינים'; current = 'מעודכנים'; failed = 'נכשלו'; manual = 'ידני'; skipped = 'דולגו'; info = 'מידע' }
    $parts = @()
    foreach ($k in $names.Keys) {
        $n = @($P.items | Where-Object { $_.status -eq $k }).Count
        if ($n -gt 0) { $parts += "$n $($names[$k])" }
    }
    if ($parts.Count -eq 0) { return 'אין פריטים' }
    return ($parts -join ' · ')
}

function Test-Cancel {
    if ($script:CancelFlag -and (Test-Path $script:CancelFlag)) { throw (New-Object System.OperationCanceledException('בוטל על ידי המשתמש')) }
}
# Inside a generic catch: let a user cancellation keep propagating instead of being recorded as an item failure.
function Test-IsCancel($Err) { return ($Err.Exception -is [System.OperationCanceledException]) }

# A source that fails as a whole gets a visible failed row, so the tiles, toast and history count it —
# a provider marked "failed" with no items would otherwise look like "everything is up to date".
function Set-ProvFailed($P, [string]$Reason) {
    $e = Add-Entry $P $P.title 'שגיאה' ''
    Set-Entry $e 'failed' $Reason
    Set-Prov $P 'failed' $Reason
}

function Invoke-Provider([string]$Id, [scriptblock]$Body) {
    $p = $script:Prov[$Id]
    if (-not (Test-ProviderEnabled $Id)) { Set-Prov $p 'skipped' 'מושבת בהגדרות'; return }
    Test-Cancel
    Set-Prov $p 'running'
    Write-Ev 1 $Id ("▶ " + $p.title) 'step'
    $sw = [Diagnostics.Stopwatch]::StartNew()
    try {
        & $Body $p
        if ($p.status -eq 'running') { Set-Prov $p 'done' (Get-ProvSummary $p) }
        Write-Ev 2 $Id ("הסתיים: {0}  ({1:N0} שניות)" -f $p.summary, $sw.Elapsed.TotalSeconds) 'ok'
    }
    catch [System.OperationCanceledException] { Set-Prov $p 'failed' 'בוטל'; throw }
    catch {
        Write-Ev 1 $Id ('שגיאה: ' + $_.Exception.Message) 'err'
        Write-Ev 3 $Id ([string]$_.ScriptStackTrace) 'err'
        Set-ProvFailed $p ('שגיאה: ' + $_.Exception.Message)
    }
}

# ════════════════════════════════════════════════════════════════════════════
#  Process + HTTP helpers (everything is echoed to the VERBOSE stream)
# ════════════════════════════════════════════════════════════════════════════
function Convert-OutBytes([byte[]]$Bytes) {
    if ($Bytes.Length -ge 4) {
        $zeros = 0; $n = [Math]::Min($Bytes.Length, 400)
        for ($i = 1; $i -lt $n; $i += 2) { if ($Bytes[$i] -eq 0) { $zeros++ } }
        if ($zeros -gt ($n / 4)) { return [Text.Encoding]::Unicode.GetString($Bytes) }
    }
    return [Text.Encoding]::UTF8.GetString($Bytes)
}

function Read-ProcStream($R, [string]$Prov, $All, [switch]$Final) {
    if (Test-Path $R.Path) {
        try {
            $fs = New-Object IO.FileStream($R.Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
            try {
                if ($fs.Length -gt $R.Pos) {
                    [void]$fs.Seek($R.Pos, [IO.SeekOrigin]::Begin)
                    $len = [int]($fs.Length - $R.Pos)
                    $bytes = New-Object byte[] $len
                    $n = $fs.Read($bytes, 0, $len)
                    $R.Pos += $n
                    if ($n -lt $len) { [Array]::Resize([ref]$bytes, $n) }
                    if ($n -gt 0) { $R.Buf += (Convert-OutBytes $bytes) }
                }
            } finally { $fs.Dispose() }
        } catch { }
    }
    $lines = @($R.Buf -split "\r\n|\n|\r")
    if ($Final) { $R.Buf = '' }
    elseif ($lines.Count -le 1) { return }
    else { $R.Buf = $lines[-1]; $lines = $lines[0..($lines.Count - 2)] }
    foreach ($ln in $lines) {
        $t = ($ln -replace '[\x00-\x08\x0B\x0C\x0E-\x1F]', '').TrimEnd()
        if ([string]::IsNullOrWhiteSpace($t)) { continue }
        if ($t -match '^\s*[-\\|/]\s*$') { continue }          # spinner frames
        if ($t -match '^[\s█▒■□\.]+(\d+%)?\s*$') { continue }  # progress bars
        [void]$All.AppendLine($t)
        if ($R.Sb) { [void]$R.Sb.AppendLine($t) }
        if (-not $R.Quiet) { Write-Ev 3 $Prov $t 'out' }
    }
}

# cmd.exe argument string for a .cmd shim:  /d /s /c ""C:\path with spaces\npm.cmd" args "quoted arg""
# /s makes cmd strip exactly the outer pair of quotes, so any number of inner quotes survives intact;
# /d skips AutoRun commands from the registry.
function Get-CmdArgs([string]$Exe, [string]$Rest) { return ('/d /s /c ""' + $Exe + '" ' + $Rest + '"') }

# -Interruptible: read-only commands (lists, probes, expand) that Stop may kill. Everything else runs to completion.
function Invoke-External([string]$File, [string]$Arguments, [string]$Prov, [int]$TimeoutSec = 1800, [string]$WorkDir, [switch]$QuietStdout, [switch]$Interruptible) {
    if (-not $WorkDir) { $WorkDir = $script:RunDir }
    Write-Ev 3 $Prov ("> " + $File + ' ' + $Arguments) 'cmd'
    $tag = [guid]::NewGuid().ToString('N').Substring(0, 8)
    $out = Join-Path $script:RunDir ("p-$tag.out"); $err = Join-Path $script:RunDir ("p-$tag.err")
    try {
        $proc = Start-Process -FilePath $File -ArgumentList $Arguments -WorkingDirectory $WorkDir -NoNewWindow -PassThru `
                    -RedirectStandardOutput $out -RedirectStandardError $err -RedirectStandardInput $script:NullIn -ErrorAction Stop
    } catch {
        Write-Ev 2 $Prov ("לא ניתן להריץ " + $File + ': ' + $_.Exception.Message) 'err'
        return [pscustomobject]@{ ExitCode = -1; Output = ''; StdOut = '' }
    }
    $null = $proc.Handle            # PS 5.1: without touching Handle, ExitCode is lost
    $readers = @(@{ Path = $out; Pos = 0L; Buf = ''; Sb = (New-Object System.Text.StringBuilder); Quiet = [bool]$QuietStdout },
                 @{ Path = $err; Pos = 0L; Buf = ''; Sb = $null; Quiet = $false })
    $all = New-Object System.Text.StringBuilder
    $sw = [Diagnostics.Stopwatch]::StartNew(); $timedOut = $false; $cancelled = $false; $noted = $false
    while ($true) {
        $exited = $proc.HasExited
        foreach ($r in $readers) { Read-ProcStream $r $Prov $all }
        Flush-State
        if ($exited) { break }
        if ($sw.Elapsed.TotalSeconds -gt $TimeoutSec) {
            # Kill the whole tree: installers (NVIDIA setup, msiexec, dpkg) run as children of what we started.
            $timedOut = $true
            $null = & (Join-Path $script:Sys32 'taskkill.exe') /PID $proc.Id /T /F 2>&1
            try { if (-not $proc.HasExited) { $proc.Kill() } } catch { }
            break
        }
        if (Test-Path $script:CancelFlag) {
            if ($Interruptible) { $cancelled = $true; $null = & (Join-Path $script:Sys32 'taskkill.exe') /PID $proc.Id /T /F 2>&1; break }
            # Installers are never killed on Stop: an interrupted pnputil/dpkg/setup leaves the machine half-updated.
            if (-not $noted) { $noted = $true; Write-Ev 2 $Prov 'בקשת עצירה התקבלה — ממתין לסיום ההתקנה הנוכחית' 'warn' }
        }
        Start-Sleep -Milliseconds 400
    }
    Start-Sleep -Milliseconds 150
    foreach ($r in $readers) { Read-ProcStream $r $Prov $all -Final }
    $code = -1
    if ($timedOut) { $code = -2 } else { try { $proc.WaitForExit(5000) | Out-Null; $code = $proc.ExitCode } catch { } }
    Remove-Item $out, $err -Force -ErrorAction SilentlyContinue
    $note = ''; if ($timedOut) { $note = ' (חרג מזמן מקסימלי)' }
    $kind = 'info'; if ($code -ne 0) { $kind = 'warn' }
    Write-Ev 3 $Prov ("exit code: {0}{1}  ({2:N1}s)" -f $code, $note, $sw.Elapsed.TotalSeconds) $kind
    if ($timedOut) { Write-Ev 2 $Prov ("{0} חרג מ-{1} דקות ונעצר" -f (Split-Path $File -Leaf), [int]($TimeoutSec / 60)) 'warn' }
    if ($cancelled) { Test-Cancel }
    return [pscustomobject]@{ ExitCode = $code; Output = $all.ToString(); StdOut = $readers[0].Sb.ToString() }
}

# ── Least privilege for per-user tools ───────────────────────────────────────────────────────────
# Rule: the elevated engine only executes programs that only Administrators/SYSTEM/TrustedInstaller can modify.
# Per-user toolchains (pip of a Python in your profile, VS Code extensions, scoop, npm globals) live in folders
# your own unelevated processes can write to, and they own per-user data: they run with your NORMAL (non-admin)
# token, through a one-shot scheduled task with RunLevel Limited. They gain nothing from admin rights anyway, and
# running them elevated would leave admin-owned files in your profile.
$script:TrustedSids = @('S-1-5-32-544', 'S-1-5-18', 'S-1-5-80-956008885-3418522649-1831038044-1853292631-2271478464')   # Admins, SYSTEM, TrustedInstaller
function Test-AdminOwned([string]$Path) {
    # File and its folder: nobody outside Admins/SYSTEM/TrustedInstaller may write, add files (DLL planting),
    # delete or re-permission. Every further ancestor up to the drive root: nobody else may delete/rename children,
    # delete the folder itself or re-permission it (otherwise the whole folder could be swapped). Owners must be trusted too.
    try {
        $full  = 0x2 -bor 0x4 -bor 0x40 -bor 0x10000 -bor 0x40000 -bor 0x80000 -bor 0x10000000 -bor 0x40000000
        $chain = 0x40 -bor 0x10000 -bor 0x40000 -bor 0x80000 -bor 0x10000000
        $t = [IO.Path]::GetFullPath($Path); $level = 0
        while ($t) {
            $acl = Get-Acl -LiteralPath $t -ErrorAction Stop
            $owner = (New-Object System.Security.Principal.NTAccount($acl.Owner)).Translate([System.Security.Principal.SecurityIdentifier]).Value
            if ($script:TrustedSids -notcontains $owner) { return $false }
            $mask = $full; if ($level -ge 2) { $mask = $chain }
            foreach ($ace in $acl.GetAccessRules($true, $true, [System.Security.Principal.SecurityIdentifier])) {
                if ($ace.AccessControlType -ne 'Allow') { continue }
                if ($level -ge 1 -and ($ace.PropagationFlags -band [System.Security.AccessControl.PropagationFlags]::InheritOnly)) { continue }
                $sid = $ace.IdentityReference.Value
                if ($sid -eq 'S-1-3-0') { continue }                                         # CREATOR OWNER: applies to new children only
                if ($script:TrustedSids -notcontains $sid -and ([int64]$ace.FileSystemRights -band $mask)) { return $false }
            }
            $parent = [IO.Path]::GetDirectoryName($t)
            if (-not $parent -or $parent -eq $t) { break }
            $t = $parent; $level++
        }
        return $true
    } catch { return $false }
}

# Exact organisation (O=) of a signing certificate — not a substring of the whole subject.
function Get-CertOrg($Cert) {
    if (-not $Cert) { return '' }
    $m = [regex]::Match([string]$Cert.Subject, '(?:^|,\s*)O=(?:"([^"]*)"|([^,]*))')
    if (-not $m.Success) { return '' }
    if ($m.Groups[1].Success) { return $m.Groups[1].Value.Trim() }
    return $m.Groups[2].Value.Trim()
}

# Same contract as Invoke-External (ExitCode / Output / StdOut, VERBOSE streaming, time limit, -Interruptible),
# but the program runs with the signed-in user's normal token. ExitCode -3 = could not start as the user.
# No window: the task runs `conhost --headless powershell -EncodedCommand ...`, which starts cmd /c inside an invisible
# console and writes the real exit code to a file. If this Windows has no headless conhost, it falls back (for the
# rest of the run) to a hidden PowerShell window, which may flash briefly.
$script:Headless = $true
# Ends the per-user tool and everything it started (Stop-ScheduledTask alone ends only the wrapper). The PID comes
# from a file the user can write, so it is honoured only for the exact process this call started: a cmd.exe, started
# after the task, whose parent is the wrapper PowerShell carrying this call's own -EncodedCommand.
function Stop-UserTree([string]$PidFile, [datetime]$Since, [string]$Marker) {
    try {
        $n = 0; if (-not [int]::TryParse(([IO.File]::ReadAllText($PidFile)).Trim(), [ref]$n)) { return }
        $c = Get-CimInstance Win32_Process -Filter "ProcessId = $n" -ErrorAction Stop
        if (-not $c -or $c.Name -ne 'cmd.exe' -or $c.CreationDate -lt $Since.AddSeconds(-2)) { return }
        $w = Get-CimInstance Win32_Process -Filter ("ProcessId = " + [int]$c.ParentProcessId) -ErrorAction Stop
        if (-not $w -or $w.Name -ne 'powershell.exe' -or -not $w.CommandLine -or $w.CommandLine.IndexOf($Marker, [StringComparison]::Ordinal) -lt 0) { return }
        # ...and both must belong to the signed-in user: the elevated engine never ends anyone else's processes.
        $me = [Security.Principal.WindowsIdentity]::GetCurrent().Name
        foreach ($proc in $c, $w) {
            $o = Invoke-CimMethod -InputObject $proc -MethodName GetOwner -ErrorAction Stop
            if (("{0}\{1}" -f $o.Domain, $o.User) -ine $me) { return }
        }
        $null = & (Join-Path $script:Sys32 'taskkill.exe') /PID $n /T /F 2>&1
    } catch { }
}
function Invoke-AsUser([string]$File, [string]$Arguments, [string]$Prov, [int]$TimeoutSec = 1800, [switch]$QuietStdout, [switch]$Interruptible, [switch]$NoRetry) {
    $dir = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'UpdateHub\run'
    New-Item -ItemType Directory -Force -Path $dir | Out-Null
    $tag = [guid]::NewGuid().ToString('N'); $out = Join-Path $dir "$tag.out"; $err = Join-Path $dir "$tag.err"; $codeFile = Join-Path $dir "$tag.code"; $pidFile = Join-Path $dir "$tag.pid"
    $line = '/d /s /c ""' + $File + '" ' + $Arguments + ' <nul >"' + $out + '" 2>"' + $err + '""'
    Write-Ev 3 $Prov ("> [הרשאות משתמש רגיל] " + $File + ' ' + $Arguments) 'cmd'
    $name = "UpdateHub-user-$tag"; $tp = '\UpdateHub\'
    $fail = [pscustomobject]@{ ExitCode = -3; Output = ''; StdOut = '' }
    $q = { param($v) "'" + ([string]$v -replace "'", "''") + "'" }
    $inner = '$p = Start-Process -FilePath ' + (& $q $script:Cmd) + ' -ArgumentList ' + (& $q $line) + ' -NoNewWindow -PassThru; ' +
             '[IO.File]::WriteAllText(' + (& $q $pidFile) + ', [string]$p.Id); ' +
             '$null = $p.Handle; $p.WaitForExit(); [IO.File]::WriteAllText(' + (& $q $codeFile) + ', [string]$p.ExitCode); exit $p.ExitCode'
    $psExe  = Join-Path $script:Sys32 'WindowsPowerShell\v1.0\powershell.exe'           # no spaces in this path: safe on any command line
    $enc    = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($inner))   # unique per call (it contains the GUID file names)
    $psArgs = '-NoProfile -NonInteractive -ExecutionPolicy Bypass -WindowStyle Hidden -EncodedCommand ' + $enc
    $usedHeadless = $script:Headless
    try {
        if ($usedHeadless) { $act = New-ScheduledTaskAction -Execute (Join-Path $script:Sys32 'conhost.exe') -Argument ('--headless ' + $psExe + ' ' + $psArgs) -WorkingDirectory $dir }
        else               { $act = New-ScheduledTaskAction -Execute $psExe -Argument $psArgs -WorkingDirectory $dir }
        $pri = New-ScheduledTaskPrincipal -UserId ([Security.Principal.WindowsIdentity]::GetCurrent().Name) -LogonType Interactive -RunLevel Limited
        $set = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -MultipleInstances IgnoreNew `
                  -ExecutionTimeLimit (New-TimeSpan -Seconds ($TimeoutSec + 300))
        Register-ScheduledTask -TaskName $name -TaskPath $tp -Action $act -Principal $pri -Settings $set -Force -ErrorAction Stop | Out-Null
        $t0 = Get-Date
        Start-ScheduledTask -TaskName $name -TaskPath $tp -ErrorAction Stop
    } catch {
        Write-Ev 2 $Prov ('לא ניתן להריץ בהרשאות משתמש רגיל: ' + $_.Exception.Message) 'err'
        Unregister-ScheduledTask -TaskName $name -TaskPath $tp -Confirm:$false -ErrorAction SilentlyContinue
        return $fail
    }
    $readers = @(@{ Path = $out; Pos = 0L; Buf = ''; Sb = (New-Object System.Text.StringBuilder); Quiet = [bool]$QuietStdout },
                 @{ Path = $err; Pos = 0L; Buf = ''; Sb = $null; Quiet = $false })
    $all = New-Object System.Text.StringBuilder
    $sw = [Diagnostics.Stopwatch]::StartNew(); $started = $false; $timedOut = $false; $cancelled = $false; $noted = $false; $code = -1; $haveCode = $false
    try {
        while ($true) {
            Start-Sleep -Milliseconds 500
            foreach ($r in $readers) { Read-ProcStream $r $Prov $all }
            Flush-State
            $st   = [string](Get-ScheduledTask -TaskName $name -TaskPath $tp -ErrorAction SilentlyContinue).State
            $info = Get-ScheduledTaskInfo -TaskName $name -TaskPath $tp -ErrorAction SilentlyContinue
            $res  = [int64]0; if ($info) { $res = [int64]$info.LastTaskResult }
            if ($st -eq 'Running' -or ($info -and $info.LastRunTime -ge $t0.AddSeconds(-2))) { $started = $true }
            # 0x41301 = running, 0x41303 = has not run yet
            if ($started -and $st -ne 'Running' -and $res -ne 0x41301 -and $res -ne 0x41303) {
                if ($res -gt 0x7FFFFFFF) { $code = [int]($res - 0x100000000L) } else { $code = [int]$res }   # uint32 → signed, like Process.ExitCode
                break
            }
            if (-not $started -and $sw.Elapsed.TotalSeconds -gt 90) { Write-Ev 2 $Prov 'המשימה בהרשאות משתמש לא התחילה (משתמש לא מחובר?)' 'err'; return $fail }
            if ($sw.Elapsed.TotalSeconds -gt $TimeoutSec) { $timedOut = $true; Stop-UserTree $pidFile $t0 $enc; Stop-ScheduledTask -TaskName $name -TaskPath $tp -ErrorAction SilentlyContinue; break }
            if (Test-Path $script:CancelFlag) {
                if ($Interruptible) { $cancelled = $true; Stop-UserTree $pidFile $t0 $enc; Stop-ScheduledTask -TaskName $name -TaskPath $tp -ErrorAction SilentlyContinue; break }
                if (-not $noted) { $noted = $true; Write-Ev 2 $Prov 'בקשת עצירה התקבלה — ממתין לסיום הפעולה הנוכחית' 'warn' }
            }
        }
        Start-Sleep -Milliseconds 300
        foreach ($r in $readers) { Read-ProcStream $r $Prov $all -Final }
        if (Test-Path -LiteralPath $codeFile) {
            $n = 0; if ([int]::TryParse(([IO.File]::ReadAllText($codeFile)).Trim(), [ref]$n)) { $code = $n; $haveCode = $true }
        }
    } finally {
        Unregister-ScheduledTask -TaskName $name -TaskPath $tp -Confirm:$false -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $out, $err, $codeFile, $pidFile -Force -ErrorAction SilentlyContinue
    }
    if (-not $haveCode -and -not $timedOut -and -not $cancelled -and $usedHeadless -and -not $NoRetry -and $sw.Elapsed.TotalSeconds -lt 60) {
        # The wrapper never ran the command (no exit-code file): this Windows has no headless conhost.
        $script:Headless = $false
        Write-Ev 2 $Prov 'conhost --headless לא זמין — ממשיך עם חלון PowerShell מוסתר (עשוי להבהב לרגע)' 'warn'
        return (Invoke-AsUser $File $Arguments $Prov $TimeoutSec -QuietStdout:$QuietStdout -Interruptible:$Interruptible -NoRetry)
    }
    if ($timedOut) { $code = -2; Write-Ev 2 $Prov ("{0} חרג מ-{1} דקות ונעצר" -f (Split-Path $File -Leaf), [int]($TimeoutSec / 60)) 'warn' }
    Write-Ev 3 $Prov ("exit code: {0}  ({1:N1}s)" -f $code, $sw.Elapsed.TotalSeconds) $(if ($code -ne 0) { 'warn' } else { 'info' })
    if ($cancelled) { Test-Cancel }
    return [pscustomobject]@{ ExitCode = $code; Output = $all.ToString(); StdOut = $readers[0].Sb.ToString() }
}

function Invoke-Tool([bool]$AsUser, [string]$File, [string]$Arguments, [string]$Prov, [int]$TimeoutSec = 1800, [switch]$QuietStdout, [switch]$Interruptible) {
    if ($AsUser) { return (Invoke-AsUser $File $Arguments $Prov $TimeoutSec -QuietStdout:$QuietStdout -Interruptible:$Interruptible) }
    if ($File -match '\.(cmd|bat)$') { return (Invoke-External $script:Cmd (Get-CmdArgs $File $Arguments) $Prov $TimeoutSec -QuietStdout:$QuietStdout -Interruptible:$Interruptible) }
    return (Invoke-External $File $Arguments $Prov $TimeoutSec -QuietStdout:$QuietStdout -Interruptible:$Interruptible)
}

function Invoke-Http([string]$Uri, [string]$Prov, [string]$Method = 'GET', [string]$Body, [int]$TimeoutSec = 45) {
    Flush-State
    Write-Ev 3 $Prov ("$Method $Uri") 'cmd'
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $p = @{ Uri = $Uri; Method = $Method; UseBasicParsing = $true; TimeoutSec = $TimeoutSec; UserAgent = $script:UA; ErrorAction = 'Stop' }
    if ($Body) { $p.Body = $Body; $p.ContentType = 'application/x-www-form-urlencoded' }
    $r = Invoke-WebRequest @p
    Write-Ev 3 $Prov ("HTTP {0} · {1:N0} bytes · {2} ms" -f [int]$r.StatusCode, $r.RawContentLength, $sw.ElapsedMilliseconds)
    return $r
}

# Streams to disk with live progress on the item (percent + MB) and honours Stop (a partial download is harmless).
function Save-Download([string]$Uri, [string]$Dest, [string]$Prov, $It = $null, [string]$Label = '') {
    Write-Ev 3 $Prov ("GET $Uri → $Dest") 'cmd'
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $req = [Net.HttpWebRequest]::Create($Uri)
    $req.UserAgent = $script:UA; $req.Timeout = 60000; $req.ReadWriteTimeout = 300000; $req.AllowAutoRedirect = $true
    $resp = $req.GetResponse()
    try {
        $total = [long]$resp.ContentLength
        $in = $resp.GetResponseStream()
        $out = [IO.File]::Create($Dest)
        try {
            $buf = New-Object byte[] 1048576; $done = 0L; $tick = [Diagnostics.Stopwatch]::StartNew()
            while (($n = $in.Read($buf, 0, $buf.Length)) -gt 0) {
                $out.Write($buf, 0, $n); $done += $n
                if ($tick.ElapsedMilliseconds -ge 1000) {
                    $tick.Restart()
                    if ($script:CancelFlag -and (Test-Path $script:CancelFlag)) { Test-Cancel }
                    $txt = "{0:N0} MB" -f ($done / 1MB)
                    if ($total -gt 0) { $txt = "{0}% · {1:N0}/{2:N0} MB" -f [int](100 * $done / $total), ($done / 1MB), ($total / 1MB) }
                    if ($It) { $It.reason = "מוריד $Label · $txt" }
                    $script:State.current = "מוריד $Label · $txt"
                    Save-State
                }
            }
        } finally { $out.Dispose(); $in.Dispose() }
    } finally { $resp.Close() }
    $len = (Get-Item $Dest).Length
    if ($total -gt 0 -and $len -ne $total) { throw ("ההורדה נקטעה ({0:N0} מתוך {1:N0} בתים)" -f $len, $total) }
    Write-Ev 2 $Prov ("הורד {0:N1} MB ב-{1:N0} שניות" -f ($len / 1MB), $sw.Elapsed.TotalSeconds)
}

function Compare-Ver([string]$A, [string]$B) {
    # returns 1 if A>B, -1 if A<B, 0 if equal, $null if not comparable
    $va = $null; $vb = $null
    if (-not [version]::TryParse(($A -replace '[^\d\.]', '').Trim('.'), [ref]$va)) { return $null }
    if (-not [version]::TryParse(($B -replace '[^\d\.]', '').Trim('.'), [ref]$vb)) { return $null }
    return $va.CompareTo($vb)
}

function Test-IsAdmin {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    return (New-Object Security.Principal.WindowsPrincipal($id)).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# ════════════════════════════════════════════════════════════════════════════
#  Safety: restore point, firmware gate, BitLocker
# ════════════════════════════════════════════════════════════════════════════
$script:RpDone = $false; $script:BlDone = $false
function Ensure-RestorePoint {
    if ($script:RpDone -or $Mode -ne 'Update' -or -not (Get-Setting 'restorePoint')) { return }
    $script:RpDone = $true
    Set-Current 'יוצר נקודת שחזור...'
    $warn = $null
    try {
        Checkpoint-Computer -Description ("UpdateHub " + $script:RunId) -RestorePointType MODIFY_SETTINGS -ErrorAction Stop -WarningVariable warn -WarningAction SilentlyContinue
        # Windows skips a new point if one was made in the last 24 h — that is reported, not hidden.
        if ($warn) { Write-Ev 1 'system' ('נקודת שחזור חדשה לא נוצרה: ' + ($warn -join ' ')) 'warn' }
        else { Write-Ev 2 'system' 'נוצרה נקודת שחזור לפני ההתקנה הראשונה' 'ok' }
    } catch {
        Write-Ev 1 'system' ('נקודת שחזור לא נוצרה (הגנת מערכת כבויה?): ' + $_.Exception.Message) 'warn'
    }
}

function Test-FirmwareGate {
    if (-not (Get-Setting 'firmware')) { return 'עדכוני קושחה מושבתים בהגדרות' }
    if (-not (Get-Setting 'firmwareRequireAC')) { return $null }
    try {
        # Desktop = Windows reports no battery at all. BatteryChargeStatus 255 ("Unknown") also has bit 128 set,
        # so that flag alone would wrongly treat a laptop as a desktop.
        $batteries = @(Get-CimInstance Win32_Battery -ErrorAction Stop)
        if ($batteries.Count -eq 0) { return $null }
        Add-Type -AssemblyName System.Windows.Forms
        $ps = [System.Windows.Forms.SystemInformation]::PowerStatus
        if ($ps.PowerLineStatus -ne [System.Windows.Forms.PowerLineStatus]::Online) { return 'קושחה מותקנת רק כשהמחשב מחובר לחשמל' }
        # Percent from WMI (PowerStatus.BatteryLifePercent clamps "unknown" to 100%). Lowest battery counts.
        $pct = ($batteries | Where-Object { $null -ne $_.EstimatedChargeRemaining } | Measure-Object -Property EstimatedChargeRemaining -Minimum).Minimum
        if ($null -eq $pct -or $pct -gt 100) { return 'מצב הסוללה לא ידוע — הקושחה לא הותקנה' }
        $min = [int](Get-Setting 'firmwareMinBattery')
        if ($pct -lt $min) { return "קושחה מותקנת רק מעל $min% סוללה (כרגע $pct%)" }
    } catch { return ('לא ניתן לבדוק מצב חשמל, הקושחה לא הותקנה: ' + $_.Exception.Message) }   # fail closed
    return $null
}

function Ensure-BitLockerSuspended {
    if ($script:BlDone) { return }
    $script:BlDone = $true
    try {
        $v = Get-BitLockerVolume -MountPoint $script:SysDrive -ErrorAction Stop
        if ([string]$v.ProtectionStatus -eq 'On') {
            Suspend-BitLocker -MountPoint $script:SysDrive -RebootCount 1 -ErrorAction Stop | Out-Null
            $script:BlOurs = $true
            Write-Ev 2 'drivers' 'BitLocker הושהה לאתחול אחד — נדרש כדי שעדכון הקושחה לא יבקש מפתח שחזור' 'warn'
        } else { Write-Ev 3 'drivers' ('BitLocker: ' + $v.ProtectionStatus) }
    } catch { Write-Ev 3 'drivers' ('BitLocker לא זמין: ' + $_.Exception.Message) }
}
# Nothing was staged after all → do not leave the disk unprotected until the next restart.
$script:BlOurs = $false
function Resume-BitLockerIfOurs {
    if (-not $script:BlOurs) { return }
    if (@($script:Prov['drivers'].items | Where-Object { $_.status -eq 'reboot' -and $_.source -match 'קושחה' }).Count -gt 0) { return }
    try { Resume-BitLocker -MountPoint $script:SysDrive -ErrorAction Stop | Out-Null; $script:BlOurs = $false; $script:BlDone = $false
          Write-Ev 2 'drivers' 'הגנת BitLocker הוחזרה (לא הוכנה קושחה)' 'ok' }
    catch { Write-Ev 2 'drivers' ('לא ניתן להחזיר את BitLocker: ' + $_.Exception.Message) 'warn' }
}

# ════════════════════════════════════════════════════════════════════════════
#  Inventory
# ════════════════════════════════════════════════════════════════════════════
function Get-MachineInfo {
    $cv   = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
    $os   = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
    $cs   = Get-CimInstance Win32_ComputerSystem  -ErrorAction SilentlyContinue
    $bios = Get-CimInstance Win32_BIOS            -ErrorAction SilentlyContinue
    $bb   = Get-CimInstance Win32_BaseBoard       -ErrorAction SilentlyContinue
    $script:OsBuild = [int]$cv.CurrentBuild
    $script:IsArm   = ($env:PROCESSOR_ARCHITECTURE -eq 'ARM64')
    try { $script:IsArm = ([string][System.Runtime.InteropServices.RuntimeInformation]::OSArchitecture -eq 'Arm64') } catch { }   # API, not env
    $dv = $cv.DisplayVersion; if (-not $dv) { $dv = $cv.ReleaseId }
    return [ordered]@{
        computer     = $env:COMPUTERNAME
        os           = [string]$os.Caption
        osVersion    = [string]$dv
        osBuild      = ("{0}.{1}" -f $cv.CurrentBuild, $cv.UBR)
        arch         = $env:PROCESSOR_ARCHITECTURE
        manufacturer = [string]$cs.Manufacturer
        model        = [string]$cs.Model
        board        = ("{0} {1}" -f $bb.Manufacturer, $bb.Product).Trim()
        bios         = ("{0} ({1})" -f $bios.SMBIOSBIOSVersion, $bios.Manufacturer)
        user         = ("{0}\{1}" -f $env:USERDOMAIN, $env:USERNAME)
    }
}

function Get-DeviceInventory {
    $devs = @(Get-PnpDevice -PresentOnly -ErrorAction SilentlyContinue)
    $keys = @('DEVPKEY_Device_HardwareIds', 'DEVPKEY_Device_DriverVersion', 'DEVPKEY_Device_DriverProvider',
              'DEVPKEY_Device_MatchingDeviceId', 'DEVPKEY_Device_ProblemCode', 'DEVPKEY_Device_DriverInfPath', 'DEVPKEY_Device_CompatibleIds')
    $byId = @{}
    foreach ($pr in @($devs | Get-PnpDeviceProperty -KeyName $keys -ErrorAction SilentlyContinue)) {
        if (-not $byId.ContainsKey($pr.InstanceId)) { $byId[$pr.InstanceId] = @{} }
        $byId[$pr.InstanceId][$pr.KeyName] = $pr.Data
    }
    $list = New-Object System.Collections.ArrayList
    foreach ($d in $devs) {
        $pp = $byId[$d.InstanceId]; if (-not $pp) { $pp = @{} }
        $name = $d.FriendlyName; if (-not $name) { $name = $d.Name }; if (-not $name) { $name = $d.InstanceId }
        $prob = 0; if ($pp['DEVPKEY_Device_ProblemCode']) { $prob = [int]$pp['DEVPKEY_Device_ProblemCode'] }
        [void]$list.Add([pscustomobject]@{
            InstanceId    = [string]$d.InstanceId
            Name          = [string]$name
            Class         = [string]$d.Class
            Status        = [string]$d.Status
            HardwareIds   = @($pp['DEVPKEY_Device_HardwareIds'] | Where-Object { $_ })
            CompatibleIds = @($pp['DEVPKEY_Device_CompatibleIds'] | Where-Object { $_ })
            DriverVersion = [string]$pp['DEVPKEY_Device_DriverVersion']
            Provider      = [string]$pp['DEVPKEY_Device_DriverProvider']
            MatchingId    = [string]$pp['DEVPKEY_Device_MatchingDeviceId']
            Problem       = $prob
            Inf           = [string]$pp['DEVPKEY_Device_DriverInfPath']
        })
    }
    return ,$list
}

# ════════════════════════════════════════════════════════════════════════════
#  Microsoft Update Catalog (drivers + UEFI firmware by Hardware ID)
# ════════════════════════════════════════════════════════════════════════════
$script:CatCache = @{}; $script:CatFailures = 0

# $script:CatLastFailed tells "the lookup failed" apart from "the catalog has nothing for this ID".
$script:CatLastFailed = $false
function Search-Catalog([string]$Query) {
    $script:CatLastFailed = $false
    if ($script:CatCache.ContainsKey($Query)) { return ,$script:CatCache[$Query] }
    $rows = New-Object System.Collections.ArrayList
    $html = $null
    for ($try = 1; $try -le 2 -and -not $html; $try++) {
        try {
            $r = Invoke-Http ($script:CatBase + '/Search.aspx?q=' + [uri]::EscapeDataString($Query)) 'drivers'
            $html = $r.Content
            $script:CatFailures = 0
        } catch {
            Write-Ev 3 'drivers' ("ניסיון $try נכשל: " + $_.Exception.Message) 'warn'
            Start-Sleep -Seconds (2 * $try)
        }
    }
    if (-not $html) { $script:CatFailures++; $script:CatLastFailed = $true; return ,$rows }
    if ($html -notmatch 'updateMatches' -and $html -notmatch 'noResultText') {
        # Neither a result table nor the "no results" banner: proxy page, outage, or a catalog redesign.
        Write-Ev 2 'drivers' 'הקטלוג החזיר דף לא צפוי (לא טבלת תוצאות ולא "אין תוצאות")' 'warn'
        $script:CatFailures++; $script:CatLastFailed = $true; return ,$rows
    }
    $rowRx  = [regex]'(?s)<tr id="([0-9a-fA-F-]{36})_R\d+">(.*?)</tr>'
    $cellRx = [regex]'(?s)_C(\d+)_R\d+"[^>]*>(.*?)</td>'
    foreach ($m in $rowRx.Matches($html)) {
        $cells = @{}
        foreach ($c in $cellRx.Matches($m.Groups[2].Value)) {
            $t = $c.Groups[2].Value -replace '<[^>]+>', ' '
            $t = [System.Net.WebUtility]::HtmlDecode($t) -replace '\s+', ' '
            $cells[[int]$c.Groups[1].Value] = $t.Trim()
        }
        [void]$rows.Add([pscustomobject]@{
            Id = $m.Groups[1].Value; Title = $cells[1]; Products = $cells[2]; Class = $cells[3]; Date = $cells[4]; Version = $cells[5]
        })
    }
    $script:CatCache[$Query] = $rows
    Start-Sleep -Milliseconds 250
    return ,$rows
}

# Minimum OS build a catalog row targets ("Windows 11 Client, version 24H2 and later" → 26100). $null = not a client OS row.
$script:Win10Map = @{ '1507' = 10240; '1511' = 10586; '1607' = 14393; '1703' = 15063; '1709' = 16299; '1803' = 17134; '1809' = 17763
                      '1903' = 18362; '1909' = 18363; '2004' = 19041; '20H2' = 19042; '21H1' = 19043; '21H2' = 19044; '22H2' = 19045 }
$script:Win11Map = @{ '21H2' = 22000; '22H2' = 22621; '23H2' = 22631; '24H2' = 26100; '25H2' = 26200; '26H1' = 28000; '26H2' = 28000 }
$script:CodeMap  = @{ 'Threshold' = 10240; 'Redstone' = 14393; 'RS5' = 17763; '19H1' = 18362; 'Vibranium' = 19041
                      'Cobalt' = 22000; 'Nickel' = 22621; 'Germanium' = 26100 }
function Get-RowMinBuild([string]$Products) {
    $builds = @()
    $rx = [regex]'(Windows 11 Client|Windows 10|Windows - Client)(?: S)?,\s*(?:version\s+)?([A-Za-z0-9]+)\s+and later'
    foreach ($m in $rx.Matches($Products)) {
        $fam = $m.Groups[1].Value; $tok = $m.Groups[2].Value; $b = $null
        if ($fam -eq 'Windows 11 Client' -and $script:Win11Map.ContainsKey($tok)) { $b = $script:Win11Map[$tok] }
        elseif ($script:Win10Map.ContainsKey($tok)) { $b = $script:Win10Map[$tok] }
        elseif ($script:CodeMap.ContainsKey($tok)) { $b = $script:CodeMap[$tok] }
        elseif ($script:Win11Map.ContainsKey($tok)) { $b = $script:Win11Map[$tok] }
        if ($b) { $builds += $b }
    }
    if ($builds.Count -eq 0) { return $null }
    return ($builds | Measure-Object -Minimum).Minimum
}

function Select-BestRow($Rows) {
    $best = $null; $bestVer = $null; $bestBuild = 0
    foreach ($r in $Rows) {
        $minB = Get-RowMinBuild $r.Products
        if (-not $minB) { continue }                                   # server-only or unknown target
        if ($minB -gt $script:OsBuild) { continue }                    # built for a newer Windows
        $isArmRow = ("$($r.Title) $($r.Products)" -match 'ARM64')
        if ($isArmRow -ne $script:IsArm) { continue }
        $v = $null
        if (-not [version]::TryParse(($r.Version -replace '[^\d\.]', ''), [ref]$v)) { continue }
        $cmp = 1; if ($bestVer) { $cmp = $v.CompareTo($bestVer) }
        if ($cmp -gt 0 -or ($cmp -eq 0 -and $minB -gt $bestBuild)) { $best = $r; $bestVer = $v; $bestBuild = $minB }
    }
    if (-not $best) { return $null }
    return [pscustomobject]@{ Row = $best; Version = $bestVer; VersionString = $best.Version; TargetBuild = $bestBuild }
}

function Select-CatalogCandidates($Devs) {
    $seen = @{}; $out = New-Object System.Collections.ArrayList
    $exRx = [string](Get-Setting 'excludeDevices')
    foreach ($d in $Devs) {
        if ($d.HardwareIds.Count -eq 0) { continue }
        if ($d.InstanceId -notmatch '^(PCI|USB|HDAUDIO|INTELAUDIO|ACPI|UEFI|SWC|HID)\\') { continue }
        $isFw      = $d.InstanceId -like 'UEFI\RES_*'
        $noDriver  = ($d.Problem -eq 28)
        $vendorDrv = ($d.Provider -and $d.Provider -notmatch '^Microsoft')
        if (-not ($isFw -or $noDriver -or $vendorDrv)) { continue }    # inbox Microsoft drivers: Windows Update owns them
        if ($d.Class -eq 'Display' -and $d.HardwareIds[0] -match 'VEN_10DE') { continue }   # NVIDIA → own provider
        if ($exRx -and $d.Name -match $exRx) { continue }
        $key = $d.HardwareIds[0].ToUpperInvariant()
        if ($seen.ContainsKey($key)) { continue }
        $seen[$key] = 1
        [void]$out.Add($d)
    }
    return ,$out
}

# IDs we are allowed to search/install against: never less specific than what the current driver matched.
function Get-SearchIds($D) {
    $ids = @($D.HardwareIds)
    $stop = $ids.Count - 1
    if ($D.MatchingId) {
        for ($i = 0; $i -lt $ids.Count; $i++) { if ($ids[$i] -ieq $D.MatchingId) { $stop = $i; break } }
    }
    $sel = @($ids[0..$stop] | Where-Object { $_ -notmatch '\\(CC_|CLASS_)' -and $_ -notmatch '^\*' })
    return @($sel | Select-Object -First 4)
}

function Find-CatalogDriver($D) {
    foreach ($id in (Get-SearchIds $D)) {
        Test-Cancel
        $rows = Search-Catalog $id
        # A failed lookup is NOT "no results": stop here instead of falling back to a less specific ID.
        if ($script:CatLastFailed) { return [pscustomobject]@{ Failed = $true } }
        Write-Ev 3 'drivers' ("{0}  →  {1} תוצאות" -f $id, $rows.Count)
        if ($rows.Count -eq 0) { continue }
        $best = Select-BestRow $rows
        if ($best) { Write-Ev 3 'drivers' ("נבחר: {0} [{1}] יעד build {2}" -f $best.Row.Title, $best.Row.Id, $best.TargetBuild) }
        else       { Write-Ev 3 'drivers' 'יש תוצאות, אך אף אחת לא מתאימה לגרסת Windows/ארכיטקטורה הזו' }
        # Stop at the most specific ID that has results — never fall back to a more generic driver.
        if ($best) { $best | Add-Member -NotePropertyName QueryId -NotePropertyValue $id }
        return $best
    }
    return $null
}

function Get-CatalogDownload([string]$UpdateId) {
    $payload = '[{"size":0,"languages":"","uidInfo":"' + $UpdateId + '","updateID":"' + $UpdateId + '"}]'
    $r = Invoke-Http ($script:CatBase + '/DownloadDialog.aspx') 'drivers' 'POST' ('updateIDs=' + [uri]::EscapeDataString($payload))
    $map = @{}
    foreach ($m in ([regex]"downloadInformation\[0\]\.files\[(\d+)\]\.(url|digest|fileName)\s*=\s*'([^']*)'").Matches($r.Content)) {
        $k = $m.Groups[1].Value
        if (-not $map.ContainsKey($k)) { $map[$k] = @{} }
        $map[$k][$m.Groups[2].Value] = $m.Groups[3].Value
    }
    $files = @()
    foreach ($k in ($map.Keys | Sort-Object)) { $files += [pscustomobject]@{ Url = $map[$k]['url']; Digest = $map[$k]['digest']; FileName = $map[$k]['fileName'] } }
    if ($files.Count -eq 0) { throw 'לא נמצא קישור הורדה בקטלוג' }
    return $files
}

# Running firmware version of a UEFI\RES_{GUID} device: the ESRT entry Windows exposes in the registry,
# falling back to the &REV_xxxxxxxx hardware ID (which is the same value, in hex). DWORDs are read unsigned.
function Get-RunningFirmware($D) {
    $gm = [regex]::Match([string]$D.InstanceId, '^UEFI\\RES_(\{[0-9A-Fa-f-]{36}\})')
    if (-not $gm.Success) { return $null }
    $k = Get-ItemProperty -Path ('HKLM:\HARDWARE\UEFI\ESRT\' + $gm.Groups[1].Value) -ErrorAction SilentlyContinue
    if ($k -and $null -ne $k.Version) {
        $low = 0L; if ($null -ne $k.LowestSupportedVersion) { $low = ([int64]$k.LowestSupportedVersion) -band 0xFFFFFFFFL }
        return @{ Version = (([int64]$k.Version) -band 0xFFFFFFFFL); Lowest = $low; Source = 'ESRT' }
    }
    foreach ($h in $D.HardwareIds) {
        if ($h -match '&REV_([0-9A-Fa-f]{1,8})$') { return @{ Version = [Convert]::ToInt64($Matches[1], 16); Lowest = 0L; Source = 'REV' } }
    }
    return $null
}
# FirmwareVersion that a capsule INF will flash:  HKR,,FirmwareVersion,%REG_DWORD%,0x0001002A
function Get-InfFirmwareVersion([string]$InfText) {
    $m = [regex]::Match($InfText, '(?im)^\s*HKR\s*,\s*,\s*FirmwareVersion\s*,\s*[^,]*,\s*(0x[0-9a-f]+|\d+)')
    if (-not $m.Success) { return $null }
    $v = $m.Groups[1].Value
    if ($v -like '0x*') { return [Convert]::ToInt64($v.Substring(2), 16) }
    return [int64]$v
}

# Decisions that do not change until the catalog row or the installed driver changes (Windows kept a better-ranked
# driver, firmware already newer): remembered so the same package is not downloaded again every week.
$script:Decisions = $null
function Get-DecisionKey($D, $M) { return ("{0}|{1}|{2}" -f $D.InstanceId, $M.Row.Id, $D.DriverVersion).ToUpperInvariant() }
function Get-Decision($D, $M) {
    if ($null -eq $script:Decisions) {
        $script:Decisions = @{}
        $p = Join-Path $DataDir 'decisions.json'
        if (Test-Path $p) {
            try { $o = Get-Content -Raw -Encoding UTF8 $p | ConvertFrom-Json -ErrorAction Stop
                  foreach ($pr in $o.PSObject.Properties) { $script:Decisions[$pr.Name] = $pr.Value } } catch { }
        }
    }
    return $script:Decisions[(Get-DecisionKey $D $M)]
}
function Save-Decision($D, $M, [string]$Status, [string]$Reason) {
    $null = Get-Decision $D $M
    $script:Decisions[(Get-DecisionKey $D $M)] = [pscustomobject]@{ status = $Status; reason = $Reason; at = (Get-Date).ToString('yyyy-MM-dd') }
    try { [IO.File]::WriteAllText((Join-Path $DataDir 'decisions.json'), ($script:Decisions | ConvertTo-Json -Depth 3 -Compress), (New-Object System.Text.UTF8Encoding($false))) } catch { }
}

function Install-CatalogDriver($D, $M, $It) {
    $isFw = $D.InstanceId -like 'UEFI\RES_*'
    if ($isFw) {
        $gate = Test-FirmwareGate
        if ($gate) { Set-Entry $It 'manual' $gate; Write-Ev 2 'drivers' ($D.Name + ': ' + $gate) 'warn'; return }
    }
    Ensure-RestorePoint
    $dir = Join-Path $script:WorkDir ([guid]::NewGuid().ToString('N'))   # unpredictable path, admin-only parent
    try {
        Set-Entry $It 'downloading' ('מוריד: ' + $M.Row.Title)
        Set-Current ('מוריד: ' + $M.Row.Title)
        New-Item -ItemType Directory -Force -Path $dir | Out-Null
        $files = Get-CatalogDownload $M.Row.Id
        foreach ($f in $files) {
            $dest = Join-Path $dir ([IO.Path]::GetFileName($f.FileName))
            Save-Download $f.Url $dest 'drivers' $It $D.Name
            if ($f.Digest) {
                $sha = [Security.Cryptography.SHA1]::Create(); $fs = [IO.File]::OpenRead($dest)
                try { $h = [Convert]::ToBase64String($sha.ComputeHash($fs)) } finally { $fs.Dispose() }
                if ($h -ne $f.Digest) { Set-Entry $It 'failed' 'קובץ שהורד לא תואם ל-hash של הקטלוג'; return }
                Write-Ev 3 'drivers' ("SHA1 תואם לקטלוג: $h") 'ok'
            }
        }
        $x = Join-Path $dir 'x'; New-Item -ItemType Directory -Force -Path $x | Out-Null
        foreach ($cab in @(Get-ChildItem $dir -Filter *.cab)) {
            $null = Invoke-External (Join-Path $script:Sys32 'expand.exe') ('"' + $cab.FullName + '" -F:* "' + $x + '"') 'drivers' 600 -Interruptible
        }
        $infs = @(Get-ChildItem $x -Recurse -Filter *.inf -ErrorAction SilentlyContinue)
        if ($infs.Count -eq 0) { Set-Entry $It 'failed' 'אין קובץ INF בחבילה'; return }

        # The package must explicitly list one of this device's IDs (at least as specific as today's match).
        $accept = @(Get-SearchIds $D); $hit = $null; $hitTxt = ''
        foreach ($inf in $infs) {
            $txt = [IO.File]::ReadAllText($inf.FullName)
            foreach ($id in $accept) { if ($txt.IndexOf($id, [StringComparison]::OrdinalIgnoreCase) -ge 0) { $hit = "$($inf.Name) ← $id"; $hitTxt = $txt; break } }
            if ($hit) { break }
        }
        if (-not $hit) { Set-Entry $It 'skipped' 'החבילה לא מיועדת במדויק לרכיב הזה — לא הותקנה'; Write-Ev 2 'drivers' ($D.Name + ': ה-INF לא מכיל את מזהה הרכיב, דילוג') 'warn'; return }
        Write-Ev 3 'drivers' ("התאמת INF: $hit") 'ok'

        if ($isFw) {
            # Compare with the firmware that is actually RUNNING (ESRT), not with the last INF Windows bound —
            # a BIOS flashed from USB/BIOS menu leaves an old INF behind, and the capsule must never be a rollback.
            $run = Get-RunningFirmware $D
            $pkgFw = Get-InfFirmwareVersion $hitTxt
            if (-not $run -or $null -eq $pkgFw) {
                Set-Entry $It 'manual' 'לא ניתן לאמת את גרסת הקושחה הפעילה מול החבילה — לא הותקנה'
                Write-Ev 2 'drivers' ("{0}: ESRT={1} · FirmwareVersion בחבילה={2} — דילוג" -f $D.Name, [bool]$run, $pkgFw) 'warn'; return
            }
            Write-Ev 3 'drivers' ("קושחה פעילה 0x{0:X8} ({1}) · בחבילה 0x{2:X8} · מינימום 0x{3:X8}" -f $run.Version, $run.Source, $pkgFw, $run.Lowest)
            if ($pkgFw -le $run.Version) { Set-Entry $It 'current' 'הקושחה הפעילה כבר בגרסה זו או חדשה יותר'; Save-Decision $D $M 'current' 'הקושחה הפעילה כבר בגרסה זו או חדשה יותר'; return }
            if ($run.Lowest -gt 0 -and $pkgFw -lt $run.Lowest) { Set-Entry $It 'skipped' 'החבילה מתחת לגרסה המינימלית שהקושחה מאפשרת'; return }
        }

        foreach ($cat in @(Get-ChildItem $x -Recurse -Filter *.cat -ErrorAction SilentlyContinue)) {
            $sig = Get-AuthenticodeSignature -FilePath $cat.FullName
            Write-Ev 3 'drivers' ("חתימה {0}: {1} · {2}" -f $cat.Name, $sig.Status, $sig.SignerCertificate.Subject)
            if ([string]$sig.Status -ne 'Valid') { Set-Entry $It 'failed' ("חתימה לא תקינה: " + $sig.Status); return }
            # Everything Microsoft distributes through the catalog is signed by Microsoft (WHQL / attestation).
            if ((Get-CertOrg $sig.SignerCertificate) -ne 'Microsoft Corporation') { Set-Entry $It 'failed' ('הקטלוג לא חתום על ידי Microsoft: ' + $sig.SignerCertificate.Subject); return }
        }

        if ($isFw) { Ensure-BitLockerSuspended }
        Set-Entry $It 'installing' ('מתקין: ' + $M.Row.Title)
        Set-Current ('מתקין: ' + $D.Name)
        $res = Invoke-External (Join-Path $script:Sys32 'pnputil.exe') ('/add-driver "' + $x + '\*.inf" /subdirs /install') 'drivers' 1200
        Start-Sleep -Seconds 2
        $newVer = [string](Get-PnpDeviceProperty -InstanceId $D.InstanceId -KeyName DEVPKEY_Device_DriverVersion -ErrorAction SilentlyContinue).Data
        Write-Ev 3 'drivers' ("אימות: גרסה לפני {0} · אחרי {1} · יעד {2}" -f $D.DriverVersion, $newVer, $M.VersionString)
        $applied = ((Compare-Ver $newVer $M.VersionString) -eq 0)
        if ($isFw -and ($applied -or $res.ExitCode -eq 3010)) {
            # The firmware "driver" is bound now; the flash itself happens during the next restart.
            Set-Entry $It 'reboot' 'הקושחה הוכנה — תיצרב באתחול הבא (אל תנתק חשמל)'
        } elseif ($isFw) {
            Set-Entry $It 'skipped' ("pnputil (קוד {0}) לא קשר את הקושחה לרכיב — לא תיצרב" -f $res.ExitCode)
            Resume-BitLockerIfOurs
        } elseif ($applied -and $res.ExitCode -eq 3010) {
            Set-Entry $It 'reboot' 'הותקן — נדרש אתחול'
        } elseif ($applied) {
            Set-Entry $It 'updated' ('עודכן ל-' + $newVer)
        } elseif ($res.ExitCode -in 0, 259, 3010) {
            $why = 'Windows השאיר את הדרייבר הנוכחי (דירוג גבוה יותר) — החבילה לא תורד שוב'
            Set-Entry $It 'skipped' $why
            Save-Decision $D $M 'skipped' $why
        } else {
            Set-Entry $It 'failed' ("pnputil החזיר קוד " + $res.ExitCode)
        }
    } catch {
        if ($isFw) { Resume-BitLockerIfOurs }
        if (Test-IsCancel $_) { throw }
        Set-Entry $It 'failed' $_.Exception.Message
        Write-Ev 2 'drivers' ($D.Name + ': ' + $_.Exception.Message) 'err'
    } finally {
        Remove-Item $dir -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# ════════════════════════════════════════════════════════════════════════════
#  Providers
# ════════════════════════════════════════════════════════════════════════════
function Run-System($P) {
    $m = Get-MachineInfo
    $script:State.machine = $m
    $null = Add-Entry $P 'Windows' 'מערכת' ("{0} {1}" -f $m.osVersion, $m.osBuild) '' 'info'
    $null = Add-Entry $P 'דגם' 'מערכת' ("{0} {1}" -f $m.manufacturer, $m.model) '' 'info'
    $null = Add-Entry $P 'BIOS' 'מערכת' $m.bios '' 'info'
    Set-Current 'סורק את כל רכיבי החומרה...'
    $script:Devices = Get-DeviceInventory
    $withDrv = @($script:Devices | Where-Object { $_.DriverVersion }).Count
    $null = Add-Entry $P 'רכיבי חומרה' 'מערכת' ("$($script:Devices.Count) רכיבים · $withDrv עם דרייבר") '' 'info'
    Write-Ev 2 'system' ("{0} · {1} {2} · build {3} · {4} רכיבי חומרה" -f $m.os, $m.manufacturer, $m.model, $m.osBuild, $script:Devices.Count)
    foreach ($d in $script:Devices) {
        if ($d.DriverVersion) { Write-Ev 3 'system' ("{0} · {1} · {2} · {3}" -f $d.Name, $d.DriverVersion, $d.Provider, ($d.HardwareIds | Select-Object -First 1)) }
    }
}

# Asynchronous WUA call with live percent, a hard time limit and Stop support (downloads only — an install is
# never aborted on Stop). Falls back to the synchronous call if this WUA build refuses null callbacks.
function Invoke-WuaJob([string]$Kind, $Obj, $It, [string]$Title, [int]$TimeoutMin) {
    try {
        # A null IUnknown for the callbacks (New-Object would drop the $null argument, so ::new is used).
        $cb = [System.Runtime.InteropServices.UnknownWrapper]::new($null)
        switch ($Kind) {
            'search'   { $job = $Obj.BeginSearch($Title, $cb, $null) }
            'download' { $job = $Obj.BeginDownload($cb, $cb, $null) }
            'install'  { $job = $Obj.BeginInstall($cb, $cb, $null) }
        }
    } catch {
        Write-Ev 3 'windowsupdate' ("מצב אסינכרוני לא זמין ($Kind), עובר לקריאה רגילה: " + $_.Exception.Message) 'warn'
        switch ($Kind) { 'search' { return $Obj.Search($Title) } 'download' { return $Obj.Download() } 'install' { return $Obj.Install() } }
    }
    $sw = [Diagnostics.Stopwatch]::StartNew(); $last = -1; $aborted = $false; $cancelled = $false; $noted = $false
    while (-not $job.IsCompleted) {
        Start-Sleep -Milliseconds 700
        if ($Kind -ne 'search') {
            $pct = -1; try { $pct = [int]$job.GetProgress().PercentComplete } catch { }
            if ($pct -ge 0 -and $pct -ne $last) {
                $last = $pct
                $verb = 'מוריד'; if ($Kind -eq 'install') { $verb = 'מתקין' }
                if ($It) { $It.reason = "$verb · $pct%" }
                $script:State.current = "$verb · $pct% · $Title"
                Save-State
                if ($pct % 10 -eq 0) { Write-Ev 3 'windowsupdate' ("$verb $pct%") }
            }
        }
        Flush-State
        if (-not $aborted -and $sw.Elapsed.TotalMinutes -gt $TimeoutMin) {
            Write-Ev 2 'windowsupdate' ("$Kind חרג מ-$TimeoutMin דקות — מבטל") 'warn'; $aborted = $true; try { $job.RequestAbort() } catch { }
        }
        if (-not $aborted -and (Test-Path $script:CancelFlag)) {
            if ($Kind -eq 'install') {
                if (-not $noted) { $noted = $true; Write-Ev 2 'windowsupdate' 'בקשת עצירה התקבלה — ממתין לסיום ההתקנה הנוכחית' 'warn' }
            } else { $cancelled = $true; $aborted = $true; try { $job.RequestAbort() } catch { } }
        }
        if ($aborted -and $sw.Elapsed.TotalMinutes -gt ($TimeoutMin + 10)) { throw "Windows Update לא הגיב לביטול ($Kind)" }
    }
    switch ($Kind) {
        'search'   { $r = $Obj.EndSearch($job) }
        'download' { $r = $Obj.EndDownload($job) }
        'install'  { $r = $Obj.EndInstall($job) }
    }
    if ($cancelled) { Test-Cancel }
    return $r
}

# Returns $null when a WU driver may be installed, otherwise @{Status; Reason}.
function Test-WuDriverSafe($U) {
    if (-not $script:Devices) { $script:Devices = Get-DeviceInventory }
    $hw = [string]$U.DriverHardwareID
    $offer = [string]$U.DriverVerVersion
    $targets = @($script:Devices | Where-Object { ($_.HardwareIds -contains $hw) -or ($_.CompatibleIds -contains $hw) })   # -contains is case-insensitive
    $exRx = [string](Get-Setting 'excludeDevices')
    if ($exRx -and (@($targets | Where-Object { $_.Name -match $exRx }).Count -gt 0 -or [string]$U.DriverModel -match $exRx)) {
        return @{ Status = 'skipped'; Reason = 'הוחרג בהגדרות (excludeDevices)' }
    }
    # Extension INFs carry their own version line (often 1.0.x) and add to the function driver rather than replace it:
    # comparing them with the function driver's version would block them forever. WU offers one only if it outranks
    # the installed extension.
    if ([string]$U.DriverClass -eq 'Extension') { return $null }
    foreach ($t in $targets) {
        if (-not $t.DriverVersion) { continue }
        $cmp = Compare-Ver $offer $t.DriverVersion
        if ($null -eq $cmp) { return @{ Status = 'skipped'; Reason = "לא ניתן להשוות $offer לגרסה המותקנת $($t.DriverVersion) — לא יותקן" } }
        if ($cmp -le 0) { return @{ Status = 'current'; Reason = "מותקנת כבר גרסה $($t.DriverVersion) (זהה או חדשה מ-$offer) — לא יותקן" } }
    }
    if ($targets.Count -eq 0) { Write-Ev 3 'windowsupdate' ("$hw לא נמצא ברשימת הרכיבים — ההתאמה נשארת של Windows Update") }
    return $null
}

function Run-WindowsUpdate($P) {
    Set-Current 'מחפש עדכונים ב-Windows Update...'
    $session = New-Object -ComObject Microsoft.Update.Session
    $session.ClientApplicationID = 'UpdateHub'
    $searcher = $session.CreateUpdateSearcher()
    $searcher.Online = $true
    try {
        $sm = New-Object -ComObject Microsoft.Update.ServiceManager
        $sm.ClientApplicationID = 'UpdateHub'
        [void]$sm.AddService2($script:MuService, 7, '')
        $searcher.ServerSelection = 3
        $searcher.ServiceID = $script:MuService
        Write-Ev 3 'windowsupdate' 'מקור: Microsoft Update (כולל Office ומוצרי Microsoft נוספים)'
    } catch { Write-Ev 3 'windowsupdate' ('Microsoft Update לא נרשם, משתמש ב-Windows Update: ' + $_.Exception.Message) 'warn' }

    $found = @{}; $list = New-Object System.Collections.ArrayList
    foreach ($crit in @("IsInstalled=0 and IsHidden=0", "IsInstalled=0 and IsHidden=0 and BrowseOnly=1")) {
        Test-Cancel
        Write-Ev 3 'windowsupdate' ("Search(`"$crit`")") 'cmd'
        try { $res = Invoke-WuaJob 'search' $searcher $null $crit 30 }
        catch {
            if (Test-IsCancel $_) { throw }
            Write-Ev 3 'windowsupdate' ('החיפוש נכשל: ' + $_.Exception.Message + ' — מנסה שוב עם ברירת המחדל') 'warn'
            $searcher.ServerSelection = 0; $res = Invoke-WuaJob 'search' $searcher $null $crit 30
        }
        Write-Ev 3 'windowsupdate' ("ResultCode={0} · {1} עדכונים" -f $res.ResultCode, $res.Updates.Count)
        for ($i = 0; $i -lt $res.Updates.Count; $i++) {
            $u = $res.Updates.Item($i); $uid = $u.Identity.UpdateID
            if (-not $found.ContainsKey($uid)) { $found[$uid] = 1; [void]$list.Add($u) }
        }
    }
    $exRx = [string](Get-Setting 'excludeWuTitle')
    if ($list.Count -eq 0) { $null = Add-Entry $P 'Windows Update' 'Windows Update' '' '' 'current'; return }
    foreach ($u in $list) {
        Test-Cancel
        $isDrv = ($u.Type -eq 2)
        $avail = ''
        if ($isDrv) { $avail = [string]$u.DriverVerVersion } elseif ($u.KBArticleIDs.Count -gt 0) { $avail = 'KB' + (@($u.KBArticleIDs) -join ',KB') }
        $kind = 'עדכון'; if ($isDrv) { $kind = 'דרייבר' }; if ($u.BrowseOnly) { $kind += ' אופציונלי' }
        $it = Add-Entry $P $u.Title ("Windows Update · $kind") '' $avail 'available'
        $cats = @($u.Categories | ForEach-Object { $_.Name }) -join ', '
        Write-Ev 3 'windowsupdate' ("{0} | Type={1} BrowseOnly={2} | {3:N1} MB | {4}" -f $u.Title, $u.Type, $u.BrowseOnly, ($u.MaxDownloadSize / 1MB), $cats)
        if ($exRx -and $u.Title -match $exRx) { Set-Entry $it 'skipped' 'הוחרג בהגדרות'; continue }
        # Feature upgrades (new Windows version) are a different risk class from monthly/optional updates.
        $isUpgrade = @($u.Categories | Where-Object { $_.CategoryID -eq '3689bdc8-b205-4af4-8d4a-a63924c5e9d5' -or $_.Name -eq 'Upgrades' }).Count -gt 0
        if ($isUpgrade -and -not (Get-Setting 'featureUpgrades')) { Set-Entry $it 'skipped' 'שדרוג גרסת Windows — כבוי (featureUpgrades בהגדרות)'; continue }
        if ($u.Title -match 'Preview|תצוגה מקדימה' -and -not (Get-Setting 'previewUpdates')) { Set-Entry $it 'skipped' 'עדכון Preview — כבוי (previewUpdates בהגדרות)'; continue }
        if ($isDrv) {
            # Same rules as the catalog path: firmware only with the running-version check, never a downgrade, honour excludeDevices.
            if ([string]$u.DriverClass -eq 'Firmware') { Set-Entry $it 'skipped' 'קושחה מותקנת בסעיף הקטלוג (שם נבדקת מול הקושחה הפעילה, חשמל ו-BitLocker)'; continue }
            $chk = Test-WuDriverSafe $u
            if ($chk) { Set-Entry $it $chk.Status $chk.Reason; Write-Ev 2 'windowsupdate' ($u.Title + ': ' + $chk.Reason); continue }
        }
        if ($Mode -ne 'Update') { continue }

        Ensure-RestorePoint
        try {
            if (-not $u.EulaAccepted) { $u.AcceptEula() }
            $coll = New-Object -ComObject Microsoft.Update.UpdateColl; [void]$coll.Add($u)
            if (-not $u.IsDownloaded) {
                Set-Entry $it 'downloading'; Set-Current ('מוריד: ' + $u.Title)
                $dl = $session.CreateUpdateDownloader(); $dl.Updates = $coll
                $dr = Invoke-WuaJob 'download' $dl $it $u.Title 120
                Write-Ev 3 'windowsupdate' ("Download ResultCode={0} HResult=0x{1:X8}" -f $dr.ResultCode, $dr.HResult)
                if ([int]$dr.ResultCode -notin 2, 3) { Set-Entry $it 'failed' ("ההורדה נכשלה: HResult 0x{0:X8}" -f $dr.HResult); continue }
            }
            Set-Entry $it 'installing'; Set-Current ('מתקין: ' + $u.Title)
            $inst = $session.CreateUpdateInstaller()
            try { $inst.AllowSourcePrompts = $false } catch { }
            $waited = 0; while ($inst.IsBusy -and $waited -lt 600) { Start-Sleep -Seconds 5; $waited += 5 }
            $inst.Updates = $coll
            $ir = Invoke-WuaJob 'install' $inst $it $u.Title 180; $r0 = $ir.GetUpdateResult(0)
            Write-Ev 3 'windowsupdate' ("Install ResultCode={0} HResult=0x{1:X8} Reboot={2}" -f $r0.ResultCode, $r0.HResult, $ir.RebootRequired)
            switch ([int]$r0.ResultCode) {
                2       { if ($ir.RebootRequired -or $r0.RebootRequired) { Set-Entry $it 'reboot' 'הותקן — נדרש אתחול' } else { Set-Entry $it 'updated' 'הותקן' } }
                3       { Set-Entry $it 'reboot' 'הותקן עם אזהרות' }
                default { Set-Entry $it 'failed' ("נכשל: ResultCode {0}, HResult 0x{1:X8}" -f $r0.ResultCode, $r0.HResult) }
            }
        } catch { if (Test-IsCancel $_) { throw }; Set-Entry $it 'failed' $_.Exception.Message }
    }
}

function Run-Drivers($P) {
    # Windows Update may have just replaced drivers: re-read versions so the catalog never re-installs what WU installed.
    $wuChanged = @($script:Prov['windowsupdate'].items | Where-Object { $_.status -in 'updated', 'reboot' }).Count
    if (-not $script:Devices -or $wuChanged -gt 0) {
        if ($wuChanged -gt 0) { Write-Ev 3 'drivers' 'Windows Update שינה רכיבים — קורא מחדש את גרסאות הדרייברים' }
        $script:Devices = Get-DeviceInventory
    }
    $cands = Select-CatalogCandidates $script:Devices
    Write-Ev 2 'drivers' ("נבדקים {0} רכיבים בקטלוג Microsoft (מתוך {1})" -f $cands.Count, $script:Devices.Count)
    foreach ($d in $cands) {
        Test-Cancel
        if ($script:CatFailures -ge 3) { Set-ProvFailed $P 'קטלוג Microsoft לא זמין כרגע — ייבדק בריצה הבאה'; return }
        $src = 'קטלוג'; if ($d.InstanceId -like 'UEFI\RES_*') { $src = 'קטלוג · קושחה' }
        $inst = $d.DriverVersion; if (-not $inst) { $inst = 'ללא דרייבר' }
        $it = Add-Entry $P $d.Name $src $inst
        Set-Current ('בודק: ' + $d.Name)
        Write-Ev 3 'drivers' ("— {0} | {1} | {2} | match={3}" -f $d.Name, $d.DriverVersion, $d.Provider, $d.MatchingId) 'step'
        $m = Find-CatalogDriver $d
        if ($m -and $m.Failed) { Set-Entry $it 'skipped' 'הבדיקה בקטלוג נכשלה (רשת?) — תיבדק בריצה הבאה'; continue }
        if (-not $m) { Set-Entry $it 'current' 'אין גרסה מתאימה בקטלוג'; continue }
        $cmp = 1; if ($d.DriverVersion) { $cmp = Compare-Ver $m.VersionString $d.DriverVersion }
        if ($cmp -le 0) { Set-Entry $it 'current' '' $m.VersionString; continue }
        $dec = Get-Decision $d $m
        if ($dec) { Set-Entry $it ([string]$dec.status) ("{0} (נבדק {1})" -f $dec.reason, $dec.at) $m.VersionString; continue }
        Set-Entry $it 'available' $m.Row.Title $m.VersionString
        Write-Ev 2 'drivers' ("{0}: {1} → {2}" -f $d.Name, $inst, $m.VersionString)
        if ($Mode -eq 'Update') { Install-CatalogDriver $d $m $it }
    }
}

function Run-Nvidia($P) {
    $gpu = Get-CimInstance Win32_VideoController | Where-Object { $_.Name -match 'NVIDIA' } | Select-Object -First 1
    if (-not $gpu) { Set-Prov $P 'skipped' 'לא נמצא כרטיס NVIDIA'; return }
    $it = Add-Entry $P $gpu.Name 'NVIDIA API' $gpu.DriverVersion
    $psid = switch -Regex ($gpu.Name) {
        # break: switch -Regex would otherwise emit EVERY matching value ("Quadro RTX 4000" → 130 163 73)
        'RTX\s+PRO|Quadro RTX' { 130; break } 'RTX\s+[56]\d{2}' { 163; break } 'RTX\s+4\d{2}' { 163; break } 'RTX\s+3\d{2}' { 120; break }
        'RTX\s+2\d{2}' { 105; break } 'GTX\s+16\d{2}' { 105; break } 'GTX\s+10\d{2}' { 98; break } 'GTX\s+9\d{2}' { 94; break }
        'Quadro' { 73; break } 'Tesla' { 69; break } default { 120 }
    }
    $isNb = ($gpu.Name -match 'Laptop|Mobile|Max-Q|MXM')
    $gpuName = ($gpu.Name -replace '^NVIDIA\s+', '' -replace '\s+with\s+Max-Q.*', '').Trim()
    $xml = [xml](Invoke-Http 'https://www.nvidia.com/Download/API/lookupValueSearch.aspx?TypeID=3' 'nvidia').Content
    $vals = @($xml.LookupValueSearch.LookupValues.LookupValue)
    $pf = $vals | Where-Object { $_.Name -eq $gpuName } | Select-Object -First 1
    if (-not $pf) { $pf = $vals | Where-Object { $_.Name -match ('^' + [regex]::Escape($gpuName) + '$') } | Select-Object -First 1 }
    if (-not $pf) { $pf = $vals | Where-Object { $_.Name -match [regex]::Escape($gpuName) } | Select-Object -First 1 }
    $pfid = '0'; if ($pf) { $pfid = [string]$pf.Value }
    Write-Ev 3 'nvidia' ("psid=$psid pfid=$pfid notebook=$isNb")
    $nb = ''; if ($isNb) { $nb = '&isNotebook=1' }
    $api = 'https://gfwsl.geforce.com/services_toolkit/services/com/nvidia/services/AjaxDriverService.php' +
           "?func=DriverManualLookup&psid=$psid&pfid=$pfid&osID=57&languageCode=1033&isWHQL=1&dltype=-1&dch=1&upCRD=0&sort1=0&numberOfResults=1$nb"
    $raw = (Invoke-Http $api 'nvidia').Content
    if ($raw -match '"Success"\s*:\s*"0"') { Set-Entry $it 'failed' 'NVIDIA לא החזיר דרייבר לכרטיס הזה'; return }
    $ver = $null; $url = $null
    if ($raw -match '"Version"\s*:\s*"([\d.]+)"') { $ver = $Matches[1] }
    if ($raw -match '"DownloadURL"\s*:\s*"(https://[^"]+\.exe)"') { $url = $Matches[1] }
    if (-not $ver -or -not $url) { Set-Entry $it 'failed' 'תשובת NVIDIA לא ניתנת לפענוח'; return }
    # Windows 32.0.15.9597  ↔  NVIDIA 595.97  (last 5 digits of the last two fields)
    $wp = $gpu.DriverVersion -split '\.'; $inst = 0
    if ($wp.Count -ge 4) { $s = $wp[2] + $wp[3]; $inst = [int]$s.Substring([Math]::Max(0, $s.Length - 5)) }
    $vp = $ver -split '\.'; $avail = [int]($vp[0] + $vp[1].PadRight(2, '0'))
    Write-Ev 3 'nvidia' ("מותקן {0} ({1}) · זמין {2} ({3})" -f $gpu.DriverVersion, $inst, $ver, $avail)
    if ($avail -le $inst) { Set-Entry $it 'current' '' $ver; return }
    Set-Entry $it 'available' '' $ver
    if ($Mode -ne 'Update') { return }
    $ndir = Join-Path $script:WorkDir ([guid]::NewGuid().ToString('N')); New-Item -ItemType Directory -Force -Path $ndir | Out-Null
    $dest = Join-Path $ndir ("nvidia-$ver.exe")   # random admin-only folder: no one can pre-plant or swap the file
    try {
        Set-Entry $it 'downloading'; Set-Current "מוריד NVIDIA $ver..."
        Save-Download $url $dest 'nvidia' $it "NVIDIA $ver"
        $sig = Get-AuthenticodeSignature $dest
        Write-Ev 3 'nvidia' ("חתימה: {0} · {1}" -f $sig.Status, $sig.SignerCertificate.Subject)
        if ([string]$sig.Status -ne 'Valid' -or (Get-CertOrg $sig.SignerCertificate) -ne 'NVIDIA Corporation') { Set-Entry $it 'failed' 'חתימת הקובץ אינה של NVIDIA'; return }
        Set-Entry $it 'installing'; Set-Current "מתקין NVIDIA $ver (המסך עשוי להבהב)..."
        $r = Invoke-External $dest '-s -noreboot' 'nvidia' 1800
        $after = (Get-CimInstance Win32_VideoController | Where-Object { $_.Name -match 'NVIDIA' } | Select-Object -First 1).DriverVersion
        Write-Ev 3 'nvidia' ("אימות: לפני {0} · אחרי {1}" -f $gpu.DriverVersion, $after)
        if ($after -ne $gpu.DriverVersion) { Set-Entry $it 'reboot' ("עודכן ל-$ver — מומלץ אתחול") }
        elseif ($r.ExitCode -in 0, 1) { Set-Entry $it 'reboot' 'ההתקנה הסתיימה — הגרסה תתעדכן אחרי אתחול' }
        else { Set-Entry $it 'failed' ("המתקין החזיר קוד " + $r.ExitCode) }
    } catch { if (Test-IsCancel $_) { throw }; Set-Entry $it 'failed' $_.Exception.Message }
    finally { Remove-Item $ndir -Recurse -Force -ErrorAction SilentlyContinue }
}

function Get-OfficeVersion {
    foreach ($rp in 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration', 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration') {
        $c = Get-ItemProperty $rp -ErrorAction SilentlyContinue
        if ($c -and $c.VersionToReport) { return $c }
    }
    return $null
}
function Run-Office($P) {
    $cfg = Get-OfficeVersion
    $c2r = Join-Path ([Environment]::GetFolderPath('CommonProgramFiles')) 'microsoft shared\ClickToRun\OfficeC2RClient.exe'
    if (-not $cfg -or -not (Test-Path $c2r)) { Set-Prov $P 'skipped' 'Microsoft 365 (Click-to-Run) לא מותקן'; return }
    $cur = [string]$cfg.VersionToReport
    $it = Add-Entry $P 'Microsoft 365 Apps' 'Office CDN' $cur
    # Latest build for THIS machine's channel, straight from the channel's version manifest.
    $base = [string]$cfg.UpdateChannel; if (-not $base) { $base = [string]$cfg.CDNBaseUrl }
    $plat = [string]$cfg.Platform; if (-not $plat) { $plat = 'x64' }
    $latest = $null
    if ($base) {
        $cabName = 'v64.cab'; if ($plat -eq 'x86') { $cabName = 'v32.cab' }
        $dir = Join-Path $script:WorkDir ([guid]::NewGuid().ToString('N')); New-Item -ItemType Directory -Force -Path $dir | Out-Null
        try {
            $cab = Join-Path $dir $cabName
            Save-Download ($base.TrimEnd('/') + '/Office/Data/' + $cabName) $cab 'office'
            $null = Invoke-External (Join-Path $script:Sys32 'expand.exe') ('"' + $cab + '" -F:VersionDescriptor.xml "' + $dir + '"') 'office' 120 -Interruptible
            [xml]$vd = Get-Content -Raw -Path (Join-Path $dir 'VersionDescriptor.xml')
            $a = $vd.Version.Available
            if ($plat -eq 'x86') { $latest = [string]$a.I320Version } else { $latest = [string]$a.I640Version }
            if (-not $latest) { $latest = [string](@($vd.Version.AvailableBuilds.AvailableBuild | Where-Object { $_.RestrictionType -eq 'Default' })[0].Version) }
            Write-Ev 3 'office' ("ערוץ: $base · פלטפורמה: $plat · מותקן $cur · אחרון בערוץ $latest")
        } catch { Write-Ev 2 'office' ('לא ניתן לקרוא את גרסת הערוץ: ' + $_.Exception.Message) 'warn' }
        finally { Remove-Item $dir -Recurse -Force -ErrorAction SilentlyContinue }
    }
    if ($latest -and (Compare-Ver $latest $cur) -le 0) { Set-Entry $it 'current' 'הגרסה האחרונה בערוץ שלך' $latest; return }
    if ($latest) { Set-Entry $it 'available' '' $latest } else { Set-Entry $it 'available' 'לא ידועה הגרסה האחרונה — מפעיל בדיקה' }
    if ($Mode -ne 'Update') { return }
    # Force-closing Office loses unsaved work, and a "scheduled" run may really be a catch-up run while the user works.
    # So: never by default; only if the user opted in (officeForceClose) AND the run is the scheduled one.
    $open = @(Get-Process -Name WINWORD, EXCEL, POWERPNT, OUTLOOK, ONENOTE, MSACCESS, MSPUB, VISIO, WINPROJ, lync -ErrorAction SilentlyContinue |
              Select-Object -ExpandProperty ProcessName -Unique)
    $force = 'false'; if ($Scheduled -and (Get-Setting 'officeForceClose')) { $force = 'true' }
    if ($open.Count -gt 0) { Write-Ev 2 'office' ("פתוחים: {0} · סגירה בכוח: {1}" -f ($open -join ', '), $force) }
    Set-Entry $it 'installing' 'Office מוריד ומתקין ברקע'
    Set-Current 'מעדכן Microsoft 365...'
    $null = Invoke-External $c2r "/update user displaylevel=false forceappshutdown=$force updatepromptuser=false" 'office' 300
    if ($open.Count -gt 0 -and $force -eq 'false') {
        Set-Entry $it 'available' ("Office פתוח ({0}) ולא נסגר בכוח — העדכון יושלם אחרי שתסגור את Office, או בריצה הבאה" -f ($open -join ', '))
        return
    }
    # C2R hands off to ClickToRunSvc; poll the real installed version.
    $minutes = 25; if (-not $latest) { $minutes = 4 }
    $deadline = (Get-Date).AddMinutes($minutes); $now = $cur
    while ((Get-Date) -lt $deadline) {
        Test-Cancel
        $It.reason = ("Office מתעדכן ברקע · עד {0:N0} דקות" -f ($deadline - (Get-Date)).TotalMinutes); Save-State -Force
        Start-Sleep -Seconds 20
        $cfgNow = Get-OfficeVersion
        if (-not $cfgNow) { continue }                      # registry briefly unavailable mid-update
        $now = [string]$cfgNow.VersionToReport
        Write-Ev 3 'office' ("גרסה מותקנת כרגע: $now")
        if ($now -ne $cur -and (-not $latest -or (Compare-Ver $now $latest) -ge 0)) { break }
    }
    if ($now -ne $cur) { Set-Entry $it 'updated' ("עודכן: $cur → $now") $now }
    elseif ($latest) { Set-Entry $it 'available' 'העדכון ממשיך ברקע (ייתכן שאפליקציית Office פתוחה) — ייבדק בריצה הבאה' }
    else { Set-Entry $it 'current' 'Office לא מצא עדכון' }
}

$script:GalleryReady = $false
function Initialize-Gallery {
    # Find-Module/Install-Module on PS 5.1 otherwise stop at an interactive NuGet/untrusted-repo prompt — invisible in a hidden run.
    if ($script:GalleryReady) { return }
    if (-not (Get-PackageProvider -ListAvailable -Name NuGet -ErrorAction SilentlyContinue | Where-Object { $_.Version -ge [version]'2.8.5.201' })) {
        Write-Ev 3 'powershell' 'מתקין NuGet provider...'
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope AllUsers -ErrorAction Stop | Out-Null
    }
    # Trust PSGallery for this run only; the user's original policy is restored when the engine finishes.
    $repo = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
    if ($repo -and $repo.InstallationPolicy -ne 'Trusted') {
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue
        $script:GalleryRestore = [string]$repo.InstallationPolicy
    }
    $script:GalleryReady = $true
}
$script:GalleryRestore = ''
function Restore-Gallery {
    if (-not $script:GalleryRestore) { return }
    try { Set-PSRepository -Name PSGallery -InstallationPolicy $script:GalleryRestore -ErrorAction Stop; $script:GalleryRestore = '' } catch { }
}

function Initialize-WinGetModule {
    if (Get-Module -ListAvailable -Name Microsoft.WinGet.Client) {
        try { Import-Module Microsoft.WinGet.Client -ErrorAction Stop; return $true }
        catch { Write-Ev 2 'winget' ('המודול לא נטען: ' + $_.Exception.Message) 'warn'; return $false }
    }
    Write-Ev 2 'winget' 'מתקין את מודול Microsoft.WinGet.Client (פעם אחת)...'
    try {
        Initialize-Gallery
        Install-Module Microsoft.WinGet.Client -Scope AllUsers -Force -AllowClobber -ErrorAction Stop
        Import-Module Microsoft.WinGet.Client -ErrorAction Stop
        return $true
    } catch { Write-Ev 2 'winget' ('המודול לא הותקן, עובר למצב שורת פקודה: ' + $_.Exception.Message) 'warn'; return $false }
}

# winget.exe on PATH is an alias in %LOCALAPPDATA%\Microsoft\WindowsApps — a folder your unelevated processes can write
# to. The elevated engine runs the real binary from the App Installer package folder (TrustedInstaller-owned) instead.
function Get-TrustedWinget {
    try {
        $pkgs = @(Get-AppxPackage -Name 'Microsoft.DesktopAppInstaller' -ErrorAction Stop | Sort-Object { [version]$_.Version } -Descending)
        foreach ($pk in $pkgs) {
            $exe = Join-Path $pk.InstallLocation 'winget.exe'
            if (-not (Test-Path -LiteralPath $exe)) { continue }
            if (Test-AdminOwned $exe) { return $exe }
            # Package folders can refuse ACL reads; C:\Program Files\WindowsApps itself is TrustedInstaller-only.
            $root = Join-Path ([Environment]::GetFolderPath('ProgramFiles')) 'WindowsApps\'
            $sig = Get-AuthenticodeSignature -FilePath $exe
            if ($exe.StartsWith($root, [StringComparison]::OrdinalIgnoreCase) -and [string]$sig.Status -eq 'Valid' -and
                (Get-CertOrg $sig.SignerCertificate) -eq 'Microsoft Corporation') { return $exe }
        }
    } catch { Write-Ev 3 'winget' ('Get-AppxPackage: ' + $_.Exception.Message) 'warn' }
    return $null
}

function Run-Winget($P) {
    $wg = Get-TrustedWinget
    if (-not $wg) {
        if (Get-Command winget.exe -ErrorAction SilentlyContinue) { Set-ProvFailed $P 'winget נמצא רק כקיצור בתיקיית המשתמש — לא יורץ כמנהל (עדכן את App Installer מה-Store)' }
        else { Set-Prov $P 'skipped' 'winget לא מותקן' }
        return
    }
    Write-Ev 3 'winget' ("winget: $wg")
    Set-Current 'מרענן מקורות winget...'
    $null = Invoke-External $wg 'source update' 'winget' 300 -Interruptible
    $exclude = @(Get-Setting 'excludeWinget')
    $useModule = Initialize-WinGetModule
    $pk = @()
    if ($useModule) {
        Set-Current 'בודק עדכונים לכל התוכנות...'
        try { $pk = @(Get-WinGetPackage -ErrorAction Stop | Where-Object { $_.IsUpdateAvailable }) }
        catch { Write-Ev 2 'winget' ('המודול נכשל בקריאת הרשימה, עובר לשורת פקודה: ' + $_.Exception.Message) 'warn'; $useModule = $false }
    }
    if ($useModule) {
        Write-Ev 2 'winget' ("{0} תוכנות עם עדכון זמין" -f $pk.Count)
        if ($pk.Count -eq 0) { $null = Add-Entry $P 'כל התוכנות' 'winget' '' '' 'current'; return }
        foreach ($pkg in $pk) {
            Test-Cancel
            $avail = [string](@($pkg.AvailableVersions)[0])
            $it = Add-Entry $P $pkg.Name ('winget · ' + $pkg.Source) ([string]$pkg.InstalledVersion) $avail 'available'
            Write-Ev 3 'winget' ("{0} [{1}] {2} → {3}" -f $pkg.Name, $pkg.Id, $pkg.InstalledVersion, $avail)
            if ($exclude -contains $pkg.Id) { Set-Entry $it 'skipped' 'הוחרג בהגדרות'; continue }
            # Ids/sources go into a command line: only the characters winget itself uses.
            if ([string]$pkg.Id -notmatch '^[A-Za-z0-9][A-Za-z0-9._+-]{0,127}$' -or ($pkg.Source -and [string]$pkg.Source -notmatch '^[A-Za-z0-9._-]{1,64}$')) {
                Set-Entry $it 'skipped' 'מזהה חבילה לא תקין — לא יותקן'; continue
            }
            if ($Mode -ne 'Update') { continue }
            Set-Entry $it 'installing'; Set-Current ('מעדכן: ' + $pkg.Name)
            try {
                # The install runs through the CLI: hard 30-minute limit, installer output streamed to VERBOSE.
                # The module is used only to list and to re-read the installed version afterwards.
                $before = [string]$pkg.InstalledVersion
                $src = ''; if ($pkg.Source) { $src = ' --source "' + $pkg.Source + '"' }
                $r = Invoke-External $wg ('upgrade --id "' + $pkg.Id + '" --exact' + $src + ' --silent --accept-package-agreements --accept-source-agreements --disable-interactivity') 'winget' 1800
                $hex = '{0:X8}' -f $r.ExitCode
                $after = [string](Get-WinGetPackage -Id $pkg.Id -MatchOption Equals -ErrorAction SilentlyContinue | Select-Object -First 1).InstalledVersion
                Write-Ev 3 'winget' ("אימות: {0} → {1} · קוד 0x{2}" -f $before, $after, $hex)
                $changed = ($after -and $after -ne $before)
                $rebootHint = ($r.ExitCode -in 3010, 1641) -or $hex -eq '8A150109'    # 0x8A150109 = reboot required to finish
                if ($changed -and $rebootHint) { Set-Entry $it 'reboot' ("עודכן ל-$after — נדרש אתחול") $after }
                elseif ($changed)               { Set-Entry $it 'updated' ("עודכן ל-$after") $after }
                elseif ($rebootHint)            { Set-Entry $it 'reboot' 'ההתקנה תושלם אחרי אתחול' }
                elseif ($r.ExitCode -eq -2)     { Set-Entry $it 'failed' 'המתקין חרג מ-30 דקות ונעצר' }
                elseif ($r.ExitCode -eq 0)      { Set-Entry $it 'failed' "המתקין דיווח הצלחה אך הגרסה נשארה $after (תוכנה פתוחה או גרסה שמדווחת אחרת — אפשר להחריג ב-excludeWinget)" }
                else                            { Set-Entry $it 'failed' ("winget החזיר 0x$hex") }
            } catch { if (Test-IsCancel $_) { throw }; Set-Entry $it 'failed' $_.Exception.Message }
        }
    } else {
        $it = Add-Entry $P 'כל התוכנות (מצב שורת פקודה)' 'winget' '' '' 'checking'
        if ($Mode -ne 'Update') { $null = Invoke-External $wg 'upgrade --include-unknown --accept-source-agreements' 'winget' 300 -Interruptible; Set-Entry $it 'info' 'ראה פירוט ב-VERBOSE'; return }
        Set-Entry $it 'installing'
        $r = Invoke-External $wg 'upgrade --all --include-unknown --silent --accept-package-agreements --accept-source-agreements --disable-interactivity' 'winget' 5400
        # Without the module the result cannot be verified per package — reported as info, not as "updated".
        if ($r.ExitCode -eq 0) { Set-Entry $it 'info' 'הסתיים ללא שגיאה — פירוט לכל תוכנה ב-VERBOSE' } else { Set-Entry $it 'failed' ("winget החזיר 0x{0:X8}" -f $r.ExitCode) }
    }
}

function Run-Store($P) {
    $it = Add-Entry $P 'אפליקציות Microsoft Store' 'Store' ''
    try {
        Get-CimInstance -Namespace 'Root\cimv2\mdm\dmmap' -ClassName 'MDM_EnterpriseModernAppManagement_AppManagement01' -ErrorAction Stop |
            Invoke-CimMethod -MethodName UpdateScanMethod -ErrorAction Stop | Out-Null
        Set-Entry $it 'info' 'סריקת Store הופעלה — Store מתקין ברקע. אפליקציות Store שמנוהלות ב-winget מאומתות בסעיף התוכנות'
    } catch { Set-Entry $it 'info' ('Store לא זמין לסריקה מרחוק: ' + $_.Exception.Message) }
}

function Run-PowerShell($P) {
    $mods = @(Get-InstalledModule -ErrorAction SilentlyContinue)
    Write-Ev 3 'powershell' ("{0} מודולים מותקנים מ-PSGallery" -f $mods.Count)
    if ($mods.Count -gt 0) { try { Initialize-Gallery } catch { Write-Ev 2 'powershell' ('PSGallery לא זמין: ' + $_.Exception.Message) 'warn'; $mods = @() } }
    $pf = [Environment]::GetFolderPath('ProgramFiles')
    foreach ($mod in $mods) {
        Test-Cancel
        $it = Add-Entry $P $mod.Name 'PSGallery' ([string]$mod.Version)
        # Modules installed for CurrentUser live in your Documents folder: the elevated engine leaves them alone
        # (it would write admin-owned files into your profile).
        if ($mod.InstalledLocation -and -not ([string]$mod.InstalledLocation).StartsWith($pf, [StringComparison]::OrdinalIgnoreCase)) {
            Set-Entry $it 'skipped' 'מותקן ל-CurrentUser (Documents) — לא מתעדכן כמנהל'; continue
        }
        try {
            $online = Find-Module -Name $mod.Name -ErrorAction Stop
            if ((Compare-Ver ([string]$online.Version) ([string]$mod.Version)) -le 0) { Set-Entry $it 'current' '' ([string]$online.Version); continue }
            Set-Entry $it 'available' '' ([string]$online.Version)
            if ($Mode -ne 'Update') { continue }
            Set-Entry $it 'installing'; Set-Current ('מעדכן מודול: ' + $mod.Name)
            # Separate process: 15-minute limit, output in VERBOSE, and a stuck module cannot stall the whole run.
            $name = $mod.Name -replace "'", "''"
            $ur = Invoke-External (Join-Path $PSHOME 'powershell.exe') ("-NoProfile -NonInteractive -Command `"[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor 3072; Update-Module -Name '$name' -Force -ErrorAction Stop -WarningAction SilentlyContinue`"") 'powershell' 900
            if ($ur.ExitCode -eq -2) { Set-Entry $it 'failed' 'חרג מ-15 דקות ונעצר'; continue }
            $after = [string](Get-InstalledModule -Name $mod.Name -ErrorAction SilentlyContinue).Version
            if ($after -eq [string]$online.Version) { Set-Entry $it 'updated' ("עודכן ל-$after") } else { Set-Entry $it 'failed' "הגרסה לא השתנתה ($after)" }
        } catch { if (Test-IsCancel $_) { throw }; Set-Entry $it 'failed' $_.Exception.Message }
    }
    $stamp = Join-Path $DataDir 'help.stamp'
    $days = [int](Get-Setting 'helpUpdateDays')
    $last = [datetime]::MinValue
    if (Test-Path $stamp) { try { $last = [datetime](Get-Content $stamp -Raw).Trim() } catch { } }
    $it = Add-Entry $P 'קבצי עזרה (Update-Help)' 'Microsoft' ''
    if (((Get-Date) - $last).TotalDays -lt $days) { Set-Entry $it 'current' ("עודכנו ב-" + $last.ToString('dd/MM')); return }
    if ($Mode -ne 'Update') { Set-Entry $it 'available'; return }
    Set-Entry $it 'installing'; Set-Current 'מעדכן קבצי עזרה של PowerShell...'
    # Separate process with a hard time limit: Update-Help can stall for a long time on unreachable help URIs.
    # en-US explicitly: there is no he-IL help, so the UI culture would make every module fail. Interruptible = harmless to stop.
    $hr = Invoke-External (Join-Path $PSHOME 'powershell.exe') '-NoProfile -NonInteractive -Command "[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor 3072; Update-Help -UICulture en-US -Force -ErrorAction SilentlyContinue"' 'powershell' 900 -Interruptible
    if ($hr.ExitCode -eq -2) { Set-Entry $it 'skipped' 'חרג מ-15 דקות — ינוסה שוב בריצה הבאה'; return }
    (Get-Date).ToString('s') | Set-Content $stamp
    # Not counted as "updated": Update-Help gives no reliable per-module result.
    Set-Entry $it 'info' "רועננו (פעם ב-$days ימים) — פירוט ב-VERBOSE"
}

function Get-RealPython {
    foreach ($n in 'python.exe', 'py.exe') {
        foreach ($c in @(Get-Command $n -All -ErrorAction SilentlyContinue)) {
            if ($c.Source -match '\\WindowsApps\\') { continue }   # Store alias stub opens the Store instead of running
            return $c.Source
        }
    }
    return $null
}
function Run-Pip($P) {
    $py = Get-RealPython
    if (-not $py) { Set-Prov $P 'skipped' 'Python לא מותקן'; return }
    Write-Ev 3 'pip' ("Python: $py")
    # A Python that only admins can modify (e.g. installed for all users in Program Files) is updated elevated;
    # one in your profile is updated with your normal rights.
    $forceUser = $false
    if ((Split-Path $py -Leaf) -ieq 'py.exe') {
        $forceUser = $true                     # until proven otherwise, the launcher's target is untrusted
        # py.exe sits in C:\Windows, but the Python it launches may live in your profile: judge the real interpreter.
        $pr = Invoke-AsUser $py '-c "import sys;print(sys.executable)"' 'pip' 60 -QuietStdout -Interruptible
        $real = @($pr.StdOut -split "`r?`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '\\pythonw?\.exe$' }) | Select-Object -First 1
        if ($real -and (Test-Path -LiteralPath $real)) { $py = $real; $forceUser = $false; Write-Ev 3 'pip' ("py.exe → $py") }
    }
    $asUser = $forceUser -or -not (Test-AdminOwned $py)
    if ($asUser) { Write-Ev 2 'pip' 'Python בפרופיל המשתמש — מעודכן בהרשאות משתמש רגיל (לא כמנהל)' }
    # Elevated Python runs isolated (-I): it ignores PYTHON* variables and the per-user site-packages, which any
    # unelevated process of yours can write to.
    $iso = ''; if (-not $asUser) { $iso = '-I ' }
    $lr = Invoke-Tool $asUser $py ($iso + '-m pip list --outdated --format=json --disable-pip-version-check') 'pip' 300 -QuietStdout -Interruptible
    $json = @($lr.StdOut -split "`r?`n" | Where-Object { $_.TrimStart().StartsWith('[') }) | Select-Object -First 1
    if ($lr.ExitCode -ne 0 -or -not $json) { Set-ProvFailed $P ("pip list נכשל (קוד {0}) — פירוט ב-VERBOSE" -f $lr.ExitCode); return }
    # PS 5.1 emits a JSON array as ONE pipeline object: assign first, then enumerate (else "a b c" becomes one package).
    $out = @(); try { $parsed = ConvertFrom-Json -InputObject $json -ErrorAction Stop; $out = @($parsed | Where-Object { $_.name }) } catch { }
    Write-Ev 3 'pip' ("{0} חבילות עם גרסה חדשה" -f $out.Count)
    if ($out.Count -eq 0) { $null = Add-Entry $P 'חבילות pip' 'PyPI' '' '' 'current'; return }
    foreach ($pkg in $out) {
        Test-Cancel
        $it = Add-Entry $P $pkg.name 'PyPI' $pkg.version $pkg.latest_version 'available'
        if ($Mode -ne 'Update') { continue }
        if ([string]$pkg.name -notmatch '^[A-Za-z0-9][A-Za-z0-9._-]*$') { Set-Entry $it 'skipped' 'שם חבילה לא תקין'; continue }
        Set-Entry $it 'installing'
        $r = Invoke-Tool $asUser $py ($iso + "-m pip install --upgrade `"$($pkg.name)`" --disable-pip-version-check") 'pip' 900
        if ($r.ExitCode -eq 0) { Set-Entry $it 'updated' ("עודכן ל-" + $pkg.latest_version) } else { Set-Entry $it 'failed' ("pip החזיר קוד " + $r.ExitCode) }
    }
    if ($Mode -eq 'Update') {
        $chk = Invoke-Tool $asUser $py ($iso + '-m pip check') 'pip' 300 -Interruptible
        if ($chk.ExitCode -ne 0) { Write-Ev 2 'pip' 'pip check מצא התנגשויות בין חבילות — פירוט ב-VERBOSE' 'warn' }
    }
}

function Run-VSCode($P) {
    $code = (Get-Command code.cmd -ErrorAction SilentlyContinue).Source
    if (-not $code) { Set-Prov $P 'skipped' 'VS Code לא מותקן'; return }
    # Extensions are per-user data in your profile: always with your normal rights.
    # Every call's exit code is checked: an empty list from a call that failed must never read as "up to date".
    $listArgs = '--list-extensions --show-versions'
    $lr = Invoke-AsUser $code $listArgs 'vscode' 180 -QuietStdout -Interruptible
    if ($lr.ExitCode -ne 0) { Set-ProvFailed $P ("לא ניתן להריץ את VS Code בהרשאות המשתמש (קוד {0}) — פירוט ב-VERBOSE" -f $lr.ExitCode); return }
    $before = @{}
    foreach ($l in @($lr.StdOut -split "`r?`n")) { if ($l -match '^(.+)@(.+)$') { $before[$Matches[1]] = $Matches[2] } }
    Write-Ev 3 'vscode' ("{0} תוספים מותקנים" -f $before.Count)
    if ($Mode -ne 'Update') { $null = Add-Entry $P ("{0} תוספים" -f $before.Count) 'Marketplace' '' '' 'info'; return }
    Set-Current 'מעדכן תוספי VS Code...'
    $ur = Invoke-AsUser $code '--update-extensions' 'vscode' 900
    $ar = Invoke-AsUser $code $listArgs 'vscode' 180 -QuietStdout
    if ($ar.ExitCode -ne 0) { Set-ProvFailed $P ("בדיקת התוספים אחרי העדכון נכשלה (קוד {0})" -f $ar.ExitCode); return }
    $changed = 0
    foreach ($l in @($ar.StdOut -split "`r?`n")) {
        if ($l -match '^(.+)@(.+)$' -and $before.ContainsKey($Matches[1]) -and $before[$Matches[1]] -ne $Matches[2]) {
            $changed++; $null = Add-Entry $P $Matches[1] 'Marketplace' $before[$Matches[1]] $Matches[2] 'updated'
        }
    }
    if ($ur.ExitCode -ne 0) { $e = Add-Entry $P 'עדכון תוספים' 'Marketplace' ''; Set-Entry $e 'failed' ("code --update-extensions החזיר קוד {0}" -f $ur.ExitCode) }
    elseif ($changed -eq 0) { $null = Add-Entry $P ("{0} תוספים נבדקו" -f $before.Count) 'Marketplace' '' '' 'current' }
}

function Run-Wsl($P) {
    $wsl = Join-Path $script:Sys32 'wsl.exe'
    # System32\wsl.exe is an inbox stub on every Windows 11 — the registered distros are the real signal.
    $lxss = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Lxss'
    $distros = @(Get-ChildItem $lxss -ErrorAction SilentlyContinue | ForEach-Object { [string](Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue).DistributionName } |
                 Where-Object { $_ -and $_ -notmatch '^docker-desktop' })
    if (-not (Test-Path $wsl) -or $distros.Count -eq 0) { Set-Prov $P 'skipped' 'אין הפצות WSL רשומות'; return }
    $it = Add-Entry $P 'WSL' 'Microsoft' ''
    if ($Mode -eq 'Update') {
        $r = Invoke-External $wsl '--update' 'wsl' 900
        if ($r.ExitCode -eq 0) { Set-Entry $it 'current' 'WSL מעודכן' } else { Set-Entry $it 'failed' ("wsl --update החזיר קוד " + $r.ExitCode) }
    } else { Set-Entry $it 'info' ("{0} הפצות רשומות" -f $distros.Count) }
    foreach ($d in $distros) {
        Test-Cancel
        $di = Add-Entry $P $d 'הפצת Linux' ''
        # The name comes from HKCU and goes into a command line: letters, digits, dot, dash, underscore only.
        if ($d -notmatch '^[A-Za-z0-9][A-Za-z0-9._-]{0,63}$') { Set-Entry $di 'manual' 'שם הפצה עם תווים לא צפויים — לא עודכנה'; continue }
        $pr = Invoke-External $wsl ('-d "' + $d + '" -- echo ok') 'wsl' 120 -Interruptible
        if ($pr.StdOut.Trim() -ne 'ok') { Set-Entry $di 'manual' ('ההפצה לא עולה: ' + (($pr.Output -replace '\s+', ' ').Trim())); continue }
        $pmr = Invoke-External $wsl ('-d "' + $d + '" -- sh -c "command -v apt-get || command -v dnf || command -v zypper || command -v pacman || command -v apk"') 'wsl' 60 -Interruptible
        $pm = ([string]$pmr.StdOut).Trim()
        # stdin from /dev/null + dpkg "keep my config" options: a config-file question must never stall a hidden run.
        $cmd = switch -Regex ($pm) {
            'apt-get' { 'exec </dev/null; export LC_ALL=C DEBIAN_FRONTEND=noninteractive; apt-get update -q && apt-get -y -q -o Dpkg::Options::=--force-confdef -o Dpkg::Options::=--force-confold upgrade' }
            'dnf'     { 'dnf -y upgrade' }
            'zypper'  { 'zypper --non-interactive update' }
            'pacman'  { 'pacman -Syu --noconfirm' }
            'apk'     { 'apk update && apk upgrade' }
            default   { $null }
        }
        if (-not $cmd) { Set-Entry $di 'manual' 'לא זוהה מנהל חבילות'; continue }
        if ($Mode -ne 'Update') {
            if ($pm -match 'apt-get') {
                $cr = Invoke-External $wsl ('-d "' + $d + '" -- sh -c "apt list --upgradable 2>/dev/null | grep -c upgradable"') 'wsl' 120 -Interruptible
                $n = 0; [void][int]::TryParse(([string]$cr.StdOut).Trim(), [ref]$n)
                if ($n -gt 0) { Set-Entry $di 'available' ("$n חבילות לפי הרשימה המקומית (apt update ירוץ בעדכון)") }
                else { Set-Entry $di 'current' 'לפי הרשימה המקומית — apt update ירוץ בעדכון' }
            } else { Set-Entry $di 'info' ("מנהל חבילות: " + (Split-Path $pm -Leaf)) }
            continue
        }
        Set-Entry $di 'installing'; Set-Current "מעדכן $d..."
        $r = Invoke-External $wsl ("-d `"$d`" -u root -- sh -c `"$cmd`"") 'wsl' 3600
        # Only "updated" when the package manager actually changed something.
        $nothing = ($r.Output -match '(?m)^0 upgraded, 0 newly installed|Nothing to do|there is nothing to do') -or
                   ($pm -match 'apk' -and $r.Output -notmatch 'Upgrading')
        if ($r.ExitCode -ne 0) { Set-Entry $di 'failed' ("קוד יציאה " + $r.ExitCode + ' — פירוט ב-VERBOSE') }
        elseif ($nothing)      { Set-Entry $di 'current' 'כל החבילות מעודכנות' }
        else                   { Set-Entry $di 'updated' 'חבילות עודכנו — פירוט ב-VERBOSE' }
    }
}

# Package names are validated before they ever reach a command line.
function Get-ChocoOutdated([string]$Choco) {
    $r = Invoke-External $Choco 'outdated -r --ignore-pinned' 'pkgmgr' 300 -QuietStdout -Interruptible
    if ($r.ExitCode -notin 0, 2) { return @{ Ok = $false; Items = @() } }          # 2 = "outdated found" (enhanced exit codes)
    $items = @($r.StdOut -split "`r?`n" | Where-Object { $_ -match '^[A-Za-z0-9][A-Za-z0-9._-]*\|[^|]*\|[^|]+\|' } |
               ForEach-Object { $f = $_ -split '\|'; [pscustomobject]@{ Name = $f[0]; Cur = $f[1]; New = $f[2] } })
    return @{ Ok = $true; Items = $items }
}
function Get-NpmOutdated([string]$Npm) {
    # npm outdated exits 1 when something is outdated, so the JSON — not the exit code — is the result.
    $r = Invoke-AsUser $Npm 'outdated -g --json' 'pkgmgr' 300 -QuietStdout -Interruptible
    if ($r.ExitCode -eq -3) { return @{ Ok = $false; Items = @() } }
    if ($r.ExitCode -eq -2) { return @{ Ok = $false; Items = @() } }
    $j = ([string]$r.StdOut).Trim()
    if (-not $j) { return @{ Ok = ($r.ExitCode -eq 0); Items = @() } }               # empty + exit 0 = nothing outdated
    try { $o = ConvertFrom-Json -InputObject $j -ErrorAction Stop } catch { return @{ Ok = $false; Items = @() } }
    if ($o.PSObject.Properties.Name -contains 'error') { return @{ Ok = $false; Items = @() } }   # {"error":{...}} on failure
    $rx = '^(@[a-z0-9][a-z0-9._~-]*/)?[a-z0-9][a-z0-9._~-]*$'
    $items = @(foreach ($pr in $o.PSObject.Properties) {
        $latest = [string]$pr.Value.latest
        if ($pr.Name -match $rx -and $latest) { [pscustomobject]@{ Name = $pr.Name; Cur = [string]$pr.Value.current; New = $latest } }
    })
    return @{ Ok = $true; Items = $items }
}

function Run-PkgMgr($P) {
    $any = $false
    # choco and npm: per-package rows from their "outdated" lists, verified by listing again after the upgrade.
    # A list that fails (exit code, error JSON, timeout) is "unknown", never "nothing outdated" — so a failed
    # re-list can not turn every row into "updated", and an error object can not become a package name.
    $choco = (Get-Command choco.exe -ErrorAction SilentlyContinue).Source
    if ($choco -and -not (Test-AdminOwned $choco)) {
        # Chocolatey needs admin rights, so it may only run from a folder that only admins can change.
        $any = $true; $e = Add-Entry $P 'Chocolatey' 'choco' ''
        Set-Entry $e 'manual' ("התיקייה של choco.exe ניתנת לשינוי ללא הרשאות מנהל — לא יורץ כמנהל ($choco)")
        $choco = $null
    }
    if ($choco) {
        $any = $true
        $old = Get-ChocoOutdated $choco
        if (-not $old.Ok) { $e = Add-Entry $P 'Chocolatey' 'choco' ''; Set-Entry $e 'failed' 'רשימת העדכונים של choco נכשלה — פירוט ב-VERBOSE' }
        elseif ($old.Items.Count -eq 0) { $null = Add-Entry $P 'Chocolatey' 'choco' '' '' 'current' }
        $rows = @{}; foreach ($o in $old.Items) { $rows[$o.Name] = Add-Entry $P $o.Name 'choco' $o.Cur $o.New 'available' }
        if ($Mode -eq 'Update' -and $rows.Count -gt 0) {
            foreach ($k in @($rows.Keys)) { $rows[$k].status = 'installing' }; Set-Current 'מעדכן חבילות Chocolatey...'; Save-State -Force
            $names = (@($rows.Keys) | Sort-Object) -join ' '
            $r = Invoke-External $choco ("upgrade $names -y --no-progress") 'pkgmgr' 5400
            $now = Get-ChocoOutdated $choco
            $rb = $r.ExitCode -in 1641, 3010
            foreach ($k in @($rows.Keys)) {
                if (-not $now.Ok) { Set-Entry $rows[$k] 'info' ("הסתיים (קוד choco {0}) — לא אומת: בדיקת הרשימה נכשלה" -f $r.ExitCode) }
                elseif (@($now.Items | Where-Object { $_.Name -eq $k }).Count -gt 0) { Set-Entry $rows[$k] 'failed' ("עדיין לא מעודכן (קוד choco {0}) — פירוט ב-VERBOSE" -f $r.ExitCode) }
                elseif ($rb) { Set-Entry $rows[$k] 'reboot' 'עודכן — נדרש אתחול' }
                else { Set-Entry $rows[$k] 'updated' ('עודכן ל-' + $rows[$k].available) }
            }
        }
    }
    $scoop = (Get-Command scoop.cmd -ErrorAction SilentlyContinue).Source
    if ($scoop) {
        $any = $true; $it = Add-Entry $P 'Scoop' 'scoop' ''
        if ($Mode -eq 'Update') {
            Set-Entry $it 'installing'
            $null = Invoke-AsUser $scoop 'update' 'pkgmgr' 600
            $r = Invoke-AsUser $scoop 'update *' 'pkgmgr' 5400
            # scoop has no machine-readable result: not counted as "updated".
            if ($r.ExitCode -eq 0) { Set-Entry $it 'info' 'הסתיים ללא שגיאה — פירוט ב-VERBOSE' } else { Set-Entry $it 'failed' ("קוד " + $r.ExitCode) }
        } else { $null = Invoke-AsUser $scoop 'status' 'pkgmgr' 300 -Interruptible; Set-Entry $it 'info' 'פירוט ב-VERBOSE' }
    }
    $npm = (Get-Command npm.cmd -ErrorAction SilentlyContinue).Source
    if ($npm) {
        $any = $true
        $old = Get-NpmOutdated $npm
        if (-not $old.Ok) { $e = Add-Entry $P 'npm (גלובלי)' 'npm' ''; Set-Entry $e 'failed' 'npm outdated נכשל (רשת?) — פירוט ב-VERBOSE' }
        elseif ($old.Items.Count -eq 0) { $null = Add-Entry $P 'npm (גלובלי)' 'npm' '' '' 'current' }
        $rows = @{}; foreach ($o in $old.Items) { $rows[$o.Name] = Add-Entry $P $o.Name 'npm' $o.Cur $o.New 'available' }
        if ($Mode -eq 'Update' -and $rows.Count -gt 0) {
            $codes = @{}
            foreach ($k in @($rows.Keys)) {
                Test-Cancel
                Set-Entry $rows[$k] 'installing'; Set-Current ('npm: ' + $k)
                $codes[$k] = (Invoke-AsUser $npm ('install -g "' + $k + '@latest"') 'pkgmgr' 900).ExitCode
            }
            $now = Get-NpmOutdated $npm
            foreach ($k in @($rows.Keys)) {
                if (-not $now.Ok) { Set-Entry $rows[$k] 'info' ("הסתיים (קוד npm {0}) — לא אומת: בדיקת הרשימה נכשלה" -f $codes[$k]) }
                elseif (@($now.Items | Where-Object { $_.Name -eq $k }).Count -gt 0) { Set-Entry $rows[$k] 'failed' ("עדיין לא מעודכן (קוד npm {0})" -f $codes[$k]) }
                else { Set-Entry $rows[$k] 'updated' ('עודכן ל-' + $rows[$k].available) }
            }
        }
    }
    if (-not $any) { Set-Prov $P 'skipped' 'לא מותקנים' }
}

function Run-Defender($P) {
    try { $before = Get-MpComputerStatus -ErrorAction Stop } catch { Set-Prov $P 'skipped' 'Defender לא פעיל (אנטי-וירוס אחר?)'; return }
    $it = Add-Entry $P 'הגדרות וירוסים' 'Microsoft' ([string]$before.AntivirusSignatureVersion)
    if ($Mode -ne 'Update') { Set-Entry $it 'info' ('עודכן לאחרונה: ' + $before.AntivirusSignatureLastUpdated); return }
    Set-Entry $it 'installing'
    try { Update-MpSignature -ErrorAction Stop } catch { Write-Ev 3 'defender' $_.Exception.Message 'warn' }
    $after = [string](Get-MpComputerStatus).AntivirusSignatureVersion
    if ($after -ne [string]$before.AntivirusSignatureVersion) { Set-Entry $it 'updated' "עודכן ל-$after" $after } else { Set-Entry $it 'current' 'כבר עדכני' $after }
}

$script:ProblemText = @{ 1 = 'לא מוגדר'; 3 = 'דרייבר פגום'; 10 = 'לא מצליח לעלות'; 12 = 'אין משאבים'; 14 = 'דורש אתחול'; 18 = 'דורש התקנה מחדש'
                         19 = 'רישום פגום'; 21 = 'בהסרה'; 24 = 'לא קיים/לא מוגדר'; 28 = 'אין דרייבר'; 31 = 'לא עובד כראוי'; 37 = 'הדרייבר נכשל'
                         39 = 'דרייבר חסר או פגום'; 43 = 'Windows עצר את ההתקן'; 48 = 'נחסם'; 52 = 'חתימה לא תקינה' }
function Run-Health($P) {
    $devs = Get-DeviceInventory
    $bad = @($devs | Where-Object { $_.Problem -ne 0 -and $_.Problem -ne 22 -and $_.Problem -ne 45 })
    foreach ($d in $bad) {
        $t = $script:ProblemText[[int]$d.Problem]; if (-not $t) { $t = 'בעיה' }
        $e = Add-Entry $P $d.Name 'Device Manager' $d.DriverVersion
        if ($d.Problem -eq 14) { Set-Entry $e 'reboot' 'ההתקן ממתין לאתחול (קוד 14)'; continue }
        Set-Entry $e 'manual' ("קוד {0}: {1}" -f $d.Problem, $t)
    }
    # Board BIOS on self-built PCs has no UEFI capsule (ESRT) → cannot be updated from Windows.
    $capsule = @($devs | Where-Object { $_.InstanceId -like 'UEFI\RES_*' })
    $m = $script:State.machine
    if ($capsule.Count -eq 0) {
        $e = Add-Entry $P ("BIOS לוח אם · " + $m.board) 'יצרן הלוח' $m.bios
        Set-Entry $e 'manual' 'ללוח הזה אין עדכון קושחה דרך Windows — עדכון BIOS נעשה מתוך תפריט ה-BIOS (EZ Flash / M-Flash / Q-Flash)'
    }
    if ($bad.Count -eq 0 -and $capsule.Count -gt 0) { $null = Add-Entry $P 'כל ההתקנים תקינים' 'Device Manager' '' '' 'current' }
    Save-State -Force
}

function Show-Toast([string]$Title, [string]$Text) {
    try {
        [void][Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime]
        [void][Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime]
        $xml = New-Object Windows.Data.Xml.Dom.XmlDocument
        $xml.LoadXml(("<toast><visual><binding template='ToastGeneric'><text>{0}</text><text>{1}</text></binding></visual></toast>" -f
            [Security.SecurityElement]::Escape($Title), [Security.SecurityElement]::Escape($Text)))
        $appId = '{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\WindowsPowerShell\v1.0\powershell.exe'
        [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($appId).Show((New-Object Windows.UI.Notifications.ToastNotification($xml)))
    } catch { Write-Ev 3 'system' ('התראה לא הוצגה: ' + $_.Exception.Message) }
}

# ════════════════════════════════════════════════════════════════════════════
#  Main
# ════════════════════════════════════════════════════════════════════════════
if ($MyInvocation.InvocationName -eq '.') { return }   # dot-sourced (tests): functions only

# ── one engine at a time ─────────────────────────────────────────────────────────────────────────
# A named mutex is atomic: two clicks, two tabs or the weekly task firing during a manual run can never start two
# engines. The server also reads it (its existence = an engine is alive). Released automatically if we crash.
$script:Mutex = New-Object System.Threading.Mutex($false, 'Global\UpdateHub.Engine')
$got = $false
try { $got = $script:Mutex.WaitOne(0) } catch [System.Threading.AbandonedMutexException] { $got = $true }   # previous engine died
if (-not $got) { Write-Host 'UpdateHub engine already running.'; exit 2 }

# ── DataDir hardening ─────────────────────────────────────────────────────────────────────────────
# Downloaded installers run elevated from DataDir\work, so only Administrators and SYSTEM may even read it (it also
# holds the device inventory and the dashboard's session key). Order matters:
#   (1) take ownership of everything WITHOUT following links (/L) — works even if a planted entry denies admins,
#   (2) protected DACL on the root, (3) children reset to inherit it, (4) delete every junction/symlink (the link,
#   never its target), (5) sweep links once more. Runs before anything is written there.
$script:EarlyLog = New-Object System.Collections.ArrayList
function Remove-Reparse([string]$Root) {
    $stack = New-Object System.Collections.Stack; $stack.Push($Root)
    while ($stack.Count -gt 0) {
        $d = [string]$stack.Pop()
        foreach ($e in @([IO.Directory]::GetFileSystemEntries($d))) {
            $a = [IO.File]::GetAttributes($e)
            if ($a -band [IO.FileAttributes]::ReparsePoint) {
                # Removes the link itself; RemoveDirectory/DeleteFile never touch a link's target.
                if ($a -band [IO.FileAttributes]::Directory) { [IO.Directory]::Delete($e, $false) } else { [IO.File]::Delete($e) }
                [void]$script:EarlyLog.Add("הוסר קישור לא צפוי מתיקיית הנתונים: $e")
            } elseif ($a -band [IO.FileAttributes]::Directory) { $stack.Push($e) }
        }
    }
}
function Protect-DataDir {
    if ((Test-Path -LiteralPath $DataDir) -and ([IO.File]::GetAttributes($DataDir) -band [IO.FileAttributes]::ReparsePoint)) {
        [IO.Directory]::Delete($DataDir, $false)
        [void]$script:EarlyLog.Add('תיקיית הנתונים עצמה הייתה קישור — הקישור הוסר ונוצרה תיקייה אמיתית')
    }
    New-Item -ItemType Directory -Force -Path $DataDir | Out-Null
    $ic = Join-Path $script:Sys32 'icacls.exe'
    $o = (& $ic $DataDir /setowner '*S-1-5-32-544' /T /C /L /Q 2>&1 | Out-String).Trim(); if ($o) { [void]$script:EarlyLog.Add("icacls setowner: $o") }
    $acl = New-Object System.Security.AccessControl.DirectorySecurity
    $acl.SetAccessRuleProtection($true, $false)
    $inh = [System.Security.AccessControl.InheritanceFlags]'ContainerInherit, ObjectInherit'
    foreach ($sidText in 'S-1-5-32-544', 'S-1-5-18') {
        $sid = New-Object System.Security.Principal.SecurityIdentifier($sidText)
        $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($sid, 'FullControl', $inh, 'None', 'Allow')))
    }
    $acl.SetOwner((New-Object System.Security.Principal.SecurityIdentifier('S-1-5-32-544')))
    Set-Acl -LiteralPath $DataDir -AclObject $acl -ErrorAction Stop
    if (@([IO.Directory]::GetFileSystemEntries($DataDir)).Count -gt 0) {
        $o = (& $ic (Join-Path $DataDir '*') /reset /T /C /L /Q 2>&1 | Out-String).Trim(); if ($o) { [void]$script:EarlyLog.Add("icacls reset: $o") }
    }
    Remove-Reparse $DataDir
    Remove-Reparse $DataDir
}

try { Protect-DataDir } catch { $script:AclError = $_.Exception.Message; New-Item -ItemType Directory -Force -Path $DataDir | Out-Null }
Remove-Item -LiteralPath (Join-Path $DataDir 'engine.lock') -Force -ErrorAction SilentlyContinue   # v1.0/1.1 lock file, replaced by the mutex
Remove-Item -LiteralPath (Join-Path $DataDir 'cancel.flag') -Force -ErrorAction SilentlyContinue

Import-Settings
Initialize-Run
foreach ($l in $script:EarlyLog) { Write-Ev 3 'system' $l 'warn' }
if ($script:AclError) { Write-Ev 1 'system' ('הרשאות תיקיית הנתונים לא הוגבלו: ' + $script:AclError) 'err' }
else {
    # Safe now: no links inside and nobody but admins can create new ones.
    Get-ChildItem -LiteralPath $script:WorkDir -Force -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    # Elevated installers (winget's download cache, self-extracting setups, choco scripts, Add-Type's compiler) use
    # TEMP: point it at an admin-only folder instead of %LOCALAPPDATA%\Temp, which unelevated processes can write to.
    $script:TmpDir = Join-Path $DataDir 'tmp'
    Remove-Item -LiteralPath $script:TmpDir -Recurse -Force -ErrorAction SilentlyContinue
    New-Item -ItemType Directory -Force -Path $script:TmpDir | Out-Null
    $env:TEMP = $script:TmpDir; $env:TMP = $script:TmpDir
}
# Leftovers of a run that was killed (power loss, task time limit): one-shot user tasks and their output files.
Get-ScheduledTask -TaskPath '\UpdateHub\' -ErrorAction SilentlyContinue | Where-Object { $_.TaskName -like 'UpdateHub-user-*' } |
    Unregister-ScheduledTask -Confirm:$false -ErrorAction SilentlyContinue
Remove-Item -Path (Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'UpdateHub\run\*') -Force -ErrorAction SilentlyContinue

# Keep the machine awake for the whole run (a laptop idling to sleep mid-download/install), restored in finally.
$script:Awake = $false
try {
    Add-Type -Namespace UpdateHub -Name Power -MemberDefinition '[System.Runtime.InteropServices.DllImport("kernel32.dll")] public static extern uint SetThreadExecutionState(uint esFlags);' -ErrorAction Stop
    $script:Awake = ([UpdateHub.Power]::SetThreadExecutionState([uint32]2147483649) -ne 0)   # ES_CONTINUOUS | ES_SYSTEM_REQUIRED
} catch { }

# Metered connection (phone hotspot, capped plan): a full run can pull several GB.
function Test-Metered {
    try {
        [void][Windows.Networking.Connectivity.NetworkInformation, Windows.Networking.Connectivity, ContentType = WindowsRuntime]
        $prof = [Windows.Networking.Connectivity.NetworkInformation]::GetInternetConnectionProfile()
        if (-not $prof) { return $false }
        $cost = $prof.GetConnectionCost()
        return (([string]$cost.NetworkCostType -in 'Fixed', 'Variable') -or $cost.Roaming -or $cost.OverDataLimit)
    } catch { return $false }
}

New-Prov 'system'        'מלאי ומערכת'                       1
New-Prov 'windowsupdate' 'Windows Update'                     3
New-Prov 'drivers'       'דרייברים וקושחה · קטלוג Microsoft'   4
New-Prov 'nvidia'        'כרטיס מסך NVIDIA'                   2
New-Prov 'office'        'Microsoft 365'                      2
New-Prov 'winget'        'תוכנות · winget'                    3
New-Prov 'store'         'Microsoft Store'                    1
New-Prov 'powershell'    'מודולי PowerShell'                  1
New-Prov 'pip'           'Python · pip'                       1
New-Prov 'vscode'        'תוספי VS Code'                      1
New-Prov 'wsl'           'WSL'                                1
New-Prov 'pkgmgr'        'Chocolatey · Scoop · npm'           1
New-Prov 'defender'      'Windows Defender'                   1
New-Prov 'health'        'תקינות התקנים'                      1
Save-State -Force

$modeText = 'עדכון מלא'; if ($Mode -eq 'Scan') { $modeText = 'סריקה בלבד' }
Write-Ev 1 'system' ("UpdateHub · $modeText · ריצה $($script:RunId)") 'step'
$final = 'completed'
try {
    if (-not (Test-IsAdmin)) { throw 'נדרשות הרשאות מנהל (הפעל דרך קיצור הדרך UpdateHub)' }
    # Fail closed: running with defaults would silently drop the user's exclusions and firmware limits.
    if ($script:SettingsError) { throw ('settings.json לא תקין, הריצה לא התחילה: ' + $script:SettingsError) }
    # Never execute downloaded installers from a folder that standard users might be able to write to.
    if ($script:AclError -and $Mode -eq 'Update') { throw 'לא ניתן לנעול את תיקיית הנתונים — העדכון לא התחיל (סריקה עדיין אפשרית)' }
    if ($Mode -eq 'Update') {
        # A cumulative update or a GPU driver needs several GB; failing halfway through is worse than not starting.
        $freeGb = [math]::Round((New-Object IO.DriveInfo($script:SysDrive)).AvailableFreeSpace / 1GB, 1)
        $minGb = [double](Get-Setting 'minFreeGB')
        if ($freeGb -lt $minGb) { throw ("יש רק $freeGb GB פנויים בכונן $($script:SysDrive) (נדרשים לפחות $minGb) — העדכון לא התחיל") }
        if (Test-Metered) {
            if ($Scheduled -and -not (Get-Setting 'allowMetered')) {
                $script:SkipRun = 'החיבור לרשת מוגדר כמוגבל (metered) — הריצה המתוזמנת דולגה (allowMetered בהגדרות)'
                throw $script:SkipRun
            }
            Write-Ev 1 'system' 'החיבור לרשת מוגדר כמוגבל (metered) — העדכון עלול להוריד כמה GB' 'warn'
        }
    }
    if (-not $script:Awake) { Write-Ev 3 'system' 'לא ניתן למנוע שינה בזמן הריצה' 'warn' }
    Invoke-Provider 'system'        { param($p) Run-System $p }
    Invoke-Provider 'windowsupdate' { param($p) Run-WindowsUpdate $p }
    Invoke-Provider 'drivers'       { param($p) Run-Drivers $p }
    Invoke-Provider 'nvidia'        { param($p) Run-Nvidia $p }
    Invoke-Provider 'office'        { param($p) Run-Office $p }
    Invoke-Provider 'winget'        { param($p) Run-Winget $p }
    Invoke-Provider 'store'         { param($p) Run-Store $p }
    Invoke-Provider 'powershell'    { param($p) Run-PowerShell $p }
    Invoke-Provider 'pip'           { param($p) Run-Pip $p }
    Invoke-Provider 'vscode'        { param($p) Run-VSCode $p }
    Invoke-Provider 'wsl'           { param($p) Run-Wsl $p }
    Invoke-Provider 'pkgmgr'        { param($p) Run-PkgMgr $p }
    Invoke-Provider 'defender'      { param($p) Run-Defender $p }
    Invoke-Provider 'health'        { param($p) Run-Health $p }
}
catch [System.OperationCanceledException] { $final = 'cancelled'; Write-Ev 1 'system' 'הריצה בוטלה' 'warn' }
catch {
  if ($script:SkipRun) {
    # A deliberate skip (metered network) is neutral, not a failure.
    $final = 'skipped'; $script:State.error = $script:SkipRun
    Write-Ev 1 'system' $script:SkipRun 'warn'
    try { $e = Add-Entry $script:Prov['system'] 'הריצה דולגה' 'UpdateHub' ''; Set-Entry $e 'skipped' $script:SkipRun } catch { }
  } else {
    $final = 'failed'; $script:State.error = $_.Exception.Message
    Write-Ev 1 'system' ('הריצה נעצרה: ' + $_.Exception.Message) 'err'
    # Visible as a failed row too, so the tiles / toast / history never say "everything is up to date".
    try { $e = Add-Entry $script:Prov['system'] 'הריצה נעצרה' 'UpdateHub' ''; Set-Entry $e 'failed' $_.Exception.Message } catch { }
  }
}
finally {
    Restore-Gallery
    if ($script:TmpDir) { Remove-Item -Path (Join-Path $script:TmpDir '*') -Recurse -Force -ErrorAction SilentlyContinue }
    if ($script:Awake) { try { [void][UpdateHub.Power]::SetThreadExecutionState([uint32]2147483648) } catch { } }   # ES_CONTINUOUS only
    foreach ($p in $script:State.providers) {
        if ($p.status -in 'pending', 'running') { $p.status = 'skipped'; if (-not $p.summary) { $p.summary = 'לא הורץ' } }
        # No item may stay "in progress" in a finished run (cancel / fatal error mid-item).
        foreach ($it in $p.items) {
            if ($it.status -eq 'checking') { $it.status = 'skipped'; $it.reason = 'לא נבדק — הריצה הופסקה' }
            elseif ($it.status -in 'downloading', 'installing') { $it.status = 'failed'; $it.reason = 'נקטע באמצע — ייבדק שוב בריצה הבאה' }
        }
    }
    $script:State.status = $final
    $script:State.finishedAt = (Get-Date).ToString('s')
    $script:State.current = ''
    Save-State -Force
    $c = $script:State.counts
    $n = { param($k) if ($c.Contains($k)) { $c[$k] } else { 0 } }
    $sum = ("עודכנו {0} · ממתינים לאתחול {1} · נכשלו {2} · ידני {3}" -f (& $n 'updated'), (& $n 'reboot'), (& $n 'failed'), (& $n 'manual'))
    Write-Ev 1 'system' ("סיום: $sum") $(if ($final -eq 'completed') { 'ok' } else { 'warn' })
    $summary = [ordered]@{ runId = $script:RunId; mode = $Mode; scheduled = [bool]$Scheduled; status = $final
                          startedAt = $script:State.startedAt; finishedAt = $script:State.finishedAt; counts = $c; rebootRequired = $script:State.rebootRequired
                          error = [string]$script:State.error }
    try { [IO.File]::WriteAllText((Join-Path $script:RunDir 'summary.json'), ($summary | ConvertTo-Json -Compress), (New-Object System.Text.UTF8Encoding($false))) } catch { }
    if ($script:EvWriter) { $script:EvWriter.Dispose() }
    # keep the last N runs
    $keep = [int](Get-Setting 'keepRuns')
    Get-ChildItem (Join-Path $DataDir 'runs') -Directory -ErrorAction SilentlyContinue | Sort-Object Name -Descending | Select-Object -Skip $keep |
        Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item (Join-Path $DataDir 'cancel.flag') -Force -ErrorAction SilentlyContinue
    $rb = ''; if ($script:State.rebootRequired) { $rb = ' · נדרש אתחול' }
    Show-Toast 'UpdateHub — הריצה הסתיימה' ($sum + $rb)
    if ($final -eq 'completed' -and $Scheduled -and $script:State.rebootRequired -and (Get-Setting 'autoRebootScheduled')) {
        & (Join-Path $script:Sys32 'shutdown.exe') /r /t 300 /c "UpdateHub: אתחול לסיום עדכונים בעוד 5 דקות (shutdown /a לביטול)"
    }
    try { $script:Mutex.ReleaseMutex() } catch { }; $script:Mutex.Dispose()
}
