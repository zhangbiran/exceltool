param(
    [Parameter(Mandatory = $true)]
    [string]$File,
    [string]$Sheet = "常量表",
    [string]$Range = "A1:C5"
)

$ErrorActionPreference = "Stop"
$resolvedFile = (Resolve-Path -LiteralPath $File).Path
$instanceFile = Join-Path $env:LOCALAPPDATA "ExcelTool\runtime\daemon.json"

function Invoke-ExcelToolJson {
    param([string[]]$Arguments)
    $output = & exceltool @Arguments 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw "exceltool $($Arguments -join ' ') failed ($LASTEXITCODE): $output"
    }
    return ($output -join "`n") | ConvertFrom-Json
}

function Invoke-TimedView {
    $watch = [Diagnostics.Stopwatch]::StartNew()
    $result = Invoke-ExcelToolJson -Arguments @(
        "view", "--file", $resolvedFile, "--sheet", $Sheet,
        "--range", $Range, "--json-full"
    )
    $watch.Stop()
    return [pscustomobject]@{
        elapsed_seconds = [Math]::Round($watch.Elapsed.TotalSeconds, 3)
        result = $result
    }
}

function Get-OwnedProcessIds {
    param([int]$RootPid)
    $all = @(Get-CimInstance Win32_Process)
    $pending = @($RootPid)
    $owned = New-Object System.Collections.Generic.HashSet[int]
    while ($pending.Count -gt 0) {
        $parent = [int]$pending[0]
        $pending = @($pending | Select-Object -Skip 1)
        if ($owned.Add($parent)) {
            $children = @($all | Where-Object { $_.ParentProcessId -eq $parent })
            $pending += @($children | ForEach-Object { [int]$_.ProcessId })
        }
    }
    return @($owned | ForEach-Object { $_ })
}

function Wait-OwnedProcessesExit {
    param(
        [int[]]$ProcessIds,
        [string]$Context
    )
    $deadline = [DateTime]::UtcNow.AddSeconds(10)
    do {
        $survivors = @(
            $ProcessIds | Where-Object {
                Get-Process -Id $_ -ErrorAction SilentlyContinue
            }
        )
        if ($survivors.Count -eq 0) { return }
        Start-Sleep -Milliseconds 100
    } while ([DateTime]::UtcNow -lt $deadline)
    throw "$Context did not clean owned LibreOffice PIDs: $($survivors -join ',')"
}

Invoke-ExcelToolJson -Arguments @("daemon", "stop", "--json") | Out-Null

$first = Invoke-TimedView
$status1 = Invoke-ExcelToolJson -Arguments @("daemon", "status", "--json")
$second = Invoke-TimedView
$status2 = Invoke-ExcelToolJson -Arguments @("daemon", "status", "--json")

if (-not $status1.running -or -not $status2.running) {
    throw "daemon did not remain running across views"
}
if ($status1.libreoffice.pid -ne $status2.libreoffice.pid) {
    throw "LibreOffice PID was not reused"
}
if ($status1.libreoffice.generation -ne $status2.libreoffice.generation) {
    throw "LibreOffice generation was not reused"
}

$exclusive = [IO.File]::Open(
    $resolvedFile,
    [IO.FileMode]::Open,
    [IO.FileAccess]::Read,
    [IO.FileShare]::None
)
$exclusive.Dispose()

$cleanOwnedProcessIds = Get-OwnedProcessIds -RootPid ([int]$status2.libreoffice.pid)
Invoke-ExcelToolJson -Arguments @("daemon", "stop", "--json") | Out-Null
Wait-OwnedProcessesExit -ProcessIds $cleanOwnedProcessIds -Context "daemon stop"

Invoke-TimedView | Out-Null
$crashStatus = Invoke-ExcelToolJson -Arguments @("daemon", "status", "--json")
$instance = Get-Content -LiteralPath $instanceFile -Raw -Encoding utf8 | ConvertFrom-Json
$ownedProcessIds = Get-OwnedProcessIds -RootPid ([int]$crashStatus.libreoffice.pid)
Stop-Process -Id ([int]$instance.pid) -Force
Wait-OwnedProcessesExit -ProcessIds $ownedProcessIds -Context "Job Object"

# Replace the stale crash record with a fresh daemon, then stop it cleanly.
Invoke-ExcelToolJson -Arguments @("daemon", "start", "--json") | Out-Null
Invoke-ExcelToolJson -Arguments @("daemon", "stop", "--json") | Out-Null

[pscustomobject]@{
    ok = $true
    file = $resolvedFile
    sheet = $Sheet
    range = $Range
    first_view_seconds = $first.elapsed_seconds
    second_view_seconds = $second.elapsed_seconds
    reused_pid = [int]$status2.libreoffice.pid
    reused_generation = [int]$status2.libreoffice.generation
    workbook_exclusive_open = $true
    clean_stop = $true
    crash_cleanup = $true
} | ConvertTo-Json -Depth 4
