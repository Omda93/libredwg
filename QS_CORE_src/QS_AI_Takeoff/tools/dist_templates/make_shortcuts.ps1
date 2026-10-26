param(
    [Parameter(Mandatory = $true)][string]$Target,
    [string]$DesktopDir = [Environment]::GetFolderPath('Desktop'),
    [string]$MenuDir = (Join-Path ([Environment]::GetFolderPath('Programs')) 'QS CORE')
)
$ErrorActionPreference = 'Stop'
$pyw = Join-Path $Target '.venv\Scripts\pythonw.exe'
if (-not (Test-Path -LiteralPath $pyw)) { throw "pythonw.exe not found: $pyw" }
$shell = New-Object -ComObject WScript.Shell
New-Item -ItemType Directory -Force -Path $MenuDir, $DesktopDir | Out-Null

function New-Link([string]$Path, [string]$Script, [string]$WorkDir, [string]$Desc) {
    $s = $shell.CreateShortcut($Path)
    $s.TargetPath = $pyw
    $s.Arguments = '"' + $Script + '"'
    $s.WorkingDirectory = $WorkDir
    $s.Description = $Desc
    $ico = Join-Path $Target 'QS_AI_Takeoff\config\qs_core.ico'
    if (Test-Path -LiteralPath $ico) { $s.IconLocation = $ico + ',0' }
    else { $s.IconLocation = (Join-Path $Target '.venv\Scripts\python.exe') + ',0' }
    $s.Save()
}

$v6 = Join-Path $Target 'DWG_Measure_Extractor_V6.py'
$qs = Join-Path $Target 'QS_AI_Takeoff\main.py'
$qsDir = Join-Path $Target 'QS_AI_Takeoff'
foreach ($dir in @($DesktopDir, $MenuDir)) {
    Remove-Item -LiteralPath (Join-Path $dir 'DWG Measure Extractor V6.lnk'), (Join-Path $dir 'Plaster and Paint Takeoff.lnk') -ErrorAction SilentlyContinue
    New-Link (Join-Path $dir 'QS CORE.lnk') $v6 $Target 'QS CORE - the all-in-one QS platform'
    New-Link (Join-Path $dir 'QS CORE Walls and Rooms.lnk') $qs $qsDir 'QS CORE - walls and rooms takeoff (plaster, paint, floor, ceiling)'
}
$u = $shell.CreateShortcut((Join-Path $MenuDir 'Uninstall QS CORE.lnk'))
$u.TargetPath = Join-Path $Target 'Uninstall.bat'
$u.WorkingDirectory = $Target
$u.Save()
Write-Output "shortcuts created in $DesktopDir and $MenuDir"
