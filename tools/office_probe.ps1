param([switch]$StartPowerPoint)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
New-Item -ItemType Directory -Force "$root/artifacts","$root/docs" | Out-Null
function Get-PeMachine([string]$Path) {
    if(!(Test-Path -LiteralPath $Path)) { return 'unavailable (Office virtual path)' }
    $s = [IO.File]::Open($Path, 'Open', 'Read', 'ReadWrite')
    try { $r = [IO.BinaryReader]::new($s); $s.Position=0x3c; $offset=$r.ReadInt32(); $s.Position=$offset+4; $m=$r.ReadUInt16(); if($m -eq 0x8664){'x64'}elseif($m -eq 0x14c){'x86'}else{'0x{0:X}' -f $m} } finally {$s.Dispose()}
}
$os = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
$office = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
$exe = Get-ChildItem $office.InstallationPath -Recurse -Filter POWERPNT.EXE | Select-Object -First 1
$compiler = Get-Command g++.exe -ErrorAction SilentlyContinue
$candidates = @($env:MSYS2_ROOT, 'C:\msys64', 'C:\msys2', 'D:\msys64') | Where-Object {$_}
$gxx = if($compiler){$compiler.Source}else{$candidates | ForEach-Object {Join-Path $_ 'ucrt64/bin/g++.exe'} | Where-Object {Test-Path $_} | Select-Object -First 1}
if(!$gxx){throw 'UCRT64 g++ not found'}
$bin = Split-Path $gxx
$cmake = Join-Path $bin cmake.exe
if(!(Test-Path $cmake)){$cmake=(Get-Command cmake.exe).Source}
$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio/Installer/vswhere.exe'
$vs = if(Test-Path $vswhere){ & $vswhere -all -format json | Out-String }else{'Not found'}
$sdk = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows Kits\Installed Roots' -ErrorAction SilentlyContinue
if($StartPowerPoint -and !(Get-Process POWERPNT -ErrorAction SilentlyContinue)) {
    $app = New-Object -ComObject PowerPoint.Application
    $app.Visible = -1
}
$modules = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | ForEach-Object {
    $processId=$_.Id
    $_.Modules | ForEach-Object {
        $physical=$_.FileName
        if(!(Test-Path -LiteralPath $physical)){
            $common=Join-Path $env:ProgramFiles 'Common Files'
            if($physical.StartsWith($common,[StringComparison]::OrdinalIgnoreCase)){
                $candidate=Join-Path $office.InstallationPath ('root/vfs/ProgramFilesCommonX64/'+$physical.Substring($common.Length).TrimStart('\'))
                if(Test-Path -LiteralPath $candidate){$physical=$candidate}
            }
        }
        $version=if(Test-Path -LiteralPath $physical){[Diagnostics.FileVersionInfo]::GetVersionInfo($physical).FileVersion}else{$_.FileVersionInfo.FileVersion}
        [pscustomobject]@{pid=$processId;filename=$_.ModuleName;path=$_.FileName;physicalPath=$physical;version=$version;base=('0x{0:X}' -f $_.BaseAddress.ToInt64());architecture=(Get-PeMachine $physical)}
    }
})
$modules | ConvertTo-Json -Depth 4 | Set-Content "$root/artifacts/modules.json"
$modules | Export-Csv "$root/artifacts/modules.csv" -NoTypeInformation
$report = @(
    '# Local environment', '', "Probed: $(Get-Date -Format o)", '',
    "Windows: $((Get-CimInstance Win32_OperatingSystem).Caption), $($os.DisplayVersion), build $($os.CurrentBuild).$($os.UBR) (registry ProductName: $($os.ProductName))",
    "Windows architecture: $env:PROCESSOR_ARCHITECTURE", '',
    "Office: $($office.ProductReleaseIds), Click-to-Run, $($office.VersionToReport), $($office.Platform)",
    "POWERPNT: $($exe.FullName)", "Executable version: $($exe.VersionInfo.FileVersion)", "PE architecture: $(Get-PeMachine $exe.FullName)", '',
    "MSYS2 root: $(Split-Path (Split-Path $bin))", "UCRT64: $(Split-Path $bin)",
    "g++: $gxx", (& $gxx --version | Out-String), "gcc: $(Join-Path $bin gcc.exe)", (& "$bin/gcc.exe" --version | Out-String),
    "Target: $(& $gxx -dumpmachine)", "Linker: $(& $gxx -print-prog-name=ld)", (& "$bin/ld.exe" --version | Out-String),
    "CMake: $cmake", (& $cmake --version | Out-String), "MinGW Windows headers: $(Test-Path "$bin/../include/windows.h")", "SDK root: $($sdk.KitsRoot10)",
    "SDK include versions: $((Get-ChildItem "$($sdk.KitsRoot10)/Include" -ErrorAction SilentlyContinue).Name -join ', ')",
    '', 'Visual Studio / Build Tools discovery:', '```json', $vs.Trim(), '```', '',
    'Loaded modules (all modules enumerated dynamically; addresses are process-specific):', '',
    '| PID | Filename | Version | Architecture | Base | Loaded path | Physical path |', '|---|---|---|---|---|---|---|'
)
$report += $modules | ForEach-Object {"| $($_.pid) | $($_.filename) | $($_.version) | $($_.architecture) | $($_.base) | $($_.path) | $($_.physicalPath) |"}
$report | Set-Content "$root/docs/environment.md"
[pscustomobject]@{gxx=$gxx;bin=$bin;cmake=$cmake;powerpoint=$exe.FullName;officeVersion=$office.VersionToReport;architecture=(Get-PeMachine $exe.FullName)} | ConvertTo-Json | Set-Content "$root/artifacts/environment.json"
Write-Output "Wrote docs/environment.md and artifacts/modules.{json,csv}; $($modules.Count) loaded modules."
