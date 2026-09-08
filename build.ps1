param([ValidateSet('Release','Debug')][string]$Configuration='Release')
$ErrorActionPreference='Stop'
$root=$PSScriptRoot
if(!(Test-Path "$root/artifacts/environment.json")){ & "$root/tools/office_probe.ps1" }
$e=Get-Content "$root/artifacts/environment.json" -Raw | ConvertFrom-Json
$env:PATH="$($e.bin);$env:PATH"
& $e.cmake -S $root -B "$root/build/$Configuration" -G 'MinGW Makefiles' "-DCMAKE_BUILD_TYPE=$Configuration" "-DCMAKE_CXX_COMPILER=$($e.gxx)"
if($LASTEXITCODE){throw 'Configure failed'}
& $e.cmake --build "$root/build/$Configuration" --parallel 4
if($LASTEXITCODE){throw 'Build failed'}
