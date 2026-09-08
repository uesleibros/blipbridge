param([ValidateSet('Release','Debug')][string]$Configuration='Release',[switch]$ResearchAddin)
$ErrorActionPreference='Stop'
if(![Environment]::Is64BitProcess){throw 'Run 64-bit PowerShell'}
$dll=(Resolve-Path "$PSScriptRoot/../build/$Configuration/BlipBridge.dll").Path
$id='{2E2E2731-C523-486B-89CB-2A89484F1E32}'
$base='HKCU:\Software\Classes'
New-Item "$base/CLSID/$id/InprocServer32" -Force | Out-Null
Set-Item "$base/CLSID/$id/InprocServer32" -Value $dll
New-ItemProperty "$base/CLSID/$id/InprocServer32" -Name ThreadingModel -Value Apartment -Force | Out-Null
New-Item "$base/BlipBridge.Engine/CLSID" -Force | Out-Null
Set-Item "$base/BlipBridge.Engine/CLSID" -Value $id
New-Item "$base/CLSID/$id/ProgID" -Force | Out-Null
Set-Item "$base/CLSID/$id/ProgID" -Value BlipBridge.Engine
if($ResearchAddin){
 $key='HKCU:\Software\Microsoft\Office\PowerPoint\Addins\BlipBridge.Engine'
 New-Item $key -Force | Out-Null
 New-ItemProperty $key -Name FriendlyName -Value 'BlipBridge research harness' -Force | Out-Null
 New-ItemProperty $key -Name Description -Value 'Explicitly connected native benchmark harness' -Force | Out-Null
 New-ItemProperty $key -Name LoadBehavior -PropertyType DWord -Value 0 -Force | Out-Null
}
Write-Output "Registered per-user x64: $dll"
