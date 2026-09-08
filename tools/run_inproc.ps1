$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
& "$PSScriptRoot/register.ps1" -ResearchAddin
$app=New-Object -ComObject PowerPoint.Application
$app.Visible=-1
$app.COMAddIns.Update()
$addin=$app.COMAddIns.Item('BlipBridge.Engine')
$addin.Connect=$true
$engine=$addin.Object
if(!$engine){throw 'Research add-in did not expose Engine'}
Write-Output "Engine host PID: $($engine.GetHostProcessId()); PowerPoint PID: $((Get-Process POWERPNT).Id)"
$result=$engine.RunBenchmarks($root)
if($result -ne 0){throw "Native experiment failed: $result"}
Write-Output 'In-process benchmarks complete'
