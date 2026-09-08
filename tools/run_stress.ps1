$ErrorActionPreference='Stop'
$app=New-Object -ComObject PowerPoint.Application
$app.COMAddIns.Update();$addin=$app.COMAddIns.Item('BlipBridge.Engine');$addin.Connect=$true
$engine=$addin.Object
$engine.RunStress($app,(Split-Path $PSScriptRoot -Parent))
'Stress complete'
