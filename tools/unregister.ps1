$ErrorActionPreference='Stop'
foreach($key in 'HKCU:\Software\Classes\CLSID\{2E2E2731-C523-486B-89CB-2A89484F1E32}','HKCU:\Software\Classes\BlipBridge.Engine','HKCU:\Software\Microsoft\Office\PowerPoint\Addins\BlipBridge.Engine'){
 if(Test-Path -LiteralPath $key){Remove-Item -LiteralPath $key -Recurse -Force}
}
