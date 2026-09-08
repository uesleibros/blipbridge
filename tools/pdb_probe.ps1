param([Parameter(Mandatory)][string]$Module,[switch]$Download)
$ErrorActionPreference='Stop'
$data=[IO.File]::ReadAllBytes((Resolve-Path -LiteralPath $Module).Path)
$pe=[BitConverter]::ToInt32($data,0x3c)
$sections=[BitConverter]::ToUInt16($data,$pe+6)
$optionalSize=[BitConverter]::ToUInt16($data,$pe+20)
$optional=$pe+24
$magic=[BitConverter]::ToUInt16($data,$optional)
$dirs=if($magic -eq 0x20b){$optional+112}else{$optional+96}
function RvaOffset([uint32]$rva){
 for($i=0;$i -lt $sections;$i++){
  $s=$optional+$optionalSize+40*$i
  $va=[BitConverter]::ToUInt32($data,$s+12);$size=[BitConverter]::ToUInt32($data,$s+16);$raw=[BitConverter]::ToUInt32($data,$s+20)
  if($rva -ge $va -and $rva -lt $va+$size){return [int]($raw+$rva-$va)}
 }
 throw 'RVA not in a file section'
}
$debug=RvaOffset ([BitConverter]::ToUInt32($data,$dirs+6*8))
$debugSize=[BitConverter]::ToUInt32($data,$dirs+6*8+4)
for($d=$debug;$d -lt $debug+$debugSize;$d+=28){
 if([BitConverter]::ToUInt32($data,$d+12) -ne 2){continue}
 $raw=[BitConverter]::ToUInt32($data,$d+24)
 if([Text.Encoding]::ASCII.GetString($data,$raw,4) -ne 'RSDS'){continue}
 [byte[]]$guidBytes=$data[($raw+4)..($raw+19)]
 $guid=[Guid]::new($guidBytes)
 $age=[BitConverter]::ToUInt32($data,$raw+20)
 $end=$raw+24;while($data[$end]){$end++}
 $name=[IO.Path]::GetFileName([Text.Encoding]::ASCII.GetString($data,$raw+24,$end-$raw-24))
 $key=$guid.ToString('N').ToUpper()+$age.ToString('X')
 $url="https://msdl.microsoft.com/download/symbols/$name/$key/$name"
 "Module: $Module";"PDB: $name";"Key: $key";"URL: $url"
 if($Download){
  $dest="$PSScriptRoot/../artifacts/symbols/$name/$key";New-Item -ItemType Directory -Force $dest | Out-Null
  try{Invoke-WebRequest -UseBasicParsing -Uri $url -OutFile "$dest/$name" -TimeoutSec 45;"Downloaded: $dest/$name"}catch{"Download failed: $($_.Exception.Message)"}
 }
}
