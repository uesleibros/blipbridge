param([string]$Path="$PSScriptRoot/../artifacts/baseline.pptx",[string]$Report='package_inspection.md')
$ErrorActionPreference='Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip=[IO.Compression.ZipFile]::OpenRead((Resolve-Path $Path).Path)
try {
 $media=@($zip.Entries | Where-Object {$_.FullName -like 'ppt/media/*'})
 $lines=@('# Saved package inspection','',"File: $([IO.Path]::GetFileName($Path))", "Media count: $($media.Count)",'','| Entry | Bytes | SHA256 |','|---|---|---|')
 foreach($e in $media){$s=$e.Open();try{$sha=[Security.Cryptography.SHA256]::Create();$hash=[BitConverter]::ToString($sha.ComputeHash($s)).Replace('-','');$lines+="| $($e.FullName) | $($e.Length) | $hash |"}finally{$s.Dispose();$sha.Dispose()}}
 foreach($e in $zip.Entries | Where-Object {$_.FullName -like 'ppt/slides/slide*.xml'}){
  $r=[IO.StreamReader]::new($e.Open());try{[xml]$x=$r.ReadToEnd()}finally{$r.Dispose()}
  $n=[Xml.XmlNamespaceManager]::new($x.NameTable);$n.AddNamespace('p','http://schemas.openxmlformats.org/presentationml/2006/main');$n.AddNamespace('a','http://schemas.openxmlformats.org/drawingml/2006/main');$n.AddNamespace('r','http://schemas.openxmlformats.org/officeDocument/2006/relationships')
  $shapes=$x.SelectNodes('//p:sp',$n);$pictures=$x.SelectNodes('//p:pic',$n);$fills=$x.SelectNodes('//p:sp/p:spPr/a:blipFill',$n)
  $lines+=@('',"$($e.FullName): normal shapes=$($shapes.Count), picture shapes=$($pictures.Count), picture-filled normal shapes=$($fills.Count)")
  $groups=$fills | ForEach-Object {$_.SelectSingleNode('a:blip',$n).GetAttribute('embed','http://schemas.openxmlformats.org/officeDocument/2006/relationships')} | Group-Object
  foreach($g in $groups){$lines+="- $($g.Name): $($g.Count) normal shape fills"}
 }
 $lines | Set-Content "$PSScriptRoot/../docs/$Report"
 $lines
}finally{$zip.Dispose()}
