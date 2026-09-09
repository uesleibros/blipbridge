$ErrorActionPreference='Stop'
$app=New-Object -ComObject PowerPoint.Application
$app.COMAddIns.Update();$addin=$app.COMAddIns.Item('BlipBridge.Engine');$addin.Connect=$true;$engine=$addin.Object
$root=Split-Path $PSScriptRoot -Parent
$pres=$app.Presentations.Add(0)
$results=@()
try{
 $slide=$pres.Slides.Add(1,12)
 $donor=$slide.Shapes.AddShape(1,10,10,100,100)
 $target=$slide.Shapes.AddShape(1,140,10,100,100)
 [byte[]]$bytes=[IO.File]::ReadAllBytes("$root/artifacts/textures/texture_64_0.png")
 $engine.MemoryFillExperiment($donor,$bytes,$root)
 $handle=$engine.RegisterTextureShape($donor)
 if($engine.GetTextureCount() -ne 1){throw 'Handle count'}
 $id=$target.Id
 $engine.ApplyTexture($target,$handle) | Out-Null
 if($target.Id -ne $id -or [int]$target.Type -ne 1 -or [int]$target.Fill.Type -ne 6){throw 'Target identity/fill failed'}
 $results+='Registered RAM-materialized donor and applied via Automation handle.'
 $engine.ReleaseTexture($handle)
 $rejected=$false;try{$engine.ApplyTexture($target,$handle)}catch{$rejected=$true}
 if(!$rejected){throw 'Stale handle accepted'}
 $results+='Stale handle rejected after ReleaseTexture.'
 $second=$engine.RegisterTextureShape($donor)
 if($second -eq $handle){throw 'Handle reused'}
 $engine.ClearTextures()
 if($engine.GetTextureCount() -ne 0){throw 'Clear count'}
 $results+='Handles are not recycled; ClearTextures releases registered entries.'
 # LoadTexture now decodes bytes into a reusable native texture. Capabilities
 # still report the backend as unvalidated, so this checks the API contract only:
 # a handle is issued, it is disjoint from donor handles, and releasing works.
 $native=$engine.LoadTexture($bytes)
 if($native -lt 0x1000000){throw 'Native texture handle overlaps the donor handle space'}
 if($engine.GetTextureCount() -ne 1){throw 'Native texture was not counted'}
 $engine.ReleaseTexture($native)
 if($engine.GetTextureCount() -ne 0){throw 'ReleaseTexture left the native texture behind'}
 $stale=$false;try{$engine.ApplyTexture($target,$native)}catch{$stale=$true}
 if(!$stale){throw 'Released native handle was still accepted'}
 $results+='LoadTexture issues a native handle in a disjoint range; release and stale checks hold.'
 $rejected=$false;try{$engine.SetImageBytes($target,$bytes)}catch{$rejected=$true}
 if(!$rejected){throw 'SetImageBytes claimed success'}
 $results+='SetImageBytes still explicitly rejects the unsupported backend.'
 $handle=$engine.RegisterTextureShape($donor)
 $donor.Delete()
 $rejected=$false;try{$engine.ApplyTexture($target,$handle)}catch{$rejected=$true}
 if(!$rejected){throw 'Deleted donor unexpectedly accepted'}
 $engine.ClearTextures()
 $results+='Deleted donor returns COM error; PowerPoint remains usable.'
 [byte[]]$invalid=@(0,1,2,3,4,5)
 $rejected=$false;try{$engine.MemoryFillExperiment($target,$invalid,$root)}catch{$rejected=$true}
 if(!$rejected){throw 'Invalid image accepted'}
 $target.Fill.UserPicture("$root/artifacts/textures/texture_64_0.png")
 $results+='Invalid image rejected; subsequent ordinary UserPicture succeeds after hook cleanup.'
 $results+='Host PID='+$engine.GetHostProcessId()
 $results | Set-Content "$root/artifacts/com_smoke.txt"
 $results
}finally{$engine.ClearTextures();$pres.Saved=-1;$pres.Close()}
