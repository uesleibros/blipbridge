$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$app=New-Object -ComObject PowerPoint.Application
$app.Visible=-1
$app.COMAddIns.Update();$addin=$app.COMAddIns.Item('BlipBridge.Engine');$addin.Connect=$true;$engine=$addin.Object
$pres=$app.Presentations.Add(-1)
try{
 $slide=$pres.Slides.Add(1,12)
 $builder=$slide.Shapes.BuildFreeform(0,100,100)
 $builder.AddNodes(0,0,240,100);$builder.AddNodes(0,0,170,240);$builder.AddNodes(0,0,100,100)
 $shape=$builder.ConvertToShape();$shape.Name='BB_memory_original'
 $before=@($shape.Id,$shape.Type,$shape.Left,$shape.Top,$shape.Width,$shape.Height,$shape.Rotation,$shape.ZOrderPosition,$shape.Nodes.Count) -join ','
 $results=@()
 foreach($file in 'texture_32_0.png','texture_64_1.png','texture_128_0.png','texture_256_1.png','texture_64_0.jpg'){
  [byte[]]$bytes=[IO.File]::ReadAllBytes("$root/artifacts/textures/$file")
  $engine.MemoryFillExperiment($shape,$bytes,$root)
  $results+="RAM fill succeeded: $file ($($bytes.Length) bytes)"
 }
 [byte[]]$bytes=[IO.File]::ReadAllBytes("$root/artifacts/textures/texture_64_1.png")
 $engine.MemoryFillExperiment($shape,$bytes,$root)
 $after=@($shape.Id,$shape.Type,$shape.Left,$shape.Top,$shape.Width,$shape.Height,$shape.Rotation,$shape.ZOrderPosition,$shape.Nodes.Count) -join ','
 if($before -ne $after){throw 'Shape identity or geometry changed'}
 if([int]$shape.Fill.Type -ne 6){throw 'Not a picture fill'}
 $shape.Nodes.SetPosition(2,250,100)
 $reference=$shape.Duplicate().Item(1)
 $reference.Fill.UserPicture("$root/artifacts/textures/texture_64_1.png")
 $reference.Export("$root/artifacts/memory_reference.png",2)
 $reference.Delete()
 $shape.Export("$root/artifacts/memory_shape.png",2)
 $pres.SaveAs("$root/artifacts/memory.pptx",24)
 $results+="Memory experiment passed identity, geometry, picture fill, node editing, and SaveAs. PID=$($engine.GetHostProcessId())"
 $results | Set-Content "$root/artifacts/memory_functional.txt"
}finally{$pres.Saved=-1;$pres.Close()}
$reopen=$app.Presentations.Open("$root/artifacts/memory.pptx",0,0,0)
try{
 $shape=$reopen.Slides.Item(1).Shapes.Item('BB_memory_original')
 if([int]$shape.Type -ne 5 -or [int]$shape.Fill.Type -ne 6){throw 'Reopen failed'}
 $shape.Export("$root/artifacts/memory_reopened.png",2)
 'Reopen preserved editable freeform and picture fill.' | Add-Content "$root/artifacts/memory_functional.txt"
}finally{$reopen.Saved=-1;$reopen.Close()}
Add-Type -AssemblyName System.Drawing
$original=[Drawing.Bitmap]::new("$root/artifacts/memory_shape.png")
try{
 foreach($name in 'memory_reference.png','memory_reopened.png'){
  $other=[Drawing.Bitmap]::new("$root/artifacts/$name")
  try{
   if($original.Width -ne $other.Width -or $original.Height -ne $other.Height){throw "Image size mismatch: $name"}
   $diff=0;for($y=0;$y -lt $original.Height;$y++){for($x=0;$x -lt $original.Width;$x++){if($original.GetPixel($x,$y).ToArgb() -ne $other.GetPixel($x,$y).ToArgb()){$diff++}}}
   "Pixel differences against ${name}: $diff" | Add-Content "$root/artifacts/memory_functional.txt"
   if($diff){throw "Rendered image mismatch: $name"}
  }finally{$other.Dispose()}
 }
}finally{$original.Dispose()}
