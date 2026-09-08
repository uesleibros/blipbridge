$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$app=New-Object -ComObject PowerPoint.Application
$app.Visible=-1
$app.COMAddIns.Update()
$addin=$app.COMAddIns.Item('BlipBridge.Engine'); $addin.Connect=$true
$engine=$addin.Object
$pres=$app.Presentations.Add(-1)
try {
 $slide=$pres.Slides.Add(1,12)
 $shape=$slide.Shapes.AddShape(1,100,100,100,100)
 $engine.TraceUserPicture($shape.Fill,$root)
 $target=$slide.Shapes.AddShape(1,230,100,100,100)
 $engine.TraceCachedApply($shape,$target,$root)
 $shape.Export("$root/artifacts/trace_shape.png",2)
 'Trace complete; fill type='+$shape.Fill.Type
}finally{$pres.Saved=-1;$pres.Close()}
