$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$app=New-Object -ComObject PowerPoint.Application
$app.COMAddIns.Update();$addin=$app.COMAddIns.Item('BlipBridge.Engine');$addin.Connect=$true;$engine=$addin.Object
$pres=$app.Presentations.Add(0)
try{
 $shape=$pres.Slides.Add(1,12).Shapes.AddShape(1,10,10,100,100)
 # Load decoder modules before taking the module snapshot in a fresh Office process.
 $shape.Fill.UserPicture((Join-Path $root 'artifacts/textures/texture_32_0.png'))
 # The target remains the same normal AutoShape throughout the research call.
 $shapeBefore = @(
  $shape.Id, $shape.Name, $shape.Type, $shape.Left, $shape.Top,
  $shape.Width, $shape.Height, $shape.Rotation, $shape.ZOrderPosition
 ) -join '|'
 $process=Get-Process -Id $engine.GetHostProcessId()
 $m=$process.Modules | Where-Object ModuleName -eq mso20win32client.dll
 $moduleList = @($process.Modules | ForEach-Object {
  @{
   name = $_.ModuleName
   base = $_.BaseAddress.ToInt64()
   size = $_.ModuleMemorySize
   version = $_.FileVersionInfo.FileVersion
  }
 })
 @{pid=$process.Id;readReturn=('0x{0:X}' -f ($m.BaseAddress.ToInt64()+0x7aede));sinkCall=('0x{0:X}' -f ($m.BaseAddress.ToInt64()+0x209fab));modules=$moduleList} | ConvertTo-Json -Depth 4 | Set-Content "$root/artifacts/decoder_target.json"
 'Ready for debugger trigger.'
 $trigger="$root/artifacts/decoder_go"
 $deadline=(Get-Date).AddSeconds(55)
 while(!(Test-Path -LiteralPath $trigger)){
  if((Get-Date) -gt $deadline){throw 'Debugger trigger timed out'}
  Start-Sleep -Milliseconds 200
 }
 Set-Content "$root/artifacts/decoder_phase.txt" 'UserPicture'
 $engine.TraceUserPicture($shape.Fill,$root)
 $shapeAfter = @(
  $shape.Id, $shape.Name, $shape.Type, $shape.Left, $shape.Top,
  $shape.Width, $shape.Height, $shape.Rotation, $shape.ZOrderPosition
 ) -join '|'
 if ($shapeAfter -ne $shapeBefore -or $shape.Fill.Type -ne 6) {
  throw 'Trace changed target identity/geometry or failed to apply picture fill'
 }
 'Trace preserved AutoShape identity, geometry, Z order and picture fill.' |
  Set-Content "$root/artifacts/decoder_shape_validation.txt"
 Set-Content "$root/artifacts/decoder_phase.txt" 'SaveAs'
 $pres.SaveAs((Join-Path $root 'artifacts/decoder_lifetime.pptx'), 24)
 Set-Content "$root/artifacts/decoder_phase.txt" 'UserPicture-and-SaveAs-complete'
 'Decoder trace call completed.'
} finally {
 Set-Content "$root/artifacts/decoder_phase.txt" 'Presentation.Close'
 $pres.Saved = -1
 $pres.Close()
 Set-Content "$root/artifacts/decoder_phase.txt" 'Presentation.Close-complete'
}
