$ErrorActionPreference='Stop'
Add-Type -AssemblyName System.Drawing
$dir=Join-Path (Split-Path $PSScriptRoot -Parent) 'artifacts/textures'
New-Item -ItemType Directory -Force $dir | Out-Null
foreach($size in 32,64,128,256){
  foreach($variant in 0,1){
    $b=[Drawing.Bitmap]::new($size,$size)
    try {
      for($y=0;$y -lt $size;$y++){for($x=0;$x -lt $size;$x++){
        $r=($x*7+$variant*97)%256; $g=($y*11+$variant*43)%256; $blue=(([int]($x/8) -bxor [int]($y/8))*63)%256
        $b.SetPixel($x,$y,[Drawing.Color]::FromArgb(255,$r,$g,$blue))
      }}
      $b.Save("$dir/texture_${size}_${variant}.png",[Drawing.Imaging.ImageFormat]::Png)
      if($size -eq 64){$b.Save("$dir/texture_${size}_${variant}.jpg",[Drawing.Imaging.ImageFormat]::Jpeg)}
    } finally {$b.Dispose()}
  }
}
Copy-Item "$dir/texture_64_0.png" "$dir/BB_TRACE_UNIQUE_TEXTURE_01.png" -Force
Copy-Item "$dir/texture_64_0.png" "$dir/identical_bytes_other_filename.png" -Force
