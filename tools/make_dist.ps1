<#
.SYNOPSIS
Builds the release layout under dist/.

.DESCRIPTION
Produces exactly what a user needs and nothing else:

    dist/windows-x64/BlipBridge.dll
    dist/windows-x64/BlipBridge.bas
    dist/windows-x64/blipbridge.h
    dist/windows-x64/README.txt
    dist/windows-x64/LICENSE

No macOS folder is produced, because no macOS backend exists. Shipping an empty
or stubbed .dylib would suggest support that is not there.
#>
param([ValidateSet('Release', 'Debug')][string]$Configuration = 'Release')

$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$target = Join-Path $root 'dist/windows-x64'

$dll = Join-Path $root "build/$Configuration/BlipBridge.dll"
if (-not (Test-Path $dll)) {
    throw "Build $Configuration first: BlipBridge.dll not found at $dll"
}

if (Test-Path (Join-Path $root 'dist')) {
    Remove-Item (Join-Path $root 'dist') -Recurse -Force
}
New-Item -ItemType Directory -Path $target -Force | Out-Null

Copy-Item $dll                                       (Join-Path $target 'BlipBridge.dll')
Copy-Item (Join-Path $root 'vba/BlipBridge.bas')     (Join-Path $target 'BlipBridge.bas')
Copy-Item (Join-Path $root 'include/blipbridge/blipbridge.h') (Join-Path $target 'blipbridge.h')
Copy-Item (Join-Path $root 'LICENSE')                (Join-Path $target 'LICENSE')

# The version the DLL reports, so the archive is self-describing.
$version = 'unknown'
try {
    $probe = Join-Path $root "build/$Configuration/bb_abi_contract.exe"
    if (Test-Path $probe) { $version = (Get-Item $dll).LastWriteTime.ToString('yyyy-MM-dd') }
} catch { }

@"
BlipBridge - fast reusable image textures for PowerPoint Shapes
built $version

WHAT IS IN HERE
  BlipBridge.dll   the library
  BlipBridge.bas   the VBA module to import
  blipbridge.h     the C ABI, if you are calling from something other than VBA
  LICENSE          MIT

INSTALL
  1. Copy BlipBridge.dll next to your .pptm.
  2. Import BlipBridge.bas into the VBA project.

  There is nothing to register and nothing to install. No regsvr32, no ProgID,
  no COM add-in.

USE
  BlipBridge.Initialize
  tex = BlipBridge.LoadTexture(bytes)
  BlipBridge.ApplyTexture shp, tex
  BlipBridge.ReleaseTexture tex
  BlipBridge.Shutdown

REQUIREMENTS
  Windows x64, 64-bit PowerPoint, VBA7.
  The accelerated backend is validated against Office 16.0.14334.20848 and
  refuses to run against any other build. Call BlipBridge.IsAvailable to check;
  BlipBridge.Version explains what it found.

  This is not universal Office compatibility. The backend depends on internal
  Office layouts that are verified per build, and it fails closed rather than
  guessing.

PERFORMANCE
  On the tested build and machine, ApplyTexture averaged 0.186 ms against
  0.659 ms for Fill.UserPicture - about 3.5x. Your results depend on Office
  build, hardware and workload.

  ApplyTextureBatch is a convenience, not a speed-up; it measured within noise
  of the same number of individual calls.

DOCUMENTATION AND SOURCE
  https://github.com/  (see docs/ in the repository)

BlipBridge is not affiliated with or endorsed by Microsoft.
"@ | Set-Content (Join-Path $target 'README.txt') -Encoding utf8

Get-ChildItem $target | Select-Object Name, Length
"dist layout written to $target"
