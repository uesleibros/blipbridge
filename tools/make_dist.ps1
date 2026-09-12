<#
.SYNOPSIS
Builds a release package for one architecture.

.DESCRIPTION
Produces exactly what a user needs and nothing else:

    dist/blipbridge-v<version>-windows-<arch>/BlipBridge-<arch>.dll
                                            /BlipBridge.bas
                                            /blipbridge.h
                                            /README.txt
                                            /LICENSE
    dist/blipbridge-v<version>-windows-<arch>.zip
    dist/blipbridge-v<version>-windows-<arch>.zip.sha256

The README inside the archive is written per architecture, because the two say
genuinely different things: both fill Shapes, but only x64 does it through the
accelerated backend. x86 must say plainly that it is the slower route rather
than leaving a user to discover it from a benchmark.

No macOS package is produced, because no macOS backend exists. Shipping an empty
or stubbed .dylib would suggest support that is not there.

.PARAMETER Architecture
x64 or x86. Must match the build under -BuildDir.

.PARAMETER Version
Release version for the archive name, e.g. 0.4.0. Defaults to the version the
header declares.
#>
param(
    [ValidateSet('x64', 'x86')][string]$Architecture = 'x64',
    [ValidateSet('Release', 'Debug')][string]$Configuration = 'Release',
    [string]$BuildDir = '',
    [string]$Version = ''
)

$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
if (-not $BuildDir) { $BuildDir = Join-Path $root "build/$Configuration" }

$dllName = "BlipBridge-$Architecture.dll"
# The build emits the architecture-suffixed name directly, so the package and a
# local build refer to the same file.
$dll = Join-Path $BuildDir $dllName
if (-not (Test-Path $dll)) {
    throw "$dllName not found in $BuildDir - build $Configuration for $Architecture first"
}

# The version the header declares is the single source of truth for the release
# number; guessing from a file timestamp produced meaningless archive names.
if (-not $Version) {
    $header = Get-Content (Join-Path $root 'include/blipbridge/blipbridge.h') -Raw
    if ($header -match 'BB_VERSION_MAJOR\s+(\d+)' ) { $major = $Matches[1] }
    if ($header -match 'BB_VERSION_MINOR\s+(\d+)' ) { $minor = $Matches[1] }
    if ($header -match 'BB_VERSION_PATCH\s+(\d+)' ) { $patch = $Matches[1] }
    if ($major -and $minor -and $patch) { $Version = "$major.$minor.$patch" } else { $Version = '0.0.0' }
}

if ($Version -notmatch '^\d+\.\d+\.\d+([-.][0-9A-Za-z.-]+)?$') {
    throw "Invalid package version: $Version. Expected a version such as 0.4.0 without the v prefix."
}

$name = "blipbridge-v$Version-windows-$Architecture"
$distRoot = Join-Path $root 'dist'
$target = Join-Path $distRoot $name
# Validate the resolved directory before recursively replacing an earlier package.
$target = [System.IO.Path]::GetFullPath($target)
$expectedParent = [System.IO.Path]::GetFullPath($distRoot)
if ((Split-Path $target -Parent) -ne $expectedParent) {
    throw "Package directory must be directly under $expectedParent"
}
if (Test-Path $target) { Remove-Item -LiteralPath $target -Recurse -Force }
New-Item -ItemType Directory -Path $target -Force | Out-Null

Copy-Item $dll (Join-Path $target $dllName)
Copy-Item (Join-Path $root 'vba/BlipBridge.bas') (Join-Path $target 'BlipBridge.bas')
Copy-Item (Join-Path $root 'include/blipbridge/blipbridge.h') (Join-Path $target 'blipbridge.h')
Copy-Item (Join-Path $root 'LICENSE') (Join-Path $target 'LICENSE')

$repo = 'https://github.com/uesleibros/blipbridge'

# The two architectures make different promises. Rather than one README hedged
# to cover both, each says exactly what is true of the binary beside it.
if ($Architecture -eq 'x64') {
    $status = @"
REQUIREMENTS
  Windows x64 and 64-bit PowerPoint, VBA7.

  This package contains the Windows x64 DLL. If your PowerPoint is 32-bit it
  will not load it, and BlipBridge.bas will tell you so in those words rather
  than letting you meet a loader error.

  To check: PowerPoint > File > Account > About PowerPoint. The first line ends
  with "64-bit" or "32-bit".

BACKEND
  Accelerated. Validated against Office 16.0.14334.20848 and refused on any
  other build, where picture fills fall back to Office's own route.

  Call BlipBridge.IsAvailable to check that BlipBridge works here at all, and
  BlipBridge.IsAccelerated to check that this is the fast path.
  BlipBridge.Version explains what it found.

  This is not universal Office compatibility. The accelerated backend depends on
  internal Office layouts verified per build, and it fails closed rather than
  guessing.

PERFORMANCE
  On the tested build and machine, ApplyTexture averaged 0.186 ms against
  0.659 ms for Fill.UserPicture - about 3.5x. UserPicture2 skips the work
  entirely when a Shape already carries that image. Your results depend on
  Office build, hardware and workload.

  ApplyTextureBatch is a convenience, not a speed-up; it measured within noise
  of the same number of individual calls.
"@
} else {
    $status = @"
REQUIREMENTS
  Windows x86 and 32-bit PowerPoint, VBA7.

  This package contains the Windows x86 DLL. If your PowerPoint is 64-bit it
  will not load it, and BlipBridge.bas will tell you so in those words.

  To check: PowerPoint > File > Account > About PowerPoint. The first line ends
  with "64-bit" or "32-bit".

BACKEND - READ THIS
  Portable, not accelerated. It works; it is not fast.

  Everything the API offers works here: textures, ApplyTexture,
  ApplyTextureRange, ApplyTextureIfChanged and its skip cache, ApplyPicture, and
  the whole image pipeline - crop, orient, resize and quad warp. Shapes get
  filled.

  It does that through Office's own Fill.UserPicture rather than through the
  accelerated path, so an apply costs what Office charges for one.
  BlipBridge.IsAccelerated returns False here and BlipBridge.IsAvailable returns
  True, which is the distinction: usable, not accelerated. Do not quote the x64
  benchmark numbers for this package.

  The accelerated backend is a reconstruction of one 64-bit Office build's
  internals. It depends on the x64 calling convention, on per-build module
  addresses, and on object layouts 32-bit Office does not share, so it is not
  built into this binary at all - shipping something that compiles and might
  corrupt a document would be worse than shipping the slower route.

  One thing this package has not had: no build of BlipBridge has been run inside
  a real 32-bit PowerPoint. The backend's behaviour is validated in a real
  PowerPoint and its 32-bit build is validated in CI, but not the two together.

  Details: $repo/blob/main/docs/windows_x86.md
"@
}

@"
BlipBridge $Version - windows-$Architecture
fast reusable image textures for PowerPoint Shapes

WHAT IS IN HERE
  $dllName   the library
  BlipBridge.bas   the VBA module to import
  blipbridge.h     the C ABI, if you are calling from something other than VBA
  LICENSE          MIT

INSTALL
  1. Copy $dllName next to your .pptm.
  2. Import BlipBridge.bas into the VBA project.

  There is nothing to register and nothing to install. No regsvr32, no ProgID,
  no COM add-in.

USE - the short way
  BlipBridge.UserPicture2 shp, "C:\textures\brick.png"

  It reads and decodes the file once however many Shapes get it, picks the
  accelerated path where the Shape class has one and Office's own path where it
  does not, and does no work at all when that Shape already carries that image.

USE - with explicit texture handles
  BlipBridge.Initialize
  tex = BlipBridge.LoadTexture(bytes)
  BlipBridge.ApplyTexture shp, tex
  BlipBridge.ReleaseTexture tex
  BlipBridge.Shutdown

$status

DOCUMENTATION AND SOURCE
  $repo

BlipBridge is not affiliated with or endorsed by Microsoft. It calls
undocumented internals of Microsoft Office, which may change or break in any
update; the version guards exist so that a change stops it rather than corrupts
anything.
"@ | Set-Content (Join-Path $target 'README.txt') -Encoding utf8

# --- archive and checksum ----------------------------------------------------
$zip = Join-Path $distRoot "$name.zip"
if (Test-Path $zip) { Remove-Item $zip -Force }
Compress-Archive -Path (Join-Path $target '*') -DestinationPath $zip

$hash = (Get-FileHash $zip -Algorithm SHA256).Hash.ToLower()
# The two-space form is what `sha256sum -c` expects, so a user can verify with
# the tool they already have rather than by eye.
"$hash  $name.zip" | Set-Content "$zip.sha256" -Encoding ascii

Get-ChildItem $target | Select-Object Name, Length
''
"package : $zip"
"sha256  : $hash"
