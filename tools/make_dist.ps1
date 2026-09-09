<#
.SYNOPSIS
Builds a release package for one architecture.

.DESCRIPTION
Produces exactly what a user needs and nothing else:

    dist/blipbridge-<version>-windows-<arch>/BlipBridge.dll
                                            /BlipBridge.bas
                                            /blipbridge.h
                                            /README.txt
                                            /LICENSE
    dist/blipbridge-<version>-windows-<arch>.zip
    dist/blipbridge-<version>-windows-<arch>.zip.sha256

The README inside the archive is written per architecture, because the two say
genuinely different things: x64 has a validated native backend, x86 does not yet
and must say so plainly rather than leaving a user to discover it.

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

$dll = Join-Path $BuildDir 'BlipBridge.dll'
if (-not (Test-Path $dll)) {
    throw "BlipBridge.dll not found at $dll - build $Configuration for $Architecture first"
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

$name = "blipbridge-$Version-windows-$Architecture"
$distRoot = Join-Path $root 'dist'
$target = Join-Path $distRoot $name
if (Test-Path $target) { Remove-Item $target -Recurse -Force }
New-Item -ItemType Directory -Path $target -Force | Out-Null

Copy-Item $dll (Join-Path $target 'BlipBridge.dll')
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

NATIVE BACKEND STATUS
  Validated against Office 16.0.14334.20848 and refused on any other build.
  Call BlipBridge.IsAvailable to check; BlipBridge.Version explains what it
  found.

  This is not universal Office compatibility. The backend depends on internal
  Office layouts verified per build, and it fails closed rather than guessing.

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

NATIVE BACKEND STATUS - READ THIS
  *** There is no accelerated backend in this package yet. ***

  The library loads, reports its version, and answers every texture call with a
  specific refusal. It will not fill anything.

  The 32-bit PowerPoint internals have not been reverse-engineered or validated,
  and the 64-bit implementation is not portable to them by recompilation - it
  depends on the x64 calling convention, on per-build module addresses, and on
  object layouts 32-bit Office does not share. Rather than ship something that
  compiles and might corrupt a document, the 64-bit backend is not built into
  this binary at all.

  This package exists so the ABI, the wrapper and the packaging can be tested on
  32-bit Office. Use it for that. Do not expect it to accelerate anything.

  Progress: $repo/blob/main/docs/windows_x86.md
"@
}

@"
BlipBridge $Version - windows-$Architecture
fast reusable image textures for PowerPoint Shapes

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
