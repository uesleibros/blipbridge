<#
.SYNOPSIS
Builds the universal Windows release package.

.DESCRIPTION
One archive holding both architectures, because the user should not have to know
which PowerPoint they have before downloading:

    blipbridge-vX.Y.Z-windows.zip
        BlipBridge-x64.dll
        BlipBridge-x86.dll
        BlipBridge.bas
        blipbridge.h
        README.md
        LICENSE
        docs/...

BlipBridge.bas picks the DLL matching the PowerPoint *process* at run time, so
both can sit beside the same presentation and only the right one is ever loaded.

Nothing from the build tree, the research DLL, the experiments or the test
fixtures goes in. Only files named here are copied, so nothing can arrive by
accident.

.PARAMETER X64Dll
Path to the built BlipBridge-x64.dll.

.PARAMETER X86Dll
Path to the built BlipBridge-x86.dll. Omit to build a package without it, which
is only appropriate when x86 genuinely could not be produced.

.PARAMETER Version
Release version without the v prefix. Defaults to what the header declares.
#>
param(
    [string]$X64Dll = '',
    [string]$X86Dll = '',
    [string]$Version = ''
)

$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent

if (-not $X64Dll) { $X64Dll = Join-Path $root 'build/Release/BlipBridge-x64.dll' }

# The header is the single source of the release number.
if (-not $Version) {
    $header = Get-Content (Join-Path $root 'include/blipbridge/blipbridge.h') -Raw
    if ($header -match 'BB_VERSION_MAJOR\s+(\d+)') { $major = $Matches[1] }
    if ($header -match 'BB_VERSION_MINOR\s+(\d+)') { $minor = $Matches[1] }
    if ($header -match 'BB_VERSION_PATCH\s+(\d+)') { $patch = $Matches[1] }
    if (-not ($major -and $minor -and $patch)) {
        throw 'Could not read the version from include/blipbridge/blipbridge.h'
    }
    $Version = "$major.$minor.$patch"
}

if ($Version -notmatch '^\d+\.\d+\.\d+([-.][0-9A-Za-z.-]+)?$') {
    throw "Invalid package version: $Version. Expected 0.5.0 without the v prefix."
}

if (-not (Test-Path $X64Dll)) {
    throw "BlipBridge-x64.dll not found at $X64Dll - build Release for x64 first."
}

$name = "blipbridge-v$Version-windows"
$distRoot = Join-Path $root 'dist'
$target = Join-Path $distRoot $name

# Resolve before deleting recursively, so a mistyped version cannot reach outside dist/.
$target = [System.IO.Path]::GetFullPath($target)
$expectedParent = [System.IO.Path]::GetFullPath($distRoot)
if ((Split-Path $target -Parent) -ne $expectedParent) {
    throw "Package directory must be directly under $expectedParent"
}
if (Test-Path $target) { Remove-Item -LiteralPath $target -Recurse -Force }
New-Item -ItemType Directory -Path $target -Force | Out-Null

Copy-Item $X64Dll (Join-Path $target 'BlipBridge-x64.dll')

$hasX86 = $false
if ($X86Dll -and (Test-Path $X86Dll)) {
    Copy-Item $X86Dll (Join-Path $target 'BlipBridge-x86.dll')
    $hasX86 = $true
}

Copy-Item (Join-Path $root 'vba/BlipBridge.bas') (Join-Path $target 'BlipBridge.bas')
Copy-Item (Join-Path $root 'include/blipbridge/blipbridge.h') (Join-Path $target 'blipbridge.h')
Copy-Item (Join-Path $root 'README.md') (Join-Path $target 'README.md')
Copy-Item (Join-Path $root 'LICENSE') (Join-Path $target 'LICENSE')
Copy-Item (Join-Path $root 'CHANGELOG.md') (Join-Path $target 'CHANGELOG.md')

# The documents a user actually needs to use the library. Research journals and
# raw evidence transcripts stay in the repository.
$docs = Join-Path $target 'docs'
New-Item -ItemType Directory -Path $docs -Force | Out-Null
foreach ($doc in @('c_abi.md', 'picture_cache.md', 'resampling.md',
                   'shape_compatibility.md', 'safety_model.md',
                   'benchmarks.md', 'windows_x86.md')) {
    $source = Join-Path $root "docs/$doc"
    if (Test-Path $source) { Copy-Item $source (Join-Path $docs $doc) }
}

# --- archive -----------------------------------------------------------------
$zip = Join-Path $distRoot "$name.zip"
if (Test-Path $zip) { Remove-Item $zip -Force }
# The folder itself goes in, not its contents, so extracting produces one
# directory rather than scattering files into the download folder. The release
# workflow builds the same layout.
Compress-Archive -Path $target -DestinationPath $zip

# --- checksums ---------------------------------------------------------------
# The two-space form is what `sha256sum -c` expects, so a user can verify with a
# tool they already have rather than by eye.
$hash = (Get-FileHash $zip -Algorithm SHA256).Hash.ToLower()
"$hash  $name.zip" | Set-Content (Join-Path $distRoot 'SHA256SUMS.txt') -Encoding ascii

# --- verify what actually went in --------------------------------------------
Add-Type -AssemblyName System.IO.Compression.FileSystem
$archive = [IO.Compression.ZipFile]::OpenRead($zip)
try {
    $entries = $archive.Entries | ForEach-Object { $_.FullName }
} finally {
    $archive.Dispose()
}

$required = @('BlipBridge-x64.dll', 'BlipBridge.bas', 'README.md', 'LICENSE')
if ($hasX86) { $required += 'BlipBridge-x86.dll' }
foreach ($entry in $required) {
    if (-not ($entries | Where-Object { $_ -like "*/$entry" -or $_ -like "*\$entry" })) {
        throw "Package is missing $entry"
    }
}

# Nothing from the build tree or the research surface may travel.
foreach ($entry in $entries) {
    foreach ($banned in @('BlipBridgeResearch', '.obj', '.o', '.ilk', '.exp', '.lib',
                          'experiments/', 'build/', 'CMakeFiles')) {
        if ($entry -like "*$banned*") { throw "Package contains $entry ($banned)" }
    }
}

''
"package  : $zip"
"sha256   : $hash"
"contents :"
$entries | Sort-Object | ForEach-Object { "  $_" }
if (-not $hasX86) {
    ''
    'NOTE: no x86 DLL was included. Pass -X86Dll to produce the universal package.'
}
