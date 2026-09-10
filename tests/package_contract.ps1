<#
.SYNOPSIS
Checks release package naming, contents and checksums with non-executable fixtures.
.DESCRIPTION
Exercises both architectures and both CMake output names without requiring an
x86 compiler. Fixtures are isolated from real release artifacts under artifacts/.
This validates packaging only, not binary architecture or Office compatibility.
#>
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

$root = Split-Path $PSScriptRoot -Parent
$fixtureRoot = Join-Path $root "artifacts/package-contract-$([Guid]::NewGuid())"
foreach ($directory in @('tools', 'vba', 'include/blipbridge', 'build')) {
    New-Item -ItemType Directory -Path (Join-Path $fixtureRoot $directory) -Force | Out-Null
}
foreach ($file in @('tools/make_dist.ps1', 'vba/BlipBridge.bas', 'include/blipbridge/blipbridge.h', 'LICENSE')) {
    Copy-Item (Join-Path $root $file) (Join-Path $fixtureRoot $file)
}

foreach ($architecture in @('x64', 'x86')) {
    foreach ($legacy in @($true, $false)) {
        $buildDir = Join-Path $fixtureRoot "build/$architecture-$legacy"
        New-Item -ItemType Directory -Path $buildDir | Out-Null
        $dllName = "BlipBridge-$architecture.dll"
        $inputName = if ($legacy) { 'BlipBridge.dll' } else { $dllName }
        $fixtureBytes = [System.Text.Encoding]::UTF8.GetBytes("Packaging fixture: $architecture-$legacy")
        [System.IO.File]::WriteAllBytes((Join-Path $buildDir $inputName), $fixtureBytes)

        & (Join-Path $fixtureRoot 'tools/make_dist.ps1') `
            -Architecture $architecture -BuildDir $buildDir -Version '0.4.0' | Out-Null

        $zipName = "blipbridge-v0.4.0-windows-$architecture.zip"
        $zipPath = Join-Path $fixtureRoot "dist/$zipName"
        $archive = [System.IO.Compression.ZipFile]::OpenRead($zipPath)
        try {
            $names = @($archive.Entries | ForEach-Object { $_.FullName })
            foreach ($expected in @($dllName, 'BlipBridge.bas', 'blipbridge.h', 'README.txt', 'LICENSE')) {
                if ($expected -notin $names) {
                    throw "$zipName is missing $expected"
                }
            }
            if ('BlipBridge.dll' -in $names) {
                throw "$zipName contains the obsolete DLL name"
            }
            $reader = [System.IO.StreamReader]::new($archive.GetEntry('README.txt').Open())
            try {
                if (-not $reader.ReadToEnd().Contains("Copy $dllName next to")) {
                    throw "$zipName has incorrect installation instructions"
                }
            }
            finally {
                $reader.Dispose()
            }
        }
        finally {
            $archive.Dispose()
        }

        $hash = (Get-FileHash $zipPath -Algorithm SHA256).Hash.ToLowerInvariant()
        if ((Get-Content "$zipPath.sha256" -Raw).Trim() -cne "$hash  $zipName") {
            throw "$zipName has an invalid checksum file"
        }
        Write-Host "PASS: $architecture, legacy output = $legacy"
    }
}
Write-Host "Packaging fixtures retained at $fixtureRoot"
