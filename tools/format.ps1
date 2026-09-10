$ErrorActionPreference = "Stop"

$clangFormatCandidates = @(
    "C:\msys64\ucrt64\bin\clang-format.exe",
    "C:\tools\msys64\ucrt64\bin\clang-format.exe"
)

$clangFormat = $null

foreach ($candidate in $clangFormatCandidates) {
    if (Test-Path $candidate) {
        $clangFormat = $candidate
        break
    }
}

if (-not $clangFormat) {
    Write-Error @"
clang-format was not found in MSYS2.

Expected something like:
    C:\msys64\ucrt64\bin\clang-format.exe

Install it from MSYS2 UCRT64 with:
    pacman -S mingw-w64-ucrt-x86_64-clang
"@
    exit 1
}

Write-Host "Using clang-format:"
Write-Host "  $clangFormat"
Write-Host ""

$files = git ls-files `
    "*.c" `
    "*.cc" `
    "*.cpp" `
    "*.cxx" `
    "*.h" `
    "*.hh" `
    "*.hpp" `
    "*.hxx"

if (-not $files) {
    Write-Host "No C/C++ source files found."
    exit 0
}

$count = 0

foreach ($file in $files) {
    Write-Host "Formatting $file"

    & $clangFormat `
        -i `
        --style=file `
        $file

    if ($LASTEXITCODE -ne 0) {
        Write-Error "clang-format failed on: $file"
        exit $LASTEXITCODE
    }

    $count++
}

Write-Host ""
Write-Host "Formatted $count source files."