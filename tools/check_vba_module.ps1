<#
.SYNOPSIS
Static checks on BlipBridge.bas that the VBA editor would otherwise find first.

.DESCRIPTION
The wrapper is not compiled by anything in CI - there is no VBA compiler on a
build agent - so a module that does not compile can and did ship. This catches
the rules that are mechanical enough to check without one.

  Optional user-defined types   VBA refuses "Optional x As SomeUdt" outright:
                                an optional parameter must be Variant or have a
                                default, and a UDT can be neither. This shipped
                                in v0.7.0 and made the whole module fail to
                                compile, which is exactly as bad as it sounds.

  Optional without a default    Legal only for Variant. Anything else needs "= x".

  Optional ordering             Once a parameter is Optional, every one after it
                                must be too.

  Public Object/Variant         The public API is strongly typed on purpose.

  Declared but missing          Every function a caller is told about has to be
                                there - the release gate greps for names, which
                                would pass on a comment.

  Doc blocks that have drifted  @function naming a different procedure, @param
                                for a parameter that no longer exists, a
                                parameter with no @param, a Function with no
                                @return. A stale @param is worse than none: it
                                tells a caller to pass something that is gone.

  Comment syntax                A line starting with * or @ that lost its
                                apostrophe is code. A comment ending in " _" does
                                not continue - VBA has no comment continuation,
                                so the next line is read as code.

## What this is not

It is not a VBA compiler, and it cannot become one here: compiling would need the
VBA project object model, which is off by default for good security reasons and
is not something a test should switch on. So this checks the rules that are
mechanical, and the definitive check remains opening the module in the editor -
which is how the defect it exists for was found.

Run it from the repo root. Exits non-zero on any finding.
#>
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$path = Join-Path $root 'vba/BlipBridge.bas'
if (-not (Test-Path $path)) { throw "Missing $path" }

$text = Get-Content $path -Raw
$findings = New-Object System.Collections.Generic.List[string]

# Types the module defines. Only these can be confused for a parameter type that
# VBA will not accept as Optional.
$userTypes = [regex]::Matches($text, '(?m)^\s*Public\s+Type\s+(\w+)') |
    ForEach-Object { $_.Groups[1].Value }
$enums = [regex]::Matches($text, '(?m)^\s*Public\s+Enum\s+(\w+)') |
    ForEach-Object { $_.Groups[1].Value }

# Procedure declarations, with their continuations joined into one line each.
$joined = $text -replace '\s_\r?\n\s*', ' '
# The parameter list has to tolerate the "()" of an array parameter. Matching
# [^)]* instead would stop at the first one - which silently truncated
# "ByRef bytes() As Byte, ByRef request As ..." to nothing useful, and would have
# missed two of the four declarations this check exists because of.
$declarations = [regex]::Matches($joined,
    '(?m)^\s*(?:Public|Private)?\s*(?:Declare\s+PtrSafe\s+)?(?:Function|Sub)\s+(\w+)\s*\(((?:[^()]|\(\s*\))*)\)')

foreach ($declaration in $declarations) {
    $name = $declaration.Groups[1].Value
    $parameters = $declaration.Groups[2].Value
    if ([string]::IsNullOrWhiteSpace($parameters)) { continue }

    $seenOptional = $false
    foreach ($parameter in ($parameters -split ',')) {
        $trimmed = $parameter.Trim()
        if (-not $trimmed) { continue }

        $isOptional = $trimmed -match '^\s*Optional\b'
        if ($isOptional) { $seenOptional = $true }
        elseif ($seenOptional -and $trimmed -notmatch '^\s*ParamArray\b') {
            $findings.Add("$name : '$trimmed' follows an Optional parameter but is not Optional")
        }
        if (-not $isOptional) { continue }

        $type = if ($trimmed -match '\bAs\s+(\w+)') { $Matches[1] } else { 'Variant' }
        $hasDefault = $trimmed -match '='

        if ($userTypes -contains $type) {
            $findings.Add("$name : 'Optional ... As $type' - VBA does not allow a user-defined type as an optional parameter")
            continue
        }
        if ($trimmed -match '\(\s*\)') {
            $findings.Add("$name : an array parameter cannot be Optional")
            continue
        }
        if (-not $hasDefault -and $type -ne 'Variant') {
            $findings.Add("$name : 'Optional ... As $type' needs a default, or it must be Variant")
        }
    }
}

# The public surface stays strongly typed.
foreach ($line in ($joined -split "`r?`n")) {
    if ($line -match '^\s*Public\s+(Function|Sub)\s' -and $line -match '\bAs\s+(Object|Variant)\b') {
        $findings.Add("public signature uses $($Matches[1]): $($line.Trim())")
    }
}

# Everything the release gates grep for must be a real declaration, not a comment.
$required = @(
    'LoadImage', 'LoadImageProcessed', 'LoadImageFromFile', 'LoadImageProcessedFromFile',
    'GetImageSize', 'ReleaseImage', 'ClearImages', 'GetImageCount',
    'CreateTextureFromImage', 'WarpImageQuad', 'ApplyImageQuad',
    'LoadTextureScaled', 'LoadTextureScaledFromFile', 'ImageRequest', 'QuadPoints',
    'ApplyTexture', 'ApplyTextureRange', 'ApplyTextureIfChanged', 'ReleaseTexture'
)
foreach ($name in $required) {
    if ($joined -notmatch "(?m)^\s*Public\s+(Function|Sub)\s+$name\s*\(") {
        $findings.Add("no public declaration of $name")
    }
}


<#
The doc blocks have to still describe the code. A signature change leaves the
block behind, and a stale @param is worse than none: it tells a caller to pass
something that no longer exists. Checked here because nothing else would notice.
#>
$blocks = 0
$openMarkers = @()
$closeMarkers = @()
$fileLines = Get-Content $path
for ($index = 0; $index -lt $fileLines.Count; $index++) {
    $trimmed = $fileLines[$index].Trim()
    if ($trimmed -eq "'/**") { $openMarkers += $index }
    if ($trimmed -eq "' */") { $closeMarkers += $index }

    # A line starting with * or @ has lost its apostrophe and is now code.
    if ($trimmed -match '^[\*@]') {
        $findings.Add("line $($index + 1) looks like a doc line but is not a comment: $trimmed")
    }
    # VBA has no comment continuation: a comment ending in " _" does not join the
    # next line, it just ends, and the next line is read as code.
    if ($trimmed.StartsWith("'") -and $fileLines[$index].TrimEnd().EndsWith(' _')) {
        $findings.Add("line $($index + 1) is a comment ending in a continuation")
    }
}
if ($openMarkers.Count -ne $closeMarkers.Count) {
    $findings.Add("$($openMarkers.Count) doc blocks opened but $($closeMarkers.Count) closed")
}

foreach ($declaration in $declarations) {
    $name = $declaration.Groups[1].Value
    $parameters = $declaration.Groups[2].Value

    # Only the public surface is documented by convention; private helpers that
    # do have a block are still checked for accuracy.
    $line = ($joined -split "`r?`n" | Where-Object {
        $_ -match "(?m)^\s*(Public|Private)\s+(Function|Sub)\s+$name\s*\("
    } | Select-Object -First 1)
    if (-not $line) { continue }

    $names = @()
    foreach ($parameter in ($parameters -split ',')) {
        $trimmed = ($parameter -replace '^\s*(Optional\s+)?(ByVal|ByRef|ParamArray)?\s*', '').Trim()
        if ($trimmed -match '^(\w+)') { $names += $Matches[1] }
    }

    # The block immediately above the declaration, if there is one.
    $start = ($fileLines | Select-String -SimpleMatch -Pattern $declaration.Groups[0].Value.Split("`n")[0].Trim() |
        Select-Object -First 1)
    if (-not $start) { continue }
    $at = $start.LineNumber - 2
    while ($at -ge 0 -and $fileLines[$at].Trim() -eq '') { $at-- }
    if ($at -lt 0 -or $fileLines[$at].Trim() -ne "' */") { continue }
    $blockEnd = $at
    while ($at -ge 0 -and $fileLines[$at].Trim() -ne "'/**") { $at-- }
    if ($at -lt 0) { continue }
    $blocks++
    $doc = ($fileLines[$at..$blockEnd] -join "`n")

    $declared = [regex]::Match($doc, '@function\s+(\w+)')
    if ($declared.Success -and $declared.Groups[1].Value -ne $name) {
        $findings.Add("$name : doc says @function $($declared.Groups[1].Value)")
    }

    $documented = [regex]::Matches($doc, '@param\s+(\w+)') | ForEach-Object { $_.Groups[1].Value }
    foreach ($param in $documented) {
        if ($names -notcontains $param) {
            $findings.Add("$name : @param $param is not a parameter")
        }
    }
    foreach ($param in $names) {
        if ($documented -notcontains $param) {
            $findings.Add("$name : parameter $param is undocumented")
        }
    }

    if ($line -match '^\s*Public\s+Function\s' -and $doc -notmatch '@return') {
        $findings.Add("$name : public Function with no @return")
    }
    if ($line -match '^\s*(Public|Private)\s+Sub\s' -and $doc -match '@return') {
        $findings.Add("$name : Sub documents a @return")
    }
}

if ($findings.Count -gt 0) {
    $findings | ForEach-Object { "  FAIL $_" }
    throw "$($findings.Count) problem(s) in BlipBridge.bas"
}

"BlipBridge.bas: $($declarations.Count) declarations and $blocks doc blocks checked, no problems"
