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

## Validate the file you ship, not a copy of it

v0.7.0 shipped a module that could not be imported, and one reason nothing
noticed is that every gate read `vba/BlipBridge.bas` out of the source tree while
the release archive carried its own copy. Anything that happened to the file
between those two points - a line-ending conversion, a truncated copy, the wrong
file entirely - was invisible.

So this takes -Path. The release workflow points it at the .bas **extracted from
the finished archive**, which is the only copy a user will ever have.

Run it from the repo root. Exits non-zero on any finding.
#>
param(
    # The module to check. Defaults to the source tree's copy; the release
    # workflow passes the one it pulled back out of the ZIP.
    [string]$Path
)

$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
if (-not $Path) { $Path = Join-Path $root 'vba/BlipBridge.bas' }
$path = $Path
if (-not (Test-Path $path)) { throw "Missing $path" }

$text = Get-Content $path -Raw
$findings = New-Object System.Collections.Generic.List[string]
$moduleFile = Split-Path $path -Leaf

<#
The file-level checks: the things that stop a module being imported at all,
before a single declaration is read, and the things a packaging step can do to a
file that no amount of source-tree checking would ever see.
#>
$bytes = [IO.File]::ReadAllBytes($path)
if ($bytes.Length -eq 0) { throw "$path is empty" }

# A .bas must open with its module name attribute. Without it the VBA editor
# refuses the import outright rather than reporting a compile error.
if ($text -notmatch '^\s*(﻿)?Attribute\s+VB_Name\s*=\s*"[^"]+"') {
    $findings.Add('the file does not begin with an Attribute VB_Name line')
}

<#
Line endings, which is where v0.7.0 went wrong in a way nothing could see.

The VBA editor writes CRLF when it exports a module; a .bas is that format. The
repository stored the module LF, so a developer with core.autocrlf=true edited
and tested a CRLF file while CI checked out LF and packaged that - and the file
users received was not the file anyone had validated. .gitattributes pins it now,
and this is what proves the pin holds all the way into the archive.
#>
$crlf = ([regex]::Matches($text, "`r`n")).Count
$bareLf = ([regex]::Matches($text, "`n")).Count - $crlf
if ($bareLf -gt 0) {
    $findings.Add("$bareLf line(s) end in LF rather than CRLF - a .bas is a CRLF format")
}
if ($text -match "`r(?!`n)") {
    $findings.Add('the file contains a bare CR')
}

# A NUL byte means the file was truncated or written in the wrong encoding, and
# nothing after it will parse.
if ($bytes -contains 0) {
    $findings.Add('the file contains a NUL byte, so it is truncated or mis-encoded')
}

<#
Block structure. Each of these must close, and an unclosed one makes everything
after it unparseable - which is the shape of damage a bad merge or a partial copy
produces, and which a name-grep gate would sail straight past.
#>
$blockPairs = @(
    @{ What = '#If'; Open = '(?m)^\s*#If\s'; Close = '(?m)^\s*#End\s+If' },
    @{ What = 'Type'; Open = '(?m)^\s*(?:Public\s+|Private\s+)?Type\s+\w+'; Close = '(?m)^\s*End\s+Type' },
    @{ What = 'Enum'; Open = '(?m)^\s*(?:Public\s+|Private\s+)?Enum\s+\w+'; Close = '(?m)^\s*End\s+Enum' }
)
foreach ($pair in $blockPairs) {
    $opened = ([regex]::Matches($text, $pair.Open)).Count
    $closed = ([regex]::Matches($text, $pair.Close)).Count
    if ($opened -ne $closed) {
        $findings.Add("$($pair.What): $opened opened but $closed closed")
    }
}

<#
Every Declare must name the library it binds to. Without Lib the module still
imports and the first call fails at run time, which is a much worse way to find
out than a refused import.
#>
$declareLines = [regex]::Matches($text, '(?m)^[ \t]*(?:Public |Private )?Declare[ \t].*$')
foreach ($declare in $declareLines) {
    if ($declare.Value -notmatch 'Lib[ \t]+"') {
        $findings.Add("a Declare has no Lib clause: $($declare.Value.Trim())")
    }
}
if ($declareLines.Count -eq 0) {
    $findings.Add('the module declares no native functions at all, so it is not the wrapper')
}

<#
Attribute lines. A .bas carries them at the top and the VBA editor writes them
in one exact shape; a malformed one is refused at import, before anything in the
module is read. Only VB_Name is required, but any Attribute present must parse.
#>
foreach ($attribute in [regex]::Matches($text, '(?m)^[ \t]*Attribute[ \t].*$')) {
    if ($attribute.Value -notmatch '^[ \t]*Attribute[ \t]+VB_[A-Za-z_]+[ \t]*=[ \t]*\S') {
        $findings.Add("malformed Attribute line: $($attribute.Value.Trim())")
    }
}

<#
Conditional compilation, by depth rather than by count. Counting #If against
#End If says nothing about order, and a #Else that is not inside an #If is just
as fatal as an unclosed block - it is a different mistake with the same symptom,
so it is worth telling them apart.
#>
$depth = 0
$lineNumber = 0
foreach ($line in ($text -split "`r?`n")) {
    $lineNumber++
    $trimmed = $line.Trim()
    if ($trimmed -match '^#If[ \t]') { $depth++ ; continue }
    if ($trimmed -match '^#End[ \t]+If') {
        $depth--
        if ($depth -lt 0) {
            $findings.Add("line ${lineNumber}: #End If with no matching #If")
            $depth = 0
        }
        continue
    }
    if ($trimmed -match '^#(Else|ElseIf)\b' -and $depth -eq 0) {
        $findings.Add("line ${lineNumber}: $trimmed outside any #If")
    }
}
if ($depth -ne 0) {
    $findings.Add("$depth conditional-compilation block(s) never closed")
}

<#
PtrSafe. This module is VBA7-only by design - it says so - and a Declare without
PtrSafe is a compile error there rather than a portability nicety. It is exactly
the sort of thing that survives review because it looks like every other line.
#>
foreach ($declare in $declareLines) {
    if ($declare.Value -notmatch '\bPtrSafe\b') {
        $findings.Add("a Declare is missing PtrSafe: $($declare.Value.Trim())")
    }
}

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
    throw "$($findings.Count) problem(s) in $moduleFile ($path)"
}

"${moduleFile}: $($declarations.Count) declarations, $blocks doc blocks, $crlf CRLF lines - no problems ($path)"
