<#
.SYNOPSIS
Proves tools/check_vba_module.ps1 actually catches what it claims to.

.DESCRIPTION
v0.7.0 shipped a VBA module that could not be imported while every release gate
reported success. The response was a gate - and a gate that has never failed is
a gate nobody has tested, which is the same position again with more scripts.

So this breaks the real module one way at a time and requires the gate to refuse
it *by name*. Failing for some other reason does not count: a check that fires on
the wrong input is how a gate ends up passing the thing it was written for.

Each case corresponds to a way a .bas becomes unimportable, and two of them are
the actual v0.7.0 defects: LF line endings and - through the optional-parameter
check the gate already had - a user-defined type as an optional parameter.

The gate runs as a child process. It writes its findings to the pipeline and then
throws, so an in-process call loses the findings and leaves only the exception;
a child process keeps both, and is how CI invokes it anyway.

Run it from anywhere. Exits non-zero if any check fails to fire.
#>
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$gate = Join-Path $PSScriptRoot 'check_vba_module.ps1'
$source = Join-Path $root 'vba/BlipBridge.bas'
if (-not (Test-Path $source)) { throw "Missing $source" }

$work = Join-Path ([IO.Path]::GetTempPath()) ('bb_vba_gate_' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $work | Out-Null

$original = [IO.File]::ReadAllText($source)

function Invoke-Gate([string]$target) {
    # No 2>&1: redirecting a native command's stderr in Windows PowerShell wraps
    # every line in an ErrorRecord and makes a successful call look like a failed
    # one. The findings matched on below go to stdout regardless.
    $previous = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        $out = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $gate -Path $target |
            Out-String
        $code = $LASTEXITCODE
    } finally {
        $ErrorActionPreference = $previous
    }
    return [pscustomobject]@{ Output = $out; Failed = ($code -ne 0) }
}

<#
Each case: how to break the module, and the wording that proves the gate noticed
that particular breakage rather than tripping over something else.
#>
$cases = @(
    @{ Name = 'LF line endings (the v0.7.0 packaging defect)'
       Expect = 'end in LF'
       Mutate = { param($t) $t -replace "`r`n", "`n" } }

    @{ Name = 'optional user-defined type (the v0.7.0 compile defect)'
       Expect = 'does not allow a user-defined type'
       Mutate = { param($t) $t -replace 'ByRef request As BlipBridgeImageRequest\)',
                                        'Optional ByRef request As BlipBridgeImageRequest)' } }

    @{ Name = 'no VB_Name attribute'
       Expect = 'Attribute VB_Name'
       Mutate = { param($t) $t -replace '^Attribute VB_Name[^\r\n]*\r\n', '' } }

    @{ Name = 'malformed Attribute line'
       Expect = 'malformed Attribute'
       Mutate = { param($t) $t -replace 'Attribute VB_Name = "BlipBridge"',
                                        'Attribute VB_Name "BlipBridge"' } }

    @{ Name = 'unclosed Type block'
       Expect = 'opened but'
       Mutate = { param($t) $t -replace '(?m)^End Type\r?\n', '' } }

    @{ Name = 'Declare with no Lib clause'
       Expect = 'no Lib clause'
       Mutate = { param($t) $t -replace 'Function BB_Init Lib "BlipBridge-x64.dll"',
                                        'Function BB_Init' } }

    @{ Name = 'Declare with no PtrSafe'
       Expect = 'missing PtrSafe'
       Mutate = { param($t) $t -replace 'Private Declare PtrSafe Function BB_Init',
                                        'Private Declare Function BB_Init' } }

    @{ Name = 'conditional-compilation directive outside any #If'
       Expect = 'outside any #If'
       Mutate = { param($t) $t -replace '(?m)^Option Explicit\r?\n', "Option Explicit`r`n#Else`r`n" } }

    @{ Name = 'unclosed #If'
       Expect = 'never closed|opened but'
       Mutate = { param($t) $t -replace '(?m)^#End If\r?\n', '' } }

    @{ Name = 'NUL byte from a truncated or mis-encoded copy'
       Expect = 'NUL byte'
       Mutate = { param($t) $t.Insert(200, [string][char]0) } }

    @{ Name = 'a doc line that lost its apostrophe'
       Expect = 'looks like a doc line but is not a comment'
       Mutate = { param($t) $t -replace "(?m)^' \* @function Initialize", '* @function Initialize' } }
)

$failures = 0
foreach ($case in $cases) {
    $broken = Join-Path $work (($case.Name -replace '\W', '_') + '.bas')
    [IO.File]::WriteAllText($broken, (& $case.Mutate $original))

    $result = Invoke-Gate $broken
    $matched = $result.Output -match $case.Expect
    if ($result.Failed -and $matched) {
        Write-Output ("  ok   refused: {0}" -f $case.Name)
    } else {
        Write-Output ("  FAIL {0}  (refused={1} named={2})" -f $case.Name, $result.Failed, $matched)
        $lines = $result.Output -split "`r?`n" | Where-Object { $_ } | Select-Object -First 3
        Write-Output ('       ' + ($lines -join ' | '))
        $failures++
    }
}

# And the unmodified module must still pass. Without this the suite would be
# satisfied by a gate that rejects everything.
$clean = Invoke-Gate $source
if (-not $clean.Failed) {
    Write-Output '  ok   the real module still passes'
} else {
    Write-Output '  FAIL the real module no longer passes'
    Write-Output ('       ' + $clean.Output)
    $failures++
}

Remove-Item -Recurse -Force $work
if ($failures -gt 0) { throw "$failures gate check(s) did not fire" }
Write-Output "VBA gate self-test: all $($cases.Count) breakages refused by name"
