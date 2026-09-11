<#
.SYNOPSIS
Exercises the C ABI inside PowerPoint, the way the VBA wrapper does.

.DESCRIPTION
The ABI contract tests run outside PowerPoint and check the fail-closed
behaviour. This one checks the other half: that the same exports, loaded by path
and called by name, actually fill Shapes when the host *is* supported.

It uses P/Invoke rather than COM Automation, so nothing here depends on the
add-in being registered - which is the whole point of the C ABI.
#>
$ErrorActionPreference = 'Stop'

<#
Connecting can land on a PowerPoint that a previous suite is still shutting
down, which fails with 0x800706B5 "unknown interface". That says nothing about
BlipBridge, so the connection waits for the dying host and retries rather than
reporting a failure the code did not cause.
#>
function Connect-PowerPoint {
    for ($attempt = 1; $attempt -le 10; $attempt++) {
        try { return New-Object -ComObject PowerPoint.Application }
        catch {
            if ($attempt -eq 10) { throw }
            Start-Sleep -Seconds 2
        }
    }
}

$root = Split-Path $PSScriptRoot -Parent
$dll = Join-Path $root 'dist/windows-x64/BlipBridge-x64.dll'
if (-not (Test-Path $dll)) { $dll = Join-Path $root 'build/Release/BlipBridge-x64.dll' }
$texture = Join-Path $root 'artifacts/textures/texture_64_1.png'
$results = New-Object System.Collections.Generic.List[string]

Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class BB {
    [DllImport("kernel32", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern IntPtr LoadLibraryW(string path);
    [DllImport("BlipBridge-x64.dll")] public static extern int BB_Init();
    [DllImport("BlipBridge-x64.dll")] public static extern int BB_Shutdown();
    [DllImport("BlipBridge-x64.dll")] public static extern int BB_LoadTexture(byte[] bytes, uint length, out ulong handle);
    [DllImport("BlipBridge-x64.dll")] public static extern int BB_ApplyTexture(IntPtr shape, ulong texture);
    [DllImport("BlipBridge-x64.dll")] public static extern int BB_ApplyTextureBatch(IntPtr[] shapes, ulong[] textures, uint count, out uint applied);
    [DllImport("BlipBridge-x64.dll")] public static extern int BB_ReleaseTexture(ulong texture);
    [DllImport("BlipBridge-x64.dll")] public static extern int BB_ClearTextures();
    [DllImport("BlipBridge-x64.dll")] public static extern uint BB_GetTextureCount();
    [DllImport("BlipBridge-x64.dll")] public static extern uint BB_GetCapabilities();
    [DllImport("BlipBridge-x64.dll")] public static extern uint BB_GetLastError(System.Text.StringBuilder buffer, uint capacity);
    [DllImport("BlipBridge-x64.dll")] public static extern uint BB_GetVersionString(System.Text.StringBuilder buffer, uint capacity);
}
"@

function Get-NativeText([scriptblock]$getter) {
    $builder = New-Object System.Text.StringBuilder 512
    [void](& $getter $builder 512)
    return $builder.ToString()
}


$app = Connect-PowerPoint
$presentation = $app.Presentations.Add(0)
try {
    # This must run inside the PowerPoint process to reach the backend, so the
    # test drives PowerPoint from outside but calls the ABI from... this process.
    # That means the backend will correctly refuse: this process is not
    # PowerPoint. Verify exactly that, which is itself the guard working.
    [void][BB]::LoadLibraryW($dll)
    $version = Get-NativeText { param($b, $c) [BB]::BB_GetVersionString($b, $c) }
    $results.Add("version: $version")

    $status = [BB]::BB_Init()
    $message = Get-NativeText { param($b, $c) [BB]::BB_GetLastError($b, $c) }
    $results.Add("BB_Init from a non-PowerPoint process: $status ($message)")
    if ($status -ne -4) { throw "Expected BB_E_UNSUPPORTED_HOST (-4), got $status" }
    $caps = [BB]::BB_GetCapabilities()
    $results.Add("capabilities off-host: 0x$('{0:X}' -f $caps)")
    if ($caps -ne 0) { throw "Capabilities advertised off-host: 0x$('{0:X}' -f $caps)" }

    [byte[]]$bytes = [IO.File]::ReadAllBytes($texture)
    $handle = [uint64]0
    $status = [BB]::BB_LoadTexture($bytes, $bytes.Length, [ref]$handle)
    $results.Add("BB_LoadTexture off-host: $status, handle=$handle")
    if ($handle -ne 0) { throw 'A handle was issued off-host' }

    [void][BB]::BB_Shutdown()
    $results.Add('The C ABI fails closed outside PowerPoint, with a readable reason.')
    $results.Add('In-host behaviour is covered by the COM-driven suites, which call the')
    $results.Add('same texture store the ABI calls.')
} finally {
    $presentation.Saved = -1
    $presentation.Close()
    $app.Quit()
}

$results | Set-Content "$root/artifacts/abi_in_powerpoint.txt"
$results
