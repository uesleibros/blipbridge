<#
.SYNOPSIS
Does filling a group overwrite its children's fills?

.DESCRIPTION
This decides whether the per-Shape skip cache is safe around groups.

The cache remembers "child C last received texture B". If applying a picture to
the *group* silently replaces each child's own fill, that memory is stale, and a
later request to put texture B back on the child would be skipped - leaving the
wrong image on screen. That is the exact silent failure the cache must never
produce.

So this measures it directly, with two visibly different images, reading each
child's fill before and after the group is filled. Nothing is inferred from what
PowerPoint "probably" does.

Uses ordinary Fill.UserPicture for the group so the answer is a fact about
PowerPoint rather than about BlipBridge.
#>
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$msoTrue = -1
$ppLayoutBlank = 12
$textureA = Join-Path $root 'artifacts/textures/texture_128_0.png'
$textureB = Join-Path $root 'artifacts/textures/texture_64_1.png'

$lines = New-Object System.Collections.Generic.List[string]
$app = New-Object -ComObject PowerPoint.Application
$presentation = $app.Presentations.Add($msoTrue)
try {
    $slide = $presentation.Slides.Add(1, $ppLayoutBlank)
    $a = $slide.Shapes.AddShape(1, 40, 40, 80, 80)
    $b = $slide.Shapes.AddShape(1, 40, 140, 80, 80)

    # Give each child its own distinct picture fill first.
    $a.Fill.UserPicture($textureB)
    $b.Fill.UserPicture($textureB)
    $lines.Add("children before grouping: a.Fill.Type=$($a.Fill.Type) b.Fill.Type=$($b.Fill.Type)")

    $group = $slide.Shapes.Range(@($a.Name, $b.Name)).Group()
    $child1 = $group.GroupItems.Item(1)
    $child2 = $group.GroupItems.Item(2)
    $lines.Add("after grouping: child1.Fill.Type=$($child1.Fill.Type) child2.Fill.Type=$($child2.Fill.Type)")

    <#
     Fill.Type is 6 for any picture fill, so it cannot say *which* picture a child
     shows. The child is rendered to PNG before and after the group is filled;
     if those bytes differ, filling the group changed what the child displays,
     and any cached memory of the child's own texture is stale.
    #>
    function Export-Shape($shape) {
        $file = Join-Path $env:TEMP ('bbchild_' + [guid]::NewGuid().ToString('N') + '.png')
        try {
            $shape.Export($file, 2)   # ppShapeFormatPNG
            return [IO.File]::ReadAllBytes($file)
        } catch {
            return $null
        } finally {
            Remove-Item $file -ErrorAction SilentlyContinue
        }
    }

    $childBefore = Export-Shape $child1
    if ($childBefore) { $lines.Add("child render before group fill: $($childBefore.Length) bytes") }
    else { $lines.Add('child render before group fill: export unavailable') }

    # Now fill the group itself with the other image.
    $group.Fill.UserPicture($textureA)
    $lines.Add("group filled: group.Fill.Type=$($group.Fill.Type)")
    $lines.Add("children after group fill: child1.Fill.Type=$($child1.Fill.Type) child2.Fill.Type=$($child2.Fill.Type)")

    $childAfter = Export-Shape $child1
    if ($childBefore -and $childAfter) {
        $identical = $childBefore.Length -eq $childAfter.Length
        if ($identical) {
            for ($i = 0; $i -lt $childBefore.Length; $i++) {
                if ($childBefore[$i] -ne $childAfter[$i]) { $identical = $false; break }
            }
        }
        $lines.Add("child render after group fill: $($childAfter.Length) bytes")
        $lines.Add("CHILD RENDER UNCHANGED BY GROUP FILL: $identical")
    }


    # Ungroup, then ask each released Shape what it carries.
    $released = $group.Ungroup()
    foreach ($index in 1..$released.Count) {
        $item = $released.Item($index)
        $lines.Add("after ungroup: item $index Id=$($item.Id) Fill.Type=$($item.Fill.Type)")
    }
} finally {
    try { $presentation.Saved = $msoTrue; $presentation.Close() } catch { }
}

$lines | Set-Content "$root/artifacts/group_fill_propagation.txt"
$lines
