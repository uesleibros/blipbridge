<#
.SYNOPSIS
Does a group child yield a complete, collision-free cache key?

.DESCRIPTION
The per-Shape skip cache is keyed by presentation + SlideID + Shape.Id. Whether a
group child can be keyed safely was assumed rather than measured, so this
measures it:

  * what is a group child's Shape.Parent - the slide, or the group?
  * does it report a SlideID, so a complete key exists at all?
  * are group children's Ids distinct from every top-level Shape's Id on the
    same slide, and from each other?
  * do Ids survive grouping and ungrouping?

Read-only apart from creating its own Shapes. Nothing here applies a fill.
#>
$ErrorActionPreference = 'Stop'
$msoTrue = -1
$ppLayoutBlank = 12

$app = New-Object -ComObject PowerPoint.Application
$presentation = $app.Presentations.Add($msoTrue)
$lines = New-Object System.Collections.Generic.List[string]
try {
    $slide = $presentation.Slides.Add(1, $ppLayoutBlank)
    $loose = @()
    for ($i = 0; $i -lt 3; $i++) {
        $loose += $slide.Shapes.AddShape(1, 20 + $i * 90, 20, 70, 70)
    }
    $g1 = $slide.Shapes.AddShape(1, 20, 140, 60, 60)
    $g2 = $slide.Shapes.AddShape(1, 100, 140, 60, 60)
    $g1Id = $g1.Id
    $g2Id = $g2.Id
    $lines.Add("before grouping: child ids $g1Id, $g2Id")

    $group = $slide.Shapes.Range(@($g1.Name, $g2.Name)).Group()
    $lines.Add("group id $($group.Id), type $($group.Type)")

    foreach ($index in 1..$group.GroupItems.Count) {
        $child = $group.GroupItems.Item($index)
        $parentKind = '<unreadable>'
        $slideId = '<none>'
        try { $parentKind = $child.Parent.Name } catch { }
        try { $slideId = $child.Parent.SlideID } catch { $slideId = '<no SlideID>' }
        $parentGroup = '<none>'
        try { $parentGroup = $child.ParentGroup.Id } catch { }
        $lines.Add("child $index : Id=$($child.Id) Parent=$parentKind SlideID=$slideId ParentGroup=$parentGroup")
    }

    # Every Id visible on the slide, top-level and nested, looking for collisions.
    $ids = New-Object System.Collections.Generic.List[string]
    foreach ($top in $slide.Shapes) {
        $ids.Add("top:$($top.Id)")
        if ($top.Type -eq 6) {
            foreach ($member in $top.GroupItems) { $ids.Add("child:$($member.Id)") }
        }
    }
    $lines.Add("all ids: $($ids -join ' ')")
    $numbers = $ids | ForEach-Object { ($_ -split ':')[1] }
    $distinct = @($numbers | Select-Object -Unique).Count
    $lines.Add("ids total $($numbers.Count), distinct $distinct")
    $lines.Add("COLLISION: $($distinct -ne $numbers.Count)")

    # Do Ids survive ungrouping?
    $released = $group.Ungroup()
    $after = @()
    foreach ($item in $released) { $after += $item.Id }
    $lines.Add("after ungroup: ids $($after -join ', ') (were $g1Id, $g2Id)")
} finally {
    try { $presentation.Saved = $msoTrue; $presentation.Close() } catch { }
}

$root = Split-Path $PSScriptRoot -Parent
$lines | Set-Content "$root/artifacts/shape_identity.txt"
$lines
