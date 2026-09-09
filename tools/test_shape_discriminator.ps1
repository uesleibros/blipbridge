<#
.SYNOPSIS
Finds a public property that separates fillable Shapes from ones that only look fillable.

.DESCRIPTION
A Connector reports Shape.Type = 1 (msoAutoShape) and presents exactly the same
PPCORE wrapper, OART FillFormat and receiver as a rectangle - so neither the type
check nor the structural walk can tell them apart. Applying to one crashes
PowerPoint, and Office's own Fill.UserPicture refuses it.

Something public must distinguish them, because Office distinguishes them. This
reads a handful of candidate properties across every class that can be created,
so the choice of guard is made from a table rather than from a guess.

Nothing here applies a fill. It only reads.
#>
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$msoTrue = -1
$ppLayoutBlank = 12
$ppLayoutText = 2
$texture = Join-Path $root 'artifacts/textures/texture_128_0.png'

function Read-Property($shape, [string]$name) {
    try {
        $value = $shape.$name
        if ($null -eq $value) { return 'null' }
        return "$value"
    } catch {
        return 'throws'
    }
}

$app = New-Object -ComObject PowerPoint.Application
$presentation = $app.Presentations.Add($msoTrue)
$rows = New-Object System.Collections.Generic.List[object]
try {
    $slide = $presentation.Slides.Add(1, $ppLayoutBlank)
    $textSlide = $presentation.Slides.Add(2, $ppLayoutText)

    $subjects = @(
        @{ Name = 'AutoShape';   Make = { $slide.Shapes.AddShape(1, 20, 20, 100, 100) } }
        @{ Name = 'Freeform';    Make = {
                $b = $slide.Shapes.BuildFreeform(1, 200, 20)
                $b.AddNodes(1, 1, 300, 20); $b.AddNodes(1, 1, 300, 120); $b.AddNodes(1, 1, 200, 20)
                $b.ConvertToShape() } }
        @{ Name = 'Connector';   Make = { $slide.Shapes.AddConnector(1, 360, 20, 460, 120) } }
        @{ Name = 'Line';        Make = { $slide.Shapes.AddLine(360, 160, 460, 220) } }
        @{ Name = 'TextBox';     Make = { $slide.Shapes.AddTextbox(1, 20, 160, 160, 50) } }
        @{ Name = 'Callout';     Make = { $slide.Shapes.AddCallout(2, 20, 240, 140, 60) } }
        @{ Name = 'Placeholder'; Make = { $textSlide.Shapes.Placeholders.Item(1) } }
        @{ Name = 'Picture';     Make = { $slide.Shapes.AddPicture($texture, $false, $true, 500, 20, 60, 60) } }
        @{ Name = 'WordArt';     Make = { $slide.Shapes.AddTextEffect(1, 'BB', 'Arial', 20, $false, $false, 200, 240) } }
        @{ Name = 'Table';       Make = { $slide.Shapes.AddTable(2, 2, 20, 320, 180, 70) } }
        @{ Name = 'Group';       Make = {
                $a = $slide.Shapes.AddShape(1, 560, 200, 40, 40)
                $b = $slide.Shapes.AddShape(1, 560, 250, 40, 40)
                $slide.Shapes.Range(@($a.Name, $b.Name)).Group() } }
    )

    foreach ($subject in $subjects) {
        $shape = $null
        try { $shape = & $subject.Make } catch {
            $rows.Add([pscustomobject]@{ Name = $subject.Name; Type = 'create failed' })
            continue
        }
        $rows.Add([pscustomobject]@{
            Name          = $subject.Name
            Type          = Read-Property $shape 'Type'
            Connector     = Read-Property $shape 'Connector'
            AutoShapeType = Read-Property $shape 'AutoShapeType'
            FillVisible   = Read-Property $shape.Fill 'Visible'
            FillType      = Read-Property $shape.Fill 'Type'
            HasChart      = Read-Property $shape 'HasChart'
            HasTable      = Read-Property $shape 'HasTable'
            HasSmartArt   = Read-Property $shape 'HasSmartArt'
        })
    }
} finally {
    try { $presentation.Saved = $msoTrue; $presentation.Close() } catch { }
}

$rows | Format-Table -AutoSize
$rows | ForEach-Object {
    '{0}|Type={1}|Connector={2}|AutoShapeType={3}|FillVisible={4}|FillType={5}' -f `
        $_.Name, $_.Type, $_.Connector, $_.AutoShapeType, $_.FillVisible, $_.FillType
} | Set-Content "$root/artifacts/shape_discriminator.txt"
