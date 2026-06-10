param(
    [Parameter(Mandatory = $true)]
    [string]$Pptx,

    [string]$Report = "",

    [switch]$FixOverflow,

    [double]$MinFontSize = 5.5,

    [double]$OverflowTolerancePt = 1.5,

    [double]$OverlapToleranceRatio = 0.18
)

$ErrorActionPreference = "Stop"

function ConvertTo-PlainText($value) {
    if ($null -eq $value) { return "" }
    return (($value -replace "`r", " ") -replace "`n", " ").Trim()
}

function Is-DecorativeText($text) {
    $t = ConvertTo-PlainText $text
    if ($t.Length -le 1) { return $true }
    if ($t -match '^[+\-−!★·…⋮0-9.]+$') { return $true }
    if ($t -match '^[Σ∑logpi\(\)\+=ᵢ\s.]+$') { return $true }
    return $false
}

function Box($shape) {
    return @{
        left = [double]$shape.Left
        top = [double]$shape.Top
        width = [double]$shape.Width
        height = [double]$shape.Height
        right = [double]($shape.Left + $shape.Width)
        bottom = [double]($shape.Top + $shape.Height)
    }
}

function IntersectArea($a, $b) {
    $x1 = [Math]::Max($a.left, $b.left)
    $y1 = [Math]::Max($a.top, $b.top)
    $x2 = [Math]::Min($a.right, $b.right)
    $y2 = [Math]::Min($a.bottom, $b.bottom)
    if ($x2 -le $x1 -or $y2 -le $y1) { return 0.0 }
    return [double](($x2 - $x1) * ($y2 - $y1))
}

function ShapeInfo($slideIndex, $shape) {
    $box = Box $shape
    $hasText = $false
    $text = ""
    $boundW = 0.0
    $boundH = 0.0
    try {
        if ($shape.HasTextFrame -and $shape.TextFrame2.HasText) {
            $text = ConvertTo-PlainText $shape.TextFrame2.TextRange.Text
            if ($text.Length -gt 0) {
                $hasText = $true
                $boundW = [double]$shape.TextFrame2.TextRange.BoundWidth
                $boundH = [double]$shape.TextFrame2.TextRange.BoundHeight
            }
        }
    } catch {}
    return @{
        slide = $slideIndex
        id = [int]$shape.Id
        name = [string]$shape.Name
        type = [int]$shape.Type
        has_text = $hasText
        text = $text
        decorative_text = (Is-DecorativeText $text)
        left = $box.left
        top = $box.top
        width = $box.width
        height = $box.height
        right = $box.right
        bottom = $box.bottom
        text_bound_width = $boundW
        text_bound_height = $boundH
    }
}

function Collect-ShapeInfos($slideIndex, $shapes, $parentName = "") {
    $items = @()
    for ($i = 1; $i -le $shapes.Count; $i++) {
        $shape = $shapes.Item($i)
        $info = ShapeInfo $slideIndex $shape
        if ($parentName -ne "") {
            $info.parent = $parentName
            $info.name = "$parentName/$($info.name)"
            $info.top_index = 0
        } else {
            $info.parent = ""
            $info.top_index = $i
        }
        $items += $info
        try {
            if ([int]$shape.Type -eq 6 -and $shape.GroupItems.Count -gt 0) {
                $items += Collect-ShapeInfos $slideIndex $shape.GroupItems $info.name
            }
        } catch {}
    }
    return $items
}

$resolved = (Resolve-Path -LiteralPath $Pptx).Path
$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = -1
$pres = $ppt.Presentations.Open($resolved, $false, $false, $false)

$issues = @()
$slidesOut = @()

try {
    for ($s = 1; $s -le $pres.Slides.Count; $s++) {
        $slide = $pres.Slides.Item($s)
        $infos = Collect-ShapeInfos $s $slide.Shapes
        for ($i = 0; $i -lt $infos.Count; $i++) {
            $info = $infos[$i]
            if ($info.has_text -and -not $info.decorative_text) {
                $shape = $null
                $overflowW = $info.text_bound_width - $info.width
                $overflowH = $info.text_bound_height - $info.height
                if ($overflowW -gt $OverflowTolerancePt -or $overflowH -gt $OverflowTolerancePt) {
                    $issues += @{
                        kind = "overflow"
                        slide = $s
                        name = $info.name
                        text = $info.text
                        box = @{
                            left = [Math]::Round($info.left, 2)
                            top = [Math]::Round($info.top, 2)
                            width = [Math]::Round($info.width, 2)
                            height = [Math]::Round($info.height, 2)
                            right = [Math]::Round($info.right, 2)
                            bottom = [Math]::Round($info.bottom, 2)
                        }
                        width = [Math]::Round($info.width, 2)
                        height = [Math]::Round($info.height, 2)
                        bound_width = [Math]::Round($info.text_bound_width, 2)
                        bound_height = [Math]::Round($info.text_bound_height, 2)
                    }
                    if ($FixOverflow) {
                        try {
                            # Best-effort: only fix top-level shapes. Grouped shapes are reported
                            # for manifest/regeneration fixes because editing them by path is fragile.
                            if ($info.parent -ne "") { continue }
                            $shape = $slide.Shapes.Item([int]$info.top_index)
                            $font = $shape.TextFrame2.TextRange.Font
                            while (($shape.TextFrame2.TextRange.BoundWidth -gt ($shape.Width - $OverflowTolerancePt) -or $shape.TextFrame2.TextRange.BoundHeight -gt ($shape.Height - $OverflowTolerancePt)) -and $font.Size -gt $MinFontSize) {
                                $font.Size = $font.Size - 0.5
                            }
                        } catch {}
                    }
                }
            }
        }

        for ($a = 0; $a -lt $infos.Count; $a++) {
            $ia = $infos[$a]
            if (-not $ia.has_text -or $ia.decorative_text) { continue }
            $boxA = @{ left=$ia.left; top=$ia.top; right=$ia.right; bottom=$ia.bottom; width=$ia.width; height=$ia.height }
            for ($b = $a + 1; $b -lt $infos.Count; $b++) {
                $ib = $infos[$b]
                $boxB = @{ left=$ib.left; top=$ib.top; right=$ib.right; bottom=$ib.bottom; width=$ib.width; height=$ib.height }
                $area = IntersectArea $boxA $boxB
                if ($area -le 0) { continue }
                $minArea = [Math]::Max(1.0, [Math]::Min($ia.width * $ia.height, $ib.width * $ib.height))
                $ratio = $area / $minArea
                if ($ratio -lt $OverlapToleranceRatio) { continue }

                if ($ib.has_text -and -not $ib.decorative_text) {
                    $issues += @{
                        kind = "text_text_overlap"
                        slide = $s
                        a = $ia.name
                        b = $ib.name
                        a_text = $ia.text
                        b_text = $ib.text
                        a_box = @{
                            left = [Math]::Round($ia.left, 2)
                            top = [Math]::Round($ia.top, 2)
                            width = [Math]::Round($ia.width, 2)
                            height = [Math]::Round($ia.height, 2)
                            right = [Math]::Round($ia.right, 2)
                            bottom = [Math]::Round($ia.bottom, 2)
                        }
                        b_box = @{
                            left = [Math]::Round($ib.left, 2)
                            top = [Math]::Round($ib.top, 2)
                            width = [Math]::Round($ib.width, 2)
                            height = [Math]::Round($ib.height, 2)
                            right = [Math]::Round($ib.right, 2)
                            bottom = [Math]::Round($ib.bottom, 2)
                        }
                        ratio = [Math]::Round($ratio, 3)
                    }
                } elseif (-not $ib.has_text) {
                    $isPanel = ($ib.name -match 'PANEL|RECT|BACKGROUND|CARD|BOX')
                    if (-not $isPanel -and $ratio -ge $OverlapToleranceRatio) {
                        $issues += @{
                            kind = "text_shape_overlap"
                            slide = $s
                            text_shape = $ia.name
                            other_shape = $ib.name
                            text = $ia.text
                            text_box = @{
                                left = [Math]::Round($ia.left, 2)
                                top = [Math]::Round($ia.top, 2)
                                width = [Math]::Round($ia.width, 2)
                                height = [Math]::Round($ia.height, 2)
                                right = [Math]::Round($ia.right, 2)
                                bottom = [Math]::Round($ia.bottom, 2)
                            }
                            other_box = @{
                                left = [Math]::Round($ib.left, 2)
                                top = [Math]::Round($ib.top, 2)
                                width = [Math]::Round($ib.width, 2)
                                height = [Math]::Round($ib.height, 2)
                                right = [Math]::Round($ib.right, 2)
                                bottom = [Math]::Round($ib.bottom, 2)
                            }
                            ratio = [Math]::Round($ratio, 3)
                        }
                    }
                }
            }
        }
        $slidesOut += @{ slide = $s; shape_count = $slide.Shapes.Count; text_shape_count = @($infos | Where-Object { $_.has_text }).Count }
    }

    if ($FixOverflow) {
        $pres.Save()
    }
} finally {
    $pres.Close()
}

$payload = @{
    ok = ($issues.Count -eq 0)
    file = $resolved
    fixed_overflow = [bool]$FixOverflow
    issue_count = $issues.Count
    slides = $slidesOut
    issues = $issues
}

$json = $payload | ConvertTo-Json -Depth 8
if ($Report -ne "") {
    $parent = Split-Path -Parent $Report
    if ($parent -and !(Test-Path -LiteralPath $parent)) { New-Item -ItemType Directory -Force -Path $parent | Out-Null }
    Set-Content -LiteralPath $Report -Value $json -Encoding UTF8
}
$json
