param(
    [Parameter(Mandatory = $true)]
    [string]$Pptx,

    [string]$Output = "",

    [double]$SlackRatio = 0.22,

    [double]$MinFontSize = 5.5,

    [double]$OverflowTolerancePt = 1.0,

    [switch]$PreferShrink
)

$ErrorActionPreference = "Stop"

function Is-DecorativeText($text) {
    if ($null -eq $text) { return $true }
    $t = (($text -replace "`r", " ") -replace "`n", " ").Trim()
    if ($t.Length -le 1) { return $true }
    if ($t -match '^[+\-−!★·…⋮0-9.]+$') { return $true }
    if ($t -match '^[Σ∑logpi\(\)\+=ᵢ\s.]+$') { return $true }
    return $false
}

function Repair-TextShape($shape, $slideWidth, $slideHeight) {
    try {
        if (-not $shape.HasTextFrame -or -not $shape.TextFrame2.HasText) { return @{changed=$false; reason="no_text"} }
        $text = $shape.TextFrame2.TextRange.Text
        if (Is-DecorativeText $text) { return @{changed=$false; reason="decorative"} }

        $changed = $false
        $shape.TextFrame2.MarginLeft = 0
        $shape.TextFrame2.MarginRight = 0
        $shape.TextFrame2.MarginTop = 0
        $shape.TextFrame2.MarginBottom = 0

        $boundW = [double]$shape.TextFrame2.TextRange.BoundWidth
        $boundH = [double]$shape.TextFrame2.TextRange.BoundHeight
        $needW = [Math]::Max($shape.Width, $boundW * (1.0 + $SlackRatio))
        $needH = [Math]::Max($shape.Height, $boundH * 1.12)

        if (-not $PreferShrink) {
            $maxW = [Math]::Max(1.0, $slideWidth - $shape.Left - 2)
            $maxH = [Math]::Max(1.0, $slideHeight - $shape.Top - 2)
            if ($needW -gt $shape.Width -and $needW -le $maxW) {
                $shape.Width = $needW
                $changed = $true
            }
            if ($needH -gt $shape.Height -and $needH -le $maxH) {
                $shape.Height = $needH
                $changed = $true
            }
        }

        $font = $shape.TextFrame2.TextRange.Font
        while (($shape.TextFrame2.TextRange.BoundWidth -gt ($shape.Width - $OverflowTolerancePt) -or $shape.TextFrame2.TextRange.BoundHeight -gt ($shape.Height - $OverflowTolerancePt)) -and $font.Size -gt $MinFontSize) {
            $font.Size = $font.Size - 0.5
            $changed = $true
        }
        return @{changed=$changed; reason="processed"}
    } catch {
        return @{changed=$false; reason=$_.Exception.Message}
    }
}

function Repair-ShapeCollection($shapes, $slideWidth, $slideHeight) {
    $count = 0
    $changed = 0
    for ($i = 1; $i -le $shapes.Count; $i++) {
        $shape = $shapes.Item($i)
        try {
            if ([int]$shape.Type -eq 6 -and $shape.GroupItems.Count -gt 0) {
                $sub = Repair-ShapeCollection $shape.GroupItems $slideWidth $slideHeight
                $count += $sub.count
                $changed += $sub.changed
                continue
            }
        } catch {}
        $count += 1
        $r = Repair-TextShape $shape $slideWidth $slideHeight
        if ($r.changed) { $changed += 1 }
    }
    return @{count=$count; changed=$changed}
}

$resolved = (Resolve-Path -LiteralPath $Pptx).Path
if ($Output -eq "") {
    $p = [IO.Path]::GetDirectoryName($resolved)
    $n = [IO.Path]::GetFileNameWithoutExtension($resolved)
    $Output = Join-Path $p ($n + "_textsafe.pptx")
}

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = -1
$pres = $ppt.Presentations.Open($resolved, $false, $false, $false)

$total = 0
$changed = 0
try {
    for ($s = 1; $s -le $pres.Slides.Count; $s++) {
        $slide = $pres.Slides.Item($s)
        $r = Repair-ShapeCollection $slide.Shapes $pres.PageSetup.SlideWidth $pres.PageSetup.SlideHeight
        $total += $r.count
        $changed += $r.changed
    }
    $target = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Output)
    $parent = Split-Path -Parent $target
    if ($parent -and !(Test-Path -LiteralPath $parent)) { New-Item -ItemType Directory -Force -Path $parent | Out-Null }
    $pres.SaveAs($target, 24)
} finally {
    $pres.Close()
}

@{
    status = "repaired"
    input = $resolved
    output = $target
    scanned_shapes = $total
    changed_text_shapes = $changed
    slack_ratio = $SlackRatio
} | ConvertTo-Json -Depth 5
