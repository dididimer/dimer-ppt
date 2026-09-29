[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string]$Pptx,

    [Parameter(Mandatory = $true, Position = 1)]
    [string]$OutDir,

    [ValidateRange(1, 16384)]
    [int]$Width = 1920,

    [ValidateRange(1, 16384)]
    [int]$Height = 1080,

    [ValidateRange(45, 3600)]
    [int]$TimeoutSeconds = 180,

    [switch]$Overwrite
)

$ErrorActionPreference = "Stop"

function Release-ComObject($object) {
    if ($null -ne $object -and [Runtime.InteropServices.Marshal]::IsComObject($object)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($object)
    }
}

$timer = [Diagnostics.Stopwatch]::StartNew()
function Assert-WithinTimeout([string]$stage) {
    if ($timer.Elapsed.TotalSeconds -gt $TimeoutSeconds) {
        throw "Timed out after $TimeoutSeconds seconds while $stage. PowerPoint can take 45 seconds or more to initialize; retry with -TimeoutSeconds 300 if startup is still in progress."
    }
}

$presentation = $null
$application = $null
$createdApplication = $false

try {
    $resolvedPptx = (Resolve-Path -LiteralPath $Pptx).Path
    $resolvedOutDir = [IO.Path]::GetFullPath($OutDir)
    [IO.Directory]::CreateDirectory($resolvedOutDir) | Out-Null
    Assert-WithinTimeout "resolving input and output paths"

    try {
        $application = [Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application")
        Write-Verbose "Using the existing PowerPoint application instance."
    } catch {
        Write-Verbose "Starting a private PowerPoint application instance."
        $application = New-Object -ComObject PowerPoint.Application
        $createdApplication = $true
    }
    Assert-WithinTimeout "initializing PowerPoint"

    # Open read-only and without a window. This presentation is always closed in finally;
    # an already-open user presentation and an existing PowerPoint process are left alone.
    $presentation = $application.Presentations.Open($resolvedPptx, -1, 0, 0)
    Assert-WithinTimeout "opening the presentation"

    $baseName = [IO.Path]::GetFileNameWithoutExtension($resolvedPptx)
    $slideCount = [int]$presentation.Slides.Count
    if ($slideCount -lt 1) {
        throw "The presentation contains no slides: $resolvedPptx"
    }

    $outputs = @()
    for ($index = 1; $index -le $slideCount; $index++) {
        Assert-WithinTimeout "exporting slide $index of $slideCount"
        $outputPath = Join-Path $resolvedOutDir ("{0}-slide-{1:D3}.png" -f $baseName, $index)
        if ((Test-Path -LiteralPath $outputPath) -and -not $Overwrite) {
            throw "Refusing to overwrite existing output: $outputPath. Use -Overwrite to replace it."
        }
        if ($Overwrite -and (Test-Path -LiteralPath $outputPath)) {
            [IO.File]::Delete($outputPath)
        }

        $presentation.Slides.Item($index).Export($outputPath, "PNG", $Width, $Height)
        if (-not (Test-Path -LiteralPath $outputPath -PathType Leaf)) {
            throw "PowerPoint returned without creating the PNG for slide ${index}: $outputPath"
        }
        $outputs += $outputPath
        Write-Host ("Exported slide {0}/{1}: {2}" -f $index, $slideCount, $outputPath)
    }
    Assert-WithinTimeout "completing PNG export"

    [pscustomobject]@{
        pptx = $resolvedPptx
        output_dir = $resolvedOutDir
        slides = $slideCount
        width_px = $Width
        height_px = $Height
        elapsed_seconds = [Math]::Round($timer.Elapsed.TotalSeconds, 1)
        outputs = $outputs
    } | ConvertTo-Json -Depth 3
} catch {
    $elapsed = [Math]::Round($timer.Elapsed.TotalSeconds, 1)
    Write-Error ("PowerPoint PNG export failed after {0}s: {1}" -f $elapsed, $_.Exception.Message)
    exit 1
} finally {
    if ($null -ne $presentation) {
        try { $presentation.Close() } catch { Write-Warning "Could not close the presentation opened by this script: $($_.Exception.Message)" }
        Release-ComObject $presentation
    }
    if ($null -ne $application) {
        if ($createdApplication) {
            try { $application.Quit() } catch { Write-Warning "Could not quit the private PowerPoint instance: $($_.Exception.Message)" }
        }
        Release-ComObject $application
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
