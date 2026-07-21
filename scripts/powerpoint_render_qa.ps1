param(
    [Parameter(Mandatory = $true)]
    [string]$PptxPath,

    [Parameter(Mandatory = $true)]
    [string]$PreviewPath,

    [Parameter(Mandatory = $true)]
    [string]$JsonPath,

    [int]$RenderWidth = 1600,
    [int]$RenderHeight = 0,
    [int]$LayoutWaitSeconds = 3
)

$ErrorActionPreference = 'Stop'
$resolvedPptx = (Resolve-Path -LiteralPath $PptxPath).Path
$previewFull = [System.IO.Path]::GetFullPath($PreviewPath)
$jsonFull = [System.IO.Path]::GetFullPath($JsonPath)
$previewDir = Split-Path -Parent $previewFull
$jsonDir = Split-Path -Parent $jsonFull
New-Item -ItemType Directory -Force -Path $previewDir, $jsonDir | Out-Null

$app = $null
$presentation = $null

try {
    $app = New-Object -ComObject PowerPoint.Application
    $presentation = $app.Presentations.Open($resolvedPptx, $true, $false, $false)
    Start-Sleep -Seconds $LayoutWaitSeconds

    $actualRenderHeight = $RenderHeight
    if ($actualRenderHeight -le 0) {
        $actualRenderHeight = [math]::Round(
            $RenderWidth * $presentation.PageSetup.SlideHeight / $presentation.PageSetup.SlideWidth
        )
    }

    $overflow = @()
    $pictures = @()
    $textCount = 0
    $shapeCount = 0
    $previews = @()

    for ($slideIndex = 1; $slideIndex -le $presentation.Slides.Count; $slideIndex++) {
        $slide = $presentation.Slides.Item($slideIndex)
        if ($presentation.Slides.Count -eq 1) {
            $slidePreview = $previewFull
        }
        else {
            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($previewFull)
            $extension = [System.IO.Path]::GetExtension($previewFull)
            $slidePreview = Join-Path $previewDir ("{0}_slide{1}{2}" -f $baseName, $slideIndex, $extension)
        }
        $slide.Export($slidePreview, 'PNG', $RenderWidth, $actualRenderHeight)
        $previews += $slidePreview

        foreach ($shape in $slide.Shapes) {
            $shapeCount++
            if ($shape.Type -eq 13) {
                $pictures += "slide${slideIndex}:$($shape.Name)"
            }
            try {
                if ($shape.HasTextFrame -and $shape.TextFrame2.HasText) {
                    $textCount++
                    $boundWidth = $shape.TextFrame2.TextRange.BoundWidth
                    $boundHeight = $shape.TextFrame2.TextRange.BoundHeight
                    if ($boundWidth -gt ($shape.Width + 1) -or $boundHeight -gt ($shape.Height + 1)) {
                        $overflow += [ordered]@{
                            slide = $slideIndex
                            shape = $shape.Name
                            bound_width = [math]::Round($boundWidth, 2)
                            bound_height = [math]::Round($boundHeight, 2)
                            box_width = [math]::Round($shape.Width, 2)
                            box_height = [math]::Round($shape.Height, 2)
                        }
                    }
                }
            }
            catch {
                # Some connector and placeholder types do not expose TextFrame2 consistently.
            }
        }
    }

    $passed = ($pictures.Count -eq 0 -and $overflow.Count -eq 0)
    $result = [ordered]@{
        pptx = $resolvedPptx
        slides = $presentation.Slides.Count
        shapes = $shapeCount
        text_shapes = $textCount
        picture_shapes = $pictures
        text_overflow = $overflow
        previews = $previews
        render_width = $RenderWidth
        render_height = $actualRenderHeight
        passed = $passed
        blocked = $false
    }
    $result | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $jsonFull -Encoding UTF8
    $result | ConvertTo-Json -Depth 8
    if (-not $passed) {
        exit 2
    }
}
catch {
    $result = [ordered]@{
        pptx = $resolvedPptx
        passed = $false
        blocked = $true
        error = $_.Exception.Message
    }
    $result | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $jsonFull -Encoding UTF8
    $result | ConvertTo-Json -Depth 4
    exit 3
}
finally {
    if ($null -ne $presentation) {
        $presentation.Close()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) | Out-Null
    }
    if ($null -ne $app) {
        $app.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
