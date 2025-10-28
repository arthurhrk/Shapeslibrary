# Batch generate PNG previews from PPTX files
# Run this script to generate all preview images for the shape library

$ErrorActionPreference = "Stop"

# Get script directory
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent $scriptDir
$assetsDir = Join-Path $projectRoot "assets"

Write-Host "Starting preview generation..." -ForegroundColor Cyan
Write-Host "Assets directory: $assetsDir" -ForegroundColor Gray

# Categories to process
$categories = @("basic", "arrows", "flowchart", "callouts")

# Stats
$totalProcessed = 0
$totalSucceeded = 0
$totalFailed = 0

# Try to get existing PowerPoint instance or create new one
$createdApp = $false
try {
    $app = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
    Write-Host "Using existing PowerPoint instance" -ForegroundColor Green
} catch {
    $app = New-Object -ComObject PowerPoint.Application
    $createdApp = $true
    Write-Host "Created new PowerPoint instance" -ForegroundColor Green
}

try {
    foreach ($category in $categories) {
        $categoryDir = Join-Path $assetsDir $category

        if (-not (Test-Path $categoryDir)) {
            Write-Host "Skipping $category - directory not found" -ForegroundColor Yellow
            continue
        }

        Write-Host "`nProcessing category: $category" -ForegroundColor Cyan

        # Get all PPTX files in the category directory
        $pptxFiles = Get-ChildItem -Path $categoryDir -Filter "*.pptx"

        foreach ($pptxFile in $pptxFiles) {
            $totalProcessed++
            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($pptxFile.Name)
            $pngPath = Join-Path $categoryDir "$baseName.png"

            # Skip if PNG already exists
            if (Test-Path $pngPath) {
                Write-Host "  [SKIP] $baseName - PNG already exists" -ForegroundColor Gray
                $totalSucceeded++
                continue
            }

            try {
                Write-Host "  [GEN]  $baseName..." -NoNewline

                # Open presentation
                $pres = $app.Presentations.Open($pptxFile.FullName, $true, $false, $false)

                # Export first slide as PNG (1600x900)
                $slide = $pres.Slides.Item(1)
                $slide.Export($pngPath, 'PNG', 1600, 900)

                # Close presentation
                $pres.Close()

                Write-Host " OK" -ForegroundColor Green
                $totalSucceeded++
            } catch {
                Write-Host " FAILED" -ForegroundColor Red
                Write-Host "    Error: $($_.Exception.Message)" -ForegroundColor Red
                $totalFailed++

                # Try to close presentation if it's open
                try {
                    if ($pres) {
                        $pres.Close()
                    }
                } catch {}
            }
        }
    }
} finally {
    # Quit PowerPoint if we created it
    if ($createdApp) {
        $app.Quit()
        Write-Host "`nClosed PowerPoint" -ForegroundColor Gray
    }
}

# Summary
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Preview Generation Complete" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Total processed: $totalProcessed" -ForegroundColor White
Write-Host "Succeeded:       $totalSucceeded" -ForegroundColor Green
Write-Host "Failed:          $totalFailed" -ForegroundColor $(if ($totalFailed -gt 0) { "Red" } else { "White" })
Write-Host "========================================" -ForegroundColor Cyan
