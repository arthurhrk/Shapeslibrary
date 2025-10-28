Param(
  [string]$OutputPath = "icon.png",
  [int]$Size = 512
)

Add-Type -AssemblyName System.Drawing

# Helper: Rounded rectangle path
function New-RoundedRectPath {
  param(
    [float]$x, [float]$y, [float]$w, [float]$h, [float]$r
  )
  $gp = New-Object System.Drawing.Drawing2D.GraphicsPath
  $d = [float]($r * 2)
  $gp.AddArc($x, $y, $d, $d, 180, 90)
  $gp.AddArc($x + $w - $d, $y, $d, $d, 270, 90)
  $gp.AddArc($x + $w - $d, $y + $h - $d, $d, $d, 0, 90)
  $gp.AddArc($x, $y + $h - $d, $d, $d, 90, 90)
  $gp.CloseFigure()
  return $gp
}

$bmp = New-Object -TypeName System.Drawing.Bitmap -ArgumentList @([int]$Size, [int]$Size, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
$gfx = [System.Drawing.Graphics]::FromImage([System.Drawing.Image]$bmp)
$gfx.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$gfx.Clear([System.Drawing.Color]::Transparent)

# Colors (PowerPoint palette inspired)
$centerColor = [System.Drawing.ColorTranslator]::FromHtml('#FFA24A')  # bright orange
$edgeColor   = [System.Drawing.ColorTranslator]::FromHtml('#D24726')  # deep PowerPoint orange/red
$shadowColor = [System.Drawing.Color]::FromArgb(70, 0, 0, 0)
$white       = [System.Drawing.Brushes]::White

# Circular background with radial gradient
$pad = [int]([math]::Round($Size * 0.03125)) # ~16 for 512
$circleRect = New-Object -TypeName System.Drawing.RectangleF -ArgumentList @([single]$pad, [single]$pad, [single]($Size - 2*$pad), [single]($Size - 2*$pad))
$circlePath = New-Object -TypeName System.Drawing.Drawing2D.GraphicsPath
$circlePath.AddEllipse($circleRect)
$pBrush = New-Object -TypeName System.Drawing.Drawing2D.PathGradientBrush -ArgumentList @($circlePath)
$pBrush.CenterPoint = New-Object -TypeName System.Drawing.PointF -ArgumentList @([single]($Size/2), [single]($Size/2))
$pBrush.CenterColor = $centerColor
$pBrush.SurroundColors = @($edgeColor)

# Optional soft outer shadow
$shadowOffset = [int]([math]::Round($Size * 0.01))
$shadowRect = New-Object -TypeName System.Drawing.RectangleF -ArgumentList @([single]($circleRect.X + $shadowOffset), [single]($circleRect.Y + $shadowOffset), [single]$circleRect.Width, [single]$circleRect.Height)
$shadowPath = New-Object -TypeName System.Drawing.Drawing2D.GraphicsPath
$shadowPath.AddEllipse($shadowRect)
$gfx.FillPath((New-Object -TypeName System.Drawing.SolidBrush -ArgumentList @($shadowColor)), $shadowPath)

$gfx.FillPath($pBrush, $circlePath)

# White shapes inside (circle, triangle, rounded-rect)
$shapeScale = $Size / 512.0

# Circle (bottom-left)
$cSize = 110 * $shapeScale
$cX = 140 * $shapeScale
$cY = 320 * $shapeScale
$gfx.FillEllipse((New-Object -TypeName System.Drawing.SolidBrush -ArgumentList @([System.Drawing.Color]::FromArgb(220,255,255,255))), [single]($cX+3), [single]($cY+3), [single]$cSize, [single]$cSize) # shadow
$gfx.FillEllipse($white, $cX, $cY, $cSize, $cSize)

# Triangle (top-left)
$t1 = New-Object -TypeName System.Drawing.PointF -ArgumentList @([single](180 * $shapeScale), [single](110 * $shapeScale))
$t2 = New-Object -TypeName System.Drawing.PointF -ArgumentList @([single](110 * $shapeScale), [single](240 * $shapeScale))
$t3 = New-Object -TypeName System.Drawing.PointF -ArgumentList @([single](250 * $shapeScale), [single](240 * $shapeScale))
$triPointsShadow = @([System.Drawing.PointF]::new([single]($t1.X+3),[single]($t1.Y+3)),[System.Drawing.PointF]::new([single]($t2.X+3),[single]($t2.Y+3)),[System.Drawing.PointF]::new([single]($t3.X+3),[single]($t3.Y+3)))
$triPoints = @($t1,$t2,$t3)
$gfx.FillPolygon((New-Object -TypeName System.Drawing.SolidBrush -ArgumentList @([System.Drawing.Color]::FromArgb(220,255,255,255))), $triPointsShadow)
$gfx.FillPolygon($white, $triPoints)

# Rounded rectangle (mid-right)
$rrX = 290 * $shapeScale
$rrY = 190 * $shapeScale
$rrW = 140 * $shapeScale
$rrH = 100 * $shapeScale
$rrR = 26 * $shapeScale
$rrShadow = New-RoundedRectPath -x ([single]($rrX+3)) -y ([single]($rrY+3)) -w ([single]$rrW) -h ([single]$rrH) -r ([single]$rrR)
$gfx.FillPath((New-Object -TypeName System.Drawing.SolidBrush -ArgumentList @([System.Drawing.Color]::FromArgb(220,255,255,255))), $rrShadow)
$rrPath = New-RoundedRectPath -x $rrX -y $rrY -w $rrW -h $rrH -r $rrR
$gfx.FillPath($white, $rrPath)

# Save PNG
$fullPath = if ([System.IO.Path]::IsPathRooted($OutputPath)) { $OutputPath } else { Join-Path -Path (Get-Location) -ChildPath $OutputPath }
$bmp.Save($fullPath, [System.Drawing.Imaging.ImageFormat]::Png)

# Cleanup
$rrPath.Dispose()
$rrShadow.Dispose()
$circlePath.Dispose()
$shadowPath.Dispose()
$pBrush.Dispose()
$gfx.Dispose()
$bmp.Dispose()

Write-Host "Icon generated at: $fullPath"
