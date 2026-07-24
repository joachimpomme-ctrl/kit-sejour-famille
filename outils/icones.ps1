# Génère les icônes PWA (marinière + badge lettre) dans wrapper/.
# Usage : powershell -ExecutionPolicy Bypass -File outils/icones.ps1 -Lettre "S"
param([string]$Lettre = "S")

Add-Type -AssemblyName System.Drawing
$dest = Join-Path $PSScriptRoot "..\wrapper"

foreach ($taille in 180, 192, 512) {
  $bmp = New-Object System.Drawing.Bitmap($taille, $taille)
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  $g.SmoothingMode = 'AntiAlias'
  $g.TextRenderingHint = 'AntiAliasGridFit'

  $craie = [System.Drawing.Color]::FromArgb(246, 244, 238)
  $marine = [System.Drawing.Color]::FromArgb(30, 51, 80)
  $g.Clear($craie)

  # Rayures marinière horizontales
  $bande = [Math]::Max(4, [int]($taille / 12))
  $pinceau = New-Object System.Drawing.SolidBrush($marine)
  for ($y = $bande; $y -lt $taille; $y += 2 * $bande) {
    $g.FillRectangle($pinceau, 0, $y, $taille, $bande)
  }

  # Badge rond central avec la lettre
  $d = [int]($taille * 0.56)
  $x = [int](($taille - $d) / 2)
  $g.FillEllipse($pinceau, $x, $x, $d, $d)
  $blanc = New-Object System.Drawing.SolidBrush($craie)
  $police = New-Object System.Drawing.Font('Arial Black', [int]($taille * 0.26), [System.Drawing.FontStyle]::Bold)
  $fmt = New-Object System.Drawing.StringFormat
  $fmt.Alignment = 'Center'; $fmt.LineAlignment = 'Center'
  $zone = New-Object System.Drawing.RectangleF($x, $x, $d, $d)
  $g.DrawString($Lettre, $police, $blanc, $zone, $fmt)

  $g.Dispose()
  $chemin = Join-Path $dest "icon-$taille.png"
  $bmp.Save($chemin, [System.Drawing.Imaging.ImageFormat]::Png)
  $bmp.Dispose()
  Write-Host "OK $chemin"
}
