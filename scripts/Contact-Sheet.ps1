<# .SYNOPSIS Create a source-bound contact sheet of actual slide PNGs on Windows. #>
param([Parameter(Mandatory=$true)][string]$PreviewDir,[Parameter(Mandatory=$true)][string]$OutputPath,[int]$Columns=4,[int]$ThumbWidth=400)
$ErrorActionPreference='Stop'
if($PSVersionTable.PSVersion.Major -lt 7){throw 'PowerShell 7 required'}
if(Test-Path -LiteralPath $OutputPath){throw 'Output exists; use a new filename'}
Add-Type -AssemblyName System.Drawing
$dir=(Resolve-Path -LiteralPath $PreviewDir).Path
$r=Get-Content -LiteralPath (Join-Path $dir 'render-report.json') -Raw | ConvertFrom-Json
if($r.slides.Count -eq 0){throw 'No rendered slides'}
if($Columns -lt 1 -or $ThumbWidth -lt 100){throw 'Invalid contact sheet dimensions'}
$first=[Drawing.Image]::FromFile((Join-Path $dir $r.slides[0].file))
$thumbHeight=[int][math]::Round($ThumbWidth*$first.Height/$first.Width);$first.Dispose()
$rows=[int][math]::Ceiling($r.slides.Count/[double]$Columns);$gap=8;$label=24
$bmp=[Drawing.Bitmap]::new($Columns*($ThumbWidth+$gap)+$gap,$rows*($thumbHeight+$label+$gap)+$gap)
$g=[Drawing.Graphics]::FromImage($bmp);$font=[Drawing.Font]::new('Microsoft YaHei',11)
try{
 $g.Clear([Drawing.Color]::FromArgb(230,235,240));$g.InterpolationMode=[Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
 for($i=0;$i -lt $r.slides.Count;$i++){
  $e=$r.slides[$i];$p=Join-Path $dir $e.file
  if((Get-FileHash -LiteralPath $p -Algorithm SHA256).Hash.ToLower() -ne $e.sha256){throw 'Rendered image changed'}
  $im=[Drawing.Image]::FromFile($p)
  try{
   $x=$gap+($i%$Columns)*($ThumbWidth+$gap);$y=$gap+[math]::Floor($i/$Columns)*($thumbHeight+$label+$gap)
   $g.DrawString(('{0:D2}' -f [int]$e.page),$font,[Drawing.Brushes]::Black,[single]$x,[single]$y)
   $g.DrawImage($im,[int]$x,[int]($y+$label),$ThumbWidth,$thumbHeight)
  }finally{$im.Dispose()}
 }
 $outputFull=[IO.Path]::GetFullPath($OutputPath)
 New-Item -ItemType Directory -Force -Path (Split-Path $outputFull) | Out-Null
 $bmp.Save($outputFull,[Drawing.Imaging.ImageFormat]::Png)
}finally{$font.Dispose();$g.Dispose();$bmp.Dispose()}
@{source_sha256=$r.source_sha256;slides=$r.slides;contact_sheet_sha256=(Get-FileHash -LiteralPath $outputFull -Algorithm SHA256).Hash.ToLower()} | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath ($outputFull+'.source.json') -Encoding utf8
Write-Output $outputFull
