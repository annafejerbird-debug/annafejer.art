param(
  [ValidateSet("wide", "a4", "both")]
  [string]$Format = "both"
)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$CatalogPath = Join-Path $Root "portfolio_compiled_works_metadata\catalog.json"
$Works = Get-Content -LiteralPath $CatalogPath -Raw | ConvertFrom-Json

$Ink = 0x1A1A1A
$Graphite = 0x3F3F3F
$Rule = 0xBDBDBD
$msoFalse = 0
$msoTrue = -1
$ppLayoutBlank = 12

function Add-TextBox {
  param($Slide, [string]$Text, [double]$Left, [double]$Top, [double]$Width, [double]$Height, [double]$Size, [bool]$Bold = $false, [int]$Color = $Ink, [string]$Align = "left", [bool]$Italic = $false)
  $shape = $Slide.Shapes.AddTextbox(1, $Left, $Top, $Width, $Height)
  $shape.TextFrame.TextRange.Text = $Text
  $shape.TextFrame.MarginLeft = 0
  $shape.TextFrame.MarginRight = 0
  $shape.TextFrame.MarginTop = 0
  $shape.TextFrame.MarginBottom = 0
  $shape.TextFrame.WordWrap = $msoTrue
  $shape.TextFrame.AutoSize = 0
  $shape.TextFrame.TextRange.Font.Name = "Arial"
  $shape.TextFrame.TextRange.Font.Size = $Size
  $shape.TextFrame.TextRange.Font.Bold = if ($Bold) { $msoTrue } else { $msoFalse }
  $shape.TextFrame.TextRange.Font.Italic = if ($Italic) { $msoTrue } else { $msoFalse }
  $shape.TextFrame.TextRange.Font.Color.RGB = $Color
  if ($Align -eq "right") { $shape.TextFrame.TextRange.ParagraphFormat.Alignment = 3 }
  elseif ($Align -eq "center") { $shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2 }
  else { $shape.TextFrame.TextRange.ParagraphFormat.Alignment = 1 }
  return $shape
}

function Add-Line {
  param($Slide, [double]$X1, [double]$Y1, [double]$X2, [double]$Y2, [int]$Color = $Rule, [double]$Width = 0.6)
  $line = $Slide.Shapes.AddLine($X1, $Y1, $X2, $Y2)
  $line.Line.ForeColor.RGB = $Color
  $line.Line.Weight = $Width
  return $line
}

function Add-ImageContain {
  param($Slide, [string]$Path, [double]$CenterX, [double]$CenterY, [double]$FrameW, [double]$FrameH, [double]$PixelW, [double]$PixelH)
  $aspect = $PixelW / $PixelH
  $drawW = [Math]::Min($FrameW, $FrameH * $aspect)
  $drawH = [Math]::Min($FrameH, $FrameW / $aspect)
  $left = $CenterX - ($drawW / 2.0)
  $top = $CenterY - ($drawH / 2.0)
  return $Slide.Shapes.AddPicture((Resolve-Path -LiteralPath $Path), $msoFalse, $msoTrue, $left, $top, $drawW, $drawH)
}

function Page-Range {
  param($Work)
  if ($Work.content_start_page -eq $Work.content_end_page) { return ("{0:D2}" -f [int]$Work.content_start_page) }
  return ("{0:D2}-{1:D2}" -f [int]$Work.content_start_page, [int]$Work.content_end_page)
}

function Caption-Meta {
  param($Work)
  $parts = @()
  if ($Work.year) { $parts += $Work.year }
  if ($Work.format) { $parts += $Work.format.ToUpperInvariant() }
  return ($parts -join "  ·  ")
}

function Caption-Details {
  param($Work)
  $parts = @()
  if ($Work.materials) { $parts += $Work.materials }
  if ($Work.size) { $parts += $Work.size }
  return ($parts -join "  ·  ")
}

function Add-Caption {
  param($Slide, $Work, [int]$WorkPage, $Cfg)
  Add-TextBox $Slide $Work.title $Cfg.CaptionX $Cfg.CaptionTitleY $Cfg.CaptionW 16 $Cfg.TitleSize $true $Ink | Out-Null
  Add-TextBox $Slide (Caption-Meta $Work) $Cfg.CaptionX $Cfg.CaptionMetaY $Cfg.CaptionW 12 $Cfg.MetaSize $false $Graphite | Out-Null
  Add-TextBox $Slide (Caption-Details $Work) $Cfg.CaptionX $Cfg.CaptionDetailsY $Cfg.CaptionW 12 $Cfg.MetaSize $false $Graphite | Out-Null
  if ($Work.location) { Add-TextBox $Slide $Work.location $Cfg.CaptionX $Cfg.CaptionLocationY $Cfg.CaptionW 12 $Cfg.LocationSize $false $Graphite "left" $true | Out-Null }
  $pageNo = [int]$Work.content_start_page + $WorkPage - 1
  Add-Line $Slide $Cfg.PageRuleX $Cfg.PageRuleY1 $Cfg.PageRuleX $Cfg.PageRuleY2 $Ink 0.45 | Out-Null
  Add-TextBox $Slide ("{0:D2}" -f $pageNo) ($Cfg.PageNumberRight - 28) $Cfg.PageNumberY 28 14 $Cfg.PageNumberSize $true $Ink "right" | Out-Null
}

function Add-TocRow {
  param($Slide, $Work, [double]$X, [double]$Y, $Cfg, [double]$MetaOffset = 15)
  Add-TextBox $Slide ("Work {0}" -f $Work.work_number) $X $Y $Cfg.TocLabelW 12 $Cfg.TocLabelSize $true $Graphite | Out-Null
  Add-TextBox $Slide $Work.title ($X + $Cfg.TocTitleOffset) $Y $Cfg.TocTitleW 20 $Cfg.TocTitleSize $true $Ink | Out-Null
  Add-TextBox $Slide (Page-Range $Work) ($X + $Cfg.TocRangeOffset - 42) $Y 42 14 $Cfg.TocRangeSize $true $Ink "right" | Out-Null
  $meta = @($Work.format.ToUpperInvariant(), $Work.year, $Work.size) -ne "" -join "  ·  "
  Add-TextBox $Slide $meta ($X + $Cfg.TocTitleOffset) ($Y + $MetaOffset) $Cfg.TocMetaW 18 $Cfg.TocMetaSize $false $Graphite | Out-Null
}

function Add-ContentPage {
  param($Presentation, $Work, [int]$WorkPage, $Cfg)
  $slide = $Presentation.Slides.Add($Presentation.Slides.Count + 1, $ppLayoutBlank)
  $imgs = @($Work.images | Where-Object { [int]$_.work_page -eq $WorkPage })
  if ($imgs.Count -eq 1) {
    $img = $imgs[0]
    Add-ImageContain $slide (Join-Path $Root $img.relative_path) $Cfg.ArtCenterX $Cfg.ArtCenterY $Cfg.ArtW $Cfg.ArtH $img.width_px $img.height_px | Out-Null
  } elseif ($imgs.Count -eq 2) {
    $a1 = [double]$imgs[0].width_px / [double]$imgs[0].height_px
    $a2 = [double]$imgs[1].width_px / [double]$imgs[1].height_px
    $h = [Math]::Min($Cfg.ArtH, ($Cfg.ArtW - $Cfg.PairGap) / ($a1 + $a2))
    $w1 = $h * $a1
    $w2 = $h * $a2
    $total = $w1 + $Cfg.PairGap + $w2
    $x1 = $Cfg.ArtCenterX - ($total / 2.0) + ($w1 / 2.0)
    $x2 = $Cfg.ArtCenterX + ($total / 2.0) - ($w2 / 2.0)
    Add-ImageContain $slide (Join-Path $Root $imgs[0].relative_path) $x1 $Cfg.ArtCenterY $w1 $h $imgs[0].width_px $imgs[0].height_px | Out-Null
    Add-ImageContain $slide (Join-Path $Root $imgs[1].relative_path) $x2 $Cfg.ArtCenterY $w2 $h $imgs[1].width_px $imgs[1].height_px | Out-Null
  }
  Add-Caption $slide $Work $WorkPage $Cfg
}

function Build-Deck {
  param([string]$Name, $Cfg)
  $app = New-Object -ComObject PowerPoint.Application
  $app.DisplayAlerts = 1
  $presentation = $app.Presentations.Add($msoTrue)
  $presentation.PageSetup.SlideWidth = $Cfg.PageW
  $presentation.PageSetup.SlideHeight = $Cfg.PageH

  $slide = $presentation.Slides.Add(1, $ppLayoutBlank)
  Add-TextBox $slide "Anna Fejer" $Cfg.CoverX $Cfg.CoverTitleY 300 48 $Cfg.CoverTitleSize $true $Ink | Out-Null
  Add-TextBox $slide "Portfolio" $Cfg.CoverX $Cfg.CoverSubY 220 36 $Cfg.CoverSubSize $false $Graphite | Out-Null
  Add-Line $slide $Cfg.CoverX $Cfg.CoverRuleY ($Cfg.CoverX + $Cfg.CoverRuleW) $Cfg.CoverRuleY $Rule 0.6 | Out-Null
  Add-TextBox $slide "Selected works, 2024-2026" $Cfg.CoverX $Cfg.CoverSmallY 220 16 $Cfg.CoverSmallSize $false $Graphite | Out-Null

  $slide = $presentation.Slides.Add(2, $ppLayoutBlank)
  Add-TextBox $slide "Contents" $Cfg.TocHeadingX $Cfg.TocHeadingY 240 42 $Cfg.TocHeadingSize $true $Ink | Out-Null
  Add-Line $slide $Cfg.TocHeadingX $Cfg.TocRuleY $Cfg.TocRuleEndX $Cfg.TocRuleY $Rule 0.6 | Out-Null
  Add-TextBox $slide "Pages" ($Cfg.TocRuleEndX - 48) ($Cfg.TocRuleY - 10) 48 14 $Cfg.TocMetaSize $false $Graphite "right" | Out-Null
  Add-Line $slide $Cfg.TocDividerX $Cfg.TocDividerY1 $Cfg.TocDividerX $Cfg.TocDividerY2 $Rule 0.4 | Out-Null
  for ($i=0; $i -lt $Works.Count; $i++) {
    $work = $Works[$i]
    if ($i -lt 8) {
      Add-TocRow $slide $work $Cfg.TocLeftX $Cfg.TocRows[$i] $Cfg ($(if ($work.work_number -eq 7) { $Cfg.TocLongMetaOffset } else { $Cfg.TocMetaOffset }))
    } else {
      Add-TocRow $slide $work $Cfg.TocRightX $Cfg.TocRows[$i-8] $Cfg $Cfg.TocMetaOffset
    }
  }

  foreach ($work in $Works) {
    for ($p=1; $p -le [int]$work.page_count; $p++) {
      Add-ContentPage $presentation $work $p $Cfg
    }
  }

  $out = Join-Path $Root $Name
  if (Test-Path -LiteralPath $out) { Remove-Item -LiteralPath $out -Force }
  $presentation.SaveAs($out)
  $presentation.Close()
  $app.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) | Out-Null
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
  Write-Host "Wrote $out"
}

$WideCfg = [pscustomobject]@{
  PageW=720.0; PageH=405.0; ArtCenterX=360.0; ArtCenterY=235.0; ArtW=616.0; ArtH=290.0; PairGap=44.0
  CaptionX=22.0; CaptionW=560.0; CaptionTitleY=345.0; CaptionMetaY=359.0; CaptionDetailsY=371.0; CaptionLocationY=383.0
  TitleSize=10.4; MetaSize=8.0; LocationSize=7.3; PageRuleX=674.0; PageRuleY1=369.0; PageRuleY2=389.0; PageNumberRight=698.0; PageNumberY=374.0; PageNumberSize=9.0
  CoverX=86.0; CoverTitleY=144.0; CoverSubY=195.0; CoverRuleY=235.0; CoverRuleW=146.0; CoverSmallY=252.0; CoverTitleSize=42.0; CoverSubSize=30.0; CoverSmallSize=10.0
  TocHeadingX=52.0; TocHeadingY=30.0; TocHeadingSize=30.0; TocRuleY=51.0; TocRuleEndX=668.0; TocDividerX=360.0; TocDividerY1=57.0; TocDividerY2=373.0
  TocLeftX=52.0; TocRightX=382.0; TocRows=@(73,113,153,193,233,273,313,361)
  TocLabelW=52.0; TocTitleOffset=58.0; TocTitleW=200.0; TocRangeOffset=286.0; TocMetaW=232.0; TocMetaOffset=15.0; TocLongMetaOffset=14.0
  TocLabelSize=9.0; TocTitleSize=11.0; TocRangeSize=10.0; TocMetaSize=8.0
}

$A4Cfg = [pscustomobject]@{
  PageW=841.8898; PageH=595.2756; ArtCenterX=421.0; ArtCenterY=343.5; ArtW=758.0; ArtH=427.0; PairGap=56.0
  CaptionX=42.0; CaptionW=650.0; CaptionTitleY=503.0; CaptionMetaY=519.0; CaptionDetailsY=533.0; CaptionLocationY=547.0
  TitleSize=11.2; MetaSize=8.6; LocationSize=8.0; PageRuleX=774.0; PageRuleY1=525.0; PageRuleY2=557.0; PageNumberRight=799.8898; PageNumberY=536.0; PageNumberSize=10.0
  CoverX=92.0; CoverTitleY=225.0; CoverSubY=285.0; CoverRuleY=329.0; CoverRuleW=176.0; CoverSmallY=355.0; CoverTitleSize=52.0; CoverSubSize=38.0; CoverSmallSize=12.0
  TocHeadingX=70.0; TocHeadingY=75.0; TocHeadingSize=34.0; TocRuleY=103.0; TocRuleEndX=772.0; TocDividerX=421.0; TocDividerY1=111.0; TocDividerY2=549.0
  TocLeftX=70.0; TocRightX=448.0; TocRows=@(135,185,235,285,335,385,435,503)
  TocLabelW=62.0; TocTitleOffset=68.0; TocTitleW=240.0; TocRangeOffset=324.0; TocMetaW=260.0; TocMetaOffset=19.0; TocLongMetaOffset=16.0
  TocLabelSize=11.0; TocTitleSize=14.0; TocRangeSize=12.5; TocMetaSize=10.0
}

if ($Format -eq "wide" -or $Format -eq "both") { Build-Deck "portfolio_from_ppt_images_editable.pptx" $WideCfg }
if ($Format -eq "a4" -or $Format -eq "both") { Build-Deck "portfolio_from_ppt_images_a4_editable.pptx" $A4Cfg }
