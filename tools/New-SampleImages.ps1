# =========================================================
#  New-SampleImages.ps1
#  動作確認用のサンプル画像(JPG)とサンプル原本(PDF)を作る
#  ※ 実運用では img\ に実際の写真・図面画像を置いてください
# =========================================================
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

$Root   = Split-Path -Parent $PSScriptRoot
$ImgDir = Join-Path $Root 'img'
$PdfDir = Join-Path $Root 'pdf'
foreach ($d in @($ImgDir, $PdfDir)) {
    if (-not (Test-Path $d)) { New-Item -ItemType Directory -Path $d | Out-Null }
}

function Get-JpFont {
    param([single]$Size, [switch]$Bold)
    $style = if ($Bold) { [System.Drawing.FontStyle]::Bold } else { [System.Drawing.FontStyle]::Regular }
    foreach ($n in @('Meiryo UI', 'Meiryo', 'Yu Gothic UI', 'MS UI Gothic', 'MS Gothic')) {
        try {
            $f = New-Object System.Drawing.Font($n, $Size, $style)
            if ($f.Name -eq $n) { return $f }
            $f.Dispose()
        } catch { }
    }
    return (New-Object System.Drawing.Font('Arial', $Size, $style))
}

function C([string]$hex) { [System.Drawing.ColorTranslator]::FromHtml($hex) }

function Draw-Box {
    param($g, [int]$x, [int]$y, [int]$w, [int]$h, [string]$fill, [string]$line, [single]$lw = 2)
    if ($fill) {
        $b = New-Object System.Drawing.SolidBrush (C $fill)
        $g.FillRectangle($b, $x, $y, $w, $h); $b.Dispose()
    }
    if ($line) {
        $p = New-Object System.Drawing.Pen((C $line), $lw)
        $g.DrawRectangle($p, $x, $y, $w, $h); $p.Dispose()
    }
}

function Draw-Ellipse {
    param($g, [int]$x, [int]$y, [int]$w, [int]$h, [string]$fill, [string]$line, [single]$lw = 2)
    if ($fill) {
        $b = New-Object System.Drawing.SolidBrush (C $fill)
        $g.FillEllipse($b, $x, $y, $w, $h); $b.Dispose()
    }
    if ($line) {
        $p = New-Object System.Drawing.Pen((C $line), $lw)
        $g.DrawEllipse($p, $x, $y, $w, $h); $p.Dispose()
    }
}

function Draw-Line {
    param($g, [int]$x1, [int]$y1, [int]$x2, [int]$y2, [string]$line, [single]$lw = 2)
    $p = New-Object System.Drawing.Pen((C $line), $lw)
    $g.DrawLine($p, $x1, $y1, $x2, $y2); $p.Dispose()
}

function Draw-Text {
    param($g, [string]$t, [int]$x, [int]$y, [single]$size = 14, [string]$color = '#333333', [switch]$Bold)
    $f = if ($Bold) { Get-JpFont -Size $size -Bold } else { Get-JpFont -Size $size }
    $b = New-Object System.Drawing.SolidBrush (C $color)
    $g.DrawString($t, $f, $b, [single]$x, [single]$y)
    $b.Dispose(); $f.Dispose()
}

function New-Image {
    param([string]$Name, [int]$W, [int]$H, [string]$Bg, [scriptblock]$Body)
    $bmp = New-Object System.Drawing.Bitmap($W, $H)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit
    $g.Clear((C $Bg))
    & $Body $g $W $H
    $path = Join-Path $ImgDir $Name
    $bmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Jpeg)
    $g.Dispose(); $bmp.Dispose()
    Write-Host "  作成: img\$Name"
}

# ---------------------------------------------------------
# 1. 工場マップ
# ---------------------------------------------------------
New-Image 'map_factory.jpg' 1000 700 '#F7F7F4' {
    param($g, $W, $H)
    for ($x = 0; $x -lt $W; $x += 50) { Draw-Line $g $x 0 $x $H '#E4E4DE' 1 }
    for ($y = 0; $y -lt $H; $y += 50) { Draw-Line $g 0 $y $W $y '#E4E4DE' 1 }
    Draw-Box $g 40 40 920 620 $null '#9A9A90' 3
    Draw-Text $g '工場マップ  1F 成形ライン' 55 55 20 '#333333' -Bold
    Draw-Text $g '※ サンプル画像です。実際の工場レイアウト図に差し替えてください。' 55 620 11 '#909090'

    # EQ-001 : x .12-.30 / y .30-.52
    Draw-Box $g 120 210 180 154 '#DCE9F7' '#0070C0' 3
    Draw-Text $g '混錬機A' 150 265 18 '#0070C0' -Bold

    # EQ-002 : x .55-.77 / y .55-.73
    Draw-Box $g 550 385 220 126 '#DCE9F7' '#0070C0' 3
    Draw-Text $g 'コンベアB' 600 430 18 '#0070C0' -Bold

    Draw-Box $g 620 120 250 120 '#EFEFEA' '#B5B5AC' 2
    Draw-Text $g '（未登録エリア）' 660 170 14 '#A0A0A0'
    Draw-Line $g 300 287 550 448 '#B5B5AC' 2
}

# ---------------------------------------------------------
# 2. 設備写真（混錬機A）
# ---------------------------------------------------------
New-Image 'eq_mixer.jpg' 900 650 '#EFF1F3' {
    param($g, $W, $H)
    Draw-Text $g '設備写真  混錬機A' 30 25 20 '#333333' -Bold
    Draw-Box $g 120 140 480 380 '#C9CFD6' '#7A828C' 3
    Draw-Ellipse $g 200 220 200 200 '#B0B8C2' '#7A828C' 3
    Draw-Text $g '本体' 270 300 16 '#4A5058'
    # 減速機部 : x .60-.78 / y .35-.57
    Draw-Box $g 540 227 162 143 '#D6EFDF' '#009650' 3
    Draw-Text $g '減速機部' 560 285 16 '#009650' -Bold
    Draw-Text $g '※ サンプル画像です。実際の設備写真に差し替えてください。' 30 600 11 '#909090'
}

# ---------------------------------------------------------
# 3. 設備写真（コンベアB）
# ---------------------------------------------------------
New-Image 'eq_conveyor.jpg' 900 650 '#EFF1F3' {
    param($g, $W, $H)
    Draw-Text $g '設備写真  コンベアB' 30 25 20 '#333333' -Bold
    Draw-Box $g 80 280 740 90 '#C9CFD6' '#7A828C' 3
    for ($x = 100; $x -lt 820; $x += 60) { Draw-Ellipse $g $x 300 50 50 '#B0B8C2' '#7A828C' 2 }
    Draw-Text $g '※ サンプル画像です。子はまだ登録されていません。' 30 600 11 '#909090'
}

# ---------------------------------------------------------
# 4. 詳細写真（減速機部）
# ---------------------------------------------------------
New-Image 'dt_gearbox.jpg' 900 650 '#F2F0EC' {
    param($g, $W, $H)
    Draw-Text $g '詳細写真  混錬機A 減速機部' 30 25 20 '#333333' -Bold
    Draw-Box $g 100 120 700 420 '#DDD8D0' '#8A8378' 3
    # 組図へ : x .30-.60 / y .30-.60
    Draw-Box $g 270 195 270 195 '#FBE6CC' '#E67800' 3
    Draw-Text $g '減速機ユニット' 310 275 18 '#E67800' -Bold
    Draw-Text $g '※ サンプル画像です。実際の部位写真に差し替えてください。' 30 600 11 '#909090'
}

# ---------------------------------------------------------
# 5. 組図 A-100
# ---------------------------------------------------------
New-Image 'as_a100.jpg' 1000 700 '#FFFFFF' {
    param($g, $W, $H)
    Draw-Box $g 20 20 960 660 $null '#333333' 3
    Draw-Box $g 30 30 940 640 $null '#888888' 1
    Draw-Text $g '減速機 組立図' 45 40 18 '#333333' -Bold

    # 図枠
    Draw-Box $g 300 150 400 380 $null '#333333' 2
    Draw-Line $g 300 250 700 250 '#333333' 1
    Draw-Line $g 300 430 700 430 '#333333' 1
    Draw-Ellipse $g 430 280 140 120 $null '#333333' 2
    Draw-Line $g 500 150 500 530 '#888888' 1

    # バルーン 1 : (.18,.28)  2 : (.46,.55)  3 : (.70,.30)
    Draw-Ellipse $g 180 196 60 60 '#FFFFFF' '#C81E28' 2
    Draw-Text $g '1' 202 208 16 '#C81E28' -Bold
    Draw-Line $g 240 226 330 240 '#C81E28' 1

    Draw-Ellipse $g 460 385 60 60 '#FFFFFF' '#C81E28' 2
    Draw-Text $g '2' 482 397 16 '#C81E28' -Bold

    Draw-Ellipse $g 700 210 60 60 '#FFFFFF' '#7850C8' 2
    Draw-Text $g '3' 722 222 16 '#7850C8' -Bold
    Draw-Line $g 700 240 640 260 '#7850C8' 1

    # 表題欄
    Draw-Box $g 640 560 330 100 $null '#333333' 2
    Draw-Line $g 640 595 970 595 '#333333' 1
    Draw-Line $g 640 628 970 628 '#333333' 1
    Draw-Text $g '図番  A-100' 650 568 13 '#333333' -Bold
    Draw-Text $g '品名  減速機 組立' 650 601 13 '#333333'
    Draw-Text $g '尺度  1:2      サンプル図面' 650 634 12 '#666666'
}

# ---------------------------------------------------------
# 6-8. 部品
# ---------------------------------------------------------
New-Image 'pt_bearing.jpg' 600 450 '#F7F7F7' {
    param($g, $W, $H)
    Draw-Text $g '購入部品  ベアリング 6205ZZ' 25 20 16 '#333333' -Bold
    Draw-Ellipse $g 180 110 240 240 '#D8D8D8' '#5A5A5A' 3
    Draw-Ellipse $g 240 170 120 120 '#FFFFFF' '#5A5A5A' 3
    Draw-Text $g '棚番 A-3-2' 240 380 14 '#C81E28' -Bold
}

New-Image 'pt_oring.jpg' 600 450 '#F7F7F7' {
    param($g, $W, $H)
    Draw-Text $g '購入部品  Oリング S-40' 25 20 16 '#333333' -Bold
    Draw-Ellipse $g 190 110 220 220 $null '#2A2A2A' 14
    Draw-Text $g '棚番 B-1-5' 240 380 14 '#C81E28' -Bold
}

New-Image 'dw_shaft.jpg' 900 650 '#FFFFFF' {
    param($g, $W, $H)
    Draw-Box $g 20 20 860 610 $null '#333333' 3
    Draw-Text $g '部品図面  出力シャフト  D-021' 40 40 18 '#333333' -Bold
    Draw-Box $g 150 260 560 90 $null '#333333' 2
    Draw-Box $g 150 285 80 40 $null '#333333' 2
    Draw-Box $g 630 275 80 60 $null '#333333' 2
    Draw-Line $g 150 305 710 305 '#888888' 1
    Draw-Text $g 'φ40  L=560   S45C' 300 380 14 '#333333'
    Draw-Text $g '棚番 C-2-1' 300 420 14 '#C81E28' -Bold
}

# ---------------------------------------------------------
# サンプル原本 PDF
# ---------------------------------------------------------
function New-SimplePdf {
    param([string]$Path, [string[]]$Lines)

    $content = "BT`n/F1 20 Tf`n"
    $y = 760
    foreach ($l in $Lines) {
        $esc = $l -replace '\\', '\\\\' -replace '\(', '\(' -replace '\)', '\)'
        $content += "1 0 0 1 60 $y Tm ($esc) Tj`n"
        $y -= 34
    }
    $content += "ET`n"

    $objs = @(
        '<</Type/Catalog/Pages 2 0 R>>',
        '<</Type/Pages/Kids[3 0 R]/Count 1>>',
        '<</Type/Page/Parent 2 0 R/MediaBox[0 0 595 842]/Resources<</Font<</F1 4 0 R>>>>/Contents 5 0 R>>',
        '<</Type/Font/Subtype/Type1/BaseFont/Helvetica>>',
        ("<</Length {0}>>`nstream`n{1}endstream" -f $content.Length, $content)
    )

    $out = "%PDF-1.4`n"
    $offsets = @()
    for ($i = 0; $i -lt $objs.Count; $i++) {
        $offsets += $out.Length
        $out += ("{0} 0 obj`n{1}`nendobj`n" -f ($i + 1), $objs[$i])
    }
    $xref = $out.Length
    $out += ("xref`n0 {0}`n" -f ($objs.Count + 1))
    $out += "0000000000 65535 f `n"
    foreach ($o in $offsets) { $out += ('{0:D10} 00000 n ' -f $o) + "`n" }
    $out += ("trailer`n<</Size {0}/Root 1 0 R>>`nstartxref`n{1}`n%%EOF`n" -f ($objs.Count + 1), $xref)

    [System.IO.File]::WriteAllText($Path, $out, [System.Text.Encoding]::ASCII)
    Write-Host ("  作成: pdf\{0}" -f (Split-Path -Leaf $Path))
}

New-SimplePdf (Join-Path $PdfDir 'A-100.pdf') @(
    'A-100  Gear Reducer Assembly',
    'SAMPLE original PDF',
    'Replace this file with the real drawing PDF.'
)
New-SimplePdf (Join-Path $PdfDir 'D-021.pdf') @(
    'D-021  Output Shaft',
    'SAMPLE original PDF',
    'Replace this file with the real drawing PDF.'
)

Write-Host ''
Write-Host 'サンプル画像・PDF を作成しました。' -ForegroundColor Green
