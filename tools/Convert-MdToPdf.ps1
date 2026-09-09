# =========================================================
#  Convert-MdToPdf.ps1
#  Markdown を HTML に整形し、Edge のヘッドレス印刷で PDF にする
#    例) powershell -ExecutionPolicy Bypass -File tools\Convert-MdToPdf.ps1
#        powershell -ExecutionPolicy Bypass -File tools\Convert-MdToPdf.ps1 -In 手順.md
# =========================================================
param(
    [string]$In = '',
    [string]$Out = '',
    [switch]$KeepHtml = $true
)
$ErrorActionPreference = 'Stop'

$Root = Split-Path -Parent $PSScriptRoot
if (-not $In) { $In = Join-Path $Root 'README.md' }
if (-not (Test-Path $In)) { throw "元ファイルが見つかりません: $In" }
$In = (Resolve-Path $In).Path
if (-not $Out) { $Out = [System.IO.Path]::ChangeExtension($In, '.pdf') }
$HtmlPath = [System.IO.Path]::ChangeExtension($Out, '.html')

# 出力先が開かれていないか先に確かめる（Acrobat などで開いていると上書きできない）
foreach ($f in @($Out, $HtmlPath)) {
    if (Test-Path $f) {
        try { $fs = [System.IO.File]::Open($f, 'Open', 'ReadWrite', 'None'); $fs.Close() }
        catch { throw "$f を更新できません。PDF ビューアやブラウザで開いている場合は閉じてから再実行してください。" }
    }
}

# --- ブラウザを探す -------------------------------------------------------
$browser = $null
foreach ($p in @(
        "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
        "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
        "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
        "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe")) {
    if (Test-Path $p) { $browser = $p; break }
}
if (-not $browser) { throw 'Edge / Chrome が見つかりませんでした。' }

# --- Markdown → HTML ------------------------------------------------------
function Esc([string]$s) {
    ($s -replace '&', '&amp;') -replace '<', '&lt;' -replace '>', '&gt;'
}
function Inline([string]$s) {
    $t = Esc $s
    $t = [regex]::Replace($t, '`([^`]+)`', { param($m) '<code>' + $m.Groups[1].Value + '</code>' })
    $t = [regex]::Replace($t, '\*\*([^*]+)\*\*', '<strong>$1</strong>')
    $t = [regex]::Replace($t, '\[([^\]]+)\]\(([^)]+)\)', '<a href="$2">$1</a>')
    return $t
}
function SplitRow([string]$s) {
    $t = $s.Trim()
    if ($t.StartsWith('|')) { $t = $t.Substring(1) }
    if ($t.EndsWith('|')) { $t = $t.Substring(0, $t.Length - 1) }
    return ($t -split '\|') | ForEach-Object { $_.Trim() }
}

$lines = [System.IO.File]::ReadAllLines($In, [System.Text.Encoding]::UTF8)
$html = New-Object System.Collections.Generic.List[string]
$list = ''      # '' | 'ul' | 'ol'
$para = New-Object System.Collections.Generic.List[string]

function Flush-Para {
    if ($script:para.Count -gt 0) {
        $script:html.Add('<p>' + ($script:para -join '<br>') + '</p>')
        $script:para.Clear()
    }
}
function Close-List {
    Flush-Para
    if ($script:list) { $script:html.Add("</$script:list>"); $script:list = '' }
}

$i = 0
while ($i -lt $lines.Count) {
    $l = $lines[$i]

    # コードブロック
    if ($l -match '^\s*```') {
        Close-List
        $i++
        $buf = @()
        while ($i -lt $lines.Count -and $lines[$i] -notmatch '^\s*```') { $buf += (Esc $lines[$i]); $i++ }
        $i++
        $html.Add('<pre><code>' + ($buf -join "`n") + '</code></pre>')
        continue
    }

    # テーブル
    if ($l -match '^\s*\|' -and ($i + 1) -lt $lines.Count -and $lines[$i + 1] -match '^\s*\|[\s:\-\|]+\|\s*$') {
        Close-List
        $head = SplitRow $l
        $align = SplitRow $lines[$i + 1]
        $i += 2
        $sty = @()
        foreach ($a in $align) {
            if ($a -match '^:.*:$') { $sty += 'center' } elseif ($a -match ':$') { $sty += 'right' } else { $sty += 'left' }
        }
        $t = New-Object System.Collections.Generic.List[string]
        $t.Add('<table><thead><tr>')
        for ($c = 0; $c -lt $head.Count; $c++) {
            $t.Add(('<th style="text-align:{0}">{1}</th>' -f $sty[[Math]::Min($c, $sty.Count - 1)], (Inline $head[$c])))
        }
        $t.Add('</tr></thead><tbody>')
        while ($i -lt $lines.Count -and $lines[$i] -match '^\s*\|') {
            $row = SplitRow $lines[$i]
            $t.Add('<tr>')
            for ($c = 0; $c -lt $row.Count; $c++) {
                $t.Add(('<td style="text-align:{0}">{1}</td>' -f $sty[[Math]::Min($c, $sty.Count - 1)], (Inline $row[$c])))
            }
            $t.Add('</tr>')
            $i++
        }
        $t.Add('</tbody></table>')
        $html.Add($t -join '')
        continue
    }

    # 見出し
    if ($l -match '^(#{1,6})\s+(.*)$') {
        Close-List
        $lv = $Matches[1].Length
        $html.Add("<h$lv>" + (Inline $Matches[2]) + "</h$lv>")
        $i++; continue
    }

    # 水平線
    if ($l -match '^\s*(-{3,}|\*{3,})\s*$') {
        Close-List
        $html.Add('<hr>')
        $i++; continue
    }

    # 箇条書き
    if ($l -match '^\s*[-*]\s+(.*)$') {
        Flush-Para
        if ($list -ne 'ul') { if ($list) { $html.Add("</$list>") }; $html.Add('<ul>'); $list = 'ul' }
        $html.Add('<li>' + (Inline $Matches[1]) + '</li>')
        $i++; continue
    }
    if ($l -match '^\s*(\d+)\.\s+(.*)$') {
        # 番号は元の値を引き継ぐ（コードブロックなどで一度切れても 1 に戻さない）
        $num = [int]$Matches[1]
        $body = $Matches[2]
        Flush-Para
        if ($list -ne 'ol') {
            if ($list) { $html.Add("</$list>") }
            if ($num -gt 1) { $html.Add('<ol start="' + $num + '">') } else { $html.Add('<ol>') }
            $list = 'ol'
        }
        $html.Add('<li>' + (Inline $body) + '</li>')
        $i++; continue
    }

    # 箇条書きの折り返し行（2 スペース以上のインデント）
    if ($list -and $l -match '^\s{2,}\S' -and $html.Count -gt 0 -and $html[$html.Count - 1].EndsWith('</li>')) {
        $last = $html[$html.Count - 1]
        $html[$html.Count - 1] = $last.Substring(0, $last.Length - 5) + '<br>' + (Inline $l.Trim()) + '</li>'
        $i++; continue
    }

    # 空行
    if ($l -match '^\s*$') { Close-List; $i++; continue }

    # 段落
    if ($list) { Close-List }
    $para.Add((Inline $l))
    $i++
}
Close-List

$title = [System.IO.Path]::GetFileNameWithoutExtension($In)
$css = @'
@page { size: A4; margin: 16mm 14mm 18mm 14mm; }
body { font-family: "Meiryo","Yu Gothic UI","Yu Gothic","MS PGothic",sans-serif;
       font-size: 10.5pt; line-height: 1.75; color: #222; }
h1 { font-size: 20pt; color:#1F4E78; border-bottom: 3px solid #1F4E78; padding-bottom: 6px; margin: 0 0 14px; }
h2 { font-size: 14pt; color:#1F4E78; border-left: 6px solid #1F4E78; padding-left: 9px;
     margin: 24px 0 8px; page-break-after: avoid; }
h3 { font-size: 12pt; margin: 16px 0 6px; page-break-after: avoid; }
p  { margin: 6px 0; }
ul, ol { margin: 6px 0; padding-left: 1.5em; }
li { margin: 3px 0; }
table { border-collapse: collapse; width: 100%; margin: 10px 0; font-size: 9.5pt;
        page-break-inside: avoid; }
th { background: #1F4E78; color: #fff; font-weight: bold; }
th, td { border: 1px solid #b9c0c7; padding: 5px 8px; vertical-align: top; }
tbody tr:nth-child(even) td { background: #f5f7fa; }
code { font-family: "Consolas","MS Gothic",monospace; font-size: 9.5pt;
       background: #eef1f4; padding: 1px 4px; border-radius: 3px; }
pre { background: #f6f8fa; border: 1px solid #dde1e6; border-left: 4px solid #1F4E78;
      padding: 10px 12px; margin: 10px 0; page-break-inside: avoid; white-space: pre-wrap; }
pre code { background: none; padding: 0; }
hr { border: 0; border-top: 1px solid #dde1e6; margin: 20px 0; }
strong { color: #12385c; }
a { color: #1F4E78; }
'@

$doc = "<!DOCTYPE html><html lang=`"ja`"><head><meta charset=`"utf-8`"><title>$title</title><style>$css</style></head><body>" +
       ($html -join "`r`n") + '</body></html>'
[System.IO.File]::WriteAllText($HtmlPath, $doc, (New-Object System.Text.UTF8Encoding($true)))

# --- HTML → PDF -----------------------------------------------------------
$tmpProfile = Join-Path $env:TEMP ('mdpdf_' + [Guid]::NewGuid().ToString('N'))
$url = 'file:///' + ($HtmlPath -replace '\\', '/')
function Q([string]$s) { if ($s -match '\s') { '"' + $s + '"' } else { $s } }
$argList = @(
    '--headless=new', '--disable-gpu', '--disable-extensions',
    (Q "--user-data-dir=$tmpProfile"),
    '--no-pdf-header-footer',
    '--run-all-compositor-stages-before-draw',
    '--virtual-time-budget=10000',
    (Q "--print-to-pdf=$Out"),
    (Q $url)
)
if (Test-Path $Out) { Remove-Item $Out -Force }
# 注) Edge は起動時に警告を stderr に出すため、& 演算子ではなく Start-Process で待つ
Start-Process -FilePath $browser -ArgumentList $argList -Wait -WindowStyle Hidden | Out-Null
Start-Sleep -Milliseconds 400
Remove-Item $tmpProfile -Recurse -Force -ErrorAction SilentlyContinue

if (-not (Test-Path $Out)) { throw "PDF を作成できませんでした: $Out" }
if (-not $KeepHtml) { Remove-Item $HtmlPath -Force -ErrorAction SilentlyContinue }

Write-Host ('作成: {0}  ({1:N0} bytes)' -f $Out, (Get-Item $Out).Length)
if ($KeepHtml) { Write-Host ('作成: {0}  ({1:N0} bytes)' -f $HtmlPath, (Get-Item $HtmlPath).Length) }
