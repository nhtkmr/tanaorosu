# =========================================================
#  Build-Workbook.ps1
#  src\ の VBA とシート定義から 予備品図面管理.xlsm を組み立てる
#  ※ Excel の「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」が必要
# =========================================================
$ErrorActionPreference = 'Stop'

$Root = Split-Path -Parent $PSScriptRoot
$Src  = Join-Path $Root 'src'
$Out  = Join-Path $Root '予備品図面管理.xlsm'

# --- 前提チェック ---------------------------------------------------------
$vbom = (Get-ItemProperty 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security' -Name AccessVBOM -ErrorAction SilentlyContinue).AccessVBOM
if ($vbom -ne 1) {
    throw @'
Excel の設定が必要です。
 ファイル > オプション > トラスト センター > トラスト センターの設定 > マクロの設定
 「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」にチェックを入れてから再実行してください。
'@
}
if (Test-Path $Out) {
    try {
        $fs = [System.IO.File]::Open($Out, 'Open', 'ReadWrite', 'None')
        $fs.Close()
    } catch {
        throw "$Out を更新できません。Excel で開いている場合は閉じてから再実行してください。"
    }
    # このスクリプトは台帳をサンプルデータに戻す。上書き前に必ず控えを取る。
    # 入力済みのデータを残したままコードだけ更新したいときは tools\Update-Workbook.ps1 を使うこと。
    $bkDir = Join-Path $Root 'backup'
    if (-not (Test-Path $bkDir)) { New-Item -ItemType Directory -Path $bkDir | Out-Null }
    $bk = Join-Path $bkDir ('予備品図面管理_' + (Get-Date -Format 'yyyyMMdd_HHmmss') + '.xlsm')
    Copy-Item $Out $bk -Force
    Write-Host "既存のブックを控えました: $bk" -ForegroundColor Yellow
    Write-Host '※ このビルドで台帳はサンプルデータに戻ります（データを残すなら Update-Workbook.ps1）' -ForegroundColor Yellow
}

# --- 定数 -----------------------------------------------------------------
$xlLeft = -4131; $xlCenter = -4108; $xlRight = -4152; $xlVCenter = -4108
$xlContinuous = 1; $xlThin = 2
$xlValidateList = 3; $xlValidAlertStop = 1; $xlBetween = 1
$xlSrcRange = 1; $xlYes = 1
$msoRoundRect = 5; $xlFreeFloating = 3
$xlOpenXMLWorkbookMacroEnabled = 52
$FONT = 'Meiryo UI'

function RGBv([int]$r, [int]$g, [int]$b) { $r + $g * 256 + $b * 65536 }

function ColLetter([int]$c) {
    $s = ''
    while ($c -gt 0) { $m = ($c - 1) % 26; $s = [string][char](65 + $m) + $s; $c = [int](($c - $m) / 26) }
    return $s
}

# COM の型解決でつまずかないよう、番地文字列 + 明示キャストで書き込む
function Set-Cell {
    param($ws, [int]$r, [int]$c, $v)
    if ($null -eq $v) { return }
    if ("$v" -eq '') { return }
    $addr = (ColLetter $c) + $r
    if ($v -is [double] -or $v -is [int] -or $v -is [long] -or $v -is [decimal] -or $v -is [single]) {
        $ws.Range($addr).Value2 = [double]$v
    } else {
        $ws.Range($addr).Value2 = [string]$v
    }
}
$CLR_HEAD   = RGBv 31 78 120
$CLR_HEAD2  = RGBv 68 114 148
$CLR_WHITE  = RGBv 255 255 255
$CLR_TEXT   = RGBv 50 50 50
$CLR_MUTED  = RGBv 130 130 130
$CLR_FRAME  = RGBv 180 180 180
$CLR_PANEL  = RGBv 247 247 244
$CLR_BTN    = RGBv 242 242 242
$CLR_BTNPRI = RGBv 0 112 192

function Add-Button {
    param($ws, [string]$name, [string]$text, [string]$macro,
          [double]$left, [double]$top, [double]$w, [double]$h,
          [int]$fill = -1, [int]$fg = -1)
    if ($fill -lt 0) { $fill = $script:CLR_BTN }
    if ($fg -lt 0) { $fg = $script:CLR_TEXT }
    $sh = $ws.Shapes.AddShape($script:msoRoundRect, $left, $top, $w, $h)
    $sh.Name = $name
    $sh.Placement = $script:xlFreeFloating
    $sh.Fill.ForeColor.RGB = $fill
    $sh.Line.ForeColor.RGB = (RGBv 165 165 165)
    $sh.Line.Weight = 0.75
    $sh.Shadow.Visible = 0
    $tf = $sh.TextFrame2
    $tf.MarginLeft = 0; $tf.MarginRight = 0; $tf.MarginTop = 0; $tf.MarginBottom = 0
    $tf.VerticalAnchor = 3          # msoAnchorMiddle
    $tf.WordWrap = 0
    $tf.TextRange.Text = $text
    $tf.TextRange.Font.Name = $script:FONT
    $tf.TextRange.Font.Size = 10
    $tf.TextRange.Font.Fill.ForeColor.RGB = $fg
    $tf.TextRange.ParagraphFormat.Alignment = 2   # msoAlignCenter
    $sh.OnAction = $macro
    return $sh
}

function Set-Header {
    param($ws, [string]$addr, [string]$text, [int]$bg = -1)
    if ($bg -lt 0) { $bg = $script:CLR_HEAD }
    $r = $ws.Range($addr)
    $r.Value2 = $text
    $r.Font.Bold = $true
    $r.Font.Color = $script:CLR_WHITE
    $r.Interior.Color = $bg
    $r.HorizontalAlignment = $script:xlLeft
    $r.VerticalAlignment = $script:xlVCenter
    $r.IndentLevel = 1
}

# --- Excel 起動 -----------------------------------------------------------
Write-Host 'Excel を起動しています...'
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.EnableEvents = $false
$xl.ScreenUpdating = $false

try {
    $wb = $xl.Workbooks.Add()
    while ($wb.Sheets.Count -gt 1) { $wb.Sheets.Item($wb.Sheets.Count).Delete() }
    $wb.Sheets.Item(1).Name = 'ビュー'
    foreach ($n in @('ノード', 'リンク', '部品一覧', '登録', '設定')) {
        $s = $wb.Sheets.Add([Type]::Missing, $wb.Sheets.Item($wb.Sheets.Count))
        $s.Name = $n
    }

    # =====================================================================
    # 設定シート
    # =====================================================================
    Write-Host '  シート: 設定'
    $ws = $wb.Sheets.Item('設定')
    $ws.Cells.Font.Name = $FONT
    $ws.Cells.Font.Size = 10
    $ws.Columns.Item('A').ColumnWidth = 34
    $ws.Columns.Item('B').ColumnWidth = 42
    $ws.Columns.Item('C').ColumnWidth = 50
    $ws.Range('A1').Value2 = '設定'
    $ws.Range('A1').Font.Size = 14
    $ws.Range('A1').Font.Bold = $true

    $cfg = @(
        @(3,  'ルートフォルダ',                 '',        '空欄 = このブックがあるフォルダ。共有先に合わせて変えられます'),
        @(4,  '画像フォルダ名',                 'img',     '表示用 JPG の置き場（ブックからの相対）'),
        @(5,  '原本フォルダ名',                 'pdf',     '原本 PDF の置き場（ブックからの相対）'),
        @(6,  'ホーム（工場マップ）のノードID', 'MAP-001', '［ホーム］ボタンで表示する項目'),
        @(7,  'ポインター既定サイズ 横(0-1)',   0.08,      '新規ポインターの初期の大きさ'),
        @(8,  'ポインター既定サイズ 縦(0-1)',   0.06,      ''),
        @(9,  'ポインターの塗りの薄さ(0-1)',    0.85,      '既定値。大きいほど薄い。0=ベタ塗り、0.9=ほぼ透明。個別に変えるには［濃さを変更］か［リンク］の「塗りの薄さ」'),
        @(10, 'ポインターの文字サイズ(pt)',     11,        '既定値。個別に変えるには［文字サイズを変更］か［リンク］の「文字サイズ」'),
        @(11, '現在表示中のノードID',           'MAP-001', '（自動で更新されます）'),
        @(12, '表示履歴',                       '',        '（自動で更新されます。［戻る］が使います）'),
        @(13, '編集モード(1=ON)',               '0',       '（自動で更新されます）')
    )
    foreach ($c in $cfg) {
        Set-Cell $ws ([int]$c[0]) 1 $c[1]
        Set-Cell $ws ([int]$c[0]) 2 $c[2]
        Set-Cell $ws ([int]$c[0]) 3 $c[3]
    }
    $ws.Range('A3:A13').Font.Bold = $true
    $ws.Range('B3:B13').Interior.Color = (RGBv 255 251 230)
    $ws.Range('B3:B13').Borders.LineStyle = $xlContinuous
    $ws.Range('B3:B13').Borders.Color = $CLR_FRAME
    $ws.Range('C3:C13').Font.Color = $CLR_MUTED
    $ws.Range('A15').Value2 = '※ 灰色の説明どおりに使ってください。B11〜B13 はツールが自動で書き換えます。'
    $ws.Range('A15').Font.Color = $CLR_MUTED

    $names = @{
        'cfgRoot' = '$B$3'; 'cfgImgDir' = '$B$4'; 'cfgPdfDir' = '$B$5'; 'cfgHome' = '$B$6'
        'cfgHsW' = '$B$7'; 'cfgHsH' = '$B$8'; 'cfgHsAlpha' = '$B$9'; 'cfgHsFont' = '$B$10'
        'cfgCurrent' = '$B$11'; 'cfgHistory' = '$B$12'; 'cfgEdit' = '$B$13'
    }
    foreach ($k in $names.Keys) { $wb.Names.Add($k, "=設定!$($names[$k])") | Out-Null }

    # =====================================================================
    # ノードシート
    # =====================================================================
    Write-Host '  シート: ノード'
    $ws = $wb.Sheets.Item('ノード')
    $ws.Cells.Font.Name = $FONT
    $ws.Cells.Font.Size = 10
    $ws.Range('A1').Value2 = 'ノード台帳 ── 1 行 = 1 つの対象（マップ／設備／詳細／組図／部品）。親子関係は［リンク］シートで管理します。'
    $ws.Range('A1').Font.Bold = $true

    $nhdr = @('ノードID', '種別', '名称', '表示用画像', '原本ファイル', '型式・品番', 'メーカー',
              '棚番', '在庫数', '単価', '購入先', '備考', '更新日')
    for ($j = 0; $j -lt $nhdr.Count; $j++) { Set-Cell $ws 2 ($j + 1) $nhdr[$j] }

    $today = (Get-Date -Format 'yyyy/MM/dd')
    $nodes = @(
        @('MAP-001', 'マップ',   '工場マップ（1F 成形ライン）', 'img\map_factory.jpg', '',              '',       '',       '',      '', '',    '',       'ポインターをクリックすると設備へ移動します', $today),
        @('EQ-001',  '設備',     '混錬機A',                     'img\eq_mixer.jpg',    '',              '',       '',       '',      '', '',    '',       '',                                           $today),
        @('EQ-002',  '設備',     'コンベアB',                   'img\eq_conveyor.jpg', '',              '',       '',       '',      '', '',    '',       '子は未登録（サンプル）',                     $today),
        @('DT-001',  '詳細',     '混錬機A 減速機部',            'img\dt_gearbox.jpg',  '',              '',       '',       '',      '', '',    '',       '',                                           $today),
        @('AS-001',  '組図',     '減速機 組立図 A-100',         'img\as_a100.jpg',     'pdf\A-100.pdf', 'A-100',  '',       '',      '', '',    '',       'JPG=表示用、PDF=原本',                       $today),
        @('PT-001',  '購入部品', 'ベアリング',                  'img\pt_bearing.jpg',  '',              '6205ZZ', 'NSK',    'A-3-2', 4,  1800,  '○○商会', '',                                           $today),
        @('PT-002',  '購入部品', 'Oリング',                     'img\pt_oring.jpg',    '',              'S-40',   'NOK',    'B-1-5', 12, 120,   '○○商会', '',                                           $today),
        @('DW-001',  '部品図面', '出力シャフト',                'img\dw_shaft.jpg',    'pdf\D-021.pdf', 'D-021',  '自社製作', 'C-2-1', 1,  42000, '△△鉄工', '要納期2週間',                                $today)
    )
    for ($i = 0; $i -lt $nodes.Count; $i++) {
        for ($j = 0; $j -lt $nhdr.Count; $j++) { Set-Cell $ws (3 + $i) ($j + 1) $nodes[$i][$j] }
    }

    $lo = $ws.ListObjects.Add($xlSrcRange, $ws.Range('A2:M' + (2 + $nodes.Count)), $null, $xlYes)
    $lo.Name = 'tblNode'
    $lo.TableStyle = 'TableStyleLight9'

    $w = @(10, 10, 26, 24, 22, 14, 12, 10, 8, 10, 14, 30, 12)
    for ($j = 0; $j -lt $w.Count; $j++) { $ws.Columns.Item($j + 1).ColumnWidth = $w[$j] }
    $ws.Range('J3:J500').NumberFormat = '#,##0'
    $ws.Rows.Item(2).RowHeight = 20

    $v = $ws.Range('B3:B500').Validation
    $v.Delete()
    $v.Add($xlValidateList, $xlValidAlertStop, $xlBetween, 'マップ,設備,詳細,組図,部品図面,購入部品') | Out-Null

    # =====================================================================
    # リンクシート
    # =====================================================================
    Write-Host '  シート: リンク'
    $ws = $wb.Sheets.Item('リンク')
    $ws.Cells.Font.Name = $FONT
    $ws.Cells.Font.Size = 10
    $ws.Range('A1').Value2 = '親子関係とポインター位置 ── X/Y/W/H は画像に対する 0〜1 の割合。空欄なら「位置未設定」（子一覧にだけ出ます）。ラベルはポインターに出す文字（空欄なら通し番号）。塗りの薄さ・文字色・文字サイズは空欄なら既定（文字色は RRGGBB か 黒/白/赤/青/緑/黄/橙）。'
    $ws.Range('A1').Font.Bold = $true

    $lhdr = @('リンクID', '親ノードID', '子ノードID', 'ラベル', 'X', 'Y', 'W', 'H', '形状', '数量', '備考', '塗りの薄さ', '文字色', '文字サイズ')
    for ($j = 0; $j -lt $lhdr.Count; $j++) { Set-Cell $ws 2 ($j + 1) $lhdr[$j] }

    $links = @(
        @('L0001', 'MAP-001', 'EQ-001',  '1', 0.12, 0.30, 0.18, 0.22, '四角', 1, '', '', '', ''),
        @('L0002', 'MAP-001', 'EQ-002',  '2', 0.55, 0.55, 0.22, 0.18, '四角', 1, '', '', '', ''),
        @('L0003', 'EQ-001',  'DT-001',  '1', 0.60, 0.35, 0.18, 0.22, '四角', 1, '', '', '', ''),
        @('L0004', 'DT-001',  'AS-001',  '1', 0.30, 0.30, 0.30, 0.30, '四角', 1, '', '', '', ''),
        @('L0005', 'AS-001',  'PT-001',  '1', 0.18, 0.28, 0.06, 0.09, '丸',   2, '', '', '', ''),
        @('L0006', 'AS-001',  'PT-002',  '2', 0.46, 0.55, 0.06, 0.09, '丸',   4, '', '', '', ''),
        @('L0007', 'AS-001',  'DW-001',  '3', 0.70, 0.30, 0.06, 0.09, '丸',   1, '', '', '', '')
    )
    for ($i = 0; $i -lt $links.Count; $i++) {
        for ($j = 0; $j -lt $lhdr.Count; $j++) { Set-Cell $ws (3 + $i) ($j + 1) $links[$i][$j] }
    }

    $lo = $ws.ListObjects.Add($xlSrcRange, $ws.Range('A2:N' + (2 + $links.Count)), $null, $xlYes)
    $lo.Name = 'tblLink'
    $lo.TableStyle = 'TableStyleLight10'

    $w = @(10, 13, 13, 9, 9, 9, 9, 9, 8, 7, 24, 11, 9, 11)
    for ($j = 0; $j -lt $w.Count; $j++) { $ws.Columns.Item($j + 1).ColumnWidth = $w[$j] }
    $ws.Range('E3:H500').NumberFormat = '0.0000'
    $ws.Range('L3:L500').NumberFormat = '0.00'
    $ws.Range('M3:M500').NumberFormat = '@'
    $ws.Range('N3:N500').NumberFormat = '0'
    $ws.Rows.Item(2).RowHeight = 20

    $v = $ws.Range('I3:I500').Validation
    $v.Delete()
    $v.Add($xlValidateList, $xlValidAlertStop, $xlBetween, '四角,丸') | Out-Null

    # =====================================================================
    # ビューシート
    # =====================================================================
    Write-Host '  シート: ビュー'
    $ws = $wb.Sheets.Item('ビュー')
    $ws.Cells.Font.Name = $FONT
    $ws.Cells.Font.Size = 10
    $ws.Columns.Item('A').ColumnWidth = 1.5
    # 画像枠 B5:Q36（16列×32行）。列幅 9.5 × 行高 18 なので縦横比 3:2
    #   ※ 枠の番地を変えるときは src\M_View.bas の FRAME_ADDR / LBL_COL 等も合わせること
    for ($c = 2; $c -le 17; $c++) { $ws.Columns.Item($c).ColumnWidth = 9.5 }
    $ws.Columns.Item('R').ColumnWidth = 1.2
    $ws.Columns.Item('S').ColumnWidth = 13
    $ws.Columns.Item('T').ColumnWidth = 15
    $ws.Columns.Item('U').ColumnWidth = 13
    $ws.Columns.Item('V').ColumnWidth = 1.2
    $ws.Columns.Item('W:X').Hidden = $true
    $ws.Rows.Item(1).RowHeight = 22
    $ws.Rows.Item(2).RowHeight = 26
    $ws.Rows.Item(3).RowHeight = 26
    $ws.Rows.Item(4).RowHeight = 26
    $ws.Range('5:60').RowHeight = 18

    # パンくず／モード表示
    $ws.Range('B1:Q1').Merge()
    $ws.Range('B1').Font.Size = 12
    $ws.Range('B1').Font.Bold = $true
    $ws.Range('B1').VerticalAlignment = $xlVCenter
    $ws.Range('S1:U1').Merge()
    $ws.Range('S1').Font.Size = 9
    $ws.Range('S1').Font.Color = $CLR_MUTED
    $ws.Range('S1').VerticalAlignment = $xlVCenter
    $ws.Range('S1').HorizontalAlignment = $xlRight

    # 画像枠
    $fr = $ws.Range('B5:Q36')
    $fr.Interior.Color = (RGBv 252 252 250)
    $fr.BorderAround($xlContinuous, $xlThin, [Type]::Missing, $CLR_FRAME) | Out-Null

    # 右パネル（S:U 列）
    Set-Header $ws 'S5:U5' '■ この項目'
    $labels = @('ノードID', '種別', '名称', '型式・品番', 'メーカー', '棚番', '在庫数',
                '単価', '購入先', '備考', '表示用画像', '原本ファイル', '使用先（親）')
    for ($i = 0; $i -lt $labels.Count; $i++) {
        $r = 6 + $i
        $ws.Cells.Item($r, 19).Value2 = $labels[$i]
        $ws.Cells.Item($r, 19).Font.Color = $CLR_MUTED
        $ws.Range("T$r`:U$r").Merge()
        $ws.Range("T$r").HorizontalAlignment = $xlLeft
    }
    $ws.Range('S6:U18').Interior.Color = $CLR_PANEL
    $ws.Range('S6:U18').Borders.LineStyle = $xlContinuous
    $ws.Range('S6:U18').Borders.Color = (RGBv 220 220 216)
    # 棚番を目立たせる
    $ws.Range('S11:U11').Interior.Color = (RGBv 255 244 244)
    $ws.Range('T11').Font.Bold = $true
    $ws.Range('T11').Font.Size = 12
    $ws.Range('T11').Font.Color = (RGBv 200 30 40)
    $ws.Range('T16:U17').Font.Size = 8
    $ws.Range('T16:U17').Font.Color = $CLR_MUTED

    Set-Header $ws 'S20:U20' '■ 子（クリックで移動）' $CLR_HEAD2
    $ws.Cells.Item(21, 19).Value2 = 'No'
    $ws.Cells.Item(21, 20).Value2 = '名称（クリックで移動）'
    $ws.Cells.Item(21, 21).Value2 = '種別・棚番'
    $ws.Range('S21:U21').Font.Bold = $true
    $ws.Range('S21:U21').Font.Color = $CLR_MUTED
    $ws.Range('S21:U21').Borders.Item(9).LineStyle = $xlContinuous   # xlEdgeBottom
    $ws.Range('S22:S60').HorizontalAlignment = $xlCenter
    $ws.Range('S22:S60').Font.Bold = $true
    $ws.Range('U22:U60').Font.Size = 9
    $ws.Range('U22:U60').Font.Color = $CLR_MUTED

    # ボタン
    $top1 = $ws.Range('B2').Top + 2
    $top2 = $ws.Range('B3').Top + 2
    $top3 = $ws.Range('B4').Top + 2
    $left = $ws.Range('B2').Left
    $btnH = 22
    $b = @(
        @('BTN_BACK',   '← 戻る',           'GoBack',              70),
        @('BTN_HOME',   'ホーム',            'GoHome',              70),
        @('BTN_PDF',    '原本を開く',        'OpenOriginal',        100),
        @('BTN_PARTS',  '棚番検索',          'GoPartsSheet',        90)
    )
    $x = $left
    foreach ($t in $b) {
        Add-Button $ws $t[0] $t[1] $t[2] $x $top1 $t[3] $btnH | Out-Null
        $x += $t[3] + 6
    }
    $b2 = @(
        @('BTN_EDIT',   '編集モード：OFF',   'ToggleEditMode',       120),
        @('BTN_ADD',    'ポインター追加',    'AddHotspot',           110),
        @('BTN_SAVE',   '位置を保存',        'SaveHotspotsAndRefresh', 100),
        @('BTN_NEW',    '子として新規登録',  'GoRegisterFromView',   140),
        @('BTN_LINK',   '既存を子に追加',    'LinkExistingChild',    134)
    )
    $x = $left
    foreach ($t in $b2) {
        Add-Button $ws $t[0] $t[1] $t[2] $x $top2 $t[3] $btnH | Out-Null
        $x += $t[3] + 6
    }
    # 3 段目: 選んだポインターの見た目を変える
    $b3 = @(
        @('BTN_ALPHA',  '濃さを変更',        'SetHotspotAlpha',      100),
        @('BTN_LABEL',  '表示名を変更',      'SetHotspotLabel',      100),
        @('BTN_COLOR',  '文字色を変更',      'SetHotspotTextColor',  100),
        @('BTN_FONT',   '文字サイズを変更',  'SetHotspotFontSize',   120)
    )
    $x = $left
    foreach ($t in $b3) {
        Add-Button $ws $t[0] $t[1] $t[2] $x $top3 $t[3] $btnH | Out-Null
        $x += $t[3] + 6
    }

    # =====================================================================
    # 部品一覧シート
    # =====================================================================
    Write-Host '  シート: 部品一覧'
    $ws = $wb.Sheets.Item('部品一覧')
    $ws.Cells.Font.Name = $FONT
    $ws.Cells.Font.Size = 10
    $w = @(10, 24, 16, 12, 12, 8, 10, 14, 26, 9)
    for ($j = 0; $j -lt $w.Count; $j++) { $ws.Columns.Item($j + 1).ColumnWidth = $w[$j] }
    $ws.Range('A1').Value2 = '予備品一覧（棚番検索）'
    $ws.Range('A1').Font.Size = 14
    $ws.Range('A1').Font.Bold = $true
    $ws.Range('A2').Value2 = '［ノード］シートの「購入部品」「部品図面」を自動で抽出します。編集は［ノード］または［登録］シートで行ってください。'
    $ws.Range('A2').Font.Color = $CLR_MUTED

    $ws.Range('A3').Value2 = '検索'
    $ws.Range('A3').Font.Bold = $true
    $ws.Range('B3:C3').Merge()
    $ws.Range('B3').Interior.Color = (RGBv 255 251 230)
    $ws.Range('B3').Borders.LineStyle = $xlContinuous
    $ws.Range('B3').Borders.Color = $CLR_FRAME
    $ws.Range('D3').HorizontalAlignment = $xlLeft
    $ws.Range('D3').Font.Color = $CLR_MUTED
    $ws.Rows.Item(3).RowHeight = 22
    $wb.Names.Add('partsQuery', '=部品一覧!$B$3') | Out-Null
    $wb.Names.Add('partsCount', '=部品一覧!$D$3') | Out-Null

    $phdr = @('ノードID', '品名', '型式・品番', 'メーカー', '棚番', '在庫', '単価', '購入先', '使用先（親）', '開く')
    for ($j = 0; $j -lt $phdr.Count; $j++) { $ws.Cells.Item(5, $j + 1).Value2 = $phdr[$j] }
    $ws.Range('A5:J5').Font.Bold = $true
    $ws.Range('A5:J5').Font.Color = $CLR_WHITE
    $ws.Range('A5:J5').Interior.Color = $CLR_HEAD
    $ws.Rows.Item(5).RowHeight = 20
    $ws.Range('E6:E500').Font.Bold = $true
    $ws.Range('E6:E500').Font.Color = (RGBv 200 30 40)
    $ws.Range('G6:G500').NumberFormat = '#,##0'
    $ws.Range('J6:J500').Font.Color = (RGBv 0 90 200)
    $ws.Range('J6:J500').HorizontalAlignment = $xlCenter
    $ws.Range('A5:J500').Borders.LineStyle = $xlContinuous
    $ws.Range('A5:J500').Borders.Color = (RGBv 220 220 216)

    $tp = $ws.Range('F3').Top + 1
    Add-Button $ws 'BTN_SEARCH' '検索'   'PartsSearch'  ($ws.Range('F3').Left)       $tp 70 20 | Out-Null
    Add-Button $ws 'BTN_ALL'    '全表示' 'PartsShowAll' ($ws.Range('F3').Left + 76)  $tp 70 20 | Out-Null
    Add-Button $ws 'BTN_TOVIEW' 'ビューへ' 'GoViewSheet' ($ws.Range('F3').Left + 152) $tp 70 20 | Out-Null

    # =====================================================================
    # 登録シート
    # =====================================================================
    Write-Host '  シート: 登録'
    $ws = $wb.Sheets.Item('登録')
    $ws.Cells.Font.Name = $FONT
    $ws.Cells.Font.Size = 10
    $ws.Columns.Item('A').ColumnWidth = 24
    $ws.Columns.Item('B').ColumnWidth = 46
    $ws.Columns.Item('C').ColumnWidth = 2
    $ws.Columns.Item('D').ColumnWidth = 22
    $ws.Columns.Item('E').ColumnWidth = 40
    $ws.Range('A1').Value2 = '新規登録'
    $ws.Range('A1').Font.Size = 14
    $ws.Range('A1').Font.Bold = $true

    $rlab = @('種別 ※必須', '名称 ※必須', '型式・品番', 'メーカー', '棚番', '在庫数', '単価',
              '購入先', '備考', '表示用画像（JPG）', '原本ファイル（PDF等）')
    for ($i = 0; $i -lt $rlab.Count; $i++) {
        $r = 3 + $i
        $ws.Cells.Item($r, 1).Value2 = $rlab[$i]
        $ws.Cells.Item($r, 1).Font.Color = $CLR_MUTED
        $ws.Rows.Item($r).RowHeight = 20
    }
    $ws.Range('A3:A4').Font.Bold = $true
    $ws.Range('A3:A4').Font.Color = (RGBv 200 30 40)
    $ws.Range('B3:B13').Interior.Color = (RGBv 255 251 230)
    $ws.Range('B3:B13').Borders.LineStyle = $xlContinuous
    $ws.Range('B3:B13').Borders.Color = $CLR_FRAME
    $ws.Range('B12:B13').Font.Size = 9

    $ws.Range('A15').Value2 = '親（ビューで表示中）'
    $ws.Range('A15').Font.Color = $CLR_MUTED
    $ws.Range('B15').Value2 = '（未選択）'
    $ws.Range('B15').Font.Bold = $true

    $rn = @{ 'regKind' = 3; 'regName' = 4; 'regModel' = 5; 'regMaker' = 6; 'regShelf' = 7
             'regStock' = 8; 'regPrice' = 9; 'regVendor' = 10; 'regNote' = 11
             'regImg' = 12; 'regPdf' = 13; 'regParent' = 15 }
    foreach ($k in $rn.Keys) { $wb.Names.Add($k, "=登録!`$B`$$($rn[$k])") | Out-Null }

    $v = $ws.Range('B3').Validation
    $v.Delete()
    $v.Add($xlValidateList, $xlValidAlertStop, $xlBetween, 'マップ,設備,詳細,組図,部品図面,購入部品') | Out-Null

    Add-Button $ws 'BTN_PICKIMG' '画像を選ぶ...' 'PickImage' ($ws.Range('D12').Left) ($ws.Range('D12').Top + 1) 110 18 | Out-Null
    Add-Button $ws 'BTN_PICKPDF' '原本を選ぶ...' 'PickPdf'   ($ws.Range('D13').Left) ($ws.Range('D13').Top + 1) 110 18 | Out-Null

    $tp = $ws.Range('A17').Top
    $lf = $ws.Range('A17').Left
    Add-Button $ws 'BTN_REGCHILD' '表示中の項目の子として登録' 'RegisterAsChild'    $lf         $tp 190 26 $CLR_BTNPRI $CLR_WHITE | Out-Null
    Add-Button $ws 'BTN_REGONE'   '単独で登録（親なし）'       'RegisterStandalone' ($lf + 196) $tp 150 26 | Out-Null
    Add-Button $ws 'BTN_REGCLR'   '入力をクリア'               'ClearRegister'      ($lf + 352) $tp 100 26 | Out-Null
    Add-Button $ws 'BTN_REGBACK'  'ビューへ戻る'               'GoViewSheet'        ($lf + 458) $tp 100 26 | Out-Null

    $help = @(
        '【使い方】',
        '1. 子を追加したい親（工場マップ・設備・組図など）を［ビュー］で表示しておく。',
        '2. このシートで 種別 と 名称 を入れ、必要なら 棚番・型式・在庫などを入力する。',
        '3. ［画像を選ぶ...］で表示用 JPG を、［原本を選ぶ...］で PDF を指定する（自動で img\ pdf\ にコピーされます）。',
        '4. ［表示中の項目の子として登録］を押すと、親の画像の中央にポインターが追加され、編集モードになります。',
        '5. ［ビュー］でポインターをドラッグして位置を合わせ、［位置を保存］を押す。',
        '',
        '・組図など PDF が原本のものは、PDF ビューアの「画像として保存」や図面ソフトの書き出しで JPG を作り、',
        '  「表示用画像」に JPG、「原本ファイル」に PDF を指定してください。ポインターは JPG の上に置きます。',
        '・棚番を管理したいものは 種別 を「購入部品」または「部品図面」にすると［部品一覧］に出ます。'
    )
    for ($i = 0; $i -lt $help.Count; $i++) {
        $ws.Cells.Item(20 + $i, 1).Value2 = $help[$i]
    }
    $ws.Range('A20:A29').Font.Color = $CLR_MUTED
    $ws.Range('A20').Font.Bold = $true
    $ws.Range('A20').Font.Color = $CLR_TEXT

    # =====================================================================
    # VBA 流し込み
    # =====================================================================
    Write-Host 'VBA を組み込んでいます...'
    $vbp = $wb.VBProject

    function Set-Code($comp, [string]$file) {
        $code = [System.IO.File]::ReadAllText($file, [System.Text.Encoding]::UTF8)
        $code = (($code -split "`r?`n") | Where-Object { $_ -notmatch '^Attribute\s' }) -join "`r`n"
        if ($comp.CodeModule.CountOfLines -gt 0) {
            $comp.CodeModule.DeleteLines(1, $comp.CodeModule.CountOfLines)
        }
        $comp.CodeModule.AddFromString($code)
    }

    foreach ($m in @('M_Data', 'M_View', 'M_Hotspot', 'M_Parts', 'M_Register')) {
        $c = $vbp.VBComponents.Add(1)      # vbext_ct_StdModule
        $c.Name = $m
        Set-Code $c (Join-Path $Src "$m.bas")
    }
    Set-Code $vbp.VBComponents.Item($wb.CodeName)                     (Join-Path $Src 'ThisWorkbook.cls')
    Set-Code $vbp.VBComponents.Item($wb.Sheets.Item('ビュー').CodeName)   (Join-Path $Src 'Sheet_View.cls')
    Set-Code $vbp.VBComponents.Item($wb.Sheets.Item('部品一覧').CodeName) (Join-Path $Src 'Sheet_Parts.cls')

    # =====================================================================
    # 仕上げ
    # =====================================================================
    foreach ($n in @('設定', '登録', '部品一覧', 'ビュー')) {
        $wb.Sheets.Item($n).Activate()
        try { $xl.ActiveWindow.DisplayGridlines = $false } catch { }
        try { $wb.Sheets.Item($n).Range('A1').Select() | Out-Null } catch { }
    }
    # 部品一覧は見出しを固定
    try {
        $wb.Sheets.Item('部品一覧').Activate()
        $wb.Sheets.Item('部品一覧').Range('A6').Select() | Out-Null
        $xl.ActiveWindow.FreezePanes = $true
    } catch { }
    $wb.Sheets.Item('ビュー').Activate()
    try { $wb.Sheets.Item('ビュー').Range('B3').Select() | Out-Null } catch { }

    $wb.SaveAs($Out, $xlOpenXMLWorkbookMacroEnabled)
    $wb.Close($false)
    Write-Host ''
    Write-Host "作成しました: $Out" -ForegroundColor Green
}
finally {
    $xl.EnableEvents = $true
    $xl.ScreenUpdating = $true
    $xl.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    [GC]::Collect()
}
