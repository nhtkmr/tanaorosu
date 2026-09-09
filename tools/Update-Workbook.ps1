# =========================================================
#  Update-Workbook.ps1
#  入力済みの台帳を残したまま、ブックを最新の画面・VBA に入れ替える
#    1. 今のブックから ［ノード］［リンク］［設定］ を読み出す
#    2. 控えを backup\ に取る
#    3. Build-Workbook.ps1 で作り直す（＝画面もボタンも VBA も最新）
#    4. 読み出した台帳を書き戻す
#  ※ 画像 (img\) と原本 (pdf\) は触りません
# =========================================================
$ErrorActionPreference = 'Stop'

$Root = Split-Path -Parent $PSScriptRoot
$Out = Join-Path $Root '予備品図面管理.xlsm'

if (-not (Test-Path $Out)) {
    Write-Host 'ブックがないので新規に作成します。'
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Build-Workbook.ps1')
    exit $LASTEXITCODE
}
try {
    $fs = [System.IO.File]::Open($Out, 'Open', 'ReadWrite', 'None'); $fs.Close()
} catch {
    throw "$Out を更新できません。Excel で開いている場合は閉じてから再実行してください。"
}

function New-Excel {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false; $xl.DisplayAlerts = $false
    $xl.EnableEvents = $false; $xl.ScreenUpdating = $false
    $xl.AutomationSecurity = 1
    return $xl
}

# ---------- 1. 今の台帳を読み出す ----------
Write-Host '今のブックから台帳を読み出しています...'
$xl = New-Excel
$nodes = @(); $links = @(); $cfg = @()
try {
    $wb = $xl.Workbooks.Open($Out)
    foreach ($t in @(@('ノード', 'tblNode'), @('リンク', 'tblLink'))) {
        $ws = $wb.Sheets.Item($t[0])
        $lo = $ws.ListObjects.Item($t[1])
        $rows = @()
        if ($lo.DataBodyRange) {
            $v = $lo.DataBodyRange.Value2
            $nr = $lo.DataBodyRange.Rows.Count
            $nc = $lo.ListColumns.Count
            for ($i = 1; $i -le $nr; $i++) {
                $row = @()
                for ($j = 1; $j -le $nc; $j++) { $row += $v.GetValue($i, $j) }
                # 完全に空の行は捨てる
                if (($row | Where-Object { "$_" -ne '' } | Measure-Object).Count -gt 0) { $rows += , $row }
            }
        }
        if ($t[1] -eq 'tblNode') { $nodes = $rows } else { $links = $rows }
    }
    $wsC = $wb.Sheets.Item('設定')
    for ($r = 3; $r -le 11; $r++) { $cfg += $wsC.Cells.Item($r, 2).Value2 }
}
finally {
    try { $wb.Close($false) } catch { }
    $xl.Quit(); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
Write-Host ("  ノード {0} 行 / リンク {1} 行" -f $nodes.Count, $links.Count)

# ---------- 2. 作り直す（Build 側で backup\ に控えが取られる） ----------
Write-Host ''
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Build-Workbook.ps1')
if ($LASTEXITCODE -ne 0) { throw 'ビルドに失敗しました。backup\ の控えを使ってください。' }

# ---------- 3. 台帳を書き戻す ----------
Write-Host ''
Write-Host '台帳を書き戻しています...'
$xl = New-Excel
try {
    $wb = $xl.Workbooks.Open($Out)

    function Write-Table($ws, $lo, $rows) {
        $first = $lo.HeaderRowRange.Row + 1
        $col = $lo.Range.Column
        $nc = $lo.ListColumns.Count
        $oldLast = $lo.HeaderRowRange.Row
        if ($lo.DataBodyRange) { $oldLast = $lo.DataBodyRange.Row + $lo.DataBodyRange.Rows.Count - 1 }

        for ($i = 0; $i -lt $rows.Count; $i++) {
            for ($j = 0; $j -lt $nc; $j++) {
                $v = $rows[$i][$j]
                $c = $ws.Cells.Item($first + $i, $col + $j)
                if ($null -eq $v -or "$v" -eq '') { $c.ClearContents() | Out-Null }
                elseif ($v -is [double] -or $v -is [int] -or $v -is [long] -or $v -is [decimal]) { $c.Value2 = [double]$v }
                else { $c.Value2 = [string]$v }
            }
        }
        $newLast = $first + [Math]::Max($rows.Count, 1) - 1
        $lo.Resize($ws.Range($ws.Cells.Item($lo.HeaderRowRange.Row, $col), $ws.Cells.Item($newLast, $col + $nc - 1)))
        if ($oldLast -gt $newLast) {
            $ws.Range($ws.Cells.Item($newLast + 1, $col), $ws.Cells.Item($oldLast, $col + $nc - 1)).ClearContents() | Out-Null
        }
    }

    Write-Table $wb.Sheets.Item('ノード') $wb.Sheets.Item('ノード').ListObjects.Item('tblNode') $nodes
    Write-Table $wb.Sheets.Item('リンク') $wb.Sheets.Item('リンク').ListObjects.Item('tblLink') $links

    # 設定は運用値（ルート・フォルダ名・ホーム・既定サイズ）だけ戻す
    $wsC = $wb.Sheets.Item('設定')
    for ($r = 3; $r -le 8; $r++) {
        $v = $cfg[$r - 3]
        if ($null -ne $v -and "$v" -ne '') {
            if ($v -is [double] -or $v -is [int]) { $wsC.Cells.Item($r, 2).Value2 = [double]$v }
            else { $wsC.Cells.Item($r, 2).Value2 = [string]$v }
        }
    }
    # 表示状態はリセット
    $wsC.Cells.Item(9, 2).Value2 = [string]$wsC.Cells.Item(6, 2).Value2   # 現在 = ホーム
    $wsC.Cells.Item(10, 2).ClearContents() | Out-Null                      # 履歴
    $wsC.Cells.Item(11, 2).Value2 = '0'                                    # 編集モード

    $wb.Save()
    $wb.Close($false)
    Write-Host ("書き戻しました: ノード {0} 行 / リンク {1} 行" -f $nodes.Count, $links.Count) -ForegroundColor Green
}
finally {
    $xl.EnableEvents = $true; $xl.ScreenUpdating = $true
    $xl.Quit(); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    [GC]::Collect()
}
