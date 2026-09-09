# =========================================================
#  New-Package.ps1
#  配布用の一式（ブック + 画像 + 原本 + 手引き）を ZIP にまとめる
#    dist\予備品図面管理_yyyyMMdd.zip
#  ※ src\ tools\ backup\ は入れません（配布先では不要）
# =========================================================
param(
    [switch]$NoSample   # サンプル画像・サンプル PDF を除いて配る場合
)
$ErrorActionPreference = 'Stop'

$Root = Split-Path -Parent $PSScriptRoot
$Book = Join-Path $Root '予備品図面管理.xlsm'
if (-not (Test-Path $Book)) { throw "ブックがありません: $Book" }

$stamp = Get-Date -Format 'yyyyMMdd'
$Dist = Join-Path $Root 'dist'
$Stage = Join-Path $Dist ('予備品図面管理_' + $stamp)
$Zip = $Stage + '.zip'

if (Test-Path $Stage) { Remove-Item $Stage -Recurse -Force }
if (Test-Path $Zip) { Remove-Item $Zip -Force }
New-Item -ItemType Directory -Path $Stage -Force | Out-Null

Copy-Item $Book $Stage
foreach ($d in @('img', 'pdf')) {
    $src = Join-Path $Root $d
    if (Test-Path $src) {
        $dst = Join-Path $Stage $d
        New-Item -ItemType Directory -Path $dst -Force | Out-Null
        $files = Get-ChildItem $src -File
        if ($NoSample) {
            $sample = @('map_factory.jpg', 'eq_mixer.jpg', 'eq_conveyor.jpg', 'dt_gearbox.jpg',
                        'as_a100.jpg', 'pt_bearing.jpg', 'pt_oring.jpg', 'dw_shaft.jpg',
                        'A-100.pdf', 'D-021.pdf')
            $files = $files | Where-Object { $sample -notcontains $_.Name }
        }
        foreach ($f in $files) { Copy-Item $f.FullName $dst }
    }
}
foreach ($f in @('README.pdf', 'README.md')) {
    $p = Join-Path $Root $f
    if (Test-Path $p) { Copy-Item $p $Stage }
}

$guide = @'
設備予備品・図面 階層ビューア　── はじめにお読みください

■ 置き場所
　このフォルダごと（予備品図面管理.xlsm と img・pdf フォルダを一緒に）
　共有フォルダへコピーしてください。
　中のファイルだけを別の場所へ移すと、画像へのリンクが切れます。

■ 最初に一度だけ（マクロを使えるようにする）
　1) 予備品図面管理.xlsm を右クリック → プロパティ
　   下のほうに「セキュリティ: このファイルは他のコンピューターから…」と出ていたら
　   「許可する」にチェック → OK
　2) Excel → ファイル → オプション → トラスト センター → トラスト センターの設定
　   → 信頼できる場所 → 新しい場所の追加 → このフォルダを指定
　   （サーバー上に置く場合は「ネットワーク上にある信頼できる場所を許可する」にもチェック）

　設定しない場合は、開いたときに出る黄色いバーの［コンテンツの有効化］を毎回押してください。

■ 使いかた
　予備品図面管理.xlsm を開く → 工場マップの上のポインターをクリックして
　設備 → 部位 → 組図 → 部品 とたどります。棚番は［部品一覧］シートで検索できます。
　詳しい手順は README.pdf をご覧ください。

■ 複数人で使うとき
　共有フォルダに 1 つ置いて、全員がそれを開きます（台帳が 1 つで済みます）。
　先に開いた人が編集でき、あとから開いた人は「読み取り専用」になります。
　読み取り専用でも、閲覧・移動・棚番検索はすべてできます。
　登録やポインター位置の保存だけができません（画面右上に表示されます）。
'@
[System.IO.File]::WriteAllText((Join-Path $Stage 'はじめにお読みください.txt'), $guide,
    (New-Object System.Text.UTF8Encoding($true)))

Compress-Archive -Path (Join-Path $Stage '*') -DestinationPath $Zip -CompressionLevel Optimal

$size = (Get-Item $Zip).Length
Write-Host ''
Write-Host ("配布パッケージを作りました: {0}  ({1:N0} bytes)" -f $Zip, $size) -ForegroundColor Green
Write-Host ("展開前のフォルダ           : {0}" -f $Stage)
Get-ChildItem $Stage -Recurse -File | ForEach-Object {
    "  " + $_.FullName.Substring($Stage.Length + 1)
}
