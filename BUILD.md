# ゼロからブックを作る手順（保守用）

`予備品図面管理.xlsm` は手で作ったものではなく、**`src\` のテキストから組み立てています**。
このファイルは、そのやり方と直し方の記録です。ツールを使うだけの人には不要です（→ `README.md`）。

---

## 1. 必要なもの

| | |
|---|---|
| OS | Windows 10 / 11 |
| Excel | デスクトップ版（Microsoft 365 / 2016 以降）。**Web 版・Mac 版では作れません** |
| PowerShell | Windows 標準の Windows PowerShell 5.1（追加インストール不要） |
| 作成・確認した環境 | Windows 11 Pro / Excel 16.0（64bit）/ 日本語ロケール（CP932） |

Python も外部ライブラリも使いません。

## 2. 事前設定（一度だけ）

Excel の設定を 1 か所だけ変えます。**PowerShell から VBA を書き込むために必須**です。

**ファイル → オプション → トラスト センター → トラスト センターの設定 → マクロの設定**
→ **「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」にチェック**

確認だけしたいときは PowerShell で:

```powershell
(Get-ItemProperty 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security').AccessVBOM   # 1 なら有効
```

`0` や空のままビルドすると、スクリプトが理由を表示して止まります。

## 3. 最短手順

`src\` と `tools\` を置いたフォルダを用意し、そこで PowerShell を開いて 2 行です。

```powershell
.\tools\New-SampleImages.ps1     # サンプル画像・PDF（任意。無くてもビルドは通る）
.\tools\Build-Workbook.ps1       # 予備品図面管理.xlsm ができる（40 秒ほど）
```

必要なのは次の 9 ファイルだけです。**これさえあれば同じブックが再現できます。**

```
tools\Build-Workbook.ps1
src\M_Data.bas  M_View.bas  M_Hotspot.bas  M_Parts.bas  M_Register.bas
src\ThisWorkbook.cls  Sheet_View.cls  Sheet_Parts.cls
```

画像が無い状態でも壊れず、ビューに「画像ファイルが見つかりません」の枠が出るだけです。
**入力済みの台帳は再現できません**（ブック本体か `backup\` にしかありません）。

## 4. 何をしているか

```
src\*.bas / *.cls  ─┐
シート・ボタン定義 ─┴→ Build-Workbook.ps1 →〔Excel を COM で自動操作〕→ .xlsm
   （Build-Workbook.ps1 の中）
```

1. Excel を非表示で起動し、新規ブックを作る
2. 6 シート（ビュー／ノード／リンク／部品一覧／登録／設定）を作り、列幅・見出し・表・入力規則・名前定義を設定
3. ボタンを図形として置き、`OnAction` にマクロ名を割り当てる
4. `src\` のコードを **`CodeModule.AddFromString`** で流し込む
   （`.bas` ファイルの Import ではなく文字列で渡す。COM 経由なら Unicode のまま届き、文字化けしない）
5. サンプル台帳（ノード 8 行・リンク 7 行）を書き込む
6. `.xlsm`（形式 52）で保存

## 5. ファイルの役割

### src（VBA。ここが正）

| ファイル | 中身 |
|---|---|
| `M_Data.bas` | 台帳へのアクセス、列位置の定数（`NC_*` `LC_*`）、ID 採番、パス変換、種別ごとの色 |
| `M_View.bas` | 画面描画（画像・情報欄・子一覧・パンくず）、履歴、戻る／ホーム、原本を開く、使用先へ移動 |
| `M_Hotspot.bas` | ポインターの描画・クリック・追加・保存、編集モード |
| `M_Parts.bas` | 部品一覧の生成と検索 |
| `M_Register.bas` | 新規登録、既存を子に追加、リンク行の作成 |
| `ThisWorkbook.cls` | 開いたとき／保存前後／閉じるときの処理 |
| `Sheet_View.cls` | ビューのクリック処理（子一覧・使用先） |
| `Sheet_Parts.cls` | 部品一覧のクリック・検索・再生成 |

### tools（PowerShell）

| ファイル | 役割 |
|---|---|
| `Build-Workbook.ps1` | ゼロから組み立て（**台帳はサンプルに戻る**） |
| `Update-Workbook.ps1` | 台帳を退避 → ビルド → 書き戻し（**通常はこちら**） |
| `New-SampleImages.ps1` | サンプル画像・PDF を生成 |
| `New-Package.ps1` | 配布用 ZIP を作る |
| `Convert-MdToPdf.ps1` | Markdown を PDF / HTML に変換 |

## 6. 直したいときの早見表

| やりたいこと | 直す場所 |
|---|---|
| ボタンを増やす・名前を変える | `Build-Workbook.ps1` の `$b`（上段）/ `$b2`（下段）。処理本体は `src\` に追加 |
| 情報欄の項目を増やす | `Build-Workbook.ps1` の `$labels` ＋ `M_View.bas` の `IR_*` 定数と `FillInfo` |
| ノード表に列を足す | `M_Data.bas` の `NC_*` と `NC_LAST` ＋ `Build-Workbook.ps1` の `$nhdr` と列幅 |
| 画像枠の大きさ | `M_View.bas` の `FRAME_ADDR` ＋ `Build-Workbook.ps1` のビューシートの列幅・行高 |
| ポインターの色 | `M_Data.bas` の `KindColor`（ラベル文字はそれを暗くした `Darken`） |
| ポインターの塗りの薄さ | **コード修正不要**。［設定］シートの「ポインターの塗りの薄さ(0-1)」。既定 0.85（大きいほど薄い） |
| 種別（マップ／設備…）を増やす | `M_Data.bas` の `KindPrefix` と `KindColor` ＋ `Build-Workbook.ps1` の入力規則 2 か所（ノード列 B・登録 B3） |
| サンプル台帳の内容 | `Build-Workbook.ps1` の `$nodes` / `$links` |
| ホーム・フォルダ名・既定サイズ・塗りの薄さ | **コード修正不要**。ブックの［設定］シートで変えられます |

直したら `.\tools\Update-Workbook.ps1`（台帳が残る）で反映します。

## 7. ハマりどころ（実際に踏んだもの）

1. **`.ps1` は UTF-8 **BOM 付き** で保存する**
   Windows PowerShell 5.1 は BOM が無いと CP932 として読むため、日本語が化けて構文エラーになります。
   ```powershell
   $t=[IO.File]::ReadAllText($p,(New-Object Text.UTF8Encoding($false)))
   [IO.File]::WriteAllText($p,$t,(New-Object Text.UTF8Encoding($true)))
   ```
2. **VBA の変数名・引数名を、手続き名や VBA 組込関数と同じにしない**
   大文字小文字は区別されないため、`Sub PushHistory` がある所で引数 `pushHistory` を作ると手続きが隠れます。
   `abs` `left` `date` など組込関数名も同じ。どちらも
   **「コンパイル エラー: Sub、Function、または Property が必要です」** になります（原因が分かりにくい）。
3. **VBA はモジュール単位の遅延コンパイル**
   あるマクロが動いても、別モジュールが壊れている可能性があります。1 つ動いたから OK とは判断しないこと。
4. **結合セルに `ClearContents` は使えない** → `.Value = ""` を使う（検索欄など）。
5. **Excel でブックを開いたままビルドすると失敗する**（スクリプトが検出して止まります）。
6. **PowerShell から Excel COM に数値を渡すとき、`Cells.Item(r,c).Value2 = $v` が型変換で落ちることがある**
   → 番地文字列（`Range("B7")`）＋明示キャストで書く（`Build-Workbook.ps1` の `Set-Cell`）。
7. **ネイティブ exe の stderr を `2>$null` などで受けない**
   PowerShell 5.1 は警告 1 行でもエラー扱いにします（Edge のヘッドレス印刷で発生）→ `Start-Process -Wait` を使う。
8. **VBE で直接 VBA を直さない**。次のビルドで消えます。直してしまったら
   そのモジュールを右クリック →「ファイルのエクスポート」で `src\` に書き戻してください。

## 8. 動作確認

ビルド後、ブックを開いて次を通せば主要機能をひととおり確認できます。

1. 開いた直後に工場マップが表示され、設備のポインターが 2 つ出る
2. ポインター → 設備 → 詳細 → 組図 → 部品 と降り、［← 戻る］で 1 段ずつ戻る
3. 組図で［原本を開く］→ PDF が既定のビューアで開く
4. ［編集モード］→ ポインターをドラッグ → ［位置を保存］→ 別のノードへ行って戻ると位置が残っている
5. ［既存を子に追加］→ 名前で検索 → 選ぶ → **中央にポインターが出る**（出ない場合は不具合）
6. ［部品一覧］で棚番を検索 → ［▶ 表示］でその部品へ飛ぶ
7. フォルダごと別の場所へコピーして開いても画像が出る（相対パスの確認）

自動での確認は、PowerShell から `Application.Run` でマクロを直接呼ぶ形で書けます
（`ShowNode` `RebuildParts` `AddChildLink` `SaveHotspots` など）。
`MsgBox` や `InputBox` を出す処理は自動化できないので、ロジックを別関数に分けて呼びます。
