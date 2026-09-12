Option Explicit

'==========================================================
' ThisWorkbook : 起動時／保存時の処理
'==========================================================

Private Sub Workbook_Open()
    Dim id As String
    On Error Resume Next
    Application.StatusBar = False
    EnsureFolders
    ' 毎回、閲覧者モード（ビューと部品一覧だけ）から始める
    LockWorkbook
    HookSaveKey
    id = GetCfg("cfgCurrent")
    If Len(id) = 0 Then id = GetCfg("cfgHome")
    ShowNode id, False
End Sub

' Ctrl+S をこのブックの保存（SaveLocked）に差し替える。他のブックに移ったら戻す
Private Sub Workbook_Activate()
    HookSaveKey
End Sub

Private Sub Workbook_Deactivate()
    On Error Resume Next
    Application.OnKey "^s"
End Sub

' ファイルは常に閲覧者モードの状態で保存する。
'   閲覧者のときはそのまま保存（もともと隠れている）。
'   編集者のときは、BeforeSave の中ではシートを隠せない（Excel が無視する）ので
'   ここでは保存を止め、Ctrl+S／［保存］ボタン（SaveLocked）で保存してもらう。
'   画像とポインター図形はブックに保存しない（ファイルを軽く保つ）
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    On Error Resume Next
    If IsEditor() Then
        Cancel = True
        MsgBox "編集者モードのときは Ctrl+S か、ビューの［保存］ボタンで保存してください。" & vbCrLf & _
               "（台帳シートを隠した状態で保存するためです）" & vbCrLf & vbCrLf & _
               "名前を付けて保存したいときは、いったん閲覧者モードに戻してください。", _
               vbInformation, "保存"
        Exit Sub
    End If
    ClearViewShapes
End Sub

Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    On Error Resume Next
    ShowNode GetCfg("cfgCurrent"), False
End Sub

' 閉じるとき
'   読み取り専用（＝他の人が先に開いている／共有の閲覧用）で開いた場合、
'   画面を描くだけでブックが変更済みになってしまうので、保存を聞かない。
'   編集者のとき、ファイルは常に閲覧者の状態で保存済み（SaveLocked 経由）なので
'   ここで保存はさせない（BeforeClose の中でもシートを隠せないため）。
'   未保存の変更がありそうなら、保存せず閉じるか戻るかを聞く。
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    Application.OnKey "^s"
    If ThisWorkbook.ReadOnly Then
        ThisWorkbook.Saved = True
    ElseIf IsEditor() Then
        If Not ThisWorkbook.Saved Then
            If MsgBox("保存していない変更があるかもしれません。" & vbCrLf & vbCrLf & _
                      "［はい］　保存せずに閉じる" & vbCrLf & _
                      "［いいえ］戻って保存する（Ctrl+S または［保存］ボタン）", _
                      vbYesNo + vbQuestion + vbDefaultButton2, "閉じる") = vbNo Then
                Cancel = True
                HookSaveKey
                Exit Sub
            End If
        End If
        ThisWorkbook.Saved = True
    End If
End Sub
