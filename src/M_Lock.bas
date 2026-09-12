Attribute VB_Name = "M_Lock"
Option Explicit

'==========================================================
' M_Lock : 閲覧者／編集者モード
'   開いた直後は閲覧者（［ビュー］［部品一覧］だけ）。合言葉で編集者になると
'   台帳シート（ノード／リンク／登録／設定）と編集ボタンが出る。
'   保存は常に閲覧者の状態で行う（マクロ無効で開かれても台帳は見えない）。
'   ※ Excel の保護は「うっかり防止」用。VBE を開ける人には合言葉が見える
'==========================================================

Public Const LOCK_PW_DEFAULT As String = "edit"    ' 初期の合言葉。運用では［設定］で変える

Private Const HIDDEN_SHEETS As String = SH_NODE & "," & SH_LINK & "," & SH_REG & "," & SH_CFG
Private Const EDIT_BUTTONS As String = _
    "BTN_EDIT,BTN_ADD,BTN_SAVE,BTN_SAVEWB,BTN_NEW,BTN_LINK,BTN_ALPHA,BTN_LABEL,BTN_COLOR,BTN_FONT"

Public Function IsEditor() As Boolean
    IsEditor = (GetCfg("cfgEditor") = "1")
End Function

Private Function LockPassword() As String
    LockPassword = GetCfg("cfgLockPw")
    If Len(LockPassword) = 0 Then LockPassword = LOCK_PW_DEFAULT
End Function

'----------------------------------------------------------
' 閲覧者モードにする：台帳シートを隠し、ブック構成を保護し、編集ボタンを消す
'----------------------------------------------------------
Public Sub LockWorkbook()
    Dim v As Variant
    On Error Resume Next
    If GetCfg("cfgEdit") = "1" Then SaveHotspots
    SetCfg "cfgEdit", "0"
    SetCfg "cfgEditor", "0"

    ThisWorkbook.Unprotect LockPassword()
    For Each v In Split(HIDDEN_SHEETS, ",")
        ThisWorkbook.Worksheets(CStr(v)).Visible = xlSheetVeryHidden
    Next v
    ThisWorkbook.Protect Password:=LockPassword(), Structure:=True
    ShowEditButtons False
    On Error GoTo 0
End Sub

'----------------------------------------------------------
' 編集者モードにする：台帳シートを出し、編集ボタンを出す
'----------------------------------------------------------
Public Sub UnlockWorkbook()
    Dim v As Variant
    On Error Resume Next
    ThisWorkbook.Unprotect LockPassword()
    For Each v In Split(HIDDEN_SHEETS, ",")
        ThisWorkbook.Worksheets(CStr(v)).Visible = xlSheetVisible
    Next v
    SetCfg "cfgEditor", "1"
    ShowEditButtons True
    HookSaveKey
    On Error GoTo 0
End Sub

' Ctrl+S をこのブックの保存（SaveLocked）に差し替える
Public Sub HookSaveKey()
    On Error Resume Next
    Application.OnKey "^s", "SaveLocked"
End Sub

Private Sub ShowEditButtons(ByVal show As Boolean)
    Dim ws As Worksheet, v As Variant, sh As Shape
    Set ws = SheetOf(SH_VIEW)
    On Error Resume Next
    For Each v In Split(EDIT_BUTTONS, ",")
        ws.Shapes(CStr(v)).Visible = IIf(show, msoTrue, msoFalse)
    Next v
    Set sh = ws.Shapes("BTN_LOGIN")
    If Not sh Is Nothing Then
        sh.TextFrame2.TextRange.Text = IIf(show, "編集者モード：ON", "編集者モード：OFF")
        If show Then
            sh.Fill.ForeColor.RGB = RGB(0, 112, 192)
            sh.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        Else
            sh.Fill.ForeColor.RGB = RGB(242, 242, 242)
            sh.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(50, 50, 50)
        End If
    End If
    On Error GoTo 0
End Sub

'----------------------------------------------------------
' 保存（Ctrl+S／［保存］ボタン）：閲覧者モードにしてから保存し、編集者なら元に戻す
'   Workbook_BeforeSave の中ではシートを隠せない（Excel が無視する）ので、
'   編集者の保存はイベントの外にあるこの手続きで行う
'----------------------------------------------------------
Public Sub SaveLocked()
    Dim errMsg As String, wasEditor As Boolean
    If ThisWorkbook.ReadOnly Then
        MsgBox "読み取り専用で開いているため保存できません。", vbExclamation, "保存"
        Exit Sub
    End If
    wasEditor = IsEditor()
    On Error GoTo Fin
    Application.ScreenUpdating = False
    LockWorkbook
    ClearViewShapes
    Application.EnableEvents = False
    ThisWorkbook.Save
Fin:
    If Err.Number <> 0 Then errMsg = Err.Description
    Application.EnableEvents = True
    If wasEditor Then UnlockWorkbook
    ShowNode GetCfg("cfgCurrent"), False
    ' モードの戻しや描画で「変更あり」になるが、ファイルは保存済みなので閉じるときに聞かれないようにする
    If Len(errMsg) = 0 Then ThisWorkbook.Saved = True
    Application.ScreenUpdating = True
    If Len(errMsg) > 0 Then
        MsgBox "保存できませんでした。" & vbCrLf & errMsg, vbExclamation, "保存"
    Else
        Application.StatusBar = "保存しました  (" & Format$(Now, "hh:mm:ss") & ")"
    End If
End Sub

'----------------------------------------------------------
' ［編集者モード］ボタン：合言葉で切り替える
'----------------------------------------------------------
Public Sub ToggleEditorMode()
    Dim s As String

    If IsEditor() Then
        LockWorkbook
        RefreshView
        Application.StatusBar = "閲覧者モードに戻しました"
        Exit Sub
    End If

    If ThisWorkbook.ReadOnly Then
        MsgBox "読み取り専用で開いているため、編集者モードにはできません。" & vbCrLf & _
               "他の人が閉じてから開き直してください。", vbExclamation, "編集者モード"
        Exit Sub
    End If

    s = InputBox("編集者の合言葉を入力してください。", "編集者モード")
    If StrPtr(s) = 0 Or Len(s) = 0 Then Exit Sub
    If StrComp(s, LockPassword(), vbBinaryCompare) <> 0 Then
        MsgBox "合言葉が違います。", vbExclamation, "編集者モード"
        Exit Sub
    End If

    UnlockWorkbook
    RefreshView
    Application.StatusBar = "編集者モードです。台帳シートと編集ボタンが使えます"
End Sub
