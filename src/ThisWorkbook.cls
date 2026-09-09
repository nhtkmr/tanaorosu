Option Explicit

'==========================================================
' ThisWorkbook : 起動時／保存時の処理
'==========================================================

Private Sub Workbook_Open()
    Dim id As String
    On Error Resume Next
    Application.StatusBar = False
    EnsureFolders
    SetCfg "cfgEdit", "0"
    id = GetCfg("cfgCurrent")
    If Len(id) = 0 Then id = GetCfg("cfgHome")
    ShowNode id, False
End Sub

' 画像とポインターはブックに保存しない（ファイルを軽く保つ）
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    On Error Resume Next
    If GetCfg("cfgEdit") = "1" Then SaveHotspots
    ClearViewShapes
End Sub

Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    On Error Resume Next
    ShowNode GetCfg("cfgCurrent"), False
End Sub

' 読み取り専用（＝他の人が先に開いている／共有の閲覧用）で開いた場合、
' 画面を描くだけでブックが変更済みになってしまうので、閉じるときに保存を聞かない
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    If ThisWorkbook.ReadOnly Then ThisWorkbook.Saved = True
End Sub
