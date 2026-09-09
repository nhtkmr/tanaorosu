Attribute VB_Name = "M_View"
Option Explicit

'==========================================================
' M_View : ビュー画面の描画とナビゲーション
'==========================================================

Public Const IMG_SHAPE  As String = "NODE_IMAGE"
Public Const PH_SHAPE   As String = "NODE_PLACEHOLDER"
Public Const FRAME_ADDR As String = "B5:M28"

' 右パネルの位置
Public Const LBL_COL = 15      ' O列 : 見出し
Public Const VAL_COL = 16      ' P列 : 値（P:Q 結合）
Public Const CH_HDR_ROW = 21   ' 子一覧の見出し行
Public Const CH_TOP_ROW = 22   ' 子一覧の先頭行
Public Const CH_MAX = 40       ' 子一覧の最大表示件数
Public Const HID_NODE_COL = 19 ' S列 : 子のノードID（非表示）
Public Const HID_LINK_COL = 20 ' T列 : リンクID（非表示）

' 情報欄の行
Public Const IR_ID = 6, IR_KIND = 7, IR_NAME = 8, IR_MODEL = 9, IR_MAKER = 10
Public Const IR_SHELF = 11, IR_STOCK = 12, IR_PRICE = 13, IR_VENDOR = 14
Public Const IR_NOTE = 15, IR_IMG = 16, IR_PDF = 17, IR_USED = 18

'----------------------------------------------------------
' メイン：ノードを表示する
'----------------------------------------------------------
' addHistory : True なら現在の項目を履歴に積んでから移動する
'   ※ 引数名を PushHistory と同じにすると VBA が手続き名を隠してしまうので注意
Public Sub ShowNode(ByVal nodeId As String, Optional ByVal addHistory As Boolean = True)
    Dim ws As Worksheet, prev As String, pic As Shape

    nodeId = Trim$(nodeId)
    If Len(nodeId) = 0 Then nodeId = GetCfg("cfgHome")
    If Not NodeExists(nodeId) Then
        MsgBox "ノード「" & nodeId & "」が見つかりません。" & vbCrLf & _
               "［ノード］シートを確認してください。", vbExclamation, "表示できません"
        Exit Sub
    End If

    ' 編集中なら、移動する前にポインター位置を保存する
    If GetCfg("cfgEdit") = "1" Then SaveHotspots

    Set ws = SheetOf(SH_VIEW)

    On Error GoTo Fin
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    prev = GetCfg("cfgCurrent")
    If addHistory And Len(prev) > 0 And StrComp(prev, nodeId, vbTextCompare) <> 0 Then
        PushHistory prev, nodeId
    End If
    SetCfg "cfgCurrent", nodeId

    On Error Resume Next
    ws.Activate
    On Error GoTo Fin

    ClearViewShapes
    FillInfo ws, nodeId
    FillChildList ws, nodeId
    DrawBreadcrumb ws

    Set pic = DrawNodeImage(ws, nodeId)
    If Not pic Is Nothing Then DrawHotspots ws, nodeId, pic
    UpdateModeIndicator

Fin:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        MsgBox "表示中にエラーが発生しました。" & vbCrLf & Err.Description, vbExclamation
        Err.Clear
    End If
End Sub

' ボタン用（引数なし）
Public Sub GoHome()
    SetCfg "cfgHistory", ""
    ShowNode GetCfg("cfgHome"), False
End Sub

Public Sub GoBack()
    Dim h As String, arr As Variant, i As Long, out As String, tgt As String
    h = GetCfg("cfgHistory")
    If Len(h) = 0 Then GoHome: Exit Sub
    arr = Split(h, ",")
    tgt = CStr(arr(UBound(arr)))
    For i = 0 To UBound(arr) - 1
        If Len(out) > 0 Then out = out & ","
        out = out & arr(i)
    Next i
    SetCfg "cfgHistory", out
    ShowNode tgt, False
End Sub

Public Sub RefreshView()
    ShowNode GetCfg("cfgCurrent"), False
End Sub

Public Sub GoPartsSheet()
    SheetOf(SH_PARTS).Activate
End Sub

'----------------------------------------------------------
' 履歴
'----------------------------------------------------------
Private Sub PushHistory(ByVal prevId As String, ByVal newId As String)
    Dim h As String, arr As Variant, i As Long, idx As Long, out As String
    h = GetCfg("cfgHistory")
    If Len(h) > 0 Then h = h & ","
    h = h & prevId

    arr = Split(h, ",")
    idx = -1
    For i = 0 To UBound(arr)
        If StrComp(CStr(arr(i)), newId, vbTextCompare) = 0 Then idx = i: Exit For
    Next i

    If idx >= 0 Then
        ' 既に通った所へ戻った → そこまで切り詰める（履歴の無限伸長を防ぐ）
        out = ""
        For i = 0 To idx - 1
            If Len(out) > 0 Then out = out & ","
            out = out & arr(i)
        Next i
    Else
        out = h
    End If
    SetCfg "cfgHistory", out
End Sub

Private Sub DrawBreadcrumb(ByVal ws As Worksheet)
    Dim h As String, arr As Variant, i As Long, s As String
    h = GetCfg("cfgHistory")
    If Len(h) > 0 Then
        arr = Split(h, ",")
        For i = 0 To UBound(arr)
            If Len(Trim$(CStr(arr(i)))) > 0 Then s = s & NodeName(CStr(arr(i))) & "  >  "
        Next i
    End If
    s = s & NodeName(GetCfg("cfgCurrent"))
    ws.Range("B1").Value = s
End Sub

'----------------------------------------------------------
' 情報欄
'----------------------------------------------------------
Private Sub FillInfo(ByVal ws As Worksheet, ByVal nodeId As String)
    Dim r As Long, kind As String, used As String
    r = FindNodeRow(nodeId)
    kind = NodeVal(nodeId, NC_KIND)

    ws.Cells(IR_ID, VAL_COL).Value = nodeId
    ws.Cells(IR_KIND, VAL_COL).Value = kind
    ws.Cells(IR_KIND, VAL_COL).Font.Color = KindColor(kind)
    ws.Cells(IR_NAME, VAL_COL).Value = NodeVal(nodeId, NC_NAME)
    ws.Cells(IR_MODEL, VAL_COL).Value = NodeVal(nodeId, NC_MODEL)
    ws.Cells(IR_MAKER, VAL_COL).Value = NodeVal(nodeId, NC_MAKER)
    ws.Cells(IR_SHELF, VAL_COL).Value = NodeVal(nodeId, NC_SHELF)
    ws.Cells(IR_STOCK, VAL_COL).Value = NodeVal(nodeId, NC_STOCK)
    ws.Cells(IR_PRICE, VAL_COL).Value = NodeVal(nodeId, NC_PRICE)
    ws.Cells(IR_VENDOR, VAL_COL).Value = NodeVal(nodeId, NC_VENDOR)
    ws.Cells(IR_NOTE, VAL_COL).Value = NodeVal(nodeId, NC_NOTE)
    ws.Cells(IR_IMG, VAL_COL).Value = NodeVal(nodeId, NC_IMG)

    If Len(NodeVal(nodeId, NC_PDF)) > 0 Then
        ws.Cells(IR_PDF, VAL_COL).Value = NodeVal(nodeId, NC_PDF)
    Else
        ws.Cells(IR_PDF, VAL_COL).Value = "（なし）"
    End If

    ' 使用先（親）はクリックで移動できるので、あるときだけリンク風に見せる
    used = UsedIn(nodeId)
    With ws.Cells(IR_USED, VAL_COL)
        If Len(used) = 0 Then
            .Value = "（なし）"
            .Font.Color = RGB(130, 130, 130)
            .Font.Underline = xlUnderlineStyleNone
        Else
            .Value = used
            .Font.Color = RGB(0, 90, 200)
            .Font.Underline = xlUnderlineStyleSingle
        End If
    End With
End Sub

'----------------------------------------------------------
' 使用先（親）へ移動する。複数あるときは番号で選ぶ
'----------------------------------------------------------
Public Sub GoToParent()
    Dim cur As String, wl As Worksheet, v As Variant
    Dim ids As Collection, msg As String, sel As String, i As Long, n As Long

    cur = GetCfg("cfgCurrent")
    Set wl = SheetOf(SH_LINK)
    Set ids = New Collection
    For Each v In ParentLinkRows(cur)
        ids.Add Trim$(CStr(wl.Cells(CLng(v), LC_PARENT).Value))
    Next v

    If ids.Count = 0 Then
        Application.StatusBar = "「" & NodeName(cur) & "」に使用先（親）はありません"
        Exit Sub
    End If
    If ids.Count = 1 Then
        ShowNode CStr(ids(1))
        Exit Sub
    End If

    msg = "「" & NodeName(cur) & "」の使用先を番号で選んでください。" & vbCrLf & vbCrLf
    For i = 1 To ids.Count
        msg = msg & i & " : " & NodeName(CStr(ids(i))) & "  (" & ids(i) & ")" & vbCrLf
    Next i
    sel = InputBox(msg, "使用先へ移動", "1")
    If Len(Trim$(sel)) = 0 Then Exit Sub
    n = Val(sel)
    If n < 1 Or n > ids.Count Then Exit Sub
    ShowNode CStr(ids(n))
End Sub

'----------------------------------------------------------
' 子一覧
'----------------------------------------------------------
Private Sub FillChildList(ByVal ws As Worksheet, ByVal nodeId As String)
    Dim rws As Collection, v As Variant, i As Long, r As Long, wl As Worksheet
    Dim childId As String, kind As String, lbl As String, memo As String, shelf As String

    Set wl = SheetOf(SH_LINK)
    ws.Range(ws.Cells(CH_TOP_ROW, LBL_COL), ws.Cells(CH_TOP_ROW + CH_MAX + 5, HID_LINK_COL)).ClearContents
    ws.Range(ws.Cells(CH_TOP_ROW, LBL_COL), ws.Cells(CH_TOP_ROW + CH_MAX + 5, LBL_COL)).Font.Color = RGB(0, 0, 0)

    Set rws = ChildLinkRows(nodeId)
    ws.Cells(CH_HDR_ROW, LBL_COL).Value = "No"
    ws.Cells(CH_HDR_ROW, VAL_COL).Value = "名称（クリックで移動）"

    If rws.Count = 0 Then
        ws.Cells(CH_TOP_ROW, VAL_COL).Value = "（子はありません）"
        Exit Sub
    End If

    i = 0
    For Each v In rws
        i = i + 1
        If i > CH_MAX Then Exit For
        r = CLng(v)
        childId = Trim$(CStr(wl.Cells(r, LC_CHILD).Value))
        kind = NodeVal(childId, NC_KIND)
        shelf = NodeVal(childId, NC_SHELF)

        lbl = Trim$(CStr(wl.Cells(r, LC_LABEL).Value))
        If Len(lbl) = 0 Then lbl = CStr(i)

        memo = kind
        If Len(shelf) > 0 Then memo = memo & " / 棚番 " & shelf
        If Len(Trim$(CStr(wl.Cells(r, LC_X).Value))) = 0 Then memo = memo & " ※位置未設定"

        ws.Cells(CH_TOP_ROW + i - 1, LBL_COL).Value = lbl
        ws.Cells(CH_TOP_ROW + i - 1, LBL_COL).Font.Color = KindColor(kind)
        ws.Cells(CH_TOP_ROW + i - 1, VAL_COL).Value = NodeName(childId)
        ws.Cells(CH_TOP_ROW + i - 1, VAL_COL + 1).Value = memo
        ws.Cells(CH_TOP_ROW + i - 1, HID_NODE_COL).Value = childId
        ws.Cells(CH_TOP_ROW + i - 1, HID_LINK_COL).Value = Trim$(CStr(wl.Cells(r, LC_ID).Value))
    Next v
End Sub

'----------------------------------------------------------
' 画像
'----------------------------------------------------------
Public Function DrawNodeImage(ByVal ws As Worksheet, ByVal nodeId As String) As Shape
    Dim p As String, fr As Range, pic As Shape, sc As Double

    Set fr = ws.Range(FRAME_ADDR)
    p = AbsPath(NodeVal(nodeId, NC_IMG))

    If Len(p) = 0 Then
        ShowPlaceholder ws, fr, "表示用画像が未登録です" & vbCrLf & _
            "［登録］シート、または［ノード］シートの「表示用画像」に JPG を設定してください。"
        Exit Function
    End If
    If Not FileExistsX(p) Then
        ShowPlaceholder ws, fr, "画像ファイルが見つかりません" & vbCrLf & p
        Exit Function
    End If

    On Error GoTo EH
    Set pic = ws.Shapes.AddPicture(p, msoTrue, msoFalse, fr.Left, fr.Top, -1, -1)
    pic.Name = IMG_SHAPE
    pic.Placement = xlFreeFloating
    pic.LockAspectRatio = msoTrue

    sc = fr.Width / pic.Width
    If (fr.Height / pic.Height) < sc Then sc = fr.Height / pic.Height
    pic.Width = pic.Width * sc
    pic.Left = fr.Left + (fr.Width - pic.Width) / 2
    pic.Top = fr.Top + (fr.Height - pic.Height) / 2

    Set DrawNodeImage = pic
    Exit Function
EH:
    ShowPlaceholder ws, fr, "画像を読み込めませんでした" & vbCrLf & p
End Function

Private Sub ShowPlaceholder(ByVal ws As Worksheet, ByVal fr As Range, ByVal msg As String)
    Dim sh As Shape
    Set sh = ws.Shapes.AddShape(msoShapeRectangle, fr.Left, fr.Top, fr.Width, fr.Height)
    sh.Name = PH_SHAPE
    sh.Placement = xlFreeFloating
    sh.Fill.ForeColor.RGB = RGB(245, 245, 245)
    sh.Line.ForeColor.RGB = RGB(190, 190, 190)
    sh.Line.DashStyle = msoLineDash
    With sh.TextFrame2.TextRange
        .Text = msg
        .Font.Size = 12
        .Font.Fill.ForeColor.RGB = RGB(120, 120, 120)
    End With
    sh.TextFrame2.WordWrap = msoTrue
End Sub

'----------------------------------------------------------
' 原本ファイル（PDF 等）を既定のアプリで開く
'----------------------------------------------------------
Public Sub OpenOriginal()
    Dim id As String, p As String
    id = GetCfg("cfgCurrent")
    p = AbsPath(NodeVal(id, NC_PDF))
    If Len(p) = 0 Then p = AbsPath(NodeVal(id, NC_IMG))

    If Len(p) = 0 Then
        MsgBox "この項目には原本ファイルが登録されていません。", vbInformation, "原本を開く"
        Exit Sub
    End If
    If Not FileExistsX(p) Then
        MsgBox "ファイルが見つかりません。" & vbCrLf & p, vbExclamation, "原本を開く"
        Exit Sub
    End If

    On Error Resume Next
    ShellExecuteA 0, "open", p, vbNullString, vbNullString, 1
    If Err.Number <> 0 Then
        Err.Clear
        ThisWorkbook.FollowHyperlink p
    End If
    On Error GoTo 0
End Sub
