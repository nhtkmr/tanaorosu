Attribute VB_Name = "M_Data"
Option Explicit

'==========================================================
' M_Data : 台帳（ノード／リンク）へのアクセスと共通処理
'==========================================================

' --- シート名 ---
Public Const SH_VIEW  As String = "ビュー"
Public Const SH_NODE  As String = "ノード"
Public Const SH_LINK  As String = "リンク"
Public Const SH_PARTS As String = "部品一覧"
Public Const SH_REG   As String = "登録"
Public Const SH_CFG   As String = "設定"

' --- ノード表の列位置 ---
Public Const NC_ID = 1, NC_KIND = 2, NC_NAME = 3, NC_IMG = 4, NC_PDF = 5
Public Const NC_MODEL = 6, NC_MAKER = 7, NC_SHELF = 8, NC_STOCK = 9
Public Const NC_PRICE = 10, NC_VENDOR = 11, NC_NOTE = 12, NC_UPD = 13
Public Const NC_LAST = 13

' --- リンク表の列位置 ---
Public Const LC_ID = 1, LC_PARENT = 2, LC_CHILD = 3, LC_LABEL = 4
Public Const LC_X = 5, LC_Y = 6, LC_W = 7, LC_H = 8
Public Const LC_SHAPE = 9, LC_QTY = 10, LC_NOTE = 11
Public Const LC_LAST = 11

#If VBA7 Then
    Public Declare PtrSafe Function ShellExecuteA Lib "shell32.dll" ( _
        ByVal hwnd As LongPtr, ByVal lpOperation As String, ByVal lpFile As String, _
        ByVal lpParameters As String, ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) As LongPtr
#Else
    Public Declare Function ShellExecuteA Lib "shell32.dll" ( _
        ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
        ByVal lpParameters As String, ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) As Long
#End If

'----------------------------------------------------------
' シート／テーブル
'----------------------------------------------------------
Public Function SheetOf(ByVal nm As String) As Worksheet
    Set SheetOf = ThisWorkbook.Worksheets(nm)
End Function

Public Function NodeTable() As ListObject
    Set NodeTable = SheetOf(SH_NODE).ListObjects("tblNode")
End Function

Public Function LinkTable() As ListObject
    Set LinkTable = SheetOf(SH_LINK).ListObjects("tblLink")
End Function

Public Function TblFirstRow(ByVal lo As ListObject) As Long
    TblFirstRow = lo.HeaderRowRange.Row + 1
End Function

Public Function TblLastRow(ByVal lo As ListObject) As Long
    If lo.DataBodyRange Is Nothing Then
        TblLastRow = lo.HeaderRowRange.Row
    Else
        TblLastRow = lo.DataBodyRange.Row + lo.DataBodyRange.Rows.Count - 1
    End If
End Function

' 末尾が空行ならそれを使い、無ければ 1 行追加する
Public Function TblAddRow(ByVal lo As ListObject) As Range
    Dim ws As Worksheet, r As Long, rg As Range
    Set ws = lo.Parent
    r = TblLastRow(lo)
    If r >= TblFirstRow(lo) Then
        Set rg = ws.Cells(r, lo.Range.Column).Resize(1, lo.ListColumns.Count)
        If Application.WorksheetFunction.CountA(rg) = 0 Then
            Set TblAddRow = rg
            Exit Function
        End If
    End If
    Set TblAddRow = lo.ListRows.Add.Range
End Function

'----------------------------------------------------------
' ノード
'----------------------------------------------------------
Public Function FindNodeRow(ByVal nodeId As String) As Long
    Dim lo As ListObject, ws As Worksheet, r As Long
    FindNodeRow = 0
    If Len(Trim$(nodeId)) = 0 Then Exit Function
    Set lo = NodeTable(): Set ws = lo.Parent
    For r = TblFirstRow(lo) To TblLastRow(lo)
        If StrComp(Trim$(CStr(ws.Cells(r, NC_ID).Value)), Trim$(nodeId), vbTextCompare) = 0 Then
            FindNodeRow = r
            Exit Function
        End If
    Next r
End Function

Public Function NodeExists(ByVal nodeId As String) As Boolean
    NodeExists = (FindNodeRow(nodeId) > 0)
End Function

Public Function NodeVal(ByVal nodeId As String, ByVal col As Long) As String
    Dim r As Long
    r = FindNodeRow(nodeId)
    If r = 0 Then Exit Function
    NodeVal = Trim$(CStr(SheetOf(SH_NODE).Cells(r, col).Value))
End Function

Public Function NodeName(ByVal nodeId As String) As String
    NodeName = NodeVal(nodeId, NC_NAME)
    If Len(NodeName) = 0 Then NodeName = nodeId
End Function

'----------------------------------------------------------
' リンク
'----------------------------------------------------------
Public Function FindLinkRow(ByVal linkId As String) As Long
    Dim lo As ListObject, ws As Worksheet, r As Long
    FindLinkRow = 0
    If Len(Trim$(linkId)) = 0 Then Exit Function
    Set lo = LinkTable(): Set ws = lo.Parent
    For r = TblFirstRow(lo) To TblLastRow(lo)
        If StrComp(Trim$(CStr(ws.Cells(r, LC_ID).Value)), Trim$(linkId), vbTextCompare) = 0 Then
            FindLinkRow = r
            Exit Function
        End If
    Next r
End Function

' 親ノードにぶら下がる子リンクの行番号（Collection of Long）
Public Function ChildLinkRows(ByVal parentId As String) As Collection
    Dim lo As ListObject, ws As Worksheet, r As Long
    Dim c As Collection
    Set c = New Collection
    Set lo = LinkTable(): Set ws = lo.Parent
    If Len(Trim$(parentId)) > 0 Then
        For r = TblFirstRow(lo) To TblLastRow(lo)
            If StrComp(Trim$(CStr(ws.Cells(r, LC_PARENT).Value)), Trim$(parentId), vbTextCompare) = 0 Then
                If Len(Trim$(CStr(ws.Cells(r, LC_CHILD).Value))) > 0 Then c.Add r
            End If
        Next r
    End If
    Set ChildLinkRows = c
End Function

' 子ノードを使っている親リンクの行番号
Public Function ParentLinkRows(ByVal childId As String) As Collection
    Dim lo As ListObject, ws As Worksheet, r As Long
    Dim c As Collection
    Set c = New Collection
    Set lo = LinkTable(): Set ws = lo.Parent
    If Len(Trim$(childId)) > 0 Then
        For r = TblFirstRow(lo) To TblLastRow(lo)
            If StrComp(Trim$(CStr(ws.Cells(r, LC_CHILD).Value)), Trim$(childId), vbTextCompare) = 0 Then
                c.Add r
            End If
        Next r
    End If
    Set ParentLinkRows = c
End Function

' この項目が使われている親の名前（読点区切り）
Public Function UsedIn(ByVal childId As String) As String
    Dim rws As Collection, v As Variant, s As String, ws As Worksheet
    Set ws = SheetOf(SH_LINK)
    Set rws = ParentLinkRows(childId)
    For Each v In rws
        If Len(s) > 0 Then s = s & "、"
        s = s & NodeName(CStr(ws.Cells(CLng(v), LC_PARENT).Value))
    Next v
    UsedIn = s
End Function

'----------------------------------------------------------
' ID 採番
'----------------------------------------------------------
Public Function KindPrefix(ByVal kind As String) As String
    Select Case Trim$(kind)
        Case "マップ":     KindPrefix = "MAP"
        Case "設備":       KindPrefix = "EQ"
        Case "詳細":       KindPrefix = "DT"
        Case "組図":       KindPrefix = "AS"
        Case "部品図面":   KindPrefix = "DW"
        Case "購入部品":   KindPrefix = "PT"
        Case Else:         KindPrefix = "ND"
    End Select
End Function

Public Function NextNodeId(ByVal kind As String) As String
    Dim lo As ListObject, ws As Worksheet, r As Long
    Dim pfx As String, mx As Long, s As String, n As Long
    pfx = KindPrefix(kind) & "-"
    Set lo = NodeTable(): Set ws = lo.Parent
    For r = TblFirstRow(lo) To TblLastRow(lo)
        s = Trim$(CStr(ws.Cells(r, NC_ID).Value))
        If UCase$(Left$(s, Len(pfx))) = UCase$(pfx) Then
            n = Val(Mid$(s, Len(pfx) + 1))
            If n > mx Then mx = n
        End If
    Next r
    NextNodeId = pfx & Format$(mx + 1, "000")
End Function

Public Function NextLinkId() As String
    Dim lo As ListObject, ws As Worksheet, r As Long, mx As Long, s As String, n As Long
    Set lo = LinkTable(): Set ws = lo.Parent
    For r = TblFirstRow(lo) To TblLastRow(lo)
        s = Trim$(CStr(ws.Cells(r, LC_ID).Value))
        If UCase$(Left$(s, 1)) = "L" Then
            n = Val(Mid$(s, 2))
            If n > mx Then mx = n
        End If
    Next r
    NextLinkId = "L" & Format$(mx + 1, "0000")
End Function

'----------------------------------------------------------
' 設定（名前定義で参照）
'----------------------------------------------------------
Public Function GetCfg(ByVal nm As String) As String
    On Error Resume Next
    GetCfg = Trim$(CStr(ThisWorkbook.Names(nm).RefersToRange.Value))
    On Error GoTo 0
End Function

Public Sub SetCfg(ByVal nm As String, ByVal v As String)
    On Error Resume Next
    ThisWorkbook.Names(nm).RefersToRange.Value = v
    On Error GoTo 0
End Sub

'----------------------------------------------------------
' パス
'----------------------------------------------------------
Public Function RootPath() As String
    Dim p As String
    p = GetCfg("cfgRoot")
    If Len(p) = 0 Then p = ThisWorkbook.Path
    If Right$(p, 1) = "\" Then p = Left$(p, Len(p) - 1)
    RootPath = p
End Function

Public Function IsAbsPath(ByVal p As String) As Boolean
    IsAbsPath = (InStr(p, ":") > 0) Or (Left$(p, 2) = "\\")
End Function

Public Function AbsPath(ByVal rel As String) As String
    rel = Trim$(rel)
    If Len(rel) = 0 Then Exit Function
    If IsAbsPath(rel) Then
        AbsPath = rel
    Else
        AbsPath = RootPath() & "\" & rel
    End If
End Function

Public Function RelPath(ByVal fullPath As String) As String
    Dim rt As String
    fullPath = Trim$(fullPath)
    rt = RootPath() & "\"
    If Len(fullPath) > Len(rt) Then
        If StrComp(Left$(fullPath, Len(rt)), rt, vbTextCompare) = 0 Then
            RelPath = Mid$(fullPath, Len(rt) + 1)
            Exit Function
        End If
    End If
    RelPath = fullPath
End Function

Public Function FileExistsX(ByVal p As String) As Boolean
    On Error Resume Next
    If Len(Trim$(p)) = 0 Then Exit Function
    FileExistsX = (Len(Dir$(p, vbNormal)) > 0)
    On Error GoTo 0
End Function

Public Sub EnsureFolders()
    On Error Resume Next
    Dim d As String
    d = RootPath() & "\" & GetCfg("cfgImgDir")
    If Len(Dir$(d, vbDirectory)) = 0 Then MkDir d
    d = RootPath() & "\" & GetCfg("cfgPdfDir")
    If Len(Dir$(d, vbDirectory)) = 0 Then MkDir d
    On Error GoTo 0
End Sub

'----------------------------------------------------------
' 種別ごとの色
'----------------------------------------------------------
Public Function KindColor(ByVal kind As String) As Long
    Select Case Trim$(kind)
        Case "マップ":   KindColor = RGB(90, 90, 90)
        Case "設備":     KindColor = RGB(0, 112, 192)
        Case "詳細":     KindColor = RGB(0, 150, 80)
        Case "組図":     KindColor = RGB(230, 120, 0)
        Case "部品図面": KindColor = RGB(120, 80, 200)
        Case "購入部品": KindColor = RGB(200, 30, 40)
        Case Else:       KindColor = RGB(90, 90, 90)
    End Select
End Function

' 種別色を暗くした色（薄い塗りの上に置く文字用）
Public Function Darken(ByVal clr As Long, ByVal f As Double) As Long
    Darken = RGB(Int((clr And &HFF) * f), _
                 Int(((clr \ &H100&) And &HFF) * f), _
                 Int(((clr \ &H10000) And &HFF) * f))
End Function

Public Function IsPartKind(ByVal kind As String) As Boolean
    IsPartKind = (kind = "購入部品" Or kind = "部品図面")
End Function
