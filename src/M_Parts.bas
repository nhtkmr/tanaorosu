Attribute VB_Name = "M_Parts"
Option Explicit

'==========================================================
' M_Parts : 予備品一覧（棚番検索）
'   ［ノード］シートから抽出して作り直す（閲覧専用）
'==========================================================

Public Const PT_HDR_ROW = 5
Public Const PT_TOP_ROW = 6

Public Sub RebuildParts()
    Dim ws As Worksheet, wn As Worksheet, lo As ListObject
    Dim r As Long, out As Long, q As String, kind As String, nodeId As String

    Set ws = SheetOf(SH_PARTS)
    Set lo = NodeTable()
    Set wn = lo.Parent
    q = Trim$(CStr(ThisWorkbook.Names("partsQuery").RefersToRange.Value))

    On Error GoTo Fin
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ws.Range(ws.Cells(PT_TOP_ROW, 1), ws.Cells(PT_TOP_ROW + 5000, 10)).ClearContents

    out = PT_TOP_ROW
    For r = TblFirstRow(lo) To TblLastRow(lo)
        kind = Trim$(CStr(wn.Cells(r, NC_KIND).Value))
        nodeId = Trim$(CStr(wn.Cells(r, NC_ID).Value))
        If IsPartKind(kind) And Len(nodeId) > 0 Then
            If Len(q) = 0 Or RowMatches(wn, r, q) Then
                ws.Cells(out, 1).Value = nodeId
                ws.Cells(out, 2).Value = wn.Cells(r, NC_NAME).Value
                ws.Cells(out, 3).Value = wn.Cells(r, NC_MODEL).Value
                ws.Cells(out, 4).Value = wn.Cells(r, NC_MAKER).Value
                ws.Cells(out, 5).Value = wn.Cells(r, NC_SHELF).Value
                ws.Cells(out, 6).Value = wn.Cells(r, NC_STOCK).Value
                ws.Cells(out, 7).Value = wn.Cells(r, NC_PRICE).Value
                ws.Cells(out, 8).Value = wn.Cells(r, NC_VENDOR).Value
                ws.Cells(out, 9).Value = UsedIn(nodeId)
                ws.Cells(out, 10).Value = "表示"
                out = out + 1
            End If
        End If
    Next r

    ThisWorkbook.Names("partsCount").RefersToRange.Value = (out - PT_TOP_ROW) & " 件"

Fin:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then Err.Clear
End Sub

' ノード表の 1 行が検索語に部分一致するか（［既存を子に追加］からも使う）
Public Function RowMatches(ByVal wn As Worksheet, ByVal r As Long, ByVal q As String) As Boolean
    Dim s As String, cols As Variant, i As Long
    cols = Array(NC_ID, NC_NAME, NC_MODEL, NC_MAKER, NC_SHELF, NC_VENDOR, NC_NOTE)
    For i = LBound(cols) To UBound(cols)
        s = CStr(wn.Cells(r, CLng(cols(i))).Value)
        If Len(s) > 0 Then
            If InStr(1, s, q, vbTextCompare) > 0 Then RowMatches = True: Exit Function
        End If
    Next i
End Function

Public Sub PartsSearch()
    RebuildParts
End Sub

Public Sub PartsShowAll()
    Application.EnableEvents = False
    ' 検索欄は結合セルなので ClearContents ではなく空文字を入れる
    ThisWorkbook.Names("partsQuery").RefersToRange.Value = ""
    Application.EnableEvents = True
    RebuildParts
End Sub

' 一覧の「表示」から該当ノードを開く
Public Sub PartsOpenRow(ByVal rowNum As Long)
    Dim ws As Worksheet, id As String
    Set ws = SheetOf(SH_PARTS)
    If rowNum < PT_TOP_ROW Then Exit Sub
    id = Trim$(CStr(ws.Cells(rowNum, 1).Value))
    If Len(id) = 0 Then Exit Sub
    ShowNode id
End Sub
