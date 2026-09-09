Option Explicit

' ［ビュー］シート：クリックで移動する
'   ・子一覧の行  → その子へ
'   ・使用先（親）→ 親へ（複数なら選択）
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim t As Range, id As String

    Set t = Target.Cells(1, 1)
    ' 単一セル、または結合セル 1 つぶんの選択だけを対象にする
    If Target.Count > t.MergeArea.Count Then Exit Sub
    If t.Column < LBL_COL Or t.Column > VAL_COL + 1 Then Exit Sub

    If t.Row = IR_USED And t.Column = VAL_COL Then
        GoToParent
        Exit Sub
    End If

    If t.Row < CH_TOP_ROW Then Exit Sub
    id = Trim$(CStr(Me.Cells(t.Row, HID_NODE_COL).Value))
    If Len(id) = 0 Then Exit Sub
    ShowNode id
End Sub
