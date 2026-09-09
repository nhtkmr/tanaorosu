Option Explicit

' ［部品一覧］シート：表示時に作り直し、検索語の変更で絞り込む
Private Sub Worksheet_Activate()
    RebuildParts
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    If Intersect(Target, Me.Range("B3")) Is Nothing Then Exit Sub
    RebuildParts
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Target.Cells.CountLarge > 1 Then Exit Sub
    If Target.Column <> 10 Then Exit Sub
    If Target.Row < PT_TOP_ROW Then Exit Sub
    If Len(Trim$(CStr(Target.Value))) = 0 Then Exit Sub
    PartsOpenRow Target.Row
End Sub
