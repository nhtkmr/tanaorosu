Attribute VB_Name = "M_Register"
Option Explicit

'==========================================================
' M_Register : 新規ノード／子リンクの登録（［登録］シート）
'==========================================================

'----------------------------------------------------------
' ファイル選択
'----------------------------------------------------------
Public Sub PickImage()
    Dim f As Variant, rel As String
    f = Application.GetOpenFilename( _
            "画像ファイル (*.jpg;*.jpeg;*.png;*.bmp;*.gif),*.jpg;*.jpeg;*.png;*.bmp;*.gif", _
            1, "表示用画像を選んでください")
    If VarType(f) = vbBoolean Then Exit Sub
    rel = ImportFile(CStr(f), GetCfg("cfgImgDir"))
    ThisWorkbook.Names("regImg").RefersToRange.Value = rel
End Sub

Public Sub PickPdf()
    Dim f As Variant, rel As String
    f = Application.GetOpenFilename( _
            "原本ファイル (*.pdf;*.jpg;*.jpeg;*.png;*.dwg;*.dxf),*.pdf;*.jpg;*.jpeg;*.png;*.dwg;*.dxf", _
            1, "原本ファイルを選んでください")
    If VarType(f) = vbBoolean Then Exit Sub
    rel = ImportFile(CStr(f), GetCfg("cfgPdfDir"))
    ThisWorkbook.Names("regPdf").RefersToRange.Value = rel
End Sub

' ルート配下へコピーして相対パスを返す（既にルート配下ならコピーしない）
Private Function ImportFile(ByVal src As String, ByVal subDir As String) As String
    Dim dstDir As String, baseName As String, ext As String, dst As String, i As Long

    If Len(Trim$(src)) = 0 Then Exit Function
    If Len(Trim$(subDir)) = 0 Then subDir = "img"
    EnsureFolders

    If StrComp(Left$(src, Len(RootPath()) + 1), RootPath() & "\", vbTextCompare) = 0 Then
        ImportFile = RelPath(src)
        Exit Function
    End If

    dstDir = RootPath() & "\" & subDir
    baseName = FileBaseName(src)
    ext = FileExt(src)
    dst = dstDir & "\" & baseName & ext
    i = 1
    Do While FileExistsX(dst)
        dst = dstDir & "\" & baseName & "_" & i & ext
        i = i + 1
    Loop

    On Error GoTo EH
    FileCopy src, dst
    ImportFile = RelPath(dst)
    Exit Function
EH:
    MsgBox "ファイルをコピーできませんでした。元のパスのまま登録します。" & vbCrLf & _
           Err.Description, vbExclamation
    ImportFile = src
End Function

Private Function FileBaseName(ByVal p As String) As String
    Dim nm As String, i As Long
    nm = Mid$(p, InStrRev(p, "\") + 1)
    i = InStrRev(nm, ".")
    If i > 1 Then nm = Left$(nm, i - 1)
    FileBaseName = nm
End Function

Private Function FileExt(ByVal p As String) As String
    Dim nm As String, i As Long
    nm = Mid$(p, InStrRev(p, "\") + 1)
    i = InStrRev(nm, ".")
    If i > 1 Then FileExt = Mid$(nm, i)
End Function

'----------------------------------------------------------
' 登録
'----------------------------------------------------------
Public Sub RegisterStandalone()
    DoRegister False
End Sub

Public Sub RegisterAsChild()
    DoRegister True
End Sub

Private Sub DoRegister(ByVal asChild As Boolean)
    Dim ws As Worksheet, rg As Range, newId As String, kind As String, nm As String
    Dim parentId As String, linkId As String

    Set ws = SheetOf(SH_REG)
    kind = Trim$(CStr(ThisWorkbook.Names("regKind").RefersToRange.Value))
    nm = Trim$(CStr(ThisWorkbook.Names("regName").RefersToRange.Value))

    If Len(kind) = 0 Or KindPrefix(kind) = "ND" Then
        MsgBox "「種別」を一覧から選んでください。", vbExclamation, "登録"
        Exit Sub
    End If
    If Len(nm) = 0 Then
        MsgBox "「名称」を入力してください。", vbExclamation, "登録"
        Exit Sub
    End If

    parentId = GetCfg("cfgCurrent")
    If asChild Then
        If Not NodeExists(parentId) Then
            MsgBox "親になる項目が表示されていません。" & vbCrLf & _
                   "［ビュー］シートで親を表示してから実行してください。", vbExclamation, "登録"
            Exit Sub
        End If
        If MsgBox("「" & NodeName(parentId) & "」の子として登録します。よろしいですか？", _
                  vbOKCancel + vbQuestion, "子として登録") <> vbOK Then Exit Sub
    End If

    ' --- ノード追加 ---
    newId = NextNodeId(kind)
    Set rg = TblAddRow(NodeTable())
    rg.Cells(1, NC_ID).Value = newId
    rg.Cells(1, NC_KIND).Value = kind
    rg.Cells(1, NC_NAME).Value = nm
    rg.Cells(1, NC_IMG).Value = GetReg("regImg")
    rg.Cells(1, NC_PDF).Value = GetReg("regPdf")
    rg.Cells(1, NC_MODEL).Value = GetReg("regModel")
    rg.Cells(1, NC_MAKER).Value = GetReg("regMaker")
    rg.Cells(1, NC_SHELF).Value = GetReg("regShelf")
    SetNum rg.Cells(1, NC_STOCK), GetReg("regStock")
    SetNum rg.Cells(1, NC_PRICE), GetReg("regPrice")
    rg.Cells(1, NC_VENDOR).Value = GetReg("regVendor")
    rg.Cells(1, NC_NOTE).Value = GetReg("regNote")
    rg.Cells(1, NC_UPD).Value = Now

    ' --- リンク追加 ---
    If asChild Then linkId = AddChildLink(parentId, newId)

    ClearRegister

    If asChild Then
        ShowNode parentId, False
        PlaceHotspotForLink linkId
    Else
        MsgBox "登録しました。" & vbCrLf & "ノードID : " & newId, vbInformation, "登録"
        ShowNode newId
    End If
End Sub

'----------------------------------------------------------
' 親子リンク
'----------------------------------------------------------
' 親子リンクを 1 行足してリンクIDを返す（ポインター座標は未設定のまま）
' 同じ子を別の親にも足せる＝1 つの項目を複数の親から参照できる
Public Function AddChildLink(ByVal parentId As String, ByVal childId As String) As String
    Dim lr As Range, linkId As String, n As Long

    If Not NodeExists(parentId) Then Exit Function
    If Not NodeExists(childId) Then Exit Function

    n = ChildLinkRows(parentId).Count
    linkId = NextLinkId()

    Set lr = TblAddRow(LinkTable())
    lr.Cells(1, LC_ID).Value = linkId
    lr.Cells(1, LC_PARENT).Value = parentId
    lr.Cells(1, LC_CHILD).Value = childId
    lr.Cells(1, LC_LABEL).Value = CStr(n + 1)
    lr.Cells(1, LC_SHAPE).Value = "四角"
    lr.Cells(1, LC_QTY).Value = 1

    AddChildLink = linkId
End Function

' 既存の項目を、いま表示している項目の子として追加する
' 既に子になっている項目も追加できる（同じ部品が図面の複数箇所にあるとき、
' 場所ごとにポインターを置ける）。リンクは 1 行増え、ポインターも 1 つ増える。
Public Sub LinkExistingChild()
    Dim parentId As String, childId As String, linkId As String
    Dim q As String, msg As String, sel As String, s As String
    Dim lo As ListObject, wn As Worksheet, r As Long
    Dim ids As Collection, over As Long, i As Long, n As Long, cnt As Long

    parentId = GetCfg("cfgCurrent")
    If Not NodeExists(parentId) Then
        MsgBox "親になる項目が表示されていません。" & vbCrLf & _
               "［ビュー］シートで親を表示してから実行してください。", vbExclamation, "既存を子に追加"
        Exit Sub
    End If

    q = Trim$(InputBox( _
        "「" & NodeName(parentId) & "」の子にする項目を探します。" & vbCrLf & _
        "名称・ノードID・型式・棚番 の一部を入力してください。", "既存を子に追加"))
    If Len(q) = 0 Then Exit Sub

    Set lo = NodeTable(): Set wn = lo.Parent
    Set ids = New Collection
    For r = TblFirstRow(lo) To TblLastRow(lo)
        s = Trim$(CStr(wn.Cells(r, NC_ID).Value))
        If Len(s) > 0 And StrComp(s, parentId, vbTextCompare) <> 0 Then
            If RowMatches(wn, r, q) Then
                If ids.Count < 20 Then
                    ids.Add s
                Else
                    over = over + 1
                End If
            End If
        End If
    Next r

    If ids.Count = 0 Then
        MsgBox "「" & q & "」に一致する項目が見つかりませんでした。", vbInformation, "既存を子に追加"
        Exit Sub
    End If

    If ids.Count = 1 Then
        childId = CStr(ids(1))
    Else
        msg = "子として追加する項目を番号で選んでください。" & vbCrLf & vbCrLf
        For i = 1 To ids.Count
            msg = msg & i & " : " & CandidateLabel(CStr(ids(i)))
            cnt = ChildLinkCount(parentId, CStr(ids(i)))
            If cnt > 0 Then msg = msg & "  ※既に子（" & cnt & " 箇所）"
            msg = msg & vbCrLf
        Next i
        If over > 0 Then msg = msg & vbCrLf & "（他 " & over & " 件。絞り込むと出ます）"
        sel = InputBox(msg, "既存を子に追加", "1")
        If Len(Trim$(sel)) = 0 Then Exit Sub
        n = Val(sel)
        If n < 1 Or n > ids.Count Then Exit Sub
        childId = CStr(ids(n))
    End If

    cnt = ChildLinkCount(parentId, childId)
    msg = "「" & NodeName(parentId) & "」の子として" & vbCrLf & _
          "「" & NodeName(childId) & "」（" & childId & "）を追加します。" & vbCrLf & vbCrLf
    If cnt > 0 Then
        msg = msg & "※ この項目は既に " & cnt & " 箇所にあります。" & vbCrLf & _
                    "   " & cnt + 1 & " 箇所目のポインターを置きます（項目は増えません）。"
    Else
        msg = msg & "※ 項目は増えません。参照が 1 つ増えるだけです。"
    End If
    If MsgBox(msg, vbOKCancel + vbQuestion, "既存を子に追加") <> vbOK Then Exit Sub

    linkId = AddChildLink(parentId, childId)
    If Len(linkId) = 0 Then Exit Sub

    ShowNode parentId, False
    PlaceHotspotForLink linkId
End Sub

' 親の下に、その子へのリンクが何本あるか（同じ子を複数箇所に置ける）
Private Function ChildLinkCount(ByVal parentId As String, ByVal childId As String) As Long
    Dim wl As Worksheet, v As Variant, n As Long
    Set wl = SheetOf(SH_LINK)
    For Each v In ChildLinkRows(parentId)
        If StrComp(Trim$(CStr(wl.Cells(CLng(v), LC_CHILD).Value)), childId, vbTextCompare) = 0 Then n = n + 1
    Next v
    ChildLinkCount = n
End Function

Private Function CandidateLabel(ByVal nodeId As String) As String
    Dim s As String
    s = NodeName(nodeId) & "  (" & nodeId & ")"
    If Len(NodeVal(nodeId, NC_MODEL)) > 0 Then s = s & "  " & NodeVal(nodeId, NC_MODEL)
    If Len(NodeVal(nodeId, NC_SHELF)) > 0 Then s = s & "  棚 " & NodeVal(nodeId, NC_SHELF)
    CandidateLabel = s
End Function

Private Function GetReg(ByVal nm As String) As String
    On Error Resume Next
    GetReg = Trim$(CStr(ThisWorkbook.Names(nm).RefersToRange.Value))
    On Error GoTo 0
End Function

Private Sub SetNum(ByVal c As Range, ByVal s As String)
    If Len(s) = 0 Then Exit Sub
    If IsNumeric(s) Then c.Value = CDbl(s) Else c.Value = s
End Sub

Public Sub ClearRegister()
    Dim ws As Worksheet
    Set ws = SheetOf(SH_REG)
    Application.EnableEvents = False
    ws.Range("B4:B13").ClearContents
    Application.EnableEvents = True
End Sub

'----------------------------------------------------------
' 画面移動
'----------------------------------------------------------
Public Sub GoRegisterFromView()
    Dim ws As Worksheet, cur As String
    Set ws = SheetOf(SH_REG)
    cur = GetCfg("cfgCurrent")
    ws.Activate
    ThisWorkbook.Names("regParent").RefersToRange.Value = _
        IIf(NodeExists(cur), NodeName(cur) & "  (" & cur & ")", "（未選択）")
    ws.Range("B3").Select
End Sub

Public Sub GoViewSheet()
    SheetOf(SH_VIEW).Activate
End Sub
