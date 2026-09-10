Attribute VB_Name = "M_Hotspot"
Option Explicit

'==========================================================
' M_Hotspot : 画像上のポインター（ホットスポット）
'   座標は画像に対する 0〜1 の相対値でリンク表に保存する
'==========================================================

Public Const HS_PREFIX As String = "HS_"

' 直前の描画で実際に画面に出したリンクID。
' 「図形が無い＝ユーザーが消した」と判断してよいのは、ここに入っているものだけ。
Private gDrawnLinks As Collection

'----------------------------------------------------------
' 画像とポインターを消す（ボタンなどは残す）
'----------------------------------------------------------
Public Sub ClearViewShapes()
    Dim ws As Worksheet, i As Long, nm As String
    Set ws = SheetOf(SH_VIEW)
    For i = ws.Shapes.Count To 1 Step -1
        nm = ws.Shapes(i).Name
        If nm = IMG_SHAPE Or nm = PH_SHAPE Or Left$(nm, Len(HS_PREFIX)) = HS_PREFIX Then
            ws.Shapes(i).Delete
        End If
    Next i
End Sub

'----------------------------------------------------------
' 描画
'----------------------------------------------------------
Public Sub DrawHotspots(ByVal ws As Worksheet, ByVal nodeId As String, ByVal pic As Shape)
    Dim wl As Worksheet, rws As Collection, v As Variant, r As Long, i As Long
    Dim x As Double, y As Double, w As Double, h As Double
    Dim pw As Double, ph As Double
    Dim childId As String, lbl As String, linkId As String
    Dim sh As Shape

    Set wl = SheetOf(SH_LINK)
    Set rws = ChildLinkRows(nodeId)
    Set gDrawnLinks = New Collection
    pw = pic.Width: ph = pic.Height

    i = 0
    For Each v In rws
        i = i + 1
        r = CLng(v)
        If Len(Trim$(CStr(wl.Cells(r, LC_X).Value))) > 0 Then
            x = Clamp(NumOf(wl.Cells(r, LC_X), 0.45), -0.2, 1.2)
            y = Clamp(NumOf(wl.Cells(r, LC_Y), 0.45), -0.2, 1.2)
            w = Clamp(NumOf(wl.Cells(r, LC_W), 0.08), 0.005, 1#)
            h = Clamp(NumOf(wl.Cells(r, LC_H), 0.06), 0.005, 1#)

            childId = Trim$(CStr(wl.Cells(r, LC_CHILD).Value))
            linkId = Trim$(CStr(wl.Cells(r, LC_ID).Value))
            lbl = Trim$(CStr(wl.Cells(r, LC_LABEL).Value))
            If Len(lbl) = 0 Then lbl = CStr(i)

            Set sh = ws.Shapes.AddShape(ShapeTypeOf(CStr(wl.Cells(r, LC_SHAPE).Value)), _
                        pic.Left + x * pw, pic.Top + y * ph, _
                        Application.WorksheetFunction.Max(w * pw, 14), _
                        Application.WorksheetFunction.Max(h * ph, 12))

            StyleHotspot sh, HS_PREFIX & linkId, lbl, _
                         KindColor(NodeVal(childId, NC_KIND)), _
                         NodeName(childId) & " (" & childId & ")"
            gDrawnLinks.Add linkId
        End If
    Next v
End Sub

Private Function ShapeTypeOf(ByVal s As String) As Long
    Select Case Trim$(s)
        Case "丸", "円", "楕円": ShapeTypeOf = msoShapeOval
        Case Else:               ShapeTypeOf = msoShapeRoundedRectangle
    End Select
End Function

Private Sub StyleHotspot(ByVal sh As Shape, ByVal nm As String, ByVal lbl As String, _
                         ByVal clr As Long, ByVal tip As String)
    sh.Name = nm
    sh.Placement = xlFreeFloating
    sh.Fill.ForeColor.RGB = clr
    sh.Fill.Transparency = Clamp(CfgNum("cfgHsAlpha", 0.85), 0#, 0.95)
    sh.Line.ForeColor.RGB = clr
    sh.Line.Weight = 1.75

    With sh.TextFrame2
        .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
        .VerticalAnchor = msoAnchorMiddle
        .WordWrap = msoFalse
        With .TextRange
            .Text = lbl
            .Font.Size = 11
            .Font.Bold = msoTrue
            .Font.Fill.ForeColor.RGB = Darken(clr, 0.7)
            .ParagraphFormat.Alignment = msoAlignCenter
        End With
    End With

    On Error Resume Next
    sh.AlternativeText = tip
    On Error GoTo 0

    If GetCfg("cfgEdit") = "1" Then
        ' 編集モード中はクリックでマクロを走らせない（ドラッグして動かせるようにする）
        sh.Line.DashStyle = msoLineDash
        sh.OnAction = ""
    Else
        sh.Line.DashStyle = msoLineSolid
        sh.OnAction = "HotspotClicked"
    End If
End Sub

'----------------------------------------------------------
' クリック → 子へ移動
'----------------------------------------------------------
Public Sub HotspotClicked()
    Dim nm As String, linkId As String, r As Long, childId As String
    On Error Resume Next
    nm = CStr(Application.Caller)
    On Error GoTo 0
    If Left$(nm, Len(HS_PREFIX)) <> HS_PREFIX Then Exit Sub

    linkId = Mid$(nm, Len(HS_PREFIX) + 1)
    r = FindLinkRow(linkId)
    If r = 0 Then Exit Sub

    childId = Trim$(CStr(SheetOf(SH_LINK).Cells(r, LC_CHILD).Value))
    If Len(childId) = 0 Then Exit Sub
    ShowNode childId
End Sub

'----------------------------------------------------------
' 追加
'----------------------------------------------------------
Public Sub AddHotspot()
    Dim ws As Worksheet, wl As Worksheet, pic As Shape
    Dim cur As String, msg As String, sel As String
    Dim cand As Collection, v As Variant, r As Long, i As Long, n As Long

    cur = GetCfg("cfgCurrent")
    Set ws = SheetOf(SH_VIEW)

    On Error Resume Next
    Set pic = ws.Shapes(IMG_SHAPE)
    On Error GoTo 0
    If pic Is Nothing Then
        MsgBox "この項目には表示用画像がありません。" & vbCrLf & _
               "先に［ノード］シートまたは［登録］シートで JPG を設定してください。", vbExclamation, "ポインター追加"
        Exit Sub
    End If

    Set wl = SheetOf(SH_LINK)
    Set cand = New Collection
    For Each v In ChildLinkRows(cur)
        r = CLng(v)
        If Len(Trim$(CStr(wl.Cells(r, LC_X).Value))) = 0 Then cand.Add r
    Next v

    If cand.Count = 0 Then
        MsgBox "位置が未設定の子がありません。" & vbCrLf & vbCrLf & _
               "・新しい子を作る → ［子として新規登録］" & vbCrLf & _
               "・既存のポインターを動かす → ［編集モード］", vbInformation, "ポインター追加"
        Exit Sub
    End If

    If cand.Count = 1 Then
        r = CLng(cand(1))
    Else
        msg = "ポインターを置く子を番号で選んでください。" & vbCrLf & vbCrLf
        For i = 1 To cand.Count
            msg = msg & i & " : " & NodeName(CStr(wl.Cells(CLng(cand(i)), LC_CHILD).Value)) & vbCrLf
        Next i
        sel = InputBox(msg, "ポインター追加", "1")
        If Len(Trim$(sel)) = 0 Then Exit Sub
        n = Val(sel)
        If n < 1 Or n > cand.Count Then Exit Sub
        r = CLng(cand(n))
    End If

    PlaceHotspotForLink Trim$(CStr(wl.Cells(r, LC_ID).Value))
End Sub

' 指定リンクのポインターを画像中央に置き、編集モードに入る
Public Sub PlaceHotspotForLink(ByVal linkId As String)
    Dim wl As Worksheet, r As Long, parentId As String
    Dim w As Double, h As Double, ofs As Double

    r = FindLinkRow(linkId)
    If r = 0 Then Exit Sub
    Set wl = SheetOf(SH_LINK)
    parentId = Trim$(CStr(wl.Cells(r, LC_PARENT).Value))

    If Len(NodeVal(parentId, NC_IMG)) = 0 Then
        MsgBox "親「" & NodeName(parentId) & "」に表示用画像がないため、" & vbCrLf & _
               "ポインターは配置しませんでした（子一覧には表示されます）。", vbInformation, "ポインター"
        Exit Sub
    End If

    w = CfgNum("cfgHsW", 0.08)
    h = CfgNum("cfgHsH", 0.06)
    If w <= 0 Then w = 0.08
    If h <= 0 Then h = 0.06
    ' 続けて追加したときに重ならないよう少しずらす
    ofs = (CountPlacedHotspots(parentId) Mod 6) * 0.03

    wl.Cells(r, LC_X).Value = Round(0.5 - w / 2 + ofs, 4)
    wl.Cells(r, LC_Y).Value = Round(0.45 - h / 2 + ofs, 4)
    wl.Cells(r, LC_W).Value = Round(w, 4)
    wl.Cells(r, LC_H).Value = Round(h, 4)

    SetCfg "cfgEdit", "1"
    RefreshView
    MsgBox "画像の中央にポインターを追加しました。" & vbCrLf & _
           "ドラッグで位置を合わせ、［位置を保存］を押してください。", vbInformation, "ポインター"
End Sub

Private Function CountPlacedHotspots(ByVal parentId As String) As Long
    Dim wl As Worksheet, v As Variant, r As Long, n As Long
    Set wl = SheetOf(SH_LINK)
    For Each v In ChildLinkRows(parentId)
        r = CLng(v)
        If Len(Trim$(CStr(wl.Cells(r, LC_X).Value))) > 0 Then n = n + 1
    Next v
    CountPlacedHotspots = n
End Function

'----------------------------------------------------------
' 保存（図形の位置 → リンク表）
'----------------------------------------------------------
Public Sub SaveHotspots()
    Dim ws As Worksheet, wl As Worksheet, pic As Shape, sh As Shape
    Dim found As Collection, v As Variant
    Dim cur As String, linkId As String, r As Long

    Set ws = SheetOf(SH_VIEW)
    On Error Resume Next
    Set pic = ws.Shapes(IMG_SHAPE)
    On Error GoTo 0
    If pic Is Nothing Then Exit Sub
    If pic.Width <= 0 Or pic.Height <= 0 Then Exit Sub

    Set wl = SheetOf(SH_LINK)
    cur = GetCfg("cfgCurrent")
    Set found = New Collection

    For Each sh In ws.Shapes
        If Left$(sh.Name, Len(HS_PREFIX)) = HS_PREFIX Then
            linkId = Mid$(sh.Name, Len(HS_PREFIX) + 1)
            r = FindLinkRow(linkId)
            If r > 0 Then
                wl.Cells(r, LC_X).Value = Round(Clamp((sh.Left - pic.Left) / pic.Width, -0.2, 1.2), 4)
                wl.Cells(r, LC_Y).Value = Round(Clamp((sh.Top - pic.Top) / pic.Height, -0.2, 1.2), 4)
                wl.Cells(r, LC_W).Value = Round(Clamp(sh.Width / pic.Width, 0.005, 1#), 4)
                wl.Cells(r, LC_H).Value = Round(Clamp(sh.Height / pic.Height, 0.005, 1#), 4)
                found.Add linkId
            End If
        End If
    Next sh

    ' 画面に出ていたのに図形が無い＝ユーザーが消した → 「位置未設定」に戻す
    ' （親子関係そのものは残す）
    ' 直前の描画に出ていないリンクは触らない。追加した直後のポインターは
    ' まだ描かれていないので、ここで消すと画面に出てこなくなる。
    For Each v In ChildLinkRows(cur)
        r = CLng(v)
        linkId = Trim$(CStr(wl.Cells(r, LC_ID).Value))
        If WasDrawn(linkId) And Not CollHas(found, linkId) Then
            wl.Range(wl.Cells(r, LC_X), wl.Cells(r, LC_H)).ClearContents
        End If
    Next v
End Sub

Private Function WasDrawn(ByVal linkId As String) As Boolean
    If gDrawnLinks Is Nothing Then Exit Function
    WasDrawn = CollHas(gDrawnLinks, linkId)
End Function

Public Sub SaveHotspotsAndRefresh()
    SaveHotspots
    RefreshView
    Application.StatusBar = "ポインターの位置を保存しました  (" & Format$(Now, "hh:mm:ss") & ")"
End Sub

'----------------------------------------------------------
' 編集モード
'----------------------------------------------------------
Public Sub ToggleEditMode()
    If GetCfg("cfgEdit") = "1" Then
        SaveHotspots
        SetCfg "cfgEdit", "0"
        RefreshView
        Application.StatusBar = "編集モードを終了し、位置を保存しました"
    Else
        SetCfg "cfgEdit", "1"
        RefreshView
        MsgBox "編集モードにしました。" & vbCrLf & vbCrLf & _
               "・ポインターをドラッグ／サイズ変更できます" & vbCrLf & _
               "・不要なポインターは選んで Delete キー（子との関係は残ります）" & vbCrLf & _
               "・終わったら［位置を保存］または［編集モード］をもう一度押してください", _
               vbInformation, "編集モード"
    End If
End Sub

Public Sub UpdateModeIndicator()
    Dim ws As Worksheet, sh As Shape, isEdit As Boolean
    Set ws = SheetOf(SH_VIEW)
    isEdit = (GetCfg("cfgEdit") = "1")

    If ThisWorkbook.ReadOnly Then
        ' 共有フォルダで他の人が先に開いているとき。閲覧・検索は全部できる
        ws.Range("O1").Value = "読み取り専用（閲覧のみ。登録や位置の保存は残りません）"
        ws.Range("O1").Font.Color = RGB(200, 90, 0)
    ElseIf isEdit Then
        ws.Range("O1").Value = "★編集モード：ポインターをドラッグ →［位置を保存］"
        ws.Range("O1").Font.Color = RGB(200, 0, 0)
    Else
        ws.Range("O1").Value = "閲覧モード：ポインターのクリックで移動"
        ws.Range("O1").Font.Color = RGB(100, 100, 100)
    End If

    On Error Resume Next
    Set sh = ws.Shapes("BTN_EDIT")
    If Not sh Is Nothing Then
        sh.TextFrame2.TextRange.Text = IIf(isEdit, "編集モード：ON", "編集モード：OFF")
        If isEdit Then
            sh.Fill.ForeColor.RGB = RGB(200, 40, 40)
            sh.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        Else
            sh.Fill.ForeColor.RGB = RGB(242, 242, 242)
            sh.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(50, 50, 50)
        End If
    End If
    On Error GoTo 0
End Sub

'----------------------------------------------------------
' 小物
'----------------------------------------------------------
Public Function NumOf(ByVal c As Range, ByVal dflt As Double) As Double
    On Error GoTo EH
    If Len(Trim$(CStr(c.Value))) = 0 Then NumOf = dflt Else NumOf = CDbl(c.Value)
    Exit Function
EH:
    NumOf = dflt
End Function

' 名前定義の設定値を数値で読む（名前が無い旧ブックでも既定値に落ちる）
Public Function CfgNum(ByVal nm As String, ByVal dflt As Double) As Double
    On Error GoTo EH
    CfgNum = NumOf(ThisWorkbook.Names(nm).RefersToRange, dflt)
    Exit Function
EH:
    CfgNum = dflt
End Function

Public Function Clamp(ByVal v As Double, ByVal lo As Double, ByVal hi As Double) As Double
    If v < lo Then v = lo
    If v > hi Then v = hi
    Clamp = v
End Function

Public Function CollHas(ByVal c As Collection, ByVal s As String) As Boolean
    Dim v As Variant
    For Each v In c
        If StrComp(CStr(v), s, vbTextCompare) = 0 Then CollHas = True: Exit Function
    Next v
End Function
