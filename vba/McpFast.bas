Attribute VB_Name = "McpFast"
' ============================================================================
' msmcp in-process engine for MicroStation V8i (VBA).
' Runs INSIDE ustation.exe, so element scans avoid cross-process COM calls.
'
' Protocol (driven by the Python MCP server):
'   1. Python writes  %LOCALAPPDATA%\msmcp\req.txt   (ASCII, key=value lines,
'      optional "---" followed by tab-separated edit commands).
'   2. Python sends key-in:  vba run [msmcp]McpFast.Main
'   3. This module writes   resp.txt  (JSON lines, non-ASCII escaped as \uXXXX)
'      and then            done.txt  (the request token) to signal completion.
' Strings coming from Python are hex-encoded UTF-16 code units (4 hex / char).
' Keep this file pure ASCII.
' ============================================================================
Option Explicit

Private gTok As String
Private gOp As String
Private gFile As String
Private gLevels As String
Private gTypes As String
Private gIds As String
Private gGeom As Boolean
Private gVmax As Long
Private gHasBox As Boolean
Private gX1 As Double, gY1 As Double, gX2 As Double, gY2 As Double
Private gCmds() As String
Private gCmdCount As Long
Private gOut As Integer
Private gOutOpen As Boolean
Private gT(0 To 7) As Double
Private gStyles As Collection
Private gWarn As String
' line style reading is expensive for custom styles (~0.5 ms/element):
' gStyleMode 0 = none, 1 = all, 2 = first 30 per level; gStyleLv = always read on these levels
Private gStyleMode As Long
Private gStyleLv As String
Private gStyleCnt(0 To 1023) As Long
Private gLvOk As Long, gLvFail As Long

' ---------------------------------------------------------------- helpers ---
Private Function BaseDir() As String
    BaseDir = Environ$("LOCALAPPDATA") & "\msmcp\"
End Function

Private Function B(ByVal v As Boolean) As String
    If v Then B = "true" Else B = "false"
End Function

Private Function N(ByVal x As Double) As String
    Dim s As String
    s = Trim$(Str$(x))
    If Left$(s, 1) = "." Then s = "0" & s
    If Left$(s, 2) = "-." Then s = "-0" & Mid$(s, 2)
    N = s
End Function

Private Function J(ByVal s As String) As String
    Dim i As Long, c As Long, r As String, ch As String
    r = """"
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        c = AscW(ch) And &HFFFF&
        If c = 34 Then
            r = r & "\"""
        ElseIf c = 92 Then
            r = r & "\\"
        ElseIf c < 32 Or c > 126 Then
            r = r & "\u" & Right$("000" & Hex$(c), 4)
        Else
            r = r & ch
        End If
    Next i
    J = r & """"
End Function

Private Function HD(ByVal h As String) As String
    Dim i As Long, r As String
    For i = 1 To Len(h) - 3 Step 4
        r = r & ChrW(CLng("&H" & Mid$(h, i, 4)))
    Next i
    HD = r
End Function

Private Function P2(ByRef p As Point3d) As String
    P2 = "[" & N(p.X) & "," & N(p.Y) & "]"
End Function

Private Sub W(ByVal s As String)
    Print #gOut, s
End Sub

Private Function RGBOf(m As ModelReference, ByVal c As Long) As Long
    On Error GoTo bad
    RGBOf = m.InternalColorToRGBColor(c)
    Exit Function
bad:
    RGBOf = -1
End Function

' Line style names are cached by number: resolving a custom style name
' (e.g. "{Tuong nha}") per element costs ~0.4 ms and dominated scan time.
Private Function StyleName(ls As LineStyle) As String
    Dim k As String, nm As String
    On Error GoTo bad
    If gStyles Is Nothing Then Set gStyles = New Collection
    k = "n" & CStr(ls.Number)
    On Error Resume Next
    nm = gStyles(k)
    If Err.Number <> 0 Then
        Err.Clear
        nm = ls.Name
        gStyles.Add nm, k
    End If
    StyleName = nm
    Exit Function
bad:
    StyleName = ""
End Function

Private Function InList(ByVal lst As String, ByVal v As Long) As Boolean
    If lst = "" Then
        InList = True
    Else
        InList = (InStr(1, lst, "," & CStr(v) & ",") > 0)
    End If
End Function

' --------------------------------------------------------------- request ---
Private Sub ReadReq()
    Dim f As Integer, ln As String, k As String, v As String, p As Long
    Dim inCmds As Boolean, a() As String
    gTok = "": gOp = "": gFile = "": gLevels = "": gTypes = "": gIds = ""
    gGeom = False: gVmax = 0: gHasBox = False: gCmdCount = 0
    gStyleMode = 0: gStyleLv = "": gLvOk = 0: gLvFail = 0
    Erase gStyleCnt
    ReDim gCmds(0 To 0)
    f = FreeFile
    Open BaseDir() & "req.txt" For Input As #f
    Do While Not EOF(f)
        Line Input #f, ln
        If inCmds Then
            If Len(ln) > 0 Then
                gCmdCount = gCmdCount + 1
                ReDim Preserve gCmds(0 To gCmdCount)
                gCmds(gCmdCount) = ln
            End If
        ElseIf ln = "---" Then
            inCmds = True
        Else
            p = InStr(1, ln, "=")
            If p > 0 Then
                k = Left$(ln, p - 1): v = Mid$(ln, p + 1)
                Select Case k
                    Case "tok": gTok = v
                    Case "op": gOp = v
                    Case "file": gFile = HD(v)
                    Case "levels": If v <> "" Then gLevels = "," & v & ","
                    Case "types": If v <> "" Then gTypes = "," & v & ","
                    Case "ids": gIds = v
                    Case "geom": gGeom = (v = "1")
                    Case "vmax": gVmax = CLng(Val(v))
                    Case "style": gStyleMode = CLng(Val(v))
                    Case "stylelv": If v <> "" Then gStyleLv = "," & v & ","
                    Case "bbox"
                        If v <> "" Then
                            a = Split(v, ",")
                            gX1 = Val(a(0)): gY1 = Val(a(1)): gX2 = Val(a(2)): gY2 = Val(a(3))
                            gHasBox = True
                        End If
                End Select
            End If
        End If
    Loop
    Close #f
End Sub

' ------------------------------------------------------------------ main ---
Public Sub Main()
    Dim f As Integer, t As Double
    t = Timer
    On Error GoTo Fail
    ReadReq
    gOut = FreeFile
    Open BaseDir() & "resp.txt" For Output As #gOut
    gOutOpen = True
    Select Case gOp
        Case "ping"
            W "{""kind"":""pong"",""file"":" & J(ActiveDesignFile.FullName) & "}"
        Case "dump"
            DoDump
        Case "apply"
            DoApply
        Case "copy"
            DoCopy
        Case Else
            W "{""kind"":""error"",""err"":" & J("unknown op " & gOp) & "}"
    End Select
    W "{""kind"":""end"",""sec"":" & N(Timer - t) & "}"
    GoTo Finish
Fail:
    If gOutOpen Then
        W "{""kind"":""error"",""err"":" & J(Err.Description & " (" & Err.Number & ")") & "}"
    End If
Finish:
    On Error Resume Next
    If gOutOpen Then Close #gOut
    gOutOpen = False
    f = FreeFile
    Open BaseDir() & "done.txt" For Output As #f
    Print #f, gTok
    Close #f
End Sub

' ------------------------------------------------------------------ dump ---
Private Sub DoDump()
    Dim df As DesignFile, m As ModelReference, ext As Boolean
    Dim lv As Level, s As String, i As Long, cnt As Long
    Dim sc As New ElementScanCriteria, ee As ElementEnumerator, el As Element
    Dim ids() As String

    Set gStyles = New Collection
    If gFile <> "" Then
        Set df = OpenDesignFileForProgram(gFile, True)
        ext = True
        Set m = df.DefaultModelReference
    Else
        Set df = ActiveDesignFile
        Set m = ActiveModelReference
    End If

    ' header: file + levels + palette
    s = ""
    For Each lv In df.Levels
        If s <> "" Then s = s & ","
        s = s & LevelJson(m, lv)
    Next lv
    W "{""kind"":""header"",""file"":" & J(df.FullName) & ",""model"":" & J(m.Name) & _
      ",""external"":" & B(ext) & ",""levels"":[" & s & "]}"
    s = ""
    For i = 0 To 255
        If i > 0 Then s = s & ","
        s = s & RGBOf(m, i)
    Next i
    W "{""kind"":""palette"",""rgb"":[" & s & "]}"

    If gIds <> "" Then
        ids = Split(gIds, ",")
        For i = LBound(ids) To UBound(ids)
            Set el = Nothing
            On Error Resume Next
            Set el = m.GetElementByID(DLongFromString(ids(i)))
            On Error GoTo 0
            If el Is Nothing Then
                W "{""kind"":""missing"",""id"":""" & ids(i) & """}"
            Else
                WriteEl m, el
                cnt = cnt + 1
            End If
        Next i
    Else
        Set ee = m.Scan(sc)
        Do While ee.MoveNext
            Set el = ee.Current
            If el.IsGraphical Then
                If Pass(el) Then
                    WriteEl m, el
                    cnt = cnt + 1
                End If
            End If
        Loop
    End If
    W "{""kind"":""count"",""n"":" & cnt & ",""prof"":[" & N(gT(1)) & "," & N(gT(2)) & "," & N(gT(3)) & "," & N(gT(4)) & "," & N(gT(5)) & "," & N(gT(6)) & "," & N(gT(7)) & "]}"
    For i = 0 To 7
        gT(i) = 0
    Next i
    If ext Then df.Close
End Sub

Private Function LevelJson(m As ModelReference, lv As Level) As String
    Dim code As Long, col As Long, wt As Long, st As String, nm As String
    On Error Resume Next
    code = -1: col = -1: wt = -1: st = "": nm = ""
    nm = lv.Name
    code = lv.Code
    col = lv.ElementColor
    wt = lv.ElementLineWeight
    st = StyleName(lv.ElementLineStyle)
    LevelJson = "{""code"":" & code & ",""name"":" & J(nm) & ",""color"":" & col & _
        ",""rgb"":" & RGBOf(m, col) & ",""weight"":" & wt & ",""style"":" & J(st) & "}"
End Function

Private Function Pass(el As Element) As Boolean
    Dim lc As Long, r As Range3d
    Pass = False
    If gTypes <> "" Then
        If Not InList(gTypes, el.Type) Then Exit Function
    End If
    If gLevels <> "" Then
        lc = -1
        On Error Resume Next
        lc = el.Level.Code
        On Error GoTo 0
        If Not InList(gLevels, lc) Then Exit Function
    End If
    If gHasBox Then
        r = el.Range
        If r.High.X < gX1 Or r.Low.X > gX2 Or r.High.Y < gY1 Or r.Low.Y > gY2 Then Exit Function
    End If
    Pass = True
End Function

Private Sub WriteEl(m As ModelReference, el As Element)
    Dim t As Long, lc As Long, col As Long, wt As Long, st As String
    Dim r As Range3d, s As String, sid As String, gg As Long
    On Error Resume Next
    t = el.Type
    sid = DLongToString(el.ID)
    Dim t0 As Double, lvo As Level
    t0 = Timer
    lc = -1
    ' background-opened files have no level table: stop asking after 200 straight failures
    If Not (gLvOk = 0 And gLvFail > 200) Then
        Set lvo = Nothing
        Set lvo = el.Level
        If lvo Is Nothing Then
            gLvFail = gLvFail + 1
        Else
            lc = lvo.Code
            gLvOk = gLvOk + 1
        End If
    End If
    gT(1) = gT(1) + Timer - t0: t0 = Timer
    col = -1: col = el.Color
    wt = -1: wt = el.LineWeight
    gT(2) = gT(2) + Timer - t0: t0 = Timer
    st = "?"
    If WantStyle(lc) Then
        st = ""
        st = StyleName(el.LineStyle)
    End If
    gT(3) = gT(3) + Timer - t0: t0 = Timer
    gg = 0: gg = el.GraphicGroup
    r = el.Range
    On Error GoTo 0
    gT(4) = gT(4) + Timer - t0: t0 = Timer
    s = "{""kind"":""el"",""id"":""" & sid & """,""t"":" & t & ",""lv"":" & lc & _
        ",""c"":" & col & ",""rgb"":" & RGBOf(m, col) & ",""w"":" & wt & ",""ls"":" & J(st) & _
        ",""gg"":" & gg & ",""r"":[" & N(r.Low.X) & "," & N(r.Low.Y) & "," & N(r.High.X) & "," & N(r.High.Y) & "]"
    gT(5) = gT(5) + Timer - t0: t0 = Timer
    s = s & Attrs(el, t)
    gT(6) = gT(6) + Timer - t0: t0 = Timer
    W s & "}"
    gT(7) = gT(7) + Timer - t0
End Sub

Private Function WantStyle(ByVal lc As Long) As Boolean
    WantStyle = False
    If gStyleLv <> "" Then
        If InStr(1, gStyleLv, "," & CStr(lc) & ",") > 0 Then
            WantStyle = True
            Exit Function
        End If
    End If
    Select Case gStyleMode
        Case 1
            WantStyle = True
        Case 2
            If lc >= 0 And lc <= 1023 Then
                If gStyleCnt(lc) < 30 Then
                    gStyleCnt(lc) = gStyleCnt(lc) + 1
                    WantStyle = True
                End If
            End If
    End Select
End Function

Private Function Verts(el As Element) As String
    Dim v() As Point3d, i As Long, s As String, k As Long
    On Error GoTo bad
    v = el.AsVertexList.GetVertices
    For i = LBound(v) To UBound(v)
        If gVmax > 0 And k >= gVmax Then Exit For
        If s <> "" Then s = s & ","
        s = s & P2(v(i))
        k = k + 1
    Next i
    Verts = s
    Exit Function
bad:
    Verts = ""
End Function

Private Function Attrs(el As Element, ByVal t As Long) As String
    Dim s As String, se As Element, ee As ElementEnumerator, vs As String, part As String
    Dim i As Long, nl As Long, txt As String, q As Point3d, q2 As Point3d
    On Error GoTo bad
    Select Case t
        Case 3, 4, 6, 11
            If gGeom Then s = ",""v"":[" & Verts(el) & "]"
        Case 12, 14
            If gGeom Then
                Set ee = el.AsComplexElement.GetSubElements
                Do While ee.MoveNext
                    Set se = ee.Current
                    part = ""
                    If se.IsVertexList Then
                        part = Verts(se)
                    ElseIf se.Type = 16 Then
                        q = se.AsArcElement.StartPoint
                        q2 = se.AsArcElement.EndPoint
                        part = P2(q) & "," & P2(q2)
                    End If
                    If part <> "" Then
                        If vs <> "" Then vs = vs & ","
                        vs = vs & part
                    End If
                Loop
                s = ",""v"":[" & vs & "]"
            End If
        Case 15
            q = el.AsEllipseElement.CenterPoint
            s = ",""o"":" & P2(q) & ",""rad"":" & N(el.AsEllipseElement.PrimaryRadius)
        Case 16
            q = el.AsArcElement.CenterPoint
            s = ",""o"":" & P2(q) & ",""rad"":" & N(el.AsArcElement.PrimaryRadius)
            If gGeom Then
                q = el.AsArcElement.StartPoint
                q2 = el.AsArcElement.EndPoint
                s = s & ",""v"":[" & P2(q) & "," & P2(q2) & "]"
            End If
        Case 17
            q = el.AsTextElement.Origin
            s = ",""tx"":" & J(el.AsTextElement.Text) & ",""o"":" & P2(q) & _
                ",""h"":" & N(el.AsTextElement.TextStyle.Height)
        Case 7
            nl = el.AsTextNodeElement.TextLinesCount
            For i = 1 To nl
                If i > 1 Then txt = txt & vbLf
                txt = txt & el.AsTextNodeElement.TextLine(i)
            Next i
            q = el.AsTextNodeElement.Origin
            s = ",""tx"":" & J(txt) & ",""o"":" & P2(q)
        Case 2
            q = el.AsCellElement.Origin
            s = ",""cell"":" & J(el.AsCellElement.Name) & ",""o"":" & P2(q)
        Case 35
            q = el.AsSharedCellElement.Origin
            s = ",""cell"":" & J(el.AsSharedCellElement.Name) & ",""o"":" & P2(q)
    End Select
    Attrs = s
    Exit Function
bad:
    Attrs = s & ",""aerr"":" & J(Err.Description)
End Function

' ------------------------------------------------------------------ copy ---
' Deep copy from a background-opened file (gFile) into the active model with
' ModelReference.CopyElement, which keeps every attribute (symbology, ByLevel,
' class, linkages) and remaps levels/styles/colors to the destination.
' Command fields: srcId, runStart, runEnd, snapped points "x,y;x,y" ("" = unchanged)
' runStart = -2 means: build a new line from the points using the copy as template.
Private Sub DoCopy()
    Dim df As DesignFile, sm As ModelReference, i As Long, f() As String, res As String
    Set df = OpenDesignFileForProgram(gFile, True)
    Set sm = df.DefaultModelReference
    For i = 1 To gCmdCount
        f = Split(gCmds(i), vbTab)
        res = ""
        gWarn = ""
        Err.Clear
        On Error Resume Next
        res = CopyOne(sm, f)
        If Err.Number <> 0 Then
            W "{""kind"":""res"",""i"":" & i & ",""ok"":false,""err"":" & J(Err.Description) & "}"
        Else
            W "{""kind"":""res"",""i"":" & i & ",""ok"":true" & res & "}"
        End If
        On Error GoTo 0
    Next i
    df.Close
End Sub

Private Function CopyOne(sm As ModelReference, f() As String) As String
    Dim src As Element, ne As Element, le As Element, p() As Point3d, chk() As Point3d
    Dim nv As Long, ra As Long, rb As Long, i As Long, inModel As Boolean, bad As Boolean
    Set src = sm.GetElementByID(DLongFromString(f(0)))
    Set ne = ActiveModelReference.CopyElement(src)
    inModel = (DLongToString(ne.ID) <> "0")
    If f(3) <> "" Then
        p = Pts(f(3))
        ra = CLng(f(1))
        rb = CLng(f(2))
        If ra >= 0 And ne.IsVertexList And (ne.Type = 3 Or ne.Type = 4) Then
            nv = UBound(ne.AsVertexList.GetVertices) + 1
            For i = nv - 1 To rb + 1 Step -1
                ne.AsVertexList.RemoveVertex i
            Next i
            For i = 1 To ra
                ne.AsVertexList.RemoveVertex 0
            Next i
            For i = 0 To UBound(p)
                ne.AsVertexList.ModifyVertex i, p(i)
            Next i
            ' verify the edit really produced the requested vertices
            chk = ne.AsVertexList.GetVertices
            bad = (UBound(chk) <> UBound(p))
            If Not bad Then
                For i = 0 To UBound(p)
                    If Abs(chk(i).X - p(i).X) > 0.0005 Or Abs(chk(i).Y - p(i).Y) > 0.0005 Then bad = True
                Next i
            End If
            If bad Then
                If inModel Then ActiveModelReference.RemoveElement ne
                Err.Raise 5, , "vertex edit check failed"
            End If
            If inModel Then
                ne.Rewrite
            Else
                ActiveModelReference.AddElement ne
            End If
        Else
            Set le = CreateLineElement1(ne, p)
            ActiveModelReference.AddElement le
            If inModel Then ActiveModelReference.RemoveElement ne
            Set ne = le
        End If
    ElseIf Not inModel Then
        ActiveModelReference.AddElement ne
    End If
    CopyOne = ",""id"":""" & DLongToString(ne.ID) & """"
End Function

' ----------------------------------------------------------------- apply ---
Private Sub DoApply()
    Dim i As Long, f() As String, res As String
    For i = 1 To gCmdCount
        f = Split(gCmds(i), vbTab)
        res = ""
        gWarn = ""
        Err.Clear
        On Error Resume Next
        res = ApplyOne(f)
        If Err.Number <> 0 Then
            W "{""kind"":""res"",""i"":" & i & ",""ok"":false,""err"":" & J(Err.Description) & "}"
        Else
            If gWarn <> "" Then res = res & ",""warn"":" & J(gWarn)
            W "{""kind"":""res"",""i"":" & i & ",""ok"":true" & res & "}"
        End If
        On Error GoTo 0
    Next i
End Sub

Private Function Pt(ByVal s As String) As Point3d
    Dim a() As String
    a = Split(s, ",")
    Pt = Point3dFromXYZ(Val(a(0)), Val(a(1)), 0)
End Function

Private Function Pts(ByVal s As String) As Point3d()
    Dim a() As String, r() As Point3d, i As Long
    a = Split(s, ";")
    ReDim r(0 To UBound(a))
    For i = 0 To UBound(a)
        r(i) = Pt(a(i))
    Next i
    Pts = r
End Function

Private Function ById(ByVal s As String) As Element
    Set ById = ActiveModelReference.GetElementByID(DLongFromString(s))
End Function

' fields: lv(code) color weight style  ("" = keep / active)
Private Sub Sym(el As Element, ByVal lv As String, ByVal col As String, ByVal wt As String, ByVal st As String)
    If lv <> "" Then el.Level = ActiveDesignFile.Levels.FindByCode(CLng(lv))
    If col <> "" Then el.Color = CLng(col)
    If wt <> "" Then el.LineWeight = CLng(wt)
    If st <> "" Then
        ' a style missing from this file must not lose the element: keep default, report a warning
        On Error Resume Next
        el.LineStyle = ActiveDesignFile.LineStyles(st)
        If Err.Number <> 0 Then gWarn = "line style not found: " & st
        Err.Clear
        On Error GoTo 0
    End If
End Sub

Private Function Rot(ByVal deg As Double) As Matrix3d
    Rot = Matrix3dFromAxisAndRotationAngle(2, deg * Pi / 180#)
End Function

Private Function ApplyOne(f() As String) As String
    Dim el As Element, p() As Point3d, ts As TextStyle, ids() As String
    Dim arr() As ChainableElement, src() As Element, i As Long, h As Double, q As Point3d, rm As Matrix3d
    Select Case f(0)
        Case "line", "linestring"
            p = Pts(f(5))
            If UBound(p) = 1 Then
                Set el = CreateLineElement2(Nothing, p(0), p(1))
            Else
                Set el = CreateLineElement1(Nothing, p)
            End If
            Sym el, f(1), f(2), f(3), f(4)
            ActiveModelReference.AddElement el
        Case "shape"
            p = Pts(f(5))
            Set el = CreateShapeElement1(Nothing, p, 0)
            Sym el, f(1), f(2), f(3), f(4)
            ActiveModelReference.AddElement el
        Case "text"
            ' f(5)=x,y f(6)=height f(7)=rotation deg f(8)=hex text
            q = Pt(f(5))
            rm = Rot(Val(f(7)))
            Set el = CreateTextElement1(Nothing, HD(f(8)), q, rm)
            h = Val(f(6))
            If h > 0 Then
                Set ts = el.AsTextElement.TextStyle
                ts.Height = h
                ts.Width = h
                Set el.AsTextElement.TextStyle = ts
            End If
            Sym el, f(1), f(2), f(3), f(4)
            ActiveModelReference.AddElement el
        Case "cell"
            ' f(5)=x,y f(6)=scale f(7)=rotation deg f(8)=hex cell name
            h = Val(f(6)): If h <= 0 Then h = 1
            q = Pt(f(5))
            rm = Rot(Val(f(7)))
            Set el = CreateCellElement2(HD(f(8)), q, Point3dFromXYZ(h, h, h), True, rm)
            Sym el, f(1), f(2), f(3), f(4)
            ActiveModelReference.AddElement el
        Case "set"
            ' f(1)=id f(2)=lv f(3)=color f(4)=weight f(5)=style
            Set el = ById(f(1))
            Sym el, f(2), f(3), f(4), f(5)
            el.Rewrite
        Case "move"
            Set el = ById(f(1))
            q = Pt(f(2))
            el.Move q
            el.Rewrite
        Case "text_set"
            ' f(1)=id f(2)=hex text
            Set el = ById(f(1))
            el.AsTextElement.Text = HD(f(2))
            el.Rewrite
        Case "delete"
            Set el = ById(f(1))
            ActiveModelReference.RemoveElement el
            ApplyOne = ",""id"":""" & f(1) & """"
            Exit Function
        Case "join"
            ' f(1..4)=sym f(5)=ids(comma) f(6)="shape"|"chain" f(7)=gap tol f(8)="1" delete originals
            ids = Split(f(5), ",")
            ReDim arr(0 To UBound(ids))
            ReDim src(0 To UBound(ids))
            For i = 0 To UBound(ids)
                Set src(i) = ById(ids(i))
                Set arr(i) = src(i)
            Next i
            If f(6) = "shape" Then
                Set el = CreateComplexShapeElement2(arr, 0, Val(f(7)))
            Else
                Set el = CreateComplexStringElement2(arr, Val(f(7)))
            End If
            Sym el, f(1), f(2), f(3), f(4)
            ActiveModelReference.AddElement el
            If f(8) = "1" Then
                For i = 0 To UBound(src)
                    ActiveModelReference.RemoveElement src(i)
                Next i
            End If
        Case Else
            Err.Raise 5, , "unknown command " & f(0)
    End Select
    ApplyOne = ",""id"":""" & DLongToString(el.ID) & """"
End Function
