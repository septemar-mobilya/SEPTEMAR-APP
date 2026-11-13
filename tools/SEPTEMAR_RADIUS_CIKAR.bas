Attribute VB_Name = "SEPTEMAR_RADIUS_CIKAR"
'==============================================================================
'  SEPTEMAR RADIUS CIKAR  -  r2
'  Aktif MONTAJDAKI tum parcalarin DAIRESEL KENARLARINI (koselerdeki yay cizgileri)
'  ve silindir/torus/kure yuzlerini SolidWorks'un kendi bilgisinden okur;
'  montajin yanina "<montaj adi>_radius.json" yazar.
'  Uygulama bu kenarlari cizgi olarak gosterir; usta cizgiye dokununca
'  SolidWorks'teki gercek R gorunur (tahmin yok).
'  Kullanim: montaji ac ve kaydet -> makroyu calistir -> cikan dosyayi
'  GLB ile ayni yere "<glb adi>.radius.json" adiyla yukle.
'==============================================================================
Option Explicit

Private gSeen As Object      ' dosya|konf -> ham yuz/kenar verisi (parca koordinatinda)
Private gMU As Object

Private Function JS(ByVal s As String) As String
    Dim i As Long, c As Long, o As String
    For i = 1 To Len(s)
        c = AscW(Mid$(s, i, 1))
        If c < 0 Then c = c + 65536
        If c = 34 Then
            o = o & "\"""
        ElseIf c = 92 Then
            o = o & "\\"
        ElseIf c < 32 Then
            o = o & " "
        ElseIf c > 126 Then
            o = o & "\u" & Right$("0000" & Hex$(c), 4)
        Else
            o = o & Chr$(c)
        End If
    Next i
    JS = o
End Function

Private Function N(ByVal v As Double) As String
    N = Replace(Format$(v, "0.000"), ",", ".")
End Function

' parca koordinatindaki nokta/yonu bilesen donusumuyle montaj koordinatina cevirir (mm)
Private Function Pt(ByVal xf As Object, ByVal x As Double, ByVal y As Double, ByVal z As Double) As String
    Dim mp As Object, r As Object, v As Variant
    Set mp = gMU.CreatePoint(Array(x, y, z))
    Set r = mp.MultiplyTransform(xf)
    v = r.ArrayData
    Pt = "[" & N(v(0) * 1000#) & "," & N(v(1) * 1000#) & "," & N(v(2) * 1000#) & "]"
End Function

Private Function Dr(ByVal xf As Object, ByVal x As Double, ByVal y As Double, ByVal z As Double) As String
    Dim mv As Object, r As Object, v As Variant
    Set mv = gMU.CreateVector(Array(x, y, z))
    Set r = mv.MultiplyTransform(xf)
    v = r.ArrayData
    Dr = "[" & N(v(0)) & "," & N(v(1)) & "," & N(v(2)) & "]"
End Function

' parcanin dairesel kenarlari: Array(R, cx,cy,cz, ax,ay,az, sx,sy,sz, mx,my,mz, ex,ey,ez)
Private Function ParcaKenarlari(ByVal pd As Object) As Collection
    Dim res As New Collection
    Dim bodies As Variant, b As Variant, edges As Variant, e As Variant
    Dim cv As Object, p As Variant, sp As Variant, ep As Variant, mp As Variant
    Dim seen As Object, key As String
    Dim t0 As Double, t1 As Double, isCl As Boolean, isPer As Boolean, cpd As Object
    Set seen = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    bodies = pd.GetBodies2(0, True)
    If IsEmpty(bodies) Then Set ParcaKenarlari = res: Exit Function
    For Each b In bodies
        edges = b.GetEdges
        If Not IsEmpty(edges) Then
            For Each e In edges
                Set cv = Nothing
                Set cv = e.GetCurve
                If Not cv Is Nothing Then
                    If cv.IsCircle Then
                        p = cv.CircleParams    ' cx,cy,cz, ax,ay,az, R
                        Set cpd = Nothing
                        Set cpd = e.GetCurveParams2        ' kenarin KIRPILMIS (gercek) baslangic/bitis parametreleri
                        If Not cpd Is Nothing Then
                            t0 = cpd.StartParam
                            t1 = cpd.EndParam
                        Else
                            cv.GetEndParams t0, t1, isCl, isPer
                        End If
                        sp = cv.Evaluate2(t0, 0)
                        ep = cv.Evaluate2(t1, 0)
                        mp = cv.Evaluate2((t0 + t1) / 2#, 0)
                        key = N(p(0)) & "|" & N(p(1)) & "|" & N(p(2)) & "|" & N(p(6)) & "|" & N(sp(0)) & "|" & N(sp(1)) & "|" & N(sp(2)) & "|" & N(ep(0)) & "|" & N(ep(1)) & "|" & N(ep(2))
                        If Not seen.Exists(key) Then
                            seen.Add key, 1
                            res.Add Array(p(6), p(0), p(1), p(2), p(3), p(4), p(5), sp(0), sp(1), sp(2), mp(0), mp(1), mp(2), ep(0), ep(1), ep(2))
                        End If
                    End If
                End If
            Next e
        End If
    Next b
    On Error GoTo 0
    Set ParcaKenarlari = res
End Function

' parcanin silindir/torus/kure yuzleri (yedek veri) - JSON parcasi
Private Function ParcaYuzleri(ByVal pd As Object, ByVal xf As Object) As String
    Dim bodies As Variant, b As Variant, faces As Variant, f As Variant
    Dim surf As Object, p As Variant, key As String, d As Object, k As Variant, r As String
    Set d = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    bodies = pd.GetBodies2(0, True)
    If IsEmpty(bodies) Then Exit Function
    For Each b In bodies
        faces = b.GetFaces
        If Not IsEmpty(faces) Then
            For Each f In faces
                Set surf = f.GetSurface
                If Not surf Is Nothing Then
                    If surf.IsCylinder Then
                        p = surf.CylinderParams
                        key = "C|" & N(p(6) * 1000#)
                        If Not d.Exists(key) Then d.Add key, "{""t"":""C"",""R"":" & N(p(6) * 1000#) & ",""c"":" & Pt(xf, p(0), p(1), p(2)) & ",""a"":" & Dr(xf, p(3), p(4), p(5)) & "}"
                    ElseIf surf.IsTorus Then
                        p = surf.TorusParams
                        key = "T|" & N(p(6) * 1000#) & "|" & N(p(7) * 1000#)
                        If Not d.Exists(key) Then d.Add key, "{""t"":""T"",""R"":" & N(p(7) * 1000#) & ",""R2"":" & N(p(6) * 1000#) & ",""c"":" & Pt(xf, p(0), p(1), p(2)) & ",""a"":" & Dr(xf, p(3), p(4), p(5)) & "}"
                    ElseIf surf.IsSphere Then
                        p = surf.SphereParams
                        key = "S|" & N(p(3) * 1000#)
                        If Not d.Exists(key) Then d.Add key, "{""t"":""S"",""R"":" & N(p(3) * 1000#) & ",""c"":" & Pt(xf, p(0), p(1), p(2)) & "}"
                    End If
                End If
            Next f
        End If
    Next b
    On Error GoTo 0
    For Each k In d.Keys
        If Len(r) > 0 Then r = r & ","
        r = r & d(k)
    Next k
    ParcaYuzleri = r
End Function

Sub main()
    Dim swApp As Object, md As Object, comps As Variant, c As Object
    Dim pd As Object, key As String, i As Long, ilk As Boolean
    Dim fso As Object, ts As Object, yol As String
    Dim xf As Object, ed As Collection, a As Variant, kn As String, ys As String
    Dim bx As Variant, sb As String, sayi As Long, kenarSay As Long

    Set swApp = Application.SldWorks
    Set md = swApp.ActiveDoc
    If md Is Nothing Then MsgBox "Once bir montaj ac.", vbExclamation: Exit Sub
    If md.GetType <> 2 Then MsgBox "Bu makro montaj (.sldasm) icindir.", vbExclamation: Exit Sub
    If Len(md.GetPathName) = 0 Then MsgBox "Montaji once kaydet.", vbExclamation: Exit Sub

    On Error Resume Next
    md.ResolveAllLightweightComponents True
    On Error GoTo 0

    Set gMU = swApp.GetMathUtility
    Set gSeen = CreateObject("Scripting.Dictionary")
    comps = md.GetComponents(False)
    If IsEmpty(comps) Then MsgBox "Montajda bilesen yok.", vbExclamation: Exit Sub

    sb = "{""v"":2,""montaj"":""" & JS(md.GetTitle) & """,""birim"":""mm"",""parcalar"":[" & vbLf
    ilk = True
    For i = LBound(comps) To UBound(comps)
        Set c = comps(i)
        If Not c Is Nothing Then
            If c.GetSuppression = 2 Then
                Set pd = Nothing
                On Error Resume Next
                Set pd = c.GetModelDoc2
                On Error GoTo 0
                If Not pd Is Nothing Then
                    If pd.GetType = 1 Then
                        key = c.GetPathName & "|" & c.ReferencedConfiguration
                        If gSeen.Exists(key) Then
                            Set ed = gSeen(key)
                        Else
                            Set ed = ParcaKenarlari(pd)
                            gSeen.Add key, ed
                        End If
                        Set xf = c.Transform2
                        ys = ParcaYuzleri(pd, xf)
                        If ed.Count > 0 Or Len(ys) > 0 Then
                            If Not ilk Then sb = sb & "," & vbLf
                            ilk = False
                            sb = sb & "{""n"":""" & JS(c.Name2) & """"
                            On Error Resume Next
                            bx = c.GetBox(False, False)
                            On Error GoTo 0
                            If Not IsEmpty(bx) Then sb = sb & ",""b"":[" & N(bx(0) * 1000#) & "," & N(bx(1) * 1000#) & "," & N(bx(2) * 1000#) & "," & N(bx(3) * 1000#) & "," & N(bx(4) * 1000#) & "," & N(bx(5) * 1000#) & "]"
                            sb = sb & ",""y"":[" & ys & "],""k"":["
                            Dim j As Long
                            For j = 1 To ed.Count
                                a = ed(j)
                                If j > 1 Then sb = sb & ","
                                sb = sb & "{""R"":" & N(a(0) * 1000#) & ",""c"":" & Pt(xf, a(1), a(2), a(3)) & ",""a"":" & Dr(xf, a(4), a(5), a(6)) & ",""p"":[" & Pt(xf, a(7), a(8), a(9)) & "," & Pt(xf, a(10), a(11), a(12)) & "," & Pt(xf, a(13), a(14), a(15)) & "]}"
                                kenarSay = kenarSay + 1
                            Next j
                            sb = sb & "]}"
                            sayi = sayi + 1
                        End If
                    End If
                End If
            End If
        End If
    Next i
    sb = sb & vbLf & "]}"

    yol = md.GetPathName
    yol = Left$(yol, InStrRev(yol, ".") - 1) & "_radius.json"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(yol, True, False)
    ts.Write sb
    ts.Close
    MsgBox "Tamam." & vbCrLf & sayi & " bilesen, " & kenarSay & " dairesel kenar." & vbCrLf & yol, vbInformation, "SEPTEMAR RADIUS CIKAR"
End Sub
