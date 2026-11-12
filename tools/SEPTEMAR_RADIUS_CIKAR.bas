Attribute VB_Name = "SEPTEMAR_RADIUS_CIKAR"
'==============================================================================
'  SEPTEMAR RADIUS CIKAR  -  r1
'  Aktif MONTAJDAKI tum parcalarin silindirik / torus / kuresel yuzlerinin
'  GERCEK yaricaplarini (SolidWorks'un kendi yuz bilgisinden) okur ve montajin
'  yanina  "<montaj adi>_radius.json"  dosyasi yazar.
'  Uygulama bu dosyayi GLB ile birlikte yukleyince, usta yuvarlak bir yuze
'  dokundugunda tahmin degil SolidWorks'teki R degeri gorunur.
'  Kullanim: montaji ac -> bu makroyu calistir -> cikan *_radius.json dosyasini
'  GLB ile ayni yere (models klasoru) "<glb adi>.radius.json" adiyla yukle.
'==============================================================================
Option Explicit

Private gSeen As Object      ' dosya|konf -> satir (ayni parca tekrar taranmasin)
Private gParca As Long

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

' bir parcanin yuz ozeti: "tip,R,..." satirlari (mm)
Private Function ParcaYuzleri(ByVal pd As Object) As String
    Dim bodies As Variant, b As Variant, faces As Variant, f As Variant
    Dim surf As Object, p As Variant, key As String, d As Object, k As Variant, r As String
    Set d = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    bodies = pd.GetBodies2(0, True)   ' 0 = swSolidBody, gorunur govdeler
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
                        If Not d.Exists(key) Then d.Add key, "{""t"":""C"",""R"":" & N(p(6) * 1000#) & ",""c"":[" & N(p(0) * 1000#) & "," & N(p(1) * 1000#) & "," & N(p(2) * 1000#) & "],""a"":[" & N(p(3)) & "," & N(p(4)) & "," & N(p(5)) & "]}"
                    ElseIf surf.IsTorus Then
                        p = surf.TorusParams
                        key = "T|" & N(p(6) * 1000#) & "|" & N(p(7) * 1000#)
                        If Not d.Exists(key) Then d.Add key, "{""t"":""T"",""R"":" & N(p(7) * 1000#) & ",""R2"":" & N(p(6) * 1000#) & ",""c"":[" & N(p(0) * 1000#) & "," & N(p(1) * 1000#) & "," & N(p(2) * 1000#) & "],""a"":[" & N(p(3)) & "," & N(p(4)) & "," & N(p(5)) & "]}"
                    ElseIf surf.IsSphere Then
                        p = surf.SphereParams
                        key = "S|" & N(p(3) * 1000#)
                        If Not d.Exists(key) Then d.Add key, "{""t"":""S"",""R"":" & N(p(3) * 1000#) & ",""c"":[" & N(p(0) * 1000#) & "," & N(p(1) * 1000#) & "," & N(p(2) * 1000#) & "]}"
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
    Dim pd As Object, key As String, fy As String, i As Long
    Dim fso As Object, ts As Object, yol As String, ilk As Boolean

    Set swApp = Application.SldWorks
    Set md = swApp.ActiveDoc
    If md Is Nothing Then MsgBox "Once bir montaj ac.", vbExclamation: Exit Sub
    If md.GetType <> 2 Then MsgBox "Bu makro montaj (.sldasm) icindir.", vbExclamation: Exit Sub

    On Error Resume Next
    md.ResolveAllLightweightComponents True
    On Error GoTo 0

    Set gSeen = CreateObject("Scripting.Dictionary")
    comps = md.GetComponents(False)    ' tum seviyeler
    If IsEmpty(comps) Then MsgBox "Montajda bilesen yok.", vbExclamation: Exit Sub

    Dim sb As String, sayi As Long
    sb = "{""v"":1,""montaj"":""" & JS(md.GetTitle) & """,""birim"":""mm"",""parcalar"":{" & vbLf
    ilk = True
    For i = LBound(comps) To UBound(comps)
        Set c = comps(i)
        If Not c Is Nothing Then
            If c.GetSuppression = 2 Then      ' 2 = swComponentFullyResolved
                Set pd = Nothing
                On Error Resume Next
                Set pd = c.GetModelDoc2
                On Error GoTo 0
                If Not pd Is Nothing Then
                    If pd.GetType = 1 Then   ' parca
                        key = c.GetPathName & "|" & c.ReferencedConfiguration
                        If gSeen.Exists(key) Then
                            fy = gSeen(key)
                        Else
                            fy = ParcaYuzleri(pd)
                            gSeen.Add key, fy
                            gParca = gParca + 1
                        End If
                        If Len(fy) > 0 Then
                            If Not ilk Then sb = sb & "," & vbLf
                            ilk = False
                            sb = sb & """" & JS(c.Name2) & """:[" & fy & "]"
                            sayi = sayi + 1
                        End If
                    End If
                End If
            End If
        End If
    Next i
    sb = sb & vbLf & "}}"

    yol = md.GetPathName
    If Len(yol) = 0 Then MsgBox "Montaji once kaydet.", vbExclamation: Exit Sub
    yol = Left$(yol, InStrRev(yol, ".") - 1) & "_radius.json"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(yol, True, False)
    ts.Write sb
    ts.Close
    MsgBox "Tamam." & vbCrLf & sayi & " bilesen, " & gParca & " farkli parca tarandi." & vbCrLf & yol, vbInformation, "SEPTEMAR RADIUS CIKAR"
End Sub
