Attribute VB_Name = "SeptemarGonder"
Option Explicit
Dim swApp As Object

Function J(ByVal s As String) As String
    s = Replace(s, "\", "\\"): s = Replace(s, """", "\""")
    s = Replace(s, vbCr, " "): s = Replace(s, vbLf, " "): s = Replace(s, vbTab, " ")
    J = s
End Function

Function Pr(cpm As Object, ByVal nm As String) As String
    Dim v As String, r As String, w As Boolean
    On Error Resume Next
    cpm.Get5 nm, False, v, r, w
    Pr = Trim(v)
End Function

Function N(ByVal s As String) As String
    s = Replace(s, ",", ".")
    If s = "" Then s = "0"
    If Not IsNumeric(Replace(s, ".", ",")) And Not IsNumeric(s) Then s = "0"
    N = s
End Function

Function Side(cpm As Object, ByVal sd As String) As String
    Side = """" & sd & """:{""ad"":""" & J(Pr(cpm, "SWOODCP_Edge" & sd & "Name")) & """,""mal"":""" & J(Pr(cpm, "SWOODCP_Edge" & sd & "Material")) & """,""kal"":" & N(Pr(cpm, "SWOODCP_Edge" & sd & "Thickness")) & ",""boy"":" & N(Pr(cpm, "SWOODCP_Edge" & sd & "Length")) & "}"
End Function

Sub AddTo(d As Object, seen As Object, ByVal mat As String, ByVal rol As String, ByVal leaf As String)
    Dim k As String
    If mat = "" Then Exit Sub
    k = mat & "|" & rol & "|" & leaf
    If seen.Exists(k) Then Exit Sub
    seen.Add k, 1
    If Not d.Exists(rol & "|" & mat) Then
        d.Add rol & "|" & mat, """" & J(leaf) & """"
    Else
        d(rol & "|" & mat) = d(rol & "|" & mat) & ",""" & J(leaf) & """"
    End If
End Sub

Function Diger(d As Object, cpm As Object, ByVal cfg As String) As String
    Dim mn As String, db As String, nms As Variant, i As Long, r As String, f As Boolean, gp As Object
    On Error Resume Next
    mn = d.GetMaterialPropertyName2(cfg, db)
    r = """mat"":""" & J(mn) & """,""db"":""" & J(db) & """,""ozellik"":{"
    f = True
    Set gp = d.Extension.CustomPropertyManager("")
    nms = Empty: nms = cpm.GetNames
    If Not IsEmpty(nms) Then
        For i = LBound(nms) To UBound(nms)
            If Not f Then r = r & ","
            f = False
            r = r & """" & J(CStr(nms(i))) & """:""" & J(Pr(cpm, CStr(nms(i)))) & """"
        Next i
    End If
    nms = Empty: nms = gp.GetNames
    If Not IsEmpty(nms) Then
        For i = LBound(nms) To UBound(nms)
            If Not f Then r = r & ","
            f = False
            r = r & """g:" & J(CStr(nms(i))) & """:""" & J(Pr(gp, CStr(nms(i)))) & """"
        Next i
    End If
    Diger = r & "}"
End Function

Function Ozel(d As Object, cpm As Object) As String
    Dim nms As Variant, i As Long, r As String, f As Boolean, gp As Object
    On Error Resume Next
    f = True
    Set gp = d.Extension.CustomPropertyManager("")
    nms = Empty: nms = cpm.GetNames
    If Not IsEmpty(nms) Then
        For i = LBound(nms) To UBound(nms)
            If Not f Then r = r & ","
            f = False
            r = r & """" & J(CStr(nms(i))) & """:""" & J(Pr(cpm, CStr(nms(i)))) & """"
        Next i
    End If
    nms = Empty: nms = gp.GetNames
    If Not IsEmpty(nms) Then
        For i = LBound(nms) To UBound(nms)
            If Not f Then r = r & ","
            f = False
            r = r & """g:" & J(CStr(nms(i))) & """:""" & J(Pr(gp, CStr(nms(i)))) & """"
        Next i
    End If
    Ozel = r
End Function

Function GetL(d As Object, ByVal rol As String, ByVal mat As String) As String
    If d.Exists(rol & "|" & mat) Then GetL = d(rol & "|" & mat) Else GetL = ""
End Function

Sub main()
    Dim doc As Object, comps As Variant, i As Long, c As Object, d As Object
    Dim cpm As Object, nm As String, leaf As String, pj As String, first As Boolean
    Dim base As String, kod As String, folder As String, st As Object
    Dim dic As Object, seen As Object, allm As Object, k As Variant, mj As String
    Dim tp As String, bt As String, cr As String, sp As Long, cnt As Long
    Dim dg As String, dgf As Boolean, dseen As Object, lf2 As String
    Dim op As String, pseen As Object, gp0 As Object
    On Error GoTo EH
    Set swApp = Application.SldWorks
    Set doc = swApp.ActiveDoc
    If doc Is Nothing Then MsgBox "Once montaji ac": Exit Sub
    If doc.GetType <> 2 Then MsgBox "Aktif dosya montaj (.sldasm) olmali": Exit Sub
    base = doc.GetPathName
    folder = Left(base, InStrRev(base, "\")) & "Septemar\"
    If Dir(folder, vbDirectory) = "" Then MkDir folder
    kod = Mid(base, InStrRev(base, "\") + 1)
    sp = InStr(kod, " ")
    If sp > 1 Then kod = Left(kod, sp - 1) Else kod = Left(kod, InStrRev(kod, ".") - 1)
    Set dic = CreateObject("Scripting.Dictionary")
    Set seen = CreateObject("Scripting.Dictionary")
    Set allm = CreateObject("Scripting.Dictionary")
    comps = doc.GetComponents(False)
    pj = "": first = True: cnt = 0: dg = "": dgf = True
    Set dseen = CreateObject("Scripting.Dictionary")
    Set pseen = CreateObject("Scripting.Dictionary")
    op = ""
    For i = LBound(comps) To UBound(comps)
        Set c = comps(i)
        Set d = Nothing
        On Error Resume Next
        Set d = c.GetModelDoc2
        On Error GoTo EH
        If Not d Is Nothing Then
            Set cpm = d.Extension.CustomPropertyManager(c.ReferencedConfiguration)
            If Pr(cpm, "SWOODCP_PanelLength") <> "" Then
                nm = c.Name2
                leaf = Mid(nm, InStrRev(nm, "/") + 1)
                If InStrRev(leaf, "-") > 0 Then
                    If IsNumeric(Mid(leaf, InStrRev(leaf, "-") + 1)) Then leaf = Left(leaf, InStrRev(leaf, "-") - 1)
                End If
                tp = Pr(cpm, "SWOODCP_TopStockMaterial")
                bt = Pr(cpm, "SWOODCP_BottomStockMaterial")
                cr = Pr(cpm, "SWOODCP_CoreStockMaterial")
                Set gp0 = d.Extension.CustomPropertyManager("")
                If tp = "" Then tp = Pr(gp0, "SWOODLAMINATETOP_Material")
                If bt = "" Then bt = Pr(gp0, "SWOODLAMINATEBOTTOM_Material")
                AddTo dic, seen, tp, "ust", leaf
                AddTo dic, seen, bt, "alt", leaf
                AddTo dic, seen, cr, "cekirdek", leaf
                If tp <> "" And Not allm.Exists(tp) Then allm.Add tp, 1
                If bt <> "" And Not allm.Exists(bt) Then allm.Add bt, 1
                If cr <> "" And Not allm.Exists(cr) Then allm.Add cr, 1
                If Not first Then pj = pj & ","
                first = False
                cnt = cnt + 1
                pj = pj & """" & J(leaf) & """:{""top"":""" & J(tp) & """,""bot"":""" & J(bt) & """,""core"":""" & J(cr) & """,""desen"":""" & J(Pr(cpm, "Desen_Takibi")) & """,""parkur"":""" & J(Pr(cpm, "Makina_Parkuru")) & """,""boy"":" & N(Pr(cpm, "SWOODCP_PanelLength")) & ",""en"":" & N(Pr(cpm, "SWOODCP_PanelWidth")) & ",""kal"":" & N(Pr(cpm, "SWOODCP_PanelThickness")) & ",""kenar"":{" & Side(cpm, "Left") & "," & Side(cpm, "Front") & "," & Side(cpm, "Right") & "," & Side(cpm, "Back") & "}}"
            ElseIf d.GetType = 1 Then
                lf2 = Mid(c.Name2, InStrRev(c.Name2, "/") + 1)
                If InStrRev(lf2, "-") > 0 Then
                    If IsNumeric(Mid(lf2, InStrRev(lf2, "-") + 1)) Then lf2 = Left(lf2, InStrRev(lf2, "-") - 1)
                End If
                If Not dseen.Exists(lf2) Then
                    dseen.Add lf2, 1
                    If Not dgf Then dg = dg & ","
                    dgf = False
                    dg = dg & """" & J(lf2) & """:{" & Diger(d, cpm, c.ReferencedConfiguration) & "}"
                End If
            End If
        End If
    Next i
    mj = "": first = True
    For Each k In allm.Keys
        If Not first Then mj = mj & ","
        first = False
        mj = mj & """" & J(CStr(k)) & """:{""ust"":[" & GetL(dic, "ust", CStr(k)) & "],""alt"":[" & GetL(dic, "alt", CStr(k)) & "],""cekirdek"":[" & GetL(dic, "cekirdek", CStr(k)) & "]}"
    Next k
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2: st.Charset = "utf-8": st.Open
    st.WriteText "{""paneller"":{" & pj & "},""malzemeler"":{" & mj & "},""diger"":{" & dg & "}}"
    st.SaveToFile folder & kod & "-sw.json", 2
    st.Close
    MsgBox "JSON hazir (" & cnt & " panel, " & allm.Count & " malzeme):" & vbCrLf & folder & kod & "-sw.json" & vbCrLf & "GLB: Visualize ile cevir."
    Exit Sub
EH:
    MsgBox "Hata: " & Err.Description & " (satir " & Erl & ")"
End Sub
