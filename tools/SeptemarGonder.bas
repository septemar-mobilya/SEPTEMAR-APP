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

Function GetL(d As Object, ByVal rol As String, ByVal mat As String) As String
    If d.Exists(rol & "|" & mat) Then GetL = d(rol & "|" & mat) Else GetL = ""
End Function

Sub main()
    Dim doc As Object, comps As Variant, i As Long, c As Object, d As Object
    Dim cpm As Object, nm As String, leaf As String, pj As String, first As Boolean
    Dim base As String, kod As String, folder As String, st As Object, e As Long, w As Long
    Dim dic As Object, seen As Object, allm As Object, k As Variant, mj As String
    Dim tp As String, bt As String, cr As String, sp As Long, cnt As Long
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
    pj = "": first = True: cnt = 0
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
    st.WriteText "{""paneller"":{" & pj & "},""malzemeler"":{" & mj & "}}"
    st.SaveToFile folder & kod & "-sw.json", 2
    st.Close
    doc.Extension.SaveAs3 folder & kod & ".glb", 0, 1, Nothing, Nothing, e, w
    If Dir(folder & kod & ".glb") = "" Then
        MsgBox "JSON hazir (" & cnt & " panel). GLB yazilamadi, onu elle kaydet." & vbCrLf & folder
    Else
        MsgBox "Hazir (" & cnt & " panel, " & allm.Count & " malzeme):" & vbCrLf & folder
    End If
    Exit Sub
EH:
    MsgBox "Hata: " & Err.Description & " (satir " & Erl & ")"
End Sub
