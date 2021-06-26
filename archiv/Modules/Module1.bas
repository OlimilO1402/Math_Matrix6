Attribute VB_Name = "Module1"
Option Explicit

Public Function CMatOp(ByVal p As Long, s As String) As CMatOp 'ByVal mRows As Byte, ByVal nCols As Byte) As CMatOp
    Set CMatOp = New CMatOp: CMatOp.New_ p, s 'mRows, nCols
End Function

Public Function CMatOpByVal(ByVal pDbl As Long, Optional ByVal mRows As Byte = 1, Optional ByVal nCols As Byte = 1) As CMatOp
    Set CMatOpByVal = New CMatOp: CMatOpByVal.NewByVal pDbl, mRows, nCols
End Function


Public Function Splitter(BolMDI As Boolean, MyOwner As Object, MyContainer As Object, Name As String, LeftTop As Control, RghtBot As Control) As Splitter
    Set Splitter = New Splitter: Splitter.New_ BolMDI, MyOwner, MyContainer, Name, LeftTop, RghtBot
End Function

'Hilfsfunktionen
Public Function DeleteMultiWS(s As String) As String
    DeleteMultiWS = Trim$(s)
    If InStr(1, s, "  ") = 0 Then Exit Function
    DeleteMultiWS = Replace(s, "  ", " ")
    DeleteMultiWS = DeleteMultiWS(DeleteMultiWS)
End Function
Public Function DeleteCRLF(s As String) As String
    DeleteCRLF = Trim$(s)
    If InStr(1, s, vbLf) = 0 Then Exit Function
    If InStr(1, s, vbCr) = 0 Then Exit Function
    DeleteCRLF = Replace(Replace(Replace(s, vbCrLf, " "), vbLf, " "), vbCr, " ")
    DeleteCRLF = DeleteCRLF(DeleteCRLF)
End Function

Public Function Max(V1, V2)
    If V1 > V2 Then Max = V1 Else Max = V2
End Function
Public Function Min(V1, V2)
    If V1 < V2 Then Min = V1 Else Min = V2
End Function
Public Function TStr(d As Double) As String
    TStr = Trim$(Str$(d))
End Function
Public Function DblParse(ByVal s As String) As Double
Try: On Error GoTo Catch
    s = Replace(Trim$(s), ",", ".")
    DblParse = Val(s)
Catch: 'out
End Function

Public Function Double_TryParse(ByVal s As String, ByRef d_out As Double) As Boolean
Try: On Error GoTo Catch
    s = Replace(Trim$(s), ",", ".")
    d_out = Val(s)
    Double_TryParse = True
Catch: 'out
End Function

Public Function PadLeft(StrVal As String, _
                        ByVal totalWidth As Long, _
                        Optional ByVal paddingChar As String) As String
    ' der String wird mit der angegebenen Länge zurückgegeben, der
    ' String wird nach rechts gerückt, und links mit PadChar aufgefüllt
    ' ist PadChar nicht angegeben, so wird mit RSet der String in
    ' Spaces eingefügt.
    If Len(paddingChar) Then
        If Len(StrVal) <= totalWidth Then _
            PadLeft = StrVal & String$(totalWidth - Len(StrVal), paddingChar)
    Else
        PadLeft = Space$(totalWidth)
        RSet PadLeft = StrVal
    End If
End Function
Public Function PadRight(StrVal As String, _
                         ByVal totalWidth As Long, _
                         Optional ByVal paddingChar As String) As String
    ' der String wird mit der angegebenen Länge zurückgegeben, der
    ' String wird nach links gerückt, und rechts mit PadChar aufgefüllt
    ' ist PadChar nicht angegeben, so wird mit LSet der String in
    ' Spaces eingefügt.
    If Len(paddingChar) Then
        If Len(StrVal) <= totalWidth Then _
            PadRight = StrVal & String$(totalWidth - Len(StrVal), paddingChar)
    Else
        PadRight = Space$(totalWidth)
        LSet PadRight = StrVal
    End If
End Function



'Sicherheitskopie
'Ausgabefunktionen
'Public Function Matrix2_ToStr(m As Matrix2) As String
''    Dim s As String: s = ""
''    With m
''        s = s & TStr(.aa) & " " & TStr(.ab) & vbCrLf
''        s = s & TStr(.ba) & " " & TStr(.bb)
''    End With
''    Matrix2_ToStr = s
'    Matrix2_ToStr = MatrixA_ToStr(Matrix2_ToArr(m), 2, 2)
'End Function
'Public Function Matrix3_ToStr(m As Matrix3) As String
''    Dim s As String: s = ""
''    With m
''        s = s & TStr(.aa) & " " & TStr(.ab) & " " & TStr(.ac) & vbCrLf
''        s = s & TStr(.ba) & " " & TStr(.bb) & " " & TStr(.bc) & vbCrLf
''        s = s & TStr(.ca) & " " & TStr(.cb) & " " & TStr(.cc)
''    End With
''    Matrix3_ToStr = s
'    Matrix3_ToStr = MatrixA_ToStr(Matrix3_ToArr(m), 3, 3)
'End Function
'Public Function Matrix4_ToStr(m As Matrix4) As String
''    Dim s As String: s = ""
''    With m
''        s = s & TStr(.aa) & " " & TStr(.ab) & " " & TStr(.ac) & " " & TStr(.ad) & vbCrLf
''        s = s & TStr(.ba) & " " & TStr(.bb) & " " & TStr(.bc) & " " & TStr(.bd) & vbCrLf
''        s = s & TStr(.ca) & " " & TStr(.cb) & " " & TStr(.cc) & " " & TStr(.cd) & vbCrLf
''        s = s & TStr(.da) & " " & TStr(.db) & " " & TStr(.dc) & " " & TStr(.dd)
''    End With
''    Matrix4_ToStr = s
'    Matrix4_ToStr = MatrixA_ToStr(Matrix4_ToArr(m), 4, 4)
'End Function
'Public Function Matrix5_ToStr(m As Matrix5) As String
''    Dim s As String: s = ""
''    With m
''        s = s & TStr(.aa) & " " & TStr(.ab) & " " & TStr(.ac) & " " & TStr(.ad) & " " & TStr(.ae) & vbCrLf
''        s = s & TStr(.ba) & " " & TStr(.bb) & " " & TStr(.bc) & " " & TStr(.bd) & " " & TStr(.be) & vbCrLf
''        s = s & TStr(.ca) & " " & TStr(.cb) & " " & TStr(.cc) & " " & TStr(.cd) & " " & TStr(.ce) & vbCrLf
''        s = s & TStr(.da) & " " & TStr(.db) & " " & TStr(.dc) & " " & TStr(.dd) & " " & TStr(.de) & vbCrLf
''        s = s & TStr(.ea) & " " & TStr(.eb) & " " & TStr(.ec) & " " & TStr(.ed) & " " & TStr(.ee)
''    End With
''    Matrix5_ToStr = s
'    Matrix5_ToStr = MatrixA_ToStr(Matrix5_ToArr(m), 5, 5)
'End Function
'Public Function Matrix6_ToStr(m As Matrix6) As String
''    Dim s As String: s = ""
''    With m
''        s = s & TStr(.aa) & " " & TStr(.ab) & " " & TStr(.ac) & " " & TStr(.ad) & " " & TStr(.ae) & " " & TStr(.af) & vbCrLf
''        s = s & TStr(.ba) & " " & TStr(.bb) & " " & TStr(.bc) & " " & TStr(.bd) & " " & TStr(.be) & " " & TStr(.bf) & vbCrLf
''        s = s & TStr(.ca) & " " & TStr(.cb) & " " & TStr(.cc) & " " & TStr(.cd) & " " & TStr(.ce) & " " & TStr(.cf) & vbCrLf
''        s = s & TStr(.da) & " " & TStr(.db) & " " & TStr(.dc) & " " & TStr(.dd) & " " & TStr(.de) & " " & TStr(.df) & vbCrLf
''        s = s & TStr(.ea) & " " & TStr(.eb) & " " & TStr(.ec) & " " & TStr(.ed) & " " & TStr(.ee) & " " & TStr(.ef) & vbCrLf
''        s = s & TStr(.fa) & " " & TStr(.fb) & " " & TStr(.fc) & " " & TStr(.fd) & " " & TStr(.fe) & " " & TStr(.ff)
''    End With
''    Matrix6_ToStr = s
'    Matrix6_ToStr = MatrixA_ToStr(Matrix6_ToArr(m), 6, 6)
'End Function
'
'
''allgemein
''Public Function MatrixT_ToStr(m As MatrixT, ByVal mRows As Long, ByVal nCols As Long) As String
''    MatrixT_ToStr = MatrixA_ToStr(m.a, mRows, nCols)
''End Function
'
'
