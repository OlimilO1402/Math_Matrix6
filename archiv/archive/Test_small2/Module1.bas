Attribute VB_Name = "Module1"
Option Explicit

Public Type Vector2
    a As Double
    b As Double
End Type

Public Type Vector3
    a As Double
    b As Double
    c As Double
End Type

'Matrix 2x2
'bsp "ab" ist das zweite Element der ersten  Zeile,
        'oder das  erste Element der zweiten Spalte
Public Type Matrix2
    aa As Double:    ab As Double
    ba As Double:    bb As Double
End Type

Public Type Matrix3 'Matrix 3x3
    aa As Double:    ab As Double:    ac As Double
    ba As Double:    bb As Double:    bc As Double
    ca As Double:    cb As Double:    cc As Double
End Type


Public Function Vec3(a As Double, b As Double, c As Double) As Vector3
    'Erzeugt einen Vector mit 3 Elementen
    With Vec3: .a = a: .b = b: .c = c: End With
End Function

Public Function Vec3_Sum(v As Vector3) As Double
    With v:     Vec3_Sum = .a + .b + .c: End With
End Function

Public Function Vec3_cross(v1 As Vector3, v2 As Vector3) As Vector3
    Dim m As Matrix3
    Mat3_Col(m, 0) = Vec3(1, 1, 1)
    Mat3_Col(m, 1) = v1
    Mat3_Col(m, 2) = v2
    Vec3_cross = Mat3_detV(m)
'    With Vec3_cross
'        .a = v1.b * v2.c - v1.c * v2.b
'        .b = v1.c * v2.a - v1.a * v2.c
'        .c = v1.a * v2.b - v1.b * v2.a
'    End With
End Function

'Aus einem Vektor Untervektoren Rauskopieren
Public Function Vec3_uvec(v As Vector3, ByVal ex As Long) As Vector2
    'Kopiert alle Elemente auﬂer ex in einen kleineren Vektor
    With Vec3_uvec
        Select Case ex
        Case 0: .a = v.b: .b = v.c
        Case 1: .a = v.a: .b = v.c
        Case 2: .a = v.a: .b = v.b
        End Select
    End With
End Function

Public Function Vec3_ToStr(v As Vector3, Optional bIsLineVec As Boolean = False) As String
    Vec3_ToStr = VectorFormat(Vec3_ToArr(v), 0, bIsLineVec)
End Function
Public Function Vec3_ToArr(v As Vector3) As Double()
    Dim da(0 To 2) As Double
    With v: da(0) = .a: da(1) = .b: da(2) = .c: End With
    Vec3_ToArr = da
End Function


Public Function Mat2(aa As Double, ab As Double, _
                     ba As Double, bb As Double) As Matrix2
    'Erzeugt eine 2x2 Matrix
    With Mat2: .aa = aa: .ab = ab
               .ba = ba: .bb = bb
    End With
End Function
Public Function Mat3(aa As Double, ab As Double, ac As Double, _
                     ba As Double, bb As Double, bc As Double, _
                     ca As Double, cb As Double, cc As Double) As Matrix3
    'Erzeugt eine 3x3 Matrix
    With Mat3: .aa = aa: .ab = ab: .ac = ac
               .ba = ba: .bb = bb: .bc = bc
               .ca = ca: .cb = cb: .cc = cc
    End With
End Function



Public Property Get Mat2_Row(m As Matrix2, ByVal index As Long) As Vector2
    With m
        Select Case index
        Case 0: Mat2_Row = Vec2(.aa, .ab)
        Case 1: Mat2_Row = Vec2(.ba, .bb)
        End Select
    End With
End Property
Public Property Let Mat2_Row(m As Matrix2, ByVal index As Long, v As Vector2)
    With m
        Select Case index
        Case 0: .aa = v.a: .ab = v.b
        Case 1: .ba = v.a: .bb = v.b
        End Select
    End With
End Property

Public Property Get Mat3_Row(m As Matrix3, ByVal index As Long) As Vector3
    With m
        Select Case index
        Case 0: Mat3_Row = Vec3(.aa, .ab, .ac)
        Case 1: Mat3_Row = Vec3(.ba, .bb, .bc)
        Case 2: Mat3_Row = Vec3(.ca, .cb, .cc)
        End Select
    End With
End Property
Public Property Let Mat3_Row(m As Matrix3, ByVal index As Long, v As Vector3)
    With m
        Select Case index
        Case 0: .aa = v.a: .ab = v.b: .ac = v.c
        Case 1: .ba = v.a: .bb = v.b: .bc = v.c
        Case 2: .ca = v.a: .cb = v.b: .cc = v.c
        End Select
    End With
End Property
Public Property Get Mat3_Col(m As Matrix3, ByVal index As Long) As Vector3
    With m
        Select Case index
        Case 0: Mat3_Col = Vec3(.aa, .ba, .ca)
        Case 1: Mat3_Col = Vec3(.ab, .bb, .cb)
        Case 2: Mat3_Col = Vec3(.ac, .bc, .cc)
        End Select
    End With
End Property
Public Property Let Mat3_Col(m As Matrix3, ByVal index As Long, v As Vector3)
    With m
        Select Case index
        Case 0: .aa = v.a: .ba = v.b: .ca = v.c
        Case 1: .ab = v.a: .bb = v.b: .cb = v.c
        Case 2: .ac = v.a: .bc = v.b: .cc = v.c
        End Select
    End With
End Property

Public Function Mat2_det(m As Matrix2) As Double
    'Berechnet die Determinante einer 2x2-Matrix
    With m
        Mat2_det = .aa * .bb - .ab * .ba
    End With
End Function
Public Function Mat3_umat(m As Matrix3, ByVal r_ex As Long, ByVal c_ex As Long) As Matrix2
    'Liefert aus einer 3x3-Matrix die 2x2-Untermatrix, indem die Zeile r_ex und die Spalte c_ex ausgelassen werden.
    Dim icex As Long
    If r_ex <> 0 Then Mat2_Row(Mat3_umat, icex) = Vec3_uvec(Mat3_Row(m, 0), c_ex): icex = icex + 1
    If r_ex <> 1 Then Mat2_Row(Mat3_umat, icex) = Vec3_uvec(Mat3_Row(m, 1), c_ex): icex = icex + 1
    If r_ex <> 2 Then Mat2_Row(Mat3_umat, icex) = Vec3_uvec(Mat3_Row(m, 2), c_ex): icex = icex + 1
End Function

Public Function Mat3_min(m As Matrix3, ByVal r_ex As Long, ByVal c_ex As Long) As Double
    'Berechnet die Minore einer 3x3-Matrix bei der die Zeile r_ex und die Spalte c_ex ausgelassen werden.
    Mat3_min = Mat2_det(Mat3_umat(m, r_ex, c_ex))
End Function

Public Function Mat3_Adj(m As Matrix3) As Matrix3
    Mat3_Adj = Mat3(Mat3_min(m, 0, 0), -Mat3_min(m, 1, 0), Mat3_min(m, 2, 0), _
                   -Mat3_min(m, 0, 1), Mat3_min(m, 1, 1), -Mat3_min(m, 2, 1), _
                    Mat3_min(m, 0, 2), -Mat3_min(m, 1, 2), Mat3_min(m, 2, 2))
End Function

'Public Function Mat3_det(m As Matrix3) As Double
'    'Berechnet die Determinante einer 3x3-Matrix
'    Dim d1 As Double, d2 As Double, d3 As Double
'    Dim d4 As Double, d5 As Double, d6 As Double
'
'    With m
'        d1 = .aa * .bb * .cc
'        d2 = .ab * .bc * .ca
'        d3 = .ac * .ba * .cb
'        d4 = -.ac * .bb * .ca
'        d5 = -.ab * .ba * .cc
'        d6 = -.aa * .bc * .cb
'    End With
'    Mat3_det = d1 + d2 + d3 + d4 + d5 + d6
'
''    With m
''        Mat3_det = .aa * .bb * .cc + .ab * .bc * .ca + .ac * .ba * .cb _
''                 - .ac * .bb * .ca - .ab * .ba * .cc - .aa * .bc * .cb
''    End With
'End Function

'Public Function Mat3_detVX(m As Matrix3) As Vector3
'    'Berechnet die Determinante einer 3x3-Matrix
'    Dim d1 As Double, d2 As Double, d3 As Double
'    Dim d4 As Double, d5 As Double, d6 As Double
'
'    With m
'        d1 = .aa * .bb * .cc
'        d2 = .ab * .bc * .ca
'        d3 = .ac * .ba * .cb
'        d4 = -.ac * .bb * .ca
'        d5 = -.ab * .ba * .cc
'        d6 = -.aa * .bc * .cb
'    End With
'    Debug.Print d1; d2; d3; d4; d5; d6
'    '-2  1 -6  4  1 -3
'
'    With Mat3_detVX
'        .a = d1 + d6
'        .b = d3 + d5
'        .c = d2 + d4
'    End With
'End Function

Public Function Mat3_detV(m As Matrix3) As Vector3
    'Berechnet die . . .
    'Determinante einer 3x3-Matrix als Ergebnisvektor des Kreuzprodukts???
    With Mat3_detV
        .a = m.aa * m.bb * m.cc - m.aa * m.bc * m.cb
        .b = m.ac * m.ba * m.cb - m.ab * m.ba * m.cc
        .c = m.ab * m.bc * m.ca - m.ac * m.bb * m.ca
    End With
End Function

Public Function Mat3_det(m As Matrix3) As Double
    'Berechnet die Determinante einer 3x3-Matrix
    Mat3_det = Vec3_Sum(Mat3_detV(m))
'    With Mat3_detV
'        .a = m.aa * m.bb * m.cc - m.aa * m.bc * m.cb
'        .b = m.ac * m.ba * m.cb - m.ab * m.ba * m.cc
'        .c = m.ab * m.bc * m.ca - m.ac * m.bb * m.ca
'    End With
End Function

Public Function VectorFormat(da() As Double, totalWidth As Long, Optional bIsLineVec As Boolean = False, Optional dFormat As Integer = -1) As String
    Dim i As Long, maxlenL As Long, maxlenR As Long
    Dim s As String, sa() As String
    Dim sdi As String, sdf As String
    Dim l As Long: l = LBound(da)
    Dim u As Long: u = UBound(da)
    For i = l To u
        If dFormat < 0 Then
            s = Trim(Str(da(i)))
        Else
            s = Replace(Format(da(i), "0." & String$(dFormat, "0")), ",", ".")
        End If
        If InStr(1, s, ".") Then
            sa = Split(s, ".")
            sdi = sa(0): If Len(sdi) = 0 Then sdi = "0"
            sdf = sa(1)
        Else
            sdi = s: sdf = ""
        End If
        maxlenL = Max(maxlenL, Len(sdi))
        maxlenR = Max(maxlenR, Len(sdf))
    Next
    ReDim sar(l To u) As String
    For i = l To u
        s = Trim(Str(da(i)))
        If InStr(1, s, ".") Then
            sa = Split(s, ".")
            sdi = sa(0): If Len(sdi) = 0 Then sdi = "0"
            sdf = sa(1)
            sar(i) = PadLeft(sdi, maxlenL) & "." & PadRight(sdf, maxlenR)
        Else
            sdi = s: sdf = ""
            sar(i) = PadLeft(sdi, maxlenL) & PadRight(sdf, maxlenR + 1)
        End If
    Next
    VectorFormat = Join(sar, IIf(bIsLineVec, " ", vbCrLf))
End Function

Public Function PadLeft(StrVal As String, _
                        ByVal totalWidth As Long, _
                        Optional ByVal paddingChar As String) As String
    ' der String wird mit der angegebenen L‰nge zur¸ckgegeben, der
    ' String wird nach rechts ger¸ckt, und links mit PadChar aufgef¸llt
    ' ist PadChar nicht angegeben, so wird mit RSet der String in
    ' Spaces eingef¸gt.
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
    ' der String wird mit der angegebenen L‰nge zur¸ckgegeben, der
    ' String wird nach links ger¸ckt, und rechts mit PadChar aufgef¸llt
    ' ist PadChar nicht angegeben, so wird mit LSet der String in
    ' Spaces eingef¸gt.
    If Len(paddingChar) Then
        If Len(StrVal) <= totalWidth Then _
            PadRight = StrVal & String$(totalWidth - Len(StrVal), paddingChar)
    Else
        PadRight = Space$(totalWidth)
        LSet PadRight = StrVal
    End If
End Function

Public Function Max(v1, v2)
    If v1 > v2 Then Max = v1 Else Max = v2
End Function


