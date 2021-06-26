Attribute VB_Name = "MMatrices"
Option Explicit
'2021-06-26 lines: 1218
'in VBC we are forced to use letters for the names of the variables of the elements in a vector or a matrix
'because otherwise we will run out of line space in some functions
Public Type Vector2
    a As Double
    b As Double
End Type
Public Type Vector3
    a As Double
    b As Double
    c As Double
End Type
Public Type Vector4
    a As Double
    b As Double
    c As Double
    d As Double
End Type
Public Type Vector5
    a As Double
    b As Double
    c As Double
    d As Double
    e As Double
End Type
Public Type Vector6
    a As Double
    b As Double
    c As Double
    d As Double
    e As Double
    f As Double
End Type
Public Type Matrix2 'Matrix 2x2
    aa As Double 'erste  Zeile, erste  Spalte
    ab As Double 'erste  Zeile, zweite Spalte
    ba As Double 'zweite Zeile, erste  Spalte
    bb As Double 'zweite Zeile, zweite Spalte
End Type
Public Type Matrix3 'Matrix 3x3
    aa As Double 'erste  Zeile, erste  Spalte
    ab As Double 'erste  Zeile, zweite Spalte
    ac As Double 'erste  Zeile, dritte Spalte
    ba As Double 'zweite Zeile, erste  Spalte
    bb As Double 'zweite Zeile, zweite Spalte
    bc As Double 'zweite Zeile, dritte Spalte
    ca As Double 'dritte Zeile, erste  Spalte
    cb As Double 'dritte Zeile, zweite Spalte
    cc As Double 'dritte Zeile, dritte Spalte
End Type
Public Type Matrix4 'Matrix 4x4
    aa As Double 'erste  Zeile, erste  Spalte
    ab As Double 'erste  Zeile, zweite Spalte
    ac As Double 'erste  Zeile, dritte Spalte
    ad As Double 'erste  Zeile, vierte Spalte
    ba As Double 'zweite Zeile, erste  Spalte
    bb As Double 'zweite Zeile, zweite Spalte
    bc As Double 'zweite Zeile, dritte Spalte
    bd As Double 'zweite Zeile, vierte Spalte
    ca As Double 'dritte Zeile, erste  Spalte
    cb As Double 'dritte Zeile, zweite Spalte
    cc As Double 'dritte Zeile, dritte Spalte
    cd As Double 'dritte Zeile, vierte Spalte
    da As Double 'vierte Zeile, erste  Spalte
    db As Double 'vierte Zeile, zweite Spalte
    dc As Double 'vierte Zeile, dritte Spalte
    dd As Double 'vierte Zeile, vierte Spalte
End Type
Public Type Matrix5 'Matrix 5x5
    aa As Double 'erste  Zeile, erste  Spalte
    ab As Double 'erste  Zeile, zweite Spalte
    ac As Double 'erste  Zeile, dritte Spalte
    ad As Double 'erste  Zeile, vierte Spalte
    ae As Double 'erste  Zeile, fünfte Spalte
    ba As Double 'zweite Zeile, erste  Spalte
    bb As Double 'zweite Zeile, zweite Spalte
    bc As Double 'zweite Zeile, dritte Spalte
    bd As Double 'zweite Zeile, vierte Spalte
    be As Double 'zweite Zeile, fünfte Spalte
    ca As Double 'dritte Zeile, erste  Spalte
    cb As Double 'dritte Zeile, zweite Spalte
    cc As Double 'dritte Zeile, dritte Spalte
    cd As Double 'dritte Zeile, vierte Spalte
    ce As Double 'dritte Zeile, fünfte Spalte
    da As Double 'vierte Zeile, erste  Spalte
    db As Double 'vierte Zeile, zweite Spalte
    dc As Double 'vierte Zeile, dritte Spalte
    dd As Double 'vierte Zeile, vierte Spalte
    de As Double 'vierte Zeile, fünfte Spalte
    ea As Double 'fünfte Zeile, erste  Spalte
    eb As Double 'fünfte Zeile, zweite Spalte
    ec As Double 'fünfte Zeile, dritte Spalte
    ed As Double 'fünfte Zeile, vierte Spalte
    ee As Double 'fünfte Zeile, fünfte Spalte
End Type
Public Type Matrix6 'Matrix 6x6
    aa As Double 'erste   Zeile, erste   Spalte
    ab As Double 'erste   Zeile, zweite  Spalte
    ac As Double 'erste   Zeile, dritte  Spalte
    ad As Double 'erste   Zeile, vierte  Spalte
    ae As Double 'erste   Zeile, fünfte  Spalte
    af As Double 'erste   Zeile, sechste Spalte
    ba As Double 'zweite  Zeile, erste   Spalte
    bb As Double 'zweite  Zeile, zweite  Spalte
    bc As Double 'zweite  Zeile, dritte  Spalte
    bd As Double 'zweite  Zeile, vierte  Spalte
    be As Double 'zweite  Zeile, fünfte  Spalte
    bf As Double 'zweite  Zeile, sechste Spalte
    ca As Double 'dritte  Zeile, erste   Spalte
    cb As Double 'dritte  Zeile, zweite  Spalte
    cc As Double 'dritte  Zeile, dritte  Spalte
    cd As Double 'dritte  Zeile, vierte  Spalte
    ce As Double 'dritte  Zeile, fünfte  Spalte
    cf As Double 'dritte  Zeile, sechste Spalte
    da As Double 'vierte  Zeile, erste   Spalte
    db As Double 'vierte  Zeile, zweite  Spalte
    dc As Double 'vierte  Zeile, dritte  Spalte
    dd As Double 'vierte  Zeile, vierte  Spalte
    de As Double 'vierte  Zeile, fünfte  Spalte
    df As Double 'vierte  Zeile, sechste Spalte
    ea As Double 'fünfte  Zeile, erste   Spalte
    eb As Double 'fünfte  Zeile, zweite  Spalte
    ec As Double 'fünfte  Zeile, dritte  Spalte
    ed As Double 'fünfte  Zeile, vierte  Spalte
    ee As Double 'fünfte  Zeile, fünfte  Spalte
    ef As Double 'fünfte  Zeile, sechste Spalte
    fa As Double 'sechste Zeile, erste   Spalte
    fb As Double 'sechste Zeile, zweite  Spalte
    fc As Double 'sechste Zeile, dritte  Spalte
    fd As Double 'sechste Zeile, vierte  Spalte
    fe As Double 'sechste Zeile, fünfte  Spalte
    ff As Double 'sechste Zeile, sechste Spalte
End Type
'Public Type MatrixT
'    a() As Double
'End Type
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef pDst As Any, ByRef pSrc As Any, ByVal nBytes As Long)
'Alias "cpymem"

'Vektoren erzeugen
Public Function Vec2(a As Double, b As Double) As Vector2
    'Erzeugt einen Vector mit 2 Elementen
    With Vec2: .a = a: .b = b: End With
End Function
Public Function Vec3(a As Double, b As Double, c As Double) As Vector3
    'Erzeugt einen Vector mit 3 Elementen
    With Vec3: .a = a: .b = b: .c = c: End With
End Function
Public Function Vec4(a As Double, b As Double, c As Double, d As Double) As Vector4
    'Erzeugt einen Vector mit 4 Elementen
    With Vec4: .a = a: .b = b: .c = c: .d = d: End With
End Function
Public Function Vec5(a As Double, b As Double, c As Double, d As Double, e As Double) As Vector5
    'Erzeugt einen Vector mit 5 Elementen
    With Vec5: .a = a: .b = b: .c = c: .d = d: .e = e: End With
End Function
Public Function Vec6(a As Double, b As Double, c As Double, d As Double, e As Double, f As Double) As Vector6
    'Erzeugt einen Vector mit 6 Elementen
    With Vec6: .a = a: .b = b: .c = c: .d = d: .e = e: .f = f: End With
End Function

'Alle Daten im Vektor löschen, bzw Variablen zu Null setzen
Public Function Vec2_Clear() As Vector2
    Dim v As Vector2: Vec2_Clear = v
End Function
Public Function Vec3_Clear() As Vector3
    Dim v As Vector3: Vec3_Clear = v
End Function
Public Function Vec4_Clear() As Vector4
    Dim v As Vector4: Vec4_Clear = v
End Function
Public Function Vec5_Clear() As Vector5
    Dim v As Vector5: Vec5_Clear = v
End Function
Public Function Vec6_Clear() As Vector6
    Dim v As Vector6: Vec6_Clear = v
End Function

'Vektoren-Addition
Public Function Vector2_add(v1 As Vector2, v2 As Vector2) As Vector2
    'Addiert 2 2er-Vectoren
    With Vector2_add:   .a = v1.a + v2.a:    .b = v1.b + v2.b:    End With
End Function
Public Function Vector3_add(v1 As Vector3, v2 As Vector3) As Vector3
    'Addiert 2 3er-Vectoren
    With Vector3_add:   .a = v1.a + v2.a:    .b = v1.b + v2.b:    .c = v1.c + v2.c:    End With
End Function
Public Function Vector4_add(v1 As Vector4, v2 As Vector4) As Vector4
    'Addiert 2 4er-Vectoren
    With Vector4_add:   .a = v1.a + v2.a:    .b = v1.b + v2.b:    .c = v1.c + v2.c:    .d = v1.d + v2.d:    End With
End Function
Public Function Vector5_add(v1 As Vector5, v2 As Vector5) As Vector5
    'Addiert 2 5er-Vectoren
    With Vector5_add:   .a = v1.a + v2.a:    .b = v1.b + v2.b:    .c = v1.c + v2.c:    .d = v1.d + v2.d:    .e = v1.e + v2.e:    End With
End Function
Public Function Vector6_add(v1 As Vector6, v2 As Vector6) As Vector6
    'Addiert 2 6er-Vectoren
    With Vector6_add:   .a = v1.a + v2.a:    .b = v1.b + v2.b:    .c = v1.c + v2.c:    .d = v1.d + v2.d:    .e = v1.e + v2.e:    .f = v1.f + v2.f:    End With
End Function

'Vektoren-Sutraktion
Public Function Vector2_sub(v1 As Vector2, v2 As Vector2) As Vector2
    'Subtrahiert 2 2er-Vectoren
    With Vector2_sub:   .a = v1.a + v2.a:    .b = v1.b + v2.b:    End With
End Function
Public Function Vector3_sub(v1 As Vector3, v2 As Vector3) As Vector3
    'Subtrahiert 2 3er-Vectoren
    With Vector3_sub:   .a = v1.a + v2.a:    .b = v1.b + v2.b:    .c = v1.c + v2.c:    End With
End Function
Public Function Vector4_sub(v1 As Vector4, v2 As Vector4) As Vector4
    'Subtrahiert 2 4er-Vectoren
    With Vector4_sub:   .a = v1.a + v2.a:    .b = v1.b + v2.b:    .c = v1.c + v2.c:    .d = v1.d + v2.d:    End With
End Function
Public Function Vector5_sub(v1 As Vector5, v2 As Vector5) As Vector5
    'Subtrahiert 2 5er-Vectoren
    With Vector5_sub:   .a = v1.a + v2.a:    .b = v1.b + v2.b:    .c = v1.c + v2.c:    .d = v1.d + v2.d:    .e = v1.e + v2.e:    End With
End Function
Public Function Vector6_sub(v1 As Vector6, v2 As Vector6) As Vector6
    'Subtrahiert 2 6er-Vectoren
    With Vector6_sub:   .a = v1.a + v2.a:    .b = v1.b + v2.b:    .c = v1.c + v2.c:    .d = v1.d + v2.d:    .e = v1.e + v2.e:    .f = v1.f + v2.f:    End With
End Function

'Vektoren mit Skalar multiplizieren
Public Function Vector2_smul(v As Vector2, ByVal s As Double) As Vector2
    'Multipliziert Alle Elemente eines 2er-Vectors mit einem Skalar
    With Vector2_smul: .a = .a * s: .b = .b * s: End With
End Function
Public Function Vector3_smul(v As Vector3, ByVal s As Double) As Vector3
    'Multipliziert Alle Elemente eines 3er-Vectors mit einem Skalar
    With Vector3_smul: .a = .a * s: .b = .b * s: .c = .c * s: End With
End Function
Public Function Vector4_smul(v As Vector4, ByVal s As Double) As Vector4
    'Multipliziert Alle Elemente eines 3er-Vectors mit einem Skalar
    With Vector4_smul: .a = .a * s: .b = .b * s: .c = .c * s: .d = .d * s: End With
End Function
Public Function Vector5_smul(v As Vector5, ByVal s As Double) As Vector5
    'Multipliziert Alle Elemente eines 3er-Vectors mit einem Skalar
    With Vector5_smul: .a = .a * s: .b = .b * s: .c = .c * s: .d = .d * s: .e = .e * s: End With
End Function
Public Function Vector6_smul(v As Vector6, ByVal s As Double) As Vector6
    'Multipliziert Alle Elemente eines 3er-Vectors mit einem Skalar
    With Vector6_smul: .a = .a * s: .b = .b * s: .c = .c * s: .d = .d * s: .e = .e * s: .f = .f * s: End With
End Function

'Aus einem Vektor Untervektoren Rauskopieren
Public Function Vector3_uvec(v As Vector3, ByVal ex As Long) As Vector2
    'Kopiert alle Elemente außer ex in einen kleineren Vektor
    With Vector3_uvec
        Select Case ex
        Case 0: .a = v.b: .b = v.c
        Case 1: .a = v.a: .b = v.c
        Case 2: .a = v.a: .b = v.b
        End Select
    End With
End Function
Public Function Vector4_uvec(v As Vector4, ByVal ex As Long) As Vector3
    'kopiert alle Elemente außer ex in einen kleineren Vektor
    With Vector4_uvec
        Select Case ex
        Case 0: .a = v.b: .b = v.c: .c = v.d
        Case 1: .a = v.a: .b = v.c: .c = v.d
        Case 2: .a = v.a: .b = v.b: .c = v.d
        Case 3: .a = v.a: .b = v.b: .c = v.c
        End Select
    End With
End Function
 Public Function Vector5_uvec(v As Vector5, ByVal ex As Long) As Vector4
    'Kopiert alle Elemente außer ex in einen kleineren Vektor
    With Vector5_uvec
        Select Case ex
        Case 0: .a = v.b: .b = v.c: .c = v.d: .d = v.e
        Case 1: .a = v.a: .b = v.c: .c = v.d: .d = v.e
        Case 2: .a = v.a: .b = v.b: .c = v.d: .d = v.e
        Case 3: .a = v.a: .b = v.b: .c = v.c: .d = v.e
        Case 4: .a = v.a: .b = v.b: .c = v.c: .d = v.d
        End Select
    End With
End Function
 Public Function Vector6_uvec(v As Vector6, ByVal ex As Long) As Vector5
    'Kopiert alle Elemente außer ex in einen kleineren Vektor
    With Vector6_uvec
        Select Case ex
        Case 0: .a = v.b: .b = v.c: .c = v.d: .d = v.e: .e = v.f
        Case 1: .a = v.a: .b = v.c: .c = v.d: .d = v.e: .e = v.f
        Case 2: .a = v.a: .b = v.b: .c = v.d: .d = v.e: .e = v.f
        Case 3: .a = v.a: .b = v.b: .c = v.c: .d = v.e: .e = v.f
        Case 4: .a = v.a: .b = v.b: .c = v.c: .d = v.d: .e = v.f
        Case 5: .a = v.a: .b = v.b: .c = v.c: .d = v.d: .e = v.e
        End Select
    End With
End Function

'Matrizen erzeugen
Public Function Mat2(aa As Double, ab As Double, _
                     ba As Double, bb As Double) As Matrix2
    'Erzeugt eine 2x2 Matrix
    With Mat2: .aa = aa: .ab = ab
               .ba = ba: .bb = bb:    End With
End Function
Public Function Mat3(aa As Double, ab As Double, ac As Double, _
                     ba As Double, bb As Double, bc As Double, _
                     ca As Double, cb As Double, cc As Double) As Matrix3
    'Erzeugt eine 3x3 Matrix
    With Mat3: .aa = aa: .ab = ab: .ac = ac
               .ba = ba: .bb = bb: .bc = bc
               .ca = ca: .cb = cb: .cc = cc:   End With
End Function
Public Function Mat4(aa As Double, ab As Double, ac As Double, ad As Double, _
                     ba As Double, bb As Double, bc As Double, bd As Double, _
                     ca As Double, cb As Double, cc As Double, cd As Double, _
                     da As Double, db As Double, dc As Double, dd As Double) As Matrix4
    'Erzeugt eine 4x4 Matrix
    With Mat4: .aa = aa: .ab = ab: .ac = ac: .ad = ad
               .ba = ba: .bb = bb: .bc = bc: .bd = bd
               .ca = ca: .cb = cb: .cc = cc: .cd = cd
               .da = da: .db = db: .dc = dc: .dd = dd:   End With
End Function
Public Function Mat5(aa As Double, ab As Double, ac As Double, ad As Double, ae As Double, _
                     ba As Double, bb As Double, bc As Double, bd As Double, be As Double, _
                     ca As Double, cb As Double, cc As Double, cd As Double, ce As Double, _
                     da As Double, db As Double, dc As Double, dd As Double, de As Double, _
                     ea As Double, eb As Double, ec As Double, ed As Double, ee As Double) As Matrix5
    'Erzeugt eine 5x5 Matrix
    With Mat5: .aa = aa: .ab = ab: .ac = ac: .ad = ad: .ae = ae
               .ba = ba: .bb = bb: .bc = bc: .bd = bd: .be = be
               .ca = ca: .cb = cb: .cc = cc: .cd = cd: .ce = ce
               .da = da: .db = db: .dc = dc: .dd = dd: .de = de
               .ea = ea: .eb = eb: .ec = ec: .ed = ed: .ee = ee:   End With
End Function
Public Function Mat6(aa As Double, ab As Double, ac As Double, ad As Double, ae As Double, af As Double, _
                     ba As Double, bb As Double, bc As Double, bd As Double, be As Double, bf As Double, _
                     ca As Double, cb As Double, cc As Double, cd As Double, ce As Double, cf As Double, _
                     da As Double, db As Double, dc As Double, dd As Double, de As Double, df As Double, _
                     ea As Double, eb As Double, ec As Double, ed As Double, ee As Double, ef As Double, _
                     fa As Double, fb As Double, fc As Double, fd As Double, fe As Double, ff As Double) As Matrix6
    'Erzeugt eine 6x6 Matrix
    With Mat6: .aa = aa: .ab = ab: .ac = ac: .ad = ad: .ae = ae: .af = af
               .ba = ba: .bb = bb: .bc = bc: .bd = bd: .be = be: .bf = bf
               .ca = ca: .cb = cb: .cc = cc: .cd = cd: .ce = ce: .cf = cf
               .da = da: .db = db: .dc = dc: .dd = dd: .de = de: .df = df
               .ea = ea: .eb = eb: .ec = ec: .ed = ed: .ee = ee: .ef = ef
               .fa = fa: .fb = fb: .fc = fc: .fd = fd: .fe = fe: .ff = ff:   End With
End Function

'Einheits-Matrizen erzeugen
Public Function Mat2_E() As Matrix2
    'Erzeugt eine 2x2 Einheits-Matrix
    With Mat2_E: .aa = 1: .bb = 1:    End With
End Function
Public Function Mat3_E() As Matrix3
    'Erzeugt eine 3x3 Einheits-Matrix
    With Mat3_E: .aa = 1: .bb = 1: .cc = 1:    End With
End Function
Public Function Mat4_E() As Matrix4
    'Erzeugt eine 4x4 Einheits-Matrix
    With Mat4_E: .aa = 1: .bb = 1: .cc = 1: .dd = 1:    End With
End Function
Public Function Mat5_E() As Matrix5
    'Erzeugt eine 5x5 Einheits-Matrix
    With Mat5_E: .aa = 1: .bb = 1: .cc = 1: .dd = 1: .ee = 1:    End With
End Function
Public Function Mat6_E() As Matrix6
    'Erzeugt eine 6x6 Einheits-Matrix
    With Mat6_E: .aa = 1: .bb = 1: .cc = 1: .dd = 1: .ee = 1: .ff = 1:    End With
End Function

'Alle Daten in der Matrix löschen, bzw Variablen zu Null setzen
Public Function Mat2_Clear() As Matrix2
    Dim m As Matrix2: Mat2_Clear = m
End Function
Public Function Mat3_Clear() As Matrix3
    Dim m As Matrix3: Mat3_Clear = m
End Function
Public Function Mat4_Clear() As Matrix4
    Dim m As Matrix4: Mat4_Clear = m
End Function
Public Function Mat5_Clear() As Matrix5
    Dim m As Matrix5: Mat5_Clear = m
End Function
Public Function Mat6_Clear() As Matrix6
    Dim m As Matrix6: Mat6_Clear = m
End Function

'Matrizen-Addition
Public Function Matrix2_add(m1 As Matrix2, m2 As Matrix2) As Matrix2
    'Addiert 2 2x2 Matrizen
    With Matrix2_add:   .aa = m1.aa + m2.aa:    .ab = m1.ab + m2.ab
                        .ba = m1.ba + m2.ba:    .bb = m1.bb + m2.bb:    End With
End Function
Public Function Matrix3_add(m1 As Matrix3, m2 As Matrix3) As Matrix3
    'Addiert 2 3x3 Matrizen
    With Matrix3_add:   .aa = m1.aa + m2.aa:    .ab = m1.ab + m2.ab:    .ac = m1.ac + m2.ac
                        .ba = m1.ba + m2.ba:    .bb = m1.bb + m2.bb:    .bc = m1.bc + m2.bc
                        .ca = m1.ca + m2.ca:    .cb = m1.cb + m2.cb:    .cc = m1.cc + m2.cc:    End With
End Function
Public Function Matrix4_add(m1 As Matrix4, m2 As Matrix4) As Matrix4
    'Addiert 2 4x4 Matrizen
    With Matrix4_add:   .aa = m1.aa + m2.aa:    .ab = m1.ab + m2.ab:    .ac = m1.ac + m2.ac:    .ad = m1.ad + m2.ad
                        .ba = m1.ba + m2.ba:    .bb = m1.bb + m2.bb:    .bc = m1.bc + m2.bc:    .bd = m1.bd + m2.bd
                        .ca = m1.ca + m2.ca:    .cb = m1.cb + m2.cb:    .cc = m1.cc + m2.cc:    .cd = m1.cd + m2.cd
                        .da = m1.da + m2.da:    .db = m1.db + m2.db:    .dc = m1.dc + m2.dc:    .dd = m1.dd + m2.dd:    End With
End Function
Public Function Matrix5_add(m1 As Matrix5, m2 As Matrix5) As Matrix5
    'Addiert 2 5x5 Matrizen
    With Matrix5_add:   .aa = m1.aa + m2.aa:    .ab = m1.ab + m2.ab:    .ac = m1.ac + m2.ac:    .ad = m1.ad + m2.ad:    .ae = m1.ae + m2.ae
                        .ba = m1.ba + m2.ba:    .bb = m1.bb + m2.bb:    .bc = m1.bc + m2.bc:    .bd = m1.bd + m2.bd:    .be = m1.be + m2.be
                        .ca = m1.ca + m2.ca:    .cb = m1.cb + m2.cb:    .cc = m1.cc + m2.cc:    .cd = m1.cd + m2.cd:    .ce = m1.ce + m2.ce
                        .da = m1.da + m2.da:    .db = m1.db + m2.db:    .dc = m1.dc + m2.dc:    .dd = m1.dd + m2.dd:    .de = m1.de + m2.de
                        .ea = m1.ea + m2.ea:    .eb = m1.eb + m2.eb:    .ec = m1.ec + m2.ec:    .ed = m1.ed + m2.ed:    .ee = m1.ee + m2.ee:    End With
End Function
Public Function Matrix6_add(m1 As Matrix6, m2 As Matrix6) As Matrix6
    'Addiert 2 6x6 Matrizen
    With Matrix6_add:   .aa = m1.aa + m2.aa:    .ab = m1.ab + m2.ab:    .ac = m1.ac + m2.ac:    .ad = m1.ad + m2.ad:    .ae = m1.ae + m2.ae:    .af = m1.af + m2.af
                        .ba = m1.ba + m2.ba:    .bb = m1.bb + m2.bb:    .bc = m1.bc + m2.bc:    .bd = m1.bd + m2.bd:    .be = m1.be + m2.be:    .bf = m1.bf + m2.bf
                        .ca = m1.ca + m2.ca:    .cb = m1.cb + m2.cb:    .cc = m1.cc + m2.cc:    .cd = m1.cd + m2.cd:    .ce = m1.ce + m2.ce:    .cf = m1.cf + m2.cf
                        .da = m1.da + m2.da:    .db = m1.db + m2.db:    .dc = m1.dc + m2.dc:    .dd = m1.dd + m2.dd:    .de = m1.de + m2.de:    .df = m1.df + m2.df
                        .ea = m1.ea + m2.ea:    .eb = m1.eb + m2.eb:    .ec = m1.ec + m2.ec:    .ed = m1.ed + m2.ed:    .ee = m1.ee + m2.ee:    .ef = m1.ef + m2.ef
                        .fa = m1.fa + m2.fa:    .fb = m1.fb + m2.fb:    .fc = m1.fc + m2.fc:    .fd = m1.fd + m2.fd:    .fe = m1.fe + m2.fe:    .ff = m1.ff + m2.ff:    End With
End Function

'Matrizen-Subraktion
Public Function Matrix2_sub(m1 As Matrix2, m2 As Matrix2) As Matrix2
    'Subtrahiert eine 2x2 Matrix m2 von Matrix m1
    With Matrix2_sub:   .aa = m1.aa - m2.aa:    .ab = m1.ab - m2.ab
                        .ba = m1.ba - m2.ba:    .bb = m1.bb - m2.bb:    End With
End Function
Public Function Matrix3_sub(m1 As Matrix3, m2 As Matrix3) As Matrix3
    'Subtrahiert eine 3x3 Matrix m2 von Matrix m1
    With Matrix3_sub:   .aa = m1.aa - m2.aa:    .ab = m1.ab - m2.ab:    .ac = m1.ac - m2.ac
                        .ba = m1.ba - m2.ba:    .bb = m1.bb - m2.bb:    .bc = m1.bc - m2.bc
                        .ca = m1.ca - m2.ca:    .cb = m1.cb - m2.cb:    .cc = m1.cc - m2.cc:    End With
End Function
Public Function Matrix4_sub(m1 As Matrix4, m2 As Matrix4) As Matrix4
    'Subtrahiert eine 4x4 Matrix m2 von Matrix m1
    With Matrix4_sub:   .aa = m1.aa - m2.aa:    .ab = m1.ab - m2.ab:    .ac = m1.ac - m2.ac:    .ad = m1.ad - m2.ad
                        .ba = m1.ba - m2.ba:    .bb = m1.bb - m2.bb:    .bc = m1.bc - m2.bc:    .bd = m1.bd - m2.bd
                        .ca = m1.ca - m2.ca:    .cb = m1.cb - m2.cb:    .cc = m1.cc - m2.cc:    .cd = m1.cd - m2.cd
                        .da = m1.da - m2.da:    .db = m1.db - m2.db:    .dc = m1.dc - m2.dc:    .dd = m1.dd - m2.dd:    End With
End Function
Public Function Matrix5_sub(m1 As Matrix5, m2 As Matrix5) As Matrix5
    'Subtrahiert eine 5x5 Matrix m2 von Matrix m1
    With Matrix5_sub:   .aa = m1.aa - m2.aa:    .ab = m1.ab - m2.ab:    .ac = m1.ac - m2.ac:    .ad = m1.ad - m2.ad:    .ae = m1.ae - m2.ae
                        .ba = m1.ba - m2.ba:    .bb = m1.bb - m2.bb:    .bc = m1.bc - m2.bc:    .bd = m1.bd - m2.bd:    .be = m1.be - m2.be
                        .ca = m1.ca - m2.ca:    .cb = m1.cb - m2.cb:    .cc = m1.cc - m2.cc:    .cd = m1.cd - m2.cd:    .ce = m1.ce - m2.ce
                        .da = m1.da - m2.da:    .db = m1.db - m2.db:    .dc = m1.dc - m2.dc:    .dd = m1.dd - m2.dd:    .de = m1.de - m2.de
                        .ea = m1.ea - m2.ea:    .eb = m1.eb - m2.eb:    .ec = m1.ec - m2.ec:    .ed = m1.ed - m2.ed:    .ee = m1.ee - m2.ee:    End With
End Function
Public Function Matrix6_sub(m1 As Matrix6, m2 As Matrix6) As Matrix6
    'Subtrahiert eine 6x6 Matrix m2 von Matrix m1
    With Matrix6_sub:   .aa = m1.aa - m2.aa:    .ab = m1.ab - m2.ab:    .ac = m1.ac - m2.ac:    .ad = m1.ad - m2.ad:    .ae = m1.ae - m2.ae:    .af = m1.af - m2.af
                        .ba = m1.ba - m2.ba:    .bb = m1.bb - m2.bb:    .bc = m1.bc - m2.bc:    .bd = m1.bd - m2.bd:    .be = m1.be - m2.be:    .bf = m1.bf - m2.bf
                        .ca = m1.ca - m2.ca:    .cb = m1.cb - m2.cb:    .cc = m1.cc - m2.cc:    .cd = m1.cd - m2.cd:    .ce = m1.ce - m2.ce:    .cf = m1.cf - m2.cf
                        .da = m1.da - m2.da:    .db = m1.db - m2.db:    .dc = m1.dc - m2.dc:    .dd = m1.dd - m2.dd:    .de = m1.de - m2.de:    .df = m1.df - m2.df
                        .ea = m1.ea - m2.ea:    .eb = m1.eb - m2.eb:    .ec = m1.ec - m2.ec:    .ed = m1.ed - m2.ed:    .ee = m1.ee - m2.ee:    .ef = m1.ef - m2.ef
                        .fa = m1.fa - m2.fa:    .fb = m1.fb - m2.fb:    .fc = m1.fc - m2.fc:    .fd = m1.fd - m2.fd:    .fe = m1.fe - m2.fe:    .ff = m1.ff - m2.ff:    End With
End Function

'Matrizen-Skalar-Multiplikation
Public Function Matrix2_smul(m As Matrix2, s As Double) As Matrix2
    'Multipliziert eine 2x2 Matrix mit einem Skalar
    With Matrix2_smul:  .aa = m.aa * s: .ab = m.ab * s
                        .ba = m.ba * s: .bb = m.bb * s:    End With
End Function
Public Function Matrix3_smul(m As Matrix3, s As Double) As Matrix3
    'Multipliziert eine 3x3 Matrix mit einem Skalar
    With Matrix3_smul:  .aa = m.aa * s: .ab = m.ab * s: .ac = m.ac * s
                        .ba = m.ba * s: .bb = m.bb * s: .bc = m.bc * s
                        .ca = m.ca * s: .cb = m.cb * s: .cc = m.cc * s:    End With
End Function
Public Function Matrix4_smul(m As Matrix4, s As Double) As Matrix4
    'Multipliziert eine 4x4 Matrix mit einem Skalar
    With Matrix4_smul:  .aa = m.aa * s: .ab = m.ab * s: .ac = m.ac * s: .ad = m.ad * s
                        .ba = m.ba * s: .bb = m.bb * s: .bc = m.bc * s: .bd = m.bd * s
                        .ca = m.ca * s: .cb = m.cb * s: .cc = m.cc * s: .cd = m.cd * s
                        .da = m.da * s: .db = m.db * s: .dc = m.dc * s: .dd = m.dd * s:    End With
End Function
Public Function Matrix5_smul(m As Matrix5, s As Double) As Matrix5
    'Multipliziert eine 5x5 Matrix mit einem Skalar
    With Matrix5_smul:  .aa = m.aa * s: .ab = m.ab * s: .ac = m.ac * s: .ad = m.ad * s: .ae = m.ae * s
                        .ba = m.ba * s: .bb = m.bb * s: .bc = m.bc * s: .bd = m.bd * s: .be = m.be * s
                        .ca = m.ca * s: .cb = m.cb * s: .cc = m.cc * s: .cd = m.cd * s: .ce = m.ce * s
                        .da = m.da * s: .db = m.db * s: .dc = m.dc * s: .dd = m.dd * s: .de = m.de * s
                        .ea = m.ea * s: .eb = m.eb * s: .ec = m.ec * s: .ed = m.ed * s: .ee = m.ee * s:    End With
End Function
Public Function Matrix6_smul(m As Matrix6, s As Double) As Matrix6
    'Multipliziert eine 6x6 Matrix mit einem Skalar
    With Matrix6_smul:  .aa = m.aa * s: .ab = m.ab * s: .ac = m.ac * s: .ad = m.ad * s: .ae = m.ae * s: .af = m.af * s
                        .ba = m.ba * s: .bb = m.bb * s: .bc = m.bc * s: .bd = m.bd * s: .be = m.be * s: .bf = m.bf * s
                        .ca = m.ca * s: .cb = m.cb * s: .cc = m.cc * s: .cd = m.cd * s: .ce = m.ce * s: .cf = m.cf * s
                        .da = m.da * s: .db = m.db * s: .dc = m.dc * s: .dd = m.dd * s: .de = m.de * s: .df = m.df * s
                        .ea = m.ea * s: .eb = m.eb * s: .ec = m.ec * s: .ed = m.ed * s: .ee = m.ee * s: .ef = m.ef * s
                        .fa = m.fa * s: .fb = m.fb * s: .fc = m.fc * s: .fd = m.fd * s: .fe = m.fe * s: .ff = m.ff * s:    End With
End Function

'Matrizen-Multiplikation
Public Function Matrix2_mul(m1 As Matrix2, m2 As Matrix2) As Matrix2
    'Multipliziert eine 2x2 Matrix m1 mit einer Matrix m2
    With Matrix2_mul:   .aa = m1.aa * m2.aa + m1.ab * m2.ba:    .ab = m1.aa * m2.ab + m1.ab * m2.bb
                        .ba = m1.ba * m2.aa + m1.bb * m2.ba:    .bb = m1.ba * m2.ab + m1.bb * m2.bb:    End With
End Function

Public Function Matrix3_mul(m1 As Matrix3, m2 As Matrix3) As Matrix3
    'Multipliziert eine 3x3 Matrix m1 mit einer Matrix m2
    With Matrix3_mul:   .aa = m1.aa * m2.aa + m1.ab * m2.ba + m1.ac * m2.ca:    .ab = m1.aa * m2.ab + m1.ab * m2.bb + m1.ac * m2.cb:    .ac = m1.aa * m2.ac + m1.ab * m2.bc + m1.ac * m2.cc
                        .ba = m1.ba * m2.aa + m1.bb * m2.ba + m1.bc * m2.ca:    .bb = m1.ba * m2.ab + m1.bb * m2.bb + m1.bc * m2.cb:    .bc = m1.ba * m2.ac + m1.bb * m2.bc + m1.bc * m2.cc
                        .ca = m1.ca * m2.aa + m1.cb * m2.ba + m1.cc * m2.ca:    .cb = m1.ca * m2.ab + m1.cb * m2.bb + m1.cc * m2.cb:    .cc = m1.ca * m2.ac + m1.cb * m2.bc + m1.cc * m2.cc:    End With
End Function
Public Function Matrix4_mul(m1 As Matrix4, m2 As Matrix4) As Matrix4
    'Multipliziert eine 4x4 Matrix m1 mit einer Matrix m2
    With Matrix4_mul:   .aa = m1.aa * m2.aa + m1.ab * m2.ba + m1.ac * m2.ca + m1.ad * m2.da:    .ab = m1.aa * m2.ab + m1.ab * m2.bb + m1.ac * m2.cb + m1.ad * m2.db:    .ac = m1.aa * m2.ac + m1.ab * m2.bc + m1.ac * m2.cc + m1.ad * m2.dc:    .ad = m1.aa * m2.ad + m1.ab * m2.bd + m1.ac * m2.cd + m1.ad * m2.dd
                        .ba = m1.ba * m2.aa + m1.bb * m2.ba + m1.bc * m2.ca + m1.bd * m2.da:    .bb = m1.ba * m2.ab + m1.bb * m2.bb + m1.bc * m2.cb + m1.bd * m2.db:    .bc = m1.ba * m2.ac + m1.bb * m2.bc + m1.bc * m2.cc + m1.bd * m2.dc:    .bd = m1.ba * m2.ad + m1.bb * m2.bd + m1.bc * m2.cd + m1.bd * m2.dd
                        .ca = m1.ca * m2.aa + m1.cb * m2.ba + m1.cc * m2.ca + m1.cd * m2.da:    .cb = m1.ca * m2.ab + m1.cb * m2.bb + m1.cc * m2.cb + m1.cd * m2.db:    .cc = m1.ca * m2.ac + m1.cb * m2.bc + m1.cc * m2.cc + m1.cd * m2.dc:    .cd = m1.ca * m2.ad + m1.cb * m2.bd + m1.cc * m2.cd + m1.cd * m2.dd
                        .da = m1.da * m2.aa + m1.db * m2.ba + m1.dc * m2.ca + m1.dd * m2.da:    .db = m1.da * m2.ab + m1.db * m2.bb + m1.dc * m2.cb + m1.dd * m2.db:    .dc = m1.da * m2.ac + m1.db * m2.bc + m1.dc * m2.cc + m1.dd * m2.dc:    .dd = m1.da * m2.ad + m1.db * m2.bd + m1.dc * m2.cd + m1.dd * m2.dd:    End With
End Function
Public Function Matrix5_mul(m1 As Matrix5, m2 As Matrix5) As Matrix5
    'Multipliziert eine 5x5 Matrix m1 mit einer Matrix m2
    With Matrix5_mul:   .aa = m1.aa * m2.aa + m1.ab * m2.ba + m1.ac * m2.ca + m1.ad * m2.da + m1.ae * m2.ea:    .ab = m1.aa * m2.ab + m1.ab * m2.bb + m1.ac * m2.cb + m1.ad * m2.db + m1.ad * m2.eb:    .ac = m1.aa * m2.ac + m1.ab * m2.bc + m1.ac * m2.cc + m1.ad * m2.dc + m1.ae * m2.ec:    .ad = m1.aa * m2.ad + m1.ab * m2.bd + m1.ac * m2.cd + m1.ad * m2.dd + m1.ae * m2.ed:    .ae = m1.aa * m2.ae + m1.ab * m2.be + m1.ac * m2.ce + m1.ad * m2.de + m1.ae * m2.ee
                        .ba = m1.ba * m2.aa + m1.bb * m2.ba + m1.bc * m2.ca + m1.bd * m2.da + m1.be * m2.ea:    .bb = m1.ba * m2.ab + m1.bb * m2.bb + m1.bc * m2.cb + m1.bd * m2.db + m1.bd * m2.eb:    .bc = m1.ba * m2.ac + m1.bb * m2.bc + m1.bc * m2.cc + m1.bd * m2.dc + m1.be * m2.ec:    .bd = m1.ba * m2.ad + m1.bb * m2.bd + m1.bc * m2.cd + m1.bd * m2.dd + m1.be * m2.ed:    .be = m1.ba * m2.ae + m1.bb * m2.be + m1.bc * m2.ce + m1.bd * m2.de + m1.be * m2.ee
                        .ca = m1.ca * m2.aa + m1.cb * m2.ba + m1.cc * m2.ca + m1.cd * m2.da + m1.ce * m2.ea:    .cb = m1.ca * m2.ab + m1.cb * m2.bb + m1.cc * m2.cb + m1.cd * m2.db + m1.cd * m2.eb:    .cc = m1.ca * m2.ac + m1.cb * m2.bc + m1.cc * m2.cc + m1.cd * m2.dc + m1.ce * m2.ec:    .cd = m1.ca * m2.ad + m1.cb * m2.bd + m1.cc * m2.cd + m1.cd * m2.dd + m1.ce * m2.ed:    .ce = m1.ca * m2.ae + m1.cb * m2.be + m1.cc * m2.ce + m1.cd * m2.de + m1.ce * m2.ee
                        .da = m1.da * m2.aa + m1.db * m2.ba + m1.dc * m2.ca + m1.dd * m2.da + m1.de * m2.ea:    .db = m1.da * m2.ab + m1.db * m2.bb + m1.dc * m2.cb + m1.dd * m2.db + m1.dd * m2.eb:    .dc = m1.da * m2.ac + m1.db * m2.bc + m1.dc * m2.cc + m1.dd * m2.dc + m1.de * m2.ec:    .dd = m1.da * m2.ad + m1.db * m2.bd + m1.dc * m2.cd + m1.dd * m2.dd + m1.de * m2.ed:    .de = m1.da * m2.ae + m1.db * m2.be + m1.dc * m2.ce + m1.dd * m2.de + m1.de * m2.ee
                        .ea = m1.ea * m2.aa + m1.eb * m2.ba + m1.ec * m2.ca + m1.ed * m2.da + m1.ee * m2.ea:    .eb = m1.ea * m2.ab + m1.eb * m2.bb + m1.ec * m2.cb + m1.ed * m2.db + m1.ed * m2.eb:    .ec = m1.ea * m2.ac + m1.eb * m2.bc + m1.ec * m2.cc + m1.ed * m2.dc + m1.ee * m2.ec:    .ed = m1.ea * m2.ad + m1.eb * m2.bd + m1.ec * m2.cd + m1.ed * m2.dd + m1.ee * m2.ed:    .ee = m1.ea * m2.ae + m1.eb * m2.be + m1.ec * m2.ce + m1.ed * m2.de + m1.ee * m2.ee:    End With
End Function
Public Function Matrix6_mul(m1 As Matrix6, m2 As Matrix6) As Matrix6
    'Multipliziert eine 6x6 Matrix m1 mit einer Matrix m2
    With Matrix6_mul:   .aa = m1.aa * m2.aa + m1.ab * m2.ba + m1.ac * m2.ca + m1.ad * m2.da + m1.ae * m2.ea + m1.af * m2.fa:    .ab = m1.aa * m2.ab + m1.ab * m2.bb + m1.ac * m2.cb + m1.ad * m2.db + m1.ad * m2.eb + m1.af * m2.fb:    .ac = m1.aa * m2.ac + m1.ab * m2.bc + m1.ac * m2.cc + m1.ad * m2.dc + m1.ae * m2.ec + m1.af * m2.fc:    .ad = m1.aa * m2.ad + m1.ab * m2.bd + m1.ac * m2.cd + m1.ad * m2.dd + m1.ae * m2.ed + m1.af * m2.fd:    .ae = m1.aa * m2.ae + m1.ab * m2.be + m1.ac * m2.ce + m1.ad * m2.de + m1.ae * m2.ee + m1.af * m2.fe:    .af = m1.aa * m2.af + m1.ab * m2.bf + m1.ac * m2.cf + m1.ad * m2.df + m1.ae * m2.ef + m1.af * m2.ff
                        .ba = m1.ba * m2.aa + m1.bb * m2.ba + m1.bc * m2.ca + m1.bd * m2.da + m1.be * m2.ea + m1.bf * m2.fa:    .bb = m1.ba * m2.ab + m1.bb * m2.bb + m1.bc * m2.cb + m1.bd * m2.db + m1.bd * m2.eb + m1.bf * m2.fb:    .bc = m1.ba * m2.ac + m1.bb * m2.bc + m1.bc * m2.cc + m1.bd * m2.dc + m1.be * m2.ec + m1.bf * m2.fc:    .bd = m1.ba * m2.ad + m1.bb * m2.bd + m1.bc * m2.cd + m1.bd * m2.dd + m1.be * m2.ed + m1.bf * m2.fd:    .be = m1.ba * m2.ae + m1.bb * m2.be + m1.bc * m2.ce + m1.bd * m2.de + m1.be * m2.ee + m1.bf * m2.fe:    .bf = m1.ba * m2.af + m1.bb * m2.bf + m1.bc * m2.cf + m1.bd * m2.df + m1.be * m2.ef + m1.bf * m2.ff
                        .ca = m1.ca * m2.aa + m1.cb * m2.ba + m1.cc * m2.ca + m1.cd * m2.da + m1.ce * m2.ea + m1.cf * m2.fa:    .cb = m1.ca * m2.ab + m1.cb * m2.bb + m1.cc * m2.cb + m1.cd * m2.db + m1.cd * m2.eb + m1.cf * m2.fb:    .cc = m1.ca * m2.ac + m1.cb * m2.bc + m1.cc * m2.cc + m1.cd * m2.dc + m1.ce * m2.ec + m1.cf * m2.fc:    .cd = m1.ca * m2.ad + m1.cb * m2.bd + m1.cc * m2.cd + m1.cd * m2.dd + m1.ce * m2.ed + m1.cf * m2.fd:    .ce = m1.ca * m2.ae + m1.cb * m2.be + m1.cc * m2.ce + m1.cd * m2.de + m1.ce * m2.ee + m1.cf * m2.fe:    .cf = m1.ca * m2.af + m1.cb * m2.bf + m1.cc * m2.cf + m1.cd * m2.df + m1.ce * m2.ef + m1.cf * m2.ff
                        .da = m1.da * m2.aa + m1.db * m2.ba + m1.dc * m2.ca + m1.dd * m2.da + m1.de * m2.ea + m1.df * m2.fa:    .db = m1.da * m2.ab + m1.db * m2.bb + m1.dc * m2.cb + m1.dd * m2.db + m1.dd * m2.eb + m1.df * m2.fb:    .dc = m1.da * m2.ac + m1.db * m2.bc + m1.dc * m2.cc + m1.dd * m2.dc + m1.de * m2.ec + m1.df * m2.fc:    .dd = m1.da * m2.ad + m1.db * m2.bd + m1.dc * m2.cd + m1.dd * m2.dd + m1.de * m2.ed + m1.df * m2.fd:    .de = m1.da * m2.ae + m1.db * m2.be + m1.dc * m2.ce + m1.dd * m2.de + m1.de * m2.ee + m1.df * m2.fe:    .df = m1.da * m2.af + m1.db * m2.bf + m1.dc * m2.cf + m1.dd * m2.df + m1.de * m2.ef + m1.df * m2.ff
                        .ea = m1.ea * m2.aa + m1.eb * m2.ba + m1.ec * m2.ca + m1.ed * m2.da + m1.ee * m2.ea + m1.ef * m2.fa:    .eb = m1.ea * m2.ab + m1.eb * m2.bb + m1.ec * m2.cb + m1.ed * m2.db + m1.ed * m2.eb + m1.ef * m2.fb:    .ec = m1.ea * m2.ac + m1.eb * m2.bc + m1.ec * m2.cc + m1.ed * m2.dc + m1.ee * m2.ec + m1.ef * m2.fc:    .ed = m1.ea * m2.ad + m1.eb * m2.bd + m1.ec * m2.cd + m1.ed * m2.dd + m1.ee * m2.ed + m1.ef * m2.fd:    .ee = m1.ea * m2.ae + m1.eb * m2.be + m1.ec * m2.ce + m1.ed * m2.de + m1.ee * m2.ee + m1.ef * m2.fe:    .ef = m1.ea * m2.af + m1.eb * m2.bf + m1.ec * m2.cf + m1.ed * m2.df + m1.ee * m2.ef + m1.ef * m2.ff
                        .fa = m1.ea * m2.aa + m1.eb * m2.ba + m1.ec * m2.ca + m1.ed * m2.da + m1.ee * m2.ea + m1.ff * m2.fa:    .fb = m1.ea * m2.ab + m1.eb * m2.bb + m1.ec * m2.cb + m1.ed * m2.db + m1.ed * m2.eb + m1.ff * m2.fb:    .fc = m1.ea * m2.ac + m1.eb * m2.bc + m1.ec * m2.cc + m1.ed * m2.dc + m1.ee * m2.ec + m1.ff * m2.fc:    .fd = m1.fa * m2.ad + m1.fb * m2.bd + m1.fc * m2.cd + m1.fd * m2.dd + m1.fe * m2.ed + m1.ff * m2.fd:    .fe = m1.fa * m2.ae + m1.fb * m2.be + m1.fc * m2.ce + m1.fd * m2.de + m1.fe * m2.ee + m1.ff * m2.fe:    .ff = m1.fa * m2.af + m1.fb * m2.bf + m1.fc * m2.cf + m1.fd * m2.df + m1.fe * m2.ef + m1.ff * m2.ff:    End With
End Function

'Multiplikation Matrix mit Vektor
Public Function Matrix2_vmul(m As Matrix2, v As Vector2) As Vector2
    'Multipliziert eine 2x2 Matrix m mit einem 2er-Vektor
    With Matrix2_vmul:   .a = m.aa * v.a + m.ab * v.b
                         .b = m.ba * v.a + m.bb * v.b:    End With
End Function
Public Function Matrix3_vmul(m As Matrix3, v As Vector3) As Vector3
    'Multipliziert eine 3x3 Matrix m mit einem 3er-Vektor
    With Matrix3_vmul:   .a = m.aa * v.a + m.ab * v.b + m.ac * v.c
                         .b = m.ba * v.a + m.bb * v.b + m.bc * v.c
                         .c = m.ca * v.a + m.cb * v.b + m.cc * v.c:    End With
End Function
Public Function Matrix4_vmul(m As Matrix4, v As Vector4) As Vector4
    'Multipliziert eine 3x3 Matrix m mit einem 3er-Vektor
    With Matrix4_vmul:   .a = m.aa * v.a + m.ab * v.b + m.ac * v.c + m.ad * v.d
                         .b = m.ba * v.a + m.bb * v.b + m.bc * v.c + m.bd * v.d
                         .c = m.ca * v.a + m.cb * v.b + m.cc * v.c + m.cd * v.d
                         .d = m.da * v.a + m.db * v.b + m.dc * v.c + m.dd * v.d:    End With
End Function
Public Function Matrix5_vmul(m As Matrix5, v As Vector5) As Vector5
    'Multipliziert eine 3x3 Matrix m mit einem 3er-Vektor
    With Matrix5_vmul:   .a = m.aa * v.a + m.ab * v.b + m.ac * v.c + m.ad * v.d + m.ae * v.e
                         .b = m.ba * v.a + m.bb * v.b + m.bc * v.c + m.bd * v.d + m.be * v.e
                         .c = m.ca * v.a + m.cb * v.b + m.cc * v.c + m.cd * v.d + m.ce * v.e
                         .d = m.da * v.a + m.db * v.b + m.dc * v.c + m.dd * v.d + m.de * v.e
                         .e = m.ea * v.a + m.eb * v.b + m.ec * v.c + m.ed * v.d + m.ee * v.e:    End With
End Function
Public Function Matrix6_vmul(m As Matrix6, v As Vector6) As Vector6
    'Multipliziert eine 3x3 Matrix m mit einem 3er-Vektor
    With Matrix6_vmul:   .a = m.aa * v.a + m.ab * v.b + m.ac * v.c + m.ad * v.d + m.ae * v.e + m.af * v.f
                         .b = m.ba * v.a + m.bb * v.b + m.bc * v.c + m.bd * v.d + m.be * v.e + m.bf * v.f
                         .c = m.ca * v.a + m.cb * v.b + m.cc * v.c + m.cd * v.d + m.ce * v.e + m.cf * v.f
                         .d = m.da * v.a + m.db * v.b + m.dc * v.c + m.dd * v.d + m.de * v.e + m.df * v.f
                         .e = m.ea * v.a + m.eb * v.b + m.ec * v.c + m.ed * v.d + m.ee * v.e + m.ef * v.f
                         .f = m.fa * v.a + m.fb * v.b + m.fc * v.c + m.fd * v.d + m.fe * v.e + m.ff * v.f:    End With
End Function

'Transponierte Matrix
Public Function Matrix2_tra(m As Matrix2) As Matrix2
    'Erzeugt die Transponierte aus einer 2x2-Matrix
    With Matrix2_tra:   .aa = m.aa:    .ab = m.ba
                        .ba = m.ab:    .bb = m.bb:    End With
End Function
Public Function Matrix3_tra(m As Matrix3) As Matrix3
    'Erzeugt die Transponierte aus einer 3x3-Matrix
    With Matrix3_tra:   .aa = m.aa:    .ab = m.ba:    .ac = m.ca
                        .ba = m.ab:    .bb = m.bb:    .bc = m.cb
                        .ca = m.ac:    .cb = m.bc:    .cc = m.cc:    End With
End Function
Public Function Matrix4_tra(m As Matrix4) As Matrix4
    'Erzeugt die Transponierte aus einer 4x4-Matrix
    With Matrix4_tra:   .aa = m.aa:    .ab = m.ba:    .ac = m.ca:    .ad = m.da
                        .ba = m.ab:    .bb = m.bb:    .bc = m.cb:    .bd = m.db
                        .ca = m.ac:    .cb = m.bc:    .cc = m.cc:    .cd = m.dc
                        .da = m.ad:    .db = m.bd:    .dc = m.cd:    .dd = m.dd:    End With
End Function
Public Function Matrix5_tra(m As Matrix5) As Matrix5
    'Erzeugt die Transponierte aus einer 5x5-Matrix
    With Matrix5_tra:   .aa = m.aa:    .ab = m.ba:    .ac = m.ca:    .ad = m.da:    .ae = m.ea
                        .ba = m.ab:    .bb = m.bb:    .bc = m.cb:    .bd = m.db:    .be = m.eb
                        .ca = m.ac:    .cb = m.bc:    .cc = m.cc:    .cd = m.dc:    .ce = m.ec
                        .da = m.ad:    .db = m.bd:    .dc = m.cd:    .dd = m.dd:    .de = m.ed
                        .ea = m.ae:    .eb = m.be:    .ec = m.ce:    .ed = m.de:    .ee = m.ee:    End With
End Function
Public Function Matrix6_tra(m As Matrix6) As Matrix6
    'Erzeugt die Transponierte aus einer 6x6-Matrix
    With Matrix6_tra:   .aa = m.aa:    .ab = m.ba:    .ac = m.ca:    .ad = m.da:    .ae = m.ea:    .af = m.fa
                        .ba = m.ab:    .bb = m.bb:    .bc = m.cb:    .bd = m.db:    .be = m.eb:    .bf = m.fb
                        .ca = m.ac:    .cb = m.bc:    .cc = m.cc:    .cd = m.dc:    .ce = m.ec:    .cf = m.fc
                        .da = m.ad:    .db = m.bd:    .dc = m.cd:    .dd = m.dd:    .de = m.ed:    .df = m.fd
                        .ea = m.ae:    .eb = m.be:    .ec = m.ce:    .ed = m.de:    .ee = m.ee:    .ef = m.fe
                        .fa = m.af:    .fb = m.bf:    .fc = m.cf:    .fd = m.df:    .fe = m.ef:    .ff = m.ff:    End With
End Function

'Matrizenvergleich
Public Function Matrix2_IsEqual(m1 As Matrix2, m2 As Matrix2) As Boolean
    'Liefert True wenn beide 2x2-Matrizen gleich sind
    Matrix2_IsEqual = (m1.aa = m2.aa And m1.ab = m2.ab) And _
                      (m1.ba = m2.ba And m1.bb = m2.bb)
End Function
Public Function Matrix3_IsEqual(m1 As Matrix3, m2 As Matrix3) As Boolean
    'Liefert True wenn beide 3x3-Matrizen gleich sind
    Matrix3_IsEqual = (m1.aa = m2.aa And m1.ab = m2.ab And m1.ac = m2.ac) And _
                      (m1.ba = m2.ba And m1.bb = m2.bb And m1.bc = m2.bc) And _
                      (m1.ca = m2.ca And m1.cb = m2.cb And m1.cc = m2.cc)
End Function
Public Function Matrix4_IsEqual(m1 As Matrix4, m2 As Matrix4) As Boolean
    'Liefert True wenn beide 4x4-Matrizen gleich sind
    Matrix4_IsEqual = (m1.aa = m2.aa And m1.ab = m2.ab And m1.ac = m2.ac And m1.ad = m2.ad) And _
                      (m1.ba = m2.ba And m1.bb = m2.bb And m1.bc = m2.bc And m1.bd = m2.bd) And _
                      (m1.ca = m2.ca And m1.cb = m2.cb And m1.cc = m2.cc And m1.cd = m2.cd) And _
                      (m1.da = m2.da And m1.db = m2.db And m1.dc = m2.dc And m1.dd = m2.dd)
End Function
Public Function Matrix5_IsEqual(m1 As Matrix5, m2 As Matrix5) As Boolean
    'Liefert True wenn beide 5x5-Matrizen gleich sind
    Matrix5_IsEqual = (m1.aa = m2.aa And m1.ab = m2.ab And m1.ac = m2.ac And m1.ad = m2.ad And m1.ae = m2.ae) And _
                      (m1.ba = m2.ba And m1.bb = m2.bb And m1.bc = m2.bc And m1.bd = m2.bd And m1.be = m2.be) And _
                      (m1.ca = m2.ca And m1.cb = m2.cb And m1.cc = m2.cc And m1.cd = m2.cd And m1.ce = m2.ce) And _
                      (m1.da = m2.da And m1.db = m2.db And m1.dc = m2.dc And m1.dd = m2.dd And m1.de = m2.de) And _
                      (m1.ea = m2.ea And m1.eb = m2.eb And m1.ec = m2.ec And m1.ed = m2.ed And m1.ee = m2.ee)
End Function
Public Function Matrix6_IsEqual(m1 As Matrix6, m2 As Matrix6) As Boolean
    'Liefert True wenn beide 6x6-Matrizen gleich sind
    Matrix6_IsEqual = (m1.aa = m2.aa And m1.ab = m2.ab And m1.ac = m2.ac And m1.ad = m2.ad And m1.ae = m2.ae And m1.af = m2.af) And _
                      (m1.ba = m2.ba And m1.bb = m2.bb And m1.bc = m2.bc And m1.bd = m2.bd And m1.be = m2.be And m1.bf = m2.bf) And _
                      (m1.ca = m2.ca And m1.cb = m2.cb And m1.cc = m2.cc And m1.cd = m2.cd And m1.ce = m2.ce And m1.cf = m2.cf) And _
                      (m1.da = m2.da And m1.db = m2.db And m1.dc = m2.dc And m1.dd = m2.dd And m1.de = m2.de And m1.df = m2.df) And _
                      (m1.ea = m2.ea And m1.eb = m2.eb And m1.ec = m2.ec And m1.ed = m2.ed And m1.ee = m2.ee And m1.ef = m2.ef) And _
                      (m1.fa = m2.fa And m1.fb = m2.fb And m1.fc = m2.fc And m1.fd = m2.fd And m1.fe = m2.fe And m1.ff = m2.ff)
End Function

'Berechnung der Determinante
Public Function Matrix2_det(m As Matrix2) As Double
    'Berechnet die Determinante einer 2x2-Matrix
    With m
        Matrix2_det = .aa * .bb - .ab * .ba
    End With
End Function
Public Function Matrix3_det(m As Matrix3) As Double
    'Berechnet die Determinante einer 3x3-Matrix
    With m
        Matrix3_det = .aa * .bb * .cc + .ab * .bc * .ca + .ac * .ba * .cb _
                    - .ac * .bb * .ca - .ab * .ba * .cc - .aa * .bc * .cb
    End With
End Function
'oder alternativ:
'Public Function Matrix3_det2(m As Matrix3) As Double
'    'Berechnet die Determinante einer 3x3-Matrix
'    With m
'        Matrix3_det2 = .aa * Matrix2_det(Mat2(.bb, .bc, .cb, .cc)) _
'                     - .ab * Matrix2_det(Mat2(.ba, .bc, .ca, .cc)) _
'                     + .ac * Matrix2_det(Mat2(.ba, .bb, .ca, .cb))
'    End With
'End Function
Public Function Matrix4_det(m As Matrix4) As Double
    'Berechnet die Determinante einer 4x4-Matrix
    With m
        Matrix4_det = .aa * Matrix3_det(Mat3(.bb, .bc, .bd, .cb, .cc, .cd, .db, .dc, .dd)) _
                    - .ab * Matrix3_det(Mat3(.ba, .bc, .bd, .ca, .cc, .cd, .da, .dc, .dd)) _
                    + .ac * Matrix3_det(Mat3(.ba, .bb, .bd, .ca, .cb, .cd, .da, .db, .dd)) _
                    - .ad * Matrix3_det(Mat3(.ba, .bb, .bc, .ca, .cb, .cc, .da, .db, .dc))
    End With
End Function
Public Function Matrix5_det(m As Matrix5) As Double
    'Berechnet die Determinante einer 5x5-Matrix
    With m
        Matrix5_det = .aa * Matrix4_det(Mat4(.bb, .bc, .bd, .be, .cb, .cc, .cd, .ce, .db, .dc, .dd, .de, .eb, .ec, .ed, .ee)) _
                    - .ab * Matrix4_det(Mat4(.ba, .bc, .bd, .be, .ca, .cc, .cd, .ce, .da, .dc, .dd, .de, .ea, .ec, .ed, .ee)) _
                    + .ac * Matrix4_det(Mat4(.ba, .bb, .bd, .be, .ca, .cb, .cd, .ce, .da, .db, .dd, .de, .ea, .eb, .ed, .ee)) _
                    - .ad * Matrix4_det(Mat4(.ba, .bb, .bc, .be, .ca, .cb, .cc, .ce, .da, .db, .dc, .de, .ea, .eb, .ec, .ee)) _
                    + .ae * Matrix4_det(Mat4(.ba, .bb, .bc, .bd, .ca, .cb, .cc, .cd, .da, .db, .dc, .dd, .ea, .eb, .ec, .ed))
    End With
End Function
Public Function Matrix6_det(m As Matrix6) As Double
    'Berechnet die Determinante einer 6x6-Matrix
    With m
        Matrix6_det = .aa * Matrix5_det(Mat5(.bb, .bc, .bd, .be, .bf, .cb, .cc, .cd, .ce, .cf, .db, .dc, .dd, .de, .df, .eb, .ec, .ed, .ee, .ef, .fb, .fc, .fd, .fe, .ff)) _
                    - .ab * Matrix5_det(Mat5(.ba, .bc, .bd, .be, .bf, .ca, .cc, .cd, .ce, .cf, .da, .dc, .dd, .de, .df, .ea, .ec, .ed, .ee, .ef, .fa, .fc, .fd, .fe, .ff)) _
                    + .ac * Matrix5_det(Mat5(.ba, .bb, .bd, .be, .bf, .ca, .cb, .cd, .ce, .cf, .da, .db, .dd, .de, .df, .ea, .eb, .ed, .ee, .ef, .fa, .fb, .fd, .fe, .ff)) _
                    - .ad * Matrix5_det(Mat5(.ba, .bb, .bc, .be, .bf, .ca, .cb, .cc, .ce, .cf, .da, .db, .dc, .de, .df, .ea, .eb, .ec, .ee, .ef, .fa, .fb, .fc, .fe, .ff)) _
                    + .ae * Matrix5_det(Mat5(.ba, .bb, .bc, .bd, .bf, .ca, .cb, .cc, .cd, .cf, .da, .db, .dc, .dd, .df, .ea, .eb, .ec, .ed, .ef, .fa, .fb, .fc, .fd, .ff)) _
                    - .af * Matrix5_det(Mat5(.ba, .bb, .bc, .bd, .be, .ca, .cb, .cc, .cd, .ce, .da, .db, .dc, .dd, .de, .ea, .eb, .ec, .ed, .ee, .fa, .fb, .fc, .fd, .fe))
    End With
End Function

'Untermatrix
Public Function Matrix2_umat(m As Matrix2, ByVal r_ex As Long, ByVal c_ex As Long) As Double
    Select Case r_ex: Case 0: Select Case c_ex: Case 0: Matrix2_umat = m.bb
                                                Case 1: Matrix2_umat = m.ba: End Select
                      Case 1: Select Case c_ex: Case 0: Matrix2_umat = m.ab
                                                Case 1: Matrix2_umat = m.aa: End Select: End Select
End Function
Public Function Matrix3_umat(m As Matrix3, ByVal r_ex As Long, ByVal c_ex As Long) As Matrix2
    'Liefert aus einer 3x3-Matrix die 2x2-Untermatrix, indem die Zeile r_ex und die Spalte c_ex ausgelassen werden.
    Dim icex As Long
    If r_ex <> 0 Then Mat2_Row(Matrix3_umat, icex) = Vector3_uvec(Mat3_Row(m, 0), c_ex): icex = icex + 1
    If r_ex <> 1 Then Mat2_Row(Matrix3_umat, icex) = Vector3_uvec(Mat3_Row(m, 1), c_ex): icex = icex + 1
    If r_ex <> 2 Then Mat2_Row(Matrix3_umat, icex) = Vector3_uvec(Mat3_Row(m, 2), c_ex): icex = icex + 1
End Function
Public Function Matrix4_umat(m As Matrix4, ByVal r_ex As Long, ByVal c_ex As Long) As Matrix3
    'Liefert aus einer 4x4-Matrix die 3x3-Untermatrix, indem die Zeile r_ex und die Spalte c_ex ausgelassen werden.
    Dim icex As Long
    If r_ex <> 0 Then Mat3_Row(Matrix4_umat, icex) = Vector4_uvec(Mat4_Row(m, 0), c_ex): icex = icex + 1
    If r_ex <> 1 Then Mat3_Row(Matrix4_umat, icex) = Vector4_uvec(Mat4_Row(m, 1), c_ex): icex = icex + 1
    If r_ex <> 2 Then Mat3_Row(Matrix4_umat, icex) = Vector4_uvec(Mat4_Row(m, 2), c_ex): icex = icex + 1
    If r_ex <> 3 Then Mat3_Row(Matrix4_umat, icex) = Vector4_uvec(Mat4_Row(m, 3), c_ex): icex = icex + 1
End Function
Public Function Matrix5_umat(m As Matrix5, ByVal r_ex As Long, ByVal c_ex As Long) As Matrix4
    'Liefert aus einer 5x5-Matrix die 4x4-Untermatrix, indem die Zeile r_ex und die Spalte c_ex ausgelassen werden.
    Dim icex As Long
    If r_ex <> 0 Then Mat4_Row(Matrix5_umat, icex) = Vector5_uvec(Mat5_Row(m, 0), c_ex): icex = icex + 1
    If r_ex <> 1 Then Mat4_Row(Matrix5_umat, icex) = Vector5_uvec(Mat5_Row(m, 1), c_ex): icex = icex + 1
    If r_ex <> 2 Then Mat4_Row(Matrix5_umat, icex) = Vector5_uvec(Mat5_Row(m, 2), c_ex): icex = icex + 1
    If r_ex <> 3 Then Mat4_Row(Matrix5_umat, icex) = Vector5_uvec(Mat5_Row(m, 3), c_ex): icex = icex + 1
    If r_ex <> 4 Then Mat4_Row(Matrix5_umat, icex) = Vector5_uvec(Mat5_Row(m, 4), c_ex): icex = icex + 1
End Function
Public Function Matrix6_umat(m As Matrix6, ByVal r_ex As Long, ByVal c_ex As Long) As Matrix5
    'Liefert aus einer 6x6-Matrix die 5x5-Untermatrix, indem die Zeile r_ex und die Spalte c_ex ausgelassen werden.
    Dim icex As Long
    If r_ex <> 0 Then Mat5_Row(Matrix6_umat, icex) = Vector6_uvec(Mat6_Row(m, 0), c_ex): icex = icex + 1
    If r_ex <> 1 Then Mat5_Row(Matrix6_umat, icex) = Vector6_uvec(Mat6_Row(m, 1), c_ex): icex = icex + 1
    If r_ex <> 2 Then Mat5_Row(Matrix6_umat, icex) = Vector6_uvec(Mat6_Row(m, 2), c_ex): icex = icex + 1
    If r_ex <> 3 Then Mat5_Row(Matrix6_umat, icex) = Vector6_uvec(Mat6_Row(m, 3), c_ex): icex = icex + 1
    If r_ex <> 4 Then Mat5_Row(Matrix6_umat, icex) = Vector6_uvec(Mat6_Row(m, 4), c_ex): icex = icex + 1
    If r_ex <> 5 Then Mat5_Row(Matrix6_umat, icex) = Vector6_uvec(Mat6_Row(m, 5), c_ex): icex = icex + 1
End Function

'Berechnet die Minoren = Determinanten der Untermatrizen
Public Function Matrix2_min(m As Matrix2, ByVal r_ex As Long, ByVal c_ex As Long) As Double
    'Berechnet die Minore einer 2x2-Matrix bei der die Zeile r_ex und die Spalte c_ex ausgelassen werden.
    Matrix2_min = Matrix2_umat(m, r_ex, c_ex)
End Function
Public Function Matrix3_min(m As Matrix3, ByVal r_ex As Long, ByVal c_ex As Long) As Double
    'Berechnet die Minore einer 3x3-Matrix bei der die Zeile r_ex und die Spalte c_ex ausgelassen werden.
    Matrix3_min = Matrix2_det(Matrix3_umat(m, r_ex, c_ex))
End Function
Public Function Matrix4_min(m As Matrix4, ByVal r_ex As Long, ByVal c_ex As Long) As Double
    'Berechnet die Minore einer 4x4-Matrix bei der die Zeile r_ex und die Spalte c_ex ausgelassen werden.
    Matrix4_min = Matrix3_det(Matrix4_umat(m, r_ex, c_ex))
End Function
Public Function Matrix5_min(m As Matrix5, ByVal r_ex As Long, ByVal c_ex As Long) As Double
    'Berechnet die Minore einer 5x5-Matrix bei der die Zeile r_ex und die Spalte c_ex ausgelassen werden.
    Matrix5_min = Matrix4_det(Matrix5_umat(m, r_ex, c_ex))
End Function
Public Function Matrix6_min(m As Matrix6, ByVal r_ex As Long, ByVal c_ex As Long) As Double
    'Berechnet die Minore einer 5x5-Matrix bei der die Zeile r_ex und die Spalte c_ex ausgelassen werden.
    Matrix6_min = Matrix5_det(Matrix6_umat(m, r_ex, c_ex))
End Function

'Adjunkte
Public Function Matrix2_Adj(m As Matrix2) As Matrix2
    With m
        Matrix2_Adj = Mat2(.bb, -.ab, _
                           -.ba, .aa)
    End With
End Function
Public Function Matrix3_Adj(m As Matrix3) As Matrix3
    Matrix3_Adj = Mat3(Matrix3_min(m, 0, 0), -Matrix3_min(m, 1, 0), Matrix3_min(m, 2, 0), _
                       -Matrix3_min(m, 0, 1), Matrix3_min(m, 1, 1), -Matrix3_min(m, 2, 1), _
                       Matrix3_min(m, 0, 2), -Matrix3_min(m, 1, 2), Matrix3_min(m, 2, 2))
End Function
Public Function Matrix4_Adj(m As Matrix4) As Matrix4
    Matrix4_Adj = Mat4(Matrix4_min(m, 0, 0), -Matrix4_min(m, 1, 0), Matrix4_min(m, 2, 0), -Matrix4_min(m, 3, 0), _
                       -Matrix4_min(m, 0, 1), Matrix4_min(m, 1, 1), -Matrix4_min(m, 2, 1), Matrix4_min(m, 3, 1), _
                       Matrix4_min(m, 0, 2), -Matrix4_min(m, 1, 2), Matrix4_min(m, 2, 2), -Matrix4_min(m, 3, 2), _
                       -Matrix4_min(m, 0, 3), Matrix4_min(m, 1, 3), -Matrix4_min(m, 2, 3), Matrix4_min(m, 3, 3))
End Function
Public Function Matrix5_Adj(m As Matrix5) As Matrix5
    Matrix5_Adj = Mat5(Matrix5_min(m, 0, 0), -Matrix5_min(m, 1, 0), Matrix5_min(m, 2, 0), -Matrix5_min(m, 3, 0), Matrix5_min(m, 4, 0), _
                       -Matrix5_min(m, 0, 1), Matrix5_min(m, 1, 1), -Matrix5_min(m, 2, 1), Matrix5_min(m, 3, 1), -Matrix5_min(m, 4, 1), _
                       Matrix5_min(m, 0, 2), -Matrix5_min(m, 1, 2), Matrix5_min(m, 2, 2), -Matrix5_min(m, 3, 2), Matrix5_min(m, 4, 2), _
                       -Matrix5_min(m, 0, 3), Matrix5_min(m, 1, 3), -Matrix5_min(m, 2, 3), Matrix5_min(m, 3, 3), -Matrix5_min(m, 4, 3), _
                       Matrix5_min(m, 0, 4), -Matrix5_min(m, 1, 4), Matrix5_min(m, 2, 4), -Matrix5_min(m, 3, 4), Matrix5_min(m, 4, 4))
End Function
Public Function Matrix6_Adj(m As Matrix6) As Matrix6
    Matrix6_Adj = Mat6(Matrix6_min(m, 0, 0), -Matrix6_min(m, 1, 0), Matrix6_min(m, 2, 0), -Matrix6_min(m, 3, 0), Matrix6_min(m, 4, 0), -Matrix6_min(m, 5, 0), _
                       -Matrix6_min(m, 0, 1), Matrix6_min(m, 1, 1), -Matrix6_min(m, 2, 1), Matrix6_min(m, 3, 1), -Matrix6_min(m, 4, 1), Matrix6_min(m, 5, 1), _
                       Matrix6_min(m, 0, 2), -Matrix6_min(m, 1, 2), Matrix6_min(m, 2, 2), -Matrix6_min(m, 3, 2), Matrix6_min(m, 4, 2), -Matrix6_min(m, 5, 2), _
                       -Matrix6_min(m, 0, 3), Matrix6_min(m, 1, 3), -Matrix6_min(m, 2, 3), Matrix6_min(m, 3, 3), -Matrix6_min(m, 4, 3), Matrix6_min(m, 5, 3), _
                       Matrix6_min(m, 0, 4), -Matrix6_min(m, 1, 4), Matrix6_min(m, 2, 4), -Matrix6_min(m, 3, 4), Matrix6_min(m, 4, 4), -Matrix6_min(m, 5, 4), _
                       -Matrix6_min(m, 0, 5), Matrix6_min(m, 1, 5), -Matrix6_min(m, 2, 5), Matrix6_min(m, 3, 5), -Matrix6_min(m, 4, 5), Matrix6_min(m, 5, 5))
End Function

'Inverse
Public Function Matrix2_inv(m As Matrix2) As Matrix2
    Dim det As Double: det = Matrix2_det(m): If det = 0 Then Exit Function
    Matrix2_inv = Matrix2_smul(Matrix2_Adj(m), 1 / det)
End Function
Public Function Matrix3_inv(m As Matrix3) As Matrix3
    Dim det As Double: det = Matrix3_det(m): If det = 0 Then Exit Function
    Matrix3_inv = Matrix3_smul(Matrix3_Adj(m), 1 / det)
End Function
Public Function Matrix4_inv(m As Matrix4) As Matrix4
    Dim det As Double: det = Matrix4_det(m): If det = 0 Then Exit Function
    Matrix4_inv = Matrix4_smul(Matrix4_Adj(m), 1 / det)
End Function
Public Function Matrix5_inv(m As Matrix5) As Matrix5
    Dim det As Double: det = Matrix5_det(m): If det = 0 Then Exit Function
    Matrix5_inv = Matrix5_smul(Matrix5_Adj(m), 1 / det)
End Function
Public Function Matrix6_inv(m As Matrix6) As Matrix6
    Dim det As Double: det = Matrix6_det(m): If det = 0 Then Exit Function
    Matrix6_inv = Matrix6_smul(Matrix6_Adj(m), 1 / det)
End Function

'Lösen von LGS
Public Function Matrix2_solve(m As Matrix2, b As Vector2) As Vector2
    Matrix2_solve = Matrix2_vmul(Matrix2_inv(m), b)
End Function
Public Function Matrix3_solve(m As Matrix3, b As Vector3) As Vector3
    Matrix3_solve = Matrix3_vmul(Matrix3_inv(m), b)
End Function
Public Function Matrix4_solve(m As Matrix4, b As Vector4) As Vector4
    Matrix4_solve = Matrix4_vmul(Matrix4_inv(m), b)
End Function
Public Function Matrix5_solve(m As Matrix5, b As Vector5) As Vector5
    Matrix5_solve = Matrix5_vmul(Matrix5_inv(m), b)
End Function
Public Function Matrix6_solve(m As Matrix6, b As Vector6) As Vector6
    Matrix6_solve = Matrix6_vmul(Matrix6_inv(m), b)
End Function

'Lesen und Schreiben
'allgemein
'Public Sub Matrix_Parse(t As String, ByVal mRows As Long, ByVal nCols As Long, ByRef a_out() As Double)
Public Function Matrix_Parse(t As String, ByVal mRows As Long, ByVal nCols As Long) As Double()
    ReDim a_out(0 To nCols - 1, 0 To mRows - 1) As Double
    Dim saLines() As String: saLines = Split(DeleteMultiWS(t), vbCrLf)
    Dim sa() As String
    Dim i As Long, j As Long
    For i = 0 To mRows - 1
        If UBound(saLines) < i Then Exit For
        sa = Split(DeleteMultiWS(saLines(i)), " ")
        For j = 0 To nCols - 1
            If UBound(sa) < j Then Exit For
            a_out(j, i) = Val(sa(j))
        Next
    Next
    Matrix_Parse = a_out
End Function
Public Function Matrix2_Parse(t As String) As Matrix2
    Dim mRows As Long: mRows = 2
    Dim nCols As Long: nCols = 2
    Dim a() As Double: a = Matrix_Parse(t, mRows, nCols)
    RtlMoveMemory Matrix2_Parse, a(0, 0), mRows * nCols * 8
End Function
Public Function Matrix3_Parse(t As String) As Matrix3
    Dim mRows As Long: mRows = 3
    Dim nCols As Long: nCols = 3
    Dim a() As Double: a = Matrix_Parse(t, mRows, nCols)
    RtlMoveMemory Matrix3_Parse, a(0, 0), mRows * nCols * 8
End Function
Public Function Matrix4_Parse(t As String) As Matrix4
    Dim mRows As Long: mRows = 4
    Dim nCols As Long: nCols = 4
    Dim a() As Double: a = Matrix_Parse(t, mRows, nCols)
    RtlMoveMemory Matrix4_Parse, a(0, 0), mRows * nCols * 8
End Function
Public Function Matrix5_Parse(t As String) As Matrix5
    Dim mRows As Long: mRows = 5
    Dim nCols As Long: nCols = 5
    Dim a() As Double: a = Matrix_Parse(t, mRows, nCols)
    RtlMoveMemory Matrix5_Parse, a(0, 0), mRows * nCols * 8
End Function
Public Function Matrix6_Parse(t As String) As Matrix6
    Dim mRows As Long: mRows = 6
    Dim nCols As Long: nCols = 6
    Dim a() As Double: a = Matrix_Parse(t, mRows, nCols)
    RtlMoveMemory Matrix6_Parse, a(0, 0), mRows * nCols * 8
End Function

Public Function Vector2_ToStr(v As Vector2, Optional bIsLineVec As Boolean) As String
    Dim sa(0 To 1) As String
    With v: sa(0) = TStr(.a): sa(1) = TStr(.b): End With
    Vector2_ToStr = Join(sa, IIf(bIsLineVec, " ", vbCrLf))
End Function
Public Function Vector3_ToStr(v As Vector3, Optional bIsLineVec As Boolean) As String
    Dim sa(0 To 2) As String
    With v: sa(0) = TStr(.a): sa(1) = TStr(.b): sa(2) = TStr(.c): End With
    Vector3_ToStr = Join(sa, IIf(bIsLineVec, " ", vbCrLf))
End Function
Public Function Vector4_ToStr(v As Vector4, Optional bIsLineVec As Boolean) As String
    Dim sa(0 To 3) As String
    With v: sa(0) = TStr(.a): sa(1) = TStr(.b): sa(2) = TStr(.c): sa(3) = TStr(.d): End With
    Vector4_ToStr = Join(sa, IIf(bIsLineVec, " ", vbCrLf))
End Function
Public Function Vector5_ToStr(v As Vector5, Optional bIsLineVec As Boolean) As String
    Dim sa(0 To 4) As String
    With v: sa(0) = TStr(.a): sa(1) = TStr(.b): sa(2) = TStr(.c): sa(3) = TStr(.d): sa(4) = TStr(.e): End With
    Vector5_ToStr = Join(sa, IIf(bIsLineVec, " ", vbCrLf))
End Function
Public Function Vector6_ToStr(v As Vector6, Optional bIsLineVec As Boolean) As String
    Dim sa(0 To 5) As String
    With v: sa(0) = TStr(.a): sa(1) = TStr(.b): sa(2) = TStr(.c): sa(3) = TStr(.d): sa(4) = TStr(.e): sa(5) = TStr(.f): End With
    Vector6_ToStr = Join(sa, IIf(bIsLineVec, " ", vbCrLf))
End Function

Public Function Matrix2_ToStr(m As Matrix2) As String
    Dim s As String: s = ""
    With m
        s = s & TStr(.aa) & " " & TStr(.ab) & vbCrLf
        s = s & TStr(.ba) & " " & TStr(.bb)
    End With
    Matrix2_ToStr = s
End Function
Public Function Matrix3_ToStr(m As Matrix3) As String
    Dim s As String: s = ""
    With m
        s = s & TStr(.aa) & " " & TStr(.ab) & " " & TStr(.ac) & vbCrLf
        s = s & TStr(.ba) & " " & TStr(.bb) & " " & TStr(.bc) & vbCrLf
        s = s & TStr(.ca) & " " & TStr(.cb) & " " & TStr(.cc)
    End With
    Matrix3_ToStr = s
End Function
Public Function Matrix4_ToStr(m As Matrix4) As String
    Dim s As String: s = ""
    With m
        s = s & TStr(.aa) & " " & TStr(.ab) & " " & TStr(.ac) & " " & TStr(.ad) & vbCrLf
        s = s & TStr(.ba) & " " & TStr(.bb) & " " & TStr(.bc) & " " & TStr(.bd) & vbCrLf
        s = s & TStr(.ca) & " " & TStr(.cb) & " " & TStr(.cc) & " " & TStr(.cd) & vbCrLf
        s = s & TStr(.da) & " " & TStr(.db) & " " & TStr(.dc) & " " & TStr(.dd)
    End With
    Matrix4_ToStr = s
End Function
Public Function Matrix5_ToStr(m As Matrix5) As String
    Dim s As String: s = ""
    With m
        s = s & TStr(.aa) & " " & TStr(.ab) & " " & TStr(.ac) & " " & TStr(.ad) & " " & TStr(.ae) & vbCrLf
        s = s & TStr(.ba) & " " & TStr(.bb) & " " & TStr(.bc) & " " & TStr(.bd) & " " & TStr(.be) & vbCrLf
        s = s & TStr(.ca) & " " & TStr(.cb) & " " & TStr(.cc) & " " & TStr(.cd) & " " & TStr(.ce) & vbCrLf
        s = s & TStr(.da) & " " & TStr(.db) & " " & TStr(.dc) & " " & TStr(.dd) & " " & TStr(.de) & vbCrLf
        s = s & TStr(.ea) & " " & TStr(.eb) & " " & TStr(.ec) & " " & TStr(.ed) & " " & TStr(.ee)
    End With
    Matrix5_ToStr = s
End Function
Public Function Matrix6_ToStr(m As Matrix6) As String
    Dim s As String: s = ""
    With m
        s = s & TStr(.aa) & " " & TStr(.ab) & " " & TStr(.ac) & " " & TStr(.ad) & " " & TStr(.ae) & " " & TStr(.af) & vbCrLf
        s = s & TStr(.ba) & " " & TStr(.bb) & " " & TStr(.bc) & " " & TStr(.bd) & " " & TStr(.be) & " " & TStr(.bf) & vbCrLf
        s = s & TStr(.ca) & " " & TStr(.cb) & " " & TStr(.cc) & " " & TStr(.cd) & " " & TStr(.ce) & " " & TStr(.cf) & vbCrLf
        s = s & TStr(.da) & " " & TStr(.db) & " " & TStr(.dc) & " " & TStr(.dd) & " " & TStr(.de) & " " & TStr(.df) & vbCrLf
        s = s & TStr(.ea) & " " & TStr(.eb) & " " & TStr(.ec) & " " & TStr(.ed) & " " & TStr(.ee) & " " & TStr(.ef) & vbCrLf
        s = s & TStr(.fa) & " " & TStr(.fb) & " " & TStr(.fc) & " " & TStr(.fd) & " " & TStr(.fe) & " " & TStr(.ff)
    End With
    Matrix6_ToStr = s
End Function

'Lesen/Schreiben von Zeilen oder Spalten aus/in eine Matrix
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

Public Property Get Mat2_Col(m As Matrix2, ByVal index As Long) As Vector2
    With m
        Select Case index
        Case 0: Mat2_Col = Vec2(.aa, .ba)
        Case 1: Mat2_Col = Vec2(.ab, .bb)
        End Select
    End With
End Property
Public Property Let Mat2_Col(m As Matrix2, ByVal index As Long, v As Vector2)
    With m
        Select Case index
        Case 0: .aa = v.a: .ba = v.b
        Case 1: .ab = v.a: .bb = v.b
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

Public Property Get Mat4_Row(m As Matrix4, ByVal index As Long) As Vector4
    With m
        Select Case index
        Case 0: Mat4_Row = Vec4(.aa, .ab, .ac, .ad)
        Case 1: Mat4_Row = Vec4(.ba, .bb, .bc, .bd)
        Case 2: Mat4_Row = Vec4(.ca, .cb, .cc, .cd)
        Case 3: Mat4_Row = Vec4(.da, .db, .dc, .dd)
        End Select
    End With
End Property
Public Property Let Mat4_Row(m As Matrix4, ByVal index As Long, v As Vector4)
    With m
        Select Case index
        Case 0: .aa = v.a: .ab = v.b: .ac = v.c: .ad = v.d
        Case 1: .ba = v.a: .bb = v.b: .bc = v.c: .bd = v.d
        Case 2: .ca = v.a: .cb = v.b: .cc = v.c: .cd = v.d
        Case 3: .da = v.a: .db = v.b: .dc = v.c: .dd = v.d
        End Select
    End With
End Property
Public Property Get Mat4_Col(m As Matrix4, ByVal index As Long) As Vector4
    With m
        Select Case index
        Case 0: Mat4_Col = Vec4(.aa, .ba, .ca, .da)
        Case 1: Mat4_Col = Vec4(.ab, .bb, .cb, .db)
        Case 2: Mat4_Col = Vec4(.ac, .bc, .cc, .dc)
        Case 3: Mat4_Col = Vec4(.ad, .bd, .cd, .dd)
        End Select
    End With
End Property
Public Property Let Mat4_Col(m As Matrix4, ByVal index As Long, v As Vector4)
    With m
        Select Case index
        Case 0: .aa = v.a: .ba = v.b: .ca = v.c: .da = v.d
        Case 1: .ab = v.a: .bb = v.b: .cb = v.c: .db = v.d
        Case 2: .ac = v.a: .bc = v.b: .cc = v.c: .dc = v.d
        Case 3: .ad = v.a: .bd = v.b: .cd = v.c: .dd = v.d
        End Select
    End With
End Property

Public Property Get Mat5_Row(m As Matrix5, ByVal index As Long) As Vector5
    With m
        Select Case index
        Case 0: Mat5_Row = Vec5(.aa, .ab, .ac, .ad, .ae)
        Case 1: Mat5_Row = Vec5(.ba, .bb, .bc, .bd, .be)
        Case 2: Mat5_Row = Vec5(.ca, .cb, .cc, .cd, .ce)
        Case 3: Mat5_Row = Vec5(.da, .db, .dc, .dd, .de)
        Case 4: Mat5_Row = Vec5(.ea, .eb, .ec, .ed, .ee)
        End Select
    End With
End Property
Public Property Let Mat5_Row(m As Matrix5, ByVal index As Long, v As Vector5)
    With m
        Select Case index
        Case 0: .aa = v.a: .ab = v.b: .ac = v.c: .ad = v.d: .ae = v.e
        Case 1: .ba = v.a: .bb = v.b: .bc = v.c: .bd = v.d: .be = v.e
        Case 2: .ca = v.a: .cb = v.b: .cc = v.c: .cd = v.d: .ce = v.e
        Case 3: .da = v.a: .db = v.b: .dc = v.c: .dd = v.d: .de = v.e
        Case 4: .ea = v.a: .eb = v.b: .ec = v.c: .ed = v.d: .ee = v.e
        End Select
    End With
End Property
Public Property Get Mat5_Col(m As Matrix5, ByVal index As Long) As Vector5
    With m
        Select Case index
        Case 0: Mat5_Col = Vec5(.aa, .ba, .ca, .da, .ea)
        Case 1: Mat5_Col = Vec5(.ab, .bb, .cb, .db, .eb)
        Case 2: Mat5_Col = Vec5(.ac, .bc, .cc, .dc, .ec)
        Case 3: Mat5_Col = Vec5(.ad, .bd, .cd, .dd, .ed)
        Case 4: Mat5_Col = Vec5(.ae, .be, .ce, .de, .ee)
        End Select
    End With
End Property
Public Property Let Mat5_Col(m As Matrix5, ByVal index As Long, v As Vector5)
    With m
        Select Case index
        Case 0: .aa = v.a: .ba = v.b: .ca = v.c: .da = v.d: .ea = v.e
        Case 1: .ab = v.a: .bb = v.b: .cb = v.c: .db = v.d: .eb = v.e
        Case 2: .ac = v.a: .bc = v.b: .cc = v.c: .dc = v.d: .ec = v.e
        Case 3: .ad = v.a: .bd = v.b: .cd = v.c: .dd = v.d: .ed = v.e
        Case 4: .ae = v.a: .be = v.b: .ce = v.c: .de = v.d: .ee = v.e
        End Select
    End With
End Property

Public Property Get Mat6_Row(m As Matrix6, ByVal index As Long) As Vector6
    With m
        Select Case index
        Case 0: Mat6_Row = Vec6(.aa, .ab, .ac, .ad, .ae, .af)
        Case 1: Mat6_Row = Vec6(.ba, .bb, .bc, .bd, .be, .bf)
        Case 2: Mat6_Row = Vec6(.ca, .cb, .cc, .cd, .ce, .cf)
        Case 3: Mat6_Row = Vec6(.da, .db, .dc, .dd, .de, .df)
        Case 4: Mat6_Row = Vec6(.ea, .eb, .ec, .ed, .ee, .ef)
        Case 5: Mat6_Row = Vec6(.fa, .fb, .fc, .fd, .fe, .ff)
        End Select
    End With
End Property
Public Property Let Mat6_Row(m As Matrix6, ByVal index As Long, v As Vector6)
    With m
        Select Case index
        Case 0: .aa = v.a: .ab = v.b: .ac = v.c: .ad = v.d: .ae = v.e: .af = v.f
        Case 1: .ba = v.a: .bb = v.b: .bc = v.c: .bd = v.d: .be = v.e: .bf = v.f
        Case 2: .ca = v.a: .cb = v.b: .cc = v.c: .cd = v.d: .ce = v.e: .cf = v.f
        Case 3: .da = v.a: .db = v.b: .dc = v.c: .dd = v.d: .de = v.e: .df = v.f
        Case 4: .ea = v.a: .eb = v.b: .ec = v.c: .ed = v.d: .ee = v.e: .ef = v.f
        Case 5: .fa = v.a: .fb = v.b: .fc = v.c: .fd = v.d: .fe = v.e: .ff = v.f
        End Select
    End With
End Property
Public Property Get Mat6_Col(m As Matrix6, ByVal index As Long) As Vector6
    With m
        Select Case index
        Case 0: Mat6_Col = Vec6(.aa, .ba, .ca, .da, .ea, .fa)
        Case 1: Mat6_Col = Vec6(.ab, .bb, .cb, .db, .eb, .fb)
        Case 2: Mat6_Col = Vec6(.ac, .bc, .cc, .dc, .ec, .fc)
        Case 3: Mat6_Col = Vec6(.ad, .bd, .cd, .dd, .ed, .fd)
        Case 4: Mat6_Col = Vec6(.ae, .be, .ce, .de, .ee, .fe)
        Case 5: Mat6_Col = Vec6(.af, .bf, .cf, .de, .ef, .ff)
        End Select
    End With
End Property
Public Property Let Mat6_Col(m As Matrix6, ByVal index As Long, v As Vector6)
    With m
        Select Case index
        Case 0: .aa = v.a: .ba = v.b: .ca = v.c: .da = v.d: .ea = v.e: .fa = v.f
        Case 1: .ab = v.a: .bb = v.b: .cb = v.c: .db = v.d: .eb = v.e: .fb = v.f
        Case 2: .ac = v.a: .bc = v.b: .cc = v.c: .dc = v.d: .ec = v.e: .fc = v.f
        Case 3: .ad = v.a: .bd = v.b: .cd = v.c: .dd = v.d: .ed = v.e: .fd = v.f
        Case 4: .ae = v.a: .be = v.b: .ce = v.c: .de = v.d: .ee = v.e: .fe = v.f
        Case 5: .af = v.a: .bf = v.b: .cf = v.c: .df = v.d: .ef = v.e: .ff = v.f
        End Select
    End With
End Property


'allgemein
'Public Function MatrixT_ToStr(m As MatrixT, ByVal mRows As Long, ByVal nCols As Long) As String
'    MatrixT_ToStr = MatrixA_ToStr(m.a, mRows, nCols)
'End Function

Public Function MatrixA_ToStr(a() As Double, ByVal mRows As Long, ByVal nCols As Long) As String
    Dim s As String ': s = ""
    Dim sl As String
    Dim i As Long, j As Long
    For i = 0 To mRows - 1
        For j = 0 To nCols - 1
            sl = sl & Str(a(j, i)) & " "
        Next
        s = s & Trim(sl) & vbCrLf
        sl = ""
    Next
    MatrixA_ToStr = s
End Function

Public Function Matrix_ToStr(ByVal p_Matrix As Long, ByVal mRows As Long, ByVal nCols As Long) As String
    'die allgemeine mathematische Anordnung  ist a(iZeile, jSpalte)
    'vgl die Speicheranordnung von VB-Arrays ist a(jSpalte, iZeile)
    If mRows = 0 Or nCols = 0 Then Exit Function
    ReDim a(0 To nCols - 1, 0 To mRows - 1) As Double
    RtlMoveMemory a(0, 0), ByVal p_Matrix, mRows * nCols * 8
    Matrix_ToStr = MatrixA_ToStr(a, mRows, nCols)
End Function

'Hilfsfunktionen
Public Function DeleteMultiWS(s As String) As String
    DeleteMultiWS = Trim$(s)
    If InStr(1, s, "  ") = 0 Then Exit Function
    DeleteMultiWS = Replace(s, "  ", " ")
    DeleteMultiWS = DeleteMultiWS(DeleteMultiWS)
End Function
Public Function Max(v1, v2)
    If v1 > v2 Then Max = v1 Else Max = v2
End Function
Public Function Min(v1, v2)
    If v1 < v2 Then Min = v1 Else Min = v2
End Function
Public Function TStr(d As Double) As String
    TStr = Trim$(Str$(d))
End Function
Public Function DblParse(ByVal s As String) As Double
    s = Replace(Trim$(s), ",", ".")
Try: On Error GoTo Catch
    DblParse = Val(s)
Catch: 'out
End Function

