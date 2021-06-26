Attribute VB_Name = "MMatrices"
Option Explicit
'Eine Zeile in VB darf maximal 1024 Zeichen lang sein, um Platz zu sparen verwenden wir die kürzest möglichen Variablennamen,
'und das sind reine Buchstaben, jeder Variablennamen muss mit einem Buchstaben beginnen.
'bei Verwendung der Form a11, a12, a13 etc werden bereits 3 Zeichen pro Variable benötigt
'Zeile, Spalte
Public Type Matrix2 'Matrix 2x2 bsp "ab" ist das zweite Element der ersten Zeile, oder das erste Element der zweiten Spalte
    aa As Double:    ab As Double
    ba As Double:    bb As Double
End Type
Public Type Matrix3 'Matrix 3x3
    aa As Double:    ab As Double:    ac As Double
    ba As Double:    bb As Double:    bc As Double
    ca As Double:    cb As Double:    cc As Double
End Type
Public Type Matrix4 'Matrix 4x4
    aa As Double:    ab As Double:    ac As Double:    ad As Double
    ba As Double:    bb As Double:    bc As Double:    bd As Double
    ca As Double:    cb As Double:    cc As Double:    cd As Double
    da As Double:    db As Double:    dc As Double:    dd As Double
End Type
Public Type Matrix5 'Matrix 5x5
    aa As Double:    ab As Double:    ac As Double:    ad As Double:    ae As Double
    ba As Double:    bb As Double:    bc As Double:    bd As Double:    be As Double
    ca As Double:    cb As Double:    cc As Double:    cd As Double:    ce As Double
    da As Double:    db As Double:    dc As Double:    dd As Double:    de As Double
    ea As Double:    eb As Double:    ec As Double:    ed As Double:    ee As Double
End Type
Public Type Matrix6 'Matrix 6x6
    aa As Double:    ab As Double:    ac As Double:    ad As Double:    ae As Double:    af As Double
    ba As Double:    bb As Double:    bc As Double:    bd As Double:    be As Double:    bf As Double
    ca As Double:    cb As Double:    cc As Double:    cd As Double:    ce As Double:    cf As Double
    da As Double:    db As Double:    dc As Double:    dd As Double:    de As Double:    df As Double
    ea As Double:    eb As Double:    ec As Double:    ed As Double:    ee As Double:    ef As Double
    fa As Double:    fb As Double:    fc As Double:    fd As Double:    fe As Double:    ff As Double
End Type
Public Type Matrix7 'Matrix 7x7
    aa As Double:    ab As Double:    ac As Double:    ad As Double:    ae As Double:    af As Double:    ag As Double
    ba As Double:    bb As Double:    bc As Double:    bd As Double:    be As Double:    bf As Double:    bg As Double
    ca As Double:    cb As Double:    cc As Double:    cd As Double:    ce As Double:    cf As Double:    cg As Double
    da As Double:    db As Double:    dc As Double:    dd As Double:    de As Double:    df As Double:    dg As Double
    ea As Double:    eb As Double:    ec As Double:    ed As Double:    ee As Double:    ef As Double:    eg As Double
    fa As Double:    fb As Double:    fc As Double:    fd As Double:    fe As Double:    ff As Double:    fg As Double
    ga As Double:    gb As Double:    gc As Double:    gd As Double:    ge As Double:    gf As Double:    gg As Double
End Type
Public Type Matrix8 'Matrix 8x8
    aa As Double:    ab As Double:    ac As Double:    ad As Double:    ae As Double:    af As Double:    ag As Double:    ah As Double
    ba As Double:    bb As Double:    bc As Double:    bd As Double:    be As Double:    bf As Double:    bg As Double:    bh As Double
    ca As Double:    cb As Double:    cc As Double:    cd As Double:    ce As Double:    cf As Double:    cg As Double:    ch As Double
    da As Double:    db As Double:    dc As Double:    dd As Double:    de As Double:    df As Double:    dg As Double:    dh As Double
    ea As Double:    eb As Double:    ec As Double:    ed As Double:    ee As Double:    ef As Double:    eg As Double:    eh As Double
    fa As Double:    fb As Double:    fc As Double:    fd As Double:    fe As Double:    ff As Double:    fg As Double:    fh As Double
    ga As Double:    gb As Double:    gc As Double:    gd As Double:    ge As Double:    gf As Double:    gg As Double:    gh As Double
    ha As Double:    hb As Double:    hc As Double:    hd As Double:    he As Double:    hf As Double:    hg As Double:    hh As Double
End Type
Public Type Matrix9 'Matrix 9x9 'here the variablename "if" occurs, this is OK in a type
    aa As Double:    ab As Double:    ac As Double:    ad As Double:    ae As Double:    af As Double:    ag As Double:    ah As Double:    ai As Double
    ba As Double:    bb As Double:    bc As Double:    bd As Double:    be As Double:    bf As Double:    bg As Double:    bh As Double:    bi As Double
    ca As Double:    cb As Double:    cc As Double:    cd As Double:    ce As Double:    cf As Double:    cg As Double:    ch As Double:    ci As Double
    da As Double:    db As Double:    dc As Double:    dd As Double:    de As Double:    df As Double:    dg As Double:    dh As Double:    di As Double
    ea As Double:    eb As Double:    ec As Double:    ed As Double:    ee As Double:    ef As Double:    eg As Double:    eh As Double:    ei As Double
    fa As Double:    fb As Double:    fc As Double:    fd As Double:    fe As Double:    ff As Double:    fg As Double:    fh As Double:    fi As Double
    ga As Double:    gb As Double:    gc As Double:    gd As Double:    ge As Double:    gf As Double:    gg As Double:    gh As Double:    gi As Double
    ha As Double:    hb As Double:    hc As Double:    hd As Double:    he As Double:    hf As Double:    hg As Double:    hh As Double:    hi As Double
    ia As Double:    ib As Double:    ic As Double:    id As Double:    ie As Double:    if As Double:    ig As Double:    ih As Double:    ii As Double
End Type
Public Type Matrix10 'Matrix 10x10
    aa As Double:    ab As Double:    ac As Double:    ad As Double:    ae As Double:    af As Double:    ag As Double:    ah As Double:    ai As Double:    aj As Double
    ba As Double:    bb As Double:    bc As Double:    bd As Double:    be As Double:    bf As Double:    bg As Double:    bh As Double:    bi As Double:    bj As Double
    ca As Double:    cb As Double:    cc As Double:    cd As Double:    ce As Double:    cf As Double:    cg As Double:    ch As Double:    ci As Double:    cj As Double
    da As Double:    db As Double:    dc As Double:    dd As Double:    de As Double:    df As Double:    dg As Double:    dh As Double:    di As Double:    dj As Double
    ea As Double:    eb As Double:    ec As Double:    ed As Double:    ee As Double:    ef As Double:    eg As Double:    eh As Double:    ei As Double:    ej As Double
    fa As Double:    fb As Double:    fc As Double:    fd As Double:    fe As Double:    ff As Double:    fg As Double:    fh As Double:    fi As Double:    fj As Double
    ga As Double:    gb As Double:    gc As Double:    gd As Double:    ge As Double:    gf As Double:    gg As Double:    gh As Double:    gi As Double:    gj As Double
    ha As Double:    hb As Double:    hc As Double:    hd As Double:    he As Double:    hf As Double:    hg As Double:    hh As Double:    hi As Double:    hj As Double
    ia As Double:    ib As Double:    ic As Double:    id As Double:    ie As Double:    if As Double:    ig As Double:    ih As Double:    ii As Double:    ij As Double
    ja As Double:    jb As Double:    jc As Double:    jd As Double:    je As Double:    jf As Double:    jg As Double:    jh As Double:    ji As Double:    jj As Double
End Type

'I ALSO WANT Singular Value Decomposition (SVD) !!!!!!!!!!!!!!!!
'WE ALSO NEED EIGENVECTOR AND EIGENVALUE!!!!!!!!!!!!!!!!

'Matrices with all variables are 1
Public Mat2_Ones As Matrix2
Public Mat3_Ones As Matrix3
Public Mat4_Ones As Matrix4
Public Mat5_Ones As Matrix5
Public Mat6_Ones As Matrix6
Public Mat7_Ones As Matrix7
Public Mat8_Ones As Matrix8
Public Mat9_Ones As Matrix9
Public Mat10_Ones As Matrix10

'Public Type MatrixT
'    a() As Double
'End Type
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef pDst As Any, ByRef pSrc As Any, ByVal nBytes As Long)
'Alias "cpymem" 'Zeile 106
Private Declare Sub PutOne Lib "msvbvm60.dll" Alias "PutMem8" (ByVal pDst As Long, Optional ByVal Src As Double = 1#)

Public Sub InitMatXOnes()
    Dim i As Long, n As Long, p As Long
    p = VarPtr(Mat2_Ones):    n = 2
    For i = 0 To n ^ 2 - 1:    PutOne p: p = p + 8:    Next
    p = VarPtr(Mat3_Ones):    n = n + 1
    For i = 0 To n ^ 2 - 1:    PutOne p: p = p + 8:    Next
    p = VarPtr(Mat4_Ones):    n = n + 1
    For i = 0 To n ^ 2 - 1:    PutOne p: p = p + 8:    Next
    p = VarPtr(Mat5_Ones):    n = n + 1
    For i = 0 To n ^ 2 - 1:    PutOne p: p = p + 8:    Next
    p = VarPtr(Mat6_Ones):    n = n + 1
    For i = 0 To n ^ 2 - 1:    PutOne p: p = p + 8:    Next
    p = VarPtr(Mat7_Ones):    n = n + 1
    For i = 0 To n ^ 2 - 1:    PutOne p: p = p + 8:    Next
    p = VarPtr(Mat8_Ones):    n = n + 1
    For i = 0 To n ^ 2 - 1:    PutOne p: p = p + 8:    Next
    p = VarPtr(Mat9_Ones):    n = n + 1
    For i = 0 To n ^ 2 - 1:    PutOne p: p = p + 8:    Next
    p = VarPtr(Mat10_Ones):    n = n + 1
    For i = 0 To n ^ 2 - 1:    PutOne p: p = p + 8:    Next
End Sub
'Einfache Matrix-Operationen
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
Public Function Mat7(aa As Double, ab As Double, ac As Double, ad As Double, ae As Double, af As Double, ag As Double, _
                     ba As Double, bb As Double, bc As Double, bd As Double, be As Double, bf As Double, bg As Double, _
                     ca As Double, cb As Double, cc As Double, cd As Double, ce As Double, cf As Double, cg As Double, _
                     da As Double, db As Double, dc As Double, dd As Double, de As Double, df As Double, dg As Double, _
                     ea As Double, eb As Double, ec As Double, ed As Double, ee As Double, ef As Double, eg As Double, _
                     fa As Double, fb As Double, fc As Double, fd As Double, fe As Double, ff As Double, fg As Double, _
                     ga As Double, gb As Double, gc As Double, gd As Double, ge As Double, gf As Double, gg As Double) As Matrix7
    'Erzeugt eine 7x7 Matrix
    With Mat7: .aa = aa: .ab = ab: .ac = ac: .ad = ad: .ae = ae: .af = af: .ag = ag
               .ba = ba: .bb = bb: .bc = bc: .bd = bd: .be = be: .bf = bf: .bg = bg
               .ca = ca: .cb = cb: .cc = cc: .cd = cd: .ce = ce: .cf = cf: .cg = cg
               .da = da: .db = db: .dc = dc: .dd = dd: .de = de: .df = df: .dg = dg
               .ea = ea: .eb = eb: .ec = ec: .ed = ed: .ee = ee: .ef = ef: .eg = eg
               .fa = fa: .fb = fb: .fc = fc: .fd = fd: .fe = fe: .ff = ff: .fg = fg
               .ga = ga: .gb = gb: .gc = gc: .gd = gd: .ge = ge: .gf = gf: .gg = gg:   End With
End Function
'not working, too much i guess
'Public Function Mat8(aa As Double, ab As Double, ac As Double, ad As Double, ae As Double, af As Double, ag As Double, ah As Double, _
'                     ba As Double, bb As Double, bc As Double, bd As Double, be As Double, bf As Double, bg As Double, bh As Double, _
'                     ca As Double, cb As Double, cc As Double, cd As Double, ce As Double, cf As Double, cg As Double, ch As Double, _
'                     da As Double, db As Double, dc As Double, dd As Double, de As Double, df As Double, dg As Double, dh As Double, _
'                     ea As Double, eb As Double, ec As Double, ed As Double, ee As Double, ef As Double, eg As Double, eh As Double, _
'                     fa As Double, fb As Double, fc As Double, fd As Double, fe As Double, ff As Double, fg As Double, fh As Double, _
'                     ga As Double, gb As Double, gc As Double, gd As Double, ge As Double, gf As Double, gg As Double, gh As Double, _
'                     ha As Double, hb As Double, hc As Double, hd As Double, he As Double, hf As Double, hg As Double, hh As Double) As Matrix8
'    'Erzeugt eine 8x8 Matrix
'    With Mat8: .aa = aa: .ab = ab: .ac = ac: .ad = ad: .ae = ae: .af = af: .ag = ag: .ah = ah
'               .ba = ba: .bb = bb: .bc = bc: .bd = bd: .be = be: .bf = bf: .bg = bg: .bh = bh
'               .ca = ca: .cb = cb: .cc = cc: .cd = cd: .ce = ce: .cf = cf: .cg = cg: .ch = ch
'               .da = da: .db = db: .dc = dc: .dd = dd: .de = de: .df = df: .dg = dg: .dh = dh
'               .ea = ea: .eb = eb: .ec = ec: .ed = ed: .ee = ee: .ef = ef: .eg = eg: .eh = eh
'               .fa = fa: .fb = fb: .fc = fc: .fd = fd: .fe = fe: .ff = ff: .fg = fg: .fh = fh
'               .ga = ga: .gb = gb: .gc = gc: .gd = gd: .ge = ge: .gf = gf: .gg = gg: .gh = gh
'               .ha = ha: .hb = hb: .hc = hc: .hd = hd: .he = he: .hf = hf: .hg = hg: .hh = hh
'    End With
'End Function
'OK folgende Möglichkeiten:
' * Paramarray, mit Array arbeiten jedes Element ist Variant also cpymem fällt aus
' * irgendwie mit Vektoren arbeiten aber wie Spalten- oder Zeilen-vektoren? OK Zeilenvektoren
Public Function Mat8(Row_a As Vector8, _
                     Row_b As Vector8, _
                     Row_c As Vector8, _
                     Row_d As Vector8, _
                     Row_e As Vector8, _
                     Row_f As Vector8, _
                     Row_g As Vector8, _
                     Row_h As Vector8) As Matrix8
    Dim p As Long: p = VarPtr(Mat8) 'size per row is 8 bytes per double and 8 variables = 8*8=64
    RtlMoveMemory ByVal p, Row_a, 64: p = p + 64
    RtlMoveMemory ByVal p, Row_b, 64: p = p + 64
    RtlMoveMemory ByVal p, Row_c, 64: p = p + 64
    RtlMoveMemory ByVal p, Row_d, 64: p = p + 64
    RtlMoveMemory ByVal p, Row_e, 64: p = p + 64
    RtlMoveMemory ByVal p, Row_f, 64: p = p + 64
    RtlMoveMemory ByVal p, Row_g, 64: p = p + 64
    RtlMoveMemory ByVal p, Row_h, 64
End Function
Public Function Mat9(Row_a As Vector9, _
                     Row_b As Vector9, _
                     Row_c As Vector9, _
                     Row_d As Vector9, _
                     Row_e As Vector9, _
                     Row_f As Vector9, _
                     Row_g As Vector9, _
                     Row_h As Vector9, _
                     Row_i As Vector9) As Matrix9
    Dim p As Long: p = VarPtr(Mat9) 'size per row is 8 bytes per double and 9 variables = 8*9=72
    RtlMoveMemory ByVal p, Row_a, 72: p = p + 72
    RtlMoveMemory ByVal p, Row_b, 72: p = p + 72
    RtlMoveMemory ByVal p, Row_c, 72: p = p + 72
    RtlMoveMemory ByVal p, Row_d, 72: p = p + 72
    RtlMoveMemory ByVal p, Row_e, 72: p = p + 72
    RtlMoveMemory ByVal p, Row_f, 72: p = p + 72
    RtlMoveMemory ByVal p, Row_g, 72: p = p + 72
    RtlMoveMemory ByVal p, Row_h, 72: p = p + 72
    RtlMoveMemory ByVal p, Row_i, 72
End Function
Public Function Mat10(Row_a As Vector10, _
                      Row_b As Vector10, _
                      Row_c As Vector10, _
                      Row_d As Vector10, _
                      Row_e As Vector10, _
                      Row_f As Vector10, _
                      Row_g As Vector10, _
                      Row_h As Vector10, _
                      Row_i As Vector10, _
                      Row_j As Vector10) As Matrix10
    Dim p As Long: p = VarPtr(Mat10) 'size per row is 8 bytes per double and 10 variables = 8*10=80
    RtlMoveMemory ByVal p, Row_a, 80: p = p + 80
    RtlMoveMemory ByVal p, Row_b, 80: p = p + 80
    RtlMoveMemory ByVal p, Row_c, 80: p = p + 80
    RtlMoveMemory ByVal p, Row_d, 80: p = p + 80
    RtlMoveMemory ByVal p, Row_e, 80: p = p + 80
    RtlMoveMemory ByVal p, Row_f, 80: p = p + 80
    RtlMoveMemory ByVal p, Row_g, 80: p = p + 80
    RtlMoveMemory ByVal p, Row_h, 80: p = p + 80
    RtlMoveMemory ByVal p, Row_i, 80
End Function


'Einheits-Matrizen erzeugen
Public Function Mat2_E() As Matrix2
    'Erzeugt eine 2x2 Einheits-Matrix, alle Variablen der Diagonale sind 1 alle Anderen sind 0
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
Public Function Mat7_E() As Matrix7
    'Erzeugt eine 7x7 Einheits-Matrix
    With Mat7_E: .aa = 1: .bb = 1: .cc = 1: .dd = 1: .ee = 1: .ff = 1: .gg = 1:    End With
End Function
Public Function Mat8_E() As Matrix8
    'Erzeugt eine 8x8 Einheits-Matrix
    With Mat8_E: .aa = 1: .bb = 1: .cc = 1: .dd = 1: .ee = 1: .ff = 1: .gg = 1: .hh = 1:    End With
End Function
Public Function Mat9_E() As Matrix9
    'Erzeugt eine 9x9 Einheits-Matrix
    With Mat9_E: .aa = 1: .bb = 1: .cc = 1: .dd = 1: .ee = 1: .ff = 1: .gg = 1: .hh = 1: .ii = 1:    End With
End Function
Public Function Mat10_E() As Matrix10
    'Erzeugt eine 10x10 Einheits-Matrix
    With Mat10_E: .aa = 1: .bb = 1: .cc = 1: .dd = 1: .ee = 1: .ff = 1: .gg = 1: .hh = 1: .ii = 1: .jj = 1:    End With
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
Public Function Matrix7_add(m1 As Matrix7, m2 As Matrix7) As Matrix7
    'Addiert 2 7x7 Matrizen
    With Matrix7_add:   .aa = m1.aa + m2.aa:    .ab = m1.ab + m2.ab:    .ac = m1.ac + m2.ac:    .ad = m1.ad + m2.ad:    .ae = m1.ae + m2.ae:    .af = m1.af + m2.af:    .ag = m1.ag + m2.ag
                        .ba = m1.ba + m2.ba:    .bb = m1.bb + m2.bb:    .bc = m1.bc + m2.bc:    .bd = m1.bd + m2.bd:    .be = m1.be + m2.be:    .bf = m1.bf + m2.bf:    .bg = m1.bg + m2.bg
                        .ca = m1.ca + m2.ca:    .cb = m1.cb + m2.cb:    .cc = m1.cc + m2.cc:    .cd = m1.cd + m2.cd:    .ce = m1.ce + m2.ce:    .cf = m1.cf + m2.cf:    .cg = m1.cg + m2.cg
                        .da = m1.da + m2.da:    .db = m1.db + m2.db:    .dc = m1.dc + m2.dc:    .dd = m1.dd + m2.dd:    .de = m1.de + m2.de:    .df = m1.df + m2.df:    .dg = m1.dg + m2.dg
                        .ea = m1.ea + m2.ea:    .eb = m1.eb + m2.eb:    .ec = m1.ec + m2.ec:    .ed = m1.ed + m2.ed:    .ee = m1.ee + m2.ee:    .ef = m1.ef + m2.ef:    .eg = m1.eg + m2.eg
                        .fa = m1.fa + m2.fa:    .fb = m1.fb + m2.fb:    .fc = m1.fc + m2.fc:    .fd = m1.fd + m2.fd:    .fe = m1.fe + m2.fe:    .ff = m1.ff + m2.ff:    .fg = m1.fg + m2.fg
                        .ga = m1.ga + m2.ga:    .gb = m1.gb + m2.gb:    .gc = m1.gc + m2.gc:    .gd = m1.gd + m2.gd:    .ge = m1.ge + m2.ge:    .gf = m1.gf + m2.gf:    .gg = m1.gg + m2.gg:    End With
End Function
Public Function Matrix8_add(m1 As Matrix8, m2 As Matrix8) As Matrix8
    'Addiert 2 8x8 Matrizen
    With Matrix8_add:   .aa = m1.aa + m2.aa:    .ab = m1.ab + m2.ab:    .ac = m1.ac + m2.ac:    .ad = m1.ad + m2.ad:    .ae = m1.ae + m2.ae:    .af = m1.af + m2.af:    .ag = m1.ag + m2.ag:    .ah = m1.ah + m2.ah
                        .ba = m1.ba + m2.ba:    .bb = m1.bb + m2.bb:    .bc = m1.bc + m2.bc:    .bd = m1.bd + m2.bd:    .be = m1.be + m2.be:    .bf = m1.bf + m2.bf:    .bg = m1.bg + m2.bg:    .bh = m1.bh + m2.bh
                        .ca = m1.ca + m2.ca:    .cb = m1.cb + m2.cb:    .cc = m1.cc + m2.cc:    .cd = m1.cd + m2.cd:    .ce = m1.ce + m2.ce:    .cf = m1.cf + m2.cf:    .cg = m1.cg + m2.cg:    .ch = m1.ch + m2.ch
                        .da = m1.da + m2.da:    .db = m1.db + m2.db:    .dc = m1.dc + m2.dc:    .dd = m1.dd + m2.dd:    .de = m1.de + m2.de:    .df = m1.df + m2.df:    .dg = m1.dg + m2.dg:    .dh = m1.dh + m2.dh
                        .ea = m1.ea + m2.ea:    .eb = m1.eb + m2.eb:    .ec = m1.ec + m2.ec:    .ed = m1.ed + m2.ed:    .ee = m1.ee + m2.ee:    .ef = m1.ef + m2.ef:    .eg = m1.eg + m2.eg:    .eh = m1.eh + m2.eh
                        .fa = m1.fa + m2.fa:    .fb = m1.fb + m2.fb:    .fc = m1.fc + m2.fc:    .fd = m1.fd + m2.fd:    .fe = m1.fe + m2.fe:    .ff = m1.ff + m2.ff:    .fg = m1.fg + m2.fg:    .fh = m1.fh + m2.fh
                        .ga = m1.ga + m2.ga:    .gb = m1.gb + m2.gb:    .gc = m1.gc + m2.gc:    .gd = m1.gd + m2.gd:    .ge = m1.ge + m2.ge:    .gf = m1.gf + m2.gf:    .gg = m1.gg + m2.gg:    .gh = m1.gh + m2.gh
                        .ha = m1.ha + m2.ha:    .hb = m1.hb + m2.hb:    .hc = m1.hc + m2.hc:    .hd = m1.hd + m2.hd:    .he = m1.he + m2.he:    .hf = m1.hf + m2.hf:    .hg = m1.hg + m2.hg:    .hh = m1.hh + m2.hh:    End With
End Function
Public Function Matrix9_add(m1 As Matrix9, m2 As Matrix9) As Matrix9
    'Addiert 2 9x9 Matrizen
    With Matrix9_add:   .aa = m1.aa + m2.aa:    .ab = m1.ab + m2.ab:    .ac = m1.ac + m2.ac:    .ad = m1.ad + m2.ad:    .ae = m1.ae + m2.ae:    .af = m1.af + m2.af:    .ag = m1.ag + m2.ag:    .ah = m1.ah + m2.ah:    .ai = m1.ai + m2.ai
                        .ba = m1.ba + m2.ba:    .bb = m1.bb + m2.bb:    .bc = m1.bc + m2.bc:    .bd = m1.bd + m2.bd:    .be = m1.be + m2.be:    .bf = m1.bf + m2.bf:    .bg = m1.bg + m2.bg:    .bh = m1.bh + m2.bh:    .bi = m1.bi + m2.bi
                        .ca = m1.ca + m2.ca:    .cb = m1.cb + m2.cb:    .cc = m1.cc + m2.cc:    .cd = m1.cd + m2.cd:    .ce = m1.ce + m2.ce:    .cf = m1.cf + m2.cf:    .cg = m1.cg + m2.cg:    .ch = m1.ch + m2.ch:    .ci = m1.ci + m2.ci
                        .da = m1.da + m2.da:    .db = m1.db + m2.db:    .dc = m1.dc + m2.dc:    .dd = m1.dd + m2.dd:    .de = m1.de + m2.de:    .df = m1.df + m2.df:    .dg = m1.dg + m2.dg:    .dh = m1.dh + m2.dh:    .di = m1.di + m2.di
                        .ea = m1.ea + m2.ea:    .eb = m1.eb + m2.eb:    .ec = m1.ec + m2.ec:    .ed = m1.ed + m2.ed:    .ee = m1.ee + m2.ee:    .ef = m1.ef + m2.ef:    .eg = m1.eg + m2.eg:    .eh = m1.eh + m2.eh:    .ei = m1.ei + m2.ei
                        .fa = m1.fa + m2.fa:    .fb = m1.fb + m2.fb:    .fc = m1.fc + m2.fc:    .fd = m1.fd + m2.fd:    .fe = m1.fe + m2.fe:    .ff = m1.ff + m2.ff:    .fg = m1.fg + m2.fg:    .fh = m1.fh + m2.fh:    .fi = m1.fi + m2.fi
                        .ga = m1.ga + m2.ga:    .gb = m1.gb + m2.gb:    .gc = m1.gc + m2.gc:    .gd = m1.gd + m2.gd:    .ge = m1.ge + m2.ge:    .gf = m1.gf + m2.gf:    .gg = m1.gg + m2.gg:    .gh = m1.gh + m2.gh:    .gi = m1.gi + m2.gi
                        .ha = m1.ha + m2.ha:    .hb = m1.hb + m2.hb:    .hc = m1.hc + m2.hc:    .hd = m1.hd + m2.hd:    .he = m1.he + m2.he:    .hf = m1.hf + m2.hf:    .hg = m1.hg + m2.hg:    .hh = m1.hh + m2.hh:    .hi = m1.hi + m2.hi
                        .ia = m1.ia + m2.ia:    .ib = m1.ib + m2.ib:    .ic = m1.ic + m2.ic:    .id = m1.id + m2.id:    .ie = m1.ie + m2.ie:    .if = m1.if + m2.if:    .ig = m1.ig + m2.ig:    .ih = m1.ih + m2.ih:    .ii = m1.ii + m2.ii:    End With
End Function
Public Function Matrix10_add(m1 As Matrix10, m2 As Matrix10) As Matrix10
    'Addiert 2 10x10 Matrizen
    With Matrix10_add:   .aa = m1.aa + m2.aa:    .ab = m1.ab + m2.ab:    .ac = m1.ac + m2.ac:    .ad = m1.ad + m2.ad:    .ae = m1.ae + m2.ae:    .af = m1.af + m2.af:    .ag = m1.ag + m2.ag:    .ah = m1.ah + m2.ah:    .ai = m1.ai + m2.ai:    .aj = m1.aj + m2.aj
                         .ba = m1.ba + m2.ba:    .bb = m1.bb + m2.bb:    .bc = m1.bc + m2.bc:    .bd = m1.bd + m2.bd:    .be = m1.be + m2.be:    .bf = m1.bf + m2.bf:    .bg = m1.bg + m2.bg:    .bh = m1.bh + m2.bh:    .bi = m1.bi + m2.bi:    .bj = m1.bj + m2.bj
                         .ca = m1.ca + m2.ca:    .cb = m1.cb + m2.cb:    .cc = m1.cc + m2.cc:    .cd = m1.cd + m2.cd:    .ce = m1.ce + m2.ce:    .cf = m1.cf + m2.cf:    .cg = m1.cg + m2.cg:    .ch = m1.ch + m2.ch:    .ci = m1.ci + m2.ci:    .cj = m1.cj + m2.cj
                         .da = m1.da + m2.da:    .db = m1.db + m2.db:    .dc = m1.dc + m2.dc:    .dd = m1.dd + m2.dd:    .de = m1.de + m2.de:    .df = m1.df + m2.df:    .dg = m1.dg + m2.dg:    .dh = m1.dh + m2.dh:    .di = m1.di + m2.di:    .dj = m1.dj + m2.dj
                         .ea = m1.ea + m2.ea:    .eb = m1.eb + m2.eb:    .ec = m1.ec + m2.ec:    .ed = m1.ed + m2.ed:    .ee = m1.ee + m2.ee:    .ef = m1.ef + m2.ef:    .eg = m1.eg + m2.eg:    .eh = m1.eh + m2.eh:    .ei = m1.ei + m2.ei:    .ej = m1.ej + m2.ej
                         .fa = m1.fa + m2.fa:    .fb = m1.fb + m2.fb:    .fc = m1.fc + m2.fc:    .fd = m1.fd + m2.fd:    .fe = m1.fe + m2.fe:    .ff = m1.ff + m2.ff:    .fg = m1.fg + m2.fg:    .fh = m1.fh + m2.fh:    .fi = m1.fi + m2.fi:    .fj = m1.fj + m2.fj
                         .ga = m1.ga + m2.ga:    .gb = m1.gb + m2.gb:    .gc = m1.gc + m2.gc:    .gd = m1.gd + m2.gd:    .ge = m1.ge + m2.ge:    .gf = m1.gf + m2.gf:    .gg = m1.gg + m2.gg:    .gh = m1.gh + m2.gh:    .gi = m1.gi + m2.gi:    .gj = m1.gj + m2.gj
                         .ha = m1.ha + m2.ha:    .hb = m1.hb + m2.hb:    .hc = m1.hc + m2.hc:    .hd = m1.hd + m2.hd:    .he = m1.he + m2.he:    .hf = m1.hf + m2.hf:    .hg = m1.hg + m2.hg:    .hh = m1.hh + m2.hh:    .hi = m1.hi + m2.hi:    .hj = m1.hj + m2.hj
                         .ia = m1.ia + m2.ia:    .ib = m1.ib + m2.ib:    .ic = m1.ic + m2.ic:    .id = m1.id + m2.id:    .ie = m1.ie + m2.ie:    .if = m1.if + m2.if:    .ig = m1.ig + m2.ig:    .ih = m1.ih + m2.ih:    .ii = m1.ii + m2.ii:    .ij = m1.ij + m2.ij
                         .ja = m1.ja + m2.ja:    .jb = m1.jb + m2.jb:    .jc = m1.jc + m2.jc:    .jd = m1.jd + m2.jd:    .je = m1.je + m2.je:    .jf = m1.jf + m2.jf:    .jg = m1.jg + m2.jg:    .jh = m1.jh + m2.jh:    .ji = m1.ji + m2.ji:    .jj = m1.jj + m2.jj:    End With
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
Public Function Matrix7_sub(m1 As Matrix7, m2 As Matrix7) As Matrix7
    'Subtrahiert eine 7x7 Matrix m2 von Matrix m1
    With Matrix7_sub:   .aa = m1.aa - m2.aa:    .ab = m1.ab - m2.ab:    .ac = m1.ac - m2.ac:    .ad = m1.ad - m2.ad:    .ae = m1.ae - m2.ae:    .af = m1.af - m2.af:    .ag = m1.ag - m2.ag
                        .ba = m1.ba - m2.ba:    .bb = m1.bb - m2.bb:    .bc = m1.bc - m2.bc:    .bd = m1.bd - m2.bd:    .be = m1.be - m2.be:    .bf = m1.bf - m2.bf:    .bg = m1.bg - m2.bg
                        .ca = m1.ca - m2.ca:    .cb = m1.cb - m2.cb:    .cc = m1.cc - m2.cc:    .cd = m1.cd - m2.cd:    .ce = m1.ce - m2.ce:    .cf = m1.cf - m2.cf:    .cg = m1.cg - m2.cg
                        .da = m1.da - m2.da:    .db = m1.db - m2.db:    .dc = m1.dc - m2.dc:    .dd = m1.dd - m2.dd:    .de = m1.de - m2.de:    .df = m1.df - m2.df:    .dg = m1.dg - m2.dg
                        .ea = m1.ea - m2.ea:    .eb = m1.eb - m2.eb:    .ec = m1.ec - m2.ec:    .ed = m1.ed - m2.ed:    .ee = m1.ee - m2.ee:    .ef = m1.ef - m2.ef:    .eg = m1.eg - m2.eg
                        .fa = m1.fa - m2.fa:    .fb = m1.fb - m2.fb:    .fc = m1.fc - m2.fc:    .fd = m1.fd - m2.fd:    .fe = m1.fe - m2.fe:    .ff = m1.ff - m2.ff:    .fg = m1.fg - m2.fg
                        .ga = m1.ga - m2.ga:    .gb = m1.gb - m2.gb:    .gc = m1.gc - m2.gc:    .gd = m1.gd - m2.gd:    .ge = m1.ge - m2.ge:    .gf = m1.gf - m2.gf:    .gg = m1.gg - m2.gg:    End With
End Function
Public Function Matrix8_sub(m1 As Matrix8, m2 As Matrix8) As Matrix8
    'Subtrahiert eine 8x8 Matrix m2 von Matrix m1
    With Matrix8_sub:   .aa = m1.aa - m2.aa:    .ab = m1.ab - m2.ab:    .ac = m1.ac - m2.ac:    .ad = m1.ad - m2.ad:    .ae = m1.ae - m2.ae:    .af = m1.af - m2.af:    .ag = m1.ag - m2.ag:    .ah = m1.ah - m2.ah
                        .ba = m1.ba - m2.ba:    .bb = m1.bb - m2.bb:    .bc = m1.bc - m2.bc:    .bd = m1.bd - m2.bd:    .be = m1.be - m2.be:    .bf = m1.bf - m2.bf:    .bg = m1.bg - m2.bg:    .bh = m1.bh - m2.bh
                        .ca = m1.ca - m2.ca:    .cb = m1.cb - m2.cb:    .cc = m1.cc - m2.cc:    .cd = m1.cd - m2.cd:    .ce = m1.ce - m2.ce:    .cf = m1.cf - m2.cf:    .cg = m1.cg - m2.cg:    .ch = m1.ch - m2.ch
                        .da = m1.da - m2.da:    .db = m1.db - m2.db:    .dc = m1.dc - m2.dc:    .dd = m1.dd - m2.dd:    .de = m1.de - m2.de:    .df = m1.df - m2.df:    .dg = m1.dg - m2.dg:    .dh = m1.dh - m2.dh
                        .ea = m1.ea - m2.ea:    .eb = m1.eb - m2.eb:    .ec = m1.ec - m2.ec:    .ed = m1.ed - m2.ed:    .ee = m1.ee - m2.ee:    .ef = m1.ef - m2.ef:    .eg = m1.eg - m2.eg:    .eh = m1.eh - m2.eh
                        .fa = m1.fa - m2.fa:    .fb = m1.fb - m2.fb:    .fc = m1.fc - m2.fc:    .fd = m1.fd - m2.fd:    .fe = m1.fe - m2.fe:    .ff = m1.ff - m2.ff:    .fg = m1.fg - m2.fg:    .fh = m1.fh - m2.fh
                        .ga = m1.ga - m2.ga:    .gb = m1.gb - m2.gb:    .gc = m1.gc - m2.gc:    .gd = m1.gd - m2.gd:    .ge = m1.ge - m2.ge:    .gf = m1.gf - m2.gf:    .gg = m1.gg - m2.gg:    .gh = m1.gh - m2.gh
                        .ha = m1.ha - m2.ha:    .hb = m1.hb - m2.hb:    .hc = m1.hc - m2.hc:    .hd = m1.hd - m2.hd:    .he = m1.he - m2.he:    .hf = m1.hf - m2.hf:    .hg = m1.hg - m2.hg:    .hh = m1.hh - m2.hh:    End With
End Function
Public Function Matrix9_sub(m1 As Matrix9, m2 As Matrix9) As Matrix9
    'Subtrahiert eine 9x9 Matrix m2 von Matrix m1
    With Matrix9_sub:   .aa = m1.aa - m2.aa:    .ab = m1.ab - m2.ab:    .ac = m1.ac - m2.ac:    .ad = m1.ad - m2.ad:    .ae = m1.ae - m2.ae:    .af = m1.af - m2.af:    .ag = m1.ag - m2.ag:    .ah = m1.ah - m2.ah:    .ai = m1.ai - m2.ai
                        .ba = m1.ba - m2.ba:    .bb = m1.bb - m2.bb:    .bc = m1.bc - m2.bc:    .bd = m1.bd - m2.bd:    .be = m1.be - m2.be:    .bf = m1.bf - m2.bf:    .bg = m1.bg - m2.bg:    .bh = m1.bh - m2.bh:    .bi = m1.bi - m2.bi
                        .ca = m1.ca - m2.ca:    .cb = m1.cb - m2.cb:    .cc = m1.cc - m2.cc:    .cd = m1.cd - m2.cd:    .ce = m1.ce - m2.ce:    .cf = m1.cf - m2.cf:    .cg = m1.cg - m2.cg:    .ch = m1.ch - m2.ch:    .ci = m1.ci - m2.ci
                        .da = m1.da - m2.da:    .db = m1.db - m2.db:    .dc = m1.dc - m2.dc:    .dd = m1.dd - m2.dd:    .de = m1.de - m2.de:    .df = m1.df - m2.df:    .dg = m1.dg - m2.dg:    .dh = m1.dh - m2.dh:    .di = m1.di - m2.di
                        .ea = m1.ea - m2.ea:    .eb = m1.eb - m2.eb:    .ec = m1.ec - m2.ec:    .ed = m1.ed - m2.ed:    .ee = m1.ee - m2.ee:    .ef = m1.ef - m2.ef:    .eg = m1.eg - m2.eg:    .eh = m1.eh - m2.eh:    .ei = m1.ei - m2.ei
                        .fa = m1.fa - m2.fa:    .fb = m1.fb - m2.fb:    .fc = m1.fc - m2.fc:    .fd = m1.fd - m2.fd:    .fe = m1.fe - m2.fe:    .ff = m1.ff - m2.ff:    .fg = m1.fg - m2.fg:    .fh = m1.fh - m2.fh:    .fi = m1.fi - m2.fi
                        .ga = m1.ga - m2.ga:    .gb = m1.gb - m2.gb:    .gc = m1.gc - m2.gc:    .gd = m1.gd - m2.gd:    .ge = m1.ge - m2.ge:    .gf = m1.gf - m2.gf:    .gg = m1.gg - m2.gg:    .gh = m1.gh - m2.gh:    .gi = m1.gi - m2.gi
                        .ha = m1.ha - m2.ha:    .hb = m1.hb - m2.hb:    .hc = m1.hc - m2.hc:    .hd = m1.hd - m2.hd:    .he = m1.he - m2.he:    .hf = m1.hf - m2.hf:    .hg = m1.hg - m2.hg:    .hh = m1.hh - m2.hh:    .hi = m1.hi - m2.hi
                        .ia = m1.ia - m2.ia:    .ib = m1.ib - m2.ib:    .ic = m1.ic - m2.ic:    .id = m1.id - m2.id:    .ie = m1.ie - m2.ie:    .if = m1.if - m2.if:    .ig = m1.ig - m2.ig:    .ih = m1.ih - m2.ih:    .ii = m1.ii - m2.ii:    End With
End Function
Public Function Matrix10_sub(m1 As Matrix10, m2 As Matrix10) As Matrix10
    'Subtrahiert eine 9x9 Matrix m2 von Matrix m1
    With Matrix10_sub:   .aa = m1.aa - m2.aa:    .ab = m1.ab - m2.ab:    .ac = m1.ac - m2.ac:    .ad = m1.ad - m2.ad:    .ae = m1.ae - m2.ae:    .af = m1.af - m2.af:    .ag = m1.ag - m2.ag:    .ah = m1.ah - m2.ah:    .ai = m1.ai - m2.ai:    .aj = m1.aj - m2.aj
                         .ba = m1.ba - m2.ba:    .bb = m1.bb - m2.bb:    .bc = m1.bc - m2.bc:    .bd = m1.bd - m2.bd:    .be = m1.be - m2.be:    .bf = m1.bf - m2.bf:    .bg = m1.bg - m2.bg:    .bh = m1.bh - m2.bh:    .bi = m1.bi - m2.bi:    .bj = m1.bj - m2.bj
                         .ca = m1.ca - m2.ca:    .cb = m1.cb - m2.cb:    .cc = m1.cc - m2.cc:    .cd = m1.cd - m2.cd:    .ce = m1.ce - m2.ce:    .cf = m1.cf - m2.cf:    .cg = m1.cg - m2.cg:    .ch = m1.ch - m2.ch:    .ci = m1.ci - m2.ci:    .cj = m1.cj - m2.cj
                         .da = m1.da - m2.da:    .db = m1.db - m2.db:    .dc = m1.dc - m2.dc:    .dd = m1.dd - m2.dd:    .de = m1.de - m2.de:    .df = m1.df - m2.df:    .dg = m1.dg - m2.dg:    .dh = m1.dh - m2.dh:    .di = m1.di - m2.di:    .dj = m1.dj - m2.dj
                         .ea = m1.ea - m2.ea:    .eb = m1.eb - m2.eb:    .ec = m1.ec - m2.ec:    .ed = m1.ed - m2.ed:    .ee = m1.ee - m2.ee:    .ef = m1.ef - m2.ef:    .eg = m1.eg - m2.eg:    .eh = m1.eh - m2.eh:    .ei = m1.ei - m2.ei:    .ej = m1.ej - m2.ej
                         .fa = m1.fa - m2.fa:    .fb = m1.fb - m2.fb:    .fc = m1.fc - m2.fc:    .fd = m1.fd - m2.fd:    .fe = m1.fe - m2.fe:    .ff = m1.ff - m2.ff:    .fg = m1.fg - m2.fg:    .fh = m1.fh - m2.fh:    .fi = m1.fi - m2.fi:    .fj = m1.fj - m2.fj
                         .ga = m1.ga - m2.ga:    .gb = m1.gb - m2.gb:    .gc = m1.gc - m2.gc:    .gd = m1.gd - m2.gd:    .ge = m1.ge - m2.ge:    .gf = m1.gf - m2.gf:    .gg = m1.gg - m2.gg:    .gh = m1.gh - m2.gh:    .gi = m1.gi - m2.gi:    .gj = m1.gj - m2.gj
                         .ha = m1.ha - m2.ha:    .hb = m1.hb - m2.hb:    .hc = m1.hc - m2.hc:    .hd = m1.hd - m2.hd:    .he = m1.he - m2.he:    .hf = m1.hf - m2.hf:    .hg = m1.hg - m2.hg:    .hh = m1.hh - m2.hh:    .hi = m1.hi - m2.hi:    .hj = m1.hj - m2.hj
                         .ia = m1.ia - m2.ia:    .ib = m1.ib - m2.ib:    .ic = m1.ic - m2.ic:    .id = m1.id - m2.id:    .ie = m1.ie - m2.ie:    .if = m1.if - m2.if:    .ig = m1.ig - m2.ig:    .ih = m1.ih - m2.ih:    .ii = m1.ii - m2.ii:    .ij = m1.ij - m2.ij
                         .ja = m1.ja - m2.ja:    .jb = m1.jb - m2.jb:    .jc = m1.jc - m2.jc:    .jd = m1.jd - m2.jd:    .je = m1.je - m2.je:    .jf = m1.jf - m2.jf:    .jg = m1.jg - m2.jg:    .jh = m1.jh - m2.jh:    .ji = m1.ji - m2.ji:    .jj = m1.jj - m2.jj:    End With
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
Public Function Matrix7_smul(m As Matrix7, s As Double) As Matrix7
    'Multipliziert eine 7x7 Matrix mit einem Skalar
    With Matrix7_smul:  .aa = m.aa * s: .ab = m.ab * s: .ac = m.ac * s: .ad = m.ad * s: .ae = m.ae * s: .af = m.af * s: .ag = m.ag * s
                        .ba = m.ba * s: .bb = m.bb * s: .bc = m.bc * s: .bd = m.bd * s: .be = m.be * s: .bf = m.bf * s: .bg = m.bg * s
                        .ca = m.ca * s: .cb = m.cb * s: .cc = m.cc * s: .cd = m.cd * s: .ce = m.ce * s: .cf = m.cf * s: .cg = m.cg * s
                        .da = m.da * s: .db = m.db * s: .dc = m.dc * s: .dd = m.dd * s: .de = m.de * s: .df = m.df * s: .dg = m.dg * s
                        .ea = m.ea * s: .eb = m.eb * s: .ec = m.ec * s: .ed = m.ed * s: .ee = m.ee * s: .ef = m.ef * s: .eg = m.eg * s
                        .fa = m.fa * s: .fb = m.fb * s: .fc = m.fc * s: .fd = m.fd * s: .fe = m.fe * s: .ff = m.ff * s: .fg = m.fg * s
                        .ga = m.ga * s: .gb = m.gb * s: .gc = m.gc * s: .gd = m.gd * s: .ge = m.ge * s: .gf = m.gf * s: .gg = m.gg * s:    End With
End Function
Public Function Matrix8_smul(m As Matrix8, s As Double) As Matrix8
    'Multipliziert eine 8x8 Matrix mit einem Skalar
    With Matrix8_smul:  .aa = m.aa * s: .ab = m.ab * s: .ac = m.ac * s: .ad = m.ad * s: .ae = m.ae * s: .af = m.af * s: .ag = m.ag * s: .ah = m.ah * s
                        .ba = m.ba * s: .bb = m.bb * s: .bc = m.bc * s: .bd = m.bd * s: .be = m.be * s: .bf = m.bf * s: .bg = m.bg * s: .bh = m.bh * s
                        .ca = m.ca * s: .cb = m.cb * s: .cc = m.cc * s: .cd = m.cd * s: .ce = m.ce * s: .cf = m.cf * s: .cg = m.cg * s: .ch = m.ch * s
                        .da = m.da * s: .db = m.db * s: .dc = m.dc * s: .dd = m.dd * s: .de = m.de * s: .df = m.df * s: .dg = m.dg * s: .dh = m.dh * s
                        .ea = m.ea * s: .eb = m.eb * s: .ec = m.ec * s: .ed = m.ed * s: .ee = m.ee * s: .ef = m.ef * s: .eg = m.eg * s: .eh = m.eh * s
                        .fa = m.fa * s: .fb = m.fb * s: .fc = m.fc * s: .fd = m.fd * s: .fe = m.fe * s: .ff = m.ff * s: .fg = m.fg * s: .fh = m.fh * s
                        .ga = m.ga * s: .gb = m.gb * s: .gc = m.gc * s: .gd = m.gd * s: .ge = m.ge * s: .gf = m.gf * s: .gg = m.gg * s: .gh = m.gh * s
                        .ha = m.ha * s: .hb = m.hb * s: .hc = m.hc * s: .hd = m.hd * s: .he = m.he * s: .hf = m.hf * s: .hg = m.hg * s: .hh = m.hh * s:    End With
End Function
Public Function Matrix9_smul(m As Matrix9, s As Double) As Matrix9
    'Multipliziert eine 8x8 Matrix mit einem Skalar
    With Matrix9_smul:  .aa = m.aa * s: .ab = m.ab * s: .ac = m.ac * s: .ad = m.ad * s: .ae = m.ae * s: .af = m.af * s: .ag = m.ag * s: .ah = m.ah * s: .ai = m.ai * s
                        .ba = m.ba * s: .bb = m.bb * s: .bc = m.bc * s: .bd = m.bd * s: .be = m.be * s: .bf = m.bf * s: .bg = m.bg * s: .bh = m.bh * s: .bi = m.bi * s
                        .ca = m.ca * s: .cb = m.cb * s: .cc = m.cc * s: .cd = m.cd * s: .ce = m.ce * s: .cf = m.cf * s: .cg = m.cg * s: .ch = m.ch * s: .ci = m.ci * s
                        .da = m.da * s: .db = m.db * s: .dc = m.dc * s: .dd = m.dd * s: .de = m.de * s: .df = m.df * s: .dg = m.dg * s: .dh = m.dh * s: .di = m.di * s
                        .ea = m.ea * s: .eb = m.eb * s: .ec = m.ec * s: .ed = m.ed * s: .ee = m.ee * s: .ef = m.ef * s: .eg = m.eg * s: .eh = m.eh * s: .ei = m.ei * s
                        .fa = m.fa * s: .fb = m.fb * s: .fc = m.fc * s: .fd = m.fd * s: .fe = m.fe * s: .ff = m.ff * s: .fg = m.fg * s: .fh = m.fh * s: .fi = m.fi * s
                        .ga = m.ga * s: .gb = m.gb * s: .gc = m.gc * s: .gd = m.gd * s: .ge = m.ge * s: .gf = m.gf * s: .gg = m.gg * s: .gh = m.gh * s: .gi = m.gi * s
                        .ha = m.ha * s: .hb = m.hb * s: .hc = m.hc * s: .hd = m.hd * s: .he = m.he * s: .hf = m.hf * s: .hg = m.hg * s: .hh = m.hh * s: .hi = m.hi * s
                        .ia = m.ia * s: .ib = m.ib * s: .ic = m.ic * s: .id = m.id * s: .ie = m.ie * s: .if = m.if * s: .ig = m.ig * s: .ih = m.ih * s: .ii = m.ii * s:    End With
End Function
Public Function Matrix10_smul(m As Matrix10, s As Double) As Matrix10
    'Multipliziert eine 8x8 Matrix mit einem Skalar
    With Matrix10_smul:  .aa = m.aa * s: .ab = m.ab * s: .ac = m.ac * s: .ad = m.ad * s: .ae = m.ae * s: .af = m.af * s: .ag = m.ag * s: .ah = m.ah * s: .ai = m.ai * s: .aj = m.aj * s
                         .ba = m.ba * s: .bb = m.bb * s: .bc = m.bc * s: .bd = m.bd * s: .be = m.be * s: .bf = m.bf * s: .bg = m.bg * s: .bh = m.bh * s: .bi = m.bi * s: .bj = m.bj * s
                         .ca = m.ca * s: .cb = m.cb * s: .cc = m.cc * s: .cd = m.cd * s: .ce = m.ce * s: .cf = m.cf * s: .cg = m.cg * s: .ch = m.ch * s: .ci = m.ci * s: .cj = m.cj * s
                         .da = m.da * s: .db = m.db * s: .dc = m.dc * s: .dd = m.dd * s: .de = m.de * s: .df = m.df * s: .dg = m.dg * s: .dh = m.dh * s: .di = m.di * s: .dj = m.dj * s
                         .ea = m.ea * s: .eb = m.eb * s: .ec = m.ec * s: .ed = m.ed * s: .ee = m.ee * s: .ef = m.ef * s: .eg = m.eg * s: .eh = m.eh * s: .ei = m.ei * s: .ej = m.ej * s
                         .fa = m.fa * s: .fb = m.fb * s: .fc = m.fc * s: .fd = m.fd * s: .fe = m.fe * s: .ff = m.ff * s: .fg = m.fg * s: .fh = m.fh * s: .fi = m.fi * s: .fj = m.fj * s
                         .ga = m.ga * s: .gb = m.gb * s: .gc = m.gc * s: .gd = m.gd * s: .ge = m.ge * s: .gf = m.gf * s: .gg = m.gg * s: .gh = m.gh * s: .gi = m.gi * s: .gj = m.gj * s
                         .ha = m.ha * s: .hb = m.hb * s: .hc = m.hc * s: .hd = m.hd * s: .he = m.he * s: .hf = m.hf * s: .hg = m.hg * s: .hh = m.hh * s: .hi = m.hi * s: .hj = m.hj * s
                         .ia = m.ia * s: .ib = m.ib * s: .ic = m.ic * s: .id = m.id * s: .ie = m.ie * s: .if = m.if * s: .ig = m.ig * s: .ih = m.ih * s: .ii = m.ii * s: .ij = m.ij * s
                         .ja = m.ja * s: .jb = m.jb * s: .jc = m.jc * s: .jd = m.jd * s: .je = m.je * s: .jf = m.jf * s: .jg = m.jg * s: .jh = m.jh * s: .ji = m.ji * s: .jj = m.jj * s:    End With
End Function

'Matrizen-Multiplikation
Public Function Matrix2_mul(m1 As Matrix2, m2 As Matrix2) As Matrix2
    'Multipliziert eine 2x2 Matrix m1 mit einer Matrix m2
    With Matrix2_mul:   .aa = m1.aa * m2.aa + m1.ab * m2.ba:    .ab = m1.aa * m2.ab + m1.ab * m2.bb
                        .ba = m1.ba * m2.aa + m1.bb * m2.ba:    .bb = m1.ba * m2.ab + m1.bb * m2.bb
    End With
End Function

Public Function Matrix3_mul(m1 As Matrix3, m2 As Matrix3) As Matrix3
    'Multipliziert eine 3x3 Matrix m1 mit einer Matrix m2
    With Matrix3_mul:   .aa = m1.aa * m2.aa + m1.ab * m2.ba + m1.ac * m2.ca:    .ab = m1.aa * m2.ab + m1.ab * m2.bb + m1.ac * m2.cb:    .ac = m1.aa * m2.ac + m1.ab * m2.bc + m1.ac * m2.cc
                        .ba = m1.ba * m2.aa + m1.bb * m2.ba + m1.bc * m2.ca:    .bb = m1.ba * m2.ab + m1.bb * m2.bb + m1.bc * m2.cb:    .bc = m1.ba * m2.ac + m1.bb * m2.bc + m1.bc * m2.cc
                        .ca = m1.ca * m2.aa + m1.cb * m2.ba + m1.cc * m2.ca:    .cb = m1.ca * m2.ab + m1.cb * m2.bb + m1.cc * m2.cb:    .cc = m1.ca * m2.ac + m1.cb * m2.bc + m1.cc * m2.cc
    End With
End Function
Public Function Matrix4_mul(m1 As Matrix4, m2 As Matrix4) As Matrix4
    'Multipliziert eine 4x4 Matrix m1 mit einer Matrix m2
    With Matrix4_mul:   .aa = m1.aa * m2.aa + m1.ab * m2.ba + m1.ac * m2.ca + m1.ad * m2.da:    .ab = m1.aa * m2.ab + m1.ab * m2.bb + m1.ac * m2.cb + m1.ad * m2.db:    .ac = m1.aa * m2.ac + m1.ab * m2.bc + m1.ac * m2.cc + m1.ad * m2.dc:    .ad = m1.aa * m2.ad + m1.ab * m2.bd + m1.ac * m2.cd + m1.ad * m2.dd
                        .ba = m1.ba * m2.aa + m1.bb * m2.ba + m1.bc * m2.ca + m1.bd * m2.da:    .bb = m1.ba * m2.ab + m1.bb * m2.bb + m1.bc * m2.cb + m1.bd * m2.db:    .bc = m1.ba * m2.ac + m1.bb * m2.bc + m1.bc * m2.cc + m1.bd * m2.dc:    .bd = m1.ba * m2.ad + m1.bb * m2.bd + m1.bc * m2.cd + m1.bd * m2.dd
                        .ca = m1.ca * m2.aa + m1.cb * m2.ba + m1.cc * m2.ca + m1.cd * m2.da:    .cb = m1.ca * m2.ab + m1.cb * m2.bb + m1.cc * m2.cb + m1.cd * m2.db:    .cc = m1.ca * m2.ac + m1.cb * m2.bc + m1.cc * m2.cc + m1.cd * m2.dc:    .cd = m1.ca * m2.ad + m1.cb * m2.bd + m1.cc * m2.cd + m1.cd * m2.dd
                        .da = m1.da * m2.aa + m1.db * m2.ba + m1.dc * m2.ca + m1.dd * m2.da:    .db = m1.da * m2.ab + m1.db * m2.bb + m1.dc * m2.cb + m1.dd * m2.db:    .dc = m1.da * m2.ac + m1.db * m2.bc + m1.dc * m2.cc + m1.dd * m2.dc:    .dd = m1.da * m2.ad + m1.db * m2.bd + m1.dc * m2.cd + m1.dd * m2.dd
    End With
End Function
Public Function Matrix5_mul(m1 As Matrix5, m2 As Matrix5) As Matrix5
    'Multipliziert eine 5x5 Matrix m1 mit einer Matrix m2
    With Matrix5_mul:   .aa = m1.aa * m2.aa + m1.ab * m2.ba + m1.ac * m2.ca + m1.ad * m2.da + m1.ae * m2.ea:    .ab = m1.aa * m2.ab + m1.ab * m2.bb + m1.ac * m2.cb + m1.ad * m2.db + m1.ad * m2.eb:    .ac = m1.aa * m2.ac + m1.ab * m2.bc + m1.ac * m2.cc + m1.ad * m2.dc + m1.ae * m2.ec:    .ad = m1.aa * m2.ad + m1.ab * m2.bd + m1.ac * m2.cd + m1.ad * m2.dd + m1.ae * m2.ed:    .ae = m1.aa * m2.ae + m1.ab * m2.be + m1.ac * m2.ce + m1.ad * m2.de + m1.ae * m2.ee
                        .ba = m1.ba * m2.aa + m1.bb * m2.ba + m1.bc * m2.ca + m1.bd * m2.da + m1.be * m2.ea:    .bb = m1.ba * m2.ab + m1.bb * m2.bb + m1.bc * m2.cb + m1.bd * m2.db + m1.bd * m2.eb:    .bc = m1.ba * m2.ac + m1.bb * m2.bc + m1.bc * m2.cc + m1.bd * m2.dc + m1.be * m2.ec:    .bd = m1.ba * m2.ad + m1.bb * m2.bd + m1.bc * m2.cd + m1.bd * m2.dd + m1.be * m2.ed:    .be = m1.ba * m2.ae + m1.bb * m2.be + m1.bc * m2.ce + m1.bd * m2.de + m1.be * m2.ee
                        .ca = m1.ca * m2.aa + m1.cb * m2.ba + m1.cc * m2.ca + m1.cd * m2.da + m1.ce * m2.ea:    .cb = m1.ca * m2.ab + m1.cb * m2.bb + m1.cc * m2.cb + m1.cd * m2.db + m1.cd * m2.eb:    .cc = m1.ca * m2.ac + m1.cb * m2.bc + m1.cc * m2.cc + m1.cd * m2.dc + m1.ce * m2.ec:    .cd = m1.ca * m2.ad + m1.cb * m2.bd + m1.cc * m2.cd + m1.cd * m2.dd + m1.ce * m2.ed:    .ce = m1.ca * m2.ae + m1.cb * m2.be + m1.cc * m2.ce + m1.cd * m2.de + m1.ce * m2.ee
                        .da = m1.da * m2.aa + m1.db * m2.ba + m1.dc * m2.ca + m1.dd * m2.da + m1.de * m2.ea:    .db = m1.da * m2.ab + m1.db * m2.bb + m1.dc * m2.cb + m1.dd * m2.db + m1.dd * m2.eb:    .dc = m1.da * m2.ac + m1.db * m2.bc + m1.dc * m2.cc + m1.dd * m2.dc + m1.de * m2.ec:    .dd = m1.da * m2.ad + m1.db * m2.bd + m1.dc * m2.cd + m1.dd * m2.dd + m1.de * m2.ed:    .de = m1.da * m2.ae + m1.db * m2.be + m1.dc * m2.ce + m1.dd * m2.de + m1.de * m2.ee
                        .ea = m1.ea * m2.aa + m1.eb * m2.ba + m1.ec * m2.ca + m1.ed * m2.da + m1.ee * m2.ea:    .eb = m1.ea * m2.ab + m1.eb * m2.bb + m1.ec * m2.cb + m1.ed * m2.db + m1.ed * m2.eb:    .ec = m1.ea * m2.ac + m1.eb * m2.bc + m1.ec * m2.cc + m1.ed * m2.dc + m1.ee * m2.ec:    .ed = m1.ea * m2.ad + m1.eb * m2.bd + m1.ec * m2.cd + m1.ed * m2.dd + m1.ee * m2.ed:    .ee = m1.ea * m2.ae + m1.eb * m2.be + m1.ec * m2.ce + m1.ed * m2.de + m1.ee * m2.ee
    End With
End Function
Public Function Matrix6_mul(m1 As Matrix6, m2 As Matrix6) As Matrix6
    'Multipliziert eine 6x6 Matrix m1 mit einer Matrix m2
    With Matrix6_mul:   .aa = m1.aa * m2.aa + m1.ab * m2.ba + m1.ac * m2.ca + m1.ad * m2.da + m1.ae * m2.ea + m1.af * m2.fa:    .ab = m1.aa * m2.ab + m1.ab * m2.bb + m1.ac * m2.cb + m1.ad * m2.db + m1.ad * m2.eb + m1.af * m2.fb:    .ac = m1.aa * m2.ac + m1.ab * m2.bc + m1.ac * m2.cc + m1.ad * m2.dc + m1.ae * m2.ec + m1.af * m2.fc:    .ad = m1.aa * m2.ad + m1.ab * m2.bd + m1.ac * m2.cd + m1.ad * m2.dd + m1.ae * m2.ed + m1.af * m2.fd:    .ae = m1.aa * m2.ae + m1.ab * m2.be + m1.ac * m2.ce + m1.ad * m2.de + m1.ae * m2.ee + m1.af * m2.fe:    .af = m1.aa * m2.af + m1.ab * m2.bf + m1.ac * m2.cf + m1.ad * m2.df + m1.ae * m2.ef + m1.af * m2.ff
                        .ba = m1.ba * m2.aa + m1.bb * m2.ba + m1.bc * m2.ca + m1.bd * m2.da + m1.be * m2.ea + m1.bf * m2.fa:    .bb = m1.ba * m2.ab + m1.bb * m2.bb + m1.bc * m2.cb + m1.bd * m2.db + m1.bd * m2.eb + m1.bf * m2.fb:    .bc = m1.ba * m2.ac + m1.bb * m2.bc + m1.bc * m2.cc + m1.bd * m2.dc + m1.be * m2.ec + m1.bf * m2.fc:    .bd = m1.ba * m2.ad + m1.bb * m2.bd + m1.bc * m2.cd + m1.bd * m2.dd + m1.be * m2.ed + m1.bf * m2.fd:    .be = m1.ba * m2.ae + m1.bb * m2.be + m1.bc * m2.ce + m1.bd * m2.de + m1.be * m2.ee + m1.bf * m2.fe:    .bf = m1.ba * m2.af + m1.bb * m2.bf + m1.bc * m2.cf + m1.bd * m2.df + m1.be * m2.ef + m1.bf * m2.ff
                        .ca = m1.ca * m2.aa + m1.cb * m2.ba + m1.cc * m2.ca + m1.cd * m2.da + m1.ce * m2.ea + m1.cf * m2.fa:    .cb = m1.ca * m2.ab + m1.cb * m2.bb + m1.cc * m2.cb + m1.cd * m2.db + m1.cd * m2.eb + m1.cf * m2.fb:    .cc = m1.ca * m2.ac + m1.cb * m2.bc + m1.cc * m2.cc + m1.cd * m2.dc + m1.ce * m2.ec + m1.cf * m2.fc:    .cd = m1.ca * m2.ad + m1.cb * m2.bd + m1.cc * m2.cd + m1.cd * m2.dd + m1.ce * m2.ed + m1.cf * m2.fd:    .ce = m1.ca * m2.ae + m1.cb * m2.be + m1.cc * m2.ce + m1.cd * m2.de + m1.ce * m2.ee + m1.cf * m2.fe:    .cf = m1.ca * m2.af + m1.cb * m2.bf + m1.cc * m2.cf + m1.cd * m2.df + m1.ce * m2.ef + m1.cf * m2.ff
                        .da = m1.da * m2.aa + m1.db * m2.ba + m1.dc * m2.ca + m1.dd * m2.da + m1.de * m2.ea + m1.df * m2.fa:    .db = m1.da * m2.ab + m1.db * m2.bb + m1.dc * m2.cb + m1.dd * m2.db + m1.dd * m2.eb + m1.df * m2.fb:    .dc = m1.da * m2.ac + m1.db * m2.bc + m1.dc * m2.cc + m1.dd * m2.dc + m1.de * m2.ec + m1.df * m2.fc:    .dd = m1.da * m2.ad + m1.db * m2.bd + m1.dc * m2.cd + m1.dd * m2.dd + m1.de * m2.ed + m1.df * m2.fd:    .de = m1.da * m2.ae + m1.db * m2.be + m1.dc * m2.ce + m1.dd * m2.de + m1.de * m2.ee + m1.df * m2.fe:    .df = m1.da * m2.af + m1.db * m2.bf + m1.dc * m2.cf + m1.dd * m2.df + m1.de * m2.ef + m1.df * m2.ff
                        .ea = m1.ea * m2.aa + m1.eb * m2.ba + m1.ec * m2.ca + m1.ed * m2.da + m1.ee * m2.ea + m1.ef * m2.fa:    .eb = m1.ea * m2.ab + m1.eb * m2.bb + m1.ec * m2.cb + m1.ed * m2.db + m1.ed * m2.eb + m1.ef * m2.fb:    .ec = m1.ea * m2.ac + m1.eb * m2.bc + m1.ec * m2.cc + m1.ed * m2.dc + m1.ee * m2.ec + m1.ef * m2.fc:    .ed = m1.ea * m2.ad + m1.eb * m2.bd + m1.ec * m2.cd + m1.ed * m2.dd + m1.ee * m2.ed + m1.ef * m2.fd:    .ee = m1.ea * m2.ae + m1.eb * m2.be + m1.ec * m2.ce + m1.ed * m2.de + m1.ee * m2.ee + m1.ef * m2.fe:    .ef = m1.ea * m2.af + m1.eb * m2.bf + m1.ec * m2.cf + m1.ed * m2.df + m1.ee * m2.ef + m1.ef * m2.ff
                        .fa = m1.fa * m2.aa + m1.fb * m2.ba + m1.fc * m2.ca + m1.fd * m2.da + m1.fe * m2.ea + m1.ff * m2.fa:    .fb = m1.fa * m2.ab + m1.fb * m2.bb + m1.fc * m2.cb + m1.fd * m2.db + m1.fd * m2.eb + m1.ff * m2.fb:    .fc = m1.fa * m2.ac + m1.fb * m2.bc + m1.fc * m2.cc + m1.fd * m2.dc + m1.fe * m2.ec + m1.ff * m2.fc:    .fd = m1.fa * m2.ad + m1.fb * m2.bd + m1.fc * m2.cd + m1.fd * m2.dd + m1.fe * m2.ed + m1.ff * m2.fd:    .fe = m1.fa * m2.ae + m1.fb * m2.be + m1.fc * m2.ce + m1.fd * m2.de + m1.fe * m2.ee + m1.ff * m2.fe:    .ff = m1.fa * m2.af + m1.fb * m2.bf + m1.fc * m2.cf + m1.fd * m2.df + m1.fe * m2.ef + m1.ff * m2.ff
    End With
End Function
Public Function Matrix7_mul(m1 As Matrix7, m2 As Matrix7) As Matrix7
    'Multipliziert eine 7x7 Matrix m1 mit einer Matrix m2
    With Matrix7_mul:   .aa = m1.aa * m2.aa + m1.ab * m2.ba + m1.ac * m2.ca + m1.ad * m2.da + m1.ae * m2.ea + m1.af * m2.fa + m1.ag * m2.ga:    .ab = m1.aa * m2.ab + m1.ab * m2.bb + m1.ac * m2.cb + m1.ad * m2.db + m1.ad * m2.eb + m1.af * m2.fb + m1.ag * m2.gb:    .ac = m1.aa * m2.ac + m1.ab * m2.bc + m1.ac * m2.cc + m1.ad * m2.dc + m1.ae * m2.ec + m1.af * m2.fc + m1.ag * m2.gc:    .ad = m1.aa * m2.ad + m1.ab * m2.bd + m1.ac * m2.cd + m1.ad * m2.dd + m1.ae * m2.ed + m1.af * m2.fd + m1.ag * m2.gd:    .ae = m1.aa * m2.ae + m1.ab * m2.be + m1.ac * m2.ce + m1.ad * m2.de + m1.ae * m2.ee + m1.af * m2.fe + m1.ag * m2.ge:    .af = m1.aa * m2.af + m1.ab * m2.bf + m1.ac * m2.cf + m1.ad * m2.df + m1.ae * m2.ef + m1.af * m2.ff + m1.ag * m2.gf:    .ag = m1.aa * m2.ag + m1.ab * m2.bg + m1.ac * m2.cg + m1.ad * m2.dg + m1.ae * m2.eg + m1.af * m2.fg + m1.ag * m2.gg
                        .ba = m1.ba * m2.aa + m1.bb * m2.ba + m1.bc * m2.ca + m1.bd * m2.da + m1.be * m2.ea + m1.bf * m2.fa + m1.bg * m2.ga:    .bb = m1.ba * m2.ab + m1.bb * m2.bb + m1.bc * m2.cb + m1.bd * m2.db + m1.bd * m2.eb + m1.bf * m2.fb + m1.bg * m2.gb:    .bc = m1.ba * m2.ac + m1.bb * m2.bc + m1.bc * m2.cc + m1.bd * m2.dc + m1.be * m2.ec + m1.bf * m2.fc + m1.bg * m2.gc:    .bd = m1.ba * m2.ad + m1.bb * m2.bd + m1.bc * m2.cd + m1.bd * m2.dd + m1.be * m2.ed + m1.bf * m2.fd + m1.bg * m2.gd:    .be = m1.ba * m2.ae + m1.bb * m2.be + m1.bc * m2.ce + m1.bd * m2.de + m1.be * m2.ee + m1.bf * m2.fe + m1.bg * m2.ge:    .bf = m1.ba * m2.af + m1.bb * m2.bf + m1.bc * m2.cf + m1.bd * m2.df + m1.be * m2.ef + m1.bf * m2.ff + m1.bg * m2.gf:    .bg = m1.ba * m2.ag + m1.bb * m2.bg + m1.bc * m2.cg + m1.bd * m2.dg + m1.be * m2.eg + m1.bf * m2.fg + m1.bg * m2.gg
                        .ca = m1.ca * m2.aa + m1.cb * m2.ba + m1.cc * m2.ca + m1.cd * m2.da + m1.ce * m2.ea + m1.cf * m2.fa + m1.cg * m2.ga:    .cb = m1.ca * m2.ab + m1.cb * m2.bb + m1.cc * m2.cb + m1.cd * m2.db + m1.cd * m2.eb + m1.cf * m2.fb + m1.cg * m2.gb:    .cc = m1.ca * m2.ac + m1.cb * m2.bc + m1.cc * m2.cc + m1.cd * m2.dc + m1.ce * m2.ec + m1.cf * m2.fc + m1.cg * m2.gc:    .cd = m1.ca * m2.ad + m1.cb * m2.bd + m1.cc * m2.cd + m1.cd * m2.dd + m1.ce * m2.ed + m1.cf * m2.fd + m1.cg * m2.gd:    .ce = m1.ca * m2.ae + m1.cb * m2.be + m1.cc * m2.ce + m1.cd * m2.de + m1.ce * m2.ee + m1.cf * m2.fe + m1.cg * m2.ge:    .cf = m1.ca * m2.af + m1.cb * m2.bf + m1.cc * m2.cf + m1.cd * m2.df + m1.ce * m2.ef + m1.cf * m2.ff + m1.cg * m2.gf:    .cg = m1.ca * m2.ag + m1.cb * m2.bg + m1.cc * m2.cg + m1.cd * m2.dg + m1.ce * m2.eg + m1.cf * m2.fg + m1.cg * m2.gg
                        .da = m1.da * m2.aa + m1.db * m2.ba + m1.dc * m2.ca + m1.dd * m2.da + m1.de * m2.ea + m1.df * m2.fa + m1.dg * m2.ga:    .db = m1.da * m2.ab + m1.db * m2.bb + m1.dc * m2.cb + m1.dd * m2.db + m1.dd * m2.eb + m1.df * m2.fb + m1.dg * m2.gb:    .dc = m1.da * m2.ac + m1.db * m2.bc + m1.dc * m2.cc + m1.dd * m2.dc + m1.de * m2.ec + m1.df * m2.fc + m1.dg * m2.gc:    .dd = m1.da * m2.ad + m1.db * m2.bd + m1.dc * m2.cd + m1.dd * m2.dd + m1.de * m2.ed + m1.df * m2.fd + m1.dg * m2.gd:    .de = m1.da * m2.ae + m1.db * m2.be + m1.dc * m2.ce + m1.dd * m2.de + m1.de * m2.ee + m1.df * m2.fe + m1.dg * m2.ge:    .df = m1.da * m2.af + m1.db * m2.bf + m1.dc * m2.cf + m1.dd * m2.df + m1.de * m2.ef + m1.df * m2.ff + m1.dg * m2.gf:    .dg = m1.da * m2.ag + m1.db * m2.bg + m1.dc * m2.cg + m1.dd * m2.dg + m1.de * m2.eg + m1.df * m2.fg + m1.dg * m2.gg
                        .ea = m1.ea * m2.aa + m1.eb * m2.ba + m1.ec * m2.ca + m1.ed * m2.da + m1.ee * m2.ea + m1.ef * m2.fa + m1.eg * m2.ga:    .eb = m1.ea * m2.ab + m1.eb * m2.bb + m1.ec * m2.cb + m1.ed * m2.db + m1.ed * m2.eb + m1.ef * m2.fb + m1.eg * m2.gb:    .ec = m1.ea * m2.ac + m1.eb * m2.bc + m1.ec * m2.cc + m1.ed * m2.dc + m1.ee * m2.ec + m1.ef * m2.fc + m1.eg * m2.gc:    .ed = m1.ea * m2.ad + m1.eb * m2.bd + m1.ec * m2.cd + m1.ed * m2.dd + m1.ee * m2.ed + m1.ef * m2.fd + m1.eg * m2.gd:    .ee = m1.ea * m2.ae + m1.eb * m2.be + m1.ec * m2.ce + m1.ed * m2.de + m1.ee * m2.ee + m1.ef * m2.fe + m1.eg * m2.ge:    .ef = m1.ea * m2.af + m1.eb * m2.bf + m1.ec * m2.cf + m1.ed * m2.df + m1.ee * m2.ef + m1.ef * m2.ff + m1.eg * m2.gf:    .eg = m1.ea * m2.ag + m1.eb * m2.bg + m1.ec * m2.cg + m1.ed * m2.dg + m1.ee * m2.eg + m1.ef * m2.fg + m1.eg * m2.gg
                        .fa = m1.fa * m2.aa + m1.fb * m2.ba + m1.fc * m2.ca + m1.fd * m2.da + m1.fe * m2.ea + m1.ff * m2.fa + m1.fg * m2.ga:    .fb = m1.fa * m2.ab + m1.fb * m2.bb + m1.fc * m2.cb + m1.fd * m2.db + m1.fd * m2.eb + m1.ff * m2.fb + m1.fg * m2.gb:    .fc = m1.fa * m2.ac + m1.fb * m2.bc + m1.fc * m2.cc + m1.fd * m2.dc + m1.fe * m2.ec + m1.ff * m2.fc + m1.fg * m2.gc:    .fd = m1.fa * m2.ad + m1.fb * m2.bd + m1.fc * m2.cd + m1.fd * m2.dd + m1.fe * m2.ed + m1.ff * m2.fd + m1.fg * m2.gd:    .fe = m1.fa * m2.ae + m1.fb * m2.be + m1.fc * m2.ce + m1.fd * m2.de + m1.fe * m2.ee + m1.ff * m2.fe + m1.fg * m2.ge:    .ff = m1.fa * m2.af + m1.fb * m2.bf + m1.fc * m2.cf + m1.fd * m2.df + m1.fe * m2.ef + m1.ff * m2.ff + m1.fg * m2.gf:    .fg = m1.fa * m2.ag + m1.fb * m2.bg + m1.fc * m2.cg + m1.fd * m2.dg + m1.fe * m2.eg + m1.ff * m2.fg + m1.fg * m2.gg
                        .ga = m1.ga * m2.aa + m1.gb * m2.ba + m1.gc * m2.ca + m1.gd * m2.da + m1.ge * m2.ea + m1.gf * m2.fa + m1.gg * m2.ga:    .gb = m1.ga * m2.ab + m1.gb * m2.bb + m1.gc * m2.cb + m1.gd * m2.db + m1.gd * m2.eb + m1.gf * m2.fb + m1.gg * m2.gb:    .gc = m1.ga * m2.ac + m1.gb * m2.bc + m1.gc * m2.cc + m1.gd * m2.dc + m1.ge * m2.ec + m1.gf * m2.fc + m1.gg * m2.gc:    .gd = m1.ga * m2.ad + m1.gb * m2.bd + m1.gc * m2.cd + m1.gd * m2.dd + m1.ge * m2.ed + m1.gf * m2.fd + m1.gg * m2.gd:    .ge = m1.ga * m2.ae + m1.gb * m2.be + m1.gc * m2.ce + m1.gd * m2.de + m1.ge * m2.ee + m1.gf * m2.fe + m1.gg * m2.ge:    .gf = m1.ga * m2.af + m1.gb * m2.bf + m1.gc * m2.cf + m1.gd * m2.df + m1.ge * m2.ef + m1.gf * m2.ff + m1.gg * m2.gf:    .gg = m1.ga * m2.ag + m1.gb * m2.bg + m1.gc * m2.cg + m1.gd * m2.dg + m1.ge * m2.eg + m1.gf * m2.fg + m1.gg * m2.gg
    End With
End Function
Public Function Matrix8_mul(m1 As Matrix8, m2 As Matrix8) As Matrix8
    'Multipliziert eine 8x8 Matrix m1 mit einer Matrix m2
    With Matrix8_mul
        .aa = m1.aa * m2.aa + m1.ab * m2.ba + m1.ac * m2.ca + m1.ad * m2.da + m1.ae * m2.ea + m1.af * m2.fa + m1.ag * m2.ga + m1.ah * m2.ha:    .ab = m1.aa * m2.ab + m1.ab * m2.bb + m1.ac * m2.cb + m1.ad * m2.db + m1.ad * m2.eb + m1.af * m2.fb + m1.ag * m2.gb + m1.ah * m2.hb:    .ac = m1.aa * m2.ac + m1.ab * m2.bc + m1.ac * m2.cc + m1.ad * m2.dc + m1.ae * m2.ec + m1.af * m2.fc + m1.ag * m2.gc + m1.ah * m2.hc:    .ad = m1.aa * m2.ad + m1.ab * m2.bd + m1.ac * m2.cd + m1.ad * m2.dd + m1.ae * m2.ed + m1.af * m2.fd + m1.ag * m2.gd + m1.ah * m2.hd:    .ae = m1.aa * m2.ae + m1.ab * m2.be + m1.ac * m2.ce + m1.ad * m2.de + m1.ae * m2.ee + m1.af * m2.fe + m1.ag * m2.ge + m1.ah * m2.he:    .af = m1.aa * m2.af + m1.ab * m2.bf + m1.ac * m2.cf + m1.ad * m2.df + m1.ae * m2.ef + m1.af * m2.ff + m1.ag * m2.gf + m1.ah * m2.hf:
        .ba = m1.ba * m2.aa + m1.bb * m2.ba + m1.bc * m2.ca + m1.bd * m2.da + m1.be * m2.ea + m1.bf * m2.fa + m1.bg * m2.ga + m1.bh * m2.ha:    .bb = m1.ba * m2.ab + m1.bb * m2.bb + m1.bc * m2.cb + m1.bd * m2.db + m1.bd * m2.eb + m1.bf * m2.fb + m1.bg * m2.gb + m1.bh * m2.hb:    .bc = m1.ba * m2.ac + m1.bb * m2.bc + m1.bc * m2.cc + m1.bd * m2.dc + m1.be * m2.ec + m1.bf * m2.fc + m1.bg * m2.gc + m1.bh * m2.hc:    .bd = m1.ba * m2.ad + m1.bb * m2.bd + m1.bc * m2.cd + m1.bd * m2.dd + m1.be * m2.ed + m1.bf * m2.fd + m1.bg * m2.gd + m1.bh * m2.hd:    .be = m1.ba * m2.ae + m1.bb * m2.be + m1.bc * m2.ce + m1.bd * m2.de + m1.be * m2.ee + m1.bf * m2.fe + m1.bg * m2.ge + m1.bh * m2.he:    .bf = m1.ba * m2.af + m1.bb * m2.bf + m1.bc * m2.cf + m1.bd * m2.df + m1.be * m2.ef + m1.bf * m2.ff + m1.bg * m2.gf + m1.bh * m2.hf:
        .ca = m1.ca * m2.aa + m1.cb * m2.ba + m1.cc * m2.ca + m1.cd * m2.da + m1.ce * m2.ea + m1.cf * m2.fa + m1.cg * m2.ga + m1.ch * m2.ha:    .cb = m1.ca * m2.ab + m1.cb * m2.bb + m1.cc * m2.cb + m1.cd * m2.db + m1.cd * m2.eb + m1.cf * m2.fb + m1.cg * m2.gb + m1.ch * m2.hb:    .cc = m1.ca * m2.ac + m1.cb * m2.bc + m1.cc * m2.cc + m1.cd * m2.dc + m1.ce * m2.ec + m1.cf * m2.fc + m1.cg * m2.gc + m1.ch * m2.hc:    .cd = m1.ca * m2.ad + m1.cb * m2.bd + m1.cc * m2.cd + m1.cd * m2.dd + m1.ce * m2.ed + m1.cf * m2.fd + m1.cg * m2.gd + m1.ch * m2.hd:    .ce = m1.ca * m2.ae + m1.cb * m2.be + m1.cc * m2.ce + m1.cd * m2.de + m1.ce * m2.ee + m1.cf * m2.fe + m1.cg * m2.ge + m1.ch * m2.he:    .cf = m1.ca * m2.af + m1.cb * m2.bf + m1.cc * m2.cf + m1.cd * m2.df + m1.ce * m2.ef + m1.cf * m2.ff + m1.cg * m2.gf + m1.ch * m2.hf:
        .da = m1.da * m2.aa + m1.db * m2.ba + m1.dc * m2.ca + m1.dd * m2.da + m1.de * m2.ea + m1.df * m2.fa + m1.dg * m2.ga + m1.dh * m2.ha:    .db = m1.da * m2.ab + m1.db * m2.bb + m1.dc * m2.cb + m1.dd * m2.db + m1.dd * m2.eb + m1.df * m2.fb + m1.dg * m2.gb + m1.dh * m2.hb:    .dc = m1.da * m2.ac + m1.db * m2.bc + m1.dc * m2.cc + m1.dd * m2.dc + m1.de * m2.ec + m1.df * m2.fc + m1.dg * m2.gc + m1.dh * m2.hc:    .dd = m1.da * m2.ad + m1.db * m2.bd + m1.dc * m2.cd + m1.dd * m2.dd + m1.de * m2.ed + m1.df * m2.fd + m1.dg * m2.gd + m1.dh * m2.hd:    .de = m1.da * m2.ae + m1.db * m2.be + m1.dc * m2.ce + m1.dd * m2.de + m1.de * m2.ee + m1.df * m2.fe + m1.dg * m2.ge + m1.dh * m2.he:    .df = m1.da * m2.af + m1.db * m2.bf + m1.dc * m2.cf + m1.dd * m2.df + m1.de * m2.ef + m1.df * m2.ff + m1.dg * m2.gf + m1.dh * m2.hf:
        .ea = m1.ea * m2.aa + m1.eb * m2.ba + m1.ec * m2.ca + m1.ed * m2.da + m1.ee * m2.ea + m1.ef * m2.fa + m1.eg * m2.ga + m1.eh * m2.ha:    .eb = m1.ea * m2.ab + m1.eb * m2.bb + m1.ec * m2.cb + m1.ed * m2.db + m1.ed * m2.eb + m1.ef * m2.fb + m1.eg * m2.gb + m1.eh * m2.hb:    .ec = m1.ea * m2.ac + m1.eb * m2.bc + m1.ec * m2.cc + m1.ed * m2.dc + m1.ee * m2.ec + m1.ef * m2.fc + m1.eg * m2.gc + m1.eh * m2.hc:    .ed = m1.ea * m2.ad + m1.eb * m2.bd + m1.ec * m2.cd + m1.ed * m2.dd + m1.ee * m2.ed + m1.ef * m2.fd + m1.eg * m2.gd + m1.eh * m2.hd:    .ee = m1.ea * m2.ae + m1.eb * m2.be + m1.ec * m2.ce + m1.ed * m2.de + m1.ee * m2.ee + m1.ef * m2.fe + m1.eg * m2.ge + m1.eh * m2.he:    .ef = m1.ea * m2.af + m1.eb * m2.bf + m1.ec * m2.cf + m1.ed * m2.df + m1.ee * m2.ef + m1.ef * m2.ff + m1.eg * m2.gf + m1.eh * m2.hf:
        .fa = m1.fa * m2.aa + m1.fb * m2.ba + m1.fc * m2.ca + m1.fd * m2.da + m1.fe * m2.ea + m1.ff * m2.fa + m1.fg * m2.ga + m1.fh * m2.ha:    .fb = m1.fa * m2.ab + m1.fb * m2.bb + m1.fc * m2.cb + m1.fd * m2.db + m1.fd * m2.eb + m1.ff * m2.fb + m1.fg * m2.gb + m1.fh * m2.hb:    .fc = m1.fa * m2.ac + m1.fb * m2.bc + m1.fc * m2.cc + m1.fd * m2.dc + m1.fe * m2.ec + m1.ff * m2.fc + m1.fg * m2.gc + m1.fh * m2.hc:    .fd = m1.fa * m2.ad + m1.fb * m2.bd + m1.fc * m2.cd + m1.fd * m2.dd + m1.fe * m2.ed + m1.ff * m2.fd + m1.fg * m2.gd + m1.fh * m2.hd:    .fe = m1.fa * m2.ae + m1.fb * m2.be + m1.fc * m2.ce + m1.fd * m2.de + m1.fe * m2.ee + m1.ff * m2.fe + m1.fg * m2.ge + m1.fh * m2.he:    .ff = m1.fa * m2.af + m1.fb * m2.bf + m1.fc * m2.cf + m1.fd * m2.df + m1.fe * m2.ef + m1.ff * m2.ff + m1.fg * m2.gf + m1.fh * m2.hf:
        .ga = m1.ga * m2.aa + m1.gb * m2.ba + m1.gc * m2.ca + m1.gd * m2.da + m1.ge * m2.ea + m1.gf * m2.fa + m1.gg * m2.ga + m1.gh * m2.ha:    .gb = m1.ga * m2.ab + m1.gb * m2.bb + m1.gc * m2.cb + m1.gd * m2.db + m1.gd * m2.eb + m1.gf * m2.fb + m1.gg * m2.gb + m1.gh * m2.hb:    .gc = m1.ga * m2.ac + m1.gb * m2.bc + m1.gc * m2.cc + m1.gd * m2.dc + m1.ge * m2.ec + m1.gf * m2.fc + m1.gg * m2.gc + m1.gh * m2.hc:    .gd = m1.ga * m2.ad + m1.gb * m2.bd + m1.gc * m2.cd + m1.gd * m2.dd + m1.ge * m2.ed + m1.gf * m2.fd + m1.gg * m2.gd + m1.gh * m2.hd:    .ge = m1.ga * m2.ae + m1.gb * m2.be + m1.gc * m2.ce + m1.gd * m2.de + m1.ge * m2.ee + m1.gf * m2.fe + m1.gg * m2.ge + m1.gh * m2.he:    .gf = m1.ga * m2.af + m1.gb * m2.bf + m1.gc * m2.cf + m1.gd * m2.df + m1.ge * m2.ef + m1.gf * m2.ff + m1.gg * m2.gf + m1.gh * m2.hf:
        .ha = m1.ha * m2.aa + m1.hb * m2.ba + m1.hc * m2.ca + m1.hd * m2.da + m1.he * m2.ea + m1.hf * m2.fa + m1.hg * m2.ga + m1.hh * m2.ha:    .hb = m1.ha * m2.ab + m1.hb * m2.bb + m1.hc * m2.cb + m1.hd * m2.db + m1.hd * m2.eb + m1.hf * m2.fb + m1.hg * m2.gb + m1.hh * m2.hb:    .hc = m1.ha * m2.ac + m1.hb * m2.bc + m1.hc * m2.cc + m1.hd * m2.dc + m1.he * m2.ec + m1.hf * m2.fc + m1.hg * m2.gc + m1.hh * m2.hc:    .hd = m1.ha * m2.ad + m1.hb * m2.bd + m1.hc * m2.cd + m1.hd * m2.dd + m1.he * m2.ed + m1.hf * m2.fd + m1.hg * m2.gd + m1.hh * m2.hd:    .he = m1.ha * m2.ae + m1.hb * m2.be + m1.hc * m2.ce + m1.hd * m2.de + m1.he * m2.ee + m1.hf * m2.fe + m1.hg * m2.ge + m1.hh * m2.he:    .hf = m1.ha * m2.af + m1.hb * m2.bf + m1.hc * m2.cf + m1.hd * m2.df + m1.he * m2.ef + m1.hf * m2.ff + m1.hg * m2.gf + m1.hh * m2.hf:
        
        .ag = m1.aa * m2.ag + m1.ab * m2.bg + m1.ac * m2.cg + m1.ad * m2.dg + m1.ae * m2.eg + m1.af * m2.fg + m1.ag * m2.gg + m1.ah * m2.hg:    .ah = m1.aa * m2.ah + m1.ab * m2.bh + m1.ac * m2.ch + m1.ad * m2.dh + m1.ae * m2.eh + m1.af * m2.fh + m1.ag * m2.gh + m1.ah * m2.hh
        .bg = m1.ba * m2.ag + m1.bb * m2.bg + m1.bc * m2.cg + m1.bd * m2.dg + m1.be * m2.eg + m1.bf * m2.fg + m1.bg * m2.gg + m1.bh * m2.hg:    .bh = m1.ba * m2.ah + m1.bb * m2.bh + m1.bc * m2.ch + m1.bd * m2.dh + m1.be * m2.eh + m1.bf * m2.fh + m1.bg * m2.gh + m1.bh * m2.hh
        .cg = m1.ca * m2.ag + m1.cb * m2.bg + m1.cc * m2.cg + m1.cd * m2.dg + m1.ce * m2.eg + m1.cf * m2.fg + m1.cg * m2.gg + m1.ch * m2.hg:    .ch = m1.ca * m2.ah + m1.cb * m2.bh + m1.cc * m2.ch + m1.cd * m2.dh + m1.ce * m2.eh + m1.cf * m2.fh + m1.cg * m2.gh + m1.ch * m2.hh
        .dg = m1.da * m2.ag + m1.db * m2.bg + m1.dc * m2.cg + m1.dd * m2.dg + m1.de * m2.eg + m1.df * m2.fg + m1.dg * m2.gg + m1.dh * m2.hg:    .dh = m1.da * m2.ah + m1.db * m2.bh + m1.dc * m2.ch + m1.dd * m2.dh + m1.de * m2.eh + m1.df * m2.fh + m1.dg * m2.gh + m1.dh * m2.hh
        .eg = m1.ea * m2.ag + m1.eb * m2.bg + m1.ec * m2.cg + m1.ed * m2.dg + m1.ee * m2.eg + m1.ef * m2.fg + m1.eg * m2.gg + m1.eh * m2.hg:    .eh = m1.ea * m2.ah + m1.eb * m2.bh + m1.ec * m2.ch + m1.ed * m2.dh + m1.ee * m2.eh + m1.ef * m2.fh + m1.eg * m2.gh + m1.eh * m2.hh
        .fg = m1.fa * m2.ag + m1.fb * m2.bg + m1.fc * m2.cg + m1.fd * m2.dg + m1.fe * m2.eg + m1.ff * m2.fg + m1.fg * m2.gg + m1.fh * m2.hg:    .fh = m1.fa * m2.ah + m1.fb * m2.bh + m1.fc * m2.ch + m1.fd * m2.dh + m1.fe * m2.eh + m1.ff * m2.fh + m1.fg * m2.gh + m1.fh * m2.hh
        .gg = m1.ga * m2.ag + m1.gb * m2.bg + m1.gc * m2.cg + m1.gd * m2.dg + m1.ge * m2.eg + m1.gf * m2.fg + m1.gg * m2.gg + m1.gh * m2.hg:    .gh = m1.ga * m2.ah + m1.gb * m2.bh + m1.gc * m2.ch + m1.gd * m2.dh + m1.ge * m2.eh + m1.gf * m2.fh + m1.gg * m2.gh + m1.gh * m2.hh
        .hg = m1.ga * m2.ag + m1.gb * m2.bg + m1.gc * m2.cg + m1.gd * m2.dg + m1.ge * m2.eg + m1.gf * m2.fg + m1.gg * m2.gg + m1.gh * m2.hg:    .hh = m1.ha * m2.ah + m1.hb * m2.bh + m1.hc * m2.ch + m1.hd * m2.dh + m1.he * m2.eh + m1.hf * m2.fh + m1.hg * m2.gh + m1.hh * m2.hh
        
    End With
End Function
Public Function Matrix9_mul(m1 As Matrix9, m2 As Matrix9) As Matrix9
    'Multipliziert eine 9x9 Matrix m1 mit einer Matrix m2
    With Matrix9_mul
        .aa = m1.aa * m2.aa + m1.ab * m2.ba + m1.ac * m2.ca + m1.ad * m2.da + m1.ae * m2.ea + m1.af * m2.fa + m1.ag * m2.ga + m1.ah * m2.ha + m1.ai * m2.ia:    .ab = m1.aa * m2.ab + m1.ab * m2.bb + m1.ac * m2.cb + m1.ad * m2.db + m1.ad * m2.eb + m1.af * m2.fb + m1.ag * m2.gb + m1.ah * m2.hb + m1.ai * m2.ib:    .ac = m1.aa * m2.ac + m1.ab * m2.bc + m1.ac * m2.cc + m1.ad * m2.dc + m1.ae * m2.ec + m1.af * m2.fc + m1.ag * m2.gc + m1.ah * m2.hc + m1.ai * m2.ic:    .ad = m1.aa * m2.ad + m1.ab * m2.bd + m1.ac * m2.cd + m1.ad * m2.dd + m1.ae * m2.ed + m1.af * m2.fd + m1.ag * m2.gd + m1.ah * m2.hd + m1.ai * m2.id:    .ae = m1.aa * m2.ae + m1.ab * m2.be + m1.ac * m2.ce + m1.ad * m2.de + m1.ae * m2.ee + m1.af * m2.fe + m1.ag * m2.ge + m1.ah * m2.he + m1.ai * m2.ie:    .af = m1.aa * m2.af + m1.ab * m2.bf + m1.ac * m2.cf + m1.ad * m2.df + m1.ae * m2.ef + m1.af * m2.ff + m1.ag * m2.gf + m1.ah * m2.hf + m1.ai * m2.if
        .ba = m1.ba * m2.aa + m1.bb * m2.ba + m1.bc * m2.ca + m1.bd * m2.da + m1.be * m2.ea + m1.bf * m2.fa + m1.bg * m2.ga + m1.bh * m2.ha + m1.bi * m2.ia:    .bb = m1.ba * m2.ab + m1.bb * m2.bb + m1.bc * m2.cb + m1.bd * m2.db + m1.bd * m2.eb + m1.bf * m2.fb + m1.bg * m2.gb + m1.bh * m2.hb + m1.bi * m2.ib:    .bc = m1.ba * m2.ac + m1.bb * m2.bc + m1.bc * m2.cc + m1.bd * m2.dc + m1.be * m2.ec + m1.bf * m2.fc + m1.bg * m2.gc + m1.bh * m2.hc + m1.bi * m2.ic:    .bd = m1.ba * m2.ad + m1.bb * m2.bd + m1.bc * m2.cd + m1.bd * m2.dd + m1.be * m2.ed + m1.bf * m2.fd + m1.bg * m2.gd + m1.bh * m2.hd + m1.bi * m2.id:    .be = m1.ba * m2.ae + m1.bb * m2.be + m1.bc * m2.ce + m1.bd * m2.de + m1.be * m2.ee + m1.bf * m2.fe + m1.bg * m2.ge + m1.bh * m2.he + m1.bi * m2.ie:    .bf = m1.ba * m2.af + m1.bb * m2.bf + m1.bc * m2.cf + m1.bd * m2.df + m1.be * m2.ef + m1.bf * m2.ff + m1.bg * m2.gf + m1.bh * m2.hf + m1.bi * m2.if
        .ca = m1.ca * m2.aa + m1.cb * m2.ba + m1.cc * m2.ca + m1.cd * m2.da + m1.ce * m2.ea + m1.cf * m2.fa + m1.cg * m2.ga + m1.ch * m2.ha + m1.ci * m2.ia:    .cb = m1.ca * m2.ab + m1.cb * m2.bb + m1.cc * m2.cb + m1.cd * m2.db + m1.cd * m2.eb + m1.cf * m2.fb + m1.cg * m2.gb + m1.ch * m2.hb + m1.ci * m2.ib:    .cc = m1.ca * m2.ac + m1.cb * m2.bc + m1.cc * m2.cc + m1.cd * m2.dc + m1.ce * m2.ec + m1.cf * m2.fc + m1.cg * m2.gc + m1.ch * m2.hc + m1.ci * m2.ic:    .cd = m1.ca * m2.ad + m1.cb * m2.bd + m1.cc * m2.cd + m1.cd * m2.dd + m1.ce * m2.ed + m1.cf * m2.fd + m1.cg * m2.gd + m1.ch * m2.hd + m1.ci * m2.id:    .ce = m1.ca * m2.ae + m1.cb * m2.be + m1.cc * m2.ce + m1.cd * m2.de + m1.ce * m2.ee + m1.cf * m2.fe + m1.cg * m2.ge + m1.ch * m2.he + m1.ci * m2.ie:    .cf = m1.ca * m2.af + m1.cb * m2.bf + m1.cc * m2.cf + m1.cd * m2.df + m1.ce * m2.ef + m1.cf * m2.ff + m1.cg * m2.gf + m1.ch * m2.hf + m1.ci * m2.if
        .da = m1.da * m2.aa + m1.db * m2.ba + m1.dc * m2.ca + m1.dd * m2.da + m1.de * m2.ea + m1.df * m2.fa + m1.dg * m2.ga + m1.dh * m2.ha + m1.di * m2.ia:    .db = m1.da * m2.ab + m1.db * m2.bb + m1.dc * m2.cb + m1.dd * m2.db + m1.dd * m2.eb + m1.df * m2.fb + m1.dg * m2.gb + m1.dh * m2.hb + m1.di * m2.ib:    .dc = m1.da * m2.ac + m1.db * m2.bc + m1.dc * m2.cc + m1.dd * m2.dc + m1.de * m2.ec + m1.df * m2.fc + m1.dg * m2.gc + m1.dh * m2.hc + m1.di * m2.ic:    .dd = m1.da * m2.ad + m1.db * m2.bd + m1.dc * m2.cd + m1.dd * m2.dd + m1.de * m2.ed + m1.df * m2.fd + m1.dg * m2.gd + m1.dh * m2.hd + m1.di * m2.id:    .de = m1.da * m2.ae + m1.db * m2.be + m1.dc * m2.ce + m1.dd * m2.de + m1.de * m2.ee + m1.df * m2.fe + m1.dg * m2.ge + m1.dh * m2.he + m1.di * m2.ie:    .df = m1.da * m2.af + m1.db * m2.bf + m1.dc * m2.cf + m1.dd * m2.df + m1.de * m2.ef + m1.df * m2.ff + m1.dg * m2.gf + m1.dh * m2.hf + m1.di * m2.if
        .ea = m1.ea * m2.aa + m1.eb * m2.ba + m1.ec * m2.ca + m1.ed * m2.da + m1.ee * m2.ea + m1.ef * m2.fa + m1.eg * m2.ga + m1.eh * m2.ha + m1.ei * m2.ia:    .eb = m1.ea * m2.ab + m1.eb * m2.bb + m1.ec * m2.cb + m1.ed * m2.db + m1.ed * m2.eb + m1.ef * m2.fb + m1.eg * m2.gb + m1.eh * m2.hb + m1.ei * m2.ib:    .ec = m1.ea * m2.ac + m1.eb * m2.bc + m1.ec * m2.cc + m1.ed * m2.dc + m1.ee * m2.ec + m1.ef * m2.fc + m1.eg * m2.gc + m1.eh * m2.hc + m1.ei * m2.ic:    .ed = m1.ea * m2.ad + m1.eb * m2.bd + m1.ec * m2.cd + m1.ed * m2.dd + m1.ee * m2.ed + m1.ef * m2.fd + m1.eg * m2.gd + m1.eh * m2.hd + m1.ei * m2.id:    .ee = m1.ea * m2.ae + m1.eb * m2.be + m1.ec * m2.ce + m1.ed * m2.de + m1.ee * m2.ee + m1.ef * m2.fe + m1.eg * m2.ge + m1.eh * m2.he + m1.ei * m2.ie:    .ef = m1.ea * m2.af + m1.eb * m2.bf + m1.ec * m2.cf + m1.ed * m2.df + m1.ee * m2.ef + m1.ef * m2.ff + m1.eg * m2.gf + m1.eh * m2.hf + m1.ei * m2.if
        .fa = m1.fa * m2.aa + m1.fb * m2.ba + m1.fc * m2.ca + m1.fd * m2.da + m1.fe * m2.ea + m1.ff * m2.fa + m1.fg * m2.ga + m1.fh * m2.ha + m1.fi * m2.ia:    .fb = m1.fa * m2.ab + m1.fb * m2.bb + m1.fc * m2.cb + m1.fd * m2.db + m1.fd * m2.eb + m1.ff * m2.fb + m1.fg * m2.gb + m1.fh * m2.hb + m1.fi * m2.ib:    .fc = m1.fa * m2.ac + m1.fb * m2.bc + m1.fc * m2.cc + m1.fd * m2.dc + m1.fe * m2.ec + m1.ff * m2.fc + m1.fg * m2.gc + m1.fh * m2.hc + m1.fi * m2.ic:    .fd = m1.fa * m2.ad + m1.fb * m2.bd + m1.fc * m2.cd + m1.fd * m2.dd + m1.fe * m2.ed + m1.ff * m2.fd + m1.fg * m2.gd + m1.fh * m2.hd + m1.fi * m2.id:    .fe = m1.fa * m2.ae + m1.fb * m2.be + m1.fc * m2.ce + m1.fd * m2.de + m1.fe * m2.ee + m1.ff * m2.fe + m1.fg * m2.ge + m1.fh * m2.he + m1.fi * m2.ie:    .ff = m1.fa * m2.af + m1.fb * m2.bf + m1.fc * m2.cf + m1.fd * m2.df + m1.fe * m2.ef + m1.ff * m2.ff + m1.fg * m2.gf + m1.fh * m2.hf + m1.fi * m2.if
        .ga = m1.ga * m2.aa + m1.gb * m2.ba + m1.gc * m2.ca + m1.gd * m2.da + m1.ge * m2.ea + m1.gf * m2.fa + m1.gg * m2.ga + m1.gh * m2.ha + m1.gi * m2.ia:    .gb = m1.ga * m2.ab + m1.gb * m2.bb + m1.gc * m2.cb + m1.gd * m2.db + m1.gd * m2.eb + m1.gf * m2.fb + m1.gg * m2.gb + m1.gh * m2.hb + m1.gi * m2.ib:    .gc = m1.ga * m2.ac + m1.gb * m2.bc + m1.gc * m2.cc + m1.gd * m2.dc + m1.ge * m2.ec + m1.gf * m2.fc + m1.gg * m2.gc + m1.gh * m2.hc + m1.gi * m2.ic:    .gd = m1.ga * m2.ad + m1.gb * m2.bd + m1.gc * m2.cd + m1.gd * m2.dd + m1.ge * m2.ed + m1.gf * m2.fd + m1.gg * m2.gd + m1.gh * m2.hd + m1.gi * m2.id:    .ge = m1.ga * m2.ae + m1.gb * m2.be + m1.gc * m2.ce + m1.gd * m2.de + m1.ge * m2.ee + m1.gf * m2.fe + m1.gg * m2.ge + m1.gh * m2.he + m1.gi * m2.ie:    .gf = m1.ga * m2.af + m1.gb * m2.bf + m1.gc * m2.cf + m1.gd * m2.df + m1.ge * m2.ef + m1.gf * m2.ff + m1.gg * m2.gf + m1.gh * m2.hf + m1.gi * m2.if
        .ha = m1.ha * m2.aa + m1.hb * m2.ba + m1.hc * m2.ca + m1.hd * m2.da + m1.he * m2.ea + m1.hf * m2.fa + m1.hg * m2.ga + m1.hh * m2.ha + m1.hi * m2.ia:    .hb = m1.ha * m2.ab + m1.hb * m2.bb + m1.hc * m2.cb + m1.hd * m2.db + m1.hd * m2.eb + m1.hf * m2.fb + m1.hg * m2.gb + m1.hh * m2.hb + m1.hi * m2.ib:    .hc = m1.ha * m2.ac + m1.hb * m2.bc + m1.hc * m2.cc + m1.hd * m2.dc + m1.he * m2.ec + m1.hf * m2.fc + m1.hg * m2.gc + m1.hh * m2.hc + m1.hi * m2.ic:    .hd = m1.ha * m2.ad + m1.hb * m2.bd + m1.hc * m2.cd + m1.hd * m2.dd + m1.he * m2.ed + m1.hf * m2.fd + m1.hg * m2.gd + m1.hh * m2.hd + m1.hi * m2.id:    .he = m1.ha * m2.ae + m1.hb * m2.be + m1.hc * m2.ce + m1.hd * m2.de + m1.he * m2.ee + m1.hf * m2.fe + m1.hg * m2.ge + m1.hh * m2.he + m1.hi * m2.ie:    .hf = m1.ha * m2.af + m1.hb * m2.bf + m1.hc * m2.cf + m1.hd * m2.df + m1.he * m2.ef + m1.hf * m2.ff + m1.hg * m2.gf + m1.hh * m2.hf + m1.hi * m2.if
        .ia = m1.ia * m2.aa + m1.ib * m2.ba + m1.ic * m2.ca + m1.id * m2.da + m1.ie * m2.ea + m1.if * m2.fa + m1.ig * m2.ga + m1.ih * m2.ha + m1.ii * m2.ia:    .ib = m1.ia * m2.ab + m1.ib * m2.bb + m1.ic * m2.cb + m1.id * m2.db + m1.id * m2.eb + m1.if * m2.fb + m1.ig * m2.gb + m1.ih * m2.hb + m1.ii * m2.ib:    .ic = m1.ia * m2.ac + m1.ib * m2.bc + m1.ic * m2.cc + m1.id * m2.dc + m1.ie * m2.ec + m1.if * m2.fc + m1.ig * m2.gc + m1.ih * m2.hc + m1.ii * m2.ic:    .id = m1.ia * m2.ad + m1.ib * m2.bd + m1.ic * m2.cd + m1.id * m2.dd + m1.ie * m2.ed + m1.if * m2.fd + m1.ig * m2.gd + m1.ih * m2.hd + m1.ii * m2.id:    .ie = m1.ia * m2.ae + m1.ib * m2.be + m1.ic * m2.ce + m1.id * m2.de + m1.ie * m2.ee + m1.if * m2.fe + m1.ig * m2.ge + m1.ih * m2.he + m1.ii * m2.ie:    .if = m1.ia * m2.af + m1.ib * m2.bf + m1.ic * m2.cf + m1.id * m2.df + m1.ie * m2.ef + m1.if * m2.ff + m1.ig * m2.gf + m1.ih * m2.hf + m1.ii * m2.if
        
        .ag = m1.aa * m2.ag + m1.ab * m2.bg + m1.ac * m2.cg + m1.ad * m2.dg + m1.ae * m2.eg + m1.af * m2.fg + m1.ag * m2.gg + m1.ah * m2.hg + m1.ai * m2.ig:    .ah = m1.aa * m2.ah + m1.ab * m2.bh + m1.ac * m2.ch + m1.ad * m2.dh + m1.ae * m2.eh + m1.af * m2.fh + m1.ag * m2.gh + m1.ah * m2.hh + m1.ai * m2.ih:    .ai = m1.aa * m2.ai + m1.ab * m2.bi + m1.ac * m2.ci + m1.ad * m2.di + m1.ae * m2.ei + m1.af * m2.fi + m1.ag * m2.gi + m1.ah * m2.hi + m1.ai * m2.ii
        .bg = m1.ba * m2.ag + m1.bb * m2.bg + m1.bc * m2.cg + m1.bd * m2.dg + m1.be * m2.eg + m1.bf * m2.fg + m1.bg * m2.gg + m1.bh * m2.hg + m1.bi * m2.ig:    .bh = m1.ba * m2.ah + m1.bb * m2.bh + m1.bc * m2.ch + m1.bd * m2.dh + m1.be * m2.eh + m1.bf * m2.fh + m1.bg * m2.gh + m1.bh * m2.hh + m1.bi * m2.ih:    .bi = m1.ba * m2.ai + m1.bb * m2.bi + m1.bc * m2.ci + m1.bd * m2.di + m1.be * m2.ei + m1.bf * m2.fi + m1.bg * m2.gi + m1.bh * m2.hi + m1.bi * m2.ii
        .cg = m1.ca * m2.ag + m1.cb * m2.bg + m1.cc * m2.cg + m1.cd * m2.dg + m1.ce * m2.eg + m1.cf * m2.fg + m1.cg * m2.gg + m1.ch * m2.hg + m1.ci * m2.ig:    .ch = m1.ca * m2.ah + m1.cb * m2.bh + m1.cc * m2.ch + m1.cd * m2.dh + m1.ce * m2.eh + m1.cf * m2.fh + m1.cg * m2.gh + m1.ch * m2.hh + m1.ci * m2.ih:    .ci = m1.ca * m2.ai + m1.cb * m2.bi + m1.cc * m2.ci + m1.cd * m2.di + m1.ce * m2.ei + m1.cf * m2.fi + m1.cg * m2.gi + m1.ch * m2.hi + m1.ci * m2.ii
        .dg = m1.da * m2.ag + m1.db * m2.bg + m1.dc * m2.cg + m1.dd * m2.dg + m1.de * m2.eg + m1.df * m2.fg + m1.dg * m2.gg + m1.dh * m2.hg + m1.di * m2.ig:    .dh = m1.da * m2.ah + m1.db * m2.bh + m1.dc * m2.ch + m1.dd * m2.dh + m1.de * m2.eh + m1.df * m2.fh + m1.dg * m2.gh + m1.dh * m2.hh + m1.di * m2.ih:    .di = m1.da * m2.ai + m1.db * m2.bi + m1.dc * m2.ci + m1.dd * m2.di + m1.de * m2.ei + m1.df * m2.fi + m1.dg * m2.gi + m1.dh * m2.hi + m1.di * m2.ii
        .eg = m1.ea * m2.ag + m1.eb * m2.bg + m1.ec * m2.cg + m1.ed * m2.dg + m1.ee * m2.eg + m1.ef * m2.fg + m1.eg * m2.gg + m1.eh * m2.hg + m1.ei * m2.ig:    .eh = m1.ea * m2.ah + m1.eb * m2.bh + m1.ec * m2.ch + m1.ed * m2.dh + m1.ee * m2.eh + m1.ef * m2.fh + m1.eg * m2.gh + m1.eh * m2.hh + m1.ei * m2.ih:    .ei = m1.ea * m2.ai + m1.eb * m2.bi + m1.ec * m2.ci + m1.ed * m2.di + m1.ee * m2.ei + m1.ef * m2.fi + m1.eg * m2.gi + m1.eh * m2.hi + m1.ei * m2.ii
        .fg = m1.fa * m2.ag + m1.fb * m2.bg + m1.fc * m2.cg + m1.fd * m2.dg + m1.fe * m2.eg + m1.ff * m2.fg + m1.fg * m2.gg + m1.fh * m2.hg + m1.fi * m2.ig:    .fh = m1.fa * m2.ah + m1.fb * m2.bh + m1.fc * m2.ch + m1.fd * m2.dh + m1.fe * m2.eh + m1.ff * m2.fh + m1.fg * m2.gh + m1.fh * m2.hh + m1.fi * m2.ih:    .fi = m1.fa * m2.ai + m1.fb * m2.bi + m1.fc * m2.ci + m1.fd * m2.di + m1.fe * m2.ei + m1.ff * m2.fi + m1.fg * m2.gi + m1.fh * m2.hi + m1.fi * m2.ii
        .gg = m1.ga * m2.ag + m1.gb * m2.bg + m1.gc * m2.cg + m1.gd * m2.dg + m1.ge * m2.eg + m1.gf * m2.fg + m1.gg * m2.gg + m1.gh * m2.hg + m1.gi * m2.ig:    .gh = m1.ga * m2.ah + m1.gb * m2.bh + m1.gc * m2.ch + m1.gd * m2.dh + m1.ge * m2.eh + m1.gf * m2.fh + m1.gg * m2.gh + m1.gh * m2.hh + m1.gi * m2.ih:    .gi = m1.ga * m2.ai + m1.gb * m2.bi + m1.gc * m2.ci + m1.gd * m2.di + m1.ge * m2.ei + m1.gf * m2.fi + m1.gg * m2.gi + m1.gh * m2.hi + m1.gi * m2.ii
        .hg = m1.ha * m2.ag + m1.hb * m2.bg + m1.hc * m2.cg + m1.hd * m2.dg + m1.he * m2.eg + m1.hf * m2.fg + m1.hg * m2.gg + m1.hh * m2.hg + m1.hi * m2.ig:    .hh = m1.ha * m2.ah + m1.hb * m2.bh + m1.hc * m2.ch + m1.hd * m2.dh + m1.he * m2.eh + m1.hf * m2.fh + m1.hg * m2.gh + m1.hh * m2.hh + m1.hi * m2.ih:    .hi = m1.ha * m2.ai + m1.hb * m2.bi + m1.hc * m2.ci + m1.hd * m2.di + m1.he * m2.ei + m1.hf * m2.fi + m1.hg * m2.gi + m1.hh * m2.hi + m1.hi * m2.ii
        .ig = m1.ia * m2.ag + m1.ib * m2.bg + m1.ic * m2.cg + m1.id * m2.dg + m1.ie * m2.eg + m1.if * m2.fg + m1.ig * m2.gg + m1.ih * m2.hg + m1.ii * m2.ig:    .ih = m1.ia * m2.ah + m1.ib * m2.bh + m1.ic * m2.ch + m1.id * m2.dh + m1.ie * m2.eh + m1.if * m2.fh + m1.ig * m2.gh + m1.ih * m2.hh + m1.ii * m2.ih:    .ii = m1.ia * m2.ai + m1.ib * m2.bi + m1.ic * m2.ci + m1.id * m2.di + m1.ie * m2.ei + m1.if * m2.fi + m1.ig * m2.gi + m1.ih * m2.hi + m1.ii * m2.ii
    End With
End Function
Public Function Matrix10_mul(m1 As Matrix10, m2 As Matrix10) As Matrix10
    'Multipliziert eine 10x10 Matrix m1 mit einer Matrix m2
    With Matrix10_mul
        .aa = m1.aa * m2.aa + m1.ab * m2.ba + m1.ac * m2.ca + m1.ad * m2.da + m1.ae * m2.ea + m1.af * m2.fa + m1.ag * m2.ga + m1.ah * m2.ha + m1.ai * m2.ia + m1.aj * m2.ja:    .ab = m1.aa * m2.ab + m1.ab * m2.bb + m1.ac * m2.cb + m1.ad * m2.db + m1.ad * m2.eb + m1.af * m2.fb + m1.ag * m2.gb + m1.ah * m2.hb + m1.ai * m2.ib + m1.aj * m2.jb:    .ac = m1.aa * m2.ac + m1.ab * m2.bc + m1.ac * m2.cc + m1.ad * m2.dc + m1.ae * m2.ec + m1.af * m2.fc + m1.ag * m2.gc + m1.ah * m2.hc + m1.ai * m2.ic + m1.aj * m2.jc:    .ad = m1.aa * m2.ad + m1.ab * m2.bd + m1.ac * m2.cd + m1.ad * m2.dd + m1.ae * m2.ed + m1.af * m2.fd + m1.ag * m2.gd + m1.ah * m2.hd + m1.ai * m2.id + m1.aj * m2.jd:    .ae = m1.aa * m2.ae + m1.ab * m2.be + m1.ac * m2.ce + m1.ad * m2.de + m1.ae * m2.ee + m1.af * m2.fe + m1.ag * m2.ge + m1.ah * m2.he + m1.ai * m2.ie + m1.aj * m2.je
        .ba = m1.ba * m2.aa + m1.bb * m2.ba + m1.bc * m2.ca + m1.bd * m2.da + m1.be * m2.ea + m1.bf * m2.fa + m1.bg * m2.ga + m1.bh * m2.ha + m1.bi * m2.ia + m1.bj * m2.ja:    .bb = m1.ba * m2.ab + m1.bb * m2.bb + m1.bc * m2.cb + m1.bd * m2.db + m1.bd * m2.eb + m1.bf * m2.fb + m1.bg * m2.gb + m1.bh * m2.hb + m1.bi * m2.ib + m1.bj * m2.jb:    .bc = m1.ba * m2.ac + m1.bb * m2.bc + m1.bc * m2.cc + m1.bd * m2.dc + m1.be * m2.ec + m1.bf * m2.fc + m1.bg * m2.gc + m1.bh * m2.hc + m1.bi * m2.ic + m1.bj * m2.jc:    .bd = m1.ba * m2.ad + m1.bb * m2.bd + m1.bc * m2.cd + m1.bd * m2.dd + m1.be * m2.ed + m1.bf * m2.fd + m1.bg * m2.gd + m1.bh * m2.hd + m1.bi * m2.id + m1.bj * m2.jd:    .be = m1.ba * m2.ae + m1.bb * m2.be + m1.bc * m2.ce + m1.bd * m2.de + m1.be * m2.ee + m1.bf * m2.fe + m1.bg * m2.ge + m1.bh * m2.he + m1.bi * m2.ie + m1.bj * m2.je
        .ca = m1.ca * m2.aa + m1.cb * m2.ba + m1.cc * m2.ca + m1.cd * m2.da + m1.ce * m2.ea + m1.cf * m2.fa + m1.cg * m2.ga + m1.ch * m2.ha + m1.ci * m2.ia + m1.cj * m2.ja:    .cb = m1.ca * m2.ab + m1.cb * m2.bb + m1.cc * m2.cb + m1.cd * m2.db + m1.cd * m2.eb + m1.cf * m2.fb + m1.cg * m2.gb + m1.ch * m2.hb + m1.ci * m2.ib + m1.cj * m2.jb:    .cc = m1.ca * m2.ac + m1.cb * m2.bc + m1.cc * m2.cc + m1.cd * m2.dc + m1.ce * m2.ec + m1.cf * m2.fc + m1.cg * m2.gc + m1.ch * m2.hc + m1.ci * m2.ic + m1.cj * m2.jc:    .cd = m1.ca * m2.ad + m1.cb * m2.bd + m1.cc * m2.cd + m1.cd * m2.dd + m1.ce * m2.ed + m1.cf * m2.fd + m1.cg * m2.gd + m1.ch * m2.hd + m1.ci * m2.id + m1.cj * m2.jd:    .ce = m1.ca * m2.ae + m1.cb * m2.be + m1.cc * m2.ce + m1.cd * m2.de + m1.ce * m2.ee + m1.cf * m2.fe + m1.cg * m2.ge + m1.ch * m2.he + m1.ci * m2.ie + m1.cj * m2.je
        .da = m1.da * m2.aa + m1.db * m2.ba + m1.dc * m2.ca + m1.dd * m2.da + m1.de * m2.ea + m1.df * m2.fa + m1.dg * m2.ga + m1.dh * m2.ha + m1.di * m2.ia + m1.dj * m2.ja:    .db = m1.da * m2.ab + m1.db * m2.bb + m1.dc * m2.cb + m1.dd * m2.db + m1.dd * m2.eb + m1.df * m2.fb + m1.dg * m2.gb + m1.dh * m2.hb + m1.di * m2.ib + m1.dj * m2.jb:    .dc = m1.da * m2.ac + m1.db * m2.bc + m1.dc * m2.cc + m1.dd * m2.dc + m1.de * m2.ec + m1.df * m2.fc + m1.dg * m2.gc + m1.dh * m2.hc + m1.di * m2.ic + m1.dj * m2.jc:    .dd = m1.da * m2.ad + m1.db * m2.bd + m1.dc * m2.cd + m1.dd * m2.dd + m1.de * m2.ed + m1.df * m2.fd + m1.dg * m2.gd + m1.dh * m2.hd + m1.di * m2.id + m1.dj * m2.jd:    .de = m1.da * m2.ae + m1.db * m2.be + m1.dc * m2.ce + m1.dd * m2.de + m1.de * m2.ee + m1.df * m2.fe + m1.dg * m2.ge + m1.dh * m2.he + m1.di * m2.ie + m1.dj * m2.je
        .ea = m1.ea * m2.aa + m1.eb * m2.ba + m1.ec * m2.ca + m1.ed * m2.da + m1.ee * m2.ea + m1.ef * m2.fa + m1.eg * m2.ga + m1.eh * m2.ha + m1.ei * m2.ia + m1.ej * m2.ja:    .eb = m1.ea * m2.ab + m1.eb * m2.bb + m1.ec * m2.cb + m1.ed * m2.db + m1.ed * m2.eb + m1.ef * m2.fb + m1.eg * m2.gb + m1.eh * m2.hb + m1.ei * m2.ib + m1.ej * m2.jb:    .ec = m1.ea * m2.ac + m1.eb * m2.bc + m1.ec * m2.cc + m1.ed * m2.dc + m1.ee * m2.ec + m1.ef * m2.fc + m1.eg * m2.gc + m1.eh * m2.hc + m1.ei * m2.ic + m1.ej * m2.jc:    .ed = m1.ea * m2.ad + m1.eb * m2.bd + m1.ec * m2.cd + m1.ed * m2.dd + m1.ee * m2.ed + m1.ef * m2.fd + m1.eg * m2.gd + m1.eh * m2.hd + m1.ei * m2.id + m1.ej * m2.jd:    .ee = m1.ea * m2.ae + m1.eb * m2.be + m1.ec * m2.ce + m1.ed * m2.de + m1.ee * m2.ee + m1.ef * m2.fe + m1.eg * m2.ge + m1.eh * m2.he + m1.ei * m2.ie + m1.ej * m2.je
        .fa = m1.fa * m2.aa + m1.fb * m2.ba + m1.fc * m2.ca + m1.fd * m2.da + m1.fe * m2.ea + m1.ff * m2.fa + m1.fg * m2.ga + m1.fh * m2.ha + m1.fi * m2.ia + m1.fj * m2.ja:    .fb = m1.fa * m2.ab + m1.fb * m2.bb + m1.fc * m2.cb + m1.fd * m2.db + m1.fd * m2.eb + m1.ff * m2.fb + m1.fg * m2.gb + m1.fh * m2.hb + m1.fi * m2.ib + m1.fj * m2.jb:    .fc = m1.fa * m2.ac + m1.fb * m2.bc + m1.fc * m2.cc + m1.fd * m2.dc + m1.fe * m2.ec + m1.ff * m2.fc + m1.fg * m2.gc + m1.fh * m2.hc + m1.fi * m2.ic + m1.fj * m2.jc:    .fd = m1.fa * m2.ad + m1.fb * m2.bd + m1.fc * m2.cd + m1.fd * m2.dd + m1.fe * m2.ed + m1.ff * m2.fd + m1.fg * m2.gd + m1.fh * m2.hd + m1.fi * m2.id + m1.fj * m2.jd:    .fe = m1.fa * m2.ae + m1.fb * m2.be + m1.fc * m2.ce + m1.fd * m2.de + m1.fe * m2.ee + m1.ff * m2.fe + m1.fg * m2.ge + m1.fh * m2.he + m1.fi * m2.ie + m1.fj * m2.je
        .ga = m1.ga * m2.aa + m1.gb * m2.ba + m1.gc * m2.ca + m1.gd * m2.da + m1.ge * m2.ea + m1.gf * m2.fa + m1.gg * m2.ga + m1.gh * m2.ha + m1.gi * m2.ia + m1.gj * m2.ja:    .gb = m1.ga * m2.ab + m1.gb * m2.bb + m1.gc * m2.cb + m1.gd * m2.db + m1.gd * m2.eb + m1.gf * m2.fb + m1.gg * m2.gb + m1.gh * m2.hb + m1.gi * m2.ib + m1.gj * m2.jb:    .gc = m1.ga * m2.ac + m1.gb * m2.bc + m1.gc * m2.cc + m1.gd * m2.dc + m1.ge * m2.ec + m1.gf * m2.fc + m1.gg * m2.gc + m1.gh * m2.hc + m1.gi * m2.ic + m1.gj * m2.jc:    .gd = m1.ga * m2.ad + m1.gb * m2.bd + m1.gc * m2.cd + m1.gd * m2.dd + m1.ge * m2.ed + m1.gf * m2.fd + m1.gg * m2.gd + m1.gh * m2.hd + m1.gi * m2.id + m1.gj * m2.jd:    .ge = m1.ga * m2.ae + m1.gb * m2.be + m1.gc * m2.ce + m1.gd * m2.de + m1.ge * m2.ee + m1.gf * m2.fe + m1.gg * m2.ge + m1.gh * m2.he + m1.gi * m2.ie + m1.gj * m2.je
        .ha = m1.ha * m2.aa + m1.hb * m2.ba + m1.hc * m2.ca + m1.hd * m2.da + m1.he * m2.ea + m1.hf * m2.fa + m1.hg * m2.ga + m1.hh * m2.ha + m1.hi * m2.ia + m1.hj * m2.ja:    .hb = m1.ha * m2.ab + m1.hb * m2.bb + m1.hc * m2.cb + m1.hd * m2.db + m1.hd * m2.eb + m1.hf * m2.fb + m1.hg * m2.gb + m1.hh * m2.hb + m1.hi * m2.ib + m1.hj * m2.jb:    .hc = m1.ha * m2.ac + m1.hb * m2.bc + m1.hc * m2.cc + m1.hd * m2.dc + m1.he * m2.ec + m1.hf * m2.fc + m1.hg * m2.gc + m1.hh * m2.hc + m1.hi * m2.ic + m1.hj * m2.jc:    .hd = m1.ha * m2.ad + m1.hb * m2.bd + m1.hc * m2.cd + m1.hd * m2.dd + m1.he * m2.ed + m1.hf * m2.fd + m1.hg * m2.gd + m1.hh * m2.hd + m1.hi * m2.id + m1.hj * m2.jd:    .he = m1.ha * m2.ae + m1.hb * m2.be + m1.hc * m2.ce + m1.hd * m2.de + m1.he * m2.ee + m1.hf * m2.fe + m1.hg * m2.ge + m1.hh * m2.he + m1.hi * m2.ie + m1.hj * m2.je
        .ia = m1.ia * m2.aa + m1.ib * m2.ba + m1.ic * m2.ca + m1.id * m2.da + m1.ie * m2.ea + m1.if * m2.fa + m1.ig * m2.ga + m1.ih * m2.ha + m1.ii * m2.ia + m1.ij * m2.ja:    .ib = m1.ia * m2.ab + m1.ib * m2.bb + m1.ic * m2.cb + m1.id * m2.db + m1.id * m2.eb + m1.if * m2.fb + m1.ig * m2.gb + m1.ih * m2.hb + m1.ii * m2.ib + m1.ij * m2.jb:    .ic = m1.ia * m2.ac + m1.ib * m2.bc + m1.ic * m2.cc + m1.id * m2.dc + m1.ie * m2.ec + m1.if * m2.fc + m1.ig * m2.gc + m1.ih * m2.hc + m1.ii * m2.ic + m1.ij * m2.jc:    .id = m1.ia * m2.ad + m1.ib * m2.bd + m1.ic * m2.cd + m1.id * m2.dd + m1.ie * m2.ed + m1.if * m2.fd + m1.ig * m2.gd + m1.ih * m2.hd + m1.ii * m2.id + m1.ij * m2.jd:    .ie = m1.ia * m2.ae + m1.ib * m2.be + m1.ic * m2.ce + m1.id * m2.de + m1.ie * m2.ee + m1.if * m2.fe + m1.ig * m2.ge + m1.ih * m2.he + m1.ii * m2.ie + m1.ij * m2.je
        .ja = m1.ja * m2.aa + m1.jb * m2.ba + m1.jc * m2.ca + m1.jd * m2.da + m1.je * m2.ea + m1.jf * m2.fa + m1.jg * m2.ga + m1.jh * m2.ha + m1.ji * m2.ia + m1.jj * m2.ja:    .jb = m1.ja * m2.ab + m1.jb * m2.bb + m1.jc * m2.cb + m1.jd * m2.db + m1.jd * m2.eb + m1.jf * m2.fb + m1.jg * m2.gb + m1.jh * m2.hb + m1.ji * m2.ib + m1.jj * m2.jb:    .jc = m1.ja * m2.ac + m1.jb * m2.bc + m1.jc * m2.cc + m1.jd * m2.dc + m1.je * m2.ec + m1.jf * m2.fc + m1.jg * m2.gc + m1.jh * m2.hc + m1.ji * m2.ic + m1.jj * m2.jc:    .jd = m1.ja * m2.ad + m1.jb * m2.bd + m1.jc * m2.cd + m1.jd * m2.dd + m1.je * m2.ed + m1.jf * m2.fd + m1.jg * m2.gd + m1.jh * m2.hd + m1.ji * m2.id + m1.jj * m2.jd:    .je = m1.ja * m2.ae + m1.jb * m2.be + m1.jc * m2.ce + m1.jd * m2.de + m1.je * m2.ee + m1.jf * m2.fe + m1.jg * m2.ge + m1.jh * m2.he + m1.ji * m2.ie + m1.jj * m2.je
        
        
        .af = m1.aa * m2.af + m1.ab * m2.bf + m1.ac * m2.cf + m1.ad * m2.df + m1.ae * m2.ef + m1.af * m2.ff + m1.ag * m2.gf + m1.ah * m2.hf + m1.ai * m2.if + m1.aj * m2.jf:    .ag = m1.aa * m2.ag + m1.ab * m2.bg + m1.ac * m2.cg + m1.ad * m2.dg + m1.ae * m2.eg + m1.af * m2.fg + m1.ag * m2.gg + m1.ah * m2.hg + m1.ai * m2.ig + m1.aj * m2.jg:    .ah = m1.aa * m2.ah + m1.ab * m2.bh + m1.ac * m2.ch + m1.ad * m2.dh + m1.ae * m2.eh + m1.af * m2.fh + m1.ag * m2.gh + m1.ah * m2.hh + m1.ai * m2.ih + m1.aj * m2.jh:    .ai = m1.aa * m2.ai + m1.ab * m2.bi + m1.ac * m2.ci + m1.ad * m2.di + m1.ae * m2.ei + m1.af * m2.fi + m1.ag * m2.gi + m1.ah * m2.hi + m1.ai * m2.ii + m1.aj * m2.ji:    .aj = m1.aa * m2.aj + m1.ab * m2.bj + m1.ac * m2.cj + m1.ad * m2.dj + m1.ae * m2.ej + m1.af * m2.fj + m1.ag * m2.gj + m1.ah * m2.hj + m1.ai * m2.ij + m1.aj * m2.jj
        .bf = m1.ba * m2.af + m1.bb * m2.bf + m1.bc * m2.cf + m1.bd * m2.df + m1.be * m2.ef + m1.bf * m2.ff + m1.bg * m2.gf + m1.bh * m2.hf + m1.bi * m2.if + m1.bj * m2.jf:    .bg = m1.ba * m2.ag + m1.bb * m2.bg + m1.bc * m2.cg + m1.bd * m2.dg + m1.be * m2.eg + m1.bf * m2.fg + m1.bg * m2.gg + m1.bh * m2.hg + m1.bi * m2.ig + m1.bj * m2.jg:    .bh = m1.ba * m2.ah + m1.bb * m2.bh + m1.bc * m2.ch + m1.bd * m2.dh + m1.be * m2.eh + m1.bf * m2.fh + m1.bg * m2.gh + m1.bh * m2.hh + m1.bi * m2.ih + m1.bj * m2.jh:    .bi = m1.ba * m2.ai + m1.bb * m2.bi + m1.bc * m2.ci + m1.bd * m2.di + m1.be * m2.ei + m1.bf * m2.fi + m1.bg * m2.gi + m1.bh * m2.hi + m1.bi * m2.ii + m1.bj * m2.ji:    .bj = m1.ba * m2.aj + m1.bb * m2.bj + m1.bc * m2.cj + m1.bd * m2.dj + m1.be * m2.ej + m1.bf * m2.fj + m1.bg * m2.gj + m1.bh * m2.hj + m1.bi * m2.ij + m1.bj * m2.jj
        .cf = m1.ca * m2.af + m1.cb * m2.bf + m1.cc * m2.cf + m1.cd * m2.df + m1.ce * m2.ef + m1.cf * m2.ff + m1.cg * m2.gf + m1.ch * m2.hf + m1.ci * m2.if + m1.cj * m2.jf:    .cg = m1.ca * m2.ag + m1.cb * m2.bg + m1.cc * m2.cg + m1.cd * m2.dg + m1.ce * m2.eg + m1.cf * m2.fg + m1.cg * m2.gg + m1.ch * m2.hg + m1.ci * m2.ig + m1.cj * m2.jg:    .ch = m1.ca * m2.ah + m1.cb * m2.bh + m1.cc * m2.ch + m1.cd * m2.dh + m1.ce * m2.eh + m1.cf * m2.fh + m1.cg * m2.gh + m1.ch * m2.hh + m1.ci * m2.ih + m1.cj * m2.jh:    .ci = m1.ca * m2.ai + m1.cb * m2.bi + m1.cc * m2.ci + m1.cd * m2.di + m1.ce * m2.ei + m1.cf * m2.fi + m1.cg * m2.gi + m1.ch * m2.hi + m1.ci * m2.ii + m1.cj * m2.ji:    .cj = m1.ca * m2.aj + m1.cb * m2.bj + m1.cc * m2.cj + m1.cd * m2.dj + m1.ce * m2.ej + m1.cf * m2.fj + m1.cg * m2.gj + m1.ch * m2.hj + m1.ci * m2.ij + m1.cj * m2.jj
        .df = m1.da * m2.af + m1.db * m2.bf + m1.dc * m2.cf + m1.dd * m2.df + m1.de * m2.ef + m1.df * m2.ff + m1.dg * m2.gf + m1.dh * m2.hf + m1.di * m2.if + m1.dj * m2.jf:    .dg = m1.da * m2.ag + m1.db * m2.bg + m1.dc * m2.cg + m1.dd * m2.dg + m1.de * m2.eg + m1.df * m2.fg + m1.dg * m2.gg + m1.dh * m2.hg + m1.di * m2.ig + m1.dj * m2.jg:    .dh = m1.da * m2.ah + m1.db * m2.bh + m1.dc * m2.ch + m1.dd * m2.dh + m1.de * m2.eh + m1.df * m2.fh + m1.dg * m2.gh + m1.dh * m2.hh + m1.di * m2.ih + m1.dj * m2.jh:    .di = m1.da * m2.ai + m1.db * m2.bi + m1.dc * m2.ci + m1.dd * m2.di + m1.de * m2.ei + m1.df * m2.fi + m1.dg * m2.gi + m1.dh * m2.hi + m1.di * m2.ii + m1.dj * m2.ji:    .dj = m1.da * m2.aj + m1.db * m2.bj + m1.dc * m2.cj + m1.dd * m2.dj + m1.de * m2.ej + m1.df * m2.fj + m1.dg * m2.gj + m1.dh * m2.hj + m1.di * m2.ij + m1.dj * m2.jj
        .ef = m1.ea * m2.af + m1.eb * m2.bf + m1.ec * m2.cf + m1.ed * m2.df + m1.ee * m2.ef + m1.ef * m2.ff + m1.eg * m2.gf + m1.eh * m2.hf + m1.ei * m2.if + m1.ej * m2.jf:    .eg = m1.ea * m2.ag + m1.eb * m2.bg + m1.ec * m2.cg + m1.ed * m2.dg + m1.ee * m2.eg + m1.ef * m2.fg + m1.eg * m2.gg + m1.eh * m2.hg + m1.ei * m2.ig + m1.ej * m2.jg:    .eh = m1.ea * m2.ah + m1.eb * m2.bh + m1.ec * m2.ch + m1.ed * m2.dh + m1.ee * m2.eh + m1.ef * m2.fh + m1.eg * m2.gh + m1.eh * m2.hh + m1.ei * m2.ih + m1.ej * m2.jh:    .ei = m1.ea * m2.ai + m1.eb * m2.bi + m1.ec * m2.ci + m1.ed * m2.di + m1.ee * m2.ei + m1.ef * m2.fi + m1.eg * m2.gi + m1.eh * m2.hi + m1.ei * m2.ii + m1.ej * m2.ji:    .ej = m1.ea * m2.aj + m1.eb * m2.bj + m1.ec * m2.cj + m1.ed * m2.dj + m1.ee * m2.ej + m1.ef * m2.fj + m1.eg * m2.gj + m1.eh * m2.hj + m1.ei * m2.ij + m1.ej * m2.jj
        .ff = m1.fa * m2.af + m1.fb * m2.bf + m1.fc * m2.cf + m1.fd * m2.df + m1.fe * m2.ef + m1.ff * m2.ff + m1.fg * m2.gf + m1.fh * m2.hf + m1.fi * m2.if + m1.fj * m2.jf:    .fg = m1.fa * m2.ag + m1.fb * m2.bg + m1.fc * m2.cg + m1.fd * m2.dg + m1.fe * m2.eg + m1.ff * m2.fg + m1.fg * m2.gg + m1.fh * m2.hg + m1.fi * m2.ig + m1.fj * m2.jg:    .fh = m1.fa * m2.ah + m1.fb * m2.bh + m1.fc * m2.ch + m1.fd * m2.dh + m1.fe * m2.eh + m1.ff * m2.fh + m1.fg * m2.gh + m1.fh * m2.hh + m1.fi * m2.ih + m1.fj * m2.jh:    .fi = m1.fa * m2.ai + m1.fb * m2.bi + m1.fc * m2.ci + m1.fd * m2.di + m1.fe * m2.ei + m1.ff * m2.fi + m1.fg * m2.gi + m1.fh * m2.hi + m1.fi * m2.ii + m1.fj * m2.ji:    .fj = m1.fa * m2.aj + m1.fb * m2.bj + m1.fc * m2.cj + m1.fd * m2.dj + m1.fe * m2.ej + m1.ff * m2.fj + m1.fg * m2.gj + m1.fh * m2.hj + m1.fi * m2.ij + m1.fj * m2.jj
        .gf = m1.ga * m2.af + m1.gb * m2.bf + m1.gc * m2.cf + m1.gd * m2.df + m1.ge * m2.ef + m1.gf * m2.ff + m1.gg * m2.gf + m1.gh * m2.hf + m1.gi * m2.if + m1.gj * m2.jf:    .gg = m1.ga * m2.ag + m1.gb * m2.bg + m1.gc * m2.cg + m1.gd * m2.dg + m1.ge * m2.eg + m1.gf * m2.fg + m1.gg * m2.gg + m1.gh * m2.hg + m1.gi * m2.ig + m1.gj * m2.jg:    .gh = m1.ga * m2.ah + m1.gb * m2.bh + m1.gc * m2.ch + m1.gd * m2.dh + m1.ge * m2.eh + m1.gf * m2.fh + m1.gg * m2.gh + m1.gh * m2.hh + m1.gi * m2.ih + m1.gj * m2.jh:    .gi = m1.ga * m2.ai + m1.gb * m2.bi + m1.gc * m2.ci + m1.gd * m2.di + m1.ge * m2.ei + m1.gf * m2.fi + m1.gg * m2.gi + m1.gh * m2.hi + m1.gi * m2.ii + m1.gj * m2.ji:    .gj = m1.ga * m2.aj + m1.gb * m2.bj + m1.gc * m2.cj + m1.gd * m2.dj + m1.ge * m2.ej + m1.gf * m2.fj + m1.gg * m2.gj + m1.gh * m2.hj + m1.gi * m2.ij + m1.gj * m2.jj
        .hf = m1.ha * m2.af + m1.hb * m2.bf + m1.hc * m2.cf + m1.hd * m2.df + m1.he * m2.ef + m1.hf * m2.ff + m1.hg * m2.gf + m1.hh * m2.hf + m1.hi * m2.if + m1.hj * m2.jf:    .hg = m1.ha * m2.ag + m1.hb * m2.bg + m1.hc * m2.cg + m1.hd * m2.dg + m1.he * m2.eg + m1.hf * m2.fg + m1.hg * m2.gg + m1.hh * m2.hg + m1.hi * m2.ig + m1.hj * m2.jg:    .hh = m1.ha * m2.ah + m1.hb * m2.bh + m1.hc * m2.ch + m1.hd * m2.dh + m1.he * m2.eh + m1.hf * m2.fh + m1.hg * m2.gh + m1.hh * m2.hh + m1.hi * m2.ih + m1.hj * m2.jh:    .hi = m1.ha * m2.ai + m1.hb * m2.bi + m1.hc * m2.ci + m1.hd * m2.di + m1.he * m2.ei + m1.hf * m2.fi + m1.hg * m2.gi + m1.hh * m2.hi + m1.hi * m2.ii + m1.hj * m2.ji:    .hj = m1.ha * m2.aj + m1.hb * m2.bj + m1.hc * m2.cj + m1.hd * m2.dj + m1.he * m2.ej + m1.hf * m2.fj + m1.hg * m2.gj + m1.hh * m2.hj + m1.hi * m2.ij + m1.hj * m2.jj
        .if = m1.ia * m2.af + m1.ib * m2.bf + m1.ic * m2.cf + m1.id * m2.df + m1.ie * m2.ef + m1.if * m2.ff + m1.ig * m2.gf + m1.ih * m2.hf + m1.ii * m2.if + m1.ij * m2.jf:    .ig = m1.ia * m2.ag + m1.ib * m2.bg + m1.ic * m2.cg + m1.id * m2.dg + m1.ie * m2.eg + m1.if * m2.fg + m1.ig * m2.gg + m1.ih * m2.hg + m1.ii * m2.ig + m1.ij * m2.jg:    .ih = m1.ia * m2.ah + m1.ib * m2.bh + m1.ic * m2.ch + m1.id * m2.dh + m1.ie * m2.eh + m1.if * m2.fh + m1.ig * m2.gh + m1.ih * m2.hh + m1.ii * m2.ih + m1.ij * m2.jh:    .ii = m1.ia * m2.ai + m1.ib * m2.bi + m1.ic * m2.ci + m1.id * m2.di + m1.ie * m2.ei + m1.if * m2.fi + m1.ig * m2.gi + m1.ih * m2.hi + m1.ii * m2.ii + m1.ij * m2.ji:    .ij = m1.ia * m2.aj + m1.ib * m2.bj + m1.ic * m2.cj + m1.id * m2.dj + m1.ie * m2.ej + m1.if * m2.fj + m1.ig * m2.gj + m1.ih * m2.hj + m1.ii * m2.ij + m1.ij * m2.jj
        .jf = m1.ja * m2.af + m1.jb * m2.bf + m1.jc * m2.cf + m1.jd * m2.df + m1.je * m2.ef + m1.jf * m2.ff + m1.jg * m2.gf + m1.jh * m2.hf + m1.ji * m2.if + m1.jj * m2.jf:    .jg = m1.ja * m2.ag + m1.jb * m2.bg + m1.jc * m2.cg + m1.jd * m2.dg + m1.je * m2.eg + m1.jf * m2.fg + m1.jg * m2.gg + m1.jh * m2.hg + m1.ji * m2.ig + m1.jj * m2.jg:    .jh = m1.ja * m2.ah + m1.jb * m2.bh + m1.jc * m2.ch + m1.jd * m2.dh + m1.je * m2.eh + m1.jf * m2.fh + m1.jg * m2.gh + m1.jh * m2.hh + m1.ji * m2.ih + m1.jj * m2.jh:    .ji = m1.ja * m2.ai + m1.jb * m2.bi + m1.jc * m2.ci + m1.jd * m2.di + m1.je * m2.ei + m1.jf * m2.fi + m1.jg * m2.gi + m1.jh * m2.hi + m1.ji * m2.ii + m1.jj * m2.ji:    .jj = m1.ja * m2.aj + m1.jb * m2.bj + m1.jc * m2.cj + m1.jd * m2.dj + m1.je * m2.ej + m1.jf * m2.fj + m1.jg * m2.gj + m1.jh * m2.hj + m1.ji * m2.ij + m1.jj * m2.jj
                
    End With
End Function

'Multiplikation Matrix mit Vektor
Public Function Matrix2_vmul(m As Matrix2, v As Vector2) As Vector2
    'Multipliziert eine 2x2 Matrix m mit einem 2er-Vektor
    With Matrix2_vmul:   .a = m.aa * v.a + m.ab * v.b
                         .b = m.ba * v.a + m.bb * v.b
    End With
End Function
Public Function Matrix3_vmul(m As Matrix3, v As Vector3) As Vector3
    'Multipliziert eine 3x3 Matrix m mit einem 3er-Vektor
    With Matrix3_vmul:   .a = m.aa * v.a + m.ab * v.b + m.ac * v.c
                         .b = m.ba * v.a + m.bb * v.b + m.bc * v.c
                         .c = m.ca * v.a + m.cb * v.b + m.cc * v.c
    End With
End Function
Public Function Matrix4_vmul(m As Matrix4, v As Vector4) As Vector4
    'Multipliziert eine 3x3 Matrix m mit einem 3er-Vektor
    With Matrix4_vmul:   .a = m.aa * v.a + m.ab * v.b + m.ac * v.c + m.ad * v.d
                         .b = m.ba * v.a + m.bb * v.b + m.bc * v.c + m.bd * v.d
                         .c = m.ca * v.a + m.cb * v.b + m.cc * v.c + m.cd * v.d
                         .d = m.da * v.a + m.db * v.b + m.dc * v.c + m.dd * v.d
    End With
End Function
Public Function Matrix5_vmul(m As Matrix5, v As Vector5) As Vector5
    'Multipliziert eine 3x3 Matrix m mit einem 3er-Vektor
    With Matrix5_vmul:   .a = m.aa * v.a + m.ab * v.b + m.ac * v.c + m.ad * v.d + m.ae * v.e
                         .b = m.ba * v.a + m.bb * v.b + m.bc * v.c + m.bd * v.d + m.be * v.e
                         .c = m.ca * v.a + m.cb * v.b + m.cc * v.c + m.cd * v.d + m.ce * v.e
                         .d = m.da * v.a + m.db * v.b + m.dc * v.c + m.dd * v.d + m.de * v.e
                         .e = m.ea * v.a + m.eb * v.b + m.ec * v.c + m.ed * v.d + m.ee * v.e
    End With
End Function
Public Function Matrix6_vmul(m As Matrix6, v As Vector6) As Vector6
    'Multipliziert eine 3x3 Matrix m mit einem 3er-Vektor
    With Matrix6_vmul:   .a = m.aa * v.a + m.ab * v.b + m.ac * v.c + m.ad * v.d + m.ae * v.e + m.af * v.f
                         .b = m.ba * v.a + m.bb * v.b + m.bc * v.c + m.bd * v.d + m.be * v.e + m.bf * v.f
                         .c = m.ca * v.a + m.cb * v.b + m.cc * v.c + m.cd * v.d + m.ce * v.e + m.cf * v.f
                         .d = m.da * v.a + m.db * v.b + m.dc * v.c + m.dd * v.d + m.de * v.e + m.df * v.f
                         .e = m.ea * v.a + m.eb * v.b + m.ec * v.c + m.ed * v.d + m.ee * v.e + m.ef * v.f
                         .f = m.fa * v.a + m.fb * v.b + m.fc * v.c + m.fd * v.d + m.fe * v.e + m.ff * v.f
    End With
End Function
Public Function Matrix7_vmul(m As Matrix7, v As Vector7) As Vector7
    'Multipliziert eine 3x3 Matrix m mit einem 3er-Vektor
    With Matrix7_vmul:   .a = m.aa * v.a + m.ab * v.b + m.ac * v.c + m.ad * v.d + m.ae * v.e + m.af * v.f + m.ag * v.g
                         .b = m.ba * v.a + m.bb * v.b + m.bc * v.c + m.bd * v.d + m.be * v.e + m.bf * v.f + m.bg * v.g
                         .c = m.ca * v.a + m.cb * v.b + m.cc * v.c + m.cd * v.d + m.ce * v.e + m.cf * v.f + m.cg * v.g
                         .d = m.da * v.a + m.db * v.b + m.dc * v.c + m.dd * v.d + m.de * v.e + m.df * v.f + m.dg * v.g
                         .e = m.ea * v.a + m.eb * v.b + m.ec * v.c + m.ed * v.d + m.ee * v.e + m.ef * v.f + m.eg * v.g
                         .f = m.fa * v.a + m.fb * v.b + m.fc * v.c + m.fd * v.d + m.fe * v.e + m.ff * v.f + m.fg * v.g
                         .g = m.ga * v.a + m.gb * v.b + m.gc * v.c + m.gd * v.d + m.ge * v.e + m.gf * v.f + m.gg * v.g
    End With
End Function
Public Function Matrix8_vmul(m As Matrix8, v As Vector8) As Vector8
    'Multipliziert eine 3x3 Matrix m mit einem 3er-Vektor
    With Matrix8_vmul:   .a = m.aa * v.a + m.ab * v.b + m.ac * v.c + m.ad * v.d + m.ae * v.e + m.af * v.f + m.ag * v.g + m.ah * v.h
                         .b = m.ba * v.a + m.bb * v.b + m.bc * v.c + m.bd * v.d + m.be * v.e + m.bf * v.f + m.bg * v.g + m.bh * v.h
                         .c = m.ca * v.a + m.cb * v.b + m.cc * v.c + m.cd * v.d + m.ce * v.e + m.cf * v.f + m.cg * v.g + m.ch * v.h
                         .d = m.da * v.a + m.db * v.b + m.dc * v.c + m.dd * v.d + m.de * v.e + m.df * v.f + m.dg * v.g + m.dh * v.h
                         .e = m.ea * v.a + m.eb * v.b + m.ec * v.c + m.ed * v.d + m.ee * v.e + m.ef * v.f + m.eg * v.g + m.eh * v.h
                         .f = m.fa * v.a + m.fb * v.b + m.fc * v.c + m.fd * v.d + m.fe * v.e + m.ff * v.f + m.fg * v.g + m.fh * v.h
                         .g = m.ga * v.a + m.gb * v.b + m.gc * v.c + m.gd * v.d + m.ge * v.e + m.gf * v.f + m.gg * v.g + m.gh * v.h
                         .h = m.ha * v.a + m.hb * v.b + m.hc * v.c + m.hd * v.d + m.he * v.e + m.hf * v.f + m.hg * v.g + m.hh * v.h
    End With
End Function
Public Function Matrix9_vmul(m As Matrix9, v As Vector9) As Vector9
    'Multipliziert eine 3x3 Matrix m mit einem 3er-Vektor
    With Matrix9_vmul:   .a = m.aa * v.a + m.ab * v.b + m.ac * v.c + m.ad * v.d + m.ae * v.e + m.af * v.f + m.ag * v.g + m.ah * v.h + m.ai * v.i
                         .b = m.ba * v.a + m.bb * v.b + m.bc * v.c + m.bd * v.d + m.be * v.e + m.bf * v.f + m.bg * v.g + m.bh * v.h + m.bi * v.i
                         .c = m.ca * v.a + m.cb * v.b + m.cc * v.c + m.cd * v.d + m.ce * v.e + m.cf * v.f + m.cg * v.g + m.ch * v.h + m.ci * v.i
                         .d = m.da * v.a + m.db * v.b + m.dc * v.c + m.dd * v.d + m.de * v.e + m.df * v.f + m.dg * v.g + m.dh * v.h + m.di * v.i
                         .e = m.ea * v.a + m.eb * v.b + m.ec * v.c + m.ed * v.d + m.ee * v.e + m.ef * v.f + m.eg * v.g + m.eh * v.h + m.ei * v.i
                         .f = m.fa * v.a + m.fb * v.b + m.fc * v.c + m.fd * v.d + m.fe * v.e + m.ff * v.f + m.fg * v.g + m.fh * v.h + m.fi * v.i
                         .g = m.ga * v.a + m.gb * v.b + m.gc * v.c + m.gd * v.d + m.ge * v.e + m.gf * v.f + m.gg * v.g + m.gh * v.h + m.gi * v.i
                         .h = m.ha * v.a + m.hb * v.b + m.hc * v.c + m.hd * v.d + m.he * v.e + m.hf * v.f + m.hg * v.g + m.hh * v.h + m.hi * v.i
                         .i = m.ia * v.a + m.ib * v.b + m.ic * v.c + m.id * v.d + m.ie * v.e + m.if * v.f + m.ig * v.g + m.ih * v.h + m.ii * v.i
    End With
End Function
Public Function Matrix10_vmul(m As Matrix10, v As Vector10) As Vector10
    'Multipliziert eine 3x3 Matrix m mit einem 3er-Vektor
    With Matrix10_vmul:   .a = m.aa * v.a + m.ab * v.b + m.ac * v.c + m.ad * v.d + m.ae * v.e + m.af * v.f + m.ag * v.g + m.ah * v.h + m.ai * v.i + m.aj * v.j
                          .b = m.ba * v.a + m.bb * v.b + m.bc * v.c + m.bd * v.d + m.be * v.e + m.bf * v.f + m.bg * v.g + m.bh * v.h + m.bi * v.i + m.bj * v.j
                          .c = m.ca * v.a + m.cb * v.b + m.cc * v.c + m.cd * v.d + m.ce * v.e + m.cf * v.f + m.cg * v.g + m.ch * v.h + m.ci * v.i + m.cj * v.j
                          .d = m.da * v.a + m.db * v.b + m.dc * v.c + m.dd * v.d + m.de * v.e + m.df * v.f + m.dg * v.g + m.dh * v.h + m.di * v.i + m.dj * v.j
                          .e = m.ea * v.a + m.eb * v.b + m.ec * v.c + m.ed * v.d + m.ee * v.e + m.ef * v.f + m.eg * v.g + m.eh * v.h + m.ei * v.i + m.ej * v.j
                          .f = m.fa * v.a + m.fb * v.b + m.fc * v.c + m.fd * v.d + m.fe * v.e + m.ff * v.f + m.fg * v.g + m.fh * v.h + m.fi * v.i + m.fj * v.j
                          .g = m.ga * v.a + m.gb * v.b + m.gc * v.c + m.gd * v.d + m.ge * v.e + m.gf * v.f + m.gg * v.g + m.gh * v.h + m.gi * v.i + m.gj * v.j
                          .h = m.ha * v.a + m.hb * v.b + m.hc * v.c + m.hd * v.d + m.he * v.e + m.hf * v.f + m.hg * v.g + m.hh * v.h + m.hi * v.i + m.hj * v.j
                          .i = m.ia * v.a + m.ib * v.b + m.ic * v.c + m.id * v.d + m.ie * v.e + m.if * v.f + m.ig * v.g + m.ih * v.h + m.ii * v.i + m.ij * v.j
                          .j = m.ia * v.a + m.jb * v.b + m.jc * v.c + m.jd * v.d + m.je * v.e + m.jf * v.f + m.jg * v.g + m.jh * v.h + m.ji * v.i + m.jj * v.j
    End With
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
Public Function Matrix7_IsEqual(m1 As Matrix7, m2 As Matrix7) As Boolean
    'Liefert True wenn beide 7x7-Matrizen gleich sind
    Matrix7_IsEqual = (m1.aa = m2.aa And m1.ab = m2.ab And m1.ac = m2.ac And m1.ad = m2.ad And m1.ae = m2.ae And m1.af = m2.af And m1.ag = m2.ag) And _
                      (m1.ba = m2.ba And m1.bb = m2.bb And m1.bc = m2.bc And m1.bd = m2.bd And m1.be = m2.be And m1.bf = m2.bf And m1.bg = m2.bg) And _
                      (m1.ca = m2.ca And m1.cb = m2.cb And m1.cc = m2.cc And m1.cd = m2.cd And m1.ce = m2.ce And m1.cf = m2.cf And m1.cg = m2.cg) And _
                      (m1.da = m2.da And m1.db = m2.db And m1.dc = m2.dc And m1.dd = m2.dd And m1.de = m2.de And m1.df = m2.df And m1.dg = m2.dg) And _
                      (m1.ea = m2.ea And m1.eb = m2.eb And m1.ec = m2.ec And m1.ed = m2.ed And m1.ee = m2.ee And m1.ef = m2.ef And m1.eg = m2.eg) And _
                      (m1.fa = m2.fa And m1.fb = m2.fb And m1.fc = m2.fc And m1.fd = m2.fd And m1.fe = m2.fe And m1.ff = m2.ff And m1.fg = m2.fg) And _
                      (m1.fa = m2.ga And m1.gb = m2.gb And m1.gc = m2.gc And m1.gd = m2.gd And m1.ge = m2.ge And m1.gf = m2.gf And m1.gg = m2.gg)
End Function
Public Function Matrix8_IsEqual(m1 As Matrix8, m2 As Matrix8) As Boolean
    'Liefert True wenn beide 8x8-Matrizen gleich sind
    Matrix8_IsEqual = (m1.aa = m2.aa And m1.ab = m2.ab And m1.ac = m2.ac And m1.ad = m2.ad And m1.ae = m2.ae And m1.af = m2.af And m1.ag = m2.ag And m1.ah = m2.ah) And _
                      (m1.ba = m2.ba And m1.bb = m2.bb And m1.bc = m2.bc And m1.bd = m2.bd And m1.be = m2.be And m1.bf = m2.bf And m1.bg = m2.bg And m1.bh = m2.bh) And _
                      (m1.ca = m2.ca And m1.cb = m2.cb And m1.cc = m2.cc And m1.cd = m2.cd And m1.ce = m2.ce And m1.cf = m2.cf And m1.cg = m2.cg And m1.ch = m2.ch) And _
                      (m1.da = m2.da And m1.db = m2.db And m1.dc = m2.dc And m1.dd = m2.dd And m1.de = m2.de And m1.df = m2.df And m1.dg = m2.dg And m1.dh = m2.dh) And _
                      (m1.ea = m2.ea And m1.eb = m2.eb And m1.ec = m2.ec And m1.ed = m2.ed And m1.ee = m2.ee And m1.ef = m2.ef And m1.eg = m2.eg And m1.eh = m2.eh) And _
                      (m1.fa = m2.fa And m1.fb = m2.fb And m1.fc = m2.fc And m1.fd = m2.fd And m1.fe = m2.fe And m1.ff = m2.ff And m1.fg = m2.fg And m1.fh = m2.fh) And _
                      (m1.fa = m2.ga And m1.gb = m2.gb And m1.gc = m2.gc And m1.gd = m2.gd And m1.ge = m2.ge And m1.gf = m2.gf And m1.gg = m2.gg And m1.gh = m2.gh) And _
                      (m1.ha = m2.ha And m1.hb = m2.hb And m1.hc = m2.hc And m1.hd = m2.hd And m1.he = m2.he And m1.hf = m2.hf And m1.hg = m2.hg And m1.hh = m2.hh)
End Function
Public Function Matrix9_IsEqual(m1 As Matrix9, m2 As Matrix9) As Boolean
    'Liefert True wenn beide 9x9-Matrizen gleich sind
    Matrix9_IsEqual = (m1.aa = m2.aa And m1.ab = m2.ab And m1.ac = m2.ac And m1.ad = m2.ad And m1.ae = m2.ae And m1.af = m2.af And m1.ag = m2.ag And m1.ah = m2.ah And m1.ai = m2.ai) And _
                      (m1.ba = m2.ba And m1.bb = m2.bb And m1.bc = m2.bc And m1.bd = m2.bd And m1.be = m2.be And m1.bf = m2.bf And m1.bg = m2.bg And m1.bh = m2.bh And m1.bi = m2.bi) And _
                      (m1.ca = m2.ca And m1.cb = m2.cb And m1.cc = m2.cc And m1.cd = m2.cd And m1.ce = m2.ce And m1.cf = m2.cf And m1.cg = m2.cg And m1.ch = m2.ch And m1.ci = m2.ci) And _
                      (m1.da = m2.da And m1.db = m2.db And m1.dc = m2.dc And m1.dd = m2.dd And m1.de = m2.de And m1.df = m2.df And m1.dg = m2.dg And m1.dh = m2.dh And m1.di = m2.di) And _
                      (m1.ea = m2.ea And m1.eb = m2.eb And m1.ec = m2.ec And m1.ed = m2.ed And m1.ee = m2.ee And m1.ef = m2.ef And m1.eg = m2.eg And m1.eh = m2.eh And m1.ei = m2.ei) And _
                      (m1.fa = m2.fa And m1.fb = m2.fb And m1.fc = m2.fc And m1.fd = m2.fd And m1.fe = m2.fe And m1.ff = m2.ff And m1.fg = m2.fg And m1.fh = m2.fh And m1.fi = m2.fi) And _
                      (m1.fa = m2.ga And m1.gb = m2.gb And m1.gc = m2.gc And m1.gd = m2.gd And m1.ge = m2.ge And m1.gf = m2.gf And m1.gg = m2.gg And m1.gh = m2.gh And m1.gi = m2.gi) And _
                      (m1.ha = m2.ha And m1.hb = m2.hb And m1.hc = m2.hc And m1.hd = m2.hd And m1.he = m2.he And m1.hf = m2.hf And m1.hg = m2.hg And m1.hh = m2.hh And m1.hi = m2.hi) And _
                      (m1.ia = m2.ia And m1.ib = m2.ib And m1.ic = m2.ic And m1.id = m2.id And m1.ie = m2.ie And m1.if = m2.if And m1.ig = m2.ig And m1.ih = m2.ih And m1.ii = m2.ii)
End Function
Public Function Matrix10_IsEqual(m1 As Matrix10, m2 As Matrix10) As Boolean
    'Liefert True wenn beide 10x10-Matrizen gleich sind
    Matrix10_IsEqual = (m1.aa = m2.aa And m1.ab = m2.ab And m1.ac = m2.ac And m1.ad = m2.ad And m1.ae = m2.ae And m1.af = m2.af And m1.ag = m2.ag And m1.ah = m2.ah And m1.ai = m2.ai And m1.aj = m2.aj) And _
                       (m1.ba = m2.ba And m1.bb = m2.bb And m1.bc = m2.bc And m1.bd = m2.bd And m1.be = m2.be And m1.bf = m2.bf And m1.bg = m2.bg And m1.bh = m2.bh And m1.bi = m2.bi And m1.bj = m2.bj) And _
                       (m1.ca = m2.ca And m1.cb = m2.cb And m1.cc = m2.cc And m1.cd = m2.cd And m1.ce = m2.ce And m1.cf = m2.cf And m1.cg = m2.cg And m1.ch = m2.ch And m1.ci = m2.ci And m1.cj = m2.cj) And _
                       (m1.da = m2.da And m1.db = m2.db And m1.dc = m2.dc And m1.dd = m2.dd And m1.de = m2.de And m1.df = m2.df And m1.dg = m2.dg And m1.dh = m2.dh And m1.di = m2.di And m1.dj = m2.dj) And _
                       (m1.ea = m2.ea And m1.eb = m2.eb And m1.ec = m2.ec And m1.ed = m2.ed And m1.ee = m2.ee And m1.ef = m2.ef And m1.eg = m2.eg And m1.eh = m2.eh And m1.ei = m2.ei And m1.ej = m2.ej) And _
                       (m1.fa = m2.fa And m1.fb = m2.fb And m1.fc = m2.fc And m1.fd = m2.fd And m1.fe = m2.fe And m1.ff = m2.ff And m1.fg = m2.fg And m1.fh = m2.fh And m1.fi = m2.fi And m1.fj = m2.fj) And _
                       (m1.ga = m2.ga And m1.gb = m2.gb And m1.gc = m2.gc And m1.gd = m2.gd And m1.ge = m2.ge And m1.gf = m2.gf And m1.gg = m2.gg And m1.gh = m2.gh And m1.gi = m2.gi And m1.gj = m2.gj) And _
                       (m1.ha = m2.ha And m1.hb = m2.hb And m1.hc = m2.hc And m1.hd = m2.hd And m1.he = m2.he And m1.hf = m2.hf And m1.hg = m2.hg And m1.hh = m2.hh And m1.hi = m2.hi And m1.hj = m2.hj) And _
                       (m1.ia = m2.ia And m1.ib = m2.ib And m1.ic = m2.ic And m1.id = m2.id And m1.ie = m2.ie And m1.if = m2.if And m1.ig = m2.ig And m1.ih = m2.ih And m1.ii = m2.ii And m1.ij = m2.ij) And _
                       (m1.ja = m2.ja And m1.jb = m2.jb And m1.jc = m2.jc And m1.jd = m2.jd And m1.je = m2.je And m1.jf = m2.jf And m1.jg = m2.jg And m1.jh = m2.jh And m1.ji = m2.ji And m1.jj = m2.jj)
End Function

'Transponierte Matrix
Public Function Matrix2_tra(m As Matrix2) As Matrix2
    'Erzeugt die Transponierte aus einer 2x2-Matrix
    With Matrix2_tra:   .aa = m.aa:    .ab = m.ba
                        .ba = m.ab:    .bb = m.bb
    End With
End Function
Public Function Matrix3_tra(m As Matrix3) As Matrix3
    'Erzeugt die Transponierte aus einer 3x3-Matrix
    With Matrix3_tra:   .aa = m.aa:    .ab = m.ba:    .ac = m.ca
                        .ba = m.ab:    .bb = m.bb:    .bc = m.cb
                        .ca = m.ac:    .cb = m.bc:    .cc = m.cc
    End With
End Function
Public Function Matrix4_tra(m As Matrix4) As Matrix4
    'Erzeugt die Transponierte aus einer 4x4-Matrix
    With Matrix4_tra:   .aa = m.aa:    .ab = m.ba:    .ac = m.ca:    .ad = m.da
                        .ba = m.ab:    .bb = m.bb:    .bc = m.cb:    .bd = m.db
                        .ca = m.ac:    .cb = m.bc:    .cc = m.cc:    .cd = m.dc
                        .da = m.ad:    .db = m.bd:    .dc = m.cd:    .dd = m.dd
    End With
End Function
Public Function Matrix5_tra(m As Matrix5) As Matrix5
    'Erzeugt die Transponierte aus einer 5x5-Matrix
    With Matrix5_tra:   .aa = m.aa:    .ab = m.ba:    .ac = m.ca:    .ad = m.da:    .ae = m.ea
                        .ba = m.ab:    .bb = m.bb:    .bc = m.cb:    .bd = m.db:    .be = m.eb
                        .ca = m.ac:    .cb = m.bc:    .cc = m.cc:    .cd = m.dc:    .ce = m.ec
                        .da = m.ad:    .db = m.bd:    .dc = m.cd:    .dd = m.dd:    .de = m.ed
                        .ea = m.ae:    .eb = m.be:    .ec = m.ce:    .ed = m.de:    .ee = m.ee
    End With
End Function
Public Function Matrix6_tra(m As Matrix6) As Matrix6
    'Erzeugt die Transponierte aus einer 6x6-Matrix
    With Matrix6_tra:   .aa = m.aa:    .ab = m.ba:    .ac = m.ca:    .ad = m.da:    .ae = m.ea:    .af = m.fa
                        .ba = m.ab:    .bb = m.bb:    .bc = m.cb:    .bd = m.db:    .be = m.eb:    .bf = m.fb
                        .ca = m.ac:    .cb = m.bc:    .cc = m.cc:    .cd = m.dc:    .ce = m.ec:    .cf = m.fc
                        .da = m.ad:    .db = m.bd:    .dc = m.cd:    .dd = m.dd:    .de = m.ed:    .df = m.fd
                        .ea = m.ae:    .eb = m.be:    .ec = m.ce:    .ed = m.de:    .ee = m.ee:    .ef = m.fe
                        .fa = m.af:    .fb = m.bf:    .fc = m.cf:    .fd = m.df:    .fe = m.ef:    .ff = m.ff
    End With
End Function
Public Function Matrix7_tra(m As Matrix7) As Matrix7
    'Erzeugt die Transponierte aus einer 6x6-Matrix
    With Matrix7_tra:   .aa = m.aa:    .ab = m.ba:    .ac = m.ca:    .ad = m.da:    .ae = m.ea:    .af = m.fa:    .ag = m.ga
                        .ba = m.ab:    .bb = m.bb:    .bc = m.cb:    .bd = m.db:    .be = m.eb:    .bf = m.fb:    .bg = m.gb
                        .ca = m.ac:    .cb = m.bc:    .cc = m.cc:    .cd = m.dc:    .ce = m.ec:    .cf = m.fc:    .cg = m.gc
                        .da = m.ad:    .db = m.bd:    .dc = m.cd:    .dd = m.dd:    .de = m.ed:    .df = m.fd:    .dg = m.gd
                        .ea = m.ae:    .eb = m.be:    .ec = m.ce:    .ed = m.de:    .ee = m.ee:    .ef = m.fe:    .eg = m.ge
                        .fa = m.af:    .fb = m.bf:    .fc = m.cf:    .fd = m.df:    .fe = m.ef:    .ff = m.ff:    .fg = m.gf
                        .ga = m.ag:    .gb = m.bg:    .gc = m.cg:    .gd = m.dg:    .ge = m.eg:    .gf = m.fg:    .gg = m.gg
    End With
End Function
Public Function Matrix8_tra(m As Matrix8) As Matrix8
    'Erzeugt die Transponierte aus einer 6x6-Matrix
    With Matrix8_tra:   .aa = m.aa:    .ab = m.ba:    .ac = m.ca:    .ad = m.da:    .ae = m.ea:    .af = m.fa:    .ag = m.ga:    .ah = m.ha
                        .ba = m.ab:    .bb = m.bb:    .bc = m.cb:    .bd = m.db:    .be = m.eb:    .bf = m.fb:    .bg = m.gb:    .bh = m.hb
                        .ca = m.ac:    .cb = m.bc:    .cc = m.cc:    .cd = m.dc:    .ce = m.ec:    .cf = m.fc:    .cg = m.gc:    .ch = m.hc
                        .da = m.ad:    .db = m.bd:    .dc = m.cd:    .dd = m.dd:    .de = m.ed:    .df = m.fd:    .dg = m.gd:    .dh = m.hd
                        .ea = m.ae:    .eb = m.be:    .ec = m.ce:    .ed = m.de:    .ee = m.ee:    .ef = m.fe:    .eg = m.ge:    .eh = m.he
                        .fa = m.af:    .fb = m.bf:    .fc = m.cf:    .fd = m.df:    .fe = m.ef:    .ff = m.ff:    .fg = m.gf:    .fh = m.hf
                        .ga = m.ag:    .gb = m.bg:    .gc = m.cg:    .gd = m.dg:    .ge = m.eg:    .gf = m.fg:    .gg = m.gg:    .gh = m.hg
                        .ha = m.ah:    .hb = m.bh:    .hc = m.ch:    .hd = m.dh:    .he = m.eh:    .hf = m.fh:    .hg = m.gh:    .hh = m.hh
    End With
End Function
Public Function Matrix9_tra(m As Matrix9) As Matrix9
    'Erzeugt die Transponierte aus einer 6x6-Matrix
    With Matrix9_tra:   .aa = m.aa:    .ab = m.ba:    .ac = m.ca:    .ad = m.da:    .ae = m.ea:    .af = m.fa:    .ag = m.ga:    .ah = m.ha:    .ai = m.ia
                        .ba = m.ab:    .bb = m.bb:    .bc = m.cb:    .bd = m.db:    .be = m.eb:    .bf = m.fb:    .bg = m.gb:    .bh = m.hb:    .bi = m.ib
                        .ca = m.ac:    .cb = m.bc:    .cc = m.cc:    .cd = m.dc:    .ce = m.ec:    .cf = m.fc:    .cg = m.gc:    .ch = m.hc:    .ci = m.ic
                        .da = m.ad:    .db = m.bd:    .dc = m.cd:    .dd = m.dd:    .de = m.ed:    .df = m.fd:    .dg = m.gd:    .dh = m.hd:    .di = m.id
                        .ea = m.ae:    .eb = m.be:    .ec = m.ce:    .ed = m.de:    .ee = m.ee:    .ef = m.fe:    .eg = m.ge:    .eh = m.he:    .ei = m.ie
                        .fa = m.af:    .fb = m.bf:    .fc = m.cf:    .fd = m.df:    .fe = m.ef:    .ff = m.ff:    .fg = m.gf:    .fh = m.hf:    .fi = m.if
                        .ga = m.ag:    .gb = m.bg:    .gc = m.cg:    .gd = m.dg:    .ge = m.eg:    .gf = m.fg:    .gg = m.gg:    .gh = m.hg:    .gi = m.ig
                        .ha = m.ah:    .hb = m.bh:    .hc = m.ch:    .hd = m.dh:    .he = m.eh:    .hf = m.fh:    .hg = m.gh:    .hh = m.hh:    .hi = m.ih
                        .ia = m.ai:    .ib = m.bi:    .ic = m.ci:    .id = m.di:    .ie = m.ei:    .if = m.fi:    .ig = m.gi:    .ih = m.hi:    .ii = m.ii
    End With
End Function
Public Function Matrix10_tra(m As Matrix10) As Matrix10
    'Erzeugt die Transponierte aus einer 6x6-Matrix
    With Matrix10_tra:   .aa = m.aa:    .ab = m.ba:    .ac = m.ca:    .ad = m.da:    .ae = m.ea:    .af = m.fa:    .ag = m.ga:    .ah = m.ha:    .ai = m.ia:    .aj = m.ja
                         .ba = m.ab:    .bb = m.bb:    .bc = m.cb:    .bd = m.db:    .be = m.eb:    .bf = m.fb:    .bg = m.gb:    .bh = m.hb:    .bi = m.ib:    .bj = m.jb
                         .ca = m.ac:    .cb = m.bc:    .cc = m.cc:    .cd = m.dc:    .ce = m.ec:    .cf = m.fc:    .cg = m.gc:    .ch = m.hc:    .ci = m.ic:    .cj = m.jc
                         .da = m.ad:    .db = m.bd:    .dc = m.cd:    .dd = m.dd:    .de = m.ed:    .df = m.fd:    .dg = m.gd:    .dh = m.hd:    .di = m.id:    .dj = m.jd
                         .ea = m.ae:    .eb = m.be:    .ec = m.ce:    .ed = m.de:    .ee = m.ee:    .ef = m.fe:    .eg = m.ge:    .eh = m.he:    .ei = m.ie:    .ej = m.je
                         .fa = m.af:    .fb = m.bf:    .fc = m.cf:    .fd = m.df:    .fe = m.ef:    .ff = m.ff:    .fg = m.gf:    .fh = m.hf:    .fi = m.if:    .fj = m.jf
                         .ga = m.ag:    .gb = m.bg:    .gc = m.cg:    .gd = m.dg:    .ge = m.eg:    .gf = m.fg:    .gg = m.gg:    .gh = m.hg:    .gi = m.ig:    .gj = m.jg
                         .ha = m.ah:    .hb = m.bh:    .hc = m.ch:    .hd = m.dh:    .he = m.eh:    .hf = m.fh:    .hg = m.gh:    .hh = m.hh:    .hi = m.ih:    .hj = m.jh
                         .ia = m.ai:    .ib = m.bi:    .ic = m.ci:    .id = m.di:    .ie = m.ei:    .if = m.fi:    .ig = m.gi:    .ih = m.hi:    .ii = m.ii:    .ij = m.ji
                         .ja = m.aj:    .jb = m.bj:    .jc = m.cj:    .jd = m.dj:    .je = m.ej:    .jf = m.fj:    .jg = m.gj:    .jh = m.hj:    .ji = m.ij:    .jj = m.jj
    End With
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
            'a_out(i, j) = DblParse(sa(j))
            a_out(j, i) = DblParse(sa(j)) 'orig ?
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
Public Function Matrix7_Parse(t As String) As Matrix7
    Dim mRows As Long: mRows = 7
    Dim nCols As Long: nCols = 7
    Dim a() As Double: a = Matrix_Parse(t, mRows, nCols)
    RtlMoveMemory Matrix7_Parse, a(0, 0), mRows * nCols * 8
End Function
Public Function Matrix8_Parse(t As String) As Matrix8
    Dim mRows As Long: mRows = 8
    Dim nCols As Long: nCols = 8
    Dim a() As Double: a = Matrix_Parse(t, mRows, nCols)
    RtlMoveMemory Matrix8_Parse, a(0, 0), mRows * nCols * 8
End Function
Public Function Matrix9_Parse(t As String) As Matrix9
    Dim mRows As Long: mRows = 9
    Dim nCols As Long: nCols = 9
    Dim a() As Double: a = Matrix_Parse(t, mRows, nCols)
    RtlMoveMemory Matrix9_Parse, a(0, 0), mRows * nCols * 8
End Function
Public Function Matrix10_Parse(t As String) As Matrix10
    Dim mRows As Long: mRows = 10
    Dim nCols As Long: nCols = 10
    Dim a() As Double: a = Matrix_Parse(t, mRows, nCols)
    RtlMoveMemory Matrix10_Parse, a(0, 0), mRows * nCols * 8
End Function

'Umwandeln in 2d-Array
Public Function Matrix2_ToArr(m As Matrix2) As Double()
    Dim d(0 To 1, 0 To 1) As Double
    With m
        d(0, 0) = .aa: d(0, 1) = .ab
        d(1, 0) = .ba: d(1, 1) = .bb
    End With
    Matrix2_ToArr = d
End Function
Public Function Matrix3_ToArr(m As Matrix3) As Double()
    Dim d(0 To 2, 0 To 2) As Double
    With m
        d(0, 0) = .aa: d(0, 1) = .ab: d(0, 2) = .ac
        d(1, 0) = .ba: d(1, 1) = .bb: d(1, 2) = .bc
        d(2, 0) = .ca: d(2, 1) = .cb: d(2, 2) = .cc
    End With
    Matrix3_ToArr = d
End Function
Public Function Matrix4_ToArr(m As Matrix4) As Double()
    Dim d(0 To 3, 0 To 3) As Double
    With m
        d(0, 0) = .aa: d(0, 1) = .ab: d(0, 2) = .ac: d(0, 3) = .ad
        d(1, 0) = .ba: d(1, 1) = .bb: d(1, 2) = .bc: d(1, 3) = .bd
        d(2, 0) = .ca: d(2, 1) = .cb: d(2, 2) = .cc: d(2, 3) = .cd
        d(3, 0) = .da: d(3, 1) = .db: d(3, 2) = .dc: d(3, 3) = .dd
    End With
    Matrix4_ToArr = d
End Function
Public Function Matrix5_ToArr(m As Matrix5) As Double()
    Dim d(0 To 4, 0 To 4) As Double
    With m
        d(0, 0) = .aa: d(0, 1) = .ab: d(0, 2) = .ac: d(0, 3) = .ad: d(0, 4) = .ae
        d(1, 0) = .ba: d(1, 1) = .bb: d(1, 2) = .bc: d(1, 3) = .bd: d(1, 4) = .be
        d(2, 0) = .ca: d(2, 1) = .cb: d(2, 2) = .cc: d(2, 3) = .cd: d(2, 4) = .ce
        d(3, 0) = .da: d(3, 1) = .db: d(3, 2) = .dc: d(3, 3) = .dd: d(3, 4) = .de
        d(4, 0) = .ea: d(4, 1) = .eb: d(4, 2) = .ec: d(4, 3) = .ed: d(4, 4) = .ee
    End With
    Matrix5_ToArr = d
End Function
Public Function Matrix6_ToArr(m As Matrix6) As Double()
    Dim d(0 To 5, 0 To 5) As Double
    With m
        d(0, 0) = .aa: d(0, 1) = .ab: d(0, 2) = .ac: d(0, 3) = .ad: d(0, 4) = .ae: d(0, 5) = .af
        d(1, 0) = .ba: d(1, 1) = .bb: d(1, 2) = .bc: d(1, 3) = .bd: d(1, 4) = .be: d(1, 5) = .bf
        d(2, 0) = .ca: d(2, 1) = .cb: d(2, 2) = .cc: d(2, 3) = .cd: d(2, 4) = .ce: d(2, 5) = .cf
        d(3, 0) = .da: d(3, 1) = .db: d(3, 2) = .dc: d(3, 3) = .dd: d(3, 4) = .de: d(3, 5) = .df
        d(4, 0) = .ea: d(4, 1) = .eb: d(4, 2) = .ec: d(4, 3) = .ed: d(4, 4) = .ee: d(4, 5) = .ef
        d(5, 0) = .fa: d(5, 1) = .fb: d(5, 2) = .fc: d(5, 3) = .fd: d(5, 4) = .fe: d(5, 5) = .ff
    End With
    Matrix6_ToArr = d
End Function
Public Function Matrix7_ToArr(m As Matrix7) As Double()
    Dim d(0 To 6, 0 To 6) As Double
    With m
        d(0, 0) = .aa: d(0, 1) = .ab: d(0, 2) = .ac: d(0, 3) = .ad: d(0, 4) = .ae: d(0, 5) = .af: d(0, 6) = .ag
        d(1, 0) = .ba: d(1, 1) = .bb: d(1, 2) = .bc: d(1, 3) = .bd: d(1, 4) = .be: d(1, 5) = .bf: d(1, 6) = .bg
        d(2, 0) = .ca: d(2, 1) = .cb: d(2, 2) = .cc: d(2, 3) = .cd: d(2, 4) = .ce: d(2, 5) = .cf: d(2, 6) = .cg
        d(3, 0) = .da: d(3, 1) = .db: d(3, 2) = .dc: d(3, 3) = .dd: d(3, 4) = .de: d(3, 5) = .df: d(3, 6) = .dg
        d(4, 0) = .ea: d(4, 1) = .eb: d(4, 2) = .ec: d(4, 3) = .ed: d(4, 4) = .ee: d(4, 5) = .ef: d(4, 6) = .eg
        d(5, 0) = .fa: d(5, 1) = .fb: d(5, 2) = .fc: d(5, 3) = .fd: d(5, 4) = .fe: d(5, 5) = .ff: d(5, 6) = .fg
        d(6, 0) = .ga: d(6, 1) = .gb: d(6, 2) = .gc: d(6, 3) = .gd: d(6, 4) = .ge: d(6, 5) = .gf: d(6, 6) = .gg
    End With
    Matrix7_ToArr = d
End Function
Public Function Matrix8_ToArr(m As Matrix8) As Double()
    Dim d(0 To 7, 0 To 7) As Double
    With m
        d(0, 0) = .aa: d(0, 1) = .ab: d(0, 2) = .ac: d(0, 3) = .ad: d(0, 4) = .ae: d(0, 5) = .af: d(0, 6) = .ag: d(0, 7) = .ah
        d(1, 0) = .ba: d(1, 1) = .bb: d(1, 2) = .bc: d(1, 3) = .bd: d(1, 4) = .be: d(1, 5) = .bf: d(1, 6) = .bg: d(1, 7) = .bh
        d(2, 0) = .ca: d(2, 1) = .cb: d(2, 2) = .cc: d(2, 3) = .cd: d(2, 4) = .ce: d(2, 5) = .cf: d(2, 6) = .cg: d(2, 7) = .ch
        d(3, 0) = .da: d(3, 1) = .db: d(3, 2) = .dc: d(3, 3) = .dd: d(3, 4) = .de: d(3, 5) = .df: d(3, 6) = .dg: d(3, 7) = .dh
        d(4, 0) = .ea: d(4, 1) = .eb: d(4, 2) = .ec: d(4, 3) = .ed: d(4, 4) = .ee: d(4, 5) = .ef: d(4, 6) = .eg: d(4, 7) = .eh
        d(5, 0) = .fa: d(5, 1) = .fb: d(5, 2) = .fc: d(5, 3) = .fd: d(5, 4) = .fe: d(5, 5) = .ff: d(5, 6) = .fg: d(5, 7) = .fh
        d(6, 0) = .ga: d(6, 1) = .gb: d(6, 2) = .gc: d(6, 3) = .gd: d(6, 4) = .ge: d(6, 5) = .gf: d(6, 6) = .gg: d(6, 7) = .gh
        d(7, 0) = .ha: d(7, 1) = .hb: d(7, 2) = .hc: d(7, 3) = .hd: d(7, 4) = .he: d(7, 5) = .hf: d(7, 6) = .hg: d(7, 7) = .hh
    End With
    Matrix8_ToArr = d
End Function
Public Function Matrix9_ToArr(m As Matrix9) As Double()
    Dim d(0 To 8, 0 To 8) As Double
    With m
        d(0, 0) = .aa: d(0, 1) = .ab: d(0, 2) = .ac: d(0, 3) = .ad: d(0, 4) = .ae: d(0, 5) = .af: d(0, 6) = .ag: d(0, 7) = .ah: d(0, 8) = .ai
        d(1, 0) = .ba: d(1, 1) = .bb: d(1, 2) = .bc: d(1, 3) = .bd: d(1, 4) = .be: d(1, 5) = .bf: d(1, 6) = .bg: d(1, 7) = .bh: d(1, 8) = .bi
        d(2, 0) = .ca: d(2, 1) = .cb: d(2, 2) = .cc: d(2, 3) = .cd: d(2, 4) = .ce: d(2, 5) = .cf: d(2, 6) = .cg: d(2, 7) = .ch: d(2, 8) = .ci
        d(3, 0) = .da: d(3, 1) = .db: d(3, 2) = .dc: d(3, 3) = .dd: d(3, 4) = .de: d(3, 5) = .df: d(3, 6) = .dg: d(3, 7) = .dh: d(3, 8) = .di
        d(4, 0) = .ea: d(4, 1) = .eb: d(4, 2) = .ec: d(4, 3) = .ed: d(4, 4) = .ee: d(4, 5) = .ef: d(4, 6) = .eg: d(4, 7) = .eh: d(4, 8) = .ei
        d(5, 0) = .fa: d(5, 1) = .fb: d(5, 2) = .fc: d(5, 3) = .fd: d(5, 4) = .fe: d(5, 5) = .ff: d(5, 6) = .fg: d(5, 7) = .fh: d(5, 8) = .fi
        d(6, 0) = .ga: d(6, 1) = .gb: d(6, 2) = .gc: d(6, 3) = .gd: d(6, 4) = .ge: d(6, 5) = .gf: d(6, 6) = .gg: d(6, 7) = .gh: d(6, 8) = .gi
        d(7, 0) = .ha: d(7, 1) = .hb: d(7, 2) = .hc: d(7, 3) = .hd: d(7, 4) = .he: d(7, 5) = .hf: d(7, 6) = .hg: d(7, 7) = .hh: d(7, 8) = .hi
        d(8, 0) = .ia: d(8, 1) = .ib: d(8, 2) = .ic: d(8, 3) = .id: d(8, 4) = .ie: d(8, 5) = .if: d(8, 6) = .ig: d(8, 7) = .ih: d(8, 8) = .ii
    End With
    Matrix9_ToArr = d
End Function
Public Function Matrix10_ToArr(m As Matrix10) As Double()
    Dim d(0 To 9, 0 To 9) As Double
    With m
        d(0, 0) = .aa: d(0, 1) = .ab: d(0, 2) = .ac: d(0, 3) = .ad: d(0, 4) = .ae: d(0, 5) = .af: d(0, 6) = .ag: d(0, 7) = .ah: d(0, 8) = .ai: d(0, 9) = .aj
        d(1, 0) = .ba: d(1, 1) = .bb: d(1, 2) = .bc: d(1, 3) = .bd: d(1, 4) = .be: d(1, 5) = .bf: d(1, 6) = .bg: d(1, 7) = .bh: d(1, 8) = .bi: d(1, 9) = .bj
        d(2, 0) = .ca: d(2, 1) = .cb: d(2, 2) = .cc: d(2, 3) = .cd: d(2, 4) = .ce: d(2, 5) = .cf: d(2, 6) = .cg: d(2, 7) = .ch: d(2, 8) = .ci: d(2, 9) = .cj
        d(3, 0) = .da: d(3, 1) = .db: d(3, 2) = .dc: d(3, 3) = .dd: d(3, 4) = .de: d(3, 5) = .df: d(3, 6) = .dg: d(3, 7) = .dh: d(3, 8) = .di: d(3, 9) = .dj
        d(4, 0) = .ea: d(4, 1) = .eb: d(4, 2) = .ec: d(4, 3) = .ed: d(4, 4) = .ee: d(4, 5) = .ef: d(4, 6) = .eg: d(4, 7) = .eh: d(4, 8) = .ei: d(4, 9) = .ej
        d(5, 0) = .fa: d(5, 1) = .fb: d(5, 2) = .fc: d(5, 3) = .fd: d(5, 4) = .fe: d(5, 5) = .ff: d(5, 6) = .fg: d(5, 7) = .fh: d(5, 8) = .fi: d(5, 9) = .fj
        d(6, 0) = .ga: d(6, 1) = .gb: d(6, 2) = .gc: d(6, 3) = .gd: d(6, 4) = .ge: d(6, 5) = .gf: d(6, 6) = .gg: d(6, 7) = .gh: d(6, 8) = .gi: d(6, 9) = .gj
        d(7, 0) = .ha: d(7, 1) = .hb: d(7, 2) = .hc: d(7, 3) = .hd: d(7, 4) = .he: d(7, 5) = .hf: d(7, 6) = .hg: d(7, 7) = .hh: d(7, 8) = .hi: d(7, 9) = .hj
        d(8, 0) = .ia: d(8, 1) = .ib: d(8, 2) = .ic: d(8, 3) = .id: d(8, 4) = .ie: d(8, 5) = .if: d(8, 6) = .ig: d(8, 7) = .ih: d(8, 8) = .ii: d(8, 9) = .ij
        d(9, 0) = .ja: d(9, 1) = .jb: d(9, 2) = .jc: d(9, 3) = .jd: d(9, 4) = .je: d(9, 5) = .jf: d(9, 6) = .jg: d(9, 7) = .jh: d(9, 8) = .ji: d(9, 9) = .jj
    End With
    Matrix10_ToArr = d
End Function


'Ausgabefunktionen
Public Function Matrix2_ToStr(m As Matrix2) As String
    Matrix2_ToStr = MatrixA_ToStr(Matrix2_ToArr(m), 2, 2)
    'Matrix2_ToStr = Matrix_ToStr(VarPtr(m), 2, 2)
End Function
Public Function Matrix3_ToStr(m As Matrix3) As String
    Matrix3_ToStr = MatrixA_ToStr(Matrix3_ToArr(m), 3, 3)
    'Matrix3_ToStr = Matrix_ToStr(VarPtr(m), 3, 3)
End Function
Public Function Matrix4_ToStr(m As Matrix4) As String
    Matrix4_ToStr = MatrixA_ToStr(Matrix4_ToArr(m), 4, 4)
    'Matrix4_ToStr = Matrix_ToStr(VarPtr(m), 4, 4)
End Function
Public Function Matrix5_ToStr(m As Matrix5) As String
    Matrix5_ToStr = MatrixA_ToStr(Matrix5_ToArr(m), 5, 5)
    'Matrix5_ToStr = Matrix_ToStr(VarPtr(m), 5, 5)
End Function
Public Function Matrix6_ToStr(m As Matrix6) As String
    Matrix6_ToStr = MatrixA_ToStr(Matrix6_ToArr(m), 6, 6)
    'Matrix6_ToStr = Matrix_ToStr(VarPtr(m), 6, 6)
End Function
Public Function Matrix7_ToStr(m As Matrix7) As String
    Matrix7_ToStr = MatrixA_ToStr(Matrix7_ToArr(m), 7, 7)
    'Matrix7_ToStr = Matrix_ToStr(VarPtr(m), 7, 7)
End Function
Public Function Matrix8_ToStr(m As Matrix8) As String
    Matrix8_ToStr = MatrixA_ToStr(Matrix8_ToArr(m), 8, 8)
    'Matrix8_ToStr = Matrix_ToStr(VarPtr(m), 8, 8)
End Function
Public Function Matrix9_ToStr(m As Matrix9) As String
    Matrix9_ToStr = MatrixA_ToStr(Matrix9_ToArr(m), 9, 9)
    'Matrix9_ToStr = Matrix_ToStr(VarPtr(m), 9, 9)
End Function
Public Function Matrix10_ToStr(m As Matrix10) As String
    Matrix10_ToStr = MatrixA_ToStr(Matrix10_ToArr(m), 10, 10)
    'Matrix10_ToStr = Matrix_ToStr(VarPtr(m), 10, 10)
End Function

Public Function MatrixA_ToStr(a() As Double, ByVal mRows As Long, ByVal nCols As Long) As String
'häöääähhh???? watn data fürn schiiiiiid
'oouuuuuu halt ja weil da wird ja untereinander ausgerichtet!!!

    Dim s As String ': s = ""
    Dim sl As String, vs As String, vsa() As String
    ReDim msa(0 To mRows - 1) As String
    Dim i As Long, j As Long
    ReDim ca(0 To mRows - 1) As Double
    For j = 0 To nCols - 1
        For i = 0 To mRows - 1
            'ca(i) = a(j, i)
            ca(i) = a(i, j)
        Next
        vsa = Split(VectorFormat(ca, 0), vbCrLf)
        For i = 0 To mRows - 1
            msa(i) = msa(i) & " " & vsa(i)
        Next
    Next
    MatrixA_ToStr = Join(msa, vbCrLf) 's

'    For i = 0 To mRows
'        For j = 0 To nCols
'            s = s & a(j, i)
'            If j < nCols Then s = s & " "
'        Next
'        If i < mRows Then s = s & vbCrLf
'    Next
'    Matrix_ToStr = s


End Function

'Function MatA_ToStr(ByVal pMat As Long, ByVal mRows As Long, ByVal nCols As Long)
'    ReDim ma(0 To mRows - 1, 0 To nCols - 1) As Double
'    RtlMoveMemory ma(0, 0), ByVal pMat, mRows * nCols * 8
'    Dim i As Long, j As Long
'    Dim s As String
'    mRows = mRows - 1
'    nCols = nCols - 1
'    For i = 0 To mRows
'        For j = 0 To nCols
'            s = s & ma(j, i)
'            If j < nCols Then s = s & " "
'        Next
'        If i < mRows Then s = s & vbCrLf
'    Next
'    MatA_ToStr = s
'End Function

Public Function Matrix_ToStr(ByVal p_Matrix As Long, ByVal mRows As Long, ByVal nCols As Long) As String
    ReDim ma(0 To mRows - 1, 0 To nCols - 1) As Double
    RtlMoveMemory ma(0, 0), ByVal p_Matrix, mRows * nCols * 8
    Dim i As Long, j As Long
    Dim s As String
    mRows = mRows - 1
    nCols = nCols - 1
    For i = 0 To mRows
        For j = 0 To nCols
            s = s & ma(j, i)
            If j < nCols Then s = s & " "
        Next
        If i < mRows Then s = s & vbCrLf
    Next
    Matrix_ToStr = s



'    'die allgemeine mathematische Anordnung  ist a(iZeile, jSpalte)
'    'vgl die Speicheranordnung von VB-Arrays ist a(jSpalte, iZeile)
'    If mRows = 0 Or nCols = 0 Then Exit Function
'    ReDim a(0 To nCols - 1, 0 To mRows - 1) As Double
'    RtlMoveMemory a(0, 0), ByVal p_Matrix, mRows * nCols * 8
'    Matrix_ToStr = MatrixA_ToStr(a, mRows, nCols)
End Function
'Zeile 588



'Fortgeschrittene Matrix-Operationen
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
Public Function Matrix7_det(m As Matrix7) As Double
    'Berechnet die Determinante einer 7x7-Matrix
    With m
        Matrix7_det = .aa * Matrix6_det(Mat6(.bb, .bc, .bd, .be, .bf, .bg, .cb, .cc, .cd, .ce, .cf, .cg, .db, .dc, .dd, .de, .df, .dg, .eb, .ec, .ed, .ee, .ef, .eg, .fb, .fc, .fd, .fe, .ff, .fg, .gb, .gc, .gd, .ge, .gf, .gg)) _
                    - .ab * Matrix6_det(Mat6(.ba, .bc, .bd, .be, .bf, .bg, .ca, .cc, .cd, .ce, .cf, .cg, .da, .dc, .dd, .de, .df, .dg, .ea, .ec, .ed, .ee, .ef, .eg, .fa, .fc, .fd, .fe, .ff, .fg, .ga, .gc, .gd, .ge, .gf, .gg)) _
                    + .ac * Matrix6_det(Mat6(.ba, .bb, .bd, .be, .bf, .bg, .ca, .cb, .cd, .ce, .cf, .cg, .da, .db, .dd, .de, .df, .dg, .ea, .eb, .ed, .ee, .ef, .eg, .fa, .fb, .fd, .fe, .ff, .fg, .ga, .gb, .gd, .ge, .gf, .gg)) _
                    - .ad * Matrix6_det(Mat6(.ba, .bb, .bc, .be, .bf, .bg, .ca, .cb, .cc, .ce, .cf, .cg, .da, .db, .dc, .de, .df, .dg, .ea, .eb, .ec, .ee, .ef, .eg, .fa, .fb, .fc, .fe, .ff, .fg, .ga, .gb, .gc, .ge, .gf, .gg)) _
                    + .ae * Matrix6_det(Mat6(.ba, .bb, .bc, .bd, .bf, .bg, .ca, .cb, .cc, .cd, .cf, .cg, .da, .db, .dc, .dd, .df, .dg, .ea, .eb, .ec, .ed, .ef, .eg, .fa, .fb, .fc, .fd, .ff, .fg, .ga, .gb, .gc, .gd, .gf, .gg)) _
                    - .af * Matrix6_det(Mat6(.ba, .bb, .bc, .bd, .be, .bg, .ca, .cb, .cc, .cd, .ce, .cg, .da, .db, .dc, .dd, .de, .dg, .ea, .eb, .ec, .ed, .ee, .eg, .fa, .fb, .fc, .fd, .fe, .fg, .ga, .gb, .gc, .gd, .ge, .gg)) _
                    + .ag * Matrix6_det(Mat6(.ba, .bb, .bc, .bd, .be, .bf, .ca, .cb, .cc, .cd, .ce, .cf, .da, .db, .dc, .dd, .de, .df, .ea, .eb, .ec, .ed, .ee, .ef, .fa, .fb, .fc, .fd, .fe, .ff, .ga, .gb, .gc, .gd, .ge, .gf))
    End With
End Function
Public Function Matrix8_det(m As Matrix8) As Double
    'Berechnet die Determinante einer 8x8-Matrix
    With m
        Matrix8_det = .aa * Matrix7_det(Mat7(.bb, .bc, .bd, .be, .bf, .bg, .bh, .cb, .cc, .cd, .ce, .cf, .cg, .ch, .db, .dc, .dd, .de, .df, .dg, .dh, .eb, .ec, .ed, .ee, .ef, .eg, .eh, .fb, .fc, .fd, .fe, .ff, .fg, .fh, .gb, .gc, .gd, .ge, .gf, .gg, .gh, .hb, .hc, .hd, .he, .hf, .hg, .hh)) _
                    - .ab * Matrix7_det(Mat7(.ba, .bc, .bd, .be, .bf, .bg, .bh, .ca, .cc, .cd, .ce, .cf, .cg, .ch, .da, .dc, .dd, .de, .df, .dg, .dh, .ea, .ec, .ed, .ee, .ef, .eg, .eh, .fa, .fc, .fd, .fe, .ff, .fg, .fh, .ga, .gc, .gd, .ge, .gf, .gg, .gh, .ha, .hc, .hd, .he, .hf, .hg, .hh)) _
                    + .ac * Matrix7_det(Mat7(.ba, .bb, .bd, .be, .bf, .bg, .bh, .ca, .cb, .cd, .ce, .cf, .cg, .ch, .da, .db, .dd, .de, .df, .dg, .dh, .ea, .eb, .ed, .ee, .ef, .eg, .eh, .fa, .fb, .fd, .fe, .ff, .fg, .fh, .ga, .gb, .gd, .ge, .gf, .gg, .gh, .ha, .hb, .hd, .he, .hf, .hg, .hh)) _
                    - .ad * Matrix7_det(Mat7(.ba, .bb, .bc, .be, .bf, .bg, .bh, .ca, .cb, .cc, .ce, .cf, .cg, .ch, .da, .db, .dc, .de, .df, .dg, .dh, .ea, .eb, .ec, .ee, .ef, .eg, .eh, .fa, .fb, .fc, .fe, .ff, .fg, .fh, .ga, .gb, .gc, .ge, .gf, .gg, .gh, .ha, .hb, .hc, .he, .hf, .hg, .hh)) _
                    + .ae * Matrix7_det(Mat7(.ba, .bb, .bc, .bd, .bf, .bg, .bh, .ca, .cb, .cc, .cd, .cf, .cg, .ch, .da, .db, .dc, .dd, .df, .dg, .dh, .ea, .eb, .ec, .ed, .ef, .eg, .eh, .fa, .fb, .fc, .fd, .ff, .fg, .fh, .ga, .gb, .gc, .gd, .gf, .gg, .gh, .ha, .hb, .hc, .hd, .hf, .hg, .hh)) _
                    - .af * Matrix7_det(Mat7(.ba, .bb, .bc, .bd, .be, .bg, .bh, .ca, .cb, .cc, .cd, .ce, .cg, .ch, .da, .db, .dc, .dd, .de, .dg, .dh, .ea, .eb, .ec, .ed, .ee, .eg, .eh, .fa, .fb, .fc, .fd, .fe, .fg, .fh, .ga, .gb, .gc, .gd, .ge, .gg, .gh, .ha, .hb, .hc, .hd, .he, .hg, .hh)) _
                    + .ag * Matrix7_det(Mat7(.ba, .bb, .bc, .bd, .be, .bf, .bh, .ca, .cb, .cc, .cd, .ce, .cf, .ch, .da, .db, .dc, .dd, .de, .df, .dh, .ea, .eb, .ec, .ed, .ee, .ef, .eh, .fa, .fb, .fc, .fd, .fe, .ff, .fh, .ga, .gb, .gc, .gd, .ge, .gf, .gh, .ha, .hb, .hc, .hd, .he, .hf, .hh)) _
                    - .ah * Matrix7_det(Mat7(.ba, .bb, .bc, .bd, .be, .bf, .bg, .ca, .cb, .cc, .cd, .ce, .cf, .cg, .da, .db, .dc, .dd, .de, .df, .dg, .ea, .eb, .ec, .ed, .ee, .ef, .eg, .fa, .fb, .fc, .fd, .fe, .ff, .fg, .ga, .gb, .gc, .gd, .ge, .gf, .gg, .ha, .hb, .hc, .hd, .he, .hf, .hg))
    End With
End Function
Public Function Matrix9_det(m As Matrix9) As Double
    'Berechnet die Determinante einer 9x9-Matrix
    With m
        Matrix9_det = .aa * Matrix8_det(Mat8(Vec8(.bb, .bc, .bd, .be, .bf, .bg, .bh, .bi), Vec8(.cb, .cc, .cd, .ce, .cf, .cg, .ch, .ci), Vec8(.db, .dc, .dd, .de, .df, .dg, .dh, .di), Vec8(.eb, .ec, .ed, .ee, .ef, .eg, .eh, .ei), Vec8(.fb, .fc, .fd, .fe, .ff, .fg, .fh, .fi), Vec8(.gb, .gc, .gd, .ge, .gf, .gg, .gh, .gi), Vec8(.hb, .hc, .hd, .he, .hf, .hg, .hh, .hi), Vec8(.ib, .ic, .id, .ie, .if, .ig, .ih, .ii))) _
                    - .ab * Matrix8_det(Mat8(Vec8(.ba, .bc, .bd, .be, .bf, .bg, .bh, .bi), Vec8(.ca, .cc, .cd, .ce, .cf, .cg, .ch, .ci), Vec8(.da, .dc, .dd, .de, .df, .dg, .dh, .di), Vec8(.ea, .ec, .ed, .ee, .ef, .eg, .eh, .ei), Vec8(.fa, .fc, .fd, .fe, .ff, .fg, .fh, .fi), Vec8(.ga, .gc, .gd, .ge, .gf, .gg, .gh, .gi), Vec8(.ha, .hc, .hd, .he, .hf, .hg, .hh, .hi), Vec8(.ia, .ic, .id, .ie, .if, .ig, .ih, .ii))) _
                    + .ac * Matrix8_det(Mat8(Vec8(.ba, .bb, .bd, .be, .bf, .bg, .bh, .bi), Vec8(.ca, .cb, .cd, .ce, .cf, .cg, .ch, .ci), Vec8(.da, .db, .dd, .de, .df, .dg, .dh, .di), Vec8(.ea, .eb, .ed, .ee, .ef, .eg, .eh, .ei), Vec8(.fa, .fb, .fd, .fe, .ff, .fg, .fh, .fi), Vec8(.ga, .gb, .gd, .ge, .gf, .gg, .gh, .gi), Vec8(.ha, .hb, .hd, .he, .hf, .hg, .hh, .hi), Vec8(.ia, .ib, .id, .ie, .if, .ig, .ih, .ii))) _
                    - .ad * Matrix8_det(Mat8(Vec8(.ba, .bb, .bc, .be, .bf, .bg, .bh, .bi), Vec8(.ca, .cb, .cc, .ce, .cf, .cg, .ch, .ci), Vec8(.da, .db, .dc, .de, .df, .dg, .dh, .di), Vec8(.ea, .eb, .ec, .ee, .ef, .eg, .eh, .ei), Vec8(.fa, .fb, .fc, .fe, .ff, .fg, .fh, .fi), Vec8(.ga, .gb, .gc, .ge, .gf, .gg, .gh, .gi), Vec8(.ha, .hb, .hc, .he, .hf, .hg, .hh, .hi), Vec8(.ia, .ib, .ic, .ie, .if, .ig, .ih, .ii))) _
                    + .ae * Matrix8_det(Mat8(Vec8(.ba, .bb, .bc, .bd, .bf, .bg, .bh, .bi), Vec8(.ca, .cb, .cc, .cd, .cf, .cg, .ch, .ci), Vec8(.da, .db, .dc, .dd, .df, .dg, .dh, .di), Vec8(.ea, .eb, .ec, .ed, .ef, .eg, .eh, .ei), Vec8(.fa, .fb, .fc, .fd, .ff, .fg, .fh, .fi), Vec8(.ga, .gb, .gc, .gd, .gf, .gg, .gh, .gi), Vec8(.ha, .hb, .hc, .hd, .hf, .hg, .hh, .hi), Vec8(.ia, .ib, .ic, .id, .if, .ig, .ih, .ii))) _
                    - .af * Matrix8_det(Mat8(Vec8(.ba, .bb, .bc, .bd, .be, .bg, .bh, .bi), Vec8(.ca, .cb, .cc, .cd, .ce, .cg, .ch, .ci), Vec8(.da, .db, .dc, .dd, .de, .dg, .dh, .di), Vec8(.ea, .eb, .ec, .ed, .ee, .eg, .eh, .ei), Vec8(.fa, .fb, .fc, .fd, .fe, .fg, .fh, .fi), Vec8(.ga, .gb, .gc, .gd, .ge, .gg, .gh, .gi), Vec8(.ha, .hb, .hc, .hd, .he, .hg, .hh, .hi), Vec8(.ia, .ib, .ic, .id, .ie, .ig, .ih, .ii))) _
                    + .ag * Matrix8_det(Mat8(Vec8(.ba, .bb, .bc, .bd, .be, .bf, .bh, .bi), Vec8(.ca, .cb, .cc, .cd, .ce, .cf, .ch, .ci), Vec8(.da, .db, .dc, .dd, .de, .df, .dh, .di), Vec8(.ea, .eb, .ec, .ed, .ee, .ef, .eh, .ei), Vec8(.fa, .fb, .fc, .fd, .fe, .ff, .fh, .fi), Vec8(.ga, .gb, .gc, .gd, .ge, .gf, .gh, .gi), Vec8(.ha, .hb, .hc, .hd, .he, .hf, .hh, .hi), Vec8(.ia, .ib, .ic, .id, .ie, .if, .ih, .ii))) _
                    - .ah * Matrix8_det(Mat8(Vec8(.ba, .bb, .bc, .bd, .be, .bf, .bg, .bi), Vec8(.ca, .cb, .cc, .cd, .ce, .cf, .cg, .ci), Vec8(.da, .db, .dc, .dd, .de, .df, .dg, .di), Vec8(.ea, .eb, .ec, .ed, .ee, .ef, .eg, .ei), Vec8(.fa, .fb, .fc, .fd, .fe, .ff, .fg, .fi), Vec8(.ga, .gb, .gc, .gd, .ge, .gf, .gg, .gi), Vec8(.ha, .hb, .hc, .hd, .he, .hf, .hg, .hi), Vec8(.ia, .ib, .ic, .id, .ie, .if, .ig, .ii))) _
                    + .ai * Matrix8_det(Mat8(Vec8(.ba, .bb, .bc, .bd, .be, .bf, .bg, .bh), Vec8(.ca, .cb, .cc, .cd, .ce, .cf, .cg, .ch), Vec8(.da, .db, .dc, .dd, .de, .df, .dg, .dh), Vec8(.ea, .eb, .ec, .ed, .ee, .ef, .eg, .eh), Vec8(.fa, .fb, .fc, .fd, .fe, .ff, .fg, .fh), Vec8(.ga, .gb, .gc, .gd, .ge, .gf, .gg, .gh), Vec8(.ha, .hb, .hc, .hd, .he, .hf, .hg, .hh), Vec8(.ia, .ib, .ic, .id, .ie, .if, .ig, .ih)))
    End With
End Function
Public Function Matrix10_det(m As Matrix10) As Double
    'Berechnet die Determinante einer 9x9-Matrix
    Dim m9 As Matrix9
    'Dim det_a As Double, det_b As Double, det_c As Double, det_d As Double, det_e As Double, det_f As Double, det_g As Double, det_h As Double, det_i As Double
    With m9
        .aa = m.bb: .ab = m.bc: .ac = m.bd: .ad = m.be: .ae = m.bf: .af = m.bg: .ag = m.bh: .ah = m.bi: .ai = m.bj
        .ba = m.cb: .bb = m.cc: .bc = m.cd: .bd = m.ce: .be = m.cf: .bf = m.cg: .bg = m.ch: .bh = m.ci: .ai = m.cj
        .ca = m.db: .cb = m.dc: .cc = m.dd: .cd = m.de: .ce = m.df: .cf = m.dg: .cg = m.dh: .ch = m.di: .ai = m.dj
        .da = m.eb: .db = m.ec: .dc = m.ed: .dd = m.ee: .de = m.ef: .df = m.eg: .dg = m.eh: .dh = m.ei: .ai = m.ej
        .ea = m.fb: .eb = m.fc: .ec = m.fd: .ed = m.fe: .ee = m.ff: .ef = m.fg: .eg = m.fh: .eh = m.fi: .ai = m.fj
        .fa = m.gb: .fb = m.gc: .fc = m.gd: .fd = m.ge: .fe = m.gf: .ff = m.gg: .fg = m.gh: .fh = m.gi: .ai = m.gj
        .ga = m.hb: .gb = m.hc: .gc = m.hd: .gd = m.he: .ge = m.hf: .gf = m.hg: .gg = m.hh: .gh = m.hi: .ai = m.hj
        .ha = m.ib: .hb = m.ic: .hc = m.id: .hd = m.ie: .he = m.if: .hf = m.ig: .hg = m.ih: .hh = m.ii: .ai = m.ij
        .ia = m.jb: .ib = m.jc: .ic = m.jd: .id = m.je: .ie = m.jf: .if = m.jg: .ig = m.jh: .ih = m.ji: .ai = m.jj
    End With
    Dim det_a As Double: det_a = Matrix9_det(m9)
    With m9
        .aa = m.ba: .ab = m.bc: .ac = m.bd: .ad = m.be: .ae = m.bf: .af = m.bg: .ag = m.bh: .ah = m.bi: .ai = m.bj
        .ba = m.ca: .bb = m.cc: .bc = m.cd: .bd = m.ce: .be = m.cf: .bf = m.cg: .bg = m.ch: .bh = m.ci: .ai = m.cj
        .ca = m.da: .cb = m.dc: .cc = m.dd: .cd = m.de: .ce = m.df: .cf = m.dg: .cg = m.dh: .ch = m.di: .ai = m.dj
        .da = m.ea: .db = m.ec: .dc = m.ed: .dd = m.ee: .de = m.ef: .df = m.eg: .dg = m.eh: .dh = m.ei: .ai = m.ej
        .ea = m.fa: .eb = m.fc: .ec = m.fd: .ed = m.fe: .ee = m.ff: .ef = m.fg: .eg = m.fh: .eh = m.fi: .ai = m.fj
        .fa = m.ga: .fb = m.gc: .fc = m.gd: .fd = m.ge: .fe = m.gf: .ff = m.gg: .fg = m.gh: .fh = m.gi: .ai = m.gj
        .ga = m.ha: .gb = m.hc: .gc = m.hd: .gd = m.he: .ge = m.hf: .gf = m.hg: .gg = m.hh: .gh = m.hi: .ai = m.hj
        .ha = m.ia: .hb = m.ic: .hc = m.id: .hd = m.ie: .he = m.if: .hf = m.ig: .hg = m.ih: .hh = m.ii: .ai = m.ij
        .ia = m.ja: .ib = m.jc: .ic = m.jd: .id = m.je: .ie = m.jf: .if = m.jg: .ig = m.jh: .ih = m.ji: .ai = m.jj
    End With
    Dim det_b As Double: det_b = Matrix9_det(m9)
    With m9
        .aa = m.ba: .ab = m.bb: .ac = m.bd: .ad = m.be: .ae = m.bf: .af = m.bg: .ag = m.bh: .ah = m.bi: .ai = m.bj
        .ba = m.ca: .bb = m.cb: .bc = m.cd: .bd = m.ce: .be = m.cf: .bf = m.cg: .bg = m.ch: .bh = m.ci: .ai = m.cj
        .ca = m.da: .cb = m.db: .cc = m.dd: .cd = m.de: .ce = m.df: .cf = m.dg: .cg = m.dh: .ch = m.di: .ai = m.dj
        .da = m.ea: .db = m.eb: .dc = m.ed: .dd = m.ee: .de = m.ef: .df = m.eg: .dg = m.eh: .dh = m.ei: .ai = m.ej
        .ea = m.fa: .eb = m.fb: .ec = m.fd: .ed = m.fe: .ee = m.ff: .ef = m.fg: .eg = m.fh: .eh = m.fi: .ai = m.fj
        .fa = m.ga: .fb = m.gb: .fc = m.gd: .fd = m.ge: .fe = m.gf: .ff = m.gg: .fg = m.gh: .fh = m.gi: .ai = m.gj
        .ga = m.ha: .gb = m.hb: .gc = m.hd: .gd = m.he: .ge = m.hf: .gf = m.hg: .gg = m.hh: .gh = m.hi: .ai = m.hj
        .ha = m.ia: .hb = m.ib: .hc = m.id: .hd = m.ie: .he = m.if: .hf = m.ig: .hg = m.ih: .hh = m.ii: .ai = m.ij
        .ia = m.ja: .ib = m.jb: .ic = m.jd: .id = m.je: .ie = m.jf: .if = m.jg: .ig = m.jh: .ih = m.ji: .ai = m.jj
    End With
    Dim det_c As Double: det_c = Matrix9_det(m9)
    With m9
        .aa = m.ba: .ab = m.bb: .ac = m.bc: .ad = m.be: .ae = m.bf: .af = m.bg: .ag = m.bh: .ah = m.bi: .ai = m.bj
        .ba = m.ca: .bb = m.cb: .bc = m.cc: .bd = m.ce: .be = m.cf: .bf = m.cg: .bg = m.ch: .bh = m.ci: .ai = m.cj
        .ca = m.da: .cb = m.db: .cc = m.dc: .cd = m.de: .ce = m.df: .cf = m.dg: .cg = m.dh: .ch = m.di: .ai = m.dj
        .da = m.ea: .db = m.eb: .dc = m.ec: .dd = m.ee: .de = m.ef: .df = m.eg: .dg = m.eh: .dh = m.ei: .ai = m.ej
        .ea = m.fa: .eb = m.fb: .ec = m.fc: .ed = m.fe: .ee = m.ff: .ef = m.fg: .eg = m.fh: .eh = m.fi: .ai = m.fj
        .fa = m.ga: .fb = m.gb: .fc = m.gc: .fd = m.ge: .fe = m.gf: .ff = m.gg: .fg = m.gh: .fh = m.gi: .ai = m.gj
        .ga = m.ha: .gb = m.hb: .gc = m.hc: .gd = m.he: .ge = m.hf: .gf = m.hg: .gg = m.hh: .gh = m.hi: .ai = m.hj
        .ha = m.ia: .hb = m.ib: .hc = m.ic: .hd = m.ie: .he = m.if: .hf = m.ig: .hg = m.ih: .hh = m.ii: .ai = m.ij
        .ia = m.ja: .ib = m.jb: .ic = m.jc: .id = m.je: .ie = m.jf: .if = m.jg: .ig = m.jh: .ih = m.ji: .ai = m.jj
    End With
    Dim det_d As Double: det_d = Matrix9_det(m9)
    With m9
        .aa = m.ba: .ab = m.bb: .ac = m.bc: .ad = m.bd: .ae = m.bf: .af = m.bg: .ag = m.bh: .ah = m.bi: .ai = m.bj
        .ba = m.ca: .bb = m.cb: .bc = m.cc: .bd = m.cd: .be = m.cf: .bf = m.cg: .bg = m.ch: .bh = m.ci: .ai = m.cj
        .ca = m.da: .cb = m.db: .cc = m.dc: .cd = m.dd: .ce = m.df: .cf = m.dg: .cg = m.dh: .ch = m.di: .ai = m.dj
        .da = m.ea: .db = m.eb: .dc = m.ec: .dd = m.ed: .de = m.ef: .df = m.eg: .dg = m.eh: .dh = m.ei: .ai = m.ej
        .ea = m.fa: .eb = m.fb: .ec = m.fc: .ed = m.fd: .ee = m.ff: .ef = m.fg: .eg = m.fh: .eh = m.fi: .ai = m.fj
        .fa = m.ga: .fb = m.gb: .fc = m.gc: .fd = m.gd: .fe = m.gf: .ff = m.gg: .fg = m.gh: .fh = m.gi: .ai = m.gj
        .ga = m.ha: .gb = m.hb: .gc = m.hc: .gd = m.hd: .ge = m.hf: .gf = m.hg: .gg = m.hh: .gh = m.hi: .ai = m.hj
        .ha = m.ia: .hb = m.ib: .hc = m.ic: .hd = m.id: .he = m.if: .hf = m.ig: .hg = m.ih: .hh = m.ii: .ai = m.ij
        .ia = m.ja: .ib = m.jb: .ic = m.jc: .id = m.jd: .ie = m.jf: .if = m.jg: .ig = m.jh: .ih = m.ji: .ai = m.jj
    End With
    Dim det_e As Double: det_e = Matrix9_det(m9)
    With m9
        .aa = m.ba: .ab = m.bb: .ac = m.bc: .ad = m.bd: .ae = m.be: .af = m.bg: .ag = m.bh: .ah = m.bi: .ai = m.bj
        .ba = m.ca: .bb = m.cb: .bc = m.cc: .bd = m.cd: .be = m.ce: .bf = m.cg: .bg = m.ch: .bh = m.ci: .ai = m.cj
        .ca = m.da: .cb = m.db: .cc = m.dc: .cd = m.dd: .ce = m.de: .cf = m.dg: .cg = m.dh: .ch = m.di: .ai = m.dj
        .da = m.ea: .db = m.eb: .dc = m.ec: .dd = m.ed: .de = m.ee: .df = m.eg: .dg = m.eh: .dh = m.ei: .ai = m.ej
        .ea = m.fa: .eb = m.fb: .ec = m.fc: .ed = m.fd: .ee = m.fe: .ef = m.fg: .eg = m.fh: .eh = m.fi: .ai = m.fj
        .fa = m.ga: .fb = m.gb: .fc = m.gc: .fd = m.gd: .fe = m.ge: .ff = m.gg: .fg = m.gh: .fh = m.gi: .ai = m.gj
        .ga = m.ha: .gb = m.hb: .gc = m.hc: .gd = m.hd: .ge = m.he: .gf = m.hg: .gg = m.hh: .gh = m.hi: .ai = m.hj
        .ha = m.ia: .hb = m.ib: .hc = m.ic: .hd = m.id: .he = m.ie: .hf = m.ig: .hg = m.ih: .hh = m.ii: .ai = m.ij
        .ia = m.ja: .ib = m.jb: .ic = m.jc: .id = m.jd: .ie = m.je: .if = m.jg: .ig = m.jh: .ih = m.ji: .ai = m.jj
    End With
    Dim det_f As Double: det_f = Matrix9_det(m9)
    With m9
        .aa = m.ba: .ab = m.bb: .ac = m.bc: .ad = m.bd: .ae = m.be: .af = m.bf: .ag = m.bh: .ah = m.bi: .ai = m.bj
        .ba = m.ca: .bb = m.cb: .bc = m.cc: .bd = m.cd: .be = m.ce: .bf = m.cf: .bg = m.ch: .bh = m.ci: .ai = m.cj
        .ca = m.da: .cb = m.db: .cc = m.dc: .cd = m.dd: .ce = m.de: .cf = m.df: .cg = m.dh: .ch = m.di: .ai = m.dj
        .da = m.ea: .db = m.eb: .dc = m.ec: .dd = m.ed: .de = m.ee: .df = m.ef: .dg = m.eh: .dh = m.ei: .ai = m.ej
        .ea = m.fa: .eb = m.fb: .ec = m.fc: .ed = m.fd: .ee = m.fe: .ef = m.ff: .eg = m.fh: .eh = m.fi: .ai = m.fj
        .fa = m.ga: .fb = m.gb: .fc = m.gc: .fd = m.gd: .fe = m.ge: .ff = m.gf: .fg = m.gh: .fh = m.gi: .ai = m.gj
        .ga = m.ha: .gb = m.hb: .gc = m.hc: .gd = m.hd: .ge = m.he: .gf = m.hf: .gg = m.hh: .gh = m.hi: .ai = m.hj
        .ha = m.ia: .hb = m.ib: .hc = m.ic: .hd = m.id: .he = m.ie: .hf = m.if: .hg = m.ih: .hh = m.ii: .ai = m.ij
        .ia = m.ja: .ib = m.jb: .ic = m.jc: .id = m.jd: .ie = m.je: .if = m.jf: .ig = m.jh: .ih = m.ji: .ai = m.jj
    End With
    Dim det_g As Double: det_g = Matrix9_det(m9)
    With m9
        .aa = m.ba: .ab = m.bb: .ac = m.bc: .ad = m.bd: .ae = m.be: .af = m.bf: .ag = m.bg: .ah = m.bi: .ai = m.bj
        .ba = m.ca: .bb = m.cb: .bc = m.cc: .bd = m.cd: .be = m.ce: .bf = m.cf: .bg = m.cg: .bh = m.ci: .ai = m.cj
        .ca = m.da: .cb = m.db: .cc = m.dc: .cd = m.dd: .ce = m.de: .cf = m.df: .cg = m.dg: .ch = m.di: .ai = m.dj
        .da = m.ea: .db = m.eb: .dc = m.ec: .dd = m.ed: .de = m.ee: .df = m.ef: .dg = m.eg: .dh = m.ei: .ai = m.ej
        .ea = m.fa: .eb = m.fb: .ec = m.fc: .ed = m.fd: .ee = m.fe: .ef = m.ff: .eg = m.fg: .eh = m.fi: .ai = m.fj
        .fa = m.ga: .fb = m.gb: .fc = m.gc: .fd = m.gd: .fe = m.ge: .ff = m.gf: .fg = m.gg: .fh = m.gi: .ai = m.gj
        .ga = m.ha: .gb = m.hb: .gc = m.hc: .gd = m.hd: .ge = m.he: .gf = m.hf: .gg = m.hg: .gh = m.hi: .ai = m.hj
        .ha = m.ia: .hb = m.ib: .hc = m.ic: .hd = m.id: .he = m.ie: .hf = m.if: .hg = m.ig: .hh = m.ii: .ai = m.ij
        .ia = m.ja: .ib = m.jb: .ic = m.jc: .id = m.jd: .ie = m.je: .if = m.jf: .ig = m.jg: .ih = m.ji: .ai = m.jj
    End With
    Dim det_h As Double: det_h = Matrix9_det(m9)
    With m9
        .aa = m.ba: .ab = m.bb: .ac = m.bc: .ad = m.bd: .ae = m.be: .af = m.bf: .ag = m.bg: .ah = m.bh: .ai = m.bj
        .ba = m.ca: .bb = m.cb: .bc = m.cc: .bd = m.cd: .be = m.ce: .bf = m.cf: .bg = m.cg: .bh = m.ch: .ai = m.cj
        .ca = m.da: .cb = m.db: .cc = m.dc: .cd = m.dd: .ce = m.de: .cf = m.df: .cg = m.dg: .ch = m.dh: .ai = m.dj
        .da = m.ea: .db = m.eb: .dc = m.ec: .dd = m.ed: .de = m.ee: .df = m.ef: .dg = m.eg: .dh = m.eh: .ai = m.ej
        .ea = m.fa: .eb = m.fb: .ec = m.fc: .ed = m.fd: .ee = m.fe: .ef = m.ff: .eg = m.fg: .eh = m.fh: .ai = m.fj
        .fa = m.ga: .fb = m.gb: .fc = m.gc: .fd = m.gd: .fe = m.ge: .ff = m.gf: .fg = m.gg: .fh = m.gh: .ai = m.gj
        .ga = m.ha: .gb = m.hb: .gc = m.hc: .gd = m.hd: .ge = m.he: .gf = m.hf: .gg = m.hg: .gh = m.hh: .ai = m.hj
        .ha = m.ia: .hb = m.ib: .hc = m.ic: .hd = m.id: .he = m.ie: .hf = m.if: .hg = m.ig: .hh = m.ih: .ai = m.ij
        .ia = m.ja: .ib = m.jb: .ic = m.jc: .id = m.jd: .ie = m.je: .if = m.jf: .ig = m.jg: .ih = m.jh: .ai = m.jj
    End With
    Dim det_i As Double: det_i = Matrix9_det(m9)
    With m9
        .aa = m.ba: .ab = m.bb: .ac = m.bc: .ad = m.bd: .ae = m.be: .af = m.bf: .ag = m.bg: .ah = m.bh: .ai = m.bi
        .ba = m.ca: .bb = m.cb: .bc = m.cc: .bd = m.cd: .be = m.ce: .bf = m.cf: .bg = m.cg: .bh = m.ch: .ai = m.ci
        .ca = m.da: .cb = m.db: .cc = m.dc: .cd = m.dd: .ce = m.de: .cf = m.df: .cg = m.dg: .ch = m.dh: .ai = m.di
        .da = m.ea: .db = m.eb: .dc = m.ec: .dd = m.ed: .de = m.ee: .df = m.ef: .dg = m.eg: .dh = m.eh: .ai = m.ei
        .ea = m.fa: .eb = m.fb: .ec = m.fc: .ed = m.fd: .ee = m.fe: .ef = m.ff: .eg = m.fg: .eh = m.fh: .ai = m.fi
        .fa = m.ga: .fb = m.gb: .fc = m.gc: .fd = m.gd: .fe = m.ge: .ff = m.gf: .fg = m.gg: .fh = m.gh: .ai = m.gi
        .ga = m.ha: .gb = m.hb: .gc = m.hc: .gd = m.hd: .ge = m.he: .gf = m.hf: .gg = m.hg: .gh = m.hh: .ai = m.hi
        .ha = m.ia: .hb = m.ib: .hc = m.ic: .hd = m.id: .he = m.ie: .hf = m.if: .hg = m.ig: .hh = m.ih: .ai = m.ii
        .ia = m.ja: .ib = m.jb: .ic = m.jc: .id = m.jd: .ie = m.je: .if = m.jf: .ig = m.jg: .ih = m.jh: .ai = m.ji
    End With
    Dim det_j As Double: det_j = Matrix9_det(m9)
        
    With m
        Matrix10_det = .aa * det_a - .ab * det_b + .ac * det_c - .ad * det_d + .ae * det_e - .af * det_f + .ag * det_g - .ah * det_h + .ai * det_i - .aj * det_j
    End With
'    With m
'        Matrix10_det = .aa * Matrix9_det(Mat9(Vec9(.bb, .bc, .bd, .be, .bf, .bg, .bh, .bi, .bj), Vec9(.cb, .cc, .cd, .ce, .cf, .cg, .ch, .ci, .cj), Vec9(.db, .dc, .dd, .de, .df, .dg, .dh, .di, .dj), Vec9(.eb, .ec, .ed, .ee, .ef, .eg, .eh, .ei, .ej), Vec9(.fb, .fc, .fd, .fe, .ff, .fg, .fh, .fi, .fj), Vec9(.gb, .gc, .gd, .ge, .gf, .gg, .gh, .gi, .gj), Vec9(.hb, .hc, .hd, .he, .hf, .hg, .hh, .hi, .hj), Vec9(.ib, .ic, .id, .ie, .if, .ig, .ih, .ii, .ij), Vec9(.jb, .jc, .jd, .je, .jf, .jg, .jh, .ji, .jj))) _
'                     - .ab * Matrix9_det(Mat9(Vec9(.ba, .bc, .bd, .be, .bf, .bg, .bh, .bi, .bj), Vec9(.ca, .cc, .cd, .ce, .cf, .cg, .ch, .ci, .cj), Vec9(.da, .dc, .dd, .de, .df, .dg, .dh, .di, .dj), Vec9(.ea, .ec, .ed, .ee, .ef, .eg, .eh, .ei, .ej), Vec9(.fa, .fc, .fd, .fe, .ff, .fg, .fh, .fi, .fj), Vec9(.ga, .gc, .gd, .ge, .gf, .gg, .gh, .gi, .gj), Vec9(.ha, .hc, .hd, .he, .hf, .hg, .hh, .hi, .hj), Vec9(.ia, .ic, .id, .ie, .if, .ig, .ih, .ii, .ij), Vec9(.ja, .jc, .jd, .je, .jf, .jg, .jh, .ji, .jj))) _
'                     + .ac * Matrix9_det(Mat9(Vec9(.ba, .bb, .bd, .be, .bf, .bg, .bh, .bi, .bj), Vec9(.ca, .cb, .cd, .ce, .cf, .cg, .ch, .ci, .cj), Vec9(.da, .db, .dd, .de, .df, .dg, .dh, .di, .dj), Vec9(.ea, .eb, .ed, .ee, .ef, .eg, .eh, .ei, .ej), Vec9(.fa, .fb, .fd, .fe, .ff, .fg, .fh, .fi, .fj), Vec9(.ga, .gb, .gd, .ge, .gf, .gg, .gh, .gi, .gj), Vec9(.ha, .hb, .hd, .he, .hf, .hg, .hh, .hi, .hj), Vec9(.ia, .ib, .id, .ie, .if, .ig, .ih, .ii, .ij), Vec9(.ja, .jb, .jd, .je, .jf, .jg, .jh, .ji, .jj))) _
'                     - .ad * Matrix9_det(Mat9(Vec9(.ba, .bb, .bc, .be, .bf, .bg, .bh, .bi, .bj), Vec9(.ca, .cb, .cc, .ce, .cf, .cg, .ch, .ci, .cj), Vec9(.da, .db, .dc, .de, .df, .dg, .dh, .di, .dj), Vec9(.ea, .eb, .ec, .ee, .ef, .eg, .eh, .ei, .ej), Vec9(.fa, .fb, .fc, .fe, .ff, .fg, .fh, .fi, .fj), Vec9(.ga, .gb, .gc, .ge, .gf, .gg, .gh, .gi, .gj), Vec9(.ha, .hb, .hc, .he, .hf, .hg, .hh, .hi, .hj), Vec9(.ia, .ib, .ic, .ie, .if, .ig, .ih, .ii, .ij), Vec9(.ja, .jb, .jc, .je, .jf, .jg, .jh, .ji, .jj))) _
'                     + .ae * Matrix9_det(Mat9(Vec9(.ba, .bb, .bc, .bd, .bf, .bg, .bh, .bi, .bj), Vec9(.ca, .cb, .cc, .cd, .cf, .cg, .ch, .ci, .cj), Vec9(.da, .db, .dc, .dd, .df, .dg, .dh, .di, .dj), Vec9(.ea, .eb, .ec, .ed, .ef, .eg, .eh, .ei, .ej), Vec9(.fa, .fb, .fc, .fd, .ff, .fg, .fh, .fi, .fj), Vec9(.ga, .gb, .gc, .gd, .gf, .gg, .gh, .gi, .gj), Vec9(.ha, .hb, .hc, .hd, .hf, .hg, .hh, .hi, .hj), Vec9(.ia, .ib, .ic, .id, .if, .ig, .ih, .ii, .ij), Vec9(.ja, .jb, .jc, .jd, .jf, .jg, .jh, .ji, .jj))) _
'                     - .af * Matrix9_det(Mat9(Vec9(.ba, .bb, .bc, .bd, .be, .bg, .bh, .bi, .bj), Vec9(.ca, .cb, .cc, .cd, .ce, .cg, .ch, .ci, .cj), Vec9(.da, .db, .dc, .dd, .de, .dg, .dh, .di, .dj), Vec9(.ea, .eb, .ec, .ed, .ee, .eg, .eh, .ei, .ej), Vec9(.fa, .fb, .fc, .fd, .fe, .fg, .fh, .fi, .fj), Vec9(.ga, .gb, .gc, .gd, .ge, .gg, .gh, .gi, .gj), Vec9(.ha, .hb, .hc, .hd, .he, .hg, .hh, .hi, .hj), Vec9(.ia, .ib, .ic, .id, .ie, .ig, .ih, .ii, .ij), Vec9(.ja, .jb, .jc, .jd, .je, .jg, .jh, .ji, .jj))) _
'                     + .ag * Matrix9_det(Mat9(Vec9(.ba, .bb, .bc, .bd, .be, .bf, .bh, .bi, .bj), Vec9(.ca, .cb, .cc, .cd, .ce, .cf, .ch, .ci, .cj), Vec9(.da, .db, .dc, .dd, .de, .df, .dh, .di, .dj), Vec9(.ea, .eb, .ec, .ed, .ee, .ef, .eh, .ei, .ej), Vec9(.fa, .fb, .fc, .fd, .fe, .ff, .fh, .fi, .fj), Vec9(.ga, .gb, .gc, .gd, .ge, .gf, .gh, .gi, .gj), Vec9(.ha, .hb, .hc, .hd, .he, .hf, .hh, .hi, .hj), Vec9(.ia, .ib, .ic, .id, .ie, .if, .ih, .ii, .ij), Vec9(.ja, .jb, .jc, .jd, .je, .jf, .jh, .ji, .jj))) _
'                     - .ah * Matrix9_det(Mat9(Vec9(.ba, .bb, .bc, .bd, .be, .bf, .bg, .bi, .bj), Vec9(.ca, .cb, .cc, .cd, .ce, .cf, .cg, .ci, .cj), Vec9(.da, .db, .dc, .dd, .de, .df, .dg, .di, .dj), Vec9(.ea, .eb, .ec, .ed, .ee, .ef, .eg, .ei, .ej), Vec9(.fa, .fb, .fc, .fd, .fe, .ff, .fg, .fi, .fj), Vec9(.ga, .gb, .gc, .gd, .ge, .gf, .gg, .gi, .gj), Vec9(.ha, .hb, .hc, .hd, .he, .hf, .hg, .hi, .hj), Vec9(.ia, .ib, .ic, .id, .ie, .if, .ig, .ii, .ij), Vec9(.ja, .jb, .jc, .jd, .je, .jf, .jg, .ji, .jj))) _
'                     + .ai * Matrix9_det(Mat9(Vec9(.ba, .bb, .bc, .bd, .be, .bf, .bg, .bh, .bj), Vec9(.ca, .cb, .cc, .cd, .ce, .cf, .cg, .ch, .cj), Vec9(.da, .db, .dc, .dd, .de, .df, .dg, .dh, .dj), Vec9(.ea, .eb, .ec, .ed, .ee, .ef, .eg, .eh, .ej), Vec9(.fa, .fb, .fc, .fd, .fe, .ff, .fg, .fh, .fj), Vec9(.ga, .gb, .gc, .gd, .ge, .gf, .gg, .gh, .gj), Vec9(.ha, .hb, .hc, .hd, .he, .hf, .hg, .hh, .hj), Vec9(.ia, .ib, .ic, .id, .ie, .if, .ig, .ih, .ij), Vec9(.ja, .jb, .jc, .jd, .je, .jf, .jg, .jh, .jj))) _
'                     - .aj * Matrix9_det(Mat9(Vec9(.ba, .bb, .bc, .bd, .be, .bf, .bg, .bh, .bi), Vec9(.ca, .cb, .cc, .cd, .ce, .cf, .cg, .ch, .ci), Vec9(.da, .db, .dc, .dd, .de, .df, .dg, .dh, .di), Vec9(.ea, .eb, .ec, .ed, .ee, .ef, .eg, .eh, .ei), Vec9(.fa, .fb, .fc, .fd, .fe, .ff, .fg, .fh, .fi), Vec9(.ga, .gb, .gc, .gd, .ge, .gf, .gg, .gh, .gi), Vec9(.ha, .hb, .hc, .hd, .he, .hf, .hg, .hh, .hi), Vec9(.ia, .ib, .ic, .id, .ie, .if, .ig, .ih, .ii), Vec9(.ja, .jb, .jc, .jd, .je, .jf, .jg, .jh, .ji)))
'    End With
End Function
Public Function Matrix10_detX(m As Matrix10) As Double
    'Berechnet die Determinante einer 9x9-Matrix
    With m
       Matrix10_detX = .aa * Matrix9_det(Mat9(Vec9(.bb, .bc, .bd, .be, .bf, .bg, .bh, .bi, .bj), Vec9(.cb, .cc, .cd, .ce, .cf, .cg, .ch, .ci, .cj), Vec9(.db, .dc, .dd, .de, .df, .dg, .dh, .di, .dj), Vec9(.eb, .ec, .ed, .ee, .ef, .eg, .eh, .ei, .ej), Vec9(.fb, .fc, .fd, .fe, .ff, .fg, .fh, .fi, .fj), Vec9(.gb, .gc, .gd, .ge, .gf, .gg, .gh, .gi, .gj), Vec9(.hb, .hc, .hd, .he, .hf, .hg, .hh, .hi, .hj), Vec9(.ib, .ic, .id, .ie, .if, .ig, .ih, .ii, .ij), Vec9(.jb, .jc, .jd, .je, .jf, .jg, .jh, .ji, .jj))) _
                     - .ab * Matrix9_det(Mat9(Vec9(.ba, .bc, .bd, .be, .bf, .bg, .bh, .bi, .bj), Vec9(.ca, .cc, .cd, .ce, .cf, .cg, .ch, .ci, .cj), Vec9(.da, .dc, .dd, .de, .df, .dg, .dh, .di, .dj), Vec9(.ea, .ec, .ed, .ee, .ef, .eg, .eh, .ei, .ej), Vec9(.fa, .fc, .fd, .fe, .ff, .fg, .fh, .fi, .fj), Vec9(.ga, .gc, .gd, .ge, .gf, .gg, .gh, .gi, .gj), Vec9(.ha, .hc, .hd, .he, .hf, .hg, .hh, .hi, .hj), Vec9(.ia, .ic, .id, .ie, .if, .ig, .ih, .ii, .ij), Vec9(.ja, .jc, .jd, .je, .jf, .jg, .jh, .ji, .jj))) _
                     + .ac * Matrix9_det(Mat9(Vec9(.ba, .bb, .bd, .be, .bf, .bg, .bh, .bi, .bj), Vec9(.ca, .cb, .cd, .ce, .cf, .cg, .ch, .ci, .cj), Vec9(.da, .db, .dd, .de, .df, .dg, .dh, .di, .dj), Vec9(.ea, .eb, .ed, .ee, .ef, .eg, .eh, .ei, .ej), Vec9(.fa, .fb, .fd, .fe, .ff, .fg, .fh, .fi, .fj), Vec9(.ga, .gb, .gd, .ge, .gf, .gg, .gh, .gi, .gj), Vec9(.ha, .hb, .hd, .he, .hf, .hg, .hh, .hi, .hj), Vec9(.ia, .ib, .id, .ie, .if, .ig, .ih, .ii, .ij), Vec9(.ja, .jb, .jd, .je, .jf, .jg, .jh, .ji, .jj))) _
                     - .ad * Matrix9_det(Mat9(Vec9(.ba, .bb, .bc, .be, .bf, .bg, .bh, .bi, .bj), Vec9(.ca, .cb, .cc, .ce, .cf, .cg, .ch, .ci, .cj), Vec9(.da, .db, .dc, .de, .df, .dg, .dh, .di, .dj), Vec9(.ea, .eb, .ec, .ee, .ef, .eg, .eh, .ei, .ej), Vec9(.fa, .fb, .fc, .fe, .ff, .fg, .fh, .fi, .fj), Vec9(.ga, .gb, .gc, .ge, .gf, .gg, .gh, .gi, .gj), Vec9(.ha, .hb, .hc, .he, .hf, .hg, .hh, .hi, .hj), Vec9(.ia, .ib, .ic, .ie, .if, .ig, .ih, .ii, .ij), Vec9(.ja, .jb, .jc, .je, .jf, .jg, .jh, .ji, .jj))) _
                     + .ae * Matrix9_det(Mat9(Vec9(.ba, .bb, .bc, .bd, .bf, .bg, .bh, .bi, .bj), Vec9(.ca, .cb, .cc, .cd, .cf, .cg, .ch, .ci, .cj), Vec9(.da, .db, .dc, .dd, .df, .dg, .dh, .di, .dj), Vec9(.ea, .eb, .ec, .ed, .ef, .eg, .eh, .ei, .ej), Vec9(.fa, .fb, .fc, .fd, .ff, .fg, .fh, .fi, .fj), Vec9(.ga, .gb, .gc, .gd, .gf, .gg, .gh, .gi, .gj), Vec9(.ha, .hb, .hc, .hd, .hf, .hg, .hh, .hi, .hj), Vec9(.ia, .ib, .ic, .id, .if, .ig, .ih, .ii, .ij), Vec9(.ja, .jb, .jc, .jd, .jf, .jg, .jh, .ji, .jj))) _
                     - .af * Matrix9_det(Mat9(Vec9(.ba, .bb, .bc, .bd, .be, .bg, .bh, .bi, .bj), Vec9(.ca, .cb, .cc, .cd, .ce, .cg, .ch, .ci, .cj), Vec9(.da, .db, .dc, .dd, .de, .dg, .dh, .di, .dj), Vec9(.ea, .eb, .ec, .ed, .ee, .eg, .eh, .ei, .ej), Vec9(.fa, .fb, .fc, .fd, .fe, .fg, .fh, .fi, .fj), Vec9(.ga, .gb, .gc, .gd, .ge, .gg, .gh, .gi, .gj), Vec9(.ha, .hb, .hc, .hd, .he, .hg, .hh, .hi, .hj), Vec9(.ia, .ib, .ic, .id, .ie, .ig, .ih, .ii, .ij), Vec9(.ja, .jb, .jc, .jd, .je, .jg, .jh, .ji, .jj))) _
                     + .ag * Matrix9_det(Mat9(Vec9(.ba, .bb, .bc, .bd, .be, .bf, .bh, .bi, .bj), Vec9(.ca, .cb, .cc, .cd, .ce, .cf, .ch, .ci, .cj), Vec9(.da, .db, .dc, .dd, .de, .df, .dh, .di, .dj), Vec9(.ea, .eb, .ec, .ed, .ee, .ef, .eh, .ei, .ej), Vec9(.fa, .fb, .fc, .fd, .fe, .ff, .fh, .fi, .fj), Vec9(.ga, .gb, .gc, .gd, .ge, .gf, .gh, .gi, .gj), Vec9(.ha, .hb, .hc, .hd, .he, .hf, .hh, .hi, .hj), Vec9(.ia, .ib, .ic, .id, .ie, .if, .ih, .ii, .ij), Vec9(.ja, .jb, .jc, .jd, .je, .jf, .jh, .ji, .jj))) _
                     - .ah * Matrix9_det(Mat9(Vec9(.ba, .bb, .bc, .bd, .be, .bf, .bg, .bi, .bj), Vec9(.ca, .cb, .cc, .cd, .ce, .cf, .cg, .ci, .cj), Vec9(.da, .db, .dc, .dd, .de, .df, .dg, .di, .dj), Vec9(.ea, .eb, .ec, .ed, .ee, .ef, .eg, .ei, .ej), Vec9(.fa, .fb, .fc, .fd, .fe, .ff, .fg, .fi, .fj), Vec9(.ga, .gb, .gc, .gd, .ge, .gf, .gg, .gi, .gj), Vec9(.ha, .hb, .hc, .hd, .he, .hf, .hg, .hi, .hj), Vec9(.ia, .ib, .ic, .id, .ie, .if, .ig, .ii, .ij), Vec9(.ja, .jb, .jc, .jd, .je, .jf, .jg, .ji, .jj))) _
                     + .ai * Matrix9_det(Mat9(Vec9(.ba, .bb, .bc, .bd, .be, .bf, .bg, .bh, .bj), Vec9(.ca, .cb, .cc, .cd, .ce, .cf, .cg, .ch, .cj), Vec9(.da, .db, .dc, .dd, .de, .df, .dg, .dh, .dj), Vec9(.ea, .eb, .ec, .ed, .ee, .ef, .eg, .eh, .ej), Vec9(.fa, .fb, .fc, .fd, .fe, .ff, .fg, .fh, .fj), Vec9(.ga, .gb, .gc, .gd, .ge, .gf, .gg, .gh, .gj), Vec9(.ha, .hb, .hc, .hd, .he, .hf, .hg, .hh, .hj), Vec9(.ia, .ib, .ic, .id, .ie, .if, .ig, .ih, .ij), Vec9(.ja, .jb, .jc, .jd, .je, .jf, .jg, .jh, .jj))) _
                     - .aj * Matrix9_det(Mat9(Vec9(.ba, .bb, .bc, .bd, .be, .bf, .bg, .bh, .bi), Vec9(.ca, .cb, .cc, .cd, .ce, .cf, .cg, .ch, .ci), Vec9(.da, .db, .dc, .dd, .de, .df, .dg, .dh, .di), Vec9(.ea, .eb, .ec, .ed, .ee, .ef, .eg, .eh, .ei), Vec9(.fa, .fb, .fc, .fd, .fe, .ff, .fg, .fh, .fi), Vec9(.ga, .gb, .gc, .gd, .ge, .gf, .gg, .gh, .gi), Vec9(.ha, .hb, .hc, .hd, .he, .hf, .hg, .hh, .hi), Vec9(.ia, .ib, .ic, .id, .ie, .if, .ig, .ih, .ii), Vec9(.ja, .jb, .jc, .jd, .je, .jf, .jg, .jh, .ji)))
    End With
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
        Case 5: Mat6_Col = Vec6(.af, .bf, .cf, .df, .ef, .ff)
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

Public Property Get Mat7_Row(m As Matrix7, ByVal index As Long) As Vector7
    With m
        Select Case index
        Case 0: Mat7_Row = Vec7(.aa, .ab, .ac, .ad, .ae, .af, .ag)
        Case 1: Mat7_Row = Vec7(.ba, .bb, .bc, .bd, .be, .bf, .bg)
        Case 2: Mat7_Row = Vec7(.ca, .cb, .cc, .cd, .ce, .cf, .cg)
        Case 3: Mat7_Row = Vec7(.da, .db, .dc, .dd, .de, .df, .dg)
        Case 4: Mat7_Row = Vec7(.ea, .eb, .ec, .ed, .ee, .ef, .eg)
        Case 5: Mat7_Row = Vec7(.fa, .fb, .fc, .fd, .fe, .ff, .fg)
        Case 6: Mat7_Row = Vec7(.ga, .gb, .gc, .gd, .ge, .gf, .gg)
        End Select
    End With
End Property
Public Property Let Mat7_Row(m As Matrix7, ByVal index As Long, v As Vector7)
    With m
        Select Case index
        Case 0: .aa = v.a: .ab = v.b: .ac = v.c: .ad = v.d: .ae = v.e: .af = v.f: .ag = v.g
        Case 1: .ba = v.a: .bb = v.b: .bc = v.c: .bd = v.d: .be = v.e: .bf = v.f: .bg = v.g
        Case 2: .ca = v.a: .cb = v.b: .cc = v.c: .cd = v.d: .ce = v.e: .cf = v.f: .cg = v.g
        Case 3: .da = v.a: .db = v.b: .dc = v.c: .dd = v.d: .de = v.e: .df = v.f: .dg = v.g
        Case 4: .ea = v.a: .eb = v.b: .ec = v.c: .ed = v.d: .ee = v.e: .ef = v.f: .eg = v.g
        Case 5: .fa = v.a: .fb = v.b: .fc = v.c: .fd = v.d: .fe = v.e: .ff = v.f: .fg = v.g
        Case 6: .ga = v.a: .gb = v.b: .gc = v.c: .gd = v.d: .ge = v.e: .gf = v.f: .gg = v.g
        End Select
    End With
End Property
Public Property Get Mat7_Col(m As Matrix7, ByVal index As Long) As Vector7
    With m
        Select Case index
        Case 0: Mat7_Col = Vec7(.aa, .ba, .ca, .da, .ea, .fa, .ga)
        Case 1: Mat7_Col = Vec7(.ab, .bb, .cb, .db, .eb, .fb, .gb)
        Case 2: Mat7_Col = Vec7(.ac, .bc, .cc, .dc, .ec, .fc, .gc)
        Case 3: Mat7_Col = Vec7(.ad, .bd, .cd, .dd, .ed, .fd, .gd)
        Case 4: Mat7_Col = Vec7(.ae, .be, .ce, .de, .ee, .fe, .ge)
        Case 5: Mat7_Col = Vec7(.af, .bf, .cf, .df, .ef, .ff, .gf)
        Case 6: Mat7_Col = Vec7(.ag, .bg, .cg, .dg, .eg, .fg, .gg)
        End Select
    End With
End Property
Public Property Let Mat7_Col(m As Matrix7, ByVal index As Long, v As Vector7)
    With m
        Select Case index
        Case 0: .aa = v.a: .ba = v.b: .ca = v.c: .da = v.d: .ea = v.e: .fa = v.f: .ga = v.g
        Case 1: .ab = v.a: .bb = v.b: .cb = v.c: .db = v.d: .eb = v.e: .fb = v.f: .gb = v.g
        Case 2: .ac = v.a: .bc = v.b: .cc = v.c: .dc = v.d: .ec = v.e: .fc = v.f: .gc = v.g
        Case 3: .ad = v.a: .bd = v.b: .cd = v.c: .dd = v.d: .ed = v.e: .fd = v.f: .gd = v.g
        Case 4: .ae = v.a: .be = v.b: .ce = v.c: .de = v.d: .ee = v.e: .fe = v.f: .ge = v.g
        Case 5: .af = v.a: .bf = v.b: .cf = v.c: .df = v.d: .ef = v.e: .ff = v.f: .gf = v.g
        Case 6: .ag = v.a: .bg = v.b: .cg = v.c: .dg = v.d: .eg = v.e: .fg = v.f: .gg = v.g
        End Select
    End With
End Property

Public Property Get Mat8_Row(m As Matrix8, ByVal index As Long) As Vector8
    With m
        Select Case index
        Case 0: Mat8_Row = Vec8(.aa, .ab, .ac, .ad, .ae, .af, .ag, .ah)
        Case 1: Mat8_Row = Vec8(.ba, .bb, .bc, .bd, .be, .bf, .bg, .bh)
        Case 2: Mat8_Row = Vec8(.ca, .cb, .cc, .cd, .ce, .cf, .cg, .ch)
        Case 3: Mat8_Row = Vec8(.da, .db, .dc, .dd, .de, .df, .dg, .dh)
        Case 4: Mat8_Row = Vec8(.ea, .eb, .ec, .ed, .ee, .ef, .eg, .eh)
        Case 5: Mat8_Row = Vec8(.fa, .fb, .fc, .fd, .fe, .ff, .fg, .fh)
        Case 6: Mat8_Row = Vec8(.ga, .gb, .gc, .gd, .ge, .gf, .gg, .gh)
        Case 7: Mat8_Row = Vec8(.ha, .hb, .hc, .hd, .he, .hf, .hg, .hh)
        End Select
    End With
End Property
Public Property Let Mat8_Row(m As Matrix8, ByVal index As Long, v As Vector8)
    With m
        Select Case index
        Case 0: .aa = v.a: .ab = v.b: .ac = v.c: .ad = v.d: .ae = v.e: .af = v.f: .ag = v.g: .ah = v.h
        Case 1: .ba = v.a: .bb = v.b: .bc = v.c: .bd = v.d: .be = v.e: .bf = v.f: .bg = v.g: .bh = v.h
        Case 2: .ca = v.a: .cb = v.b: .cc = v.c: .cd = v.d: .ce = v.e: .cf = v.f: .cg = v.g: .ch = v.h
        Case 3: .da = v.a: .db = v.b: .dc = v.c: .dd = v.d: .de = v.e: .df = v.f: .dg = v.g: .dh = v.h
        Case 4: .ea = v.a: .eb = v.b: .ec = v.c: .ed = v.d: .ee = v.e: .ef = v.f: .eg = v.g: .eh = v.h
        Case 5: .fa = v.a: .fb = v.b: .fc = v.c: .fd = v.d: .fe = v.e: .ff = v.f: .fg = v.g: .fh = v.h
        Case 6: .ga = v.a: .gb = v.b: .gc = v.c: .gd = v.d: .ge = v.e: .gf = v.f: .gg = v.g: .gh = v.h
        Case 7: .ha = v.a: .hb = v.b: .hc = v.c: .hd = v.d: .he = v.e: .hf = v.f: .hg = v.g: .hh = v.h
        End Select
    End With
End Property
Public Property Get Mat8_Col(m As Matrix8, ByVal index As Long) As Vector8
    With m
        Select Case index
        Case 0: Mat8_Col = Vec8(.aa, .ba, .ca, .da, .ea, .fa, .ga, .ha)
        Case 1: Mat8_Col = Vec8(.ab, .bb, .cb, .db, .eb, .fb, .gb, .hb)
        Case 2: Mat8_Col = Vec8(.ac, .bc, .cc, .dc, .ec, .fc, .gc, .hc)
        Case 3: Mat8_Col = Vec8(.ad, .bd, .cd, .dd, .ed, .fd, .gd, .hd)
        Case 4: Mat8_Col = Vec8(.ae, .be, .ce, .de, .ee, .fe, .ge, .he)
        Case 5: Mat8_Col = Vec8(.af, .bf, .cf, .df, .ef, .ff, .gf, .hf)
        Case 6: Mat8_Col = Vec8(.ag, .bg, .cg, .dg, .eg, .fg, .gg, .hg)
        Case 7: Mat8_Col = Vec8(.ah, .bh, .ch, .dh, .eh, .fh, .gh, .hh)
        End Select
    End With
End Property
Public Property Let Mat8_Col(m As Matrix8, ByVal index As Long, v As Vector8)
    With m
        Select Case index
        Case 0: .aa = v.a: .ba = v.b: .ca = v.c: .da = v.d: .ea = v.e: .fa = v.f: .ga = v.g: .ha = v.h
        Case 1: .ab = v.a: .bb = v.b: .cb = v.c: .db = v.d: .eb = v.e: .fb = v.f: .gb = v.g: .hb = v.h
        Case 2: .ac = v.a: .bc = v.b: .cc = v.c: .dc = v.d: .ec = v.e: .fc = v.f: .gc = v.g: .hc = v.h
        Case 3: .ad = v.a: .bd = v.b: .cd = v.c: .dd = v.d: .ed = v.e: .fd = v.f: .gd = v.g: .hd = v.h
        Case 4: .ae = v.a: .be = v.b: .ce = v.c: .de = v.d: .ee = v.e: .fe = v.f: .ge = v.g: .he = v.h
        Case 5: .af = v.a: .bf = v.b: .cf = v.c: .df = v.d: .ef = v.e: .ff = v.f: .gf = v.g: .hf = v.h
        Case 6: .ag = v.a: .bg = v.b: .cg = v.c: .dg = v.d: .eg = v.e: .fg = v.f: .gg = v.g: .hg = v.h
        Case 7: .ah = v.a: .bh = v.b: .ch = v.c: .dh = v.d: .eh = v.e: .fh = v.f: .gh = v.g: .hh = v.h
        End Select
    End With
End Property

Public Property Get Mat9_Row(m As Matrix9, ByVal index As Long) As Vector9
    With m
        Select Case index
        Case 0: Mat9_Row = Vec9(.aa, .ab, .ac, .ad, .ae, .af, .ag, .ah, .ai)
        Case 1: Mat9_Row = Vec9(.ba, .bb, .bc, .bd, .be, .bf, .bg, .bh, .bi)
        Case 2: Mat9_Row = Vec9(.ca, .cb, .cc, .cd, .ce, .cf, .cg, .ch, .ci)
        Case 3: Mat9_Row = Vec9(.da, .db, .dc, .dd, .de, .df, .dg, .dh, .di)
        Case 4: Mat9_Row = Vec9(.ea, .eb, .ec, .ed, .ee, .ef, .eg, .eh, .ei)
        Case 5: Mat9_Row = Vec9(.fa, .fb, .fc, .fd, .fe, .ff, .fg, .fh, .fi)
        Case 6: Mat9_Row = Vec9(.ga, .gb, .gc, .gd, .ge, .gf, .gg, .gh, .gi)
        Case 7: Mat9_Row = Vec9(.ha, .hb, .hc, .hd, .he, .hf, .hg, .hh, .hi)
        Case 8: Mat9_Row = Vec9(.ia, .ib, .ic, .id, .ie, .if, .ig, .ih, .ii)
        End Select
    End With
End Property
Public Property Let Mat9_Row(m As Matrix9, ByVal index As Long, v As Vector9)
    With m
        Select Case index
        Case 0: .aa = v.a: .ab = v.b: .ac = v.c: .ad = v.d: .ae = v.e: .af = v.f: .ag = v.g: .ah = v.h: .ai = v.i
        Case 1: .ba = v.a: .bb = v.b: .bc = v.c: .bd = v.d: .be = v.e: .bf = v.f: .bg = v.g: .bh = v.h: .bi = v.i
        Case 2: .ca = v.a: .cb = v.b: .cc = v.c: .cd = v.d: .ce = v.e: .cf = v.f: .cg = v.g: .ch = v.h: .ci = v.i
        Case 3: .da = v.a: .db = v.b: .dc = v.c: .dd = v.d: .de = v.e: .df = v.f: .dg = v.g: .dh = v.h: .di = v.i
        Case 4: .ea = v.a: .eb = v.b: .ec = v.c: .ed = v.d: .ee = v.e: .ef = v.f: .eg = v.g: .eh = v.h: .ei = v.i
        Case 5: .fa = v.a: .fb = v.b: .fc = v.c: .fd = v.d: .fe = v.e: .ff = v.f: .fg = v.g: .fh = v.h: .fi = v.i
        Case 6: .ga = v.a: .gb = v.b: .gc = v.c: .gd = v.d: .ge = v.e: .gf = v.f: .gg = v.g: .gh = v.h: .gi = v.i
        Case 7: .ha = v.a: .hb = v.b: .hc = v.c: .hd = v.d: .he = v.e: .hf = v.f: .hg = v.g: .hh = v.h: .hi = v.i
        Case 8: .ia = v.a: .ib = v.b: .ic = v.c: .id = v.d: .ie = v.e: .if = v.f: .ig = v.g: .ih = v.h: .ii = v.i
        End Select
    End With
End Property
Public Property Get Mat9_Col(m As Matrix9, ByVal index As Long) As Vector9
    With m
        Select Case index
        Case 0: Mat9_Col = Vec9(.aa, .ba, .ca, .da, .ea, .fa, .ga, .ha, .ia)
        Case 1: Mat9_Col = Vec9(.ab, .bb, .cb, .db, .eb, .fb, .gb, .hb, .ib)
        Case 2: Mat9_Col = Vec9(.ac, .bc, .cc, .dc, .ec, .fc, .gc, .hc, .ic)
        Case 3: Mat9_Col = Vec9(.ad, .bd, .cd, .dd, .ed, .fd, .gd, .hd, .id)
        Case 4: Mat9_Col = Vec9(.ae, .be, .ce, .de, .ee, .fe, .ge, .he, .ie)
        Case 5: Mat9_Col = Vec9(.af, .bf, .cf, .df, .ef, .ff, .gf, .hf, .if)
        Case 6: Mat9_Col = Vec9(.ag, .bg, .cg, .dg, .eg, .fg, .gg, .hg, .ig)
        Case 7: Mat9_Col = Vec9(.ah, .bh, .ch, .dh, .eh, .fh, .gh, .hh, .ih)
        Case 8: Mat9_Col = Vec9(.ai, .bi, .ci, .di, .ei, .fi, .gi, .hi, .ii)
        End Select
    End With
End Property
Public Property Let Mat9_Col(m As Matrix9, ByVal index As Long, v As Vector9)
    With m
        Select Case index
        Case 0: .aa = v.a: .ba = v.b: .ca = v.c: .da = v.d: .ea = v.e: .fa = v.f: .ga = v.g: .ha = v.h: .ia = v.i
        Case 1: .ab = v.a: .bb = v.b: .cb = v.c: .db = v.d: .eb = v.e: .fb = v.f: .gb = v.g: .hb = v.h: .ib = v.i
        Case 2: .ac = v.a: .bc = v.b: .cc = v.c: .dc = v.d: .ec = v.e: .fc = v.f: .gc = v.g: .hc = v.h: .ic = v.i
        Case 3: .ad = v.a: .bd = v.b: .cd = v.c: .dd = v.d: .ed = v.e: .fd = v.f: .gd = v.g: .hd = v.h: .id = v.i
        Case 4: .ae = v.a: .be = v.b: .ce = v.c: .de = v.d: .ee = v.e: .fe = v.f: .ge = v.g: .he = v.h: .ie = v.i
        Case 5: .af = v.a: .bf = v.b: .cf = v.c: .df = v.d: .ef = v.e: .ff = v.f: .gf = v.g: .hf = v.h: .if = v.i
        Case 6: .ag = v.a: .bg = v.b: .cg = v.c: .dg = v.d: .eg = v.e: .fg = v.f: .gg = v.g: .hg = v.h: .ig = v.i
        Case 7: .ah = v.a: .bh = v.b: .ch = v.c: .dh = v.d: .eh = v.e: .fh = v.f: .gh = v.g: .hh = v.h: .ih = v.i
        Case 8: .ai = v.a: .bi = v.b: .ci = v.c: .di = v.d: .ei = v.e: .fi = v.f: .gi = v.g: .hi = v.h: .ii = v.i
        End Select
    End With
End Property

Public Property Get Mat10_Row(m As Matrix10, ByVal index As Long) As Vector10
    With m
        Select Case index
        Case 0: Mat10_Row = Vec10(.aa, .ab, .ac, .ad, .ae, .af, .ag, .ah, .ai, .aj)
        Case 1: Mat10_Row = Vec10(.ba, .bb, .bc, .bd, .be, .bf, .bg, .bh, .bi, .bj)
        Case 2: Mat10_Row = Vec10(.ca, .cb, .cc, .cd, .ce, .cf, .cg, .ch, .ci, .cj)
        Case 3: Mat10_Row = Vec10(.da, .db, .dc, .dd, .de, .df, .dg, .dh, .di, .dj)
        Case 4: Mat10_Row = Vec10(.ea, .eb, .ec, .ed, .ee, .ef, .eg, .eh, .ei, .ej)
        Case 5: Mat10_Row = Vec10(.fa, .fb, .fc, .fd, .fe, .ff, .fg, .fh, .fi, .fj)
        Case 6: Mat10_Row = Vec10(.ga, .gb, .gc, .gd, .ge, .gf, .gg, .gh, .gi, .gj)
        Case 7: Mat10_Row = Vec10(.ha, .hb, .hc, .hd, .he, .hf, .hg, .hh, .hi, .hj)
        Case 8: Mat10_Row = Vec10(.ia, .ib, .ic, .id, .ie, .if, .ig, .ih, .ii, .ij)
        Case 9: Mat10_Row = Vec10(.ja, .jb, .jc, .jd, .je, .jf, .jg, .jh, .ji, .jj)
        End Select
    End With
End Property
Public Property Let Mat10_Row(m As Matrix10, ByVal index As Long, v As Vector10)
    With m
        Select Case index
        Case 0: .aa = v.a: .ab = v.b: .ac = v.c: .ad = v.d: .ae = v.e: .af = v.f: .ag = v.g: .ah = v.h: .ai = v.i: .aj = v.j
        Case 1: .ba = v.a: .bb = v.b: .bc = v.c: .bd = v.d: .be = v.e: .bf = v.f: .bg = v.g: .bh = v.h: .bi = v.i: .bj = v.j
        Case 2: .ca = v.a: .cb = v.b: .cc = v.c: .cd = v.d: .ce = v.e: .cf = v.f: .cg = v.g: .ch = v.h: .ci = v.i: .cj = v.j
        Case 3: .da = v.a: .db = v.b: .dc = v.c: .dd = v.d: .de = v.e: .df = v.f: .dg = v.g: .dh = v.h: .di = v.i: .dj = v.j
        Case 4: .ea = v.a: .eb = v.b: .ec = v.c: .ed = v.d: .ee = v.e: .ef = v.f: .eg = v.g: .eh = v.h: .ei = v.i: .ej = v.j
        Case 5: .fa = v.a: .fb = v.b: .fc = v.c: .fd = v.d: .fe = v.e: .ff = v.f: .fg = v.g: .fh = v.h: .fi = v.i: .fj = v.j
        Case 6: .ga = v.a: .gb = v.b: .gc = v.c: .gd = v.d: .ge = v.e: .gf = v.f: .gg = v.g: .gh = v.h: .gi = v.i: .gj = v.j
        Case 7: .ha = v.a: .hb = v.b: .hc = v.c: .hd = v.d: .he = v.e: .hf = v.f: .hg = v.g: .hh = v.h: .hi = v.i: .hj = v.j
        Case 8: .ia = v.a: .ib = v.b: .ic = v.c: .id = v.d: .ie = v.e: .if = v.f: .ig = v.g: .ih = v.h: .ii = v.i: .ij = v.j
        Case 9: .ja = v.a: .jb = v.b: .jc = v.c: .jd = v.d: .je = v.e: .jf = v.f: .jg = v.g: .jh = v.h: .ji = v.i: .jj = v.j
        End Select
    End With
End Property
Public Property Get Mat10_Col(m As Matrix10, ByVal index As Long) As Vector10
    With m
        Select Case index
        Case 0: Mat10_Col = Vec10(.aa, .ba, .ca, .da, .ea, .fa, .ga, .ha, .ia, .ja)
        Case 1: Mat10_Col = Vec10(.ab, .bb, .cb, .db, .eb, .fb, .gb, .hb, .ib, .jb)
        Case 2: Mat10_Col = Vec10(.ac, .bc, .cc, .dc, .ec, .fc, .gc, .hc, .ic, .jc)
        Case 3: Mat10_Col = Vec10(.ad, .bd, .cd, .dd, .ed, .fd, .gd, .hd, .id, .jd)
        Case 4: Mat10_Col = Vec10(.ae, .be, .ce, .de, .ee, .fe, .ge, .he, .ie, .je)
        Case 5: Mat10_Col = Vec10(.af, .bf, .cf, .df, .ef, .ff, .gf, .hf, .if, .jf)
        Case 6: Mat10_Col = Vec10(.ag, .bg, .cg, .dg, .eg, .fg, .gg, .hg, .ig, .jg)
        Case 7: Mat10_Col = Vec10(.ah, .bh, .ch, .dh, .eh, .fh, .gh, .hh, .ih, .jh)
        Case 8: Mat10_Col = Vec10(.ai, .bi, .ci, .di, .ei, .fi, .gi, .hi, .ii, .ji)
        Case 9: Mat10_Col = Vec10(.aj, .bj, .cj, .dj, .ej, .fj, .gj, .hj, .ij, .jj)
        End Select
    End With
End Property
Public Property Let Mat10_Col(m As Matrix10, ByVal index As Long, v As Vector10)
    With m
        Select Case index
        Case 0: .aa = v.a: .ba = v.b: .ca = v.c: .da = v.d: .ea = v.e: .fa = v.f: .ga = v.g: .ha = v.h: .ia = v.i: .ja = v.j
        Case 1: .ab = v.a: .bb = v.b: .cb = v.c: .db = v.d: .eb = v.e: .fb = v.f: .gb = v.g: .hb = v.h: .ib = v.i: .jb = v.j
        Case 2: .ac = v.a: .bc = v.b: .cc = v.c: .dc = v.d: .ec = v.e: .fc = v.f: .gc = v.g: .hc = v.h: .ic = v.i: .jc = v.j
        Case 3: .ad = v.a: .bd = v.b: .cd = v.c: .dd = v.d: .ed = v.e: .fd = v.f: .gd = v.g: .hd = v.h: .id = v.i: .jd = v.j
        Case 4: .ae = v.a: .be = v.b: .ce = v.c: .de = v.d: .ee = v.e: .fe = v.f: .ge = v.g: .he = v.h: .ie = v.i: .je = v.j
        Case 5: .af = v.a: .bf = v.b: .cf = v.c: .df = v.d: .ef = v.e: .ff = v.f: .gf = v.g: .hf = v.h: .if = v.i: .jf = v.j
        Case 6: .ag = v.a: .bg = v.b: .cg = v.c: .dg = v.d: .eg = v.e: .fg = v.f: .gg = v.g: .hg = v.h: .ig = v.i: .jg = v.j
        Case 7: .ah = v.a: .bh = v.b: .ch = v.c: .dh = v.d: .eh = v.e: .fh = v.f: .gh = v.g: .hh = v.h: .ih = v.i: .jh = v.j
        Case 8: .ai = v.a: .bi = v.b: .ci = v.c: .di = v.d: .ei = v.e: .fi = v.f: .gi = v.g: .hi = v.h: .ii = v.i: .ji = v.j
        Case 9: .aj = v.a: .bj = v.b: .cj = v.c: .dj = v.d: .ej = v.e: .fj = v.f: .gj = v.g: .hj = v.h: .ij = v.i: .jj = v.j
        End Select
    End With
End Property


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
Public Function Matrix7_umat(m As Matrix7, ByVal r_ex As Long, ByVal c_ex As Long) As Matrix6
    'Liefert aus einer 7x7-Matrix die 6x6-Untermatrix, indem die Zeile r_ex und die Spalte c_ex ausgelassen werden.
    Dim icex As Long
    If r_ex <> 0 Then Mat6_Row(Matrix7_umat, icex) = Vector7_uvec(Mat7_Row(m, 0), c_ex): icex = icex + 1
    If r_ex <> 1 Then Mat6_Row(Matrix7_umat, icex) = Vector7_uvec(Mat7_Row(m, 1), c_ex): icex = icex + 1
    If r_ex <> 2 Then Mat6_Row(Matrix7_umat, icex) = Vector7_uvec(Mat7_Row(m, 2), c_ex): icex = icex + 1
    If r_ex <> 3 Then Mat6_Row(Matrix7_umat, icex) = Vector7_uvec(Mat7_Row(m, 3), c_ex): icex = icex + 1
    If r_ex <> 4 Then Mat6_Row(Matrix7_umat, icex) = Vector7_uvec(Mat7_Row(m, 4), c_ex): icex = icex + 1
    If r_ex <> 5 Then Mat6_Row(Matrix7_umat, icex) = Vector7_uvec(Mat7_Row(m, 5), c_ex): icex = icex + 1
    If r_ex <> 6 Then Mat6_Row(Matrix7_umat, icex) = Vector7_uvec(Mat7_Row(m, 6), c_ex): icex = icex + 1
End Function
Public Function Matrix8_umat(m As Matrix8, ByVal r_ex As Long, ByVal c_ex As Long) As Matrix7
    'Liefert aus einer 8x8-Matrix die 7x7-Untermatrix, indem die Zeile r_ex und die Spalte c_ex ausgelassen werden.
    Dim icex As Long
    If r_ex <> 0 Then Mat7_Row(Matrix8_umat, icex) = Vector8_uvec(Mat8_Row(m, 0), c_ex): icex = icex + 1
    If r_ex <> 1 Then Mat7_Row(Matrix8_umat, icex) = Vector8_uvec(Mat8_Row(m, 1), c_ex): icex = icex + 1
    If r_ex <> 2 Then Mat7_Row(Matrix8_umat, icex) = Vector8_uvec(Mat8_Row(m, 2), c_ex): icex = icex + 1
    If r_ex <> 3 Then Mat7_Row(Matrix8_umat, icex) = Vector8_uvec(Mat8_Row(m, 3), c_ex): icex = icex + 1
    If r_ex <> 4 Then Mat7_Row(Matrix8_umat, icex) = Vector8_uvec(Mat8_Row(m, 4), c_ex): icex = icex + 1
    If r_ex <> 5 Then Mat7_Row(Matrix8_umat, icex) = Vector8_uvec(Mat8_Row(m, 5), c_ex): icex = icex + 1
    If r_ex <> 6 Then Mat7_Row(Matrix8_umat, icex) = Vector8_uvec(Mat8_Row(m, 6), c_ex): icex = icex + 1
    If r_ex <> 7 Then Mat7_Row(Matrix8_umat, icex) = Vector8_uvec(Mat8_Row(m, 7), c_ex): icex = icex + 1
End Function
Public Function Matrix9_umat(m As Matrix9, ByVal r_ex As Long, ByVal c_ex As Long) As Matrix8
    'Liefert aus einer 9x9-Matrix die 8x8-Untermatrix, indem die Zeile r_ex und die Spalte c_ex ausgelassen werden.
    Dim icex As Long
    If r_ex <> 0 Then Mat8_Row(Matrix9_umat, icex) = Vector9_uvec(Mat9_Row(m, 0), c_ex): icex = icex + 1
    If r_ex <> 1 Then Mat8_Row(Matrix9_umat, icex) = Vector9_uvec(Mat9_Row(m, 1), c_ex): icex = icex + 1
    If r_ex <> 2 Then Mat8_Row(Matrix9_umat, icex) = Vector9_uvec(Mat9_Row(m, 2), c_ex): icex = icex + 1
    If r_ex <> 3 Then Mat8_Row(Matrix9_umat, icex) = Vector9_uvec(Mat9_Row(m, 3), c_ex): icex = icex + 1
    If r_ex <> 4 Then Mat8_Row(Matrix9_umat, icex) = Vector9_uvec(Mat9_Row(m, 4), c_ex): icex = icex + 1
    If r_ex <> 5 Then Mat8_Row(Matrix9_umat, icex) = Vector9_uvec(Mat9_Row(m, 5), c_ex): icex = icex + 1
    If r_ex <> 6 Then Mat8_Row(Matrix9_umat, icex) = Vector9_uvec(Mat9_Row(m, 6), c_ex): icex = icex + 1
    If r_ex <> 7 Then Mat8_Row(Matrix9_umat, icex) = Vector9_uvec(Mat9_Row(m, 7), c_ex): icex = icex + 1
    If r_ex <> 8 Then Mat8_Row(Matrix9_umat, icex) = Vector9_uvec(Mat9_Row(m, 8), c_ex): icex = icex + 1
End Function
Public Function Matrix10_umat(m As Matrix10, ByVal r_ex As Long, ByVal c_ex As Long) As Matrix9
    'Liefert aus einer 10x10-Matrix die 9x9-Untermatrix, indem die Zeile r_ex und die Spalte c_ex ausgelassen werden.
    Dim icex As Long
    If r_ex <> 0 Then Mat9_Row(Matrix10_umat, icex) = Vector10_uvec(Mat10_Row(m, 0), c_ex): icex = icex + 1
    If r_ex <> 1 Then Mat9_Row(Matrix10_umat, icex) = Vector10_uvec(Mat10_Row(m, 1), c_ex): icex = icex + 1
    If r_ex <> 2 Then Mat9_Row(Matrix10_umat, icex) = Vector10_uvec(Mat10_Row(m, 2), c_ex): icex = icex + 1
    If r_ex <> 3 Then Mat9_Row(Matrix10_umat, icex) = Vector10_uvec(Mat10_Row(m, 3), c_ex): icex = icex + 1
    If r_ex <> 4 Then Mat9_Row(Matrix10_umat, icex) = Vector10_uvec(Mat10_Row(m, 4), c_ex): icex = icex + 1
    If r_ex <> 5 Then Mat9_Row(Matrix10_umat, icex) = Vector10_uvec(Mat10_Row(m, 5), c_ex): icex = icex + 1
    If r_ex <> 6 Then Mat9_Row(Matrix10_umat, icex) = Vector10_uvec(Mat10_Row(m, 6), c_ex): icex = icex + 1
    If r_ex <> 7 Then Mat9_Row(Matrix10_umat, icex) = Vector10_uvec(Mat10_Row(m, 7), c_ex): icex = icex + 1
    If r_ex <> 8 Then Mat9_Row(Matrix10_umat, icex) = Vector10_uvec(Mat10_Row(m, 8), c_ex): icex = icex + 1
    If r_ex <> 9 Then Mat9_Row(Matrix10_umat, icex) = Vector10_uvec(Mat10_Row(m, 9), c_ex): icex = icex + 1
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
    'Berechnet die Minore einer 6x6-Matrix bei der die Zeile r_ex und die Spalte c_ex ausgelassen werden.
    Matrix6_min = Matrix5_det(Matrix6_umat(m, r_ex, c_ex))
End Function
Public Function Matrix7_min(m As Matrix7, ByVal r_ex As Long, ByVal c_ex As Long) As Double
    'Berechnet die Minore einer 6x6-Matrix bei der die Zeile r_ex und die Spalte c_ex ausgelassen werden.
    Matrix7_min = Matrix6_det(Matrix7_umat(m, r_ex, c_ex))
End Function
Public Function Matrix8_min(m As Matrix8, ByVal r_ex As Long, ByVal c_ex As Long) As Double
    'Berechnet die Minore einer 6x6-Matrix bei der die Zeile r_ex und die Spalte c_ex ausgelassen werden.
    Matrix8_min = Matrix7_det(Matrix8_umat(m, r_ex, c_ex))
End Function
Public Function Matrix9_min(m As Matrix9, ByVal r_ex As Long, ByVal c_ex As Long) As Double
    'Berechnet die Minore einer 6x6-Matrix bei der die Zeile r_ex und die Spalte c_ex ausgelassen werden.
    Matrix9_min = Matrix8_det(Matrix9_umat(m, r_ex, c_ex))
End Function
Public Function Matrix10_min(m As Matrix10, ByVal r_ex As Long, ByVal c_ex As Long) As Double
    'Berechnet die Minore einer 6x6-Matrix bei der die Zeile r_ex und die Spalte c_ex ausgelassen werden.
    Matrix10_min = Matrix9_det(Matrix10_umat(m, r_ex, c_ex))
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
Public Function Matrix7_Adj(m As Matrix7) As Matrix7
    Matrix7_Adj = Mat7(Matrix7_min(m, 0, 0), -Matrix7_min(m, 1, 0), Matrix7_min(m, 2, 0), -Matrix7_min(m, 3, 0), Matrix7_min(m, 4, 0), -Matrix7_min(m, 5, 0), Matrix7_min(m, 6, 0), _
                       -Matrix7_min(m, 0, 1), Matrix7_min(m, 1, 1), -Matrix7_min(m, 2, 1), Matrix7_min(m, 3, 1), -Matrix7_min(m, 4, 1), Matrix7_min(m, 5, 1), -Matrix7_min(m, 6, 1), _
                       Matrix7_min(m, 0, 2), -Matrix7_min(m, 1, 2), Matrix7_min(m, 2, 2), -Matrix7_min(m, 3, 2), Matrix7_min(m, 4, 2), -Matrix7_min(m, 5, 2), Matrix7_min(m, 6, 2), _
                       -Matrix7_min(m, 0, 3), Matrix7_min(m, 1, 3), -Matrix7_min(m, 2, 3), Matrix7_min(m, 3, 3), -Matrix7_min(m, 4, 3), Matrix7_min(m, 5, 3), -Matrix7_min(m, 6, 3), _
                       Matrix7_min(m, 0, 4), -Matrix7_min(m, 1, 4), Matrix7_min(m, 2, 4), -Matrix7_min(m, 3, 4), Matrix7_min(m, 4, 4), -Matrix7_min(m, 5, 4), Matrix7_min(m, 6, 4), _
                       -Matrix7_min(m, 0, 5), Matrix7_min(m, 1, 5), -Matrix7_min(m, 2, 5), Matrix7_min(m, 3, 5), -Matrix7_min(m, 4, 5), Matrix7_min(m, 5, 5), -Matrix7_min(m, 6, 5), _
                       Matrix7_min(m, 0, 6), -Matrix7_min(m, 1, 6), Matrix7_min(m, 2, 6), -Matrix7_min(m, 3, 6), Matrix7_min(m, 4, 6), -Matrix7_min(m, 5, 6), Matrix7_min(m, 6, 6))
End Function
Public Function Matrix8_Adj(m As Matrix8) As Matrix8
    Matrix8_Adj = Mat8(Vec8(Matrix8_min(m, 0, 0), -Matrix8_min(m, 1, 0), Matrix8_min(m, 2, 0), -Matrix8_min(m, 3, 0), Matrix8_min(m, 4, 0), -Matrix8_min(m, 5, 0), Matrix8_min(m, 6, 0), -Matrix8_min(m, 7, 0)), _
                       Vec8(-Matrix8_min(m, 0, 1), Matrix8_min(m, 1, 1), -Matrix8_min(m, 2, 1), Matrix8_min(m, 3, 1), -Matrix8_min(m, 4, 1), Matrix8_min(m, 5, 1), -Matrix8_min(m, 6, 1), Matrix8_min(m, 7, 1)), _
                       Vec8(Matrix8_min(m, 0, 2), -Matrix8_min(m, 1, 2), Matrix8_min(m, 2, 2), -Matrix8_min(m, 3, 2), Matrix8_min(m, 4, 2), -Matrix8_min(m, 5, 2), Matrix8_min(m, 6, 2), -Matrix8_min(m, 7, 2)), _
                       Vec8(-Matrix8_min(m, 0, 3), Matrix8_min(m, 1, 3), -Matrix8_min(m, 2, 3), Matrix8_min(m, 3, 3), -Matrix8_min(m, 4, 3), Matrix8_min(m, 5, 3), -Matrix8_min(m, 6, 3), Matrix8_min(m, 7, 3)), _
                       Vec8(Matrix8_min(m, 0, 4), -Matrix8_min(m, 1, 4), Matrix8_min(m, 2, 4), -Matrix8_min(m, 3, 4), Matrix8_min(m, 4, 4), -Matrix8_min(m, 5, 4), Matrix8_min(m, 6, 4), -Matrix8_min(m, 7, 4)), _
                       Vec8(-Matrix8_min(m, 0, 5), Matrix8_min(m, 1, 5), -Matrix8_min(m, 2, 5), Matrix8_min(m, 3, 5), -Matrix8_min(m, 4, 5), Matrix8_min(m, 5, 5), -Matrix8_min(m, 6, 5), Matrix8_min(m, 7, 5)), _
                       Vec8(Matrix8_min(m, 0, 6), -Matrix8_min(m, 1, 6), Matrix8_min(m, 2, 6), -Matrix8_min(m, 3, 6), Matrix8_min(m, 4, 6), -Matrix8_min(m, 5, 6), Matrix8_min(m, 6, 6), -Matrix8_min(m, 7, 6)), _
                       Vec8(-Matrix8_min(m, 0, 7), Matrix8_min(m, 1, 7), -Matrix8_min(m, 2, 7), Matrix8_min(m, 3, 7), -Matrix8_min(m, 4, 7), Matrix8_min(m, 5, 7), -Matrix8_min(m, 6, 7), Matrix8_min(m, 7, 7)))
End Function
Public Function Matrix9_Adj(m As Matrix9) As Matrix9
    Matrix9_Adj = Mat9(Vec9(Matrix9_min(m, 0, 0), -Matrix9_min(m, 1, 0), Matrix9_min(m, 2, 0), -Matrix9_min(m, 3, 0), Matrix9_min(m, 4, 0), -Matrix9_min(m, 5, 0), Matrix9_min(m, 6, 0), -Matrix9_min(m, 7, 0), Matrix9_min(m, 8, 0)), _
                       Vec9(-Matrix9_min(m, 0, 1), Matrix9_min(m, 1, 1), -Matrix9_min(m, 2, 1), Matrix9_min(m, 3, 1), -Matrix9_min(m, 4, 1), Matrix9_min(m, 5, 1), -Matrix9_min(m, 6, 1), Matrix9_min(m, 7, 1), -Matrix9_min(m, 8, 1)), _
                       Vec9(Matrix9_min(m, 0, 2), -Matrix9_min(m, 1, 2), Matrix9_min(m, 2, 2), -Matrix9_min(m, 3, 2), Matrix9_min(m, 4, 2), -Matrix9_min(m, 5, 2), Matrix9_min(m, 6, 2), -Matrix9_min(m, 7, 2), Matrix9_min(m, 8, 2)), _
                       Vec9(-Matrix9_min(m, 0, 3), Matrix9_min(m, 1, 3), -Matrix9_min(m, 2, 3), Matrix9_min(m, 3, 3), -Matrix9_min(m, 4, 3), Matrix9_min(m, 5, 3), -Matrix9_min(m, 6, 3), Matrix9_min(m, 7, 3), -Matrix9_min(m, 8, 3)), _
                       Vec9(Matrix9_min(m, 0, 4), -Matrix9_min(m, 1, 4), Matrix9_min(m, 2, 4), -Matrix9_min(m, 3, 4), Matrix9_min(m, 4, 4), -Matrix9_min(m, 5, 4), Matrix9_min(m, 6, 4), -Matrix9_min(m, 7, 4), Matrix9_min(m, 8, 4)), _
                       Vec9(-Matrix9_min(m, 0, 5), Matrix9_min(m, 1, 5), -Matrix9_min(m, 2, 5), Matrix9_min(m, 3, 5), -Matrix9_min(m, 4, 5), Matrix9_min(m, 5, 5), -Matrix9_min(m, 6, 5), Matrix9_min(m, 7, 5), -Matrix9_min(m, 8, 5)), _
                       Vec9(Matrix9_min(m, 0, 6), -Matrix9_min(m, 1, 6), Matrix9_min(m, 2, 6), -Matrix9_min(m, 3, 6), Matrix9_min(m, 4, 6), -Matrix9_min(m, 5, 6), Matrix9_min(m, 6, 6), -Matrix9_min(m, 7, 6), Matrix9_min(m, 8, 6)), _
                       Vec9(-Matrix9_min(m, 0, 7), Matrix9_min(m, 1, 7), -Matrix9_min(m, 2, 7), Matrix9_min(m, 3, 7), -Matrix9_min(m, 4, 7), Matrix9_min(m, 5, 7), -Matrix9_min(m, 6, 7), Matrix9_min(m, 7, 7), -Matrix9_min(m, 8, 7)), _
                       Vec9(Matrix9_min(m, 0, 8), -Matrix9_min(m, 1, 8), Matrix9_min(m, 2, 8), -Matrix9_min(m, 3, 8), Matrix9_min(m, 4, 8), -Matrix9_min(m, 5, 8), Matrix9_min(m, 6, 8), -Matrix9_min(m, 7, 8), Matrix9_min(m, 8, 8)))
End Function
Public Function Matrix10_Adj(m As Matrix10) As Matrix10
    Matrix10_Adj = Mat10(Vec10(Matrix10_min(m, 0, 0), -Matrix10_min(m, 1, 0), Matrix10_min(m, 2, 0), -Matrix10_min(m, 3, 0), Matrix10_min(m, 4, 0), -Matrix10_min(m, 5, 0), Matrix10_min(m, 6, 0), -Matrix10_min(m, 7, 0), Matrix10_min(m, 8, 0), -Matrix10_min(m, 9, 0)), _
                         Vec10(-Matrix10_min(m, 0, 1), Matrix10_min(m, 1, 1), -Matrix10_min(m, 2, 1), Matrix10_min(m, 3, 1), -Matrix10_min(m, 4, 1), Matrix10_min(m, 5, 1), -Matrix10_min(m, 6, 1), Matrix10_min(m, 7, 1), -Matrix10_min(m, 8, 1), Matrix10_min(m, 9, 1)), _
                         Vec10(Matrix10_min(m, 0, 2), -Matrix10_min(m, 1, 2), Matrix10_min(m, 2, 2), -Matrix10_min(m, 3, 2), Matrix10_min(m, 4, 2), -Matrix10_min(m, 5, 2), Matrix10_min(m, 6, 2), -Matrix10_min(m, 7, 2), Matrix10_min(m, 8, 2), -Matrix10_min(m, 9, 2)), _
                         Vec10(-Matrix10_min(m, 0, 3), Matrix10_min(m, 1, 3), -Matrix10_min(m, 2, 3), Matrix10_min(m, 3, 3), -Matrix10_min(m, 4, 3), Matrix10_min(m, 5, 3), -Matrix10_min(m, 6, 3), Matrix10_min(m, 7, 3), -Matrix10_min(m, 8, 3), Matrix10_min(m, 9, 3)), _
                         Vec10(Matrix10_min(m, 0, 4), -Matrix10_min(m, 1, 4), Matrix10_min(m, 2, 4), -Matrix10_min(m, 3, 4), Matrix10_min(m, 4, 4), -Matrix10_min(m, 5, 4), Matrix10_min(m, 6, 4), -Matrix10_min(m, 7, 4), Matrix10_min(m, 8, 4), -Matrix10_min(m, 9, 4)), _
                         Vec10(-Matrix10_min(m, 0, 5), Matrix10_min(m, 1, 5), -Matrix10_min(m, 2, 5), Matrix10_min(m, 3, 5), -Matrix10_min(m, 4, 5), Matrix10_min(m, 5, 5), -Matrix10_min(m, 6, 5), Matrix10_min(m, 7, 5), -Matrix10_min(m, 8, 5), Matrix10_min(m, 9, 5)), _
                         Vec10(Matrix10_min(m, 0, 6), -Matrix10_min(m, 1, 6), Matrix10_min(m, 2, 6), -Matrix10_min(m, 3, 6), Matrix10_min(m, 4, 6), -Matrix10_min(m, 5, 6), Matrix10_min(m, 6, 6), -Matrix10_min(m, 7, 6), Matrix10_min(m, 8, 6), -Matrix10_min(m, 9, 6)), _
                         Vec10(-Matrix10_min(m, 0, 7), Matrix10_min(m, 1, 7), -Matrix10_min(m, 2, 7), Matrix10_min(m, 3, 7), -Matrix10_min(m, 4, 7), Matrix10_min(m, 5, 7), -Matrix10_min(m, 6, 7), Matrix10_min(m, 7, 7), -Matrix10_min(m, 8, 7), Matrix10_min(m, 9, 7)), _
                         Vec10(Matrix10_min(m, 0, 8), -Matrix10_min(m, 1, 8), Matrix10_min(m, 2, 8), -Matrix10_min(m, 3, 8), Matrix10_min(m, 4, 8), -Matrix10_min(m, 5, 8), Matrix10_min(m, 6, 8), -Matrix10_min(m, 7, 8), Matrix10_min(m, 8, 8), -Matrix10_min(m, 9, 8)), _
                         Vec10(-Matrix10_min(m, 0, 9), Matrix10_min(m, 1, 9), -Matrix10_min(m, 2, 9), Matrix10_min(m, 3, 9), -Matrix10_min(m, 4, 9), Matrix10_min(m, 5, 9), -Matrix10_min(m, 6, 9), Matrix10_min(m, 7, 9), -Matrix10_min(m, 8, 9), Matrix10_min(m, 9, 9)))
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
Public Function Matrix7_inv(m As Matrix7) As Matrix7
    Dim det As Double: det = Matrix7_det(m): If det = 0 Then Exit Function
    Matrix7_inv = Matrix7_smul(Matrix7_Adj(m), 1 / det)
End Function
Public Function Matrix8_inv(m As Matrix8) As Matrix8
    Dim det As Double: det = Matrix8_det(m): If det = 0 Then Exit Function
    Matrix8_inv = Matrix8_smul(Matrix8_Adj(m), 1 / det)
End Function
Public Function Matrix9_inv(m As Matrix9) As Matrix9
    Dim det As Double: det = Matrix9_det(m): If det = 0 Then Exit Function
    Matrix9_inv = Matrix9_smul(Matrix9_Adj(m), 1 / det)
End Function
Public Function Matrix10_inv(m As Matrix10) As Matrix10
    Dim det As Double: det = Matrix10_det(m): If det = 0 Then Exit Function
    Matrix10_inv = Matrix10_smul(Matrix10_Adj(m), 1 / det)
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
Public Function Matrix7_solve(m As Matrix7, b As Vector7) As Vector7
    Matrix7_solve = Matrix7_vmul(Matrix7_inv(m), b)
End Function
Public Function Matrix8_solve(m As Matrix8, b As Vector8) As Vector8
    Matrix8_solve = Matrix8_vmul(Matrix8_inv(m), b)
End Function
Public Function Matrix9_solve(m As Matrix9, b As Vector9) As Vector9
    Matrix9_solve = Matrix9_vmul(Matrix9_inv(m), b)
End Function
Public Function Matrix10_solve(m As Matrix10, b As Vector10) As Vector10
    Matrix10_solve = Matrix10_vmul(Matrix10_inv(m), b)
End Function

Public Function Matrix2_Rnd() As Matrix2
    Dim d() As Double: d = Matrix_Random(2)
    RtlMoveMemory Matrix2_Rnd, d(0), 2 ^ 2 * 8
End Function
Public Function Matrix3_Rnd() As Matrix3
    Dim d() As Double: d = Matrix_Random(3)
    RtlMoveMemory Matrix3_Rnd, d(0), 3 ^ 2 * 8
End Function
Public Function Matrix4_Rnd() As Matrix4
    Dim d() As Double: d = Matrix_Random(4)
    RtlMoveMemory Matrix4_Rnd, d(0), 4 ^ 2 * 8
End Function
Public Function Matrix5_Rnd() As Matrix5
    Dim d() As Double: d = Matrix_Random(5)
    RtlMoveMemory Matrix5_Rnd, d(0), 5 ^ 2 * 8
End Function
Public Function Matrix6_Rnd() As Matrix6
    Dim d() As Double: d = Matrix_Random(6)
    RtlMoveMemory Matrix6_Rnd, d(0), 6 ^ 2 * 8
End Function
Public Function Matrix7_Rnd() As Matrix7
    Dim d() As Double: d = Matrix_Random(7)
    RtlMoveMemory Matrix7_Rnd, d(0), 7 ^ 2 * 8
End Function
Public Function Matrix8_Rnd() As Matrix8
    Dim d() As Double: d = Matrix_Random(8)
    RtlMoveMemory Matrix8_Rnd, d(0), 8 ^ 2 * 8
End Function
Public Function Matrix9_Rnd() As Matrix9
    Dim d() As Double: d = Matrix_Random(9)
    RtlMoveMemory Matrix9_Rnd, d(0), 9 ^ 2 * 8
End Function
Public Function Matrix10_Rnd() As Matrix10
    Dim d() As Double: d = Matrix_Random(10)
    RtlMoveMemory Matrix10_Rnd, d(0), 10 ^ 2 * 8
End Function

Public Function Matrix_Random(ByVal rc As Byte) As Double()
    Dim u As Long: u = rc * rc - 1
    ReDim d(0 To u) As Double
    Randomize
    Dim i As Long
    For i = 0 To u
        d(i) = Rnd() * 200 - 100
    Next
    Matrix_Random = d
End Function
