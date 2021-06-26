Attribute VB_Name = "MMatAlt"
Option Explicit

'Public Function Matrix4_detX(m As Matrix4) As Double
'    'Berechnet die Determinante einer 4x4-Matrix
'    With m
'        Matrix4_detX = .aa * Matrix3_det(Mat3(.bb, .bc, .bd, .cb, .cc, .cd, .db, .dc, .dd)) _
'                     - .ab * Matrix3_det(Mat3(.ba, .bc, .bd, .ca, .cc, .cd, .da, .dc, .dd)) _
'                     + .ac * Matrix3_det(Mat3(.ba, .bb, .bd, .ca, .cb, .cd, .da, .db, .dd)) _
'                     - .ad * Matrix3_det(Mat3(.ba, .bb, .bc, .ca, .cb, .cc, .da, .db, .dc))
'    End With
'End Function
'
'Public Function Matrix5_detX(m As Matrix5) As Double
'    'Berechnet die Determinante einer 5x5-Matrix
'    With m
'        Matrix5_detX = .aa * Matrix4_detX(Mat4(.bb, .bc, .bd, .be, .cb, .cc, .cd, .ce, .db, .dc, .dd, .de, .eb, .ec, .ed, .ee)) _
'                     - .ab * Matrix4_detX(Mat4(.ba, .bc, .bd, .be, .ca, .cc, .cd, .ce, .da, .dc, .dd, .de, .ea, .ec, .ed, .ee)) _
'                     + .ac * Matrix4_detX(Mat4(.ba, .bb, .bd, .be, .ca, .cb, .cd, .ce, .da, .db, .dd, .de, .ea, .eb, .ed, .ee)) _
'                     - .ad * Matrix4_detX(Mat4(.ba, .bb, .bc, .be, .ca, .cb, .cc, .ce, .da, .db, .dc, .de, .ea, .eb, .ec, .ee)) _
'                     + .ae * Matrix4_detX(Mat4(.ba, .bb, .bc, .bd, .ca, .cb, .cc, .cd, .da, .db, .dc, .dd, .ea, .eb, .ec, .ed))
'    End With
'End Function
'
'Public Function Matrix6_detX(m As Matrix6) As Double
'    'Berechnet die Determinante einer 6x6-Matrix
'    With m
'        Matrix6_detX = .aa * Matrix5_detX(Mat5(.bb, .bc, .bd, .be, .bf, .cb, .cc, .cd, .ce, .cf, .db, .dc, .dd, .de, .df, .eb, .ec, .ed, .ee, .ef, .fb, .fc, .fd, .fe, .ff)) _
'                     - .ab * Matrix5_detX(Mat5(.ba, .bc, .bd, .be, .bf, .ca, .cc, .cd, .ce, .cf, .da, .dc, .dd, .de, .df, .ea, .ec, .ed, .ee, .ef, .fa, .fc, .fd, .fe, .ff)) _
'                     + .ac * Matrix5_detX(Mat5(.ba, .bb, .bd, .be, .bf, .ca, .cb, .cd, .ce, .cf, .da, .db, .dd, .de, .df, .ea, .eb, .ed, .ee, .ef, .fa, .fb, .fd, .fe, .ff)) _
'                     - .ad * Matrix5_detX(Mat5(.ba, .bb, .bc, .be, .bf, .ca, .cb, .cc, .ce, .cf, .da, .db, .dc, .de, .df, .ea, .eb, .ec, .ee, .ef, .fa, .fb, .fc, .fe, .ff)) _
'                     + .ae * Matrix5_detX(Mat5(.ba, .bb, .bc, .bd, .bf, .ca, .cb, .cc, .cd, .cf, .da, .db, .dc, .dd, .df, .ea, .eb, .ec, .ed, .ef, .fa, .fb, .fc, .fd, .ff)) _
'                     - .af * Matrix5_detX(Mat5(.ba, .bb, .bc, .bd, .be, .ca, .cb, .cc, .cd, .ce, .da, .db, .dc, .dd, .de, .ea, .eb, .ec, .ed, .ee, .fa, .fb, .fc, .fd, .fe))
'    End With
'End Function
'
'Public Function Matrix7_detX(m As Matrix7) As Double
'    'Berechnet die Determinante einer 7x7-Matrix
'    With m
'        Matrix7_detX = .aa * Matrix6_detX(Mat6(.bb, .bc, .bd, .be, .bf, .bg, .cb, .cc, .cd, .ce, .cf, .cg, .db, .dc, .dd, .de, .df, .dg, .eb, .ec, .ed, .ee, .ef, .eg, .fb, .fc, .fd, .fe, .ff, .fg, .gb, .gc, .gd, .ge, .gf, .gg)) _
'                     - .ab * Matrix6_detX(Mat6(.ba, .bc, .bd, .be, .bf, .bg, .ca, .cc, .cd, .ce, .cf, .cg, .da, .dc, .dd, .de, .df, .dg, .ea, .ec, .ed, .ee, .ef, .eg, .fa, .fc, .fd, .fe, .ff, .fg, .ga, .gc, .gd, .ge, .gf, .gg)) _
'                     + .ac * Matrix6_detX(Mat6(.ba, .bb, .bd, .be, .bf, .bg, .ca, .cb, .cd, .ce, .cf, .cg, .da, .db, .dd, .de, .df, .dg, .ea, .eb, .ed, .ee, .ef, .eg, .fa, .fb, .fd, .fe, .ff, .fg, .ga, .gb, .gd, .ge, .gf, .gg)) _
'                     - .ad * Matrix6_detX(Mat6(.ba, .bb, .bc, .be, .bf, .bg, .ca, .cb, .cc, .ce, .cf, .cg, .da, .db, .dc, .de, .df, .dg, .ea, .eb, .ec, .ee, .ef, .eg, .fa, .fb, .fc, .fe, .ff, .fg, .ga, .gb, .gc, .ge, .gf, .gg)) _
'                     + .ae * Matrix6_detX(Mat6(.ba, .bb, .bc, .bd, .bf, .bg, .ca, .cb, .cc, .cd, .cf, .cg, .da, .db, .dc, .dd, .df, .dg, .ea, .eb, .ec, .ed, .ef, .eg, .fa, .fb, .fc, .fd, .ff, .fg, .ga, .gb, .gc, .gd, .gf, .gg)) _
'                     - .af * Matrix6_detX(Mat6(.ba, .bb, .bc, .bd, .be, .bg, .ca, .cb, .cc, .cd, .ce, .cg, .da, .db, .dc, .dd, .de, .dg, .ea, .eb, .ec, .ed, .ee, .eg, .fa, .fb, .fc, .fd, .fe, .fg, .ga, .gb, .gc, .gd, .ge, .gg)) _
'                     + .ag * Matrix6_detX(Mat6(.ba, .bb, .bc, .bd, .be, .bf, .ca, .cb, .cc, .cd, .ce, .cf, .da, .db, .dc, .dd, .de, .df, .ea, .eb, .ec, .ed, .ee, .ef, .fa, .fb, .fc, .fd, .fe, .ff, .ga, .gb, .gc, .gd, .ge, .gf))
'    End With
'End Function
'
'Public Function Matrix8_detX(m As Matrix8) As Double
'    'Berechnet die Determinante einer 8x8-Matrix
'    With m
'        Matrix8_detX = .aa * Matrix7_detX(Mat7(.bb, .bc, .bd, .be, .bf, .bg, .bh, .cb, .cc, .cd, .ce, .cf, .cg, .ch, .db, .dc, .dd, .de, .df, .dg, .dh, .eb, .ec, .ed, .ee, .ef, .eg, .eh, .fb, .fc, .fd, .fe, .ff, .fg, .fh, .gb, .gc, .gd, .ge, .gf, .gg, .gh, .hb, .HC, .hd, .he, .hf, .hg, .hh)) _
'                     - .ab * Matrix7_detX(Mat7(.ba, .bc, .bd, .be, .bf, .bg, .bh, .ca, .cc, .cd, .ce, .cf, .cg, .ch, .da, .dc, .dd, .de, .df, .dg, .dh, .ea, .ec, .ed, .ee, .ef, .eg, .eh, .fa, .fc, .fd, .fe, .ff, .fg, .fh, .ga, .gc, .gd, .ge, .gf, .gg, .gh, .ha, .HC, .hd, .he, .hf, .hg, .hh)) _
'                     + .ac * Matrix7_detX(Mat7(.ba, .bb, .bd, .be, .bf, .bg, .bh, .ca, .cb, .cd, .ce, .cf, .cg, .ch, .da, .db, .dd, .de, .df, .dg, .dh, .ea, .eb, .ed, .ee, .ef, .eg, .eh, .fa, .fb, .fd, .fe, .ff, .fg, .fh, .ga, .gb, .gd, .ge, .gf, .gg, .gh, .ha, .hb, .hd, .he, .hf, .hg, .hh)) _
'                     - .ad * Matrix7_detX(Mat7(.ba, .bb, .bc, .be, .bf, .bg, .bh, .ca, .cb, .cc, .ce, .cf, .cg, .ch, .da, .db, .dc, .de, .df, .dg, .dh, .ea, .eb, .ec, .ee, .ef, .eg, .eh, .fa, .fb, .fc, .fe, .ff, .fg, .fh, .ga, .gb, .gc, .ge, .gf, .gg, .gh, .ha, .hb, .HC, .he, .hf, .hg, .hh)) _
'                     + .ae * Matrix7_detX(Mat7(.ba, .bb, .bc, .bd, .bf, .bg, .bh, .ca, .cb, .cc, .cd, .cf, .cg, .ch, .da, .db, .dc, .dd, .df, .dg, .dh, .ea, .eb, .ec, .ed, .ef, .eg, .eh, .fa, .fb, .fc, .fd, .ff, .fg, .fh, .ga, .gb, .gc, .gd, .gf, .gg, .gh, .ha, .hb, .HC, .hd, .hf, .hg, .hh)) _
'                     - .af * Matrix7_detX(Mat7(.ba, .bb, .bc, .bd, .be, .bg, .bh, .ca, .cb, .cc, .cd, .ce, .cg, .ch, .da, .db, .dc, .dd, .de, .dg, .dh, .ea, .eb, .ec, .ed, .ee, .eg, .eh, .fa, .fb, .fc, .fd, .fe, .fg, .fh, .ga, .gb, .gc, .gd, .ge, .gg, .gh, .ha, .hb, .HC, .hd, .he, .hg, .hh)) _
'                     + .ag * Matrix7_detX(Mat7(.ba, .bb, .bc, .bd, .be, .bf, .bh, .ca, .cb, .cc, .cd, .ce, .cf, .ch, .da, .db, .dc, .dd, .de, .df, .dh, .ea, .eb, .ec, .ed, .ee, .ef, .eh, .fa, .fb, .fc, .fd, .fe, .ff, .fh, .ga, .gb, .gc, .gd, .ge, .gf, .gh, .ha, .hb, .HC, .hd, .he, .hf, .hh)) _
'                     - .ah * Matrix7_detX(Mat7(.ba, .bb, .bc, .bd, .be, .bf, .bg, .ca, .cb, .cc, .cd, .ce, .cf, .cg, .da, .db, .dc, .dd, .de, .df, .dg, .ea, .eb, .ec, .ed, .ee, .ef, .eg, .fa, .fb, .fc, .fd, .fe, .ff, .fg, .ga, .gb, .gc, .gd, .ge, .gf, .gg, .ha, .hb, .HC, .hd, .he, .hf, .hg))
'    End With
'End Function
'Public Function Matrix9_detX(m As Matrix9) As Double
'    'Berechnet die Determinante einer 9x9-Matrix
'    With m
'        Matrix9_detX = .aa * Matrix8_detX(Mat8(Vec8(.bb, .bc, .bd, .be, .bf, .bg, .bh, .bi), Vec8(.cb, .cc, .cd, .ce, .cf, .cg, .ch, .ci), Vec8(.db, .dc, .dd, .de, .df, .dg, .dh, .di), Vec8(.eb, .ec, .ed, .ee, .ef, .eg, .eh, .ei), Vec8(.fb, .fc, .fd, .fe, .ff, .fg, .fh, .fi), Vec8(.gb, .gc, .gd, .ge, .gf, .gg, .gh, .gi), Vec8(.hb, .HC, .hd, .he, .hf, .hg, .hh, .hi), Vec8(.ib, .ic, .id, .ie, .if, .ig, .ih, .ii))) _
'                     - .ab * Matrix8_detX(Mat8(Vec8(.ba, .bc, .bd, .be, .bf, .bg, .bh, .bi), Vec8(.ca, .cc, .cd, .ce, .cf, .cg, .ch, .ci), Vec8(.da, .dc, .dd, .de, .df, .dg, .dh, .di), Vec8(.ea, .ec, .ed, .ee, .ef, .eg, .eh, .ei), Vec8(.fa, .fc, .fd, .fe, .ff, .fg, .fh, .fi), Vec8(.ga, .gc, .gd, .ge, .gf, .gg, .gh, .gi), Vec8(.ha, .HC, .hd, .he, .hf, .hg, .hh, .hi), Vec8(.ia, .ic, .id, .ie, .if, .ig, .ih, .ii))) _
'                     + .ac * Matrix8_detX(Mat8(Vec8(.ba, .bb, .bd, .be, .bf, .bg, .bh, .bi), Vec8(.ca, .cb, .cd, .ce, .cf, .cg, .ch, .ci), Vec8(.da, .db, .dd, .de, .df, .dg, .dh, .di), Vec8(.ea, .eb, .ed, .ee, .ef, .eg, .eh, .ei), Vec8(.fa, .fb, .fd, .fe, .ff, .fg, .fh, .fi), Vec8(.ga, .gb, .gd, .ge, .gf, .gg, .gh, .gi), Vec8(.ha, .hb, .hd, .he, .hf, .hg, .hh, .hi), Vec8(.ia, .ib, .id, .ie, .if, .ig, .ih, .ii))) _
'                     - .ad * Matrix8_detX(Mat8(Vec8(.ba, .bb, .bc, .be, .bf, .bg, .bh, .bi), Vec8(.ca, .cb, .cc, .ce, .cf, .cg, .ch, .ci), Vec8(.da, .db, .dc, .de, .df, .dg, .dh, .di), Vec8(.ea, .eb, .ec, .ee, .ef, .eg, .eh, .ei), Vec8(.fa, .fb, .fc, .fe, .ff, .fg, .fh, .fi), Vec8(.ga, .gb, .gc, .ge, .gf, .gg, .gh, .gi), Vec8(.ha, .hb, .HC, .he, .hf, .hg, .hh, .hi), Vec8(.ia, .ib, .ic, .ie, .if, .ig, .ih, .ii))) _
'                     + .ae * Matrix8_detX(Mat8(Vec8(.ba, .bb, .bc, .bd, .bf, .bg, .bh, .bi), Vec8(.ca, .cb, .cc, .cd, .cf, .cg, .ch, .ci), Vec8(.da, .db, .dc, .dd, .df, .dg, .dh, .di), Vec8(.ea, .eb, .ec, .ed, .ef, .eg, .eh, .ei), Vec8(.fa, .fb, .fc, .fd, .ff, .fg, .fh, .fi), Vec8(.ga, .gb, .gc, .gd, .gf, .gg, .gh, .gi), Vec8(.ha, .hb, .HC, .hd, .hf, .hg, .hh, .hi), Vec8(.ia, .ib, .ic, .id, .if, .ig, .ih, .ii))) _
'                     - .af * Matrix8_detX(Mat8(Vec8(.ba, .bb, .bc, .bd, .be, .bg, .bh, .bi), Vec8(.ca, .cb, .cc, .cd, .ce, .cg, .ch, .ci), Vec8(.da, .db, .dc, .dd, .de, .dg, .dh, .di), Vec8(.ea, .eb, .ec, .ed, .ee, .eg, .eh, .ei), Vec8(.fa, .fb, .fc, .fd, .fe, .fg, .fh, .fi), Vec8(.ga, .gb, .gc, .gd, .ge, .gg, .gh, .gi), Vec8(.ha, .hb, .HC, .hd, .he, .hg, .hh, .hi), Vec8(.ia, .ib, .ic, .id, .ie, .ig, .ih, .ii))) _
'                     + .ag * Matrix8_detX(Mat8(Vec8(.ba, .bb, .bc, .bd, .be, .bf, .bh, .bi), Vec8(.ca, .cb, .cc, .cd, .ce, .cf, .ch, .ci), Vec8(.da, .db, .dc, .dd, .de, .df, .dh, .di), Vec8(.ea, .eb, .ec, .ed, .ee, .ef, .eh, .ei), Vec8(.fa, .fb, .fc, .fd, .fe, .ff, .fh, .fi), Vec8(.ga, .gb, .gc, .gd, .ge, .gf, .gh, .gi), Vec8(.ha, .hb, .HC, .hd, .he, .hf, .hh, .hi), Vec8(.ia, .ib, .ic, .id, .ie, .if, .ih, .ii))) _
'                     - .ah * Matrix8_detX(Mat8(Vec8(.ba, .bb, .bc, .bd, .be, .bf, .bg, .bi), Vec8(.ca, .cb, .cc, .cd, .ce, .cf, .cg, .ci), Vec8(.da, .db, .dc, .dd, .de, .df, .dg, .di), Vec8(.ea, .eb, .ec, .ed, .ee, .ef, .eg, .ei), Vec8(.fa, .fb, .fc, .fd, .fe, .ff, .fg, .fi), Vec8(.ga, .gb, .gc, .gd, .ge, .gf, .gg, .gi), Vec8(.ha, .hb, .HC, .hd, .he, .hf, .hg, .hi), Vec8(.ia, .ib, .ic, .id, .ie, .if, .ig, .ii))) _
'                     + .ai * Matrix8_detX(Mat8(Vec8(.ba, .bb, .bc, .bd, .be, .bf, .bg, .bh), Vec8(.ca, .cb, .cc, .cd, .ce, .cf, .cg, .ch), Vec8(.da, .db, .dc, .dd, .de, .df, .dg, .dh), Vec8(.ea, .eb, .ec, .ed, .ee, .ef, .eg, .eh), Vec8(.fa, .fb, .fc, .fd, .fe, .ff, .fg, .fh), Vec8(.ga, .gb, .gc, .gd, .ge, .gf, .gg, .gh), Vec8(.ha, .hb, .HC, .hd, .he, .hf, .hg, .hh), Vec8(.ia, .ib, .ic, .id, .ie, .if, .ig, .ih)))
'    End With
'End Function
'Public Function Matrix10_detX(m As Matrix10) As Double
'    'Berechnet die Determinante einer 10x10-Matrix
'    With m
'       Matrix10_detX = .aa * Matrix9_detX(Mat9(Vec9(.bb, .bc, .bd, .be, .bf, .bg, .bh, .bi, .bj), Vec9(.cb, .cc, .cd, .ce, .cf, .cg, .ch, .ci, .cj), Vec9(.db, .dc, .dd, .de, .df, .dg, .dh, .di, .dj), Vec9(.eb, .ec, .ed, .ee, .ef, .eg, .eh, .ei, .ej), Vec9(.fb, .fc, .fd, .fe, .ff, .fg, .fh, .fi, .fj), Vec9(.gb, .gc, .gd, .ge, .gf, .gg, .gh, .gi, .gj), Vec9(.hb, .HC, .hd, .he, .hf, .hg, .hh, .hi, .hj), Vec9(.ib, .ic, .id, .ie, .if, .ig, .ih, .ii, .ij), Vec9(.jb, .jc, .jd, .je, .jf, .jg, .jh, .ji, .jj))) _
'                     - .ab * Matrix9_detX(Mat9(Vec9(.ba, .bc, .bd, .be, .bf, .bg, .bh, .bi, .bj), Vec9(.ca, .cc, .cd, .ce, .cf, .cg, .ch, .ci, .cj), Vec9(.da, .dc, .dd, .de, .df, .dg, .dh, .di, .dj), Vec9(.ea, .ec, .ed, .ee, .ef, .eg, .eh, .ei, .ej), Vec9(.fa, .fc, .fd, .fe, .ff, .fg, .fh, .fi, .fj), Vec9(.ga, .gc, .gd, .ge, .gf, .gg, .gh, .gi, .gj), Vec9(.ha, .HC, .hd, .he, .hf, .hg, .hh, .hi, .hj), Vec9(.ia, .ic, .id, .ie, .if, .ig, .ih, .ii, .ij), Vec9(.ja, .jc, .jd, .je, .jf, .jg, .jh, .ji, .jj))) _
'                     + .ac * Matrix9_detX(Mat9(Vec9(.ba, .bb, .bd, .be, .bf, .bg, .bh, .bi, .bj), Vec9(.ca, .cb, .cd, .ce, .cf, .cg, .ch, .ci, .cj), Vec9(.da, .db, .dd, .de, .df, .dg, .dh, .di, .dj), Vec9(.ea, .eb, .ed, .ee, .ef, .eg, .eh, .ei, .ej), Vec9(.fa, .fb, .fd, .fe, .ff, .fg, .fh, .fi, .fj), Vec9(.ga, .gb, .gd, .ge, .gf, .gg, .gh, .gi, .gj), Vec9(.ha, .hb, .hd, .he, .hf, .hg, .hh, .hi, .hj), Vec9(.ia, .ib, .id, .ie, .if, .ig, .ih, .ii, .ij), Vec9(.ja, .jb, .jd, .je, .jf, .jg, .jh, .ji, .jj))) _
'                     - .ad * Matrix9_detX(Mat9(Vec9(.ba, .bb, .bc, .be, .bf, .bg, .bh, .bi, .bj), Vec9(.ca, .cb, .cc, .ce, .cf, .cg, .ch, .ci, .cj), Vec9(.da, .db, .dc, .de, .df, .dg, .dh, .di, .dj), Vec9(.ea, .eb, .ec, .ee, .ef, .eg, .eh, .ei, .ej), Vec9(.fa, .fb, .fc, .fe, .ff, .fg, .fh, .fi, .fj), Vec9(.ga, .gb, .gc, .ge, .gf, .gg, .gh, .gi, .gj), Vec9(.ha, .hb, .HC, .he, .hf, .hg, .hh, .hi, .hj), Vec9(.ia, .ib, .ic, .ie, .if, .ig, .ih, .ii, .ij), Vec9(.ja, .jb, .jc, .je, .jf, .jg, .jh, .ji, .jj))) _
'                     + .ae * Matrix9_detX(Mat9(Vec9(.ba, .bb, .bc, .bd, .bf, .bg, .bh, .bi, .bj), Vec9(.ca, .cb, .cc, .cd, .cf, .cg, .ch, .ci, .cj), Vec9(.da, .db, .dc, .dd, .df, .dg, .dh, .di, .dj), Vec9(.ea, .eb, .ec, .ed, .ef, .eg, .eh, .ei, .ej), Vec9(.fa, .fb, .fc, .fd, .ff, .fg, .fh, .fi, .fj), Vec9(.ga, .gb, .gc, .gd, .gf, .gg, .gh, .gi, .gj), Vec9(.ha, .hb, .HC, .hd, .hf, .hg, .hh, .hi, .hj), Vec9(.ia, .ib, .ic, .id, .if, .ig, .ih, .ii, .ij), Vec9(.ja, .jb, .jc, .jd, .jf, .jg, .jh, .ji, .jj))) _
'                     - .af * Matrix9_detX(Mat9(Vec9(.ba, .bb, .bc, .bd, .be, .bg, .bh, .bi, .bj), Vec9(.ca, .cb, .cc, .cd, .ce, .cg, .ch, .ci, .cj), Vec9(.da, .db, .dc, .dd, .de, .dg, .dh, .di, .dj), Vec9(.ea, .eb, .ec, .ed, .ee, .eg, .eh, .ei, .ej), Vec9(.fa, .fb, .fc, .fd, .fe, .fg, .fh, .fi, .fj), Vec9(.ga, .gb, .gc, .gd, .ge, .gg, .gh, .gi, .gj), Vec9(.ha, .hb, .HC, .hd, .he, .hg, .hh, .hi, .hj), Vec9(.ia, .ib, .ic, .id, .ie, .ig, .ih, .ii, .ij), Vec9(.ja, .jb, .jc, .jd, .je, .jg, .jh, .ji, .jj))) _
'                     + .ag * Matrix9_detX(Mat9(Vec9(.ba, .bb, .bc, .bd, .be, .bf, .bh, .bi, .bj), Vec9(.ca, .cb, .cc, .cd, .ce, .cf, .ch, .ci, .cj), Vec9(.da, .db, .dc, .dd, .de, .df, .dh, .di, .dj), Vec9(.ea, .eb, .ec, .ed, .ee, .ef, .eh, .ei, .ej), Vec9(.fa, .fb, .fc, .fd, .fe, .ff, .fh, .fi, .fj), Vec9(.ga, .gb, .gc, .gd, .ge, .gf, .gh, .gi, .gj), Vec9(.ha, .hb, .HC, .hd, .he, .hf, .hh, .hi, .hj), Vec9(.ia, .ib, .ic, .id, .ie, .if, .ih, .ii, .ij), Vec9(.ja, .jb, .jc, .jd, .je, .jf, .jh, .ji, .jj))) _
'                     - .ah * Matrix9_detX(Mat9(Vec9(.ba, .bb, .bc, .bd, .be, .bf, .bg, .bi, .bj), Vec9(.ca, .cb, .cc, .cd, .ce, .cf, .cg, .ci, .cj), Vec9(.da, .db, .dc, .dd, .de, .df, .dg, .di, .dj), Vec9(.ea, .eb, .ec, .ed, .ee, .ef, .eg, .ei, .ej), Vec9(.fa, .fb, .fc, .fd, .fe, .ff, .fg, .fi, .fj), Vec9(.ga, .gb, .gc, .gd, .ge, .gf, .gg, .gi, .gj), Vec9(.ha, .hb, .HC, .hd, .he, .hf, .hg, .hi, .hj), Vec9(.ia, .ib, .ic, .id, .ie, .if, .ig, .ii, .ij), Vec9(.ja, .jb, .jc, .jd, .je, .jf, .jg, .ji, .jj))) _
'                     + .ai * Matrix9_detX(Mat9(Vec9(.ba, .bb, .bc, .bd, .be, .bf, .bg, .bh, .bj), Vec9(.ca, .cb, .cc, .cd, .ce, .cf, .cg, .ch, .cj), Vec9(.da, .db, .dc, .dd, .de, .df, .dg, .dh, .dj), Vec9(.ea, .eb, .ec, .ed, .ee, .ef, .eg, .eh, .ej), Vec9(.fa, .fb, .fc, .fd, .fe, .ff, .fg, .fh, .fj), Vec9(.ga, .gb, .gc, .gd, .ge, .gf, .gg, .gh, .gj), Vec9(.ha, .hb, .HC, .hd, .he, .hf, .hg, .hh, .hj), Vec9(.ia, .ib, .ic, .id, .ie, .if, .ig, .ih, .ij), Vec9(.ja, .jb, .jc, .jd, .je, .jf, .jg, .jh, .jj))) _
'                     - .aj * Matrix9_detX(Mat9(Vec9(.ba, .bb, .bc, .bd, .be, .bf, .bg, .bh, .bi), Vec9(.ca, .cb, .cc, .cd, .ce, .cf, .cg, .ch, .ci), Vec9(.da, .db, .dc, .dd, .de, .df, .dg, .dh, .di), Vec9(.ea, .eb, .ec, .ed, .ee, .ef, .eg, .eh, .ei), Vec9(.fa, .fb, .fc, .fd, .fe, .ff, .fg, .fh, .fi), Vec9(.ga, .gb, .gc, .gd, .ge, .gf, .gg, .gh, .gi), Vec9(.ha, .hb, .HC, .hd, .he, .hf, .hg, .hh, .hi), Vec9(.ia, .ib, .ic, .id, .ie, .if, .ig, .ih, .ii), Vec9(.ja, .jb, .jc, .jd, .je, .jf, .jg, .jh, .ji)))
'    End With
'End Function
'
'Public Function Matrix9_detXX(m As Matrix9) As Double
'    'Berechnet die Determinante einer 9x9-Matrix
'    'Entwicklung nach der letzten Zeile
'    Dim md As Matrix8
'    With md
'        .aa = m.ab: .ab = m.ac: .ac = m.ad: .ad = m.ae: .ae = m.af: .af = m.ag: .ag = m.ah: .ah = m.ai
'        .ba = m.bb: .bb = m.bc: .bc = m.bd: .bd = m.be: .be = m.bf: .bf = m.bg: .bg = m.bh: .bh = m.bi
'        .ca = m.cb: .cb = m.cc: .cc = m.cd: .cd = m.ce: .ce = m.cf: .cf = m.cg: .cg = m.ch: .ch = m.ci
'        .da = m.db: .db = m.dc: .dc = m.dd: .dd = m.de: .de = m.df: .df = m.dg: .dg = m.dh: .dh = m.di
'        .ea = m.eb: .eb = m.ec: .ec = m.ed: .ed = m.ee: .ee = m.ef: .ef = m.eg: .eg = m.eh: .eh = m.ei
'        .fa = m.fb: .fb = m.fc: .fc = m.fd: .fd = m.fe: .fe = m.ff: .ff = m.fg: .fg = m.fh: .fh = m.fi
'        .ga = m.gb: .gb = m.gc: .gc = m.gd: .gd = m.ge: .ge = m.gf: .gf = m.gg: .gg = m.gh: .gh = m.gi
'        .ha = m.hb: .hb = m.HC: .HC = m.hd: .hd = m.he: .he = m.hf: .hf = m.hg: .hg = m.hh: .hh = m.hi
'    End With
'    Dim det_a As Double: det_a = Matrix8_det(md)
'    With md
'        .aa = m.aa: .ab = m.ac: .ac = m.ad: .ad = m.ae: .ae = m.af: .af = m.ag: .ag = m.ah: .ah = m.ai
'        .ba = m.ba: .bb = m.bc: .bc = m.bd: .bd = m.be: .be = m.bf: .bf = m.bg: .bg = m.bh: .bh = m.bi
'        .ca = m.ca: .cb = m.cc: .cc = m.cd: .cd = m.ce: .ce = m.cf: .cf = m.cg: .cg = m.ch: .ch = m.ci
'        .da = m.da: .db = m.dc: .dc = m.dd: .dd = m.de: .de = m.df: .df = m.dg: .dg = m.dh: .dh = m.di
'        .ea = m.ea: .eb = m.ec: .ec = m.ed: .ed = m.ee: .ee = m.ef: .ef = m.eg: .eg = m.eh: .eh = m.ei
'        .fa = m.fa: .fb = m.fc: .fc = m.fd: .fd = m.fe: .fe = m.ff: .ff = m.fg: .fg = m.fh: .fh = m.fi
'        .ga = m.ga: .gb = m.gc: .gc = m.gd: .gd = m.ge: .ge = m.gf: .gf = m.gg: .gg = m.gh: .gh = m.gi
'        .ha = m.ha: .hb = m.HC: .HC = m.hd: .hd = m.he: .he = m.hf: .hf = m.hg: .hg = m.hh: .hh = m.hi
'    End With
'    Dim det_b As Double: det_b = Matrix8_det(md)
'    With md
'        .aa = m.aa: .ab = m.ab: .ac = m.ad: .ad = m.ae: .ae = m.af: .af = m.ag: .ag = m.ah: .ah = m.ai
'        .ba = m.ba: .bb = m.bb: .bc = m.bd: .bd = m.be: .be = m.bf: .bf = m.bg: .bg = m.bh: .bh = m.bi
'        .ca = m.ca: .cb = m.cb: .cc = m.cd: .cd = m.ce: .ce = m.cf: .cf = m.cg: .cg = m.ch: .ch = m.ci
'        .da = m.da: .db = m.db: .dc = m.dd: .dd = m.de: .de = m.df: .df = m.dg: .dg = m.dh: .dh = m.di
'        .ea = m.ea: .eb = m.eb: .ec = m.ed: .ed = m.ee: .ee = m.ef: .ef = m.eg: .eg = m.eh: .eh = m.ei
'        .fa = m.fa: .fb = m.fb: .fc = m.fd: .fd = m.fe: .fe = m.ff: .ff = m.fg: .fg = m.fh: .fh = m.fi
'        .ga = m.ga: .gb = m.gb: .gc = m.gd: .gd = m.ge: .ge = m.gf: .gf = m.gg: .gg = m.gh: .gh = m.gi
'        .ha = m.ha: .hb = m.hb: .HC = m.hd: .hd = m.he: .he = m.hf: .hf = m.hg: .hg = m.hh: .hh = m.hi
'    End With
'    Dim det_c As Double: det_c = Matrix8_det(md)
'    With md
'        .aa = m.aa: .ab = m.ab: .ac = m.ac: .ad = m.ae: .ae = m.af: .af = m.ag: .ag = m.ah: .ah = m.ai
'        .ba = m.ba: .bb = m.bb: .bc = m.bc: .bd = m.be: .be = m.bf: .bf = m.bg: .bg = m.bh: .bh = m.bi
'        .ca = m.ca: .cb = m.cb: .cc = m.cc: .cd = m.ce: .ce = m.cf: .cf = m.cg: .cg = m.ch: .ch = m.ci
'        .da = m.da: .db = m.db: .dc = m.dc: .dd = m.de: .de = m.df: .df = m.dg: .dg = m.dh: .dh = m.di
'        .ea = m.ea: .eb = m.eb: .ec = m.ec: .ed = m.ee: .ee = m.ef: .ef = m.eg: .eg = m.eh: .eh = m.ei
'        .fa = m.fa: .fb = m.fb: .fc = m.fc: .fd = m.fe: .fe = m.ff: .ff = m.fg: .fg = m.fh: .fh = m.fi
'        .ga = m.ga: .gb = m.gb: .gc = m.gc: .gd = m.ge: .ge = m.gf: .gf = m.gg: .gg = m.gh: .gh = m.gi
'        .ha = m.ha: .hb = m.hb: .HC = m.HC: .hd = m.he: .he = m.hf: .hf = m.hg: .hg = m.hh: .hh = m.hi
'    End With
'    Dim det_d As Double: det_d = Matrix8_det(md)
'    With md
'        .aa = m.aa: .ab = m.ab: .ac = m.ac: .ad = m.ad: .ae = m.af: .af = m.ag: .ag = m.ah: .ah = m.ai
'        .ba = m.ba: .bb = m.bb: .bc = m.bc: .bd = m.bd: .be = m.bf: .bf = m.bg: .bg = m.bh: .bh = m.bi
'        .ca = m.ca: .cb = m.cb: .cc = m.cc: .cd = m.cd: .ce = m.cf: .cf = m.cg: .cg = m.ch: .ch = m.ci
'        .da = m.da: .db = m.db: .dc = m.dc: .dd = m.dd: .de = m.df: .df = m.dg: .dg = m.dh: .dh = m.di
'        .ea = m.ea: .eb = m.eb: .ec = m.ec: .ed = m.ed: .ee = m.ef: .ef = m.eg: .eg = m.eh: .eh = m.ei
'        .fa = m.fa: .fb = m.fb: .fc = m.fc: .fd = m.fd: .fe = m.ff: .ff = m.fg: .fg = m.fh: .fh = m.fi
'        .ga = m.ga: .gb = m.gb: .gc = m.gc: .gd = m.gd: .ge = m.gf: .gf = m.gg: .gg = m.gh: .gh = m.gi
'        .ha = m.ha: .hb = m.hb: .HC = m.HC: .hd = m.hd: .he = m.hf: .hf = m.hg: .hg = m.hh: .hh = m.hi
'    End With
'    Dim det_e As Double: det_e = Matrix8_det(md)
'    With md
'        .aa = m.aa: .ab = m.ab: .ac = m.ac: .ad = m.ad: .ae = m.ae: .af = m.ag: .ag = m.ah: .ah = m.ai
'        .ba = m.ba: .bb = m.bb: .bc = m.bc: .bd = m.bd: .be = m.be: .bf = m.bg: .bg = m.bh: .bh = m.bi
'        .ca = m.ca: .cb = m.cb: .cc = m.cc: .cd = m.cd: .ce = m.ce: .cf = m.cg: .cg = m.ch: .ch = m.ci
'        .da = m.da: .db = m.db: .dc = m.dc: .dd = m.dd: .de = m.de: .df = m.dg: .dg = m.dh: .dh = m.di
'        .ea = m.ea: .eb = m.eb: .ec = m.ec: .ed = m.ed: .ee = m.ee: .ef = m.eg: .eg = m.eh: .eh = m.ei
'        .fa = m.fa: .fb = m.fb: .fc = m.fc: .fd = m.fd: .fe = m.fe: .ff = m.fg: .fg = m.fh: .fh = m.fi
'        .ga = m.ga: .gb = m.gb: .gc = m.gc: .gd = m.gd: .ge = m.ge: .gf = m.gg: .gg = m.gh: .gh = m.gi
'        .ha = m.ha: .hb = m.hb: .HC = m.HC: .hd = m.hd: .he = m.he: .hf = m.hg: .hg = m.hh: .hh = m.hi
'    End With
'    Dim det_f As Double: det_f = Matrix8_det(md)
'    With md
'        .aa = m.aa: .ab = m.ab: .ac = m.ac: .ad = m.ad: .ae = m.ae: .af = m.af: .ag = m.ah: .ah = m.ai
'        .ba = m.ba: .bb = m.bb: .bc = m.bc: .bd = m.bd: .be = m.be: .bf = m.bf: .bg = m.bh: .bh = m.bi
'        .ca = m.ca: .cb = m.cb: .cc = m.cc: .cd = m.cd: .ce = m.ce: .cf = m.cf: .cg = m.ch: .ch = m.ci
'        .da = m.da: .db = m.db: .dc = m.dc: .dd = m.dd: .de = m.de: .df = m.df: .dg = m.dh: .dh = m.di
'        .ea = m.ea: .eb = m.eb: .ec = m.ec: .ed = m.ed: .ee = m.ee: .ef = m.ef: .eg = m.eh: .eh = m.ei
'        .fa = m.fa: .fb = m.fb: .fc = m.fc: .fd = m.fd: .fe = m.fe: .ff = m.ff: .fg = m.fh: .fh = m.fi
'        .ga = m.ga: .gb = m.gb: .gc = m.gc: .gd = m.gd: .ge = m.ge: .gf = m.gf: .gg = m.gh: .gh = m.gi
'        .ha = m.ha: .hb = m.hb: .HC = m.HC: .hd = m.hd: .he = m.he: .hf = m.hf: .hg = m.hh: .hh = m.hi
'    End With
'    Dim det_g As Double: det_g = Matrix8_det(md)
'    With md
'        .aa = m.aa: .ab = m.ab: .ac = m.ac: .ad = m.ad: .ae = m.ae: .af = m.af: .ag = m.ag: .ah = m.ai
'        .ba = m.ba: .bb = m.bb: .bc = m.bc: .bd = m.bd: .be = m.be: .bf = m.bf: .bg = m.bg: .bh = m.bi
'        .ca = m.ca: .cb = m.cb: .cc = m.cc: .cd = m.cd: .ce = m.ce: .cf = m.cf: .cg = m.cg: .ch = m.ci
'        .da = m.da: .db = m.db: .dc = m.dc: .dd = m.dd: .de = m.de: .df = m.df: .dg = m.dg: .dh = m.di
'        .ea = m.ea: .eb = m.eb: .ec = m.ec: .ed = m.ed: .ee = m.ee: .ef = m.ef: .eg = m.eg: .eh = m.ei
'        .fa = m.fa: .fb = m.fb: .fc = m.fc: .fd = m.fd: .fe = m.fe: .ff = m.ff: .fg = m.fg: .fh = m.fi
'        .ga = m.ga: .gb = m.gb: .gc = m.gc: .gd = m.gd: .ge = m.ge: .gf = m.gf: .gg = m.gg: .gh = m.gi
'        .ha = m.ha: .hb = m.hb: .HC = m.HC: .hd = m.hd: .he = m.he: .hf = m.hf: .hg = m.hg: .hh = m.hi
'    End With
'    Dim det_h As Double: det_h = Matrix8_det(md)
'    With md
'        .aa = m.aa: .ab = m.ab: .ac = m.ac: .ad = m.ad: .ae = m.ae: .af = m.af: .ag = m.ag: .ah = m.ah
'        .ba = m.ba: .bb = m.bb: .bc = m.bc: .bd = m.bd: .be = m.be: .bf = m.bf: .bg = m.bg: .bh = m.bh
'        .ca = m.ca: .cb = m.cb: .cc = m.cc: .cd = m.cd: .ce = m.ce: .cf = m.cf: .cg = m.cg: .ch = m.ch
'        .da = m.da: .db = m.db: .dc = m.dc: .dd = m.dd: .de = m.de: .df = m.df: .dg = m.dg: .dh = m.dh
'        .ea = m.ea: .eb = m.eb: .ec = m.ec: .ed = m.ed: .ee = m.ee: .ef = m.ef: .eg = m.eg: .eh = m.eh
'        .fa = m.fa: .fb = m.fb: .fc = m.fc: .fd = m.fd: .fe = m.fe: .ff = m.ff: .fg = m.fg: .fh = m.fh
'        .ga = m.ga: .gb = m.gb: .gc = m.gc: .gd = m.gd: .ge = m.ge: .gf = m.gf: .gg = m.gg: .gh = m.gh
'        .ha = m.ha: .hb = m.hb: .HC = m.HC: .hd = m.hd: .he = m.he: .hf = m.hf: .hg = m.hg: .hh = m.hh
'    End With
'    Dim det_i As Double: det_i = Matrix8_det(md)
'         '+   -   +   -   +   -   +   -   +
'    With m
'        Matrix9_detXX = .ia * det_a _
'                      - .ib * det_b _
'                      + .ic * det_c _
'                      - .id * det_d _
'                      + .ie * det_e _
'                      - .if * det_f _
'                      + .ig * det_g _
'                      - .ih * det_h _
'                      + .ii * det_i
'    End With
'End Function
'
'
