VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17055
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   17055
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnRunSomeTests 
      Caption         =   "Run some tests"
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Random Matrices"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton BtnCalc 
      Caption         =   "="
      Height          =   375
      Left            =   10800
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   11400
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   3
      Top             =   600
      Width           =   5655
   End
   Begin VB.ComboBox CmbOps 
      Height          =   315
      Left            =   5160
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   5760
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   1
      Top             =   600
      Width           =   5535
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   0
      Top             =   600
      Width           =   5655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command4_Click()

End Sub

'zum Überprüfen der Berechnung siehe z.B.:
'http://matrizen-rechner.de/
'https://matrixcalc.org/de/
'https://rechneronline.de/lineare-algebra/matrizen.php
'https://rechneronline.de/lineare-algebra/gleichungssysteme.php


'Matrix-Tipp
'einfache Matrizenoperationen für kleine Matrizen bis 6x6
'Addieren (add), Subtrahieren (sub), Multiplizieren (mul), Skalar-Multiplikation (smul), Transponierte (tra), Determinante (det), Adjunkte (adj), Inverse (inv)
'add, sub, mul, tra, det, adj, inv
Private Sub Form_Load()
    Me.Caption = "Matrizen - bis 6x6"
    With CmbOps
        .Font = "Courier New": .FontSize = 10
        .AddItem "  +  ": .AddItem "  -  ": .AddItem "  *  ": .AddItem " tra ": .AddItem " det ": .AddItem " adj ": .AddItem " inv ": .AddItem "  =  "
        .AddItem "A*x=b" 'noch nicht implementiert
        .ListIndex = 0
    End With
    Randomize
    Text1.Text = GetRandomMat
    Text2.Text = Text1.Text
End Sub

Private Sub BtnCalc_Click()

    Dim A2 As Matrix2, b2 As Matrix2, C2 As Matrix2
    Dim A3 As Matrix3, b3 As Matrix3, C3 As Matrix3
    Dim a4 As Matrix4, b4 As Matrix4, C4 As Matrix4
    Dim A5 As Matrix5, b5 As Matrix5, C5 As Matrix5
    Dim A6 As Matrix6, b6 As Matrix6, C6 As Matrix6
    'die maximalen Größen der Matrizen herausfinden
    Dim A_mRows As Long, A_nCols As Long
    Dim B_mRows As Long, B_nCols As Long
    GetRowsCols Text1.Text, A_mRows, A_nCols
    GetRowsCols Text2.Text, B_mRows, B_nCols
    Dim maxrc As Long: maxrc = Max(A_mRows, Max(B_mRows, Max(A_nCols, B_nCols)))
    
    Select Case maxrc
    Case 2: A2 = Matrix2_Parse(Text1.Text): b2 = Matrix2_Parse(Text2.Text)
    Case 3: A3 = Matrix3_Parse(Text1.Text): b3 = Matrix3_Parse(Text2.Text)
    Case 4: a4 = Matrix4_Parse(Text1.Text): b4 = Matrix4_Parse(Text2.Text)
    Case 5: A5 = Matrix5_Parse(Text1.Text): b5 = Matrix5_Parse(Text2.Text)
    Case 6: A6 = Matrix6_Parse(Text1.Text): b6 = Matrix6_Parse(Text2.Text)
    End Select
    'Dim C6 As Matrix6
    Dim scalar As Double
    Dim op As String: op = CmbOps.List(CmbOps.ListIndex)
    If op = "  *  " And (A_mRows = 1 And A_nCols = 1) Or (B_mRows = 1 And B_nCols = 1) Then
        op = "smul"
        scalar = Val(Text2.Text)
    End If
    Select Case op
    Case "  +  "
        Select Case maxrc
        Case 2: C2 = Matrix2_add(A2, b2):  LSet C6 = C2
        Case 3: C3 = Matrix3_add(A3, b3):  LSet C6 = C3
        Case 4: C4 = Matrix4_add(a4, b4):  LSet C6 = C4
        Case 5: C5 = Matrix5_add(A5, b5):  LSet C6 = C5
        Case 6: C6 = Matrix6_add(A6, b6)
        End Select
    Case "  -  "
        Select Case maxrc
        Case 2: C2 = Matrix2_sub(A2, b2): LSet C6 = C2
        Case 3: C3 = Matrix3_sub(A3, b3): LSet C6 = C3
        Case 4: C4 = Matrix4_sub(a4, b4): LSet C6 = C4
        Case 5: C5 = Matrix5_sub(A5, b5): LSet C6 = C5
        Case 6: C6 = Matrix6_sub(A6, b6)
        End Select
    Case "smul"
        Select Case maxrc
        Case 2: C2 = Matrix2_smul(A2, scalar): LSet C6 = C2
        Case 3: C3 = Matrix3_smul(A3, scalar): LSet C6 = C3
        Case 4: C4 = Matrix4_smul(a4, scalar): LSet C6 = C4
        Case 5: C5 = Matrix5_smul(A5, scalar): LSet C6 = C5
        Case 6: C6 = Matrix6_smul(A6, scalar)
        End Select
    Case "  *  "
        Select Case maxrc
        Case 2: C2 = Matrix2_mul(A2, b2): LSet C6 = C2
        Case 3: C3 = Matrix3_mul(A3, b3): LSet C6 = C3
        Case 4: C4 = Matrix4_mul(a4, b4): LSet C6 = C4
        Case 5: C5 = Matrix5_mul(A5, b5): LSet C6 = C5
        Case 6: C6 = Matrix6_mul(A6, b6)
        End Select
    Case " tra "
        Select Case maxrc
        Case 2: C2 = Matrix2_tra(A2): LSet C6 = C2
        Case 3: C3 = Matrix3_tra(A3): LSet C6 = C3
        Case 4: C4 = Matrix4_tra(a4): LSet C6 = C4
        Case 5: C5 = Matrix5_tra(A5): LSet C6 = C5
        Case 6: C6 = Matrix6_tra(A6)
        End Select
    Case "  =  "
        Dim b As Boolean
        Select Case maxrc
        Case 2: b = Matrix2_IsEqual(A2, b2)
        Case 3: b = Matrix3_IsEqual(A3, b3)
        Case 4: b = Matrix4_IsEqual(a4, b4)
        Case 5: b = Matrix5_IsEqual(A5, b5)
        Case 6: b = Matrix6_IsEqual(A6, b6)
        End Select
    Case " det "
        Dim d As Double, d2 As Double
        Select Case maxrc
        Case 2: d = Matrix2_det(A2)
        Case 3: d = Matrix3_det(A3) ': d2 = Matrix3_det2(A3)
        Case 4: d = Matrix4_det(a4)
        Case 5: d = Matrix5_det(A5)
        Case 6: d = Matrix6_det(A6)
        End Select
    Case " adj "
        Select Case maxrc
        Case 2: C2 = Matrix2_Adj(A2):  LSet C6 = C2
        Case 3: C3 = Matrix3_Adj(A3):  LSet C6 = C3
        Case 4: C4 = Matrix4_Adj(a4):  LSet C6 = C4
        Case 5: C5 = Matrix5_Adj(A5):  LSet C6 = C5
        Case 6: C6 = Matrix6_Adj(A6)
        End Select
    Case " inv "
        Select Case maxrc
        Case 2: C2 = Matrix2_inv(A2):  LSet C6 = C2
        Case 3: C3 = Matrix3_inv(A3):  LSet C6 = C3
        Case 4: C4 = Matrix4_inv(a4):  LSet C6 = C4
        Case 5: C5 = Matrix5_inv(A5):  LSet C6 = C5
        Case 6: C6 = Matrix6_inv(A6)
        End Select
    Case "A*x=b"
        Dim vb2 As Vector2, vb3 As Vector3, vb4 As Vector4, vb5 As Vector5, vb6 As Vector6
        Dim vx2 As Vector2, vx3 As Vector3, vx4 As Vector4, vx5 As Vector5, vx6 As Vector6
        'z.B.: A[-2 -2 | -1  1] * x[-2 | 0] = b [ 4 | 2 ];
        'z.B.: A[ 5  6 |  3  4] * x[-13|12] = b [ 7 | 8 ];
        Select Case maxrc
        Case 2: vb2 = Mat2_Col(Matrix2_Parse(Text3.Text), 0) ': Debug.Print Vector2_ToStr(vb2)
                vx2 = Matrix2_solve(A2, vb2):                 'Debug.Print Vector2_ToStr(vx2) '  LSet C6 = C2
        Case 3: vb3 = Mat3_Col(Matrix3_Parse(Text3.Text), 0) ': Debug.Print Vector3_ToStr(vb3)
                vx3 = Matrix3_solve(A3, vb3):                 'Debug.Print Vector3_ToStr(vx3) '  LSet C6 = C2
        Case 4: vb4 = Mat4_Col(Matrix4_Parse(Text3.Text), 0) ': Debug.Print Vector3_ToStr(vb3)
                vx4 = Matrix4_solve(a4, vb4):                 'Debug.Print Vector3_ToStr(vx3) '  LSet C6 = C2
        Case 5: vb5 = Mat5_Col(Matrix5_Parse(Text3.Text), 0) ': Debug.Print Vector3_ToStr(vb3)
                vx5 = Matrix5_solve(A5, vb5):                 'Debug.Print Vector3_ToStr(vx3) '  LSet C6 = C2
        Case 6: vb6 = Mat6_Col(Matrix6_Parse(Text3.Text), 0) ': Debug.Print Vector3_ToStr(vb3)
                vx6 = Matrix6_solve(A6, vb6):                 'Debug.Print Vector3_ToStr(vx3) '  LSet C6 = C2
        End Select
    End Select
    
    Select Case maxrc
    Case 2: Text1.Text = Matrix2_ToStr(A2): Text2.Text = Matrix2_ToStr(b2)
    Case 3: Text1.Text = Matrix3_ToStr(A3): Text2.Text = Matrix3_ToStr(b3)
    Case 4: Text1.Text = Matrix4_ToStr(a4): Text2.Text = Matrix4_ToStr(b4)
    Case 5: Text1.Text = Matrix5_ToStr(A5): Text2.Text = Matrix5_ToStr(b5)
    Case 6: Text1.Text = Matrix6_ToStr(A6): Text2.Text = Matrix6_ToStr(b6)
    End Select
    'If Not op = "smul" Then
    'End If
    Dim t As String
    If op = "  =  " Then
        t = b
        Text3.Text = t
    ElseIf op = " det " Then
        t = Str(d) '& vbCrLf & Str(d2)
        Text3.Text = t
    ElseIf op = "A*x=b" Then
        Select Case maxrc
        Case 2: t = Vector2_ToStr(vx2, False)
        Case 3: t = Vector3_ToStr(vx3, False)
        Case 4: t = Vector4_ToStr(vx4, False)
        Case 5: t = Vector5_ToStr(vx5, False)
        Case 6: t = Vector6_ToStr(vx6, False)
        End Select
        Text2.Text = t
    Else
        t = Matrix_ToStr(VarPtr(C6), maxrc, maxrc)
        Text3.Text = t
    End If



End Sub

Sub GetRowsCols(t As String, ByRef rows_out As Long, ByRef cols_out As Long)
    rows_out = GetRows(t):    cols_out = GetCols(t)
End Sub
Function GetRows(t As String) As Long
    Dim s As String: s = DeleteMultiWS(t)
    Dim sa() As String: sa = Split(s, vbCrLf)
    Dim i As Long
    GetRows = UBound(sa) + 1
    For i = UBound(sa) To 0 Step -1
        If Len(sa(i)) = 0 Then GetRows = GetRows - 1 Else Exit For
    Next
End Function
Function GetCols(t As String) As Long
    Dim s As String: s = DeleteMultiWS(t)
    Dim sa() As String: sa = Split(s, vbCrLf)
    Dim i As Long
    For i = 0 To UBound(sa)
        GetCols = Max(GetCols, UBound(Split(Trim(sa(i)), " ")) + 1)
    Next
End Function

Private Sub Command2_Click()
    Text1.Text = GetRandomMat
    Text2.Text = GetRandomMat 'Text1.Text
    Text3.Text = ""
    BtnCalc_Click
End Sub

Function GetRandomMat() As String
    Dim r As Long: r = Rnd * 10 + 2
    Select Case r
    Case 2: GetRandomMat = GetMat2
    Case 3: GetRandomMat = GetMat3
    Case 4: GetRandomMat = GetMat4
    Case 5: GetRandomMat = GetMat5
    Case 6: GetRandomMat = GetMat6
    Case Else
            GetRandomMat = GetRndMat
    End Select
End Function

Function GetRndMat() As String
    Dim size As Long: size = Rnd * 4 + 2
    Dim s As String
    'ReDim a(0 To size - 1, 0 To size - 1)
    Dim i As Long, j As Long
    For i = 1 To size 'UBound(a, 1)
        For j = 1 To size 'UBound(a, 2)
            'a(i, j) =  Rnd * 200 * Sgn(Rnd - 0.5)
            s = s & Replace(Format(Rnd * 200 * Sgn(Rnd - 0.5), "0.00"), ",", ".") & " "
        Next
        s = s & vbCrLf
    Next
    GetRndMat = s
End Function
Function GetMat2() As String
    GetMat2 = "11 12" & vbCrLf & _
              "21 22"
End Function
Function GetMat3() As String
    GetMat3 = "11 12 13" & vbCrLf & _
              "21 22 23" & vbCrLf & _
              "31 32 33"
End Function
Function GetMat4() As String
    GetMat4 = "11 12 13 14" & vbCrLf & _
              "21 22 23 24" & vbCrLf & _
              "31 32 33 34" & vbCrLf & _
              "41 42 43 44"
End Function
Function GetMat5() As String
    GetMat5 = "11 12 13 14 15" & vbCrLf & _
              "21 22 23 24 25" & vbCrLf & _
              "31 32 33 34 35" & vbCrLf & _
              "41 42 43 44 45" & vbCrLf & _
              "51 52 53 54 55"
End Function
Function GetMat6() As String
    GetMat6 = "11 12 13 14 15 16" & vbCrLf & _
              "21 22 23 24 25 26" & vbCrLf & _
              "31 32 33 34 35 36" & vbCrLf & _
              "41 42 43 44 45 46" & vbCrLf & _
              "51 52 53 54 55 56" & vbCrLf & _
              "61 62 63 64 65 66"
End Function

Private Sub Form_Resize()
    Dim brdr: brdr = 2 * Screen.TwipsPerPixelX
    Dim L, t, W, H
    Dim L2
    L = brdr: t = Text1.Top
    W = (Me.ScaleWidth - 4 * brdr) / 3
    H = Me.ScaleHeight - t - brdr
    If W > 0 And H > 0 Then
        Text1.Move L, t, W, H
        L = L + Text1.Width + brdr
        L2 = L - (brdr + CmbOps.Width) / 2
        CmbOps.Move L2, t - CmbOps.Height
        Text2.Move L, t, W, H
        L = L + Text2.Width + brdr
        L2 = L - (brdr + BtnCalc.Width) / 2
        BtnCalc.Move L2, t - BtnCalc.Height
        Text3.Move L, t, W, H
    End If
End Sub

Private Sub BtnRunSomeTests_Click()
    Dim a4 As Matrix4
    Dim t As Single: t = Timer
    If MsgBox("Starting some tests, this may take a while." & vbCrLf & "All matrices in the view will be cleared. Proceed anyway?", vbOKCancel) = vbCancel Then Exit Sub
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    
    Test_ToStrParseDet
    Test_umat3
    Test_umat4
    Test_umat5
    Test_umat6
    
    Test_PropGetLetRowCol_2
    Test_PropGetLetRowCol_3
    Test_PropGetLetRowCol_4
    Test_PropGetLetRowCol_5
    Test_PropGetLetRowCol_6
    
    Test_MatE_6
    'All tests passed!
    t = Timer - t
    BtnRunSomeTests.Caption = t
    CmbOps.SetFocus
End Sub

Sub Test_ToStrParseDet()
    Dim a4 As Matrix4
    MsgBox "Create a 4x4-Matrix"
    a4 = Mat4(11, 12, 13, -14, _
              21, 22, 23, 24, _
              31, 32, 33, 34, _
              41, 42, 43, 44)
    Text1.Text = Matrix4_ToStr(a4)
    Dim t As String: t = Text1.Text
    MsgBox "Parse the 4x4-Matrix"
    a4 = Matrix4_Parse(t)
    Text2.Text = Matrix4_ToStr(a4)
    Dim d As Double: d = Matrix4_det(a4)
    MsgBox "The determinant of the matrix is: " & d
    Text2.Text = d
End Sub
Sub Test_umat3()
    MsgBox "Create a 3x3-matrix"
    Dim m As Matrix3: m = Mat3(11, 12, 13, _
                               21, 22, 23, _
                               31, 32, 33)
    Text1.Text = Matrix3_ToStr(m)
    Dim i As Long, j As Long
    For i = 0 To 2: For j = 0 To 2
        MsgBox "2x2-submatrix " & i & " " & j
        Dim um As Matrix2: um = Matrix3_umat(m, i, j)
        Text2.Text = Matrix2_ToStr(um)
    Next: Next
End Sub
Sub Test_umat4()
    MsgBox "Create a 4x4-matrix"
    Dim m As Matrix4: m = Mat4(11, 12, 13, 14, _
                               21, 22, 23, 24, _
                               31, 32, 33, 34, _
                               41, 42, 43, 44)
    Text1.Text = Matrix4_ToStr(m)
    Dim i As Long, j As Long
    For i = 0 To 3: For j = 0 To 3
        MsgBox "3x3-submatrix " & i & " " & j
        Dim um As Matrix3: um = Matrix4_umat(m, i, j)
        Text2.Text = Matrix3_ToStr(um)
    Next: Next
End Sub
Sub Test_umat5()
    MsgBox "Create a 5x5-matrix"
    Dim m As Matrix5: m = Mat5(11, 12, 13, 14, 15, _
                               21, 22, 23, 24, 25, _
                               31, 32, 33, 34, 35, _
                               41, 42, 43, 44, 45, _
                               51, 52, 53, 54, 55)
    Text1.Text = Matrix5_ToStr(m)
    Dim i As Long, j As Long
    For i = 0 To 4: For j = 0 To 4
        MsgBox "4x4-submatrix " & i & " " & j
        Dim um As Matrix4: um = Matrix5_umat(m, i, j)
        Text2.Text = Matrix4_ToStr(um)
    Next: Next
End Sub
Sub Test_umat6()
    MsgBox "Create a 6x6-matrix"
    Dim m As Matrix6: m = Mat6(11, 12, 13, 14, 15, 16, _
                               21, 22, 23, 24, 25, 26, _
                               31, 32, 33, 34, 35, 36, _
                               41, 42, 43, 44, 45, 46, _
                               51, 52, 53, 54, 55, 56, _
                               61, 62, 63, 64, 65, 66)
    Text1.Text = Matrix6_ToStr(m)
    Dim i As Long, j As Long
    For i = 0 To 5: For j = 0 To 5
        MsgBox "5x5-submatrix " & i & " " & j
        Dim um As Matrix5: um = Matrix6_umat(m, i, j)
        Text2.Text = Matrix5_ToStr(um)
    Next: Next
End Sub

'Test Property Get/Let Row / Col
Sub Test_PropGetLetRowCol_2()
    Dim m As Matrix2
    Dim v As Vector2
    'OK vielleicht machen wirs so:
    
    ' * eine leere mit Nullen besetzte 2x2-Matrix erzeugen und anzeigen in Text1
    MsgBox "Create an empty 2x2-matrix m": Text1.Text = Matrix2_ToStr(m)
    
    ' * den 1. Zeilen-Vektor erzeugen und anzeigen in Text2
    MsgBox "Create a 2x1-vector v": v = Vec2(11, 12): Text2.Text = Vector2_ToStr(v)
    ' * den Vektor in die Matrix an die 0.-Zeile schreiben und die Matrix in Text1 anzeigen
    MsgBox "Write the vector to row 0 of m (Property Let Row)": Mat2_Row(m, 0) = v: Text1.Text = Matrix2_ToStr(m): Text2.Text = ""
        
    ' * den 2. Zeilen-Vektor erzeugen und anzeigen in Text2
    MsgBox "Create a 2x1-vector v": v = Vec2(21, 22): Text2.Text = Vector2_ToStr(v)
    ' * den Vektor in die Matrix an die erste Zeile schreiben und die Matrix in Text1 anzeigen
    
    MsgBox "Write the vector to row 1 of m (Property Let Row)": Mat2_Row(m, 1) = v: Text1.Text = Matrix2_ToStr(m): Text2.Text = ""
    
    
    
    ' * die 2x2-Matrix löschen und anzeigen in Text1
    MsgBox "Clear the 2x2-matrix m": m = Mat2_Clear:   Text1.Text = Matrix2_ToStr(m)
    
    ' * den 1. Spalten-Vektor erzeugen und anzeigen in Text2
    MsgBox "Create a 2x1-vector v":  v = Vec2(11, 21): Text2.Text = Vector2_ToStr(v)
    ' * den Vektor in die Matrix an die 0.-Spalte schreiben und die Matrix in Text1 anzeigen
    MsgBox "Write the vector to col 0 of m (Property Let Col)": Mat2_Col(m, 0) = v: Text1.Text = Matrix2_ToStr(m): Text2.Text = ""
    
    ' * den 2. Spalten-Vektor erzeugen und anzeigen in Text2
    MsgBox "Create a 2x1-vector v": v = Vec2(12, 22): Text2.Text = Vector2_ToStr(v)
    ' * den Vektor in die Matrix an die erste Zeile schreiben und die Matrix in Text1 anzeigen
    MsgBox "Write the vector to col 1 of m (Property Let Col)": Mat2_Col(m, 1) = v: Text1.Text = Matrix2_ToStr(m): Text2.Text = ""
    ' * u.s.w.
End Sub
Sub Test_PropGetLetRowCol_3()
    Dim m As Matrix3
    Dim v As Vector3

    MsgBox "Create an empty 3x3-matrix m": Text1.Text = Matrix3_ToStr(m)
    MsgBox "Create a 3x1-vector":  v = Vec3(11, 12, 13): Text2.Text = Vector3_ToStr(v)
    MsgBox "Write the vector to row 0 of m (Property Let Row)": Mat3_Row(m, 0) = v: Text1.Text = Matrix3_ToStr(m): Text2.Text = ""
        
    MsgBox "Create a 3x1-vector v":  v = Vec3(21, 22, 23): Text2.Text = Vector3_ToStr(v)
    MsgBox "Write the vector to row 1 of m (Property Let Row)": Mat3_Row(m, 1) = v: Text1.Text = Matrix3_ToStr(m): Text2.Text = ""
        
    MsgBox "Create a 3x1-vector v":  v = Vec3(31, 32, 33): Text2.Text = Vector3_ToStr(v)
    MsgBox "Write the vector to row 2 of m (Property Let Row)": Mat3_Row(m, 2) = v: Text1.Text = Matrix3_ToStr(m): Text2.Text = ""
    
    
    MsgBox "Clear the 3x3-matrix m": m = Mat3_Clear:   Text1.Text = Matrix3_ToStr(m)
    MsgBox "Create a 3x1-vector v":  v = Vec3(11, 21, 31): Text2.Text = Vector3_ToStr(v)
    MsgBox "Write the vector to col 0 of m (Property Let Col)": Mat3_Col(m, 0) = v: Text1.Text = Matrix3_ToStr(m): Text2.Text = ""
    
    MsgBox "Create a 3x1-vector v":  v = Vec3(12, 22, 32): Text2.Text = Vector3_ToStr(v)
    MsgBox "Write the vector to col 1 of m (Property Let Col)": Mat3_Col(m, 1) = v: Text1.Text = Matrix3_ToStr(m): Text2.Text = ""
    
    MsgBox "Create a 3x1-vector v":  v = Vec3(13, 23, 33): Text2.Text = Vector3_ToStr(v)
    MsgBox "Write the vector to col 2 of m (Property Let Col)": Mat3_Col(m, 2) = v: Text1.Text = Matrix3_ToStr(m): Text2.Text = ""

'    Dim m1 As Matrix3, m2 As Matrix3
'    Mat3_Row(m1, 0) = Vec3(11, 12, 13)
'    Mat3_Row(m1, 1) = Vec3(21, 22, 23)
'    Mat3_Row(m1, 2) = Vec3(31, 32, 33)
'    MsgBox Matrix3_ToStr(m1)
'    MsgBox Vector3_ToStr(Mat3_Row(m1, 0), True)
'    MsgBox Vector3_ToStr(Mat3_Row(m1, 1), True)
'    MsgBox Vector3_ToStr(Mat3_Row(m1, 2), True)
'
'    Mat3_Col(m2, 0) = Vec3(11, 21, 31)
'    Mat3_Col(m2, 1) = Vec3(12, 22, 32)
'    Mat3_Col(m2, 2) = Vec3(13, 23, 33)
'    MsgBox Matrix3_ToStr(m2)
'    MsgBox Vector3_ToStr(Mat3_Col(m2, 0), False)
'    MsgBox Vector3_ToStr(Mat3_Col(m2, 1), False)
'    MsgBox Vector3_ToStr(Mat3_Col(m2, 2), False)
End Sub
Sub Test_PropGetLetRowCol_4()
    Dim m As Matrix4
    Dim v As Vector4

    MsgBox "Create an empty 4x4-matrix m": Text1.Text = Matrix4_ToStr(m)
    MsgBox "Create a 4x1-vector v":  v = Vec4(11, 12, 13, 14): Text2.Text = Vector4_ToStr(v)
    MsgBox "Write the vector to row 0 of m (Property Let Row)": Mat4_Row(m, 0) = v: Text1.Text = Matrix4_ToStr(m): Text2.Text = ""
        
    MsgBox "Create a 4x1-vector v":  v = Vec4(21, 22, 23, 24): Text2.Text = Vector4_ToStr(v)
    MsgBox "Write the vector to row 1 of m (Property Let Row)": Mat4_Row(m, 1) = v: Text1.Text = Matrix4_ToStr(m): Text2.Text = ""
        
    MsgBox "Create a 4x1-vector v":  v = Vec4(31, 32, 33, 34): Text2.Text = Vector4_ToStr(v)
    MsgBox "Write the vector to row 2 of m (Property Let Row)": Mat4_Row(m, 2) = v: Text1.Text = Matrix4_ToStr(m): Text2.Text = ""
        
    MsgBox "Create a 4x1-vector v":  v = Vec4(41, 42, 43, 44): Text2.Text = Vector4_ToStr(v)
    MsgBox "Write the vector to row 3 of m (Property Let Row)": Mat4_Row(m, 3) = v: Text1.Text = Matrix4_ToStr(m): Text2.Text = ""
    
    
    MsgBox "Clear the 4x4-matrix m": m = Mat4_Clear:   Text1.Text = Matrix4_ToStr(m)
    MsgBox "Create a 4x1-vector v":  v = Vec4(11, 21, 31, 41): Text2.Text = Vector4_ToStr(v)
    MsgBox "Write the vector to col 0 of m (Property Let Col)": Mat4_Col(m, 0) = v: Text1.Text = Matrix4_ToStr(m): Text2.Text = ""
    
    MsgBox "Create a 4x1-vector v":  v = Vec4(12, 22, 32, 42): Text2.Text = Vector4_ToStr(v)
    MsgBox "Write the vector to col 1 of m (Property Let Col)": Mat4_Col(m, 1) = v: Text1.Text = Matrix4_ToStr(m): Text2.Text = ""
    
    MsgBox "Create a 4x1-vector v":  v = Vec4(13, 23, 33, 43): Text2.Text = Vector4_ToStr(v)
    MsgBox "Write the vector to col 2 of m (Property Let Col)": Mat4_Col(m, 2) = v: Text1.Text = Matrix4_ToStr(m): Text2.Text = ""
    
    MsgBox "Create a 4x1-vector v":  v = Vec4(14, 24, 34, 44): Text2.Text = Vector4_ToStr(v)
    MsgBox "Write the vector to col 3 of m (Property Let Col)": Mat4_Col(m, 3) = v: Text1.Text = Matrix4_ToStr(m): Text2.Text = ""

'    Dim m1 As Matrix4, m2 As Matrix4
'    Mat4_Row(m1, 0) = Vec4(11, 12, 13, 14)
'    Mat4_Row(m1, 1) = Vec4(21, 22, 23, 24)
'    Mat4_Row(m1, 2) = Vec4(31, 32, 33, 34)
'    Mat4_Row(m1, 3) = Vec4(41, 42, 43, 44)
'    MsgBox Matrix4_ToStr(m1)
'    MsgBox Vector4_ToStr(Mat4_Row(m1, 0), True)
'    MsgBox Vector4_ToStr(Mat4_Row(m1, 1), True)
'    MsgBox Vector4_ToStr(Mat4_Row(m1, 2), True)
'    MsgBox Vector4_ToStr(Mat4_Row(m1, 3), True)
'
'    Mat4_Col(m2, 0) = Vec4(11, 21, 31, 41)
'    Mat4_Col(m2, 1) = Vec4(12, 22, 32, 42)
'    Mat4_Col(m2, 2) = Vec4(13, 23, 33, 43)
'    Mat4_Col(m2, 3) = Vec4(14, 24, 34, 44)
'    MsgBox Matrix4_ToStr(m2)
'    MsgBox Vector4_ToStr(Mat4_Col(m2, 0), False)
'    MsgBox Vector4_ToStr(Mat4_Col(m2, 1), False)
'    MsgBox Vector4_ToStr(Mat4_Col(m2, 2), False)
'    MsgBox Vector4_ToStr(Mat4_Col(m2, 3), False)
End Sub
Sub Test_PropGetLetRowCol_5()
    Dim m As Matrix5
    Dim v As Vector5

    MsgBox "Create an empty 5x5-matrix m": Text1.Text = Matrix5_ToStr(m)
    MsgBox "Create a 5x1-vector v":  v = Vec5(11, 12, 13, 14, 15): Text2.Text = Vector5_ToStr(v)
    MsgBox "Write the vector to row 0 of m (Property Let Row)": Mat5_Row(m, 0) = v: Text1.Text = Matrix5_ToStr(m): Text2.Text = ""
        
    MsgBox "Create a 5x1-vector v":  v = Vec5(21, 22, 23, 24, 25): Text2.Text = Vector5_ToStr(v)
    MsgBox "Write the vector to row 1 of m (Property Let Row)": Mat5_Row(m, 1) = v: Text1.Text = Matrix5_ToStr(m): Text2.Text = ""
        
    MsgBox "Create a 5x1-vector v":  v = Vec5(31, 32, 33, 34, 35): Text2.Text = Vector5_ToStr(v)
    MsgBox "Write the vector to row 2 of m (Property Let Row)": Mat5_Row(m, 2) = v: Text1.Text = Matrix5_ToStr(m): Text2.Text = ""
        
    MsgBox "Create a 5x1-vector v":  v = Vec5(41, 42, 43, 44, 45): Text2.Text = Vector5_ToStr(v)
    MsgBox "Write the vector to row 3 of m (Property Let Row)": Mat5_Row(m, 3) = v: Text1.Text = Matrix5_ToStr(m): Text2.Text = ""
        
    MsgBox "Create a 5x1-vector v":  v = Vec5(51, 52, 53, 54, 55): Text2.Text = Vector5_ToStr(v)
    MsgBox "Write the vector to row 4 of m (Property Let Row)": Mat5_Row(m, 4) = v: Text1.Text = Matrix5_ToStr(m): Text2.Text = ""
    
    
    MsgBox "Clear the 5x5-matrix m": m = Mat5_Clear:   Text1.Text = Matrix5_ToStr(m)
    MsgBox "Create a 5x1-vector v":  v = Vec5(11, 21, 31, 41, 51): Text2.Text = Vector5_ToStr(v)
    MsgBox "Write the vector to col 0 of m (Property Let Col)": Mat5_Col(m, 0) = v: Text1.Text = Matrix5_ToStr(m): Text2.Text = ""
    
    MsgBox "Create a 5x1-vector v":  v = Vec5(12, 22, 32, 42, 52): Text2.Text = Vector5_ToStr(v)
    MsgBox "Write the vector to col 1 of m (Property Let Col)": Mat5_Col(m, 1) = v: Text1.Text = Matrix5_ToStr(m): Text2.Text = ""
    
    MsgBox "Create a 5x1-vector v":  v = Vec5(13, 23, 33, 43, 53): Text2.Text = Vector5_ToStr(v)
    MsgBox "Write the vector to col 2 of m (Property Let Col)": Mat5_Col(m, 2) = v: Text1.Text = Matrix5_ToStr(m): Text2.Text = ""
    
    MsgBox "Create a 5x1-vector v":  v = Vec5(14, 24, 34, 44, 54): Text2.Text = Vector5_ToStr(v)
    MsgBox "Write the vector to col 3 of m (Property Let Col)": Mat5_Col(m, 3) = v: Text1.Text = Matrix5_ToStr(m): Text2.Text = ""
    
    MsgBox "Create a 5x1-vector v":  v = Vec5(15, 25, 35, 45, 55): Text2.Text = Vector5_ToStr(v)
    MsgBox "Write the vector to col 4 of m (Property Let Col)": Mat5_Col(m, 4) = v: Text1.Text = Matrix5_ToStr(m): Text2.Text = ""

'    Dim m1 As Matrix5, m2 As Matrix5
'    Mat5_Row(m1, 0) = Vec5(11, 12, 13, 14, 15)
'    Mat5_Row(m1, 1) = Vec5(21, 22, 23, 24, 25)
'    Mat5_Row(m1, 2) = Vec5(31, 32, 33, 34, 35)
'    Mat5_Row(m1, 3) = Vec5(41, 42, 43, 44, 45)
'    Mat5_Row(m1, 4) = Vec5(51, 52, 53, 54, 55)
'    MsgBox Matrix5_ToStr(m1)
'    MsgBox Vector5_ToStr(Mat5_Row(m1, 0), True)
'    MsgBox Vector5_ToStr(Mat5_Row(m1, 1), True)
'    MsgBox Vector5_ToStr(Mat5_Row(m1, 2), True)
'    MsgBox Vector5_ToStr(Mat5_Row(m1, 3), True)
'    MsgBox Vector5_ToStr(Mat5_Row(m1, 4), True)
'
'    Mat5_Col(m2, 0) = Vec5(11, 21, 31, 41, 51)
'    Mat5_Col(m2, 1) = Vec5(12, 22, 32, 42, 52)
'    Mat5_Col(m2, 2) = Vec5(13, 23, 33, 43, 53)
'    Mat5_Col(m2, 3) = Vec5(14, 24, 34, 44, 54)
'    Mat5_Col(m2, 4) = Vec5(15, 25, 35, 45, 55)
'    MsgBox Matrix5_ToStr(m2)
'    MsgBox Vector5_ToStr(Mat5_Col(m2, 0), False)
'    MsgBox Vector5_ToStr(Mat5_Col(m2, 1), False)
'    MsgBox Vector5_ToStr(Mat5_Col(m2, 2), False)
'    MsgBox Vector5_ToStr(Mat5_Col(m2, 3), False)
'    MsgBox Vector5_ToStr(Mat5_Col(m2, 4), False)
End Sub
Sub Test_PropGetLetRowCol_6()
    Dim m As Matrix6
    Dim v As Vector6

    MsgBox "Create an empty 6x6-matrix m": Text1.Text = Matrix6_ToStr(m)
    MsgBox "Create a 6x1-vector v":  v = Vec6(11, 12, 13, 14, 15, 16): Text2.Text = Vector6_ToStr(v)
    MsgBox "Write the vector to row 0 of m (Property Let Row)": Mat6_Row(m, 0) = v: Text1.Text = Matrix6_ToStr(m): Text2.Text = ""
        
    MsgBox "Create a 6x1-vector v":  v = Vec6(21, 22, 23, 24, 25, 26): Text2.Text = Vector6_ToStr(v)
    MsgBox "Write the vector to row 1 of m (Property Let Row)": Mat6_Row(m, 1) = v: Text1.Text = Matrix6_ToStr(m): Text2.Text = ""
        
    MsgBox "Create a 6x1-vector v":  v = Vec6(31, 32, 33, 34, 35, 36): Text2.Text = Vector6_ToStr(v)
    MsgBox "Write the vector to row 2 of m (Property Let Row)": Mat6_Row(m, 2) = v: Text1.Text = Matrix6_ToStr(m): Text2.Text = ""
        
    MsgBox "Create a 6x1-vector v":  v = Vec6(41, 42, 43, 44, 45, 46): Text2.Text = Vector6_ToStr(v)
    MsgBox "Write the vector to row 3 of m (Property Let Row)": Mat6_Row(m, 3) = v: Text1.Text = Matrix6_ToStr(m): Text2.Text = ""
        
    MsgBox "Create a 6x1-vector v":  v = Vec6(51, 52, 53, 54, 55, 56): Text2.Text = Vector6_ToStr(v)
    MsgBox "Write the vector to row 4 of m (Property Let Row)": Mat6_Row(m, 4) = v: Text1.Text = Matrix6_ToStr(m): Text2.Text = ""
        
    MsgBox "Create a 6x1-vector v":  v = Vec6(61, 62, 63, 64, 65, 66): Text2.Text = Vector6_ToStr(v)
    MsgBox "Write the vector to row 5 of m (Property Let Row)": Mat6_Row(m, 5) = v: Text1.Text = Matrix6_ToStr(m): Text2.Text = ""
    
    
    MsgBox "Clear the 6x6-matrix m": m = Mat6_Clear:   Text1.Text = Matrix6_ToStr(m)
    MsgBox "Create a 6x1-vector v":  v = Vec6(11, 21, 31, 41, 51, 61): Text2.Text = Vector6_ToStr(v)
    MsgBox "Write the vector to col 0 of m (Property Let Col)": Mat6_Col(m, 0) = v: Text1.Text = Matrix6_ToStr(m): Text2.Text = ""
    
    MsgBox "Create a 6x1-vector v":  v = Vec6(12, 22, 32, 42, 52, 62): Text2.Text = Vector6_ToStr(v)
    MsgBox "Write the vector to col 1 of m (Property Let Col)": Mat6_Col(m, 1) = v: Text1.Text = Matrix6_ToStr(m): Text2.Text = ""
    
    MsgBox "Create a 6x1-vector v":  v = Vec6(13, 23, 33, 43, 53, 63): Text2.Text = Vector6_ToStr(v)
    MsgBox "Write the vector to col 2 of m (Property Let Col)": Mat6_Col(m, 2) = v: Text1.Text = Matrix6_ToStr(m): Text2.Text = ""
    
    MsgBox "Create a 6x1-vector v":  v = Vec6(14, 24, 34, 44, 54, 64): Text2.Text = Vector6_ToStr(v)
    MsgBox "Write the vector to col 3 of m (Property Let Col)": Mat6_Col(m, 3) = v: Text1.Text = Matrix6_ToStr(m): Text2.Text = ""
    
    MsgBox "Create a 6x1-vector v":  v = Vec6(15, 25, 35, 45, 55, 65): Text2.Text = Vector6_ToStr(v)
    MsgBox "Write the vector to col 4 of m (Property Let Col)": Mat6_Col(m, 4) = v: Text1.Text = Matrix6_ToStr(m): Text2.Text = ""
    
    MsgBox "Create a 6x1-vector v":  v = Vec6(16, 26, 36, 46, 56, 66): Text2.Text = Vector6_ToStr(v)
    MsgBox "Write the vector to col 5 of m (Property Let Col)": Mat6_Col(m, 5) = v: Text1.Text = Matrix6_ToStr(m): Text2.Text = ""

'    Dim m1 As Matrix6, m2 As Matrix6
'    Mat6_Row(m1, 0) = Vec6(11, 12, 13, 14, 15, 16)
'    Mat6_Row(m1, 1) = Vec6(21, 22, 23, 24, 25, 26)
'    Mat6_Row(m1, 2) = Vec6(31, 32, 33, 34, 35, 36)
'    Mat6_Row(m1, 3) = Vec6(41, 42, 43, 44, 45, 46)
'    Mat6_Row(m1, 4) = Vec6(51, 52, 53, 54, 55, 56)
'    Mat6_Row(m1, 5) = Vec6(61, 62, 63, 64, 65, 66)
'    MsgBox Matrix6_ToStr(m1)
'    MsgBox Vector6_ToStr(Mat6_Row(m1, 0), True)
'    MsgBox Vector6_ToStr(Mat6_Row(m1, 1), True)
'    MsgBox Vector6_ToStr(Mat6_Row(m1, 2), True)
'    MsgBox Vector6_ToStr(Mat6_Row(m1, 3), True)
'    MsgBox Vector6_ToStr(Mat6_Row(m1, 4), True)
'    MsgBox Vector6_ToStr(Mat6_Row(m1, 5), True)
'
'    Mat6_Col(m2, 0) = Vec6(11, 21, 31, 41, 51, 61)
'    Mat6_Col(m2, 1) = Vec6(12, 22, 32, 42, 52, 62)
'    Mat6_Col(m2, 2) = Vec6(13, 23, 33, 43, 53, 63)
'    Mat6_Col(m2, 3) = Vec6(14, 24, 34, 44, 54, 64)
'    Mat6_Col(m2, 4) = Vec6(15, 25, 35, 45, 55, 65)
'    Mat6_Col(m2, 5) = Vec6(16, 26, 36, 46, 56, 66)
'    MsgBox Matrix6_ToStr(m2)
'    MsgBox Vector6_ToStr(Mat6_Col(m2, 0), False)
'    MsgBox Vector6_ToStr(Mat6_Col(m2, 1), False)
'    MsgBox Vector6_ToStr(Mat6_Col(m2, 2), False)
'    MsgBox Vector6_ToStr(Mat6_Col(m2, 3), False)
'    MsgBox Vector6_ToStr(Mat6_Col(m2, 4), False)
'    MsgBox Vector6_ToStr(Mat6_Col(m2, 5), False)
End Sub

Sub Test_MatE_6()
    Dim m As Matrix6: m = Mat6_E
    MsgBox "Create 6x6-unity-matrix"
    Text1.Text = Matrix6_ToStr(m)
End Sub
