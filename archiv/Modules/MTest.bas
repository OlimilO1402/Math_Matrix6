Attribute VB_Name = "MTest"
Option Explicit

'ein paar alte Testroutinen
'
'Private Sub BtnTest10_Click()
'    Text1.Text = Form1.GetMat10
'    Text2.Text = Form1.GetMat10
'
'End Sub
'
'Private Sub Command2_Click()
'    MyMat10_1 = MMatrices.Matrix10_Rnd
'    MyMat10_2 = MyMat10_1
'
'    Text1.Text = Matrix10_ToStr(MyMat10_1)
'    Text2.Text = Matrix10_ToStr(MyMat10_2)
'
'End Sub
'
'Private Sub Command8_Click()
'    Dim dt As Single, det10_1 As Double, det10_2 As Double
'
'    dt = Timer
'    det10_1 = MMatrices.Matrix10_det(MyMat10_1)
'    dt = Timer - dt
'
'    Text1.Text = dt & " " & det10_1
'
'    dt = Timer
'    det10_2 = MMatAlt.Matrix10_detX(MyMat10_2)
'    dt = Timer - dt
'
'    Text2.Text = dt & " " & det10_2
'
'
'    Text3.Text = det10_1 = det10_2
'
'End Sub
'
'
'Private Sub Command3_Click()
'    Dim a4 As Matrix4
'    Dim T As String
'    a4 = Mat4(11, 12, 13, -14, _
'              21, 22, 23, 24, _
'              31, 32, 33, 34, _
'              41, 42, 43, 44)
'    T = MMatrices.Matrix4_ToStr(a4)
'    MsgBox T
'    a4 = Matrix4_Parse(T)
'    T = MMatrices.Matrix4_ToStr(a4)
'    MsgBox T
'    Dim d As Double: d = Matrix4_det(a4)
'    MsgBox d
'End Sub
'
'Private Sub Command4_Click()
'    'Test2
'    'Test3
'    'Test4
'    'Test5
'    'Test6
'    Test_umat5
'End Sub
'
'Sub Test_umat3()
'    Dim m As Matrix3: m = Mat3(11, 12, 13, _
'                               21, 22, 23, _
'                               31, 32, 33)
'    Dim i As Long, j As Long
'    For i = 0 To 2: For j = 0 To 2
'        Dim um As Matrix2: um = Matrix3_umat(m, i, j)
'        Debug.Print Matrix2_ToStr(um)
'    Next: Next
'End Sub
'Sub Test_umat4()
'    Dim m As Matrix4: m = Mat4(11, 12, 13, 14, _
'                               21, 22, 23, 24, _
'                               31, 32, 33, 34, _
'                               41, 42, 43, 44)
'    Dim i As Long, j As Long
'    For i = 0 To 3: For j = 0 To 3
'        Dim um As Matrix3: um = Matrix4_umat(m, i, j)
'        Debug.Print Matrix3_ToStr(um)
'    Next: Next
'End Sub
'Sub Test_umat5()
'    Dim m As Matrix5: m = Mat5(11, 12, 13, 14, 15, _
'                               21, 22, 23, 24, 25, _
'                               31, 32, 33, 34, 35, _
'                               41, 42, 43, 44, 45, _
'                               51, 52, 53, 54, 55)
'    Dim i As Long, j As Long
'    For i = 0 To 4: For j = 0 To 4
'        Dim um As Matrix4: um = Matrix5_umat(m, i, j)
'        Debug.Print MMatrices.Matrix4_ToStr(um)
'    Next: Next
'End Sub
'Sub Test_umat6()
'    Dim m As Matrix6: m = Mat6(11, 12, 13, 14, 15, 16, _
'                               21, 22, 23, 24, 25, 26, _
'                               31, 32, 33, 34, 35, 36, _
'                               41, 42, 43, 44, 45, 46, _
'                               51, 52, 53, 54, 55, 56, _
'                               61, 62, 63, 64, 65, 66)
'    Dim i As Long, j As Long
'    For i = 0 To 5: For j = 0 To 5
'        Dim um As Matrix5: um = Matrix6_umat(m, i, j)
'        Debug.Print Matrix5_ToStr(um)
'    Next: Next
'End Sub
'
''Test Property Get/Let Row / Col
'Sub Test2()
'    Dim m1 As Matrix2, m2 As Matrix2
'    Mat2_Row(m1, 0) = Vec2(11, 12)
'    Mat2_Row(m1, 1) = Vec2(21, 22)
'    MsgBox Matrix2_ToStr(m1)
'    MsgBox Vector2_ToStr(Mat2_Row(m1, 0), True)
'    MsgBox Vector2_ToStr(Mat2_Row(m1, 1), True)
'
'    Mat2_Col(m2, 0) = Vec2(11, 21)
'    Mat2_Col(m2, 1) = Vec2(12, 22)
'    MsgBox Matrix2_ToStr(m2)
'    MsgBox Vector2_ToStr(Mat2_Col(m2, 0), False)
'    MsgBox Vector2_ToStr(Mat2_Col(m2, 1), False)
'End Sub
'Sub Test3()
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
'End Sub
'Sub Test4()
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
'End Sub
'Sub Test5()
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
'End Sub
'Sub Test6()
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
'End Sub
'
'Private Sub Command5_Click()
'    MsgBox MMatrices.Matrix6_ToStr(Mat6_E)
'End Sub
'
'Private Sub Command6_Click()
'    Dim v6 As Vector6: v6 = Vec6(0.1, 0.000000000001, 13.2, 15.678, 123.5, 1089.4)
'    Text1.Text = Vector6_ToStr(v6)
'End Sub
'
