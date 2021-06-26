VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   7065
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox Text3 
      Height          =   2175
      Left            =   2760
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   2175
      Left            =   1440
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   2175
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    With Text1
        .FontName = "Consolas"
        .FontSize = 11
        .Text = ""
    End With
    Set Text2.Font = Text1.Font
    Set Text3.Font = Text2.Font
End Sub
Function NewFont(fnam As String, siz As Byte) As StdFont
    NewFont.Name = fnam
    NewFont.Size = siz
End Function
Private Sub Command1_Click()
    
    Dim v1 As Vector3: v1 = Vec3(1, 4, 3)
    Dim v2 As Vector3: v2 = Vec3(-2, 1.2, -1)
    
    Dim vErg As Vector3: vErg = Vec3_cross(v1, v2)
    Text1.Text = Vec3_ToStr(vErg)
    
    
    Dim v0 As Vector3: v0 = Vec3(1, 1, 1)   'Hilfsvektor
    Dim m As Matrix3
    Mat3_Col(m, 0) = v0
    Mat3_Col(m, 1) = v1
    Mat3_Col(m, 2) = v2
    Text2.Text = Mat3_det(m)
    
    Text3.Text = Vec3_ToStr(Mat3_detV(m))
    
    
    Dim m2 As Matrix3
    Mat3_Row(m2, 0) = v1
    Mat3_Row(m2, 1) = v2
    Mat3_Row(m2, 2) = v0
    
    Dim mAdj As Matrix3: mAdj = Mat3_Adj(m2)
    
    vErg = Mat3_Col(mAdj, 2)
    
    Debug.Print Vec3_ToStr(vErg)
    
End Sub
