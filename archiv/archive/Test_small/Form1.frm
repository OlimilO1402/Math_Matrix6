VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3615
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8835
   LinkTopic       =   "Form1"
   ScaleHeight     =   3615
   ScaleWidth      =   8835
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   120
      Width           =   2415
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
      Height          =   2535
      Left            =   2640
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   600
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2415
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
      Height          =   2535
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'so wis dasteht, Zeile, Spalte
Private Type Mat3
    aa As Double:    ab As Double:    ac As Double
    ba As Double:    bb As Double:    bc As Double
    ca As Double:    cb As Double:    cc As Double
End Type
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef pDst As Any, ByRef pSrc As Any, ByVal nBytes As Long)

Dim m As Mat3
Dim ma() As Double


Private Sub Form_Load()
    With m
        .aa = 11: .ab = 12: .ac = 13
        .ba = 21: .bb = 22: .bc = 23
        .ca = 31: .cb = 32: .cc = 33
    End With
End Sub
Private Sub Command1_Click()
    Text1.Text = Mat3_ToStr(m)
End Sub
Private Function Mat3_ToStr(m As Mat3) As String
    Dim s As String
    With m
        s = s & .aa & " " & .ab & " " & .ac & vbCrLf
        s = s & .ba & " " & .bb & " " & .bc & vbCrLf
        s = s & .ca & " " & .cb & " " & .cc & vbCrLf
    End With
    Mat3_ToStr = s
End Function

Private Sub Command2_Click()
    Text2.Text = MatA_ToStr(VarPtr(m), 3, 3)
End Sub
Function MatA_ToStr(ByVal pMat As Long, ByVal mRows As Long, ByVal nCols As Long)
    ReDim ma(0 To mRows - 1, 0 To nCols - 1) As Double
    RtlMoveMemory ma(0, 0), ByVal pMat, mRows * nCols * 8
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
    MatA_ToStr = s
End Function

