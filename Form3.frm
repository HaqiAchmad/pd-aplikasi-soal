VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   6225
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10860
   LinkTopic       =   "Form3"
   ScaleHeight     =   6225
   ScaleWidth      =   10860
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command22 
      Caption         =   "Next"
      Height          =   495
      Left            =   9120
      TabIndex        =   26
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton Command21 
      Caption         =   "Prevous"
      Height          =   495
      Left            =   7320
      TabIndex        =   25
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton Command20 
      Caption         =   "20"
      Height          =   495
      Left            =   9840
      TabIndex        =   24
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton Command19 
      Caption         =   "19"
      Height          =   495
      Left            =   9240
      TabIndex        =   23
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton Command18 
      Caption         =   "18"
      Height          =   495
      Left            =   8640
      TabIndex        =   22
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton Command17 
      Caption         =   "17"
      Height          =   495
      Left            =   8040
      TabIndex        =   21
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton Command16 
      Caption         =   "16"
      Height          =   495
      Left            =   7440
      TabIndex        =   20
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton Command15 
      Caption         =   "15"
      Height          =   495
      Left            =   9840
      TabIndex        =   19
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton Command14 
      Caption         =   "14"
      Height          =   495
      Left            =   9240
      TabIndex        =   18
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton Command13 
      Caption         =   "13"
      Height          =   495
      Left            =   8640
      TabIndex        =   17
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton Command12 
      Caption         =   "12"
      Height          =   495
      Left            =   8040
      TabIndex        =   16
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton Command11 
      Caption         =   "11"
      Height          =   495
      Left            =   7440
      TabIndex        =   15
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton Command10 
      Caption         =   "10"
      Height          =   495
      Left            =   9840
      TabIndex        =   14
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Command9 
      Caption         =   "9"
      Height          =   495
      Left            =   9240
      TabIndex        =   13
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Command8 
      Caption         =   "8"
      Height          =   495
      Left            =   8640
      TabIndex        =   12
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Command7 
      Caption         =   "7"
      Height          =   495
      Left            =   8040
      TabIndex        =   11
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      Caption         =   "6"
      Height          =   495
      Left            =   7440
      TabIndex        =   10
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      Caption         =   "5"
      Height          =   495
      Left            =   9840
      TabIndex        =   9
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "4"
      Height          =   495
      Left            =   9240
      TabIndex        =   8
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "3"
      Height          =   495
      Left            =   8640
      TabIndex        =   7
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "2"
      Height          =   495
      Left            =   8040
      TabIndex        =   6
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   495
      Left            =   7440
      TabIndex        =   5
      Top             =   1320
      Width           =   495
   End
   Begin VB.OptionButton Option5 
      Caption         =   "E. Real"
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   5280
      Width           =   3975
   End
   Begin VB.OptionButton Option4 
      Caption         =   "D. Byte"
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   4560
      Width           =   3975
   End
   Begin VB.OptionButton Option3 
      Caption         =   "C. Boolean"
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   3840
      Width           =   3975
   End
   Begin VB.OptionButton Option2 
      Caption         =   "B. Char"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   3120
      Width           =   3975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "A. String"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   2400
      Width           =   3975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Pilih Soal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   29
      Top             =   600
      Width           =   3015
   End
   Begin VB.Line Line1 
      X1              =   6960
      X2              =   6960
      Y1              =   0
      Y2              =   6480
   End
   Begin VB.Label Label2 
      Caption         =   "Tipe data yang hanya mempunyai nilai TRUE dan FALSE saja adalah"
      Height          =   855
      Left            =   0
      TabIndex        =   28
      Top             =   1080
      Width           =   6615
   End
   Begin VB.Label Label1 
      Caption         =   "Soal Ke 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   27
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form3.Hide
Form1.Show
End Sub

Private Sub Command10_Click()
Form3.Hide
Form10.Show
End Sub

Private Sub Command11_Click()
Form3.Hide
Form11.Show
End Sub

Private Sub Command12_Click()
Form3.Hide
Form12.Show
End Sub

Private Sub Command13_Click()
Form3.Hide
Form13.Show
End Sub

Private Sub Command14_Click()
Form3.Hide
Form14.Show
End Sub

Private Sub Command15_Click()
Form3.Hide
Form15.Show
End Sub

Private Sub Command16_Click()
Form3.Hide
Form16.Show
End Sub

Private Sub Command17_Click()
Form3.Hide
Form17.Show
End Sub

Private Sub Command18_Click()
Form3.Hide
Form18.Show
End Sub

Private Sub Command19_Click()
Form3.Hide
Form19.Show
End Sub

Private Sub Command2_Click()
Form3.Hide
Form2.Show
End Sub

Private Sub Command20_Click()
Form3.Hide
Form20.Show
End Sub

Private Sub Command21_Click()
Form3.Hide
Form2.Show
End Sub

Private Sub Command22_Click()
If Option3.Value = True Then
Form3.Hide
Form4.Show
ElseIf Option3.Value = False Then
MsgBox "Coba Lagi"

End If

End Sub

Private Sub Command4_Click()
Form3.Hide
Form4.Show
End Sub

Private Sub Command5_Click()
Form3.Hide
Form5.Show
End Sub

Private Sub Command6_Click()
Form3.Hide
Form6.Show

End Sub

Private Sub Command7_Click()
Form3.Hide
Form7.Show
End Sub

Private Sub Command8_Click()
Form3.Hide
Form8.Show
End Sub

Private Sub Command9_Click()
Form3.Hide
Form9.Show
End Sub

Private Sub Form_Load()
Option1.Value = False
Option2.Value = False
Option3.Value = False
Option4.Value = False
Option5.Value = False


End Sub


