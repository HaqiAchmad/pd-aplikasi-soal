VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   6495
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10830
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   10830
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option1 
      Caption         =   "A. Membuat algoritma"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   26
      Top             =   2400
      Width           =   3975
   End
   Begin VB.OptionButton Option2 
      Caption         =   "B. Membuat program"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   25
      Top             =   3120
      Width           =   3975
   End
   Begin VB.OptionButton Option3 
      Caption         =   "C. Membeli komputer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   24
      Top             =   3840
      Width           =   3975
   End
   Begin VB.OptionButton Option4 
      Caption         =   "D. Proses"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   23
      Top             =   4560
      Width           =   3975
   End
   Begin VB.OptionButton Option5 
      Caption         =   "E. Mempelajari program"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   22
      Top             =   5280
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   21
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      TabIndex        =   20
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8640
      TabIndex        =   19
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9240
      TabIndex        =   18
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9840
      TabIndex        =   17
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   16
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Command7 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      TabIndex        =   15
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Command8 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8640
      TabIndex        =   14
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Command9 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9240
      TabIndex        =   13
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Command10 
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9840
      TabIndex        =   12
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Command11 
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   11
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton Command12 
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      TabIndex        =   10
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton Command13 
      Caption         =   "13"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8640
      TabIndex        =   9
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton Command14 
      Caption         =   "14"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9240
      TabIndex        =   8
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton Command15 
      Caption         =   "15"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9840
      TabIndex        =   7
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton Command16 
      Caption         =   "16"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   6
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton Command17 
      Caption         =   "17"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      TabIndex        =   5
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton Command18 
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8640
      TabIndex        =   4
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton Command19 
      Caption         =   "19"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9240
      TabIndex        =   3
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton Command20 
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9840
      TabIndex        =   2
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton Command21 
      Caption         =   "Prevous"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      TabIndex        =   1
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton Command22 
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9120
      TabIndex        =   0
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Soal Ke 1"
      Height          =   375
      Left            =   240
      TabIndex        =   29
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label d 
      Caption         =   "Dalam menyusun suatu program langkah pertama yang harus dilakukan adalah...."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   28
      Top             =   1080
      Width           =   6615
   End
   Begin VB.Line Line1 
      X1              =   6960
      X2              =   6960
      Y1              =   0
      Y2              =   6480
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Pilih Soal"
      Height          =   375
      Left            =   7440
      TabIndex        =   27
      Top             =   600
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command10_Click()
Form1.Hide
Form10.Show
End Sub

Private Sub Command11_Click()
Form1.Hide
Form11.Show
End Sub

Private Sub Command12_Click()
Form1.Hide
Form12.Show
End Sub

Private Sub Command13_Click()
Form1.Hide
Form13.Show
End Sub

Private Sub Command14_Click()
Form1.Hide
for114.Show
End Sub

Private Sub Command15_Click()
Form1.Hide
Form15.Show
End Sub

Private Sub Command16_Click()
Form1.Hide
Form16.Show
End Sub

Private Sub Command17_Click()
Form1.Hide
Form17.Show
End Sub

Private Sub Command18_Click()
Form1.Hide
Form18.Show
End Sub

Private Sub Command19_Click()
Form1.Hide
Form19.Show
End Sub

Private Sub Command2_Click()
Form1.Hide
Form2.Show
End Sub

Private Sub Command20_Click()
Form1.Hide
Form20.Show
End Sub

Private Sub Command21_Click()
Form1.Hide
Form20.Show
End Sub

Private Sub Command22_Click()
If Option1.Value = True Then
Form1.Hide
Form2.Show
ElseIf Option1.Value = False Then
MsgBox "Coba Lagi"

End If

End Sub

Private Sub Command3_Click()
Form1.Hide
Form3.Show
End Sub

Private Sub Command4_Click()
Form1.Hide
Form4.Show
End Sub

Private Sub Command5_Click()
Form1.Hide
Form5.Show
End Sub

Private Sub Command6_Click()
Form1.Hide
Form6.Show

End Sub

Private Sub Command7_Click()
Form1.Hide
Form7.Show
End Sub

Private Sub Command8_Click()
Form1.Hide
Form8.Show
End Sub

Private Sub Command9_Click()
Form1.Hide
Form9.Show
End Sub

Private Sub Form_Load()
Option1.Value = False
Option2.Value = False
Option3.Value = False
Option4.Value = False
Option5.Value = False


End Sub

