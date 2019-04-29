VERSION 5.00
Begin VB.Form Login2 
   Caption         =   "Login2"
   ClientHeight    =   6195
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10845
   LinkTopic       =   "Form2"
   ScaleHeight     =   6195
   ScaleWidth      =   10845
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Mulai Ujian"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   3
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label Label3 
      Height          =   495
      Left            =   6840
      TabIndex        =   2
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label2 
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Pemrograman Dasar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   0
      Top             =   720
      Width           =   5175
   End
End
Attribute VB_Name = "Login2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Login2.Hide
Form1.Show
End Sub

Private Sub Form_Load()
Label2.Caption = Time
Label3.Caption = Date
End Sub
