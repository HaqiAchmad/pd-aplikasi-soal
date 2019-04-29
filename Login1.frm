VERSION 5.00
Begin VB.Form Login1 
   Caption         =   "Login1"
   ClientHeight    =   6240
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10800
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
   ScaleHeight     =   6240
   ScaleWidth      =   10800
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Keluar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6480
      TabIndex        =   10
      Top             =   4320
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Height          =   420
      Left            =   8160
      TabIndex        =   9
      Top             =   1800
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Height          =   420
      Left            =   8160
      TabIndex        =   8
      Top             =   2400
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Height          =   420
      Left            =   8160
      TabIndex        =   7
      Top             =   3000
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   420
      Left            =   8160
      TabIndex        =   2
      Top             =   1200
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8280
      TabIndex        =   0
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "No Absen"
      Height          =   375
      Left            =   6000
      TabIndex        =   6
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Kelas"
      Height          =   375
      Left            =   6000
      TabIndex        =   5
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Jurusan"
      Height          =   375
      Left            =   6000
      TabIndex        =   4
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Nama"
      Height          =   375
      Left            =   6000
      TabIndex        =   3
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Masukan data diri anda"
      Height          =   375
      Left            =   6000
      TabIndex        =   1
      Top             =   480
      Width           =   4335
   End
   Begin VB.Image Image1 
      Height          =   4140
      Left            =   960
      Picture         =   "Login1.frx":0000
      Top             =   480
      Width           =   3345
   End
   Begin VB.Line Line1 
      X1              =   5400
      X2              =   5400
      Y1              =   0
      Y2              =   6240
   End
End
Attribute VB_Name = "Login1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Login1.Hide
Login2.Show
End Sub

Private Sub Command2_Click()
End
End Sub
