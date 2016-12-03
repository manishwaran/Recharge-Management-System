VERSION 5.00
Begin VB.Form Form10 
   Caption         =   "Form10"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form10"
   Picture         =   "Form10.frx":0000
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command6 
      Caption         =   "MAIN  MENU"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6240
      TabIndex        =   7
      Top             =   7800
      Width           =   3255
   End
   Begin VB.PictureBox Picture1 
      DragMode        =   1  'Automatic
      Height          =   375
      Left            =   19320
      Picture         =   "Form10.frx":E8D7C
      ScaleHeight     =   315
      ScaleWidth      =   915
      TabIndex        =   6
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "SPECIAL PACK"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6240
      TabIndex        =   4
      Top             =   6600
      Width           =   3255
   End
   Begin VB.CommandButton Command4 
      Caption         =   "TOPUP PACK"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6240
      TabIndex        =   3
      Top             =   5400
      Width           =   3255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "NET PACK"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6240
      TabIndex        =   2
      Top             =   4320
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "MESSAGE PACK"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6240
      TabIndex        =   1
      Top             =   3240
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "NETWORK"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6240
      TabIndex        =   0
      Top             =   2160
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "MANIPULATE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   855
      Left            =   5280
      TabIndex        =   5
      Top             =   360
      Width           =   5175
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form10.Hide
Form11.Show
End Sub

Private Sub Command2_Click()
Form10.Hide
Form12.Show
End Sub

Private Sub Command3_Click()
Form10.Hide
Form13.Show
End Sub

Private Sub Command4_Click()
Form10.Hide
Form14.Show
End Sub

Private Sub Command5_Click()
Form10.Hide
Form15.Show
End Sub


Private Sub Command6_Click()
Form10.Hide
Form9.Show
End Sub


Private Sub Picture1_Click()
End
End Sub
