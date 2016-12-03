VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form6"
   Picture         =   "Form61.frx":0000
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command5 
      Caption         =   "BACK"
      Height          =   495
      Left            =   7920
      TabIndex        =   6
      Top             =   7080
      Width           =   3135
   End
   Begin VB.PictureBox Picture1 
      DragMode        =   1  'Automatic
      Height          =   375
      Left            =   19320
      Picture         =   "Form61.frx":126BB5
      ScaleHeight     =   315
      ScaleWidth      =   915
      TabIndex        =   5
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "SPECIAL  PACK"
      Height          =   495
      Left            =   7920
      TabIndex        =   4
      Top             =   5880
      Width           =   3135
   End
   Begin VB.CommandButton Command3 
      Caption         =   "NET  PACK"
      Height          =   495
      Left            =   7920
      TabIndex        =   3
      Top             =   4560
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      Caption         =   "TOPUP  PACK"
      Height          =   495
      Left            =   7920
      TabIndex        =   2
      Top             =   3360
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "MESSAGE PACK"
      Height          =   495
      Left            =   7920
      TabIndex        =   1
      Top             =   2040
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "BROWSE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7080
      TabIndex        =   0
      Top             =   360
      Width           =   4815
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form6.Hide
Form16.Show
End Sub

Private Sub Command2_Click()
Form6.Hide
Form17.Show
End Sub

Private Sub Command3_Click()
Form6.Hide
Form18.Show
End Sub

Private Sub Command4_Click()
Form6.Hide
Form19.Show
End Sub

Private Sub Command5_Click()
Form6.Hide
Form1.Show
End Sub

Private Sub Picture1_Click()
End
End Sub
