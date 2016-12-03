VERSION 5.00
Begin VB.Form Form17 
   Caption         =   "Form17"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form17"
   Picture         =   "Form17.frx":0000
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "BACK"
      Height          =   495
      Left            =   11400
      TabIndex        =   10
      Top             =   6360
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      DragMode        =   1  'Automatic
      Height          =   375
      Left            =   19320
      Picture         =   "Form17.frx":AF498
      ScaleHeight     =   315
      ScaleWidth      =   915
      TabIndex        =   9
      Top             =   0
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form17.frx":B8C7D
      Left            =   8400
      List            =   "Form17.frx":B8C90
      TabIndex        =   4
      Text            =   "SELECT AMOUNT"
      Top             =   2520
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CHECK"
      Height          =   495
      Left            =   6840
      TabIndex        =   3
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   495
      Left            =   9120
      TabIndex        =   2
      Top             =   6360
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   8400
      TabIndex        =   1
      Top             =   3600
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   8400
      TabIndex        =   0
      Top             =   4800
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "AMOUNT"
      Height          =   375
      Left            =   6720
      TabIndex        =   8
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "TOP UP PACK"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6840
      TabIndex        =   7
      Top             =   480
      Width           =   4935
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "VALIDITY"
      Height          =   375
      Left            =   6720
      TabIndex        =   6
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "DESCRIPTION"
      Height          =   375
      Left            =   6720
      TabIndex        =   5
      Top             =   4920
      Width           =   1215
   End
End
Attribute VB_Name = "Form17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim connEmp As ADODB.Connection
Dim rsEmp As ADODB.Recordset
Dim COMM As ADODB.Command
Dim CM As ADODB.Command

Private Sub Command1_Click()
Set rsEmp = New ADODB.Recordset
rsEmp.Open "select validity,des from TOP where amt = '" & Combo1.Text & "'", connEmp, adOpenKeyset, adLockReadOnly, adCmdText
Text1.Text = Trim(rsEmp!validity)
Text2.Text = Trim(rsEmp!des)
rsEmp.Close
Set rsEmp = Nothing
End Sub

Private Sub Command2_Click()
Form1.Text6.Text = Combo1.Text
Form16.Hide
Form1.Show
End Sub

Private Sub Command3_Click()
Form17.Hide
Form6.Show
End Sub

Private Sub Form_Load()
Set connEmp = New ADODB.Connection
connEmp.Open "Provider=OraOLEDB.Oracle.1;Password=sa;Persist Security Info=True;User ID=system;Data Source=SARAVANAN"
connEmp.CursorLocation = adUseClient
Set rsEmp = New ADODB.Recordset
rsEmp.Open "select amt from TOP where net = '" & Form1.Combo1.Text & "'", connEmp, adOpenDynamic
Do While rsEmp.EOF <> True
Combo1.AddItem (rsEmp.Fields("amt"))
rsEmp.MoveNext
Loop
End Sub



Private Sub Picture1_Click()
End
End Sub
