VERSION 5.00
Begin VB.Form Form16 
   Caption         =   "Form16"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form16"
   Picture         =   "Form16.frx":0000
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "BACK"
      Height          =   495
      Left            =   10680
      TabIndex        =   10
      Top             =   6480
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      DragMode        =   1  'Automatic
      Height          =   375
      Left            =   19320
      Picture         =   "Form16.frx":FB4A8
      ScaleHeight     =   315
      ScaleWidth      =   915
      TabIndex        =   9
      Top             =   0
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   8160
      TabIndex        =   7
      Top             =   4920
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   8160
      TabIndex        =   6
      Top             =   3720
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   495
      Left            =   8400
      TabIndex        =   5
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CHECK"
      Height          =   495
      Left            =   6000
      TabIndex        =   4
      Top             =   6480
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form16.frx":104C8D
      Left            =   8160
      List            =   "Form16.frx":104CA0
      TabIndex        =   1
      Text            =   "SELECT AMOUNT"
      Top             =   2640
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "AMOUNT"
      Height          =   375
      Left            =   6480
      TabIndex        =   8
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "DESCRIPTION"
      Height          =   375
      Left            =   6480
      TabIndex        =   3
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "VALIDITY"
      Height          =   375
      Left            =   6480
      TabIndex        =   2
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "MESSAGE  PACK"
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
      Left            =   6360
      TabIndex        =   0
      Top             =   600
      Width           =   4935
   End
End
Attribute VB_Name = "Form16"
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
rsEmp.Open "select validity,des from msg where amt = '" & Combo1.Text & "'", connEmp, adOpenKeyset, adLockReadOnly, adCmdText
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
Form16.Hide
Form6.Show
End Sub

Private Sub Form_Load()
Set connEmp = New ADODB.Connection
connEmp.Open "Provider=OraOLEDB.Oracle.1;Password=sa;Persist Security Info=True;User ID=system;Data Source=SARAVANAN"
connEmp.CursorLocation = adUseClient
Set rsEmp = New ADODB.Recordset
rsEmp.Open "select amt from msg where net = '" & Form1.Combo1.Text & "'", connEmp, adOpenDynamic
Do While rsEmp.EOF <> True
Combo1.AddItem (rsEmp.Fields("amt"))
rsEmp.MoveNext
Loop
End Sub


Private Sub Picture1_Click()
End
End Sub
