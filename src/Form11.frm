VERSION 5.00
Begin VB.Form Form11 
   Caption         =   "Form11"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form11"
   Picture         =   "Form11.frx":0000
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      Caption         =   "BACK"
      Height          =   495
      Left            =   9480
      TabIndex        =   7
      Top             =   4080
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      DragMode        =   1  'Automatic
      Height          =   375
      Left            =   19320
      Picture         =   "Form11.frx":C53C4
      ScaleHeight     =   315
      ScaleWidth      =   915
      TabIndex        =   6
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "MAIN  MENU"
      Height          =   495
      Left            =   11760
      TabIndex        =   4
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "DELETE"
      Height          =   495
      Left            =   7080
      TabIndex        =   3
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ADD"
      Height          =   495
      Left            =   4560
      TabIndex        =   2
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   6720
      TabIndex        =   0
      Top             =   2520
      Width           =   4575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "NETWORK"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7080
      TabIndex        =   5
      Top             =   600
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "NETWORK"
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   2520
      Width           =   1335
   End
End
Attribute VB_Name = "Form11"
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
rsEmp.Open "select * from SCHEME where NET = '" & Text1.Text & "'", connEmp, adOpenKeyset, adLockReadOnly, adCmdText
If (Text1.Text = Empty) Then
MsgBox "ROLL_NOT_FOUND", vbRetryCancel, "WARNING"
ElseIf rsEmp.RecordCount <> 0 Then
MsgBox "Employee Code Already Exists !"
rsEmp.Close
Set rsEmp = Nothing
Exit Sub
Else
Set rsEmp = New ADODB.Recordset
rsEmp.Open "select * from SCHEME where NET = '" & Text1.Text & "'", connEmp, adOpenKeyset, adLockPessimistic, adCmdText
rsEmp.AddNew
rsEmp!NET = Trim(Text1.Text)
rsEmp.Update
connEmp.Execute "commit"
rsEmp.Close
Set rsEmp = Nothing
MsgBox "Added Succesfully !"
Form1.Command8_Click
End If
End Sub

Private Sub Command2_Click()
Form11.Hide
Form9.Show
End Sub

Private Sub Command3_Click()
If (Text1.Text = Empty) Then
MsgBox "ROLL_NOT_FOUND", vbRetryCancel, "WARNING"
ElseIf (MsgBox("Are you sure to delete ?", vbYesNo) = vbYes) Then
Set rsEmp = New ADODB.Recordset
rsEmp.Open "select * from SCHEME where NET = '" & Text1.Text & "'", connEmp, adOpenKeyset, adLockPessimistic, adCmdText
rsEmp.Delete
connEmp.Execute "commit"
rsEmp.Close
Set rsEmp = Nothing
MsgBox "Deleted Succesfully !"
Form1.Command8_Click
End If
End Sub

Private Sub Command4_Click()
Form11.Hide
Form10.Show
End Sub

Private Sub Form_Load()
Set connEmp = New ADODB.Connection
connEmp.Open "Provider=OraOLEDB.Oracle.1;Password=sa;Persist Security Info=True;User ID=system;Data Source=SARAVANAN"
connEmp.CursorLocation = adUseClient
Form1.BackColor = vbYellow
End Sub

Private Sub Picture1_Click()
End
End Sub
