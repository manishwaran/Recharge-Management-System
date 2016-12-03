VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "Form7"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form7"
   Picture         =   "Form7.frx":0000
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      DragMode        =   1  'Automatic
      Height          =   375
      Left            =   19320
      Picture         =   "Form7.frx":945AA
      ScaleHeight     =   315
      ScaleWidth      =   915
      TabIndex        =   6
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SIGN  UP"
      Height          =   495
      Left            =   6000
      TabIndex        =   5
      Top             =   4560
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "LOGIN"
      Height          =   495
      Left            =   6000
      TabIndex        =   4
      Top             =   3600
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   6000
      TabIndex        =   1
      Top             =   2520
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   6000
      TabIndex        =   0
      Top             =   1320
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "PASSWORD"
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "USER  NAME"
      Height          =   615
      Left            =   2280
      TabIndex        =   2
      Top             =   1440
      Width           =   1695
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim connEmp As ADODB.Connection
Dim rsEmp As ADODB.Recordset

Private Sub Command1_Click()
Set rsEmp = New ADODB.Recordset
If (Text1.Text = Empty Or Text2.Text = Empty) Then
MsgBox "ROLL_NOT_FOUND", vbRetryCancel, "WARNING"
Else
    rsEmp.Open "select password from login where username='" & Text1.Text & "'", connEmp, adOpenKeyset, adLockReadOnly, adCmdText
    rsEmp.MoveFirst
    If (rsEmp!Password = Text2.Text) Then
        MsgBox "Login success!"
        Form7.Hide
        Form9.Show
        rsEmp.Close
    Else
        MsgBox (" INVALID PASSWORD OR USER NAME ")
    End If
End If
End Sub

Private Sub Command2_Click()
Set rsEmp = New ADODB.Recordset
rsEmp.Open "select * from LOGIN where PASSWORD= '" & Text2.Text & "'", connEmp, adOpenKeyset, adLockReadOnly, adCmdText
If (Text1.Text = Empty Or Text2.Text = Empty) Then
MsgBox "USERNAME OR PASSWORD CANNOT BE EMPTY !", vbRetryCancel, "WARNING"
ElseIf rsEmp.RecordCount <> 0 Then
MsgBox "ACCOUNT ALREADY PRESENT !"
rsEmp.Close
Set rsEmp = Nothing
ElseIf rsEmp.RecordCount = 0 Then
Set rsEmp = New ADODB.Recordset
rsEmp.Open "select * from LOGIN where PASSWORD = '" & Text2.Text & "'", connEmp, adOpenKeyset, adLockPessimistic, adCmdText
rsEmp.AddNew
rsEmp!UserName = Trim(Text1.Text)
rsEmp!Password = Trim(Text2.Text)
rsEmp.Update
connEmp.Execute "commit"
rsEmp.Close
Set rsEmp = Nothing
MsgBox "SIGNED UP SUCCESSFULLY !"
Form7.Hide
Form9.Show
Else
MsgBox ("ACCOUNT NOT FOUND !")
End If
End Sub

Public Sub Form_Load()
Set connEmp = New ADODB.Connection
connEmp.Open "Provider=OraOLEDB.Oracle.1;Password=sa;Persist Security Info=True;User ID=system;Data Source=SARAVANAN"
connEmp.CursorLocation = adUseClient
End Sub


Private Sub Picture1_Click()
End
End Sub
