VERSION 5.00
Begin VB.Form Form12 
   Caption         =   "Form12"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form12"
   Picture         =   "Form12.frx":0000
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command6 
      Caption         =   "BACK"
      Height          =   495
      Left            =   12960
      TabIndex        =   18
      Top             =   7560
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      DragMode        =   1  'Automatic
      Height          =   375
      Left            =   19320
      Picture         =   "Form12.frx":7BA37
      ScaleHeight     =   315
      ScaleWidth      =   915
      TabIndex        =   17
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      Caption         =   "MAIN  MENU"
      Height          =   495
      Left            =   15360
      TabIndex        =   16
      Top             =   7560
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "CLEAR"
      Height          =   495
      Left            =   10680
      TabIndex        =   14
      Top             =   7560
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "VIEW"
      Height          =   495
      Left            =   8520
      TabIndex        =   13
      Top             =   7560
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "UPDATE"
      Height          =   495
      Left            =   6120
      TabIndex        =   12
      Top             =   7560
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "DELETE"
      Height          =   495
      Left            =   3960
      TabIndex        =   11
      Top             =   7560
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   9360
      TabIndex        =   5
      Top             =   6240
      Width           =   2175
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   9360
      TabIndex        =   4
      Top             =   5040
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   9360
      TabIndex        =   3
      Top             =   3840
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   9360
      TabIndex        =   2
      Top             =   2640
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   9360
      TabIndex        =   1
      Top             =   1560
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ADD"
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   7560
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "MESSAGE  PACK"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6360
      TabIndex        =   15
      Top             =   240
      Width           =   6015
   End
   Begin VB.Label Label5 
      Caption         =   "DESCRIPTION"
      Height          =   255
      Left            =   6720
      TabIndex        =   10
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "VALIDITY"
      Height          =   255
      Left            =   6720
      TabIndex        =   9
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "AMOUNT"
      Height          =   255
      Left            =   6720
      TabIndex        =   8
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "NETWORK"
      Height          =   375
      Left            =   6720
      TabIndex        =   7
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "ID"
      Height          =   375
      Left            =   6720
      TabIndex        =   6
      Top             =   1560
      Width           =   1575
   End
End
Attribute VB_Name = "Form12"
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
rsEmp.Open "select * from MSG where ID = '" & Text1.Text & "'", connEmp, adOpenKeyset, adLockReadOnly, adCmdText
If (Text1.Text = Empty) Then
MsgBox "ROLL_NOT_FOUND", vbRetryCancel, "WARNING"
ElseIf rsEmp.RecordCount <> 0 Then
MsgBox "Employee Code Already Exists !"
rsEmp.Close
Set rsEmp = Nothing
Exit Sub
Else
Set rsEmp = New ADODB.Recordset
rsEmp.Open "select * from MSG where ID = '" & Text1.Text & "'", connEmp, adOpenKeyset, adLockPessimistic, adCmdText
rsEmp.AddNew
rsEmp!ID = Trim(Text1.Text)
rsEmp!NET = Trim(Text2.Text)
rsEmp!amt = Trim(Text3.Text)
rsEmp!validity = Trim(Text4.Text)
rsEmp!des = Trim(Text5.Text)
rsEmp.Update
connEmp.Execute "commit"
rsEmp.Close
Set rsEmp = Nothing
MsgBox "Added Succesfully !"
Command5_Click
End If
End Sub

Private Sub Command2_Click()
If (Text1.Text = Empty) Then
MsgBox "ROLL_NOT_FOUND", vbRetryCancel, "WARNING"
ElseIf (MsgBox("Are you sure to delete ?", vbYesNo) = vbYes) Then
Set rsEmp = New ADODB.Recordset
rsEmp.Open "select * from MSG where ID = '" & Text1.Text & "'", connEmp, adOpenKeyset, adLockPessimistic, adCmdText
rsEmp.Delete
connEmp.Execute "commit"
rsEmp.Close
Set rsEmp = Nothing
MsgBox "Deleted Succesfully !"
End If
Command5_Click
End Sub

Private Sub Command3_Click()
If (Text1.Text = Empty) Then
MsgBox "ROLL_NOT_FOUND", vbRetryCancel, "WARNING"
ElseIf (MsgBox("Are you sure to edit ?", vbYesNo) = vbYes) Then
Set rsEmp = New ADODB.Recordset
rsEmp.Open "select * from MSG where ID = '" & Text1.Text & "'", connEmp, adOpenKeyset, adLockPessimistic, adCmdText
rsEmp!ID = Trim(Text1.Text)
rsEmp!NET = Trim(Text2.Text)
rsEmp!amt = Trim(Text3.Text)
rsEmp!validity = Trim(Text4.Text)
rsEmp!des = Trim(Text5.Text)
rsEmp.Update
connEmp.Execute "commit"
rsEmp.Close
Set rsEmp = Nothing
MsgBox "Edited Succesfully !"
Command5_Click
End If
End Sub

Private Sub Command4_Click()
Set rsEmp = New ADODB.Recordset
rsEmp.Open "select * from MSG where id = '" & Text1.Text & "'", connEmp, adOpenKeyset, adLockReadOnly, adCmdText
If (Text1.Text = Empty) Then
MsgBox "ROLL_NOT_FOUND", vbRetryCancel, "WARNING"
ElseIf rsEmp.RecordCount <> 0 Then
Text2.Text = rsEmp!NET
Text3.Text = rsEmp!amt
Text4.Text = rsEmp!validity
Text5.Text = rsEmp!des
MsgBox "Viewed Succesfully !"
Else
MsgBox "Message pack Id Not Exists !"
End If
rsEmp.Close
Set rsEmp = Nothing
End Sub

Private Sub Command5_Click()
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
End Sub

Private Sub Command6_Click()
Form12.Hide
Form10.Show
End Sub

Private Sub Command7_Click()
Form12.Hide
Form9.Show
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
