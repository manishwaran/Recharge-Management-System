VERSION 5.00
Begin VB.Form Form9 
   Caption         =   "Form9"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form9"
   Picture         =   "Form9.frx":0000
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "SIGN  OUT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5880
      TabIndex        =   4
      Top             =   7080
      Width           =   5055
   End
   Begin VB.PictureBox Picture1 
      DragMode        =   1  'Automatic
      Height          =   375
      Left            =   19320
      Picture         =   "Form9.frx":3588C
      ScaleHeight     =   315
      ScaleWidth      =   915
      TabIndex        =   3
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SCHEME  MANAGEMENT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5880
      TabIndex        =   1
      Top             =   5040
      Width           =   5055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "RECHARGE  MANAGEMENT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5880
      TabIndex        =   0
      Top             =   3000
      Width           =   5055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "MAIN MENU"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   855
      Left            =   4920
      TabIndex        =   2
      Top             =   840
      Width           =   6615
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim connEmp As ADODB.Connection
Dim rsEmp As ADODB.Recordset
Dim rsemp1 As ADODB.Recordset
Dim COMM As ADODB.Command
Dim CM As ADODB.Command
Dim CM1 As ADODB.Command
Dim CM2 As ADODB.Command
Dim CM3 As ADODB.Command
Dim CM4 As ADODB.Command
Dim CM5 As ADODB.Command

Private Sub create_click()

Set COMM = New ADODB.Command
Set CM = New ADODB.Command
CM.CommandType = adCmdText
CM.CommandText = "CREATE TABLE RECHARGE(R_ID INT,C_ID INT,NET varchar(30),MODEL VARCHAR(30),C_NAME varchar(30),CITY varchar(30),PH_NO INT,AMT varchar(30),PRIMARY KEY(R_ID))"
CM.ActiveConnection = connEmp
CM.Execute

Set CM1 = New ADODB.Command
CM1.CommandType = adCmdText
CM1.CommandText = "create table scheme(net varchar(50),primary key(net))"
CM1.ActiveConnection = connEmp
CM1.Execute

Set CM2 = New ADODB.Command
CM2.CommandType = adCmdText
CM2.CommandText = "create table msg(net varchar(50),id int,amt int,validity int,des varchar(50),primary key(id))"
CM2.ActiveConnection = connEmp
CM2.Execute

Set CM3 = New ADODB.Command
CM3.CommandType = adCmdText
CM3.CommandText = "create table top(net varchar(50),id int,amt int,validity int,des varchar(50),primary key(id))"
CM3.ActiveConnection = connEmp
CM3.Execute

Set CM4 = New ADODB.Command
CM4.CommandType = adCmdText
CM4.CommandText = "create table spl(net varchar(50),id int,amt int,validity int,des varchar(50),primary key(id))"
CM4.ActiveConnection = connEmp
CM4.Execute

Set CM5 = New ADODB.Command
CM5.CommandType = adCmdText
CM5.CommandText = "create table netp(net varchar(50),id int,amt int,validity int,des varchar(50),primary key(id))"
CM5.ActiveConnection = connEmp
CM5.Execute


End Sub
Private Sub Command1_Click()
Form9.Hide
Form1.Show
End Sub

Private Sub Command2_Click()
Form9.Hide
Form10.Show
End Sub

Private Sub Command3_Click()
Form9.Hide
Form7.Text1.Text = ""
Form7.Text2.Text = ""
Form7.Show
MsgBox ("SIGNED OUT SUCCESSFULLY !")

End Sub

Private Sub Command4_Click()
ElseIf (MsgBox("Are you sure to delete...", vbYesNo) = vbYes) Then
Set COMM = New ADODB.Command
Set CM = New ADODB.Command
CM.CommandType = adCmdText
CM.CommandText = "drop table recharge"
CM.ActiveConnection = connEmp
CM.Execute

Set COMM = New ADODB.Command
Set CM1 = New ADODB.Command
CM1.CommandType = adCmdText
CM1.CommandText = "drop table scheme"
CM1.ActiveConnection = connEmp
CM1.Execute

Set COMM = New ADODB.Command
Set CM2 = New ADODB.Command
CM2.CommandType = adCmdText
CM2.CommandText = "drop table msg"
CM2.ActiveConnection = connEmp
CM2.Execute

Set COMM = New ADODB.Command
Set CM3 = New ADODB.Command
CM3.CommandType = adCmdText
CM3.CommandText = "drop table top"
CM3.ActiveConnection = connEmp
CM3.Execute

Set COMM = New ADODB.Command
Set CM4 = New ADODB.Command
CM4.CommandType = adCmdText
CM4.CommandText = "drop table netp"
CM4.ActiveConnection = connEmp
CM4.Execute

Set COMM = New ADODB.Command
Set CM5 = New ADODB.Command
CM5.CommandType = adCmdText
CM5.CommandText = "drop table spl"
CM5.ActiveConnection = connEmp
CM5.Execute

create_click
End If
End Sub

Private Sub Form_Load()

Set connEmp = New ADODB.Connection
connEmp.Open "Provider=OraOLEDB.Oracle.1;Password=sa;Persist Security Info=True;User ID=system;Data Source=SARAVANAN"
connEmp.CursorLocation = adUseClient

End Sub


Private Sub Picture1_Click()
End
End Sub
