VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      DragMode        =   1  'Automatic
      Height          =   375
      Left            =   19320
      Picture         =   "Form1.frx":106989
      ScaleHeight     =   315
      ScaleWidth      =   915
      TabIndex        =   24
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton Command9 
      Caption         =   "MAIN MENU"
      Height          =   735
      Left            =   12720
      TabIndex        =   23
      Top             =   9600
      Width           =   2415
   End
   Begin VB.CommandButton Command8 
      Caption         =   "CLEAR"
      Height          =   735
      Left            =   12720
      TabIndex        =   22
      Top             =   7800
      Width           =   2415
   End
   Begin VB.CommandButton Command6 
      Caption         =   "UPDATE"
      Height          =   735
      Left            =   12720
      TabIndex        =   21
      Top             =   6000
      Width           =   2415
   End
   Begin VB.CommandButton Command5 
      Caption         =   "FETCH"
      Height          =   735
      Left            =   12720
      TabIndex        =   20
      Top             =   4680
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      Caption         =   "DELETE"
      Height          =   735
      Left            =   12720
      TabIndex        =   19
      Top             =   2760
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "INSERT"
      Height          =   735
      Left            =   12720
      TabIndex        =   18
      Top             =   720
      Width           =   2415
   End
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   6240
      TabIndex        =   16
      Top             =   3840
      Width           =   2895
   End
   Begin VB.Frame Frame2 
      Caption         =   "CUSTOMER"
      Height          =   4455
      Left            =   1920
      TabIndex        =   8
      Top             =   5880
      Width           =   9495
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   4080
         TabIndex        =   13
         Top             =   2880
         Width           =   2415
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   4080
         TabIndex        =   12
         Top             =   1800
         Width           =   2415
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   4080
         TabIndex        =   10
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label5 
         Caption         =   "CITY"
         Height          =   375
         Left            =   1440
         TabIndex        =   14
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "CUSTOMER  NAME"
         Height          =   375
         Left            =   1440
         TabIndex        =   11
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "CUTOMER  ID"
         Height          =   495
         Left            =   1440
         TabIndex        =   9
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.OptionButton Option2 
      Caption         =   "POSTPAID"
      Height          =   495
      Left            =   8040
      TabIndex        =   4
      Top             =   2520
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "PREPAID"
      Height          =   495
      Left            =   6240
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "RECHARGE"
      Height          =   4815
      Left            =   1920
      TabIndex        =   0
      Top             =   600
      Width           =   9495
      Begin VB.CommandButton Command10 
         Caption         =   "CHECK"
         Height          =   495
         Left            =   7800
         TabIndex        =   25
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "BROWSE"
         Height          =   375
         Left            =   5160
         TabIndex        =   7
         Top             =   4200
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   4320
         TabIndex        =   6
         Top             =   480
         Width           =   2895
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Form1.frx":11016E
         Left            =   4320
         List            =   "Form1.frx":11017E
         TabIndex        =   5
         Text            =   "SELECT NETWORK"
         Top             =   2640
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   4320
         TabIndex        =   2
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Label Label7 
         Caption         =   "AMOUNT"
         Height          =   375
         Left            =   1440
         TabIndex        =   17
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "PHONE  NUMBER"
         Height          =   255
         Left            =   1440
         TabIndex        =   15
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "RECHARGE  ID"
         Height          =   375
         Left            =   1440
         TabIndex        =   1
         Top             =   1440
         Width           =   2055
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim connEmp As ADODB.Connection
Dim rsEmp As ADODB.Recordset
Dim rsemp1 As ADODB.Recordset
Dim COMM As ADODB.Command
Dim CM As ADODB.Command
Private Sub Command1_Click()
If (Combo1.Text = "SELECT NETWORK") Then
    MsgBox ("SELECT A NETWORK !")
Else
    Form1.Hide
    Form6.Show
End If
End Sub

Private Sub Command10_Click()
Set rsEmp = New ADODB.Recordset
rsEmp.Open "select MAX(R_ID) from recharge", connEmp, adOpenKeyset, adLockReadOnly, adCmdText
Text1.Text = Trim(rsEmp(0)) + 1
rsEmp.Close
Set rsEmp = Nothing
Set rsemp1 = New ADODB.Recordset
rsemp1.Open "select distinct(c_id),c_name,city,model,net from recharge where ph_no = '" & Text2.Text & "'", connEmp, adOpenKeyset, adLockReadOnly, adCmdText
If (Text2.Text = Empty) Then
    Command8_Click
    MsgBox ("Enter the phone number ! ")
ElseIf (rsemp1.RecordCount <> 0) Then
    Text3.Text = Trim(rsemp1!c_id)
    Text4.Text = Trim(rsemp1(1))
    Text5.Text = Trim(rsemp1(2))
    Combo1.Text = Trim(rsemp1(4))
    If Trim(rsemp1!MODEL) = "POSTPAID" Then
    Option2.Value = True
    Else
    Option1.Value = True
    End If
End If
rsemp1.Close
Set rsemp1 = Nothing
End Sub

Private Sub Command2_Click()
Set COMM = New ADODB.Command
Set CM = New ADODB.Command
CM.CommandType = adCmdText
CM.CommandText = "drop table recharge"
CM.ActiveConnection = connEmp
CM.Execute

COMM.CommandType = adCmdText
COMM.CommandText = "CREATE TABLE RECHARGE(R_ID INT,C_ID INT,NET varchar(30),MODEL VARCHAR(30),C_NAME varchar(30),CITY varchar(30),PH_NO INT,AMT varchar(30),PRIMARY KEY(R_ID))"
COMM.ActiveConnection = connEmp
COMM.Execute
MsgBox ("TABLE CREATED !")
End Sub

Private Sub Command3_Click()
Set rsEmp = New ADODB.Recordset
rsEmp.Open "select * from recharge where r_id = '" & Text1.Text & "'", connEmp, adOpenKeyset, adLockReadOnly, adCmdText
If (Text1.Text = Empty) Then
MsgBox "ROLL_NOT_FOUND", vbRetryCancel, "WARNING"
ElseIf rsEmp.RecordCount <> 0 Then
MsgBox "Employee Code Already Exists !"
rsEmp.Close
Set rsEmp = Nothing
Exit Sub
Else
Set rsEmp = New ADODB.Recordset
rsEmp.Open "select * from recharge where r_id = '" & Text1.Text & "'", connEmp, adOpenKeyset, adLockPessimistic, adCmdText
rsEmp.AddNew
rsEmp!R_ID = Trim(Text1.Text)
rsEmp!c_id = Trim(Text3.Text)
If Option1.Value = True Then
rsEmp!MODEL = "PREPAID"
ElseIf Option2.Value = True Then
rsEmp!MODEL = "POSTPAID"
End If
rsEmp!C_NAME = Trim(Text4.Text)
rsEmp!CITY = Trim(Text5.Text)
rsEmp!PH_NO = Trim(Text2.Text)
rsEmp!amt = Trim(Text6.Text)
rsEmp!NET = Trim(Combo1.Text)
rsEmp.Update
connEmp.Execute "commit"
rsEmp.Close
Set rsEmp = Nothing
MsgBox "Added Succesfully..."
Command8_Click
End If
End Sub

Private Sub Command4_Click()
If (Text1.Text = Empty) Then
MsgBox "ROLL_NOT_FOUND", vbRetryCancel, "WARNING"
ElseIf (MsgBox("Are you sure to delete...", vbYesNo) = vbYes) Then
Set rsEmp = New ADODB.Recordset
rsEmp.Open "select * from RECHARGE where R_ID = '" & Text1.Text & "'", connEmp, adOpenKeyset, adLockPessimistic, adCmdText
rsEmp.Delete
connEmp.Execute "commit"
rsEmp.Close
Set rsEmp = Nothing
MsgBox "Deleted Succesfully..."
Command8_Click
End If
End Sub

Private Sub Command5_Click()
Set rsEmp = New ADODB.Recordset
rsEmp.Open "select * from RECHARGE where r_id = '" & Text1.Text & "'", connEmp, adOpenKeyset, adLockReadOnly, adCmdText
If (Text1.Text = Empty) Then
MsgBox "ROLL_NOT_FOUND", vbRetryCancel, "WARNING"
ElseIf rsEmp.RecordCount <> 0 Then
Text2.Text = Trim(rsEmp!PH_NO)
If Trim(rsEmp!MODEL) = "POSTPAID" Then
Option2.Value = True
Else
Option1.Value = True
End If
Text3.Text = Trim(rsEmp!c_id)
Text4.Text = Trim(rsEmp!C_NAME)
Text5.Text = Trim(rsEmp!CITY)
Text6.Text = Trim(rsEmp!amt)
Combo1.Text = Trim(rsEmp!NET)
MsgBox "Viewed Succesfully !"
Else
MsgBox "Recharge Id Not Exists !"
End If
rsEmp.Close
Set rsEmp = Nothing
End Sub

Private Sub Command6_Click()
If (Text1.Text = Empty) Then
MsgBox "ROLL_NOT_FOUND", vbRetryCancel, "WARNING"
ElseIf (MsgBox("Are you sure to edit ?", vbYesNo) = vbYes) Then
Set rsEmp = New ADODB.Recordset
rsEmp.Open "select * from RECHARGE where R_ID = '" & Text1.Text & "'", connEmp, adOpenKeyset, adLockPessimistic, adCmdText
rsEmp!R_ID = Trim(Text1.Text)
rsEmp!PH_NO = Trim(Text2.Text)
If Option1.Value = True Then
rsEmp!MODEL = "PREPAID"
ElseIf Option2.Value = True Then
rsEmp!Gender = "POSTPAID"
End If
rsEmp!c_id = Trim(Text3.Text)
rsEmp!C_NAME = Trim(Text4.Text)
rsEmp!CITY = Trim(Text5.Text)
rsEmp!amt = Trim(Text6.Text)
rsEmp!NET = Trim(Combo1.Text)
rsEmp.Update
connEmp.Execute "commit"
rsEmp.Close
Set rsEmp = Nothing
MsgBox "Edited Succesfully !"
Command8_Click
End If
End Sub

Public Sub Command7_Click()
Form1.Hide
Form9.Show
End Sub

Public Sub Command8_Click()
Text1.Text = ""
Text2.Text = ""
Option1.Value = False
Option2.Value = False
Combo1.Text = "SELECT NETWORK"
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
End Sub

Private Sub Command9_Click()
Form1.Hide
Form9.Show
End Sub

Private Sub Form_Load()
Form1.Hide
Form8.Show
Set connEmp = New ADODB.Connection
connEmp.Open "Provider=OraOLEDB.Oracle.1;Password=sa;Persist Security Info=True;User ID=system;Data Source=SARAVANAN"
connEmp.CursorLocation = adUseClient
Form1.BackColor = vbYellow
End Sub

Private Sub Picture1_Click()
End
End Sub

