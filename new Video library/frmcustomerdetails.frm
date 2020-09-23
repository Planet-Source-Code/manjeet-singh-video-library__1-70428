VERSION 5.00
Begin VB.Form frmcustomerdetails 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Customer Details Form"
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7575
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "frmcustomerdetails.frx":0000
   ScaleHeight     =   5910
   ScaleWidth      =   7575
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C0C0C0&
      Caption         =   "E&XIT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Exit the window"
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&DELETE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "To delete a record"
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton cmdfind 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&FIND"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "To find some record"
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton cmdmodify 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&MODIFY"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "To modify the record"
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&ADD"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "To add a record"
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton cmdsave 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&SAVE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "To save the modified record"
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton cmdclear 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&CLEAR"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "To clear all text boxes"
      Top             =   3240
      Width           =   1455
   End
   Begin VB.TextBox txtemailID 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      MaxLength       =   30
      TabIndex        =   13
      Top             =   3840
      Width           =   1815
   End
   Begin VB.TextBox txtmobno 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      MaxLength       =   10
      TabIndex        =   12
      Top             =   3240
      Width           =   1815
   End
   Begin VB.TextBox txtphoneno 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      MaxLength       =   8
      TabIndex        =   11
      Top             =   2640
      Width           =   1815
   End
   Begin VB.TextBox txtaddress 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      MaxLength       =   30
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox txtname 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      MaxLength       =   20
      TabIndex        =   9
      Top             =   1440
      Width           =   1785
   End
   Begin VB.TextBox txtcustid 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2400
      MaxLength       =   3
      TabIndex        =   8
      Top             =   840
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   3840
      Picture         =   "frmcustomerdetails.frx":28D34
      ScaleHeight     =   3495
      ScaleWidth      =   2295
      TabIndex        =   7
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CUSTOMER DETAILS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   6
      Top             =   240
      Width           =   3735
   End
   Begin VB.Label lblcustid 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer ID         :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label lblcustname 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name   :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label lblcustadd 
      BackStyle       =   0  'Transparent
      Caption         =   "Address                :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label lblcustphoneno 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone No.            :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label lblcustmobno 
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile No.          :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label lblcustemail 
      BackStyle       =   0  'Transparent
      Caption         =   "Email ID              :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   3960
      Width           =   975
   End
End
Attribute VB_Name = "frmcustomerdetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const str As String = "select * from customer"
Dim str1 As String
Dim con As New Connection
Dim rs As New Recordset
Dim estatus As Boolean
Dim ins As String
Dim j As Integer


Private Sub cmdAdd_Click()

If Validate = True Then
con.BeginTrans
ins = "insert into Customer(CUSTID,CUSTNAME,CUSTADD,CUSTPHONENO,CUSTMOBNO,CUSTEMAIL)"
ins = ins & "values(" & txtcustid & "," & " '" & txtname & "'" & "," & "'" & txtaddress & "'" _
& "," & txtphoneno & "," & txtmobno & "," & "'" & txtemailID & "'" & ")"
Debug.Print ins
If estatus = False Then
con.Execute ins
End If
con.CommitTrans
MsgBox "Your Data has been saved.."
Call clearfields
'txtcustid.SetFocus
End If
txtcustid.Enabled = True
rs.MoveLast
txtcustid = rs!custid + 1
a = txtcustid + 1
txtcustid = a
txtcustid.Enabled = False

End Sub

Private Sub cmdclear_Click()
Call clearfields
rs.MoveLast
txtcustid = rs!custid + 1


End Sub

Private Sub cmdDelete_Click()
Dim yes As String
Dim k As Integer
b = txtcustid
k = InputBox("Enter the record no you want to delete..", "Delete..")
rs.MoveFirst
While Not rs.EOF
If rs!custid = Val(k) Then
yes = MsgBox("Are you sure you want to delete this record ..", vbYesNo)
If yes = vbYes Then
rs.Delete
a = txtcustid - 1
txtcustid = a
If rs.EOF = True Then
rs.MoveLast
End If
End If
End If
rs.MoveNext
Wend
If rs.EOF Then
MsgBox "Record Does not Exist.."
End If
End Sub

Private Sub cmdExit_Click()
MDIForm1.Show
frmcustomerdetails.Hide
End Sub

Private Sub cmdfind_Click()
Dim i As Integer
Dim flag As Boolean
i = InputBox("Enter the CustomerID you want to search..", "Find")
rs.MoveFirst
While Not rs.EOF
If rs!custid = Val(i) Then
txtcustid.Enabled = False
Call display
flag = True
End If
rs.MoveNext
Wend
If flag = False Then
MsgBox "No Record Exist..", vbCritical
End If
cmddelete.Enabled = True
cmdmodify.Enabled = True

End Sub

Public Sub display()
txtcustid.Text = rs!custid
txtname.Text = rs!custname
txtaddress.Text = rs!CUSTADD
txtphoneno.Text = rs!CUSTPHONENO
txtmobno.Text = rs!custmobno
txtemailID.Text = rs!CUSTEMAIL
End Sub

Private Sub cmdmodify_Click()

j = InputBox("Enter the CustomerID you want to Modify..", "Modify")
rs.MoveFirst
While Not rs.EOF
If rs!custid = Val(j) Then
txtcustid.Enabled = False
Call display
End If
rs.MoveNext
Wend

txtcustid.Enabled = True

End Sub

Private Sub cmdsave_Click()
If Validate = True Then
con.BeginTrans
ins = "update Customer set custname = " & "'" & txtname & "'" & "," & "custadd = " & "'" & txtaddress & "'" & "," _
& "custphoneno = " & txtphoneno & "," & "custmobno = " & txtmobno & "," & " custemail = " & "'" & txtemailID & "'" & " where custid = " & Val(j)
Debug.Print ins
If estatus = False Then
con.Execute ins
End If
con.CommitTrans
MsgBox "Your Data has been Modified.."
End If
Call clearfields
rs.MoveLast
txtcustid = rs!custid + 1
'a = txtcustid + 1
'txtcustid = a

End Sub

Private Sub Form_Load()

con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Videolibrary.mdb;Persist Security Info=False"
rs.Open str, con, adOpenDynamic, adLockOptimistic

txtemailID.Text = "@"
txtemailID.SelStart = 0
rs.MoveLast
txtcustid = rs!custid + 1
txtcustid.Enabled = False
End Sub

Public Sub clearfields()
txtcustid.Text = ""
txtname.Text = ""
txtaddress.Text = ""
txtphoneno.Text = ""
txtmobno.Text = ""
txtemailID.Text = ""
txtemailID.Text = "@"
End Sub



Public Function Validate()
Dim flag As Boolean
If Not IsNumeric(txtcustid) And txtcustid = "" Then
MsgBox "Please enter Customer ID and it should be Numeric..", vbInformation
ElseIf txtname = "" Then
MsgBox "Please enter Customer name..", vbInformation
ElseIf txtaddress = "" Then
MsgBox "Please enter your Address..", vbInformation
ElseIf txtphoneno = "" And Not IsNumeric(txtphoneno) Then
MsgBox "Please enter Phone No..", vbInformation
ElseIf txtmobno = "" And Not IsNumeric(txtmobno) Then
MsgBox "Please enter Mobile No & it should be Numeric..", vbInformation
ElseIf txtemailID = "" Then
MsgBox "Please enter your Email ID..", vbInformation
Else
flag = True
End If
Validate = flag
End Function

Private Sub txtemailID_GotFocus()
txtemailID.SelStart = 0
End Sub
