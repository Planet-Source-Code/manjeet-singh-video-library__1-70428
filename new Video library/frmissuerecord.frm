VERSION 5.00
Begin VB.Form frmissuerecord 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Issue Record"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7245
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   7245
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtdiskqty 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   3960
      TabIndex        =   26
      Top             =   4560
      Width           =   1095
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
      Height          =   345
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "To find some record"
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox txtissueno 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   3960
      TabIndex        =   0
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox txtcustomerid 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   3960
      TabIndex        =   1
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox txtdiskno 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   3960
      TabIndex        =   2
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox txtissuedate 
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "M/d/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   3960
      TabIndex        =   3
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox txtreturndate 
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "M/d/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   3960
      TabIndex        =   4
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox txtadvancepaid 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   3960
      TabIndex        =   5
      Top             =   3120
      Width           =   1110
   End
   Begin VB.TextBox txtbalanceamount 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   3960
      TabIndex        =   6
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox txtnoofdays 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   3960
      TabIndex        =   7
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdadd 
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
      Height          =   345
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "To add a new record"
      Top             =   600
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
      Height          =   345
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "To modify records"
      Top             =   1560
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
      Height          =   345
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "To clear all text boxes"
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton cmdexit 
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
      Height          =   345
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Exit this window"
      Top             =   3480
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
      Height          =   345
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "To save the modified data"
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton cmddelete 
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
      Height          =   345
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "To delete a record"
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label lbldiskqty 
      BackStyle       =   0  'Transparent
      Caption         =   "Disk Quantity"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   240
      TabIndex        =   25
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Note : Disk rent is Rs.30/per day"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   24
      Top             =   5520
      Width           =   7095
   End
   Begin VB.Label lblissue 
      BackStyle       =   0  'Transparent
      Caption         =   "Issue No"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   240
      TabIndex        =   23
      Top             =   720
      Width           =   3975
   End
   Begin VB.Label lblcust_id 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer ID"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   240
      TabIndex        =   22
      Top             =   1200
      Width           =   3975
   End
   Begin VB.Label lbldiskno 
      BackStyle       =   0  'Transparent
      Caption         =   "Disk No"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   240
      TabIndex        =   21
      Top             =   1680
      Width           =   3975
   End
   Begin VB.Label lblissuedate 
      BackStyle       =   0  'Transparent
      Caption         =   "Issue Date(MM/DD/YYYY)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   240
      TabIndex        =   20
      Top             =   2160
      Width           =   3615
   End
   Begin VB.Label lblreturndate 
      BackStyle       =   0  'Transparent
      Caption         =   "Return Date(MM/DD/YYYY)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   240
      TabIndex        =   19
      Top             =   2640
      Width           =   3735
   End
   Begin VB.Label lbladvpaid 
      BackStyle       =   0  'Transparent
      Caption         =   "Advance Paid"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   240
      TabIndex        =   18
      Top             =   3120
      Width           =   3615
   End
   Begin VB.Label lblbalamt 
      BackStyle       =   0  'Transparent
      Caption         =   "Balance Amount"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   240
      TabIndex        =   17
      Top             =   3600
      Width           =   3615
   End
   Begin VB.Label lblnoofdays 
      BackStyle       =   0  'Transparent
      Caption         =   "Disk was Issued for"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   240
      TabIndex        =   16
      Top             =   4080
      Width           =   3015
   End
   Begin VB.Label lblissuerecord 
      BackStyle       =   0  'Transparent
      Caption         =   "Issue Record"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   2640
      TabIndex        =   15
      Top             =   0
      Width           =   2535
   End
End
Attribute VB_Name = "frmissuerecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const str As String = "select * from issuerec"
 Const str1 As String = "select diskqty from DISC"
  Dim con As New Connection
Dim rs As New Recordset
Dim rs1 As New Recordset
Dim estatus As Boolean
Dim ins1 As String
Dim ins As String
Dim i As Integer
Dim j As Integer
Dim f As Integer

Private Sub cmdAdd_Click()
rs1.MoveLast
a = rs1!diskqty - 1
txtdiskqty.Text = a
If addStatus = True Then
con.BeginTrans
ins = "insert into Issuerec(issueno,custid,DISKNO,issuedt,advpaid,diskqty)"
ins = ins & "values(" & txtissueno & "," & txtcustomerid & "," & txtdiskno _
& "," & "'" & txtissuedate & "'" & "," & txtadvancepaid & "," & txtdiskqty.Text & ")"

Debug.Print ins
If estatus = False Then
con.Execute ins
End If
con.CommitTrans
'con.BeginTrans
'Dim upd As String
'upd = "update DISC set diskqty = " & a & " where DISKNO = " & txtdiskno.Text
'Debug.Print upd
'con.Execute upd
'con.CommitTrans
MsgBox "Your Data has been saved.."
Call clearfields
'txtcustid.SetFocus
End If

txtissueno.Enabled = True
rs.MoveLast
txtissueno = rs!issueno + 1
txtissueno.Enabled = False

End Sub

Private Sub cmdclear_Click()
Call clearfields
rs.MoveLast
txtissueno = rs!issueno + 1
End Sub

Private Sub cmdDelete_Click()
Dim flag As Boolean
Dim yes As String
Dim k As Long
k = InputBox("Enter the record no you want to delete..", "Delete..")
rs.MoveFirst
While Not rs.EOF
If rs!issueno = Val(k) Then
yes = MsgBox("Are you sure you want to delete this record ..", vbYesNo)
If yes = vbYes Then
rs.Delete
flag = True
If rs.EOF = True Then
rs.MoveLast
End If
End If
End If
rs.MoveNext
Wend
If flag = False Then
MsgBox "Record does not exist ..", vbCritical
End If
End Sub

Private Sub cmdExit_Click()
MDIForm1.Show
frmissuerecord.Hide
End Sub

Public Sub display()
On Error Resume Next
txtissueno.Text = rs!issueno
txtcustomerid.Text = rs!custid
txtdiskno.Text = rs!diskno
txtdiskrate.Text = rs!DISKRT
txtissuedate.Text = rs!issuedt
txtreturndate.Text = rs!returndt
txtadvancepaid.Text = rs!advpaid
txtbalanceamount.Text = rs!balamt
txtnoofdays.Text = rs!noofdays
txtdiskqty.Text = rs!diskqty

End Sub

Private Sub cmdfind_Click()

Dim flag As Boolean
i = InputBox("Enter the Customer ID you want to search..", "Find")
f = i
rs.MoveFirst
While Not rs.EOF
If rs!custid = Val(i) Then
txtcustomerid.Enabled = False
Call display
If txtreturndate.Text <> "" Then
a = txtissuedate.Text
b = txtreturndate.Text
c = Val(txtadvancepaid.Text)
d = DateDiff("y", a, b)
txtnoofdays.Text = d & " Day"
If d <> 0 Then
e = d * 30 - c
txtbalanceamount.Text = "Rs " & e
End If

End If
flag = True
End If
rs.MoveNext
Wend
If flag = False Then
MsgBox "No Record Exist..", vbCritical
End If
cmddelete.Enabled = True
cmdmodify.Enabled = True

'If txtreturndate.Text <> "" Then
'a = txtissuedate.Text
'b = txtreturndate.Text
'c = txtadvancepaid.Text
'd = DateDiff("y", a, b)
'If d <> 0 Then
'e = d * 30 - c
'txtbalanceamount.Text = e
'End If
'Else
'MsgBox "Please Enter the Return date..", vbInformation
'End If
End Sub

Private Sub cmdmodify_Click()
Dim flag As Boolean
j = InputBox("Enter the Customer ID you want to Modify..", "Modify")
f = j
rs.MoveFirst
While Not rs.EOF
    If rs!custid = Val(j) Then
        txtcustomerid.Enabled = False
        txtissueno.Enabled = False
        txtdiskno.Enabled = False
        txtadvancepaid.Enabled = False
        txtdiskqty.Enabled = False
        txtissuedate.Enabled = False
        Call display
        
        flag = True
        If txtreturndate.Text <> "" Then
            a = txtissuedate.Text
            b = txtreturndate.Text
            c = Val(txtadvancepaid.Text)
            d = DateDiff("y", a, b)
            txtnoofdays.Text = d & " Day"
            If d <> 0 Then
                e = d * 30 - c
                txtbalanceamount.Text = "Rs " & e
            End If
            
'Else
'MsgBox "Please Enter the Return date..", vbInformation
        End If
        If txtreturndate.Text = "" Then
            txtreturndate.Text = Date
            End If
    End If
rs.MoveNext
Wend

If flag = False Then
MsgBox "Record does not exist ..", vbCritical
End If


End Sub

Private Sub cmdsave_Click()
If txtreturndate.Text <> "" Then
a = txtissuedate.Text
b = txtreturndate.Text
c = Val(txtadvancepaid.Text)
d = DateDiff("y", a, b)
txtnoofdays.Text = d
If d <> 0 Then
e = d * 30 - c
txtbalanceamount.Text = e
End If
Else
MsgBox "Please Enter the Return date..", vbInformation
End If
rs1.MoveLast
a = rs1!diskqty + 1
txtdiskqty.Text = a
If Validate = True Then
con.BeginTrans
ins = "update Issuerec set custid = " & txtcustomerid & "," & "diskno = " & txtdiskno & "," _
& "issuedt = " & "'" & txtissuedate & "'" & "," & "returndt = " & "'" & txtreturndate & "'" & "," & " advpaid = " & txtadvancepaid & "," & " balamt = " & txtbalanceamount & "," & "noofdays = " & txtnoofdays.Text & "," & "diskqty = " & txtdiskqty.Text & " where issueno = " & Val(f)
Debug.Print ins

If estatus = False Then
con.Execute ins

End If

con.CommitTrans

con.BeginTrans
Dim upd As String
upd = "update DISC set diskqty = " & a & " where DISKNO = " & txtdiskno.Text
Debug.Print upd
con.Execute upd
con.CommitTrans
MsgBox "Your Data has been Modified.."

End If

'Call clearfields

End Sub

Private Sub Form_Load()

con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Videolibrary.mdb;Persist Security Info=False"
rs.Open str, con, adOpenDynamic, adLockOptimistic
rs1.Open str1, con, adOpenDynamic, adLockOptimistic
 
 rs.MoveLast
txtissueno = rs!issueno + 1
txtissueno.Enabled = False


 End Sub

Public Sub clearfields()
txtissueno.Text = ""
txtcustomerid.Text = ""
txtdiskno.Text = ""

txtissuedate.Text = ""
txtreturndate.Text = ""
txtadvancepaid.Text = ""
txtbalanceamount.Text = ""
txtnoofdays.Text = ""
txtdiskqty.Text = ""
txtcustomerid.Enabled = True
txtcustomerid.SetFocus
End Sub

Public Function Validate()

Dim flag As Boolean

If Not IsNumeric(txtissueno) And txtissueno = "" Then
MsgBox "Please enter Issue no and it should be Numeric..", vbInformation

ElseIf txtcustid = "" And Not IsNumeric(txtcustid) Then
MsgBox "Please enter Customer id & it should be Numeric..", vbInformation

ElseIf txtdiskno = "" Then
MsgBox "Please enter your Disk no & it should be Numeric..", vbInformation


ElseIf txtissuedate = "" And Not IsNumeric(txtissuedate) Then
MsgBox "Please enter Issue date & it should be Numeric..", vbInformation

ElseIf txtreturndate = "" And Not IsNumeric(txtreturndate) Then
MsgBox "Please enter Return date & it should be Numeric..", vbInformation

ElseIf txtadvancepaid = "" And Not IsNumeric(txtreturndate) Then
MsgBox "Please enter your Advance paid & it should be Numeric..", vbInformation

ElseIf txtbalanceamount = "" And Not IsNumeric(txtreturndate) Then
MsgBox "Please enter your Balance paid & it should be Numeric..", vbInformation

ElseIf txtbalanceamount.Text <> "" Then
MsgBox "This field should not be filled by you..", vbInformation

ElseIf txtdiskqty.Text <> "" Then
MsgBox "This field should not be filled by you..", vbInf

ElseIf txtnoofdays.Text <> "" Then
MsgBox "This field should not be filled by you..", vbInformation


Else
flag = True

End If

Validate = flag

End Function

Private Sub txtadvancepaid_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii >= Asc("a") _
And KeyAscii <= Asc("z") Then
KeyAscii = 0
End If
End Sub

Private Sub txtbalanceamount_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii >= Asc("a") _
And KeyAscii <= Asc("z") Then
KeyAscii = 0
End If
End Sub

Private Sub txtcustomerid_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii >= Asc("a") _
And KeyAscii <= Asc("z") Then
KeyAscii = 0
End If
End Sub

Private Sub txtdiskno_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii >= Asc("a") _
And KeyAscii <= Asc("z") Then
KeyAscii = 0
End If
End Sub


'Private Sub txtdiskrate_KeyPress(KeyAscii As Integer)
'If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii >= Asc("a") _
'And KeyAscii <= Asc("z") Then
'KeyAscii = 0
'End If
'End Sub

Private Sub txtissuedate_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii >= Asc("a") _
And KeyAscii <= Asc("z") Then
KeyAscii = 0
End If
End Sub

Private Sub txtissueno_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii >= Asc("a") _
And KeyAscii <= Asc("z") Then
KeyAscii = 0
End If
End Sub


Private Sub txtnoofdays_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii >= Asc("a") _
And KeyAscii <= Asc("z") Then
KeyAscii = 0
End If
End Sub

Private Sub txtreturndate_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii >= Asc("a") _
And KeyAscii <= Asc("z") Then
KeyAscii = 0
End If
End Sub






Public Function addStatus()
Dim flag1 As Boolean

If Not IsNumeric(txtissueno) And txtissueno = "" Then
MsgBox "Please enter Issue no and it should be Numeric..", vbInformation

ElseIf txtcustid = "" And Not IsNumeric(txtcustid) Then
MsgBox "Please enter Customer id & it should be Numeric..", vbInformation

ElseIf txtdiskno = "" Then
MsgBox "Please enter your Disk no & it should be Numeric..", vbInformation

ElseIf txtreturndate.Text <> "" Then
MsgBox "This field should be field when the customer returns the Disk.."

ElseIf txtbalanceamount.Text <> "" Then
MsgBox "This field should not be filled by you..", vbInformation

ElseIf txtdiskqty.Text <> "" Then
MsgBox "This field should not be filled by you..", vbInf

ElseIf txtnoofdays.Text <> "" Then
MsgBox "This field should not be filled by you..", vbInformation

ElseIf txtissuedate = "" And Not IsNumeric(txtissuedate) Then
MsgBox "Please enter Issue date & it should be Numeric..", vbInformation


ElseIf txtadvancepaid = "" And Not IsNumeric(txtreturndate) Then
MsgBox "Please enter your Advance paid & it should be Numeric..", vbInformation

Else
flag1 = True

End If

addStatus = flag1
End Function
