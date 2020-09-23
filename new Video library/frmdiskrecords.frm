VERSION 5.00
Begin VB.Form frmdiskrecords 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Disk records"
   ClientHeight    =   5310
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7035
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   7035
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtdiskqty 
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
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   3240
      TabIndex        =   25
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox txtdiskno 
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
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   3240
      TabIndex        =   0
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox txtdiskname 
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
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   3240
      TabIndex        =   1
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox txtdisktype 
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
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   3240
      TabIndex        =   2
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox txtdisklang 
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
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   3240
      TabIndex        =   3
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox txtdiskdesc 
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
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   3240
      TabIndex        =   4
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox txtdiskrt 
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
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   3240
      TabIndex        =   5
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox txtdisksno 
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
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   3240
      TabIndex        =   6
      Top             =   4080
      Width           =   1455
   End
   Begin VB.TextBox txtdiskrno 
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
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   3240
      TabIndex        =   7
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&EXIT"
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
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "To exit the window"
      Top             =   3840
      Width           =   1335
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
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "To clear the textboxes"
      Top             =   3360
      Width           =   1335
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
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "To delete a record"
      Top             =   2880
      Width           =   1335
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
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "To find the disk records"
      Top             =   2400
      Width           =   1335
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
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "To modify the records"
      Top             =   1920
      Width           =   1335
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
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "To save the changes made by modify option"
      Top             =   1440
      Width           =   1335
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
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "To add a new Disk record"
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label lbldiskqty 
      BackStyle       =   0  'Transparent
      Caption         =   "Disk Quantity"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   600
      TabIndex        =   24
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label lbldiskno 
      BackStyle       =   0  'Transparent
      Caption         =   "Disk No                               "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   600
      TabIndex        =   23
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lbldisknaame 
      BackStyle       =   0  'Transparent
      Caption         =   "Disk Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   600
      TabIndex        =   22
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label lbldisktype 
      BackStyle       =   0  'Transparent
      Caption         =   "Disk Type"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   600
      TabIndex        =   21
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label lbldisklang 
      BackStyle       =   0  'Transparent
      Caption         =   "Disk Language"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   600
      TabIndex        =   20
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label lbldiskdesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Disk Description"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   600
      TabIndex        =   19
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label lbldiskrt 
      BackStyle       =   0  'Transparent
      Caption         =   "Disk Rate"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   600
      TabIndex        =   18
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label lbldiskshelfno 
      BackStyle       =   0  'Transparent
      Caption         =   "Disk Shelf No"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   600
      TabIndex        =   17
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label lbldiskrackno 
      BackStyle       =   0  'Transparent
      Caption         =   "Disk Rack No"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   600
      TabIndex        =   16
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Label lbldisk 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Disk Records"
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
      Left            =   960
      TabIndex        =   15
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "frmdiskrecords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Const str As String = "select * from DISC "
 Dim str1 As String
 Dim con As New Connection
Dim rs As New Recordset
Dim estatus As Boolean
Dim ins As String
Dim j As Integer
Dim m As Integer

Private Sub cmdAdd_Click()
If Validate = True Then
con.BeginTrans
ins = "insert into DISC(DISKNO,DISKNAME,DISKTYPE,disklang,DISKDESC,DISKRT,DISKSHELFNO,DISKRACKNO,diskqty)"
ins = ins & "values(" & txtdiskno & "," & "'" & txtdiskname & "'" & "," & "'" & txtdisktype & "'" _
& "," & "'" & txtdisklang & "'" & "," & "'" & txtdiskdesc & "'" & "," & txtdiskrt & "," & txtdisksno & "," & txtdiskrno & "," & txtdiskqty.Text & ")"
Debug.Print ins
If estatus = False Then
con.Execute ins
End If
con.CommitTrans
MsgBox "Your Data has been saved.."
frmdiskrecords.Refresh
Call clearfields

End If
txtdiskno.Enabled = True
rs.MoveLast
txtdiskno = rs!diskno + 1
txtdiskno.Enabled = False

End Sub

Private Sub cmdclear_Click()
Call clearfields
txtdiskno.Text = m
End Sub

Private Sub cmdDelete_Click()
Dim yes As String
Dim k As Integer
k = InputBox("Enter the record no you want to delete..", "Delete..")
rs.MoveFirst
While Not rs.EOF
If rs!diskno = Val(k) Then
yes = MsgBox("Are you sure you want to delete this record ..", vbYesNo)
If yes = vbYes Then
rs.Delete
If rs.EOF = True Then
rs.MoveLast
End If
End If
End If
rs.MoveNext
Wend
End Sub

Private Sub cmdExit_Click()
frmdiskrecords.Hide
MDIForm1.Show

End Sub

Private Sub cmdfind_Click()
Dim i As Integer
Dim flag As Boolean
On Error Resume Next
i = InputBox("Enter the Disk No you want to search..", "Find")

rs.MoveFirst
While Not rs.EOF
If rs!diskno = Val(i) Then
txtdiskno.Enabled = False
Call display
Call txtDisable
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
txtdiskno.Text = rs!diskno
txtdiskname.Text = rs!DISKNAME
txtdisktype.Text = rs!disktype
txtdisklang.Text = rs!disklang
txtdiskdesc.Text = rs!diskdesc
txtdiskrt.Text = rs!DISKRT
txtdisksno.Text = rs!diskshelfno
txtdiskrno.Text = rs!diskrackno
txtdiskqty.Text = rs!diskqty
End Sub

Private Sub cmdmodify_Click()
On Error Resume Next
j = InputBox("Enter the diskno you want to Modify..", "Modify")
rs.MoveFirst
While Not rs.EOF
If rs!diskno = Val(j) Then
txtdiskno.Enabled = False
Call display
txtdiskno.Enabled = False
Call txtEnable
End If
rs.MoveNext
Wend


 
End Sub

Private Sub cmdsave_Click()
If Validate = True Then
con.BeginTrans
ins = "update DISC set DISKNAME = " & "'" & txtdiskname & "'" & "," & " DISKTYPE = " & "'" & txtdisktype & "'" & "," _
& " disklang = " & "'" & txtdisklang & "'" & "," & " DISKDESC = " & "'" & txtdiskdesc & "'" & "," & " DISKRT = " & txtdiskrt & "," & " DISKSHELFNO = " & txtdisksno & "," & " DISKRACKKNO = " & txtdiskrno & "," & " diskqty = " & txtdiskqty & " where DISKNO = " & Val(j)
Debug.Print ins
If estatus = False Then
con.Execute ins
End If
con.CommitTrans
MsgBox "Your Data has been Modified.."
End If
Call clearfields
End Sub

Private Sub Form_Load()

con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Videolibrary.mdb;Persist Security Info=False"
rs.Open str, con, adOpenDynamic, adLockOptimistic
rs.MoveLast
txtdiskno = rs!diskno + 1
m = txtdiskno.Text
txtdiskno.Enabled = False

End Sub

Public Sub clearfields()
txtdiskno.Text = ""
txtdiskname.Text = ""
txtdisktype.Text = ""
txtdisklang.Text = ""
txtdiskdesc.Text = ""
txtdiskrt.Text = ""
txtdisksno.Text = ""
txtdiskrno.Text = ""
txtdiskqty.Text = ""
txtdiskno.Enabled = False
'txtdiskname.SetFocus

End Sub

Public Function Validate()
Dim flag As Boolean
If Not IsNumeric(txtdiskno) And txtdiskno = "" Then
MsgBox "Please enter Disk No and it should be Numeric..", vbInformation

ElseIf txtdiskname = "" Then
MsgBox "Please enter Disk Name and it should be character..", vbInformation

ElseIf txtdisktype = "" Then
MsgBox "Please enter Disk Type it should be character..", vbInformation

ElseIf txtdisklang = "" Then
MsgBox "Please enter Disk Lang it should be character..", vbInformation

ElseIf txtdiskdesc = "" Then
MsgBox "Please enter Disk Desc it should be character..", vbInformation

ElseIf txtdiskrt = "" And Not IsNumeric(txtphoneno) Then
MsgBox "Please enter your Disk RT & it should be Numeric....", vbInformation

ElseIf txtdisksno = "" And Not IsNumeric(txtphoneno) Then
MsgBox "Please enter your Disk SNO & it should be Numeric....", vbInformation

ElseIf txtdiskrno = "" And Not IsNumeric(txtphoneno) Then
MsgBox "Please enter your Disk RNO & it should be Numeric....", vbInformation

Else
flag = True
End If
Validate = flag
End Function





Private Sub txtdiskdesc_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
KeyAscii = 0
End If
End Sub

Private Sub txtdisklang_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
KeyAscii = 0
End If
End Sub

Private Sub txtdiskname_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
KeyAscii = 0
End If
End Sub

Private Sub txtdiskno_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii >= Asc("a") _
And KeyAscii <= Asc("z") Then
KeyAscii = 0
End If
End Sub

Private Sub txtdiskrno_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii >= Asc("a") _
And KeyAscii <= Asc("z") Then
KeyAscii = 0
End If
End Sub

Private Sub txtdiskrt_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii >= Asc("a") _
And KeyAscii <= Asc("z") Then
KeyAscii = 0
End If
End Sub

Private Sub txtdisksno_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii >= Asc("a") _
And KeyAscii <= Asc("z") Then
KeyAscii = 0
End If
End Sub

Private Sub txtdisktype_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
KeyAscii = 0
End If
End Sub



Public Sub txtEnable()
txtdiskname.Enabled = True
txtdisktype.Enabled = True
txtdisklang.Enabled = True
txtdiskdesc.Enabled = True
txtdiskrt.Enabled = True
txtdisksno.Enabled = True
txtdiskrno.Enabled = True
txtdiskqty.Enabled = True
End Sub

Public Sub txtDisable()
txtdiskname.Enabled = False
txtdisktype.Enabled = False
txtdisklang.Enabled = False
txtdiskdesc.Enabled = False
txtdiskrt.Enabled = False
txtdisksno.Enabled = False
txtdiskrno.Enabled = False
txtdiskqty.Enabled = False
End Sub
