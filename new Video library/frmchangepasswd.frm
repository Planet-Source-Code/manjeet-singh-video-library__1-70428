VERSION 5.00
Begin VB.Form frmchangepasswd 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Change Password Form"
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7575
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "frmchangepasswd.frx":0000
   ScaleHeight     =   5910
   ScaleWidth      =   7575
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame framnewpasswd 
      BackColor       =   &H00FFFFFF&
      Caption         =   "New Password"
      ForeColor       =   &H000000FF&
      Height          =   2655
      Left            =   720
      TabIndex        =   15
      Top             =   3240
      Width           =   5415
      Begin VB.TextBox txtusrname 
         Appearance      =   0  'Flat
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   3000
         MaxLength       =   10
         TabIndex        =   5
         Top             =   240
         Width           =   1300
      End
      Begin VB.TextBox txtnewpasswd 
         Appearance      =   0  'Flat
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   3000
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   840
         Width           =   1300
      End
      Begin VB.TextBox txtconfirmpasswd 
         Appearance      =   0  'Flat
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   3000
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   1440
         Width           =   1300
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "Ok"
         Height          =   495
         Left            =   600
         Picture         =   "frmchangepasswd.frx":228C
         TabIndex        =   8
         Top             =   1920
         Width           =   1215
      End
      Begin VB.CommandButton cmdclear 
         Caption         =   "Clear"
         Height          =   495
         Left            =   2160
         Picture         =   "frmchangepasswd.frx":2BC4
         TabIndex        =   9
         Top             =   1920
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   495
         Left            =   3720
         Picture         =   "frmchangepasswd.frx":360B
         TabIndex        =   10
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label lblusrname 
         BackStyle       =   0  'Transparent
         Caption         =   "New User Name"
         Height          =   495
         Left            =   840
         TabIndex        =   18
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblnewpasswd 
         BackStyle       =   0  'Transparent
         Caption         =   "New Password"
         Height          =   495
         Left            =   840
         TabIndex        =   17
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblconfirmpasswd 
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm Password"
         Height          =   495
         Left            =   840
         TabIndex        =   16
         Top             =   1560
         Width           =   1815
      End
   End
   Begin VB.Frame framcoldpasswd 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Password Validation"
      ForeColor       =   &H000000FF&
      Height          =   2295
      Left            =   720
      TabIndex        =   12
      Top             =   720
      Width           =   5415
      Begin VB.CommandButton cmdok1 
         Caption         =   "Ok"
         Height          =   495
         Left            =   600
         Picture         =   "frmchangepasswd.frx":4131
         TabIndex        =   2
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton cmdclear1 
         Caption         =   "Clear"
         Height          =   495
         Left            =   2160
         Picture         =   "frmchangepasswd.frx":4A69
         TabIndex        =   3
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton cmdcancel1 
         Caption         =   "Cancel"
         Height          =   495
         Left            =   3720
         Picture         =   "frmchangepasswd.frx":54B0
         TabIndex        =   4
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox txtoldpasswd 
         Appearance      =   0  'Flat
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   2880
         MaxLength       =   4
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   840
         Width           =   1300
      End
      Begin VB.TextBox txtusername 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   2880
         TabIndex        =   0
         Top             =   240
         Width           =   1300
      End
      Begin VB.Label lbloldpassword 
         BackStyle       =   0  'Transparent
         Caption         =   "Old Password"
         Height          =   495
         Left            =   720
         TabIndex        =   14
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label lblusername 
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
         Height          =   495
         Left            =   720
         TabIndex        =   13
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Label lblChangePassword 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CHANGE PASSWORD"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   11
      Top             =   240
      Width           =   4335
   End
End
Attribute VB_Name = "frmchangepasswd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const str As String = "select * from passwd"
 Dim str1 As String
 Dim con As New Connection
Dim rs As New Recordset
Dim estatus As Boolean
Dim ins As String
Dim j As Integer

Private Sub cmdcancel_Click()
cmdOk.Enabled = True
cmdclear.Enabled = True
cmdcancel.Enabled = True
cmdok1.Enabled = True
cmdcancel1.Enabled = True
cmdclear1.Enabled = True
txtusrname.Enabled = False
txtoldpasswd.Enabled = True
txtusername.Enabled = True
txtnewpasswd.Enabled = False
txtconfirmpasswd.Enabled = False
cmdOk.Enabled = False
cmdclear.Enabled = False
cmdcancel.Enabled = False
txtconfirmpasswd.Text = ""
txtnewpasswd.Text = ""
txtusrname.Text = ""
txtusername.Text = ""
txtoldpasswd.Text = ""
End Sub

Private Sub cmdcancel1_Click()
MDIForm1.Show
frmchangepasswd.Hide
End Sub

Private Sub cmdclear_Click()
txtconfirmpasswd.Text = ""
txtnewpasswd.Text = ""

End Sub

Private Sub cmdOk_Click()

If txtnewpasswd.Text <> "" And txtconfirmpasswd.Text <> "" And txtnewpasswd.Text = txtconfirmpasswd Then
con.BeginTrans
ins = "update passwd set passwd = " & txtnewpasswd & "," & " username = " & "'" & txtusrname & "'" & " where username = " & "'" & txtusername & "'"
Debug.Print ins
If estatus = False Then
con.Execute ins
End If
con.CommitTrans
MsgBox "Your Password has been Changed.."
Call cmdcancel_Click
txtusername.Text = ""
txtoldpasswd.Text = ""
Else
MsgBox "The Password does not match or the field is empty..", vbCritical
txtconfirmpasswd.Text = ""
txtnewpasswd.Text = ""
End If

End Sub

Private Sub Form_Load()
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Videolibrary.mdb;Persist Security Info=False"
rs.Open str, con, adOpenDynamic, adLockOptimistic
txtnewpasswd.Enabled = False
txtconfirmpasswd.Enabled = False
cmdOk.Enabled = False
cmdclear.Enabled = False
cmdcancel.Enabled = False
End Sub

Private Sub cmdclear1_Click()
txtusername.Text = ""
txtoldpasswd.Text = ""
End Sub

Private Sub cmdok1_Click()

If rs!UserName = StrConv(txtusername, vbLowerCase) And rs!passwd = txtoldpasswd Then
txtoldpasswd.Enabled = False
txtusername.Enabled = False
txtnewpasswd.Enabled = True
txtconfirmpasswd.Enabled = True
cmdok1.Enabled = False
cmdcancel1.Enabled = False
cmdclear1.Enabled = False
cmdOk.Enabled = True
cmdclear.Enabled = True
cmdcancel.Enabled = True
txtusrname.Enabled = True
txtusrname.SetFocus
Else
MsgBox "Access Denied..", vbCritical
txtusername.Text = ""
txtoldpasswd.Text = ""
End If

End Sub
