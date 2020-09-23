VERSION 5.00
Begin VB.Form frmloginid 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Login Form"
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   7575
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmloginid.frx":0000
   ScaleHeight     =   5910
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtpassword 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   3240
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3000
      Width           =   1300
   End
   Begin VB.TextBox txtusername 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3240
      MaxLength       =   10
      TabIndex        =   0
      Top             =   2280
      Width           =   1300
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H80000013&
      Caption         =   "Cancel"
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
      Left            =   4800
      MaskColor       =   &H00808080&
      Picture         =   "frmloginid.frx":4AC6
      TabIndex        =   4
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H80000013&
      Caption         =   "Ok"
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
      Left            =   1320
      MaskColor       =   &H00808080&
      Picture         =   "frmloginid.frx":55EC
      TabIndex        =   2
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H80000013&
      Caption         =   "Clear"
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
      Left            =   3000
      MaskColor       =   &H00808080&
      Picture         =   "frmloginid.frx":5F24
      TabIndex        =   3
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      DrawMode        =   14  'Copy Pen
      X1              =   600
      X2              =   6975
      Y1              =   600
      Y2              =   615
   End
   Begin VB.Image Image1 
      Height          =   525
      Left            =   960
      Picture         =   "frmloginid.frx":696B
      Top             =   0
      Width           =   5625
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Peth, Pune-411001"
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
      Left            =   2520
      TabIndex        =   9
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Housing Society, Bhavani"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   8
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "7, Lila Co-Operaive"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label lblusername 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "User Name :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label lblPassword 
      BackStyle       =   0  'Transparent
      Caption         =   "Password :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   3000
      Width           =   1695
   End
End
Attribute VB_Name = "frmloginid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const str As String = "Select * from passwd"
Dim con As New Connection
Dim rs As New Recordset
Dim flag As Boolean
Private Sub cmdcancel_Click()
a = MsgBox("Are you sure? You wanna Exit..", vbInformation + vbYesNo, "Message")
If a = 7 Then
txtusername.Text = ""
txtpassword.Text = ""
Else
End
End If

End Sub

Private Sub cmdclear_Click()
txtusername.Text = ""
txtpassword.Text = ""

End Sub

Private Sub cmdOk_Click()
Call Validate
If rs!UserName = txtusername And rs!passwd = txtpassword Then
frmprogressbar.Show
txtusername.SetFocus
frmloginid.Hide
Call cmdclear_Click
frmprogressbar.Refresh
Else
MsgBox "Please enter correct password.."
If rs!UserName = txtusername And rs!passwd <> txtpassword Then
txtpassword.Text = ""
Else
txtpassword.Text = ""
txtusername.Text = ""
txtusername.SetFocus
End If
End If

End Sub



Private Sub Form_Load()
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Videolibrary.mdb;Persist Security Info=False"
rs.Open str, con, adOpenDynamic, adLockOptimistic
rs.MoveFirst
'Timer1.Enabled = False
'ProgressBar1.Value = 0

End Sub



Private Sub txtpassword_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii >= Asc("a") _
And KeyAscii <= Asc("z") Then
KeyAscii = 0
End If
If KeyAscii = 13 Then
Call cmdOk_Click
End If
End Sub

Public Sub Validate()
If Trim(txtusername) = "" Then
MsgBox "Please enter User Name .. ", vbInformation
ElseIf Trim(txtpassword) = "" And Not IsNumeric(txtpassword) Then
MsgBox "Please enter Password & it should be Numeric.. "
End If
End Sub
