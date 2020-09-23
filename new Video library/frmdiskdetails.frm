VERSION 5.00
Begin VB.Form frmdiskdetails 
   Caption         =   "Disk Details Form"
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
   Picture         =   "frmdiskdetails.frx":0000
   ScaleHeight     =   5910
   ScaleWidth      =   7575
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtdiskqty 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3720
      MaxLength       =   3
      TabIndex        =   21
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4800
      Picture         =   "frmdiskdetails.frx":4C2EF
      TabIndex        =   3
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   3000
      Picture         =   "frmdiskdetails.frx":4CE15
      TabIndex        =   2
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   495
      Left            =   1200
      Picture         =   "frmdiskdetails.frx":4D85C
      TabIndex        =   1
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox txtlang 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3720
      TabIndex        =   19
      Top             =   4560
      Width           =   1300
   End
   Begin VB.TextBox txtprice 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3720
      TabIndex        =   18
      Top             =   4080
      Width           =   1300
   End
   Begin VB.TextBox txtdesc 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3720
      TabIndex        =   17
      Top             =   3600
      Width           =   1300
   End
   Begin VB.TextBox txtdisktype 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3720
      TabIndex        =   16
      Top             =   3120
      Width           =   1300
   End
   Begin VB.TextBox txtrackno 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3720
      MaxLength       =   3
      TabIndex        =   15
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox txtshelfno 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3720
      MaxLength       =   3
      TabIndex        =   14
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox txtdiskno 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3720
      MaxLength       =   3
      TabIndex        =   13
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox txtmoviename 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3720
      TabIndex        =   0
      Top             =   720
      Width           =   1300
   End
   Begin VB.Label lbldiskqty 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Disk Quantity"
      Height          =   375
      Left            =   720
      TabIndex        =   20
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label lbldisklang 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Disk Language"
      Height          =   375
      Left            =   720
      TabIndex        =   12
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label lblDiskprice 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Disk Price"
      Height          =   375
      Left            =   720
      TabIndex        =   11
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label lblDiskdesc 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Disk Description"
      Height          =   375
      Left            =   720
      TabIndex        =   10
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label lblDisktype 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Disk Type"
      Height          =   375
      Left            =   720
      TabIndex        =   9
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label lblRackno 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Disk Rack No."
      Height          =   375
      Left            =   720
      TabIndex        =   8
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label lblshelfno 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Disk Shelf No."
      Height          =   375
      Left            =   720
      TabIndex        =   7
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label lbldiskno 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Disk No"
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblmoviename 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter Movie Name"
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label lblDiskdetails 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "DISK DETAILS"
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
      Left            =   1440
      TabIndex        =   4
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmdiskdetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Const str As String = "select * from disc"
 Dim str1 As String
 Dim con As New Connection
Dim rs1 As New Recordset
Dim estatus As Boolean
Dim ins As String
Dim j As Integer

Private Sub cmdcancel_Click()
MDIForm1.Show
frmdiskdetails.Hide
End Sub

Private Sub cmdclear_Click()
txtmoviename.Enabled = True
txtdiskno.Text = ""
txtshelfno.Text = ""
txtrackno.Text = ""
txtdisktype.Text = ""
txtdesc.Text = ""
txtlang.Text = ""
txtprice.Text = ""
txtmoviename = ""
txtdiskqty.Text = ""
txtmoviename.SetFocus
End Sub

Private Sub cmdOk_Click()
Dim flag As Boolean
rs1.MoveFirst
While Not rs1.EOF
If rs1!DISKNAME = StrConv(txtmoviename, vbProperCase) Then
Call display
txtmoviename.Enabled = False

flag = True
End If
rs1.MoveNext
Wend
If flag = False Then
MsgBox "Sorry this Movie Disc is not Available..", vbCritical
txtmoviename.Text = ""
txtmoviename.Enabled = True

End If
End Sub

Private Sub Form_Load()
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Videolibrary.mdb;Persist Security Info=False"
rs1.Open str, con, adOpenDynamic, adLockOptimistic
End Sub

Public Sub display()
txtmoviename.Text = rs1!DISKNAME
txtprice.Text = "Rs "
txtdiskno.Text = rs1!diskno
txtshelfno.Text = rs1!diskshelfno
txtrackno.Text = rs1!diskrackno
txtdisktype.Text = StrConv(rs1!disktype, vbProperCase)
txtdesc.Text = StrConv(rs1!diskdesc, vbProperCase)
txtlang.Text = StrConv(rs1!disklang, vbProperCase)
txtprice.Text = txtprice.Text & rs1!DISKRT
txtdiskqty.Text = rs1!diskqty
End Sub



Private Sub txtmoviename_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call cmdOk_Click
End If
End Sub
