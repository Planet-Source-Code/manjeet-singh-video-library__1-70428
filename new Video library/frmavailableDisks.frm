VERSION 5.00
Begin VB.Form frmavailableDisks 
   Caption         =   "Available Disks Form"
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
   Picture         =   "frmavailableDisks.frx":0000
   ScaleHeight     =   5910
   ScaleWidth      =   7575
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstCDList 
      ForeColor       =   &H000000FF&
      Height          =   2910
      ItemData        =   "frmavailableDisks.frx":AC07
      Left            =   3720
      List            =   "frmavailableDisks.frx":AC09
      TabIndex        =   10
      Top             =   1200
      Width           =   2295
   End
   Begin VB.TextBox txttotal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   5400
      TabIndex        =   8
      Top             =   5040
      Width           =   615
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2040
      Picture         =   "frmavailableDisks.frx":AC0B
      TabIndex        =   4
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ok"
      Height          =   495
      Left            =   600
      Picture         =   "frmavailableDisks.frx":B731
      TabIndex        =   3
      Top             =   4320
      Width           =   1215
   End
   Begin VB.ComboBox cmboDiskLang 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   720
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2880
      Width           =   1815
   End
   Begin VB.ComboBox cmboDiskType 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   720
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label lblDiskname 
      BackStyle       =   0  'Transparent
      Caption         =   "Disk Name :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3720
      TabIndex        =   9
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Total Disks :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3840
      TabIndex        =   7
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Label lblDisklang 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Disk Language"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label lblCdType 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Disk Type"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label lbltitle 
      BackColor       =   &H00004080&
      Caption         =   "AVAILABLE DISK'S"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "frmavailableDisks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim con As New Connection
Dim rs As New Recordset
'Dim str As String
'str = "select DISKNAME,diskqty from DISC"


Private Sub cmdcancel_Click()
MDIForm1.Show
frmavailableDisks.Hide
End Sub

Private Sub cmdOk_Click()
Dim string1 As String
Dim j As Integer
j = 0
lstCDList.Clear

string1 = "select diskname from disc"
string1 = string1 & " where disktype = " & "'" & cmboDiskType.List(cmboDiskType.ListIndex) & "'" & " and disklang = " & "'" & cmboDiskLang.List(cmboDiskLang.ListIndex) & "'"

rs.Open string1, con, adOpenDynamic, adLockOptimistic
Debug.Print ins

rs.MoveFirst

While Not rs.EOF
lstCDList.AddItem rs!DISKNAME
j = j + 1
rs.MoveNext
Wend
txttotal.Text = j
rs.Close

End Sub

Private Sub Form_Load()
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Videolibrary.mdb;Persist Security Info=False"
cmboDiskType.AddItem "Video CD"
cmboDiskType.AddItem "Mp3"
cmboDiskLang.AddItem "English"
cmboDiskLang.AddItem "Hindi"
cmboDiskLang.AddItem "Marathi"



End Sub
