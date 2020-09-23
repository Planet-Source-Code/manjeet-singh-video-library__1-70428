VERSION 5.00
Begin VB.Form frmcustomerrecords 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Customer records"
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6555
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   6615
   ScaleWidth      =   6555
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtcustomerno 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   3360
      TabIndex        =   0
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox txtcustomername 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   3360
      TabIndex        =   1
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox txtcustomeraddress 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   3360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox txtcustomerphoneno 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   3360
      TabIndex        =   3
      Top             =   2160
      Width           =   1815
   End
   Begin VB.TextBox txtdiskno 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   3360
      TabIndex        =   4
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox txtdiskname 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   3360
      TabIndex        =   5
      Top             =   3120
      Width           =   1815
   End
   Begin VB.TextBox txtissuedate 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   3360
      TabIndex        =   6
      Top             =   3600
      Width           =   1815
   End
   Begin VB.TextBox txtissueno 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   3360
      TabIndex        =   7
      Top             =   4080
      Width           =   735
   End
   Begin VB.TextBox txtreturndate 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   3360
      TabIndex        =   8
      Top             =   4560
      Width           =   1815
   End
   Begin VB.TextBox txtadvancepaid 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   3360
      TabIndex        =   9
      Top             =   5040
      Width           =   1815
   End
   Begin VB.TextBox txtbalanceamount 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   3360
      TabIndex        =   10
      Top             =   5520
      Width           =   1815
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
      Height          =   360
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "To find a record"
      Top             =   6000
      Width           =   1215
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
      Height          =   360
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "To clear all records"
      Top             =   6000
      Width           =   1215
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
      Height          =   360
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Exit this window"
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Label lblcustno 
      BackColor       =   &H00FFC0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Customer No"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   720
      TabIndex        =   25
      Top             =   720
      Width           =   3135
   End
   Begin VB.Label lblcustname 
      BackColor       =   &H00FFC0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   720
      TabIndex        =   24
      Top             =   1200
      Width           =   3135
   End
   Begin VB.Label lbladd 
      BackColor       =   &H00FFC0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Address"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   720
      TabIndex        =   23
      Top             =   1680
      Width           =   3135
   End
   Begin VB.Label lblcustphno 
      BackColor       =   &H00FFC0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile No"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   720
      TabIndex        =   22
      Top             =   2160
      Width           =   3135
   End
   Begin VB.Label lbldiskno 
      BackColor       =   &H00FFC0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Disk No"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   720
      TabIndex        =   21
      Top             =   2640
      Width           =   3135
   End
   Begin VB.Label lbldiskname 
      BackColor       =   &H00FFC0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Disk Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   720
      TabIndex        =   20
      Top             =   3120
      Width           =   3135
   End
   Begin VB.Label lblissuedt 
      BackColor       =   &H00FFC0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Issue Date"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   720
      TabIndex        =   19
      Top             =   3600
      Width           =   3135
   End
   Begin VB.Label lblissueno 
      BackColor       =   &H00FFC0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Issue No"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   720
      TabIndex        =   18
      Top             =   4080
      Width           =   3135
   End
   Begin VB.Label lblreturndt 
      BackColor       =   &H00FFC0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Return Date"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   720
      TabIndex        =   17
      Top             =   4560
      Width           =   3135
   End
   Begin VB.Label lbladvpaid 
      BackColor       =   &H00FFC0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Advance Paid"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   720
      TabIndex        =   16
      Top             =   5040
      Width           =   3135
   End
   Begin VB.Label lblbalamt 
      BackColor       =   &H00FFC0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Balance Amount"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   720
      TabIndex        =   15
      Top             =   5520
      Width           =   3135
   End
   Begin VB.Label lblcustomer 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Records"
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
      Left            =   1080
      TabIndex        =   14
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmcustomerrecords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const str As String = "select * from CustomerRecord"
Dim str1 As String
Dim con As New Connection
Dim rs As New Recordset
Dim rs1 As New Recordset
Dim estatus As Boolean
Dim ins As String
Dim j As Integer
Dim i As Integer

Private Sub cmdclear_Click()
txtcustomerno.Text = ""
txtcustomername.Text = ""
txtcustomeraddress.Text = ""
txtcustomerphoneno.Text = ""
txtadvancepaid = ""
txtbalanceamount = ""
txtdiskname.Text = ""
txtdiskno.Text = ""
txtissueno.Text = ""
txtissuedate.Text = ""
txtreturndate.Text = ""
End Sub

Private Sub cmdExit_Click()
MDIForm1.Show
frmcustomerrecords.Hide
End Sub

Private Sub cmdfind_Click()

Dim flag As Boolean
i = InputBox("Enter the CustomerID you want to search..", "Find")
rs.MoveFirst
While Not rs.EOF
If rs!custid = Val(i) Then
txtcustomerno.Enabled = False
Call display

If txtreturndate.Text = "" Then
txtbalanceamount.Text = "Disk Not Returned Yet"
End If
flag = True

End If
rs.MoveNext
Wend
If flag = False Then
MsgBox "No Record Exist..", vbCritical
End If

End Sub

Public Sub display()
txtcustomerno.Text = rs!custid
txtcustomername.Text = rs!custname
txtcustomeraddress.Text = rs!CUSTADD
txtcustomerphoneno.Text = rs!custmobno
txtadvancepaid = rs!advpaid
txtbalanceamount = rs!balamt
txtdiskname.Text = rs!DISKNAME
txtdiskno.Text = rs!diskno
txtissueno.Text = rs!issueno
txtissuedate.Text = rs!issuedt
'txtreturndate.Text = rs!returndt
If rs!returndt = "" Then
txtreturndate.Text = "Disk Not Returned Yet"

ElseIf rs!returndt <> "" Then
txtreturndate.Text = rs!returndt
txtadvancepaid.Text = 0
txtbalanceamount.Text = 0
End If
End Sub

Private Sub Form_Load()
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Videolibrary.mdb;Persist Security Info=False"
rs.Open str, con, adOpenDynamic, adLockOptimistic

End Sub
