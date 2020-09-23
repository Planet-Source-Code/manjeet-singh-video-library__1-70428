VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Music World Video Library"
   ClientHeight    =   8190
   ClientLeft      =   165
   ClientTop       =   -1185
   ClientWidth     =   11880
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":0000
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   9060
      Left            =   0
      Picture         =   "MDIForm1.frx":1B7D
      ScaleHeight     =   9000
      ScaleWidth      =   11820
      TabIndex        =   0
      Top             =   0
      Width           =   11880
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1050
         Left            =   2640
         Picture         =   "MDIForm1.frx":10BB8
         ScaleHeight     =   1050
         ScaleWidth      =   5550
         TabIndex        =   2
         Top             =   -120
         Width           =   5550
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3135
         Left            =   9000
         Picture         =   "MDIForm1.frx":12616
         ScaleHeight     =   3135
         ScaleWidth      =   2295
         TabIndex        =   1
         Top             =   120
         Width           =   2295
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00008000&
         BorderStyle     =   4  'Dash-Dot
         BorderWidth     =   6
         X1              =   2640
         X2              =   7920
         Y1              =   960
         Y2              =   960
      End
   End
   Begin VB.Menu mnumaster 
      Caption         =   "Master"
      Index           =   0
      Begin VB.Menu mnuavailabledisks 
         Caption         =   "Available disks"
         Index           =   1
      End
      Begin VB.Menu mnudiskdetails 
         Caption         =   "Disk details"
         Index           =   2
      End
      Begin VB.Menu mnudiskrecord 
         Caption         =   "Disk record"
         Index           =   3
      End
      Begin VB.Menu mnubgchange 
         Caption         =   "Background Change"
      End
      Begin VB.Menu mnufreemp3 
         Caption         =   "Free Sample Mp3"
      End
   End
   Begin VB.Menu mnutransaction 
      Caption         =   "Transaction"
      Index           =   4
      Begin VB.Menu mnucustomerdetails 
         Caption         =   "Add Customer"
         Index           =   5
      End
      Begin VB.Menu mnuissuerecord 
         Caption         =   "Issue Records"
         Index           =   6
      End
      Begin VB.Menu mnucustomerrecord 
         Caption         =   "Customer Records"
         Index           =   7
      End
      Begin VB.Menu mnuchangepasswd 
         Caption         =   "Change Password"
         Shortcut        =   ^P
      End
   End
   Begin VB.Menu mnureports 
      Caption         =   "Reports"
      Index           =   8
   End
   Begin VB.Menu mnuexit 
      Caption         =   "Exit"
      Index           =   9
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub mnuavailabledisks_Click(Index As Integer)
frmavailableDisks.Show
MDIForm1.Hide
End Sub

Private Sub mnubgchange_Click()
Background.Show
MDIForm1.Hide
End Sub

Private Sub mnuchangepasswd_Click()
frmchangepasswd.Show
MDIForm1.Hide
End Sub

Private Sub mnucustomerdetails_Click(Index As Integer)
frmcustomerdetails.Show
MDIForm1.Hide
End Sub

Private Sub mnucustomerrecord_Click(Index As Integer)
frmcustomerrecords.Show
MDIForm1.Hide
End Sub

Private Sub mnudiskdetails_Click(Index As Integer)
frmdiskdetails.Show
MDIForm1.Hide
End Sub

Private Sub mnudiskrecord_Click(Index As Integer)
frmdiskrecords.Show
MDIForm1.Hide
End Sub

Private Sub mnuexit_Click(Index As Integer)
frmloginid.Show
MDIForm1.Hide
End Sub

Private Sub mnufreemp3_Click()
frmsamplemusic.Show
MDIForm1.Hide
End Sub

Private Sub mnuissuerecord_Click(Index As Integer)
frmissuerecord.Show
MDIForm1.Hide
End Sub

