VERSION 5.00
Object = "{BF3128D8-55B8-11D4-8ED4-00E07D815373}#1.0#0"; "MBPrgBar.ocx"
Begin VB.Form frmprogressbar 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1395
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3615
   LinkTopic       =   "Form1"
   ScaleHeight     =   1395
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   310
      Left            =   3360
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   3840
      Top             =   0
   End
   Begin MBProgressBar.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   873
      BorderStyle     =   6
      Smooth          =   -1  'True
      BarStartColor   =   255
      BarEndColor     =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackPicture     =   "frmprogressbar.frx":0000
      BarPicture      =   "frmprogressbar.frx":12552
      TextAfterCaption=   "%"
   End
   Begin VB.Label LBLD 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PLEASE WAIT "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frmprogressbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim d As Integer



Private Sub Form_Activate()
Timer1.Enabled = True
Timer2.Enabled = True
End Sub

Private Sub Form_Load()

Timer1.Enabled = True
ProgressBar1.Value = 0
End Sub


Private Sub Timer1_Timer()
If ProgressBar1.Value < 100 Then
ProgressBar1.Value = ProgressBar1.Value + 10
If ProgressBar1.Value >= 1 And ProgressBar1.Value <= 50 Then
ProgressBar1.TextColor = vbBlack
ElseIf ProgressBar1.Value > 50 And ProgressBar1.Value <= 100 Then
ProgressBar1.TextColor = vbWhite
End If
Else
MDIForm1.Show
frmprogressbar.Hide
ProgressBar1.Value = 0
Timer1.Enabled = False
End If

End Sub

Private Sub Timer2_Timer()
d = d + 1
 If d <= 4 Then
    LBLD.Caption = LBLD.Caption & "."
 End If
If d > 5 Then
    d = 0
    LBLD.Caption = ""
    Exit Sub
End If
If ProgressBar1.Value >= 100 Then
Timer2.Enabled = False
End If
End Sub
