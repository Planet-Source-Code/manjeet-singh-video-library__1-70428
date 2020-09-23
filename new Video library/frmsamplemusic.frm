VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BF3128D8-55B8-11D4-8ED4-00E07D815373}#1.0#0"; "MBPrgBar.ocx"
Begin VB.Form frmsamplemusic 
   Caption         =   "Sample Music"
   ClientHeight    =   5925
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7455
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "frmsamplemusic.frx":0000
   ScaleHeight     =   5925
   ScaleWidth      =   7455
   StartUpPosition =   2  'CenterScreen
   Begin MBProgressBar.ProgressBar ProgressBar5 
      Height          =   1575
      Left            =   1320
      TabIndex        =   10
      Top             =   2160
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   2778
      BorderStyle     =   6
      CaptionType     =   0
      Value           =   30
      Percentage      =   30
      BarDirection    =   2
      BackColor       =   16777215
      BarStartColor   =   255
      BarEndColor     =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackPicture     =   "frmsamplemusic.frx":4E73
      BarPicture      =   "frmsamplemusic.frx":4E8F
   End
   Begin VB.Timer Timer2 
      Interval        =   40
      Left            =   6000
      Top             =   480
   End
   Begin VB.Timer Timer1 
      Interval        =   15
      Left            =   6480
      Top             =   480
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C0C0C0&
      Caption         =   "E&xit"
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
      Left            =   5520
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4680
      Width           =   1215
   End
   Begin VB.ListBox lstsamplesongs 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2325
      ItemData        =   "frmsamplemusic.frx":4EAB
      Left            =   4920
      List            =   "frmsamplemusic.frx":4EAD
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   1560
      Width           =   1815
   End
   Begin MCI.MMControl MMControl1 
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   3960
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   873
      _Version        =   393216
      EjectEnabled    =   -1  'True
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.CommandButton cmdopenfile 
      Appearance      =   0  'Flat
      DisabledPicture =   "frmsamplemusic.frx":4EAF
      DownPicture     =   "frmsamplemusic.frx":59AE
      Height          =   855
      Left            =   1560
      Picture         =   "frmsamplemusic.frx":61D3
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4560
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7440
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MBProgressBar.ProgressBar ProgressBar4 
      Height          =   495
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   873
      BorderStyle     =   6
      CaptionType     =   0
      Value           =   25
      Percentage      =   25
      Smooth          =   -1  'True
      BackColor       =   16777215
      BarStartColor   =   255
      BarEndColor     =   255
      TextColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackPicture     =   "frmsamplemusic.frx":69F8
      BarPicture      =   "frmsamplemusic.frx":6A14
      TextAfterCaption=   "Free Mp3 Songs"
   End
   Begin MBProgressBar.ProgressBar ProgressBar3 
      Height          =   5415
      Left            =   6840
      TabIndex        =   7
      Top             =   0
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   9551
      CaptionType     =   0
      Value           =   25
      Percentage      =   25
      Smooth          =   -1  'True
      BarDirection    =   3
      VerticalText    =   -1  'True
      BackColor       =   16777215
      BarStartColor   =   65535
      BarEndColor     =   65535
      TextColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackPicture     =   "frmsamplemusic.frx":6A30
      BarPicture      =   "frmsamplemusic.frx":6A4C
      TextBeforeCaption=   "Free Mp3 Songs"
   End
   Begin MBProgressBar.ProgressBar ProgressBar2 
      Height          =   495
      Left            =   600
      TabIndex        =   8
      Top             =   5400
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   873
      BorderStyle     =   6
      CaptionType     =   0
      Value           =   25
      Percentage      =   25
      Smooth          =   -1  'True
      BarDirection    =   1
      BackColor       =   16777215
      BarStartColor   =   16711680
      BarEndColor     =   16711680
      TextColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackPicture     =   "frmsamplemusic.frx":6A68
      BarPicture      =   "frmsamplemusic.frx":6A84
      TextAfterCaption=   "Free Mp3 Songs"
   End
   Begin MBProgressBar.ProgressBar ProgressBar1 
      DragIcon        =   "frmsamplemusic.frx":6AA0
      Height          =   5415
      Left            =   0
      TabIndex        =   9
      Top             =   480
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   9551
      MouseIcon       =   "frmsamplemusic.frx":796A
      BorderStyle     =   7
      CaptionType     =   0
      Value           =   25
      Percentage      =   25
      Smooth          =   -1  'True
      BarDirection    =   2
      VerticalText    =   -1  'True
      BackColor       =   16777215
      BarStartColor   =   128
      BarEndColor     =   33023
      TextColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackPicture     =   "frmsamplemusic.frx":8844
      BarPicture      =   "frmsamplemusic.frx":8860
      TextAfterCaption=   "Free Mp3 Songs"
   End
   Begin MBProgressBar.ProgressBar ProgressBar6 
      Height          =   1575
      Left            =   1920
      TabIndex        =   11
      Top             =   2160
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   2778
      BorderStyle     =   6
      CaptionType     =   0
      Value           =   25
      Percentage      =   25
      BarDirection    =   2
      BackColor       =   16777215
      BarStartColor   =   16711680
      BarEndColor     =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackPicture     =   "frmsamplemusic.frx":887C
      BarPicture      =   "frmsamplemusic.frx":8898
   End
   Begin MBProgressBar.ProgressBar ProgressBar7 
      Height          =   1575
      Left            =   2520
      TabIndex        =   12
      Top             =   2160
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   2778
      BorderStyle     =   6
      CaptionType     =   0
      Value           =   15
      Percentage      =   15
      BarDirection    =   2
      BackColor       =   16777215
      BarStartColor   =   0
      BarEndColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackPicture     =   "frmsamplemusic.frx":88B4
      BarPicture      =   "frmsamplemusic.frx":88D0
   End
   Begin MBProgressBar.ProgressBar ProgressBar8 
      Height          =   1575
      Left            =   1680
      TabIndex        =   13
      Top             =   2160
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   2778
      BorderStyle     =   6
      CaptionType     =   0
      Value           =   15
      Percentage      =   15
      Smooth          =   -1  'True
      BarDirection    =   2
      BackColor       =   16777215
      BarStartColor   =   65280
      BarEndColor     =   65280
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackPicture     =   "frmsamplemusic.frx":88EC
      BarPicture      =   "frmsamplemusic.frx":8908
   End
   Begin MBProgressBar.ProgressBar ProgressBar9 
      Height          =   1575
      Left            =   2280
      TabIndex        =   14
      Top             =   2160
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   2778
      BorderStyle     =   6
      CaptionType     =   0
      Value           =   15
      Percentage      =   15
      Smooth          =   -1  'True
      BarDirection    =   2
      BackColor       =   16777215
      BarStartColor   =   65280
      BarEndColor     =   65280
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackPicture     =   "frmsamplemusic.frx":8924
      BarPicture      =   "frmsamplemusic.frx":8940
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   4
      FillStyle       =   7  'Diagonal Cross
      Height          =   1815
      Left            =   1200
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Sample Songs"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Listen Free Mp3 Music"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   840
      TabIndex        =   2
      Top             =   600
      Width           =   5775
   End
End
Attribute VB_Name = "frmsamplemusic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdExit_Click()
MDIForm1.Show
frmsamplemusic.Hide
End Sub

Private Sub cmdopenfile_Click()
CommonDialog1.Filter = "WAV(*.wav)|*.wav"
CommonDialog1.ShowOpen
MMControl1.Command = "Close"
MMControl1.FileName = CommonDialog1.FileName
MMControl1.Command = "open"
End Sub

Private Sub Form_Load()
MMControl1.Notify = False
MMControl1.Wait = True
MMControl1.Shareable = False
MMControl1.DeviceType = "WaveAudio"
MMControl1.UpdateInterval = 500

lstsamplesongs.AddItem "Kala Chasma"
lstsamplesongs.AddItem "Akh matkave"
lstsamplesongs.AddItem "Just Chill"
lstsamplesongs.AddItem "Salaam Namaste"
lstsamplesongs.AddItem "Aap Ke Kashish"
lstsamplesongs.AddItem "Sun Zara"
lstsamplesongs.AddItem "Jay Sean"
lstsamplesongs.AddItem "Jazzy Romeo"

Timer1.Enabled = True
ProgressBar1.Value = 0
ProgressBar2.Value = 0
ProgressBar3.Value = 0
ProgressBar4.Value = 0

Timer2.Enabled = False
End Sub

Private Sub Form_Terminate()
MMControl1.Command = "Close"
End Sub


Private Sub MMControl1_PauseClick(Cancel As Integer)
Timer2.Enabled = False
End Sub

Private Sub MMControl1_PlayClick(Cancel As Integer)
cmdopenfile.Enabled = False

Timer2.Enabled = True
ProgressBar5.Value = 0
ProgressBar6.Value = 0
ProgressBar7.Value = 0
ProgressBar8.Value = 0
ProgressBar9.Value = 0
End Sub


Private Sub MMControl1_StatusUpdate()
If (MMControl1.Mode <> mciModePlay And _
MMControl1.Mode <> mciModePause) Then
cmdopenfile.Enabled = True
End If
End Sub





Private Sub MMControl1_StopClick(Cancel As Integer)
Timer2.Enabled = False
End Sub

Private Sub Timer1_Timer()
  
If ProgressBar1.Value < 100 Then
ProgressBar1.Value = ProgressBar1.Value + 3
ProgressBar2.Value = ProgressBar2.Value + 3
ProgressBar3.Value = ProgressBar3.Value + 3
ProgressBar4.Value = ProgressBar4.Value + 3
Else
ProgressBar1.Value = 0
ProgressBar2.Value = 0
ProgressBar3.Value = 0
ProgressBar4.Value = 0
End If
End Sub

Private Sub Timer2_Timer()
If ProgressBar5.Value < 100 Then
ProgressBar5.Value = ProgressBar5.Value + 2
ProgressBar6.Value = ProgressBar6.Value + 4
ProgressBar7.Value = ProgressBar7.Value + 1
ProgressBar8.Value = ProgressBar8.Value + 3
ProgressBar9.Value = ProgressBar9.Value + 2
Else
ProgressBar5.Value = 0
ProgressBar6.Value = 0
ProgressBar7.Value = 0
ProgressBar8.Value = 0
ProgressBar9.Value = 0
End If
End Sub
