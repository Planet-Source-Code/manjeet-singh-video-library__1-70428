VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl UserControl1 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   Picture         =   "UserControl1.ctx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.TextBox txtSpeed 
      Height          =   285
      Left            =   4080
      TabIndex        =   1
      Top             =   3240
      Width           =   615
   End
   Begin VB.Timer tmrTimer 
      Interval        =   500
      Left            =   0
      Top             =   0
   End
   Begin MSComCtl2.UpDown updSpeed 
      Height          =   255
      Left            =   3840
      TabIndex        =   0
      Top             =   3240
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      _Version        =   393216
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Sub tmrTimer_Timer()
Call DrawShape
End Sub

Private Sub updSpeed_Change()
tmrTimer.Interval = updSpeed.Value
   txtSpeed.Text = updSpeed.Value
End Sub

Private Sub UserControl_Initialize()
txtSpeed.Text = tmrTimer.Interval
   
   With updSpeed
      ' Make txtSpeed the buddy control
      .BuddyControl = txtSpeed
      
      .Min = 0       ' Minimum value
      .Max = 1000    ' Maximum value
      .Wrap = True   ' When min is exceeded wrap to 1000
                     ' and when max is exceeded wrap to 0
                     
                     
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  This line was added to ensure that the control begins with '
'  the default value of 500 given to the buddy control.       '
'  Without this fix, the control does not increment/decrement '
'  properly the first time an arrow is clicked.               '

      .Value = tmrTimer.Interval   ' Set default at 500
      
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      
      .Increment = 100  ' Increment/decrement amount
   End With
End Sub

Private Sub DrawShape()
   Dim x As Single, y As Single
   Dim totalRadians As Single, r As Single
   Dim a As Single, theta As Single
   
   Call Randomize
   Scale (3, -3)-(-3, 3)         ' Change scale
   totalRadians = 8 * Atn(1)     ' Circle in Radians
   
   ForeColor = QBColor(Rnd() * 15)
   
   a = 3 * Rnd()  ' Offset used in equation
      
   For theta = 0 To totalRadians Step 0.01
      r = a * Sin(10 * theta) ' Multi-Leaved Rose
      x = r * Cos(theta)      ' y coordinate
      y = r * Sin(theta)      ' x coordinate
      PSet (x, y)             ' Turn pixel on
   Next theta
   
End Sub


