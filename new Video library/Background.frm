VERSION 5.00
Begin VB.Form Background 
   BackColor       =   &H8000000E&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "CHANGE BACKGROUND"
   ClientHeight    =   10125
   ClientLeft      =   165
   ClientTop       =   480
   ClientWidth     =   12330
   Icon            =   "Background.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10125
   ScaleWidth      =   12330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdback 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Back To Main Menu"
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
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "To find a record"
      Top             =   7320
      Width           =   2415
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   11760
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   11760
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Image Image12 
      Height          =   1695
      Left            =   9000
      Picture         =   "Background.frx":08CA
      Stretch         =   -1  'True
      Top             =   960
      Width           =   2535
   End
   Begin VB.Image Image11 
      Height          =   1575
      Left            =   9000
      Picture         =   "Background.frx":19305
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   2535
   End
   Begin VB.Image Image10 
      Height          =   1695
      Left            =   9000
      Picture         =   "Background.frx":7DF99
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   1695
      Left            =   3240
      Picture         =   "Background.frx":8479D
      Stretch         =   -1  'True
      Top             =   960
      Width           =   2535
   End
   Begin VB.Image Image2 
      Height          =   1695
      Left            =   240
      Picture         =   "Background.frx":9E729
      Stretch         =   -1  'True
      Top             =   960
      Width           =   2535
   End
   Begin VB.Image Image3 
      Height          =   1575
      Left            =   240
      Picture         =   "Background.frx":AA929
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   2535
   End
   Begin VB.Image Image4 
      Height          =   1695
      Left            =   3240
      Picture         =   "Background.frx":B81A5
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   2535
   End
   Begin VB.Image Image5 
      Height          =   1575
      Left            =   3240
      Picture         =   "Background.frx":C7AF1
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   2535
   End
   Begin VB.Image Image6 
      Height          =   1575
      Left            =   6120
      Picture         =   "Background.frx":D2D40
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   2535
   End
   Begin VB.Image Image7 
      Height          =   1695
      Left            =   240
      Picture         =   "Background.frx":F818E
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   2535
   End
   Begin VB.Image Image8 
      Height          =   1695
      Left            =   6120
      Picture         =   "Background.frx":11612B
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   2535
   End
   Begin VB.Image Image9 
      Height          =   1695
      Left            =   6120
      Picture         =   "Background.frx":11F0A0
      Stretch         =   -1  'True
      Top             =   960
      Width           =   2535
   End
   Begin VB.Shape Shape1 
      Height          =   7335
      Left            =   0
      Top             =   -120
      Width           =   11775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Click on the Picture to have it as background of main window."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   -600
      TabIndex        =   0
      Top             =   120
      Width           =   11535
   End
End
Attribute VB_Name = "Background"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As String

Private Sub cmdBack_Click()
MDIForm1.Show
Background.Hide
End Sub



Private Sub Image1_Click()
MDIForm1.Picture1.Picture = Image1.Picture
MsgBox "Your Background Is Changed!!!!"
MDIForm1.Show
Unload Me
End Sub

Private Sub Image10_Click()
MDIForm1.Picture1.Picture = Image10.Picture
MsgBox "Your Background Is Changed!!!!"
MDIForm1.Show
Unload Me
End Sub

Private Sub Image11_Click()
MDIForm1.Picture1.Picture = Image11.Picture
MsgBox "Your Background Is Changed!!!!"
MDIForm1.Show
Unload Me
End Sub

Private Sub Image12_Click()
MDIForm1.Picture1.Picture = Image12.Picture
MsgBox "Your Background Is Changed!!!!"
MDIForm1.Show
Unload Me
End Sub

Private Sub Image2_Click()
MDIForm1.Picture1.Picture = Image2.Picture
MsgBox "Your Background Is Changed!!!!"
MDIForm1.Show
Unload Me
End Sub

Private Sub Image3_Click()
MDIForm1.Picture1.Picture = Image3.Picture
MsgBox "Your Background Is Changed!!!!"
MDIForm1.Show
Unload Me
End Sub

Private Sub Image4_Click()
MDIForm1.Picture1.Picture = Image4.Picture
MsgBox "Your Background Is Changed!!!!"
MDIForm1.Show
Unload Me
End Sub

Private Sub Image5_Click()
MDIForm1.Picture1.Picture = Image5.Picture

MsgBox "Your Background Is Changed!!!!"
MDIForm1.Show
Unload Me
End Sub

Private Sub Image6_Click()
MDIForm1.Picture1.Picture = Image6.Picture
MsgBox "Your Background Is Changed!!!!"
MDIForm1.Show
Unload Me
End Sub

Private Sub Image7_Click()
MDIForm1.Picture1.Picture = Image7.Picture
MsgBox "Your Background Is Changed!!!!"
MDIForm1.Show
Unload Me
End Sub

Private Sub Image8_Click()
MDIForm1.Picture1.Picture = Image8.Picture
MsgBox "Your Background Is Changed!!!!"
MDIForm1.Show
Unload Me
End Sub

Private Sub Image9_Click()
MDIForm1.Picture1.Picture = Image9.Picture
MsgBox "Your Background Is Changed!!!!"
MDIForm1.Show
Unload Me
End Sub
