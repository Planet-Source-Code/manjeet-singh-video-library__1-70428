VERSION 5.00
Object = "{3B008041-905A-11D1-B4AE-444553540000}#1.0#0"; "VSOCX6.OCX"
Begin VB.Form fFlood 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Automatic resizing and flooding"
   ClientHeight    =   5040
   ClientLeft      =   2310
   ClientTop       =   1905
   ClientWidth     =   8700
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5040
   ScaleWidth      =   8700
   Begin vsOcx6LibCtl.vsElastic VSElastic1 
      Height          =   2595
      Index           =   1
      Left            =   600
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1245
      Width           =   7545
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   1
      MousePointer    =   0
      _ConvInfo       =   -1
      Version         =   600
      BackColor       =   -2147483633
      ForeColor       =   16711680
      FloodColor      =   192
      ForeColorDisabled=   -2147483631
      Caption         =   "Nest Elastics that automatically resize without writing any code."
      Align           =   5
      Appearance      =   1
      AutoSizeChildren=   0
      BorderWidth     =   12
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   4
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   0   'False
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   0
      GridCols        =   0
      _GridInfo       =   ""
   End
   Begin vsOcx6LibCtl.vsElastic Flood 
      Height          =   600
      Index           =   0
      Left            =   600
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   645
      Width           =   7545
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   1
      MousePointer    =   0
      _ConvInfo       =   -1
      Version         =   600
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   65535
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   1
      Appearance      =   1
      AutoSizeChildren=   0
      BorderWidth     =   6
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   1
      FloodPercent    =   25
      CaptionPos      =   4
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   0   'False
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   0
      GridCols        =   0
      _GridInfo       =   ""
   End
   Begin vsOcx6LibCtl.vsElastic StatusBar 
      Height          =   630
      Left            =   600
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3840
      Width           =   7545
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   1
      MousePointer    =   0
      _ConvInfo       =   -1
      Version         =   600
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   192
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   2
      Appearance      =   1
      AutoSizeChildren=   2
      BorderWidth     =   6
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   1
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   0   'False
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   0
      GridCols        =   0
      _GridInfo       =   ""
      Begin vsOcx6LibCtl.vsElastic VSElastic3 
         Height          =   450
         Index           =   1
         Left            =   5025
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   90
         Width           =   2430
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   1
         MousePointer    =   0
         _ConvInfo       =   -1
         Version         =   600
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FloodColor      =   192
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         Appearance      =   1
         AutoSizeChildren=   1
         BorderWidth     =   6
         ChildSpacing    =   6
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   1
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   0
         TagSplit        =   0   'False
         PicturePos      =   4
         CaptionStyle    =   0
         ResizeFonts     =   0   'False
         GridRows        =   0
         GridCols        =   0
         _GridInfo       =   ""
         Begin VB.CommandButton Command1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "B"
            Height          =   270
            Index           =   9
            Left            =   1650
            TabIndex        =   12
            Top             =   90
            Width           =   690
         End
         Begin VB.CommandButton Command1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "B"
            Height          =   270
            Index           =   8
            Left            =   870
            TabIndex        =   11
            Top             =   90
            Width           =   690
         End
         Begin VB.CommandButton Command1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "B"
            Height          =   270
            Index           =   7
            Left            =   90
            TabIndex        =   10
            Top             =   90
            Width           =   690
         End
      End
      Begin vsOcx6LibCtl.vsElastic VSElastic2 
         Height          =   450
         Index           =   1
         Left            =   90
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   90
         Width           =   4875
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   1
         MousePointer    =   0
         _ConvInfo       =   -1
         Version         =   600
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FloodColor      =   192
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         Appearance      =   1
         AutoSizeChildren=   1
         BorderWidth     =   6
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   1
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   0
         TagSplit        =   0   'False
         PicturePos      =   4
         CaptionStyle    =   0
         ResizeFonts     =   0   'False
         GridRows        =   0
         GridCols        =   0
         _GridInfo       =   ""
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   2
            Left            =   90
            TabIndex        =   15
            Text            =   "When the form is resized, check this text box and the B buttons."
            Top             =   90
            Width           =   4695
         End
      End
   End
   Begin vsOcx6LibCtl.vsElastic Flood 
      Height          =   3825
      Index           =   2
      Left            =   0
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   645
      Width           =   600
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   1
      MousePointer    =   0
      _ConvInfo       =   -1
      Version         =   600
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   16711935
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   3
      Appearance      =   1
      AutoSizeChildren=   0
      BorderWidth     =   6
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   3
      FloodPercent    =   25
      CaptionPos      =   4
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   0   'False
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   0
      GridCols        =   0
      _GridInfo       =   ""
   End
   Begin vsOcx6LibCtl.vsElastic Flood 
      Height          =   570
      Index           =   1
      Left            =   0
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   4470
      Width           =   8145
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   1
      MousePointer    =   0
      _ConvInfo       =   -1
      Version         =   600
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   16776960
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   2
      Appearance      =   1
      AutoSizeChildren=   0
      BorderWidth     =   6
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   2
      FloodPercent    =   25
      CaptionPos      =   4
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   0   'False
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   0
      GridCols        =   0
      _GridInfo       =   ""
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4095
      Top             =   3600
   End
   Begin vsOcx6LibCtl.vsElastic Toolbar 
      Height          =   645
      Left            =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   8145
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   1
      MousePointer    =   0
      _ConvInfo       =   -1
      Version         =   600
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   192
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   1
      Appearance      =   1
      AutoSizeChildren=   2
      BorderWidth     =   6
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   1
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   0   'False
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   0
      GridCols        =   0
      _GridInfo       =   ""
      Begin vsOcx6LibCtl.vsElastic VSElastic2 
         Height          =   465
         Index           =   0
         Left            =   90
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   90
         Width           =   4980
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   1
         MousePointer    =   0
         _ConvInfo       =   -1
         Version         =   600
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FloodColor      =   192
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         Appearance      =   1
         AutoSizeChildren=   1
         BorderWidth     =   6
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   1
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   0
         TagSplit        =   0   'False
         PicturePos      =   4
         CaptionStyle    =   0
         ResizeFonts     =   0   'False
         GridRows        =   0
         GridCols        =   0
         _GridInfo       =   ""
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   90
            TabIndex        =   14
            Text            =   "When the form is resized, check this text box and the A buttons."
            Top             =   90
            Width           =   4800
         End
      End
      Begin vsOcx6LibCtl.vsElastic VSElastic3 
         Height          =   465
         Index           =   0
         Left            =   5130
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   90
         Width           =   2925
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   1
         MousePointer    =   0
         _ConvInfo       =   -1
         Version         =   600
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FloodColor      =   192
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         Appearance      =   1
         AutoSizeChildren=   1
         BorderWidth     =   6
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   1
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   0
         TagSplit        =   0   'False
         PicturePos      =   4
         CaptionStyle    =   0
         ResizeFonts     =   0   'False
         GridRows        =   0
         GridCols        =   0
         _GridInfo       =   ""
         Begin VB.CommandButton Command1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "A"
            Height          =   285
            Index           =   4
            Left            =   90
            TabIndex        =   7
            Top             =   90
            Width           =   885
         End
         Begin VB.CommandButton Command1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "A"
            Height          =   285
            Index           =   2
            Left            =   1035
            TabIndex        =   6
            Top             =   90
            Width           =   870
         End
         Begin VB.CommandButton Command1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "A"
            Height          =   285
            Index           =   0
            Left            =   1965
            TabIndex        =   5
            Top             =   90
            Width           =   870
         End
      End
   End
   Begin vsOcx6LibCtl.vsElastic Flood 
      Height          =   5040
      Index           =   3
      Left            =   8145
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   555
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   1
      MousePointer    =   0
      _ConvInfo       =   -1
      Version         =   600
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   65280
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   4
      Appearance      =   1
      AutoSizeChildren=   0
      BorderWidth     =   6
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   4
      FloodPercent    =   25
      CaptionPos      =   4
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   0   'False
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   0
      GridCols        =   0
      _GridInfo       =   ""
   End
End
Attribute VB_Name = "fFlood"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Timer1_Timer()
  Dim j%
  Static i%

  If i > 97 Then i = 0
  i = i + 3

  For j = 0 To 3
    Flood(j).FloodPercent = i
    Flood(j).Caption = i & "%"
  Next j

End Sub

