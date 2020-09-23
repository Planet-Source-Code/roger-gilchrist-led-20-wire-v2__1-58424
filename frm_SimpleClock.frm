VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_SimpleClock 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "simple clock with LED20 DEMO"
   ClientHeight    =   5700
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   6165
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   6165
   StartUpPosition =   3  'Windows Default
   Begin prjSimpleClock.LED20 LEDClassic 
      Height          =   1695
      Left            =   720
      TabIndex        =   21
      Top             =   2880
      Width           =   975
      _extentx        =   1720
      _extenty        =   2990
      onwid           =   10
      val             =   "1"
   End
   Begin VB.CommandButton cmdScrollFont 
      Caption         =   ">"
      Height          =   375
      Index           =   1
      Left            =   2640
      TabIndex        =   20
      Top             =   4800
      Width           =   735
   End
   Begin prjSimpleClock.LED20 LEDStyled 
      Height          =   1695
      Left            =   3480
      TabIndex        =   19
      Top             =   2880
      Width           =   975
      _extentx        =   1720
      _extenty        =   2990
      fcol            =   16777088
      wcol            =   12632064
      onwid           =   10
      char            =   2
      val             =   "1"
   End
   Begin VB.CommandButton cmdScrollFont 
      Caption         =   "<"
      Height          =   375
      Index           =   0
      Left            =   1920
      TabIndex        =   18
      Top             =   4800
      Width           =   735
   End
   Begin prjSimpleClock.LED20 LEDSimple 
      Height          =   1695
      Left            =   2160
      TabIndex        =   17
      Top             =   2880
      Width           =   975
      _extentx        =   1720
      _extenty        =   2990
      fcol            =   255
      wcol            =   8438015
      onwid           =   10
      char            =   1
      val             =   "1"
   End
   Begin MSComDlg.CommonDialog cdlSimpleClock 
      Left            =   5640
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin prjSimpleClock.LED20 LEDLogo 
      Height          =   735
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   1560
      Width           =   495
      _extentx        =   450
      _extenty        =   661
      fcol            =   255
      wcol            =   192
      onwid           =   7
      char            =   1
      val             =   "L"
   End
   Begin prjSimpleClock.LED20 LEDClock 
      Height          =   975
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   750
      _extentx        =   1323
      _extenty        =   1720
      fcol            =   8454016
      onwid           =   5
      val             =   ""
   End
   Begin VB.Timer tmrUpdate 
      Interval        =   1000
      Left            =   5640
      Top             =   720
   End
   Begin prjSimpleClock.LED20 LEDClock 
      Height          =   975
      Index           =   1
      Left            =   750
      TabIndex        =   1
      Top             =   120
      Width           =   750
      _extentx        =   1323
      _extenty        =   1720
      fcol            =   8454016
      onwid           =   5
      val             =   ""
   End
   Begin prjSimpleClock.LED20 LEDClock 
      Height          =   975
      Index           =   2
      Left            =   1620
      TabIndex        =   2
      Top             =   120
      Width           =   750
      _extentx        =   1323
      _extenty        =   1720
      fcol            =   8454016
      onwid           =   5
      val             =   ""
   End
   Begin prjSimpleClock.LED20 LEDClock 
      Height          =   975
      Index           =   3
      Left            =   2370
      TabIndex        =   3
      Top             =   120
      Width           =   750
      _extentx        =   1323
      _extenty        =   1720
      fcol            =   8454016
      onwid           =   5
      val             =   ""
   End
   Begin prjSimpleClock.LED20 LEDClock 
      Height          =   735
      Index           =   4
      Left            =   3240
      TabIndex        =   4
      Top             =   120
      Width           =   750
      _extentx        =   1323
      _extenty        =   1296
      fcol            =   8454016
      onwid           =   4
      val             =   ""
   End
   Begin prjSimpleClock.LED20 LEDClock 
      Height          =   735
      Index           =   5
      Left            =   3960
      TabIndex        =   5
      Top             =   120
      Width           =   750
      _extentx        =   1323
      _extenty        =   1296
      fcol            =   8454016
      shad            =   0   'False
      onwid           =   4
      val             =   ""
   End
   Begin prjSimpleClock.LED20 LEDClock 
      Height          =   270
      Index           =   6
      Left            =   4740
      TabIndex        =   6
      Top             =   120
      Width           =   345
      _extentx        =   609
      _extenty        =   476
      fcol            =   8454016
      val             =   "A"
   End
   Begin prjSimpleClock.LED20 LEDClock 
      Height          =   270
      Index           =   7
      Left            =   5205
      TabIndex        =   7
      Top             =   120
      Width           =   345
      _extentx        =   609
      _extenty        =   476
      fcol            =   8454016
      val             =   "M"
   End
   Begin prjSimpleClock.LED20 LEDLogo 
      Height          =   735
      Index           =   1
      Left            =   855
      TabIndex        =   9
      Top             =   1560
      Width           =   495
      _extentx        =   450
      _extenty        =   661
      fcol            =   255
      wcol            =   192
      onwid           =   7
      char            =   1
      val             =   "E"
   End
   Begin prjSimpleClock.LED20 LEDLogo 
      Height          =   735
      Index           =   2
      Left            =   1470
      TabIndex        =   10
      Top             =   1560
      Width           =   495
      _extentx        =   450
      _extenty        =   661
      fcol            =   255
      wcol            =   192
      onwid           =   7
      char            =   1
      val             =   "D"
   End
   Begin prjSimpleClock.LED20 LEDLogo 
      Height          =   735
      Index           =   3
      Left            =   2085
      TabIndex        =   11
      Top             =   1560
      Width           =   495
      _extentx        =   450
      _extenty        =   661
      fcol            =   255
      wcol            =   192
      onwid           =   7
      char            =   1
      val             =   "2"
   End
   Begin prjSimpleClock.LED20 LEDLogo 
      Height          =   735
      Index           =   4
      Left            =   2700
      TabIndex        =   12
      Top             =   1560
      Width           =   495
      _extentx        =   450
      _extenty        =   661
      fcol            =   255
      wcol            =   192
      onwid           =   7
      char            =   1
      val             =   "0"
   End
   Begin prjSimpleClock.LED20 LEDLogo 
      Height          =   735
      Index           =   5
      Left            =   3555
      TabIndex        =   13
      Top             =   1560
      Width           =   495
      _extentx        =   450
      _extenty        =   661
      fcol            =   255
      wcol            =   192
      onwid           =   7
      char            =   2
      val             =   "D"
   End
   Begin prjSimpleClock.LED20 LEDLogo 
      Height          =   735
      Index           =   6
      Left            =   4170
      TabIndex        =   14
      Top             =   1560
      Width           =   495
      _extentx        =   450
      _extenty        =   661
      fcol            =   255
      wcol            =   192
      onwid           =   7
      char            =   2
      val             =   "E"
   End
   Begin prjSimpleClock.LED20 LEDLogo 
      Height          =   735
      Index           =   7
      Left            =   4785
      TabIndex        =   15
      Top             =   1560
      Width           =   495
      _extentx        =   450
      _extenty        =   661
      fcol            =   255
      wcol            =   192
      onwid           =   7
      char            =   2
      val             =   "M"
   End
   Begin prjSimpleClock.LED20 LEDLogo 
      Height          =   735
      Index           =   8
      Left            =   5400
      TabIndex        =   16
      Top             =   1560
      Width           =   495
      _extentx        =   450
      _extenty        =   661
      fcol            =   255
      wcol            =   192
      onwid           =   7
      char            =   2
      val             =   "O"
   End
   Begin VB.Label lblFontName 
      BackStyle       =   0  'Transparent
      Caption         =   "Classic        Simple       Styled"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   22
      Top             =   2520
      Width           =   3855
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuClock 
      Caption         =   "Clock Settings"
      Begin VB.Menu mnuFont 
         Caption         =   "Font"
         Begin VB.Menu mnuFontopt 
            Caption         =   "Classic"
            Index           =   0
         End
         Begin VB.Menu mnuFontopt 
            Caption         =   "Simple"
            Index           =   1
         End
         Begin VB.Menu mnuFontopt 
            Caption         =   "Styled"
            Index           =   2
         End
      End
      Begin VB.Menu mnuCol 
         Caption         =   "Color"
         Begin VB.Menu mnuColOpt 
            Caption         =   "Fore"
            Index           =   0
         End
         Begin VB.Menu mnuColOpt 
            Caption         =   "Back"
            Index           =   1
         End
         Begin VB.Menu mnuColOpt 
            Caption         =   "Wire"
            Index           =   2
         End
      End
      Begin VB.Menu mnuShadow 
         Caption         =   "Shadow ON"
      End
      Begin VB.Menu mnuWid 
         Caption         =   "OnWidth"
         Begin VB.Menu mnuWidOpt 
            Caption         =   "2"
            Index           =   0
         End
         Begin VB.Menu mnuWidOpt 
            Caption         =   "3"
            Index           =   1
         End
         Begin VB.Menu mnuWidOpt 
            Caption         =   "4"
            Index           =   2
         End
         Begin VB.Menu mnuWidOpt 
            Caption         =   "5"
            Index           =   3
         End
         Begin VB.Menu mnuWidOpt 
            Caption         =   "6"
            Index           =   4
         End
         Begin VB.Menu mnuWidOpt 
            Caption         =   "7"
            Index           =   5
         End
         Begin VB.Menu mnuWidOpt 
            Caption         =   "8"
            Index           =   6
         End
         Begin VB.Menu mnuWidOpt 
            Caption         =   "9"
            Index           =   7
         End
         Begin VB.Menu mnuWidOpt 
            Caption         =   "10"
            Index           =   8
         End
      End
   End
End
Attribute VB_Name = "frm_SimpleClock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Const Chars     As String = "1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Private ChNo            As Long

Private Sub cmdScrollFont_Click(Index As Integer)

  Select Case Index
   Case 0
    ChNo = ChNo - 1
    If ChNo < 1 Then
      ChNo = Len(Chars)
    End If
   Case 1
    ChNo = ChNo + 1
    If ChNo > Len(Chars) Then
      ChNo = 1
    End If
  End Select
  LEDClassic.Value = Mid$(Chars, ChNo, 1)
  LEDSimple.Value = Mid$(Chars, ChNo, 1)
  LEDStyled.Value = Mid$(Chars, ChNo, 1)

End Sub

Private Sub Form_Load()

  MenuChecking mnuWidOpt, LEDClock(0).ONWidth - 2
  ChNo = 1

End Sub

Private Sub mnuClock_Click()

  'v2 safer than setting from IDE

  MenuChecking mnuFontopt, LEDClock(0).CharSet

End Sub

Private Sub mnuColOpt_Click(Index As Integer)

  Dim I       As Long
  Dim TestCol As Long

  Select Case Index
   Case 0
    TestCol = LEDClock(0).ForeColor
   Case 1
    TestCol = LEDClock(0).BackColor
   Case 2
    TestCol = LEDClock(0).WireColor
  End Select
  With cdlSimpleClock
    .Flags = cdlCCRGBInit Or cdlCCFullOpen
    .Color = TestCol
    .ShowColor
    TestCol = .Color
  End With
  Select Case Index
   Case 0
    For I = 0 To 7
      LEDClock(I).ForeColor = TestCol
    Next I
   Case 1
    For I = 0 To 7
      LEDClock(I).BackColor = TestCol
    Next I
   Case 2
    For I = 0 To 7
      LEDClock(I).WireColor = TestCol
    Next I
  End Select

End Sub

Private Sub mnuExit_Click()

  tmrUpdate.Enabled = False
  Unload Me

End Sub

Private Sub mnuFontopt_Click(Index As Integer)

  Dim I As Long

  MenuChecking mnuFontopt, Index
  For I = 0 To 5
    LEDClock(I).Value = ""
    LEDClock(I).CharSet = Index
  Next I

End Sub

Private Sub mnuShadow_Click()

  Dim I As Long

  For I = 0 To 7
    LEDClock(I).OFFShadow = Not LEDClock(I).OFFShadow
  Next I
  mnuShadow.Caption = "Shadow " & IIf(LEDClock(0).OFFShadow, "OFF", "ON")

End Sub

Private Sub mnuWidOpt_Click(Index As Integer)

  Dim I As Long

  MenuChecking mnuWidOpt, Index
  For I = 0 To 5
    LEDClock(I).ONWidth = Index + 2
  Next I

End Sub

Private Sub tmrUpdate_Timer()

  Dim RndF     As Long
  Dim RndB     As Long
  Dim RndW     As Long
  Dim RndS     As Boolean
  Dim strTime2 As String
  Dim I        As Long

  strTime2 = Format$(Time, "hh:mm:ss AM/PM")
  LEDClock(0).Value = Left$(strTime2, 1)
  LEDClock(1).Value = Mid$(strTime2, 2, 1)
  LEDClock(2).Value = Mid$(strTime2, 4, 1)
  LEDClock(3).Value = Mid$(strTime2, 5, 1)
  LEDClock(4).Value = Mid$(strTime2, 7, 1)
  LEDClock(5).Value = Mid$(strTime2, 8, 1)
  LEDClock(6).Value = Left$(Right$(strTime2, 2), 1)
  RndF = Rnd * vbWhite
  RndB = Rnd * vbWhite
  RndW = Rnd * vbWhite
  RndS = Rnd > 0.5
  For I = 0 To LEDLogo.Count - 1
    With LEDLogo(I) 'Only affects the count property of the control array
      .ForeColor = RndF
      .WireColor = RndW
      .BackColor = RndB
      .OFFShadow = RndS
    End With 'LEDLogo
  Next I

End Sub

':)Code Fixer V2.8.9 (22/01/2005 9:47:36 AM) 3 + 147 = 150 Lines Thanks Ulli for inspiration and lots of code.
