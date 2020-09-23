VERSION 5.00
Begin VB.UserControl LED20 
   BackColor       =   &H80000008&
   ClientHeight    =   645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2985
   LockControls    =   -1  'True
   ScaleHeight     =   645
   ScaleWidth      =   2985
   Begin VB.Line Wire 
      BorderColor     =   &H00E0E0E0&
      Index           =   19
      X1              =   2805
      X2              =   2805
      Y1              =   240
      Y2              =   400
   End
   Begin VB.Line Wire 
      BorderColor     =   &H00E0E0E0&
      Index           =   18
      X1              =   2670
      X2              =   2670
      Y1              =   240
      Y2              =   400
   End
   Begin VB.Line Wire 
      BorderColor     =   &H00E0E0E0&
      Index           =   17
      X1              =   2535
      X2              =   2535
      Y1              =   240
      Y2              =   400
   End
   Begin VB.Line Wire 
      BorderColor     =   &H00E0E0E0&
      Index           =   16
      X1              =   2400
      X2              =   2400
      Y1              =   240
      Y2              =   400
   End
   Begin VB.Line Wire 
      BorderColor     =   &H00E0E0E0&
      Index           =   15
      X1              =   2265
      X2              =   2265
      Y1              =   240
      Y2              =   400
   End
   Begin VB.Line Wire 
      BorderColor     =   &H00E0E0E0&
      Index           =   14
      X1              =   2130
      X2              =   2130
      Y1              =   240
      Y2              =   400
   End
   Begin VB.Line Wire 
      BorderColor     =   &H00E0E0E0&
      Index           =   13
      X1              =   1995
      X2              =   1995
      Y1              =   240
      Y2              =   400
   End
   Begin VB.Line Wire 
      BorderColor     =   &H00E0E0E0&
      Index           =   12
      X1              =   1860
      X2              =   1860
      Y1              =   240
      Y2              =   400
   End
   Begin VB.Line Wire 
      BorderColor     =   &H00E0E0E0&
      Index           =   11
      X1              =   1725
      X2              =   1725
      Y1              =   240
      Y2              =   400
   End
   Begin VB.Line Wire 
      BorderColor     =   &H00E0E0E0&
      Index           =   10
      X1              =   1590
      X2              =   1590
      Y1              =   240
      Y2              =   400
   End
   Begin VB.Line Wire 
      BorderColor     =   &H00E0E0E0&
      Index           =   9
      X1              =   1455
      X2              =   1455
      Y1              =   240
      Y2              =   400
   End
   Begin VB.Line Wire 
      BorderColor     =   &H00E0E0E0&
      Index           =   8
      X1              =   1320
      X2              =   1320
      Y1              =   240
      Y2              =   400
   End
   Begin VB.Line Wire 
      BorderColor     =   &H00E0E0E0&
      Index           =   7
      X1              =   1185
      X2              =   1185
      Y1              =   240
      Y2              =   400
   End
   Begin VB.Line Wire 
      BorderColor     =   &H00E0E0E0&
      Index           =   6
      X1              =   1050
      X2              =   1050
      Y1              =   240
      Y2              =   400
   End
   Begin VB.Line Wire 
      BorderColor     =   &H00E0E0E0&
      Index           =   5
      X1              =   915
      X2              =   915
      Y1              =   240
      Y2              =   400
   End
   Begin VB.Line Wire 
      BorderColor     =   &H00E0E0E0&
      Index           =   4
      X1              =   780
      X2              =   780
      Y1              =   240
      Y2              =   400
   End
   Begin VB.Line Wire 
      BorderColor     =   &H00E0E0E0&
      Index           =   3
      X1              =   645
      X2              =   645
      Y1              =   240
      Y2              =   400
   End
   Begin VB.Line Wire 
      BorderColor     =   &H00E0E0E0&
      Index           =   2
      X1              =   510
      X2              =   510
      Y1              =   240
      Y2              =   400
   End
   Begin VB.Line Wire 
      BorderColor     =   &H00E0E0E0&
      Index           =   1
      X1              =   375
      X2              =   375
      Y1              =   240
      Y2              =   400
   End
   Begin VB.Line Wire 
      BorderColor     =   &H00E0E0E0&
      Index           =   0
      X1              =   240
      X2              =   240
      Y1              =   240
      Y2              =   400
   End
End
Attribute VB_Name = "LED20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Just a quickie LED UserControl
'LED20 because it uses 20 'Wires' to display characters (numbered 0 to 19 in case you're wondering)
'
'NEW in V2
'Changed Font consturction and display routines to use arrays
'added an Enum to simplify creating new fonts
'changed Property Mode to better name: CharSet
'changed InArray to InArray2 which handles pre-defined arrays
'uses the Count property of Wire control array rather than hard coded values (this potentially allows you to add new Wires
'
'This is a spin off from the LED clock challange but definitely not a contender (way too many lines)
'If you look at the control in IDE you'll see that there has been no attempt to position the wires ahead of time
'In case you do want to edit it please note that the form has been locked to stop unnecessary fiddling about.
'
'PROPERTIES
'Value      = Letter(Ucase only)/Number to display set to "" <blank string> for all Wires off
'CharSet     = eBlock| eSimple| EStyles the 3 character sets available (see Procdures below)
'BackColour = backfield of control
'WireColor  = colour to use for Off Wires. Usually a duller version of the ForColor
'             Set to match BackColor and OffShadow = False for invisible
'ForeColor  = colour to display Value/On Wires
'ONWidth    = how thick the OnWires wires are (2 to ???( depends on size of control))
'OFFShadow  = True  = Off Wires overlay On Wires (more machine look)
'           = False = On Wires overlay Off Wires (smoother characters)
'
'see UserControl_Initialize for default values
'
'PROCEDURES
'LiteUp  = switchboard decides which Font to use
'FontBlockyBasic = Initialise LEDWireData to the basic 7 line numerals (not activated see sub for details)
'FontBlocky  = Initialise LEDWireData to the 20 line chracters used by CharSet = eBlock
'FontSimple  = Initialise LEDWireData to the 20 line chracters used by CharSet = eSimple
'FontStyled  = Initialise LEDWireData to the 20 line chracters used by CharSet = eStyled
'SetLine = makes positioning the lines easer by making code more readable(called by 'UserControl_Resize')
'InArray2 = simple test that a value is in a pre-defined array
'DoOnColour = Change Colour of On Wires to ForeColour
'UserControl_Initialize = Set intial appearance when you create a new control
'UserControl_Resize = layout and size lines.
'
'Suggestions/improvements welcome
'
'ADDING/EDITING CHARACTER SETS
'see 'FontSimple' for how to create your own character generating Procedure
'
'If you craft your own set of characters
'and send them to me I'll add them to future versions
'
Private m_value           As String
Public Enum blkCharSet
  eBlock
  eSimple
  eStyled
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private eBlock, eSimple, eStyled
#End If
'this enum makes it a bit easier to create new fonts
Private Enum eWireNum
  eTopL
  eTopR
  eRightT
  eRightB
  eBotR
  eBotL
  eLeftB
  eLeftT
  eMidT
  eMidB
  eMidL
  eMidR
  eCto11
  eCto4
  eCto2
  eCto7
  e9to12
  e12to3
  e3to6
  e6to9
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private eTopL, eTopR, eRightT, eRightB, eBotR, eBotL, eLeftB, eLeftT, eMidT, eMidB, eMidL, eMidR, eCto11, eCto4, eCto2, eCto7, e9to12, e12to3
Private e3to6, e6to9
#End If
Private m_FColour         As OLE_COLOR
Private m_WColour         As OLE_COLOR
Private m_OnWidth         As Long
Private m_OFFShadow       As Boolean
Private m_CharSet         As blkCharSet
Private LEDWireData()     As Variant

Public Property Get BackColor() As OLE_COLOR

  BackColor = UserControl.BackColor

End Property

Public Property Let BackColor(ByVal NewCol As OLE_COLOR)

  UserControl.BackColor = NewCol
  PropertyChanged "BCol"

End Property

Public Property Get CharSet() As blkCharSet

  CharSet = m_CharSet

End Property

Public Property Let CharSet(NewCharSet As blkCharSet)

  m_CharSet = NewCharSet
  PropertyChanged "Char"

End Property

Private Sub ClearData()

  'clears the data arrays before setting a new font
  
  Dim I As Long

  For I = 0 To Wire.Count - 1
    LEDWireData(I) = Array()
  Next I

End Sub

Private Sub DoOnColour()

  Dim I As Long

  For I = 0 To Wire.Count - 1
    Wire(I).BorderColor = WireColor
    If m_OFFShadow Then
      Wire(I).ZOrder 0
    End If
  Next I
  For I = 0 To Wire.Count - 1
    If Wire(I).BorderWidth > 1 Then
      Wire(I).BorderColor = ForeColor
      If Not m_OFFShadow Then
        Wire(I).ZOrder 0
      End If
    End If
  Next I

End Sub

Private Sub FontBlocky()

  'numbers only the classic LED numerals

  ClearData
  LEDWireData(eTopL) = Array(0, 2, 3, 5, 6, 7, 8, 9, "A", "B", "C", "D", "E", "F", "G", "M", "O", "P", "Q", "R", "S", "T", "Z")
  LEDWireData(eTopR) = Array(0, 2, 3, 5, 6, 7, 8, 9, "A", "B", "C", "D", "E", "F", "G", "M", "O", "P", "Q", "R", "S", "T", "Z")
  '
  LEDWireData(eRightT) = Array(0, 1, 2, 3, 4, 7, 8, 9, "A", "B", "D", "H", "I", "J", "M", "N", "O", "P", "Q", "R", "U", "V", "W", _
                               "Y", "Z", "X")
  LEDWireData(eRightB) = Array(0, 1, 3, 4, 5, 6, 7, 8, 9, "A", "B", "D", "G", "H", "I", "J", "K", "M", "N", "O", "Q", "S", "U", _
                               "W")
  '
  LEDWireData(eBotR) = Array(0, 2, 3, 5, 6, 8, 9, "B", "C", "D", "E", "J", "L", "N", "O", "S", "U", "W", "Z")
  LEDWireData(eBotL) = Array(0, 2, 3, 5, 6, 8, 9, "B", "C", "D", "E", "J", "L", "O", "S", "U", "V", "W", "Z")
  '
  LEDWireData(eLeftB) = Array(0, 2, 6, 8, "A", "B", "C", "D", "E", "F", "H", "J", "K", "L", "M", "N", "O", "P", "R", "U", "V", _
                              "W", "Z", "Y", "X")
  LEDWireData(eLeftT) = Array(0, 4, 5, 6, 8, 9, "A", "B", "C", "D", "E", "F", "G", "H", "K", "L", "M", "N", "O", "P", "Q", "R", _
                              "S", "U", "V", "W", "Y", "X")
  '
  LEDWireData(eMidT) = Array("K", "M", "T")
  LEDWireData(eMidB) = Array("M", "N", "R", "T", "V", "W", "X")
  '
  LEDWireData(eMidL) = Array(2, 3, 4, 5, 6, 8, 9, "A", "B", "E", "F", "G", "H", "K", "N", "P", "Q", "R", "S", "Y", "Z", "X")
  LEDWireData(eMidR) = Array(2, 3, 4, 5, 6, 8, 9, "A", "B", "E", "F", "G", "H", "K", "P", "Q", "R", "S", "V", "Y", "Z", "X")

End Sub

Private Sub FontBlockyBasic()

  'numbers only; the classic LED numerals
  '(Not Used)
  'If you want the smallest possible UserControl to add to a project
  'delete the other fonts and adjust LiteUp to call only this procedure
  ClearData
  LEDWireData(eTopL) = Array(0, 2, 3, 5, 6, 7, 8, 9)
  LEDWireData(eTopR) = Array(0, 2, 3, 5, 6, 7, 8, 9)
  LEDWireData(eRightT) = Array(0, 1, 2, 3, 4, 7, 8, 9)
  LEDWireData(eRightB) = Array(0, 1, 3, 4, 5, 6, 7, 8, 9)
  LEDWireData(eBotR) = Array(0, 2, 3, 5, 6, 8, 9)
  LEDWireData(eBotL) = Array(0, 2, 3, 5, 6, 8, 9)
  LEDWireData(eLeftB) = Array(0, 2, 6, 8)
  LEDWireData(eLeftT) = Array(0, 4, 5, 6, 8, 9)
  LEDWireData(eMidL) = Array(2, 3, 4, 5, 6, 8, 9)
  LEDWireData(eMidR) = Array(2, 3, 4, 5, 6, 8, 9)

End Sub

Private Sub FontSimple()

  ClearData ' make sure you clear out the old data before setting a new font
  '           otherwise members of LEDWireData not reset will remain in the array
  'BUILD YOUR OWN FONT
  '
  '1. Create a copy of 'FontSimple' changing the name to suit
  '2. Add an Enum member to Enum blkCharSet ; use the format 'e'+FontName
  '3. Add a 'Case 'e'+FontName to 'LiteUp' that calls your font
  '
  '4. To test your font reset one of the LED20's at the bottom of the demo to use it
  '
  '5. For each character use the chart to work out which lines to turn on
  '   (apologies for the messy ordering system, it was an on-the-fly idea,
  '    roughly clock-wise for each group, Egdes, MidLines, Radials, Daigonals)
  '   and add the character to that line's activating array
  '  NOTE LED20 is programmed to use only Upper Case letters
  '  search for Ucase and make any changes you want to use mixed cases.
  '  But the supplied fonts will not recognize lower case characters and return blanks
  '
  '     ---TopL----'--TopR----'
  '    | \       / | \       /|
  '    |  \  9to12 | 12to3  / |
  'LeftT   \   /   |    \  /  RightT
  '    |    \ /   MidT   \/   |
  '    |    /\     |     /\   |
  '    |   / Cto11 | Cto2  \  |
  '    |  /      \ |        \ |
  '    | /        \| /       \|
  '    -----MidL---'---MidR---'
  '    |\         /|\        /|
  '    | \       / | \      / |
  '    |  \   Cto7 |  Cto4 /  |
  '    |   \  /    |   \  /   |
  'LeftB    \/    MidB  \/     RightB
  '    |    /\     |    /\    |
  '    |   /  \    |   /  \   |
  '    |  /  6to9  | 3to6  \  |
  '    | /      \  |/       \ |
  '     ----BotL---'---BotR---'
  LEDWireData(eTopL) = Array(3, 7, 8, "B", "D", "P", "R", "T", "Z")
  LEDWireData(eTopR) = Array(0, 2, 3, 5, 6, 7, 8, 9, "B", "C", "E", "F", "G", "R", "S", "T", "Z")
  LEDWireData(eRightT) = Array(0, 2, 8, 9, "H", "M", "N", "U", "V", "W")
  LEDWireData(eRightB) = Array("A", "G", "H", "M", "N", "W")
  LEDWireData(eBotR) = Array(2, 3, 5, 6, 8, "B", "C", "E", "L", "Z")
  LEDWireData(eBotL) = Array(0, 2, 3, 5, 6, 8, 9, "B", "D", "L", "U", "S", "Z")
  LEDWireData(eLeftB) = Array(0, 6, 8, "A", "B", "D", "F", "H", "K", "L", "M", "N", "P", "U", "R", "W")
  LEDWireData(eLeftT) = Array("B", "D", "H", "K", "L", "M", "N", "P", "R", "U", "V", "W")
  LEDWireData(eMidT) = Array(1, 4, "I", "J", "T")
  LEDWireData(eMidB) = Array(1, 4, "I", "J", "T", "Y")
  LEDWireData(eMidL) = Array(4, 5, 6, 8, 9, "A", "B", "E", "F", "H", "K", "P", "R", "S")
  LEDWireData(eMidR) = Array(2, 4, 8, 9, "A", "E", "F", "H", "P", "S")
  LEDWireData(eCto11) = Array(8, "M", "N", "X", "Y")
  LEDWireData(eCto4) = Array(5, 3, 6, 8, "B", "K", "N", "Q", "R", "W", "X")
  LEDWireData(eCto2) = Array(3, 7, "B", "K", "M", "R", "X", "Y", "Z")
  LEDWireData(eCto7) = Array(2, 7, "W", "X", "Z")
  LEDWireData(e9to12) = Array(0, 2, 4, 5, 6, 9, "A", "C", "E", "F", "G", "O", "Q", "S")
  LEDWireData(e12to3) = Array("A", "D", "O", "P", "Q")
  LEDWireData(e3to6) = Array(0, 9, "D", "G", "O", "Q", "S", "U", "V")
  LEDWireData(e6to9) = Array("C", "E", "G", "J", "O", "Q", "V")

End Sub

Private Sub FontStyled()

  ClearData
  LEDWireData(eTopL) = Array(7, "D", "T")
  LEDWireData(eTopR) = Array(2, 3, 4, 5, 6, 7, 8, 9, 0, "A", "B", "C", "D", "E", "F", "G", "J", "M", "N", "O", "P", "Q", "R", _
                             "S", "T", "Z")
  LEDWireData(eRightT) = Array(2, 3, 9, 0, "A", "B", "D", "H", "J", "M", "N", "O", "P", "Q", "R", "U", "V", "W", "X", "Y")
  LEDWireData(eRightB) = Array(8, 0, "A", "B", "G", "H", "J", "K", "M", "N", "O", "Q", "S", "U", "W")
  LEDWireData(eBotR) = Array(2, 8, 0, "B", "C", "E", "G", "J", "L", "O", "Q", "S", "U", "W", "Z")
  LEDWireData(eBotL) = Array(2, 3, 5, 6, 8, 0, "B", "C", "D", "E", "G", "J", "L", "O", "Q", "S", "U", "W", "Z")
  LEDWireData(eLeftB) = Array(6, 0, "A", "B", "C", "D", "E", "F", "G", "H", "K", "L", "M", "N", "O", "P", "Q", "R", "T", "U", _
                              "W")
  LEDWireData(eLeftT) = Array("D", "K")
  LEDWireData(eMidT) = Array(1, "M")
  LEDWireData(eMidB) = Array(1, 4, 7, "I", "M", "W")
  LEDWireData(eMidL) = Array(4, 5, 6, 8, 9, "A", "B", "E", "F", "H", "K", "P", "R", "S", "X", "Y")
  LEDWireData(eMidR) = Array(2, 3, 4, 5, 6, 8, 9, "A", "B", "E", "F", "H", "K", "P", "R", "S", "X", "Y")
  LEDWireData(eCto4) = Array("Q", "R", "X")
  LEDWireData(eCto2) = Array(4, 7, 8, "I", "K", "Z")
  LEDWireData(eCto7) = Array(2, 8, "J", "X", "Y", "Z")
  LEDWireData(e9to12) = Array(1, 2, 3, 4, 5, 6, 8, 9, 0, "A", "B", "C", "E", "F", "G", "H", "L", "M", "N", "O", "P", "Q", "R", _
                              "S", "T", "U", "V", "W", "X", "Y", "Z")
  LEDWireData(e12to3) = Array()
  LEDWireData(e3to6) = Array(3, 5, 6, 9, "D", "V")
  LEDWireData(e6to9) = Array("V")

End Sub

Public Property Get ForeColor() As OLE_COLOR

  ForeColor = m_FColour

End Property

Public Property Let ForeColor(ByVal NewCol As OLE_COLOR)

  m_FColour = NewCol
  PropertyChanged "FCol"
  DoOnColour

End Property

Private Function InArray2(ByVal Test As String, _
                          arrD As Variant) As Boolean

  'v2 changed to handle new array based data
  
  Dim I As Long

  For I = LBound(arrD) To UBound(arrD)
    If Test = arrD(I) Then
      InArray2 = True
    End If
  Next I

End Function

Private Sub LiteUp()

  If ONWidth Then ' safety to prevent trying to draw if ONWidth has not been set
    Select Case CharSet
     Case eBlock
      FontBlocky
     Case eSimple
      FontSimple
     Case eStyled
      FontStyled
    End Select
    SetWires
  End If

End Sub

Public Property Get OFFShadow() As Boolean

  OFFShadow = m_OFFShadow

End Property

Public Property Let OFFShadow(ByVal BShow As Boolean)

  m_OFFShadow = BShow

End Property

Public Property Get ONWidth() As Long

  ONWidth = m_OnWidth

End Property

Public Property Let ONWidth(ByVal NewWid As Long)

  If NewWid > 0 Then
    m_OnWidth = NewWid
    PropertyChanged "ONWid"
    LiteUp
  End If

End Property

Private Sub SetLine(ByVal lno As Long, _
                    ByVal w1 As Long, _
                    ByVal h1 As Long, _
                    ByVal w2 As Long, _
                    ByVal h2 As Long, _
                    Optional ByVal Col As Long = vbWhite)

  'this just makes it easier to see what is going on in UserControl_Resize

  With Wire(lno)
    .X1 = w1
    .Y1 = h1
    .X2 = w2
    .Y2 = h2
    .BorderColor = Col
  End With

End Sub

Private Sub SetWires()

  'v2 sets the WireWidth from the FontData
  ' does all the settings of wire in one procedure
  
  Dim I   As Long
  Dim bOn As Boolean

  For I = 0 To Wire.Count - 1
    bOn = InArray2(m_value, LEDWireData(I))
    With Wire(I)
      .BorderWidth = IIf(bOn, ONWidth, 1)
      .BorderColor = IIf(bOn, ForeColor, WireColor)
      .ZOrder IIf(m_OFFShadow, 0, 1)
    End With 'Wire(I)
  Next I

End Sub

Private Sub UserControl_Initialize()

  ReDim LEDWireData(Wire.Count - 1) As Variant
  ForeColor = vbGreen
  WireColor = &HC000&
  BackColor = vbBlack
  CharSet = eSimple
  ONWidth = 3
  OFFShadow = True

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

  With PropBag
    ForeColor = .ReadProperty("FCol", vbGreen)
    WireColor = .ReadProperty("WCol", &HC000&)
    BackColor = .ReadProperty("BCol", vbBlack)
    OFFShadow = .ReadProperty("Shad", True)
    ONWidth = .ReadProperty("OnWid", 3)
    CharSet = .ReadProperty("Char", eBlock)
    Value = .ReadProperty("Val", " ")
  End With 'PropBag

End Sub

Private Sub UserControl_Resize()

  'position and size the lines
  
  Dim Hi   As Long
  Dim Wid  As Long
  Dim HMin As Long
  Dim WMin As Long
  Dim HMax As Long
  Dim WMax As Long
  Dim WMid As Long
  Dim HMid As Long

  Hi = UserControl.Height
  Wid = UserControl.Width
  HMin = Hi * 0.05
  WMin = Wid * 0.05
  HMax = Hi - HMin
  WMax = Wid - WMin
  WMid = Wid / 2
  HMid = Hi / 2
  'Edges
  SetLine eTopL, WMin, HMin, WMid, HMin
  SetLine eTopR, WMid, HMin, WMax, HMin
  SetLine eRightT, WMax, HMin, WMax, HMid
  SetLine eRightB, WMax, HMid, WMax, HMax
  SetLine eBotR, WMax, HMax, WMid, HMax
  SetLine eBotL, WMid, HMax, WMin, HMax
  SetLine eLeftB, WMin, HMax, WMin, HMid
  SetLine eLeftT, WMin, HMid, WMin, HMin
  'Midlines
  SetLine eMidT, WMid, HMin, WMid, HMid
  SetLine eMidR, WMid, HMid, WMax, HMid
  SetLine eMidB, WMid, HMid, WMid, HMax
  SetLine eMidL, WMin, HMid, WMid, HMid
  'Radials
  SetLine eCto11, WMin, HMin, WMid, HMid
  SetLine eCto2, WMax, HMin, WMid, HMid
  SetLine eCto4, WMid, HMid, WMax, HMax
  SetLine eCto7, WMid, HMid, WMin, HMax
  'Diagonals
  SetLine e9to12, WMin, HMid, WMid, HMin
  SetLine e12to3, WMid, HMin, WMax, HMid
  SetLine e3to6, WMax, HMid, WMid, HMax
  SetLine e6to9, WMid, HMax, WMin, HMid

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

  With PropBag
    .WriteProperty "FCol", ForeColor, vbGreen
    .WriteProperty "WCol", WireColor, &HC000&
    .WriteProperty "BCol", BackColor, vbBlack
    .WriteProperty "Val", Value, " "
    .WriteProperty "OnWid", ONWidth, 3
    .WriteProperty "Shad", OFFShadow, True
    .WriteProperty "Char", CharSet, eBlock
  End With 'PropBag

End Sub

Public Property Get Value() As String

  Value = m_value

End Property

Public Property Let Value(ByVal NewValue As String)

  m_value = UCase$(NewValue)
  LiteUp
  PropertyChanged "Val"

End Property

Public Property Get WireColor() As OLE_COLOR

  WireColor = m_WColour

End Property

Public Property Let WireColor(ByVal NewCol As OLE_COLOR)

  m_WColour = NewCol
  PropertyChanged "WCol"
  DoOnColour

End Property

':)Code Fixer V2.8.9 (22/01/2005 9:47:41 AM) 87 + 426 = 513 Lines Thanks Ulli for inspiration and lots of code.

