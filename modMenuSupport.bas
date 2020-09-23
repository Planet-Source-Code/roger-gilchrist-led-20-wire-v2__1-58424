Attribute VB_Name = "modMenuSupport"
Option Explicit

Public Sub MenuChecking(Ctrl As Variant, _
                        ByVal chekMe As Long)

  'generic for controlling Checks in menu arrays
  'unchecks all except the selected member
  'requires continuous menu array
  
  Dim I As Long

  For I = 0 To Ctrl.Count - 1
    Ctrl(I).Checked = I = chekMe
  Next I

End Sub

':)Code Fixer V2.8.9 (22/01/2005 9:47:43 AM) 1 + 17 = 18 Lines Thanks Ulli for inspiration and lots of code.

