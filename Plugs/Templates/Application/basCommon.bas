Attribute VB_Name = "basCommon"
Option Explicit

'
' Common Module
'
'
' Shared with Host App and PlugUtilities.  Holds certain maintenance and misc.
'  functions that are shared between projects.
'
'
Public Function X(Value As Long) As Long
   '
   ' Instead of specifying pixels in terms of twips, we specify in terms of pixels,
   '  and do the twip translation dynamically.
   '
   X = (Value * Screen.TwipsPerPixelX)
End Function

Public Function Y(Value As Long) As Long
   '
   ' Instead of specifying pixels in terms of twips, we specify in terms of pixels,
   '  and do the twip translation dynamically.
   '
   Y = (Value * Screen.TwipsPerPixelY)
End Function

Public Sub Delay(inSngSeconds As Single)
   '
   ' Slight delay
   '
   inSngSeconds = (inSngSeconds * 300000)
   
   Do While (inSngSeconds > 0)
      inSngSeconds = (inSngSeconds - 1)
      DoEvents
   Loop
End Sub

