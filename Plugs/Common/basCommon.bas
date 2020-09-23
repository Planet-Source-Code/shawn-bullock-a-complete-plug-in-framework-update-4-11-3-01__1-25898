Attribute VB_Name = "basCommon"
Option Explicit

' All source code pertaining to this plug-in architecture is copyright(c) 2001
'  Shawn Bullock.  Author reserves all rights and is granting you limited use
'  to use for educational purposes.  You may only redistribute this in its
'  entirety and without modification unless you include a copy of the original
'  work with this notice.  You may not profit or resell this source code.
'
' Portions of this source code is copyright by Steve McMahon
'  www.vbaccellerator.com. and is subject to the terms and conditions set forth
'  by him.
'
' All source code is provided as-is and without warranty of any kind, implied
'  or not.  Author is not responsible for any damage created by it or misuse of
'  source code and functionality intended by the author or not.
'

'
' Common Module
'
'
' Shared with Dynamic Word and PlugUtilities.  Holds certain maintenance and misc.
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
