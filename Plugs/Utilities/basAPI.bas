Attribute VB_Name = "basAPI"
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
Public Declare Function SetParent Lib "user32" ( _
         ByVal hWndChild As Long, _
         ByVal hWndNewParent As Long) _
      As Long

Public Declare Function GetParent Lib "user32" ( _
         ByVal hwnd As Long) _
      As Long

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
         ByVal hwnd As Long, _
         ByVal nIndex As Long, _
         ByVal dwNewLong As Long _
      ) As Long

Public Declare Function GetWindowLong Lib "user32" _
         Alias "GetWindowLongA" (ByVal hwnd As Long, _
         ByVal nIndex As Long _
      ) As Long



Public Const GWL_WNDPROC = (-4)
Public Const GWL_HINSTANCE = (-6)
Public Const GWL_HWNDPARENT = (-8)
Public Const GWL_STYLE = (-16)
Public Const GWL_EXSTYLE = (-20)
Public Const GWL_USERDATA = (-21)
Public Const GWL_ID = (-12)
Public Const WS_CHILD = &H40000000


