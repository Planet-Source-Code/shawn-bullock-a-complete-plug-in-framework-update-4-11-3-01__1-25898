Attribute VB_Name = "basMain"
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
Public Application As Application
Public gAppVisible As Boolean
Public gAppEvents As Events

Public gMain As frmMain


Public Sub Main()
   '
   ' If we are already running, don't launch a second instance
   '
   Delay 0.5
   
   If (App.PrevInstance) Then
      End
   End If
   
   ' Are we embedded?
   '
   Select Case Command
      Case "-Embedding"
         '
         ' When you call CreateObject(), Windows passes "-Embedding" as a command-line
         '  parameter so you're application can handle it accordingly
         '
         gAppVisible = False
         
      Case Else
         gAppVisible = True
      
   End Select
   
   Set Application = New Application
End Sub

Public Sub Exit_Application()
   gAppEvents.Send_Quit
   Application.Plugins.UnLoadPlugins
   
   Set Application = Nothing
   Set gAppEvents = Nothing
   End
End Sub
