Attribute VB_Name = "basMain"
Option Explicit

Public Application As Application
Public gAppVisible As Boolean
Public gAppEvents As Events

Public gMain As frmMain


Public Sub Main()
   '
   ' If we are already running, don't launch a second instance
   '
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

