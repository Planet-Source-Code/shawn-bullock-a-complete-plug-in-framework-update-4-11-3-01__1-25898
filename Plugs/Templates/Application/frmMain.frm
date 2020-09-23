VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "HostApplication1"
   ClientHeight    =   2790
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5865
   LinkTopic       =   "Form1"
   ScaleHeight     =   2790
   ScaleWidth      =   5865
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mbrFile 
      Caption         =   "&File"
      Begin VB.Menu miFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu miToolsModules 
         Caption         =   "&Modules"
         Begin VB.Menu miModules 
            Caption         =   "&Module Manager..."
            Index           =   0
         End
         Begin VB.Menu miModules 
            Caption         =   "-"
            Index           =   1
         End
      End
      Begin VB.Menu miToolsBr1 
         Caption         =   "-"
      End
      Begin VB.Menu miToolsSettings 
         Caption         =   "&Settings..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MODULE_MANAGER As Long = 0
Private Const MODULE_BREAK As Long = 1


Public Function AddModuleToModules( _
      isModuleName As String, _
      isModuleIdentifier As String _
   ) As Long
   
   ' Not required for the framework.  Only provided as a means
   '  to demonstrate one of many ways to add a module the
   '  menus.
   '
   Dim lnCount As Long
   
   For lnCount = 0 To miModules.Count - 1
      '
      ' Make sure it doesn't already exist
      '
      If (miModules.Item(lnCount).Tag = isModuleIdentifier) Then
         AddModuleToModules = -(PLUG_ERROR_MODULE_ALREADY_EXISTS)
         Exit Function
      End If
   Next
   
   lnCount = (miModules.Count)
   
   ' Hook it
   '
   Load miModules(lnCount)
   
   ' Name it
   '
   miModules(lnCount).Caption = isModuleName
   miModules(lnCount).Tag = isModuleIdentifier
   miModules(lnCount).Visible = True
   
   AddModuleToModules = lnCount
End Function
   
Public Function RemoveModuleFromModules( _
      isModuleIdentifier _
   ) As Long
      
   ' Not specifically required for the framework.  Only provided
   '  as a sample to demonstrate one of many ways to hook a
   '  module into the application, er, unhook the previously
   '  hooked.
   '
   Dim lnCount As Long
   
   DoEvents
   For lnCount = 0 To miModules.Count - 1
      If (miModules(lnCount).Tag = isModuleIdentifier) Then
         Unload miModules(lnCount)
      End If
   Next
   
End Function




Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If (Not Application Is Nothing) Then
      Application.Quit
   End If
End Sub

Private Sub miFileExit_Click()
   If (Not Application Is Nothing) Then
      Application.Quit
   End If
End Sub

Private Sub miModules_Click(Index As Integer)
   '
   ' Not required for the framework.  This is a means simply
   '  for the application to deal with a selected menu item
   '  that was selected in the modules menu.  This means a
   '  module was invoked.
   '
   Select Case Index
      Case MODULE_MANAGER
         Application.Plugins.ShowModuleManager
         
      Case Else
         '
         ' Module functionality was invoked, deal with it.
         '
         gAppEvents.Send_ModuleSelected miModules(Index).Tag
         
   End Select
   
End Sub

Private Sub miToolsSettings_Click()
   Application.Settings.ShowSettings
End Sub
