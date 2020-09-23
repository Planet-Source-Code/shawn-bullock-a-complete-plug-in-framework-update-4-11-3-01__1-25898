VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Application"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Plugins"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member1" ,"Settings"
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
Private WithEvents UtilityEvents As PlugUtilities.UtilityEvents
Attribute UtilityEvents.VB_VarHelpID = -1

Public Settings As PlugUtilities.Settings
Public Plugins As PlugUtilities.Plugins
Public ActiveDocument As ActiveDocument

Private mnGenKey As Long

Private msEditorKey As String
Private msApplicationKey As String


Public Sub Quit()
   Exit_Application
End Sub

Public Property Let Visible(Value As Boolean)
   gAppVisible = Value
   Plugins.Visible = Value
   gMain.Visible = gAppVisible
End Property

Public Property Get Visible() As Boolean
   Visible = gAppVisible
End Property

Public Property Get Name() As String
   Name = App.Title
End Property

 
Private Sub PopulateSettings()

   msEditorKey = "Editor"
   msApplicationKey = "Application"

      
   ' Add the editor node to the treeview
   '
   Settings.AddNode msApplicationKey, "Editor", msEditorKey
   
End Sub


Public Function AddModuleToModules( _
      isModuleName As String, _
      isModuleIdentifier As String _
   ) As Long
   
   AddModuleToModules = gMain.AddModuleToModules(isModuleName, isModuleIdentifier)
End Function
   
Public Function RemoveModuleFromModules( _
      isModuleIdentifier _
   ) As Long
   
   RemoveModuleFromModules = gMain.RemoveModuleFromModules(isModuleIdentifier)
End Function

Private Sub Class_Initialize()
   Set ActiveDocument = New ActiveDocument
   Set gAppEvents = New Events
   Set Plugins = New Plugins
   Set gMain = New frmMain
   
   Set UtilityEvents = Plugins.UtilityEvents
   Set Settings = Plugins.Settings
   
   With Plugins
      .ApplicationName = App.Title
      .ApplicationPath = App.Path
      .AppVersion = App.Major & "." & App.Minor & "." & App.Revision
      
      .ApplicationPointer = Me
      .Visible = gAppVisible
      
      .AppEventPointer = gAppEvents
      
      ' Put the settings in order, this must be first before the initialize event
      '  or things won't work correctly
      '
      PopulateSettings
      
      ' Load the plugins, after PopulateSettings because the plugs may need to add
      '  themselves to the Modules node
      '
      .LoadPlugins
      
      Do While (.AllPluginsLoaded = False)
         DoEvents
      Loop
      
      ' Notify all loaded plugins that the application has initialized
      '
      gAppEvents.Send_Initialize
      
   End With
   
   Visible = gAppVisible
End Sub

Private Sub Class_Terminate()
   Set ActiveDocument = Nothing
   Set UtilityEvents = Nothing
   Set gAppEvents = Nothing
   Set Settings = Nothing
   Set Plugins = Nothing
   
   Unload gMain
   Set gMain = Nothing
End Sub


Private Sub UtilityEvents_SelectSetting(Setting As String)
   Select Case Setting
      Case "Editor"
         Beep
   
      Case "About:Application"
         Settings.SetActiveSetting frmAbout
         
   End Select
   
End Sub
