VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Application"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private WithEvents UtilityEvents As PlugUtilities.UtilityEvents
Attribute UtilityEvents.VB_VarHelpID = -1

Public Settings As PlugUtilities.Settings
Public Plugins As PlugUtilities.Plugins

Private mnGenKey As Long

' Node Keys for the settings manager.  There should be a key for
'  each node group, and a key for each node within the group.
'  You will only be able to intercept an event in the settings
'  manager by the key value.
'
Private msApplicationKey As String
Private msHostApplicationGeneralKey As String



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

Public Sub PopulateSettings()
   msHostApplicationGeneralKey = "General"
   msApplicationKey = "Application"
   
   ' Add the editor node to the treeview
   '
   Settings.AddNode msApplicationKey, "General", msHostApplicationGeneralKey
     
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
   Set UtilityEvents = Nothing
   Set gAppEvents = Nothing
   Set Settings = Nothing
   Set Plugins = Nothing
   
   Unload gMain
   Set gMain = Nothing
End Sub

Private Sub UtilityEvents_SelectSetting(Setting As String)
   Select Case Setting
      Case msHostApplicationGeneralKey
         Application.Settings.SetActiveSetting frmConfigGeneral
      
      Case "About:Application"
         Application.Settings.SetActiveSetting frmSettingAbout
         
   End Select
   
End Sub


