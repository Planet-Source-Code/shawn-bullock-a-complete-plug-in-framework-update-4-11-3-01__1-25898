VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IPluginHook
Implements IPluginRegister

Private WithEvents AppEvents As Dynamic_Word.Events
Attribute AppEvents.VB_VarHelpID = -1
Private WithEvents UtilityEvents As PlugUtilities.UtilityEvents
Attribute UtilityEvents.VB_VarHelpID = -1

Private gModPlugTest As String

' Only use the host events you need.  You don't need to act on all
'  events that are fired.
'


' Class specfics
'
Private Sub Class_Terminate()
   If (Not Application Is Nothing) Then
      Set Application = Nothing
   End If
End Sub






'
'
' IPluginHook Interface
'
'
'
Public Property Get IPluginHook_Application() As Object
   Set IPluginHook_Application = Application
End Property

Public Property Let IPluginHook_Application(Value As Object)
   Set Application = Value
End Property


Public Property Let IPluginHook_AppEvents(Value As Object)
   Set AppEvents = Value
End Property

Public Property Get IPluginHook_AppEvents() As Object
   Set IPluginHook_AppEvents = AppEvents
End Property


Public Property Let IPluginHook_UtilityEvents(Value As Object)
   Set UtilityEvents = Value
End Property

Public Property Get IPluginHook_UtilityEvents() As Object
   IPluginHook_UtilityEvents = UtilityEvents
End Property


Private Function IPluginHook_HookPlugin() As Long
   With Application.Settings
      gModPlugTest = .AddModuleGroup(App.Title)
      .AddModuleNode "Configure", gModPlugTest, "config" & App.Title
   End With
End Function

Private Function IPluginHook_UnHookPlugin() As Long
   With Application.Settings
      .RemoveModuleGroup gModPlugTest
   End With
End Function



'
'
' IPluginRegister Interface
'
'
'
Public Property Get IPluginRegister_Description() As String
   IPluginRegister_Description = "Periodically save changes automatically."
End Property

Public Property Get IPluginRegister_Display() As String
   IPluginRegister_Display = "AutoSave"
End Property

Public Property Get IPluginRegister_Interface() As String
   IPluginRegister_Interface = "dwAutoSave.Main"
End Property

Public Property Get IPluginRegister_Name() As String
   IPluginRegister_Name = "dwAutoSave"
End Property

Public Property Get IPluginRegister_Path() As String
   IPluginRegister_Path = "<Defualt>"
End Property

Public Property Get IPluginRegister_Startup() As Boolean
   IPluginRegister_Startup = True
End Property

Public Property Get IPluginRegister_Supports() As String
   IPluginRegister_Supports = "1.1.0"
End Property




'
' Event handling for utility events (as opposed to application
'  events.
'
Private Sub UtilityEvents_SelectSetting(Setting As String)
   '
   ' Each case statement should be followed by the key of a previsously
   '  hooked module in the settings manager.
   '
   Select Case Setting
      Case "config" & App.Title
         Application.Settings.SetActiveSetting frmConfig
         
      Case "About:" & App.Title
         Application.Settings.SetActiveSetting frmAbout
         
   End Select
End Sub

' Event handling for utility events (as opposed to application
'  events.
'
Private Sub UtilityEvents_ViewSettings()
   'Message "View Settings"
End Sub

