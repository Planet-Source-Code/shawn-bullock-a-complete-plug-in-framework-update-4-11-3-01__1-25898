VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
Implements IPluginHook
Implements IPluginRegister

Private WithEvents AppEvents As Dynamic_Word.Events
Attribute AppEvents.VB_VarHelpID = -1
Private WithEvents UtilityEvents As PlugUtilities.UtilityEvents
Attribute UtilityEvents.VB_VarHelpID = -1

Private gModPlugTest As String


Private Sub AppEvents_Initialize()
   Message "Initialize"
End Sub

Private Sub AppEvents_Quit()
   Message "Quit"
End Sub

Private Sub AppEvents_ViewSettings()
   Message "View Settings"
End Sub

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
      gModPlugTest = .AddModuleGroup("PlugTest")
      .AddModuleNode "Test1", gModPlugTest, "Test1"
      .AddModuleNode "Test2", gModPlugTest, "Test2"
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
Public Property Get Description() As String
   Description = "Test module to test application events."
End Property

Public Property Get Interface() As String
   Interface = "PlugTest.Main"
End Property

Public Property Get Display() As String
   Display = "PlugTest"
End Property

Public Property Get Name() As String
   Name = "PlugTest"
End Property

Public Property Get Path() As String
   Path = "<Defualt>"
End Property

Public Property Get Startup() As Boolean
   Startup = True
End Property

Public Property Get Supports() As String
   Supports = "1.0.45"
End Property

Public Property Get IPluginRegister_Description() As String
   IPluginRegister_Description = Description
End Property

Public Property Get IPluginRegister_Display() As String
   IPluginRegister_Display = Display
End Property

Public Property Get IPluginRegister_Interface() As String
   IPluginRegister_Interface = Interface
End Property

Public Property Get IPluginRegister_Name() As String
   IPluginRegister_Name = Name
End Property

Public Property Get IPluginRegister_Path() As String
   IPluginRegister_Path = Path
End Property

Public Property Get IPluginRegister_Startup() As Boolean
   IPluginRegister_Startup = Startup
End Property

Public Property Get IPluginRegister_Supports() As String
   IPluginRegister_Supports = Supports
End Property



'
' Plugin Utility Events
'
'
Private Sub UtilityEvents_SelectSetting(Setting As String)
   Select Case Setting
      Case "Test1"
         Application.Settings.SetActiveSetting frmConfig
         
      Case "Test2"
         Message Setting
      
         
   End Select
End Sub
   
Private Sub UtilityEvents_ViewSettings()
   Message "View Settings"
End Sub
