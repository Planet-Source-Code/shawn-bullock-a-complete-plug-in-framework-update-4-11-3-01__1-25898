VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Plugin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
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
Implements IPluginVersion
Implements IPlugin


Public Key As String


' Version
'
Private mCompanyName As String
Private mCopyright As String
Private mMajor As Long
Private mMinor As Long
Private mRevision As Long
Private mTrademarks As String
Private mVersion As String

' Actuals
'
Private mPlugDescription As String
Private mPlugInterface As String
Private mPlugName As String
Private mPlugPath As String
Private mPlugStartup As Boolean
Private mPlugIsActive As Boolean
Private mPlugDisplay As String
Private mPlugSupports As String
Private mPlugStatus As Long

' Defaults
'
Private mDefPlugDescription As String
Private mDefPlugInterface As String
Private mDefPlugName As String
Private mDefPlugPath As String
Private mDefPlugStartup As Boolean
Private mDefPlugDisplay As String
Private mDefPlugSupports As String
Private mDefPlugStartupCount As Long

' Other
'
Private moPlugin As IPluginHook
Private moModuleManager As IModuleManager
Private Registry As Registry
Private mUtilityEvents As UtilityEvents




'
' IPlugin Interface
'
'
' All let statements: don't set if already set
'
'
Private Property Let IPlugin_Description(Value As String)
   If (mPlugDescription = "") Then
      mDefPlugDescription = Value
      mPlugDescription = Value
   End If
End Property

Public Property Get IPlugin_Description() As String
   IPlugin_Description = mPlugDescription
End Property


Public Property Let IPlugin_Display(Value As String)
   If (mPlugDisplay = "") Then
      mDefPlugDisplay = Value
      mPlugDisplay = Value
   End If
End Property

Public Property Get IPlugin_Display() As String
   IPlugin_Display = mPlugDisplay
End Property


Private Property Let IPlugin_Interface(Value As String)
   If (mPlugInterface = "") Then
      mDefPlugInterface = Value
      mPlugInterface = Value
   End If
End Property

Public Property Get IPlugin_Interface() As String
   IPlugin_Interface = mPlugInterface
End Property


Private Property Let IPlugin_Name(Value As String)
   If (mPlugName = "") Then
      mDefPlugName = Value
      mPlugName = Value
   End If
End Property

Public Property Get IPlugin_Name() As String
   IPlugin_Name = mPlugName
End Property


Private Property Let IPlugin_Path(Value As String)
   If (mPlugPath = "") Then
      mDefPlugPath = Value
      mPlugPath = Value
   End If
End Property

Public Property Get IPlugin_Path() As String
   IPlugin_Path = mPlugPath
End Property


Private Property Let IPlugin_Startup(Value As Boolean)
   If (mDefPlugStartupCount = 0) Then
      mDefPlugStartupCount = 1
      mDefPlugStartup = Value
   End If
   
   IPlugin_UpdateDefault "Startup", CStr(Value)
   mPlugStartup = Value
End Property

Public Property Get IPlugin_Startup() As Boolean
   IPlugin_Startup = mPlugStartup
End Property


Public Property Let IPlugin_Supports(Value As String)
   If (mPlugSupports = "") Then
      mDefPlugSupports = Value
      mPlugSupports = Value
   End If
End Property

Public Property Get IPlugin_Supports() As String
   IPlugin_Supports = mPlugSupports
End Property


Public Property Let IPlugin_IsActive(Value As Boolean)
   mPlugIsActive = Value
End Property

Public Property Get IPlugin_IsActive() As Boolean
   IPlugin_IsActive = mPlugIsActive
End Property


Public Property Let IPlugin_Status(Value As Long)
   mPlugStatus = Value
End Property

Public Property Get IPlugin_Status() As Long
   IPlugin_Status = mPlugStatus
End Property


Public Function IPlugin_Load() As Long

   ' Before we can load the plugin and keep a reference, we must make sure we have the
   '  interface key.  If not, we return an error indicated that there is not a valid
   '  interface.
   '
   If (mPlugInterface = "") Then
      '
      ' No valid interface.  It's possible this plugin was added to the collection but not
      '  retrieved from the registry.  It shouldn't happen that way if you use the proper
      '  methods in Plugins.
      '
      IPlugin_Load = PLUG_ERROR_NO_VALID_INTERFACE
      Exit Function
   End If
      
   ' Add the plugin listing to the module manager form
   '
   Set moModuleManager = gModuleManager
       moModuleManager.AddModule mPlugDisplay, mPlugDescription, mPlugStartup, PLUG_COLOR_NORMAL
       moModuleManager.UpdateModuleStatus mPlugDisplay, PLUG_STATUS_INACTIVE

End Function

Public Function IPlugin_UnLoad() As Long
   '
   ' Unload the plug in.
   '
   If (moPlugin Is Nothing) Then
      '
      ' The plugin isn't loaded
      '
      IPlugin_UnLoad = PLUG_ERROR_NOT_LOADED
      Exit Function
   End If
   
   ' Deactivate the plugin first
   '
   IPlugin_Deactivate
   
   ' Remove it from the Module Manager
   '
   moModuleManager.UpdateModuleStatus mPlugDisplay, PLUG_STATUS_UNLOADING
   IPlugin_UnLoad = moModuleManager.RemoveModule(mPlugDisplay)
   
End Function


Public Function IPlugin_Activate() As Long
On Error GoTo HandleError:

   '
   ' After the plugin is registered, it must be loaded.  By default, if the startup
   '  attribute is True, then when Load() is called it is also activated.  If it is
   '  false, then we have to manually activate the plug in.  A plug-in may be in the
   '  Module Manager, but it is not hooked into the interface until Activate is
   '  called.
   '
   
   ' Make sure we are not already active, if so, we return Success
   '
   If (mPlugIsActive = True) Then
      IPlugin_Activate = PLUG_ERROR_SUCCESS
      Exit Function
   End If
   
   ' Load the interface
   '
   Set moPlugin = CreateObject(mPlugInterface)

   With moPlugin
      '
      ' Set the pointer to the application
      '
      .Application = gPlugins.ApplicationPointer
      .AppEvents = gPlugins.AppEventPointer
      .UtilityEvents = gUtilityEvents
   End With
   
   ' The plugin must support a minimum version of the host application.  If the host is
   '  too old, then we must indicate that it's not supported.
   '
   If (gPlugins.AppVersion < mPlugSupports) Then
      moModuleManager.UpdateModuleColor mPlugDisplay, PLUG_COLOR_RED
      moModuleManager.UpdateModuleStatus mPlugDisplay, PLUG_STATUS_OLD_HOST
      
      IPlugin_Activate = PLUG_ERROR_BAD_VERSION
      Exit Function
   End If
   
   ' Hook the plugin into the interface
   '
   moPlugin.HookPlugin
   
   ' Indicate in the Module Manager that the plugin is loaded.
   '
   moModuleManager.UpdateModuleStatus mPlugDisplay, PLUG_STATUS_ACTIVE
   moModuleManager.UpdateModuleState mPlugDisplay, True
   mPlugIsActive = True
   
   IPlugin_Activate = PLUG_ERROR_SUCCESS

   
Xit:
   
   Exit Function
   
   
HandleError:
   
   Select Case Err.Number
      Case 429
         '
         ' Can't create the ActiveX object.
         '
         IPlugin_Activate = PLUG_ERROR_LIBRARY_NOT_FOUND
         moModuleManager.UpdateModuleColor mPlugDisplay, PLUG_COLOR_RED
         moModuleManager.UpdateModuleStatus mPlugDisplay, PLUG_STATUS_ERROR_ACTIVATING
      
      Case 13
         '
         ' Version not the same, GUID Different
         '
         IPlugin_Activate = PLUG_ERROR_BAD_VERSION
         moModuleManager.UpdateModuleColor mPlugDisplay, PLUG_COLOR_RED
         moModuleManager.UpdateModuleStatus mPlugDisplay, PLUG_STATUS_BAD_VERSION
      
      Case Else
         '
         ' Any other error not yet accounted for
         '
         IPlugin_Activate = PLUG_ERROR_FAILED
         moModuleManager.UpdateModuleColor mPlugDisplay, PLUG_COLOR_RED
         moModuleManager.UpdateModuleStatus mPlugDisplay, PLUG_STATUS_UNKNOWN
   
   End Select
   
End Function

Public Function IPlugin_Deactivate() As Long
On Error GoTo HandleError:

   '
   ' Deactivate the plugin this will automatically be called by UnLoad.  However, in
   '  the case of runtime deactivation, may also be called individually.  This is also
   '  called when you uncheck an item in the Module Manager.
   '
   
   ' If we are already inactive, return success
   '
   If (mPlugIsActive = False) Then
      IPlugin_Deactivate = PLUG_ERROR_SUCCESS
      Exit Function
   End If
   
   ' Unhook our plugin from the interface
   '
   moPlugin.UnHookPlugin
   
   ' Relieve our active reference to the plugin
   '
   Set moPlugin = Nothing
   mPlugIsActive = False
   
   ' Indicate that is it not active in the Module Manager
   '
   moModuleManager.UpdateModuleState mPlugDisplay, False
   moModuleManager.UpdateModuleStatus mPlugDisplay, PLUG_STATUS_INACTIVE
   
Xit:
   
   Exit Function
   
HandleError:
   moModuleManager.UpdateModuleColor mPlugDisplay, PLUG_COLOR_RED
   moModuleManager.UpdateModuleStatus mPlugDisplay, PLUG_STATUS_ERROR_DEACTIVATING
   
   IPlugin_Deactivate = PLUG_ERROR_UNABLE_TO_DEACTIVATE
   
End Function

Public Function IPlugin_UpdateDefault(isNode As String, Value As String) As Long
   
   With Registry
      .ClassKey = HKEY_LOCAL_MACHINE
      .SectionKey = "Software\" & gsAppKey & mDefPlugName & "\"
      
      Select Case isNode
         Case "Startup"
            .ValueKey = isNode
            .ValueType = REG_SZ
            .Value = Value
            
         Case Else
            '
            
      End Select
   End With
   
   IPlugin_UpdateDefault = PLUG_ERROR_SUCCESS
End Function

Public Function IPlugin_GetDefault(isNode As String) As Variant
   Dim lsTemp As Variant
   
   Select Case isNode
      Case "Description"
         lsTemp = mDefPlugDescription
         
      Case "Display"
         lsTemp = mDefPlugDisplay
         
      Case "Interface"
         lsTemp = mDefPlugInterface
         
      Case "Name"
         lsTemp = mDefPlugName
         
      Case "Path"
         lsTemp = mDefPlugPath
         
      Case "Startup"
         lsTemp = mDefPlugStartup
         
      Case "Supports"
         lsTemp = mDefPlugSupports
         
      Case Else
         lsTemp = ""
         
   End Select
   
   IPlugin_GetDefault = lsTemp
End Function




'
' IPluginVersion Interface
'
'
' All Let statements don't set if already set
'
'
Public Property Let IPluginVersion_CompanyName(Value As String)
   If (mCompanyName = "") Then
      mCompanyName = Value
   End If
End Property

Public Property Get IPluginVersion_CompanyName() As String
   IPluginVersion_CompanyName = mCompanyName
End Property


Public Property Let IPluginVersion_Copyright(Value As String)
   If (mCopyright = "") Then
      mCopyright = Value
   End If
End Property

Public Property Get IPluginVersion_Copyright() As String
   IPluginVersion_Copyright = mCopyright
End Property


Public Property Let IPluginVersion_Major(Value As Long)
   If (mMajor < 0) Then
      mMajor = Value
   End If
End Property

Public Property Get IPluginVersion_Major() As Long
   IPluginVersion_Major = mMajor
End Property


Public Property Let IPluginVersion_Minor(Value As Long)
   If (mMinor < 0) Then
      mMinor = Value
   End If
End Property

Public Property Get IPluginVersion_Minor() As Long
   IPluginVersion_Minor = mMinor
End Property


Public Property Let IPluginVersion_Revision(Value As Long)
   If (mRevision < 0) Then
      mRevision = Value
   End If
End Property

Public Property Get IPluginVersion_Revision() As Long
   IPluginVersion_Revision = mRevision
End Property


Public Property Let IPluginVersion_Trademarks(Value As String)
   If (mTrademarks = "") Then
      mTrademarks = Value
   End If
End Property

Public Property Get IPluginVersion_Trademarks() As String
   IPluginVersion_Trademarks = mTrademarks
End Property


Public Property Get IPluginVersion_Version() As String
   IPluginVersion_Version = mMajor & "." & mMinor & "." & mRevision
End Property







Private Sub Class_Initialize()
   '
   ' Set the default versions.  Each of these values, when less than 0, indicates that
   '  the value has not been set.  Once it's set, it can not be changed.  By setting it,
   '  the number becomes 0 or positive, therefore we know it's not default when attempting
   '  to modify the values.
   '
   mMajor = -1
   mMinor = -1
   mRevision = -1
   
   Set Registry = New Registry
End Sub

Private Sub Class_Terminate()
   Set moModuleManager = Nothing
   Set moPlugin = Nothing
   Set Registry = Nothing
End Sub

