VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Plugins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"Plugin"
Attribute VB_Ext_KEY = "Member0" ,"Plugin"
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
Private Registry As Registry
Private mCol As Collection
Private mApplicationName As String

Private mbAllPlugsLoaded As Boolean
Private moAppEvents As Object

Private msAppVersion As String

' moAppObject is declared as object because we want the plugins collection to not
'  have any dependancies on the host application.  When the application creates
'  the moAppObject, it will set it to a pointer of its application object.  In turn,
'  when the plugins are loaded, they will recieve a pointer to this object.  However,
'  the plugins know the Application interface, so it's early bound for the plugins.
'
Private moAppObject As Object


Public Function ShowModuleManager() As Long
   '
   ' Show the plug-in registry editor
   '
   If (gAppVisible = True) Then
      gModuleManager.Show vbModal
   End If
   
End Function

Public Function RegisterPlugin( _
      inStrPluginPath As String _
   ) As Long
      
   Dim lsFileExt As String    '
   Dim lsResult As String     ' Result of path verification
   Dim lnResult As Long       ' Result of registrar
   
   ' If an empty string was passed, we exit indicating to try again with a valid
   '  parameter.
   '
   If (inStrPluginPath = "") Then
      RegisterPlugin = PLUG_ERROR_PARAMETER_REQUIRED
      Exit Function
   End If
   
   ' The name of the plugin is not valid
   '
   If (Len(inStrPluginPath) <= 4) Then
      RegisterPlugin = PLUG_ERROR_INCORRECT_PARAMETER_CRITERIA
      Exit Function
   End If
   
   ' Get the extension of the plugin in question
   '
   lsFileExt = LCase(Right(inStrPluginPath, 4))
   
   ' Registering a plugin can be tricky: we might have a dll, exe, or plc file selected.
   '  We must at all times verify that the file specified exists in the default or
   '  the specified location.  If we have a Plugin Configuration File (.plc) then we
   '  must also make sure that a corresponding exe or dll exists in the same location
   '  as the plc file.  If both an exe and a dll exist in the same location, we
   '  provide an error indicating that we don't know which to register.  If an exe or
   '  a dll neither exist, we provide an error code indicating so.
   '
   ' We first check the file type we are working with, then we attempt to verify all
   '  necessary files exist were we say they do.
   '
   lsResult = VerifyPluginPath(inStrPluginPath)
   
   If (lsResult <> "") Then
      Select Case lsFileExt
         Case ".plc"
            lnResult = RegisterPLC(lsResult)
                        
         Case ".dll"
            lnResult = RegisterDLL(lsResult)
            
         Case ".exe"
            lnResult = RegisterEXE(lsResult)
            
         Case Else
            RegisterPlugin = PLUG_ERROR_INCORRECT_PARAMETER_CRITERIA
            Exit Function
            
      End Select
      
      ' Evaluate the result code from the registrar
      '
      Select Case lnResult
         Case PLUG_ERROR_SUCCESS
            RegisterPlugin = lnResult
            
         Case Else
            RegisterPlugin = PLUG_ERROR_FAILED
            
      End Select
   End If   ' (lsResult <> "")

End Function

Public Function UnRegisterPlugin( _
      inStrPlugin As String _
   ) As Long
            
   ' Unregister the specified plug-in
   '
   Dim loPlugin As IPlugin
   Dim lnResult As Long
   Dim lsTemp As String
   Dim lsPath As String
   
   If (inStrPlugin = "") Then
      UnRegisterPlugin = PLUG_ERROR_PARAMETER_REQUIRED
      Exit Function
   End If
   
   For Each loPlugin In Me
      If (loPlugin.Display = inStrPlugin) Then
         lsTemp = (gsAppKey & loPlugin.Name)
         
         With Registry
            .ClassKey = HKEY_LOCAL_MACHINE
            .SectionKey = "Software\" & lsTemp & "\"
            .ValueType = REG_SZ
            .ValueKey = "Path"
            
            lsPath = .Value
            
            UnLoadPlugin (loPlugin.Name)
            lnResult = Shell("regsvr32 """ & lsPath & """ -u /s")
            
            .ValueKey = ""
            .DeleteKey
            
         End With
      End If
   Next
   
   UnRegisterPlugin = PLUG_ERROR_SUCCESS
End Function

Public Function LoadPlugins() As Long
   '
   ' Load all the registered plugins.  Should only be called before the program begins.
   '  We get our list of plugins from the \Modules\ Key of the registry.  We do not
   '  actively scan the hard disk for new plugs in certain directories.  Once we have
   '  a match, we add the plugin to the collection.
   '
   '
   ' Keys         Description
   ' ----------   -----------
   ' Name       : Name of plugin
   ' Startup    : Load at Startup, True/False
   ' Path       : Location of library
   ' Interface  : ProjectName.MainClass
   ' Description: Description of Plugin
   ' Display    : Display name of plugin
   '
   '
   Dim lsSections() As String
   Dim lnKeyCount As Long
   Dim lnKeyPosition As Long
      
   With Registry
      '
      ' Make sure we're working with the proper registry key
      '
      .ClassKey = HKEY_LOCAL_MACHINE
      .SectionKey = "Software\" & gsAppKey
      
      ' Get the module keys
      '
      .EnumerateSections lsSections(), lnKeyCount
      
      ' If there were no matches, handle it accordingly
      '
      If (lnKeyCount = 0) Then
         LoadPlugins = PLUG_ERROR_NO_VALUES
         mbAllPlugsLoaded = True
         Exit Function
      End If
      
      ' Loop through the Modules key and load each plugin accordingly
      '
      For lnKeyPosition = 1 To lnKeyCount
         LoadPlugin lsSections(lnKeyPosition)
      Next  ' lnKeyPosition
   End With ' Registry
      
   ' Set the status
   '
   LoadPlugins = PLUG_ERROR_SUCCESS
   
End Function

Public Function UnLoadPlugins() As Long
   '
   ' Unload all the loaded plugins.  Should only be called when the program terminates.
   '
   Dim loPlugin As Plugin

   ' Loop through the Plugins collection and unload each loaded module
   '
   For Each loPlugin In Me
      UnLoadPlugin loPlugin.Key
   Next

   UnLoadPlugins = PLUG_ERROR_SUCCESS
End Function

Public Function LoadPlugin(inStrPluginName As String) As Long
   '
   ' LoadPlugin is setup so it can be called dynamically at run-time, or during the
   '  initialization phase.  All registered plugins are initially loaded.  Before a
   '  plugin can be loaded during run-time, RegisterPlugin must be called and then
   '  LoadPlugin will succeed, because the module needs registry entries before it
   '  can be used.  If the plugin is not wanting to load during initialization, then
   '  UnloadPlugin must be called, and then UnRegisterPlugin and it will not load
   '  next time the program initializes.
   '
   Dim loPlugin As IPlugin
   Dim loTmpPlugin As Plugin
   
   Dim lsValues() As String
   Dim lnValueCount As Long
   Dim lnValuePosition As Long
   
   '
   ' Load the specified plugin.  May be called through the object model, usually called
   '  from the Module Manager to activate a plugin.
   '
   With Registry
      '
      ' Make sure we're working with the proper registry key
      '
      .ClassKey = HKEY_LOCAL_MACHINE
      .SectionKey = "Software\" & gsAppKey & inStrPluginName & "\"
      
      ' Get the values
      '
      .EnumerateValues lsValues, lnValueCount
      
      ' If we didn't find anything, we return an error
      '
      If (lnValueCount = 0) Then
         LoadPlugin = PLUG_ERROR_REGISTRY_NOT_FOUND
         Exit Function
      End If
      
      ' If the plugin already exists, don't add it again
      '
      For Each loTmpPlugin In Me
         If (loTmpPlugin.Key = inStrPluginName) Then
            LoadPlugin = PLUG_ERROR_MODULE_ALREADY_EXISTS
            Exit Function
         End If
      Next
      
      ' Allocate a new plugin object in the collection
      '
      Set loPlugin = Add(inStrPluginName)
   
      ' Populate the plugin object with appropriate values
      '
      For lnValuePosition = 1 To lnValueCount
         '
         ' Set the key we are working with
         '
         .ValueKey = lsValues(lnValuePosition)
         
         Select Case .ValueKey
            Case "Name"
               loPlugin.Name = .Value
               
            Case "Description"
               loPlugin.Description = .Value
               
            Case "Startup"
               loPlugin.Startup = .Value
               
            Case "Interface"
               loPlugin.Interface = .Value
               
            Case "Path"
               loPlugin.Path = .Value
            
            Case "Display"
               loPlugin.Display = .Value
            
            Case "Supports"
               loPlugin.Supports = .Value
               
         End Select  ' .ValueKey := lsValues(lnValuePosition)
         
      Next  ' lnValuePosition
      
      ' Now we load the plugin.  If StatupMode = True, then we also activate it, else,
      '  it's only loaded.
      '
      loPlugin.Load
      
      If (loPlugin.Startup = True) Then
         loPlugin.Activate
      End If
      
   End With ' Registry
   
   If (Not loPlugin Is Nothing) Then
      Set loPlugin = Nothing
   End If
   
   ' Signal to the application that we are complete loading all plugins
   '
   mbAllPlugsLoaded = True
   
   LoadPlugin = PLUG_ERROR_SUCCESS
   
End Function

Public Function UnLoadPlugin(inStrPluginName As String) As Long
   '
   ' Unload a loaded plugin.  May be called through the object model, usually called
   '  from the Module Manager to deactivate a plugin.
   '
   Dim lnModuleCount As Long
   
   ' By default, we assume the plugin is not loaded.
   '
   UnLoadPlugin = PLUG_ERROR_NOT_LOADED
   
   ' Find the plugin
   '
   If (Count > 0) Then
      For lnModuleCount = 1 To Count
         If (Item(lnModuleCount).Key = inStrPluginName) Then
            Item(lnModuleCount).IPlugin_UnLoad
            Remove lnModuleCount
            
            UnLoadPlugin = PLUG_ERROR_SUCCESS
            Exit For
         End If
      Next
   End If
   
End Function


Public Function VerifyPluginPath(isPlugin As String) As String
   '
   ' Either the plugin is located in the default Modules directory, or in a
   '  specified location, or no where at all.  Return the full path of the
   '  file in question.
   '
   If (InStr(1, isPlugin, "\") = 0) Then
      '
      ' \Modules\ directory (Default)
      '
      If (Dir(gsAppPath & "\Modules\" & isPlugin) = "") Then
         GoTo Xit:
      Else
         VerifyPluginPath = gsAppPath & "\Modules\" & isPlugin
         GoTo Xit:
      End If
   End If
   
   If (Dir(isPlugin) = "") Then
      '
      ' Specific location
      '
      GoTo Xit:
   Else
      VerifyPluginPath = isPlugin
      GoTo Xit:
   End If

Xit:
   
   Exit Function
   
End Function


Public Function GetNameFromPath(isFilePath As String) As String
   '
   ' Return the name of a file based on the path or actual name with extension provided
   '
   Dim lsTemp As String
   Dim lnCount As Long
   Dim lnFileStart As Long
   Dim lsTempFile As String
   
   ' Figure out if there is any subfolder, or not
   
   ' Reverse the string, because if we do, than the "\" preceding the file name will
   '  be the first in the string
   '
   For lnCount = Len(isFilePath) To 1 Step -1
      lsTemp = (lsTemp & Mid(isFilePath, lnCount, 1))
   Next
   
   ' Get the position of the backslash
   '
   lnFileStart = InStr(1, lsTemp, "\")
   
   ' Get the filename (at this point, it's reversed)
   '
   lsTempFile = Left(lsTemp, lnFileStart - 1)
   
   ' Reverse it back
   '
   lsTemp = ""
   
   For lnCount = Len(lsTempFile) To 1 Step -1
      lsTemp = (lsTemp & Mid(lsTempFile, lnCount, 1))
   Next
   
   ' Strip the extension from the filename
   '
   lsTemp = Left(lsTemp, (Len(lsTemp) - 4))
   
   ' Return
   '
   GetNameFromPath = lsTemp
   
End Function

Private Function RegisterPLC(isPluginPath As String) As Long
   '
End Function

Private Function RegisterDLL(isPluginPath As String) As Long
On Error GoTo HandleError:

   '
   ' Our plugin is a dll.  We will register it.  We must first check the registry
   '  and see if there is an entry in the [Path] key that matches [isPluginPath].  If
   '  so, then it was partially registered from a PLC file (most likely) and we
   '  will have to retrieve the Interface and then register the plugin.
   '
   Dim lnResult As Long
   Dim lsProjName As String
   Dim loPlugReg As IPluginRegister

   
   ' The very first thing, is to register with Windows the dll in question.
   '
   lnResult = Shell("regsvr32 """ & isPluginPath & """ /s")
   
   ' Create the object, based on the interface, to complete the registry
   '
   lsProjName = GetNameFromPath(isPluginPath)
   lsProjName = (lsProjName & ".Main")
   
   Set loPlugReg = CreateObject(lsProjName)
   
   ' Register the plugin
   '
   With Registry
      .ClassKey = HKEY_LOCAL_MACHINE
      .SectionKey = "Software\" & gsAppKey & loPlugReg.Name & "\"
      .CreateKey
            
      .ValueType = REG_EXPAND_SZ
      
      .ValueKey = "Name"
      .Value = loPlugReg.Name
      
      .ValueKey = "Description"
      .Value = loPlugReg.Description
      
      .ValueKey = "Path"
      .Value = isPluginPath
      
      .ValueKey = "Interface"
      .Value = loPlugReg.Interface
      
      .ValueKey = "Display"
      .Value = loPlugReg.Display
      
      .ValueKey = "Startup"
      .Value = loPlugReg.Startup
      
      .ValueKey = "Supports"
      .Value = loPlugReg.Supports
      
   End With
   
   RegisterDLL = PLUG_ERROR_SUCCESS
   
Xit:
   
   ' Do all of our cleanup work here
   '
   If (Not loPlugReg Is Nothing) Then
      Set loPlugReg = Nothing
   End If
   
   Exit Function
   
   
HandleError:

   Select Case Err.Number
      Case 429
         RegisterDLL = PLUG_ERROR_LIBRARY_NOT_FOUND
      
      Case Else
         RegisterDLL = PLUG_ERROR_FAILED
         
   End Select
   
   GoTo Xit:

End Function

Private Function RegisterEXE(isPluginPath As String) As Long
   '
End Function


Public Property Let ApplicationName(Value As String)
   '
   ' The plugin manager needs to know what the name of the current application is so
   '  it can use the program registry key to contain the modules information.  In this
   '  case, if the name has already been set, we don't set it again.  This prevents it
   '  from being changed during execution.
   '
   If (mApplicationName = "") Then
      mApplicationName = Value
      gsAppKey = Value & "\Modules\"
   End If
   
End Property

Public Property Get ApplicationName() As String
   ApplicationName = mApplicationName
End Property

Public Property Let ApplicationPath(Value As String)
   '
   ' The plugin manager needs to know the application path so it can locate the modules
   '  folder.  Once the name has been set, it cannot be changed.  This prevents it from
   '  being changed during execution.
   '
   If (gsAppPath = "") Then
      gsAppPath = Value
   End If
   
End Property

Public Property Get ApplicationPath() As String
   ApplicationPath = gsAppPath
End Property

Public Property Let ApplicationPointer(Value As Object)
   If (moAppObject Is Nothing) Then
      Set moAppObject = Value
   End If
End Property

Public Property Get ApplicationPointer() As Object
   Set ApplicationPointer = moAppObject
End Property

Public Property Let Visible(Value As Boolean)
   gAppVisible = Value
End Property

Public Property Get Visible() As Boolean
   Visible = gAppVisible
End Property

Public Property Let AppEventPointer(Value As Object)
   Set moAppEvents = Value
End Property

Public Property Get AppEventPointer() As Object
   Set AppEventPointer = moAppEvents
End Property

Public Property Get AllPluginsLoaded() As Boolean
   AllPluginsLoaded = mbAllPlugsLoaded
End Property

Public Property Get UtilityEvents() As UtilityEvents
   Set UtilityEvents = gUtilityEvents
End Property

Public Property Let AppVersion(Value As String)
   msAppVersion = Value
End Property

Public Property Get AppVersion() As String
   AppVersion = msAppVersion
End Property

Public Property Get Settings() As Settings
   Set Settings = gSettings
End Property





'
'
' ClassBuilder Creations.
'
'
'
Public Function Add(Key As String, Optional sKey As String) As Plugin
   Dim objNewMember As Plugin
   Set objNewMember = New Plugin
   
   objNewMember.Key = Key
   If Len(sKey) = 0 Then
      mCol.Add objNewMember
   Else
      mCol.Add objNewMember, sKey
   End If
   
   Set Add = objNewMember
   Set objNewMember = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As Plugin
Attribute Item.VB_UserMemId = 0
   Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
   Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
   mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
   Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
   Set gModuleManager = New frmModuleManager
   Set gUtilityEvents = New UtilityEvents
   Set gSettings = New Settings
   Set Registry = New Registry
   Set mCol = New Collection
   Set gPlugins = Me
End Sub

Private Sub Class_Terminate()
   Set gUtilityEvents = Nothing
   Set gModuleManager = Nothing
   Set gSettings = Nothing
   Set Registry = Nothing
   Set gPlugins = Nothing
   Set mCol = Nothing
   
   If (Not moAppEvents Is Nothing) Then
      Set moAppEvents = Nothing
   End If
   
   If (Not moAppObject Is Nothing) Then
      Set moAppObject = Nothing
   End If
End Sub

