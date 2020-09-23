VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IModuleManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
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
Public Function AddModule( _
      isModule As String, _
      isDescription As String, _
      ibStartup As Boolean, _
      inColor As PLUGIN_COLOR_CODES _
   ) As Long
   '
End Function

Public Function UpdateModuleStartupMode( _
      isModule As String, _
      ibStartup As Boolean _
   ) As Long
   '
End Function

Public Function UpdateModuleColor( _
      isModule As String, _
      inColor As PLUGIN_COLOR_CODES _
   ) As Long
   '
End Function

Public Function RemoveModule( _
      isModule As String _
   ) As Long
   '
End Function

Public Function UpdateModuleState( _
      isModule As String, _
      ibLoaded As Boolean _
   ) As Long
   '
End Function

Public Function UpdateModuleStatus( _
      isModule As String, _
      inStatus As PLUGIN_STATUS_MODE _
   ) As Long
   '
End Function

Public Property Get ModuleStatus( _
      isModule As String _
   ) As Long
   '
End Property
