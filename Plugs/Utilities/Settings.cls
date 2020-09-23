VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
Private m_frmSettings As frmSettings
Private mnGenKey As Long

Private msModulesKey As String
Private msApplicationKey As String
Private msConflictKey As String



Public Function AddGroup( _
      isGroupName As String, _
      Optional isGroupKey As String _
   ) As String
   
   Dim loNode As Node
   Dim lsKey As String
   
   If (IsMissing(isGroupKey)) Then
      lsKey = GenerateKey()
   Else
      lsKey = isGroupKey
   End If
   
   Set loNode = m_frmSettings.tvwModules.Nodes.Add(, , lsKey, isGroupName)
       loNode.Tag = "About:" & isGroupName
       loNode.Expanded = True
   
   Set loNode = Nothing
   AddGroup = lsKey
End Function

Public Function AddNode( _
      isGroup As String, _
      isNodeName As String, _
      Optional isIdentifier As String _
   ) As String
   
   Dim loNode As Node
   Dim lsKey As String
   
   lsKey = GenerateKey
   
   Set loNode = m_frmSettings.tvwModules.Nodes.Add(isGroup, tvwChild, lsKey, isNodeName)
       loNode.Expanded = True
   
   If (Not IsMissing(isIdentifier)) Then
      loNode.Tag = isIdentifier
   End If
   
   Set loNode = Nothing
   AddNode = lsKey
End Function

Public Function AddModuleGroup( _
      isGroupName As String _
   ) As String
   
   Dim loNode As Node
   Dim lsKey As String
   
   lsKey = GenerateKey
   
   Set loNode = m_frmSettings.tvwModules.Nodes.Add(msModulesKey, tvwChild, lsKey, isGroupName)
       loNode.Tag = "About:" & isGroupName
       loNode.Expanded = True
   
   Set loNode = Nothing
   AddModuleGroup = lsKey
End Function

Public Function AddModuleNode( _
      isModule As String, _
      Optional isGroup As String = "Modules", _
      Optional isIdentifier As String _
   ) As String
   
   Dim lsKey As String
   
   lsKey = AddNode(isGroup, isModule, isIdentifier)
   AddModuleNode = lsKey
End Function

Public Function RemoveGroup(isGroupName) As String
   '
End Function

Public Function RemoveKey(isKeyName) As String
   '
End Function

Public Function RemoveModuleGroup(isGroupKey As String) As Long
On Error GoTo HandleError:

   Dim loNode As Node
   
   For Each loNode In m_frmSettings.tvwModules.Nodes
      If (Not loNode.Parent Is Nothing) Then
         If (loNode.Parent.Key = isGroupKey) Then
            m_frmSettings.tvwModules.Nodes.Remove loNode.Index
         End If
      End If
   Next
   
   For Each loNode In m_frmSettings.tvwModules.Nodes
      If (loNode.Key = isGroupKey) Then
         m_frmSettings.tvwModules.Nodes.Remove loNode.Index
      End If
   Next

   RemoveModuleGroup = PLUG_ERROR_SUCCESS
      
Xit:
   
   Exit Function
   
HandleError:
   If (Err.Number = 35606) Then
      Resume Next
   End If
   
End Function

Private Function GenerateKey() As String
   mnGenKey = (mnGenKey + 1)
   GenerateKey = ("Key" & CStr(mnGenKey))
End Function

Public Sub ShowSettings()
   If (gAppVisible = True) Then
      gUtilityEvents.Send_ViewSettings
      m_frmSettings.Show
   End If
End Sub

Public Sub HideSettings()
   m_frmSettings.Hide
End Sub

Public Function SetActiveSetting(SettingScreen As Object) As Long
   SetActiveSetting = m_frmSettings.SetActiveSetting(SettingScreen)
End Function

Public Property Get ApplicationKey() As String
   ApplicationKey = msApplicationKey
End Property

Public Property Get ModulesKey() As String
   ModulesKey = msModulesKey
End Property

Public Property Get ConflictKey() As String
   ConflictKey = msConflictKey
End Property


Private Sub Class_Initialize()
   Set m_frmSettings = New frmSettings
   
   msApplicationKey = "Application"
   msConflictKey = "Conflicts"
   msModulesKey = "Modules"
   
   ' Add the application key to the treeview
   '
   AddGroup "Application", ApplicationKey
   
   ' Add the modules group to the tree veiw of the settings dialog
   '
   AddGroup "Modules", ModulesKey
   
End Sub

Private Sub Class_Terminate()
   Set m_frmSettings = Nothing
End Sub
