VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmModuleManager 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Module Manager"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   8475
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAction 
      Caption         =   "&Add"
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   6060
      TabIndex        =   6
      Top             =   3540
      Width           =   555
   End
   Begin VB.TextBox txtPlugin 
      Height          =   285
      Left            =   60
      TabIndex        =   5
      Top             =   3540
      Width           =   5955
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "..."
      Height          =   315
      Index           =   3
      Left            =   6660
      TabIndex        =   4
      Top             =   3540
      Width           =   315
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   7920
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   7200
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   7200
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvwPlugins 
      Height          =   3435
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   6059
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Plug-in Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Startup / Status"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CheckBox chkStartup 
      Caption         =   "&Startup"
      Height          =   255
      Left            =   7140
      TabIndex        =   0
      Top             =   3540
      Width           =   1215
   End
   Begin VB.Menu mnuModuleOptions 
      Caption         =   "&Module Options"
      Visible         =   0   'False
      Begin VB.Menu miUnRegister 
         Caption         =   "&UnRegister..."
      End
   End
End
Attribute VB_Name = "frmModuleManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Implements IModuleManager



Private Const CMD_OK = 0
Private Const CMD_CANCEL = 1
Private Const CMD_ADD = 2
Private Const CMD_BROWSE = 3

Private mbStartup As Boolean
Private mbActive As Boolean
Private mnStatus As Long


Private Sub chkStartup_Click()
   '
   Dim lsModule As String

   ' Get the item selected, this is who we'll deal with for now
   '
   lsModule = lvwPlugins.SelectedItem
   
   If (chkStartup.Value = vbChecked) Then
      IModuleManager_UpdateModuleStartupMode lsModule, True
   Else
      IModuleManager_UpdateModuleStartupMode lsModule, False
   End If
   
End Sub

Private Sub cmdAction_Click(Index As Integer)
   Select Case Index
      Case CMD_OK
         Hide
         
      Case CMD_CANCEL
         Hide
         
      Case CMD_ADD
         gPlugins.RegisterPlugin txtPlugin
         gPlugins.LoadPlugin gPlugins.GetNameFromPath(gPlugins.VerifyPluginPath(txtPlugin))
         
      Case CMD_BROWSE
         BrowseForPlugin
         
   End Select
End Sub

Public Sub BrowseForPlugin()
   '
   ' Browse for the plugin.  Default to the application default directory.  There are
   '  three types of files supported.
   '
   ' File            Description
   ' ----            -----------
   '
   ' DLL             Dynamic Link Library.  This is the preferred plugin type.
   ' EXE             ActiveX Executable.  Both a plugin and an application.
   ' PLC             Plugin Configuration File.  If a <filename>.plc file is located
   '                  in the same directory as <filename>.dll/exe, the Module
   '                  Manager will get the plugin settings from the plc file instead
   '                  if "<filename>.main" as an Interface.  This is in the case
   '                  that the project name of the plugin is different than the
   '                  filename.  Else, "<filename>.main" is the default.  The .plc
   '                  file is an INI file with the following structure:
   '
   '
   ' Group           Key                     Description
   ' -----           ---                     -----------
   '
   ' [Plugin]        Interface               Required
   '                 Name                    Required
   '                 Startup                 Optional
   '                 Path                    Automatic
   '                 Description             Optional
   '
   ' [Special]       ...                     ...
   '
   '
   ' If any of these keys are present in the config file, then their respective string
   '  is inserted into the key of the \appname\modules\<Name>\ node.  If any other
   '  special registry keys need to be set, then they must be placed in the [Special]
   '  group and they will be placed accordingly as \modules\<name>\Special\... and
   '  will default their type to string.
   '
   '
   With CD
      .InitDir = gPlugins.ApplicationPath & "\Modules\"
      .Filter = "Dynamic Link Libraries (*.dll)|*.dll|" & _
                "Executable Files (*.exe)|*.exe|" & _
                "Plugin Configuration Files (*.plc)|*.plc"

      
      .ShowOpen
      
      If (.FileName <> "") Then
         If (Left(.FileName, (Len(.FileName) - Len(.FileTitle))) = .InitDir) Then
            txtPlugin = .FileTitle
         Else
            txtPlugin = .FileName
         End If
      End If
   End With
End Sub

Private Sub Form_Load()
   With lvwPlugins
      .ColumnHeaders(1).Width = x(111)
      .ColumnHeaders(2).Width = x(227)
      .ColumnHeaders(3).Width = x(119)
   End With
End Sub



'
'
' IModuleManager Interface
'
'
'
Public Function IModuleManager_AddModule( _
      isModule As String, _
      isDescription As String, _
      ibStartup As Boolean, _
      inColor As PLUGIN_COLOR_CODES _
   ) As Long
   
   Dim loNode As ListItem
   Dim lnCount As Long
   
   ' If the module is already listed, we won't attempt to list it again.  This
   '  avoids a non-unique key error in the listview.  Return success.
   '
   If (lvwPlugins.ListItems.Count > 0) Then
      For lnCount = 1 To lvwPlugins.ListItems.Count
         If (lvwPlugins.ListItems(lnCount) = isModule) Then
            IModuleManager_AddModule = PLUG_ERROR_SUCCESS
            Exit Function
         End If
      Next
   End If
   
   ' Add the module to the listing
   '
   With lvwPlugins
      Set loNode = .ListItems.Add(, isModule, isModule)
          loNode.SubItems(1) = isDescription
          IModuleManager_UpdateModuleStartupMode isModule, ibStartup
          loNode.ForeColor = inColor
          
   End With
   
   IModuleManager_AddModule = PLUG_ERROR_SUCCESS
End Function

Public Function IModuleManager_RemoveModule( _
      isModule As String _
   ) As Long

   Dim lnCount As Long
   
   If (lvwPlugins.ListItems.Count > 0) Then
      For lnCount = 1 To lvwPlugins.ListItems.Count
         If (lvwPlugins.ListItems(lnCount) = isModule) Then
            lvwPlugins.ListItems.Remove lnCount
            Exit For
         End If
      Next
   Else
      IModuleManager_RemoveModule = PLUG_ERROR_NO_VALUES
      Exit Function
   End If
   
   IModuleManager_RemoveModule = PLUG_ERROR_SUCCESS
   
End Function

Public Function IModuleManager_UpdateModuleColor( _
      isModule As String, _
      inColor As PLUGIN_COLOR_CODES _
   ) As Long

   Dim loNode As ListItem
   
   ' Get a pointer to the module in the list
   '
   Set loNode = GetModuleItem(isModule)
   If (Not loNode Is Nothing) Then
      loNode.ForeColor = inColor
   End If
   
   IModuleManager_UpdateModuleColor = PLUG_ERROR_SUCCESS
End Function

Public Function IModuleManager_UpdateModuleStartupMode( _
      isModule As String, _
      ibStartup As Boolean _
   ) As Long

   Dim loNode As ListItem
   Dim loPlugin As IPlugin
   
   ' Get a pointer to the module in the list
   '
   Set loNode = GetModuleItem(isModule)
   If (Not loNode Is Nothing) Then
      mbStartup = ibStartup
      
      For Each loPlugin In gPlugins
         If (loPlugin.Display = loNode.Text) Then
            loPlugin.Startup = ibStartup
            mnStatus = loPlugin.Status
         End If
      Next
      
      IModuleManager_UpdateModuleStatus isModule, mnStatus
      
   End If
   
End Function

Public Function IModuleManager_UpdateModuleState( _
      isModule As String, _
      ibLoaded As Boolean _
   ) As Long
      
   Dim loNode As ListItem
   
   ' Get a pointer to the module reference in the list
   '
   Set loNode = GetModuleItem(isModule)
   If (Not loNode Is Nothing) Then
      loNode.Checked = ibLoaded
   End If
   
   ' Remember that state
   '
   mbActive = ibLoaded
   
   IModuleManager_UpdateModuleState = PLUG_ERROR_SUCCESS
   
End Function

Public Function IModuleManager_UpdateModuleStatus( _
      isModule As String, _
      inStatus As PLUGIN_STATUS_MODE _
   ) As Long
      
   Dim loNode As ListItem
   Dim lsTemp As String
   Dim loPlugin As IPlugin
   
   ' Get a pointer to the module reference in the list
   '
   Set loNode = GetModuleItem(isModule)
   If (Not loNode Is Nothing) Then
      Select Case inStatus
         Case PLUG_STATUS_ACTIVE
            lsTemp = "Active"
            
         Case PLUG_STATUS_ERROR_ACTIVATING
            lsTemp = "Error Activating"
            
         Case PLUG_STATUS_ERROR_DEACTIVATING
            lsTemp = "Error Deactivating"
            
         Case PLUG_STATUS_ERROR_LOADING
            lsTemp = "Error Loading"
            
         Case PLUG_STATUS_ERROR_UNLOADING
            lsTemp = "Error Unloading"
            
         Case PLUG_STATUS_INACTIVE
            lsTemp = "Inactive"
            
         Case PLUG_STATUS_LOADING
            lsTemp = "Loading"
            
         Case PLUG_STATUS_UNKNOWN
            lsTemp = "Unkown"
            
         Case PLUG_STATUS_UNLOADING
            lsTemp = "Unloading"
         
         Case PLUG_STATUS_BAD_VERSION
            lsTemp = "Bad Version"
         
         Case PLUG_STATUS_OLD_HOST
            lsTemp = "Old Host Version"
         
         Case PLUG_STATUS_MODULE_NOT_LOADED
            ' Not used for display
         
         Case PLUG_STATUS_FAILED_MISERABLY
            lsTemp = "Failed Miserably"
            
         Case Else
            '
            ' There shouldn't be anything else
            '
            lsTemp = "Invalid Parameter"
            
      End Select
      
      loNode.SubItems(2) = mbStartup & " / " & lsTemp
      IModuleManager_UpdateModuleStatus = PLUG_ERROR_SUCCESS
   Else
      '
      ' We couldn't find the module we are looking for
      '
      IModuleManager_UpdateModuleStatus = PLUG_ERROR_LISTING_NOT_FOUND
      
   End If
   
   ' Update the status
   '
   Set loPlugin = GetPlugin(isModule)
   If (Not loPlugin Is Nothing) Then
      loPlugin.Status = inStatus
   End If
   
End Function

Public Property Get FileName() As String
   FileName = CD.FileTitle
End Property

Public Property Get IModuleManager_ModuleStatus( _
      isModule As String _
   ) As Long
   
   Dim loPlugin As IPlugin
   
   ' Get a pointer to the module in question
   '
   Set loPlugin = GetPlugin(isModule)
   
   ' Return the status
   '
   If (Not loPlugin Is Nothing) Then
      IModuleManager_ModuleStatus = loPlugin.Status
   Else
      IModuleManager_ModuleStatus = PLUG_STATUS_MODULE_NOT_LOADED
   End If
   
   ' Cleanup after ourselves
   '
   Set loPlugin = Nothing
End Property


Private Function GetModuleItem( _
      isModule As String _
   ) As ListItem
      
   '
   ' Find the module requested and return a pointer to it.  If the function failes,
   '  a null will be returned.  We must check in the colling procedure to see if
   '  the return value != Nothing.
   '
   Dim lnCount As Long
   
   If (lvwPlugins.ListItems.Count > 0) Then
      For lnCount = 1 To lvwPlugins.ListItems.Count
         If (lvwPlugins.ListItems(lnCount) = isModule) Then
            Set GetModuleItem = lvwPlugins.ListItems(lnCount)
            Exit Function
         End If
      Next
   End If
   
   ' Object is not loaded
   '
   Set GetModuleItem = Nothing
   
End Function

Private Function GetPlugin( _
      isModule As String _
   ) As IPlugin
      
   Dim loPlugin As IPlugin
   
   For Each loPlugin In gPlugins
      If (loPlugin.Display = isModule) Then
         Set GetPlugin = loPlugin
      End If
   Next
End Function
      

Private Sub lvwPlugins_ItemCheck(ByVal Item As MSComctlLib.ListItem)
   '
   ' If we become checked, we activate the plugin.  If we become unchecked, we deactivate
   '  the plugin.  If we are red, we always keep the item unchecked.
   '
   Dim loPlugin As IPlugin
   
   If (Item.Checked = True) Then
      If (Item.ForeColor = PLUG_COLOR_RED) Then
         Item.Checked = False
      End If
      
      For Each loPlugin In gPlugins
         With loPlugin
            If (.Display = Item) Then
               .Activate
            End If
         End With
      Next
   Else
      For Each loPlugin In gPlugins
         With loPlugin
            If (.Display = Item) Then
               mbStartup = .Startup
               .Deactivate
            End If
         End With
      Next
   End If
End Sub

Private Sub lvwPlugins_ItemClick(ByVal Item As MSComctlLib.ListItem)
   Dim loPlugin As IPlugin
   
   For Each loPlugin In gPlugins
      If (loPlugin.Display = Item) Then
         If (loPlugin.Startup = True) Then
            chkStartup = vbChecked
         Else
            chkStartup = vbUnchecked
         End If
      End If
   Next
End Sub

Private Sub lvwPlugins_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   Select Case Button
      Case 2
         PopupMenu mnuModuleOptions
         
   End Select
   
End Sub

Private Sub miUnRegister_Click()
   Select Case MsgBox("Are you sure you want to UnRegister this module?", vbYesNo, "Remove Module")
      Case vbYes
         gPlugins.UnRegisterPlugin (lvwPlugins.SelectedItem)

   End Select
End Sub

Private Sub txtPlugin_Change()
   If (txtPlugin <> "") Then
      cmdAction(CMD_ADD).Enabled = True
   Else
      cmdAction(CMD_ADD).Enabled = False
   End If
End Sub
