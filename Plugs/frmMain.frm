VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dynamic Word"
   ClientHeight    =   5895
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   9180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   9180
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CD 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox rtfDocument 
      Height          =   5895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   10398
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmMain.frx":0000
   End
   Begin VB.Menu mbrFile 
      Caption         =   "&File"
      Begin VB.Menu miOpen 
         Caption         =   "&Open..."
      End
      Begin VB.Menu miBreak1 
         Caption         =   "-"
      End
      Begin VB.Menu miSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu miSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu miBreak2 
         Caption         =   "-"
      End
      Begin VB.Menu miFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu miFind 
         Caption         =   "&Find..."
         Shortcut        =   ^F
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
         Shortcut        =   ^T
      End
   End
End
Attribute VB_Name = "frmMain"
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
Private Const MODULE_MANAGER As Long = 0
Private Const MODULE_BREAK As Long = 1


Public Function AddModuleToModules( _
      isModuleName As String, _
      isModuleIdentifier As String _
   ) As Long
   
   Dim lnCount As Long
   
   DoEvents
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

Private Sub miFind_Click()
   gAppEvents.Send_MenuSelected miFind.Caption, miFind.Tag
End Sub


Private Sub miModules_Click(Index As Integer)
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

Private Sub miOpen_Click()
   With CD
      .ShowOpen
      
      If (.FileName <> "") Then
         Application.ActiveDocument.OpenFile .FileName
      End If
   End With
End Sub

Private Sub miSave_Click()
   With CD
      If (Application.ActiveDocument.DocPath <> "") Then
         Application.ActiveDocument.SaveFile (Application.ActiveDocument.DocPath)
      Else
         miSaveAs_Click
      End If
   End With
End Sub

Private Sub miSaveAs_Click()
   With CD
      .Filter = "RTF Document (*.rtf)|*.rtf"
      .ShowSave
      
      If (.FileName <> "") Then
         Application.ActiveDocument.SaveFile .FileName
      End If
   End With
End Sub

Private Sub miToolsSettings_Click()
   Application.Settings.ShowSettings
End Sub


Private Sub rtfDocument_Change()
   Application.ActiveDocument.Changed = True
End Sub
