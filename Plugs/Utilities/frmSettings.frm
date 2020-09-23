VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   4755
   ClientLeft      =   225
   ClientTop       =   510
   ClientWidth     =   9915
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   9915
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   9000
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin MSComctlLib.ImageList imgSettings 
      Left            =   180
      Top             =   4020
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Frame fraConfig 
      Height          =   4695
      Left            =   2700
      TabIndex        =   1
      Top             =   0
      Width           =   6195
   End
   Begin MSComctlLib.TreeView tvwModules 
      Height          =   4635
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   8176
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   441
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   7
      FullRowSelect   =   -1  'True
      BorderStyle     =   1
      Appearance      =   1
   End
End
Attribute VB_Name = "frmSettings"
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
Private mSettingScreen As Form


Private Sub cmdOK_Click()
   Hide
End Sub

Private Sub tvwModules_NodeClick(ByVal Node As MSComctlLib.Node)
   If (Node.Tag <> "") Then
      gUtilityEvents.Send_SelectSetting (Node.Tag)
   End If
End Sub

Public Function SetActiveSetting(SettingScreen As Object) As Long
   Dim lnWindowStyle As Long
   Dim lnResult As Long
   
   If (Not mSettingScreen Is Nothing) Then
      mSettingScreen.Hide
      Set mSettingScreen = Nothing
   End If
   
   Set mSettingScreen = SettingScreen

   DoEvents
   If (fraConfig.hwnd <> GetParent(mSettingScreen.hwnd)) Then
      SetParent mSettingScreen.hwnd, fraConfig.hwnd
      
      lnWindowStyle = GetWindowLong(mSettingScreen.hwnd, GWL_STYLE)
      lnResult = SetWindowLong(mSettingScreen.hwnd, GWL_STYLE, lnWindowStyle Or WS_CHILD)
      
   End If
   
   mSettingScreen.Top = X(8)
   mSettingScreen.Left = X(2)
   
   mSettingScreen.Show
   
End Function
