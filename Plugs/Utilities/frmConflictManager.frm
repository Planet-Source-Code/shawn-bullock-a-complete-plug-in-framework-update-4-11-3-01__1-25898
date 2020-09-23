VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConflictManager 
   BorderStyle     =   0  'None
   Caption         =   "Conflict Manager"
   ClientHeight    =   4395
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lvwHooks 
      Height          =   2595
      Left            =   2460
      TabIndex        =   1
      Top             =   360
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   4577
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Owner"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.ListBox lstFunctions 
      Height          =   2595
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2115
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Function Name"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2115
   End
End
Attribute VB_Name = "frmConflictManager"
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
Private Sub Form_Activate()
   If (Visible = True) Then
      'Text1.SetFocus
   End If
End Sub
