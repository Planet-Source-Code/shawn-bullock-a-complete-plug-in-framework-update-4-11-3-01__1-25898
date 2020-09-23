VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmConfig 
   BorderStyle     =   0  'None
   Caption         =   "dwAutoSave"
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
   Visible         =   0   'False
   Begin MSComDlg.CommonDialog CD 
      Left            =   5460
      Top             =   660
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Left            =   5520
      Top             =   120
   End
   Begin VB.CheckBox chkSave 
      Height          =   195
      Left            =   540
      TabIndex        =   3
      Top             =   720
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.TextBox txtMinutes 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Text            =   "5"
      Top             =   720
      Width           =   555
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Save every                  minutes"
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   2
      Top             =   720
      Width           =   2235
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Autosave options:"
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   3315
   End
End
Attribute VB_Name = "frmConfig"
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
Private mnCount As Long


Private Sub Form_Activate()
   txtMinutes.SetFocus
End Sub

Private Sub chkSave_Click()
   txtMinutes_LostFocus
End Sub

Private Sub txtMinutes_Change()
   txtMinutes_LostFocus
End Sub

Private Sub txtMinutes_LostFocus()
   If (chkSave.Value = vbChecked) Then
      Timer1.Interval = 60000
      mnCount = 0
   Else
      Timer1.Interval = 0
   End If
End Sub

Private Sub Timer1_Timer()
   mnCount = (mnCount + 1)
   
   If (mnCount < txtMinutes) Then
      Exit Sub
   End If
   
   If (Application.ActiveDocument.Changed = False) Then
      mnCount = 0
      Exit Sub
   End If
   
   If (Application.ActiveDocument.DocPath = "") Then
      With CD
         .Filter = "RTF Document (*.rtf)|*.rtf"
         .ShowSave
         
         If (.FileName <> "") Then
            Application.ActiveDocument.SaveFile (.FileName)
         End If
      End With
   Else
      Application.ActiveDocument.SaveFile (Application.ActiveDocument.DocPath)
   End If
   
   mnCount = 0
End Sub
