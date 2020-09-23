VERSION 5.00
Begin VB.Form frmConfig 
   BorderStyle     =   0  'None
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
   Begin VB.CheckBox chkStayOnTop 
      Caption         =   "Stay on &Top"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   900
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.CheckBox chkCase 
      Caption         =   "Case &Sensitive"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   540
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Text Search Defaults:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
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
Private Sub chkCase_Click()
   If (chkCase.Value = vbChecked) Then
      gbCaseSensitive = True
   Else
      gbCaseSensitive = False
   End If
End Sub

Private Sub chkStayOnTop_Click()
   If (chkStayOnTop.Value = vbChecked) Then
      gbStayOnTop = True
   Else
      gbStayOnTop = False
   End If
End Sub

Private Sub Form_Activate()
   chkCase.SetFocus
End Sub

Private Sub Form_Load()
   If (chkCase.Value = vbChecked) Then
      gbCaseSensitive = True
   Else
      gbCaseSensitive = False
   End If

   If (chkStayOnTop.Value = vbChecked) Then
      gbStayOnTop = True
   Else
      gbStayOnTop = False
   End If
End Sub
