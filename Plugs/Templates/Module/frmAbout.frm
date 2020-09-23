VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
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
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   780
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmAbout.frx":0000
      Top             =   840
      Width           =   4575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Plugin Module"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   5175
   End
End
Attribute VB_Name = "frmAbout"
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
   '
   ' We must set some control on the form to have focus, that way when the form
   '  becomes active in the setting screen, it won't flicker
   '
   Text1.SetFocus
End Sub
