VERSION 5.00
Begin VB.Form frmConfigGeneral 
   BorderStyle     =   0  'None
   Caption         =   "Settings"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   240
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   1155
      Left            =   360
      TabIndex        =   0
      Text            =   "Host Applications general settings would go here."
      Top             =   300
      Width           =   5355
   End
End
Attribute VB_Name = "frmConfigGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
    Text1.SetFocus
End Sub
