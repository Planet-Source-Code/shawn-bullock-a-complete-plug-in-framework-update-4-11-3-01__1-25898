VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find"
   ClientHeight    =   1275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   4110
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   480
      TabIndex        =   3
      Top             =   60
      Width           =   2295
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   3000
      TabIndex        =   2
      Top             =   900
      Width           =   1035
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find &Next"
      Height          =   315
      Index           =   1
      Left            =   3000
      TabIndex        =   1
      Top             =   420
      Width           =   1035
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Height          =   315
      Index           =   0
      Left            =   3000
      TabIndex        =   0
      Top             =   60
      Width           =   1035
   End
   Begin VB.Label label1 
      Caption         =   "Find:"
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   60
      Width           =   735
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Target As String
Private FoundPos As Long


Private Declare Function SetWindowPos Lib "user32" ( _
      ByVal hWnd As Long, _
      ByVal hWndInsertAfter As Long, _
      ByVal X As Long, _
      ByVal Y As Long, _
      ByVal cx As Long, _
      ByVal cy As Long, _
      ByVal wFlags As Long _
   ) As Long


Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2


Private Sub cmdCancel_Click()
   txtSearch.Text = ""
   Me.Hide
End Sub

Private Sub cmdFind_Click(Index As Integer)
   
   If Index = 0 Then
      FoundPos = 0
      
      If (gbCaseSensitive = True) Then
         Target = txtSearch.Text
      Else
         Target = UCase(txtSearch.Text)
      End If
   End If
   
   If Target = "" Then
      Exit Sub
   End If
   
   If (gbCaseSensitive = True) Then
      FoundPos = InStr((FoundPos + 1), Application.ActiveDocument.Text, Target)
   Else
      FoundPos = InStr((FoundPos + 1), UCase(Application.ActiveDocument.Text), Target)
   End If
   
   If FoundPos > 0 Then
      Application.ActiveDocument.SetFocus
      
      Application.ActiveDocument.SelStart = (FoundPos - 1)
      Application.ActiveDocument.SelLength = Len(Target)
      
      cmdFind(1).Enabled = True
   Else
      MsgBox ("No matches found")
      
      Application.ActiveDocument.SelStart = 0
      cmdFind(1).Enabled = False
   End If
   
End Sub

Private Sub Form_Activate()
   If (frmConfig.chkCase.Value = vbChecked) Then
      gbCaseSensitive = True
   Else
      gbCaseSensitive = False
   End If
   
   SetTopMost
End Sub

Private Sub SetTopMost()
   Dim lResult As Long
    
   If (gbStayOnTop = True) Then
      lResult = SetWindowPos( _
            frmFind.hWnd, _
            HWND_TOPMOST, _
            0, _
            0, _
            0, _
            0, _
            FLAGS _
         )
   Else
      lResult = SetWindowPos( _
            frmFind.hWnd, _
            HWND_NOTOPMOST, _
            0, _
            0, _
            0, _
            0, _
            FLAGS _
         )
   End If
End Sub

