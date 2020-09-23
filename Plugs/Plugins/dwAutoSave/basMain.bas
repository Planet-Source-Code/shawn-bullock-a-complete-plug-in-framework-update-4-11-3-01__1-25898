Attribute VB_Name = "basMain"
Option Explicit

Public Application As Dynamic_Word.Application


Public Sub Message(isText As String)
    If (Application.Visible = True) Then
        MsgBox isText
    End If
End Sub
