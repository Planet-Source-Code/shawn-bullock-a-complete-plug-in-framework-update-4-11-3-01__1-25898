Attribute VB_Name = "basMain"
Option Explicit

Public Application As Object '[Host].Application  ' Be sure to reference


Public Sub Message(isText As String)
    If (Application.Visible = True) Then
        MsgBox isText
    End If
End Sub
