VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Events"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event Initialize()
Public Event Quit()
Public Event ModuleSelected(Module As String)
Public Event MenuSelected(MenuItem As String, MenuTag As Variant)


Public Sub Send_Initialize()
   '
   ' All plugins will be loaded, and then we will send the initialize event and each
   '  plugin will recieve it.  Once they recieve it, they can respond accordingly
   '
   RaiseEvent Initialize
End Sub

Public Sub Send_Quit()
   RaiseEvent Quit
End Sub

Public Sub Send_ModuleSelected(Module As String)
   RaiseEvent ModuleSelected(Module)
End Sub

Public Sub Send_MenuSelected(MenuItem As String, MenuTag As Variant)
   RaiseEvent MenuSelected(MenuItem, MenuTag)
End Sub

