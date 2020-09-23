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
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
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
Public Event Initialize()
Public Event Quit()
Public Event ModuleSelected(Module As String)
Public Event MenuSelected(MenuItem As String, MenuTag As Variant)
Public Event OpenDocument()
Public Event SaveDocument()
Public Event SaveDocumentAs()



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

Public Sub Send_OpenDocument()
   RaiseEvent OpenDocument
End Sub

Public Sub Send_SaveDocument()
   RaiseEvent SaveDocument
End Sub

Public Sub Send_SaveDocumentAs()
   RaiseEvent SaveDocumentAs
End Sub
