VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ActiveDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mbChanged As Boolean
Private msDocPath As String


Public Property Let SelStart(Value As Long)
   gMain.rtfDocument.SelStart = Value
End Property

Public Property Get SelStart() As Long
   SelStart = gMain.rtfDocument.SelStart
End Property


Public Property Let SelLength(Value As Long)
   gMain.rtfDocument.SelLength = Value
End Property

Public Property Get SelLength() As Long
   SelLength = gMain.rtfDocument.SelLength
End Property


Public Property Let SelText(Value As String)
   gMain.rtfDocument.SelText = Value
End Property

Public Property Get SelText() As String
   SelText = gMain.rtfDocument.SelText
End Property


Public Property Let TextRTF(Value As String)
   gMain.rtfDocument.TextRTF = Value
End Property

Public Property Get TextRTF() As String
   TextRTF = gMain.rtfDocument.TextRTF
End Property


Public Property Let Text(Value As String)
   gMain.rtfDocument.Text = Value
End Property

Public Property Get Text() As String
   Text = gMain.rtfDocument.Text
End Property


Public Property Let Changed(Value As Boolean)
   mbChanged = Value
End Property

Public Property Get Changed() As Boolean
   Changed = mbChanged
End Property


Public Property Let DocPath(Value As String)
   msDocPath = Value
End Property

Public Property Get DocPath() As String
   DocPath = msDocPath
End Property


Public Sub SaveFile(FilePath As String)
   gAppEvents.Send_SaveDocument
   
   gMain.rtfDocument.SaveFile (FilePath)
   DocPath = FilePath
   Changed = False
End Sub

Public Sub OpenFile(FilePath As String)
   gAppEvents.Send_OpenDocument
   
   gMain.rtfDocument.LoadFile (FilePath)
   DocPath = FilePath
End Sub


Public Sub SetFocus()
   gMain.rtfDocument.SetFocus
End Sub

