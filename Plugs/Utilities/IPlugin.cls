VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IPlugin"
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
Public Property Let Name(Value As String)
   '
End Property

Public Property Get Name() As String
   '
End Property


Public Property Let Startup(Value As Boolean)
   '
End Property

Public Property Get Startup() As Boolean
   '
End Property


Public Property Let Path(Value As String)
   '
End Property

Public Property Get Path() As String
   '
End Property


Public Property Let Interface(Value As String)
   '
End Property

Public Property Get Interface() As String
   '
End Property


Public Property Let Description(Value As String)
   '
End Property

Public Property Get Description() As String
   '
End Property


Public Property Let Display(Value As String)
   '
End Property

Public Property Get Display() As String
   '
End Property


Public Property Let Supports(Value As String)
   '
End Property

Public Property Get Supports() As String
   '
End Property


Public Function Load() As Long
   '
End Function

Public Function UnLoad() As Long
   '
End Function


Public Function Activate() As Long
   '
End Function

Public Function Deactivate() As Long
   '
End Function


Public Property Let IsActive(Value As Boolean)
   '
End Property

Public Property Get IsActive() As Boolean
   '
End Property


Public Property Let Status(Value As Long)
   '
End Property

Public Property Get Status() As Long
   '
End Property


Public Function UpdateDefault(isNode As String, Value As String) As Long
   '
End Function

Public Function GetDefault(isNode As String) As Variant
   '
End Function


