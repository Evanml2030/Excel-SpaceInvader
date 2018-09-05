VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ScaleItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private plargestSize As Long

Public Property Get MaxSize() As Long
    MaxSize = plargestSize
End Property

Public Property Let MaxSize(ByRef value As Long)
    plargestSize = value
End Property


