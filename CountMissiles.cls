VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CountMissiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pCount As Long

Public Property Get Count() As Long
    Count = pCount
End Property

Private Property Let Count(ByRef value As Long)
    pCount = value
End Property

Public Sub IncrementMissileCount()
    pCount = pCount + 1
End Sub


