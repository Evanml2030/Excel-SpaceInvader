RSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SpaceObjectCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pcount As Long

Public Property Get Count() As Long
    Count = pcount
End Property

Public Property Let Count(ByRef value As Long)
    pcount = value
End Property

Public Sub IncrementCount()
    pcount = pcount + 1
End Sub
