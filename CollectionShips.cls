VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CollectionShips"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private CollectionShips As Collection
Private pCount As Long

Private Sub Class_Initialize()
    Set CollectionShips = New Collection
End Sub

Private Sub Class_Terminate()
    Set CollectionShips = Nothing
End Sub

Public Sub Add(obj As ISpaceObject)
    CollectionShips.Add obj
End Sub

Public Sub Remove(index As Variant)
    CollectionShips.Remove index
End Sub

Public Property Get Item(index As Variant) As ISpaceObject
    Set Item = CollectionShips.Item(index)
End Property

Property Get Count() As Long
    Count = CollectionShips.Count
End Property

Public Sub Clear()
    Set CollectionShips = New Collection
End Sub

