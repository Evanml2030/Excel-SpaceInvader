VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CollectionMissiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private CollectionMissiles As Collection
Private pCount As Long

Private Sub Class_Initialize()
    Set CollectionMissiles = New Collection
End Sub

Private Sub Class_Terminate()
    Set CollectionMissiles = Nothing
End Sub

Public Sub Add(obj As ISpaceObject)
    CollectionMissiles.Add obj
End Sub

Public Sub Remove(index As Variant)
    CollectionMissiles.Remove index
End Sub

Public Property Get Item(index As Variant) As ISpaceObject
    Set Item = CollectionMissiles.Item(index)
End Property

Property Get Count() As Long
    Count = CollectionMissiles.Count
End Property

Public Sub Clear()
    Set CollectionMissiles = New Collection
End Sub

