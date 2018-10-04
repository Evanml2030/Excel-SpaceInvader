VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CollectionIncomingSpaceObjects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private CollectionIncomingSpaceObjects As Collection

Private Sub Class_Initialize()
    Set CollectionIncomingSpaceObjects = New Collection
End Sub

Private Sub Class_Terminate()
    Set CollectionIncomingSpaceObjects = Nothing
End Sub

Public Sub Add(obj As ISpaceObject)
    CollectionIncomingSpaceObjects.Add obj
End Sub

Public Sub Remove(index As Variant)
    CollectionIncomingSpaceObjects.Remove index
End Sub

Public Property Get Item(index As Variant) As ISpaceObject
    Set Item = CollectionIncomingSpaceObjects.Item(index)
End Property

Property Get Count() As Long
    Count = CollectionIncomingSpaceObjects.Count
End Property

Public Sub Clear()
    Set CollectionIncomingSpaceObjects = New Collection
End Sub

