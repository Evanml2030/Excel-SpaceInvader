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
Private CollectionIncSpaceObjects As Collection

Private Sub Class_Initialize()
    Set CollectionIncSpaceObjects = New Collection
End Sub

Private Sub Class_Terminate()
    Set CollectionIncSpaceObjects = Nothing
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = CollectionIncSpaceObjects.[_NewEnum]
End Property

Public Sub Add(obj As IBoundControl)
    CollectionIncSpaceObjects.Add obj
End Sub

Public Sub remove(index As Variant)
    CollectionIncSpaceObjects.remove index
End Sub

Public Property Get Item(index As Variant) As IBoundControl
Attribute Item.VB_UserMemId = 0
    Set Item = CollectionIncSpaceObjects.Item(index)
End Property

Property Get Count() As Long
    Count = CollectionIncSpaceObjects.Count
End Property

Public Sub Clear()
    Set CollectionIncSpaceObjects = New Collection
End Sub
