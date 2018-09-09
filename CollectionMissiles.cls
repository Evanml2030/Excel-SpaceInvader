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
Private CollectionMissles As Collection

Private Sub Class_Initialize()
    Set CollectionMissles = New Collection
End Sub

Private Sub Class_Terminate()
    Set CollectionMissles = Nothing
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = CollectionMissles.[_NewEnum]
End Property

Public Sub Add(obj As IBoundControl)
    CollectionMissles.Add obj
End Sub

Public Sub remove(index As Variant)
    CollectionMissles.remove index
End Sub

Public Property Get Item(index As Variant) As IBoundControl
Attribute Item.VB_UserMemId = 0
    Set Item = CollectionMissles.Item(index)
End Property

Property Get Count() As Long
    Count = CollectionMissles.Count
End Property

Public Sub Clear()
    Set CollectionMissles = New Collection
End Sub
