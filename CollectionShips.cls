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

Private Sub Class_Initialize()
    Set CollectionShips = New Collection
End Sub

Private Sub Class_Terminate()
    Set CollectionShips = Nothing
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = CollectionShips.[_NewEnum]
End Property

Public Sub Add(obj As IBoundControl)
    CollectionShips.Add obj
End Sub

Public Sub remove(index As Variant)
    CollectionShips.remove index
End Sub

Public Property Get Item(index As Variant) As IBoundControl
Attribute Item.VB_UserMemId = 0
    Set Item = CollectionShips.Item(index)
End Property

Property Get Count() As Long
    Count = CollectionShips.Count
End Property

Public Sub Clear()
    Set CollectionShips = New Collection
End Sub
