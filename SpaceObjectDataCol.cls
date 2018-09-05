VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SpaceObjectDataCol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private SpaceObjectsData As Collection

Private Sub Class_Initialize()
    Set SpaceObjectsData = New Collection
End Sub

Private Sub Class_Terminate()
    Set SpaceObjectsData = Nothing
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = SpaceObjectsData.[_NewEnum]
End Property

Public Sub Add(obj As ISpaceObject)
    SpaceObjectsData.Add obj
End Sub

Public Sub Remove(index As Variant)
    SpaceObjectsData.Remove index
End Sub

Public Property Get Item(index As Variant) As ISpaceObject
Attribute Item.VB_UserMemId = 0
    Set Item = SpaceObjectsData.Item(index)
End Property

Property Get Count() As Long
    Count = SpaceObjectsData.Count
End Property

Public Sub Clear()
    Set SpaceObjectsData = New Collection
End Sub
