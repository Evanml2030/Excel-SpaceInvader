VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SpaceObjectCntrlsCol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private SpaceObjectsCntrls As Collection

Private Sub Class_Initialize()
    Set SpaceObjectsCntrls = New Collection
End Sub

Private Sub Class_Terminate()
    Set SpaceObjectsCntrls = Nothing
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = SpaceObjectsCntrls.[_NewEnum]
End Property

Public Sub Add(obj As Control)
    SpaceObjectsCntrls.Add obj
End Sub

Public Sub Remove(index As Variant)
    SpaceObjectsCntrls.Remove index
End Sub

Public Property Get Item(index As Variant) As Control
Attribute Item.VB_UserMemId = 0
    Set Item = SpaceObjectsCntrls.Item(index)
End Property

Property Get Count() As Long
    Count = SpaceObjectsCntrls.Count
End Property

Public Sub Clear()
    Set SpaceObjectsCntrls = New Collection
End Sub

