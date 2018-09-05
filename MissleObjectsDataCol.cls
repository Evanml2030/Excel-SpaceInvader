VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MissleObjectsDataCol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private MissleObjectsData As Collection

Private Sub Class_Initialize()
    Set MissleObjectsData = New Collection
End Sub

Private Sub Class_Terminate()
    Set MissleObjectsData = Nothing
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = MissleObjectsData.[_NewEnum]
End Property

Public Sub Add(obj As missle)
    MissleObjectsData.Add obj
End Sub

Public Sub Remove(index As Variant)
    MissleObjectsData.Remove index
End Sub

Public Property Get Item(index As Variant) As missle
Attribute Item.VB_UserMemId = 0
    Set Item = MissleObjectsData.Item(index)
End Property

Property Get Count() As Long
    Count = MissleObjectsData.Count
End Property

Public Sub Clear()
    Set MissleObjectsData = New Collection
End Sub
