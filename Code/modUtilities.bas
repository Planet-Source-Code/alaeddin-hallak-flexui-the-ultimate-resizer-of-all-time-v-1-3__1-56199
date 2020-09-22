Attribute VB_Name = "modUtilities"
Option Explicit

Public Function GetCtrlNameWithIndex(ctrl As Object) As String

    On Error Resume Next
    
    GetCtrlNameWithIndex = ctrl.Name & IIf(ctrl.Index = -1, "", "(" & ctrl.Index & ")")
    
    If Err.Number <> 0 Then GetCtrlNameWithIndex = ctrl.Name
    
End Function

Public Function IsSupportedControl(obj As Object) As Boolean

  IsSupportedControl = (Not (TypeOf obj Is FlexUI) And _
                        Not (TypeOf obj Is Menu) And _
                        SupportsContainerProperty(obj))
                        
End Function

Private Function SupportsContainerProperty(obj As Object)
    Static str As String
    On Error Resume Next
    Err.Clear
    str = obj.Container.Name
    SupportsContainerProperty = Not (Err.Number = 438)
End Function

Public Function CheckBit(ByVal Value As Variant, ByVal BitNo As Integer) As Variant
    CheckBit = (Value And (BitNo)) = BitNo
End Function
