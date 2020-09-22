VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsControlTracker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private m_ControlArray As clsControlArray
Attribute m_ControlArray.VB_VarUserMemId = 0
Private m_ParentControls As Object
Private m_intNumResizerControlsPresent As Integer
Private m_intParentControlsCount As Integer

Private Sub Class_Initialize()
    Set m_ControlArray = New clsControlArray
End Sub

Public Property Get ParentControls() As Object
    Set ParentControls = m_ParentControls
End Property

Public Property Set ParentControls(ByVal vNewValue As Object)
    Set m_ParentControls = vNewValue
End Property

Public Property Get ControlArray() As clsControlArray
    Set ControlArray = m_ControlArray
End Property


Public Sub InitControlArray()
    
    If m_ParentControls Is Nothing Then Exit Sub

    Dim ctrl As Object
    Dim tmp As clsControlInfo
    
    For Each ctrl In m_ParentControls
        
        If modUtilities.IsSupportedControl(ctrl) Then
            m_ControlArray.Add ctrl, None
        End If
        
    Next
    
End Sub


Public Sub CheckControlsCount()
    
    If m_ParentControls Is Nothing Then Exit Sub
    'Count parent controls
    Dim ctrl As Object
    
    m_intNumResizerControlsPresent = 0
    m_intParentControlsCount = 0
    
    For Each ctrl In m_ParentControls
        
        If TypeOf ctrl Is FlexUI Then
            m_intNumResizerControlsPresent = m_intNumResizerControlsPresent + 1
        
        ElseIf modUtilities.IsSupportedControl(ctrl) Then
               m_intParentControlsCount = m_intParentControlsCount + 1
               
        End If

    Next
    
    If m_intParentControlsCount > m_ControlArray.Count Then
        'User have added controls since the last time we filled the array
        AddAddedControls
    ElseIf m_intParentControlsCount < m_ControlArray.Count Then
        'User have deleted controls since the last time we filled the array
        RemoveDeletedControls
    End If
    
End Sub

Public Function Serialize() As String
    
    Dim i As clsControlInfo
    Dim str As String
    
    For Each i In m_ControlArray
        str = str & i.Key & "|" & i.Style & "*"
    Next
    
    Serialize = str
    
End Function

Public Sub Deserialize(ByVal str As String)
    
    m_ControlArray.Clear
    
    Dim i As Integer
    Dim tokens() As String, tmp() As String
    Dim ctrl As Object
        
    tokens = Split(str, "*", , vbBinaryCompare)
   
    For i = 0 To UBound(tokens) - 1
        tmp = Split(tokens(i), "|", , vbBinaryCompare)
        
        Set ctrl = GetControlRef(tmp(0))
        
        If Not ctrl Is Nothing Then
            m_ControlArray.Add ctrl, Val(tmp(1))
        End If
    Next i
    
End Sub


Private Sub AddAddedControls()

    If m_ParentControls Is Nothing Then Exit Sub

    Dim ctrl As Object
    
    For Each ctrl In m_ParentControls
        
        If modUtilities.IsSupportedControl(ctrl) Then
           
            If Not m_ControlArray.ContainsKey(modUtilities.GetCtrlNameWithIndex(ctrl)) Then
                
                m_ControlArray.Add ctrl, None
                
            End If
        
        End If
    Next
    
End Sub

Private Sub RemoveDeletedControls()
    
    If m_ParentControls Is Nothing Then Exit Sub

    Dim i As clsControlInfo, ctrl As Object
    Dim bDeletedCtrl As Boolean
    
    'On Error GoTo Err
    
    For Each i In m_ControlArray
    
        bDeletedCtrl = True
        
        For Each ctrl In m_ParentControls
            If ctrl Is i.ControlRef Then
                bDeletedCtrl = False
                Exit For
            End If
        Next
        
        If bDeletedCtrl Then m_ControlArray.Remove i.Key
        
    Next
    
Err:
    
End Sub

Private Function GetControlRef(ByVal strCtrlName As String) As Object
    Dim ctrl As Object
    
    If m_ParentControls Is Nothing Then
        Set GetControlRef = Nothing
        Exit Function
    End If
    
    For Each ctrl In m_ParentControls
        If modUtilities.GetCtrlNameWithIndex(ctrl) = strCtrlName Then
            Set GetControlRef = ctrl
            Exit Function
        End If
    Next ctrl
    
    Set GetControlRef = Nothing
End Function

Private Sub Class_Terminate()
    Set m_ControlArray = Nothing
End Sub
