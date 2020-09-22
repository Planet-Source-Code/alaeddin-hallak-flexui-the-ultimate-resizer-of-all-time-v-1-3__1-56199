VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsControlArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsControlInfo"
Attribute VB_Ext_KEY = "Member0" ,"clsControlInfo"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable to hold collection
Private mCol As Collection

Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub

Public Sub Add(ControlRef As Object, Style As AnchorResizeStyle, Optional Width As Integer = 0, Optional Height As Integer = 0, Optional Top As Integer = 0, Optional Left As Integer = 0, Optional ContainerWidth As Integer = 0, Optional ContainerHeight As Integer = 0)
             
        'create a new object
        Dim objNewMember As clsControlInfo
        Set objNewMember = New clsControlInfo
    
        'set the properties passed into the method
        objNewMember.ContainerWidth = ContainerWidth
        objNewMember.ContainerHeight = ContainerHeight
        objNewMember.Width = Width
        objNewMember.Height = Height
        objNewMember.Top = Top
        objNewMember.Left = Left
        objNewMember.Style = Style
        objNewMember.Key = modUtilities.GetCtrlNameWithIndex(ControlRef)
        
        Set objNewMember.ControlRef = ControlRef
       
        mCol.Add objNewMember, objNewMember.Key
        
        'return the object created
        Set objNewMember = Nothing

End Sub

Public Property Get Item(vntIndexKey As Variant) As clsControlInfo
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
    On Error GoTo ErrHandler
  Set Item = mCol(vntIndexKey)
  
  Exit Property
  
ErrHandler:
  Err.Raise 5, "Item", "Key was not found in collection"
End Property


Public Function ContainsKey(ByVal strKey As String) As Boolean

    On Error Resume Next
    Dim x As clsControlInfo
    Set x = mCol.Item(strKey)
    ContainsKey = (Err.Number = 0)
    
End Function

Public Sub Clear()
    
    While mCol.Count > 0
        mCol.Remove 1
    Wend
    
End Sub

Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)
    
    If IsNumeric(vntIndexKey) Then
        mCol.Remove vntIndexKey
    Else
        If vntIndexKey <> "" Then
            mCol.Remove vntIndexKey
        End If
    End If
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property




Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

