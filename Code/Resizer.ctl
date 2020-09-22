VERSION 5.00
Begin VB.UserControl FlexUI 
   ClientHeight    =   450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   465
   InvisibleAtRuntime=   -1  'True
   Picture         =   "Resizer.ctx":0000
   PropertyPages   =   "Resizer.ctx":0CCA
   ScaleHeight     =   450
   ScaleWidth      =   465
   ToolboxBitmap   =   "Resizer.ctx":0D02
   Windowless      =   -1  'True
End
Attribute VB_Name = "FlexUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'==================================================================================================
'Note: Credits go to Paul Caton (Paul_Caton@hotmail.com) for writing this awsome subclassing code
'==================================================================================================

'Subclasser declarations

Private Const WM_SIZE                As Long = &H5
Private Const WM_GETMINMAXINFO       As Long = &H24
      
Private Enum eMsgWhen
  MSG_AFTER = 1                                                                         'Message calls back after the original (previous) WndProc
  MSG_BEFORE = 2                                                                        'Message calls back before the original (previous) WndProc
  MSG_BEFORE_AND_AFTER = MSG_AFTER Or MSG_BEFORE                                        'Message calls back before and after the original (previous) WndProc
End Enum

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type
      
Private Const ALL_MESSAGES           As Long = -1                                       'All messages added or deleted
Private Const GMEM_FIXED             As Long = 0                                        'Fixed memory GlobalAlloc flag
Private Const GWL_WNDPROC            As Long = -4                                       'Get/SetWindow offset to the WndProc procedure address
Private Const PATCH_04               As Long = 88                                       'Table B (before) address patch offset
Private Const PATCH_05               As Long = 93                                       'Table B (before) entry count patch offset
Private Const PATCH_08               As Long = 132                                      'Table A (after) address patch offset
Private Const PATCH_09               As Long = 137                                      'Table A (after) entry count patch offset

Private Type tSubData                                                                   'Subclass data type
  hWnd                               As Long                                            'Handle of the window being subclassed
  nAddrSub                           As Long                                            'The address of our new WndProc (allocated memory).
  nAddrOrig                          As Long                                            'The address of the pre-existing WndProc
  nMsgCntA                           As Long                                            'Msg after table entry count
  nMsgCntB                           As Long                                            'Msg before table entry count
  aMsgTblA()                         As Long                                            'Msg after table array
  aMsgTblB()                         As Long                                            'Msg Before table array
End Type

Private sc_aSubData()                As tSubData                                        'Subclass data array

Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Sub CopyMemoryToMinMaxInfo Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As MINMAXINFO, ByVal hpvSource As Long, ByVal cbCopy As Long)
Private Declare Sub CopyMemoryFromMinMaxInfo Lib "kernel32" Alias "RtlMoveMemory" (ByVal hpvDest As Long, hpvSource As MINMAXINFO, ByVal cbCopy As Long)
'==================================================================================================

Public Enum AnchorResizeStyle
    None = 32
    LeftAnchor = 33
    RightAnchor = 34
    ResizeHorizontally = 35
    TopAnchor = 36
    TopAnchor_LeftAnchor = 37
    TopAnchor_RightAnchor = 38
    TopAnchor_ResizeHorizontally = 39
    BottomAnchor = 40
    BottomAnchor_LeftAnchor = 41
    BottomAnchor_RightAnchor = 42
    BottomAnchor_ResizeHorizontally = 43
    ResizeVertically = 44
    ResizeVertically_LeftAnchor = 45
    ResizeVertically_RightAnchor = 46
    ResizeVertically_ResizeHorizontally = 47
    RTLEffect = 16
End Enum

Private m_ControlTracker As clsControlTracker
Private m_intMinFormHeight As Integer, m_intMinFormWidth As Integer
Private m_intMaxFormHeight As Integer, m_intMaxFormWidth As Integer

Private Sub UserControl_Initialize()

    Set m_ControlTracker = New clsControlTracker
   
End Sub

Private Sub UserControl_InitProperties()
    
    Set m_ControlTracker.ParentControls = UserControl.Parent.controls
    
    MinFormWidth = UserControl.Parent.Width / Screen.TwipsPerPixelX
    MinFormHeight = UserControl.Parent.Height / Screen.TwipsPerPixelY
    Enabled = True
    MaxFormWidth = -1
    MaxFormHeight = -1
    
    m_ControlTracker.InitControlArray
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
        
    Enabled = PropBag.ReadProperty("Enabled", True)
    MinFormWidth = PropBag.ReadProperty("MinFormWidth", UserControl.Parent.Width / Screen.TwipsPerPixelX)
    MinFormHeight = PropBag.ReadProperty("MinFormHeight", UserControl.Parent.Height / Screen.TwipsPerPixelY)
    MaxFormWidth = PropBag.ReadProperty("MaxFormWidth", -1)
    MaxFormHeight = PropBag.ReadProperty("MaxFormHeight", -1)
    
    Set m_ControlTracker.ParentControls = UserControl.Parent.controls
   
    m_ControlTracker.Deserialize PropBag.ReadProperty("ControlsSetupString", "")
    
    m_ControlTracker.CheckControlsCount
    
    If Ambient.UserMode Then
    
        InitResizer
    
        With UserControl.Parent
            Call Subclass_Start(.hWnd)
            Call Subclass_AddMsg(.hWnd, WM_GETMINMAXINFO, MSG_AFTER)
            Call Subclass_AddMsg(.hWnd, WM_SIZE, MSG_AFTER)
        End With
        
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    With PropBag
        m_ControlTracker.CheckControlsCount
        .WriteProperty "ControlsSetupString", ControlsSetupString, ""
        .WriteProperty "Enabled", UserControl.Enabled, True
        .WriteProperty "MinFormWidth", m_intMinFormWidth, -1
        .WriteProperty "MinFormHeight", m_intMinFormHeight, -1
        .WriteProperty "MaxFormWidth", m_intMaxFormWidth, -1
        .WriteProperty "MaxFormHeight", m_intMaxFormHeight, -1
    End With
    
End Sub

'======================================================================================================
'Subclass handler - MUST be the first Public routine in this file. That includes public properties also

Public Sub zSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lng_hWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)
Attribute zSubclass_Proc.VB_MemberFlags = "40"
'Parameters:
  'bBefore  - Indicates whether the the message is being processed before or after the default handler - only really needed if a message is set to callback both before & after.
  'bHandled - Set this variable to True in a 'before' callback to prevent the message being subsequently processed by the default handler... and if set, an 'after' callback
  'lReturn  - Set this variable as per your intentions and requirements, see the MSDN documentation for each individual message value.
  'hWnd     - The window handle
  'uMsg     - The message number
  'wParam   - Message related data
  'lParam   - Message related data
'Notes:
  'If you really know what you're doing, it's possible to change the values of the
  'hWnd, uMsg, wParam and lParam parameters in a 'before' callback so that different
  'values get passed to the default handler.. and optionaly, the 'after' callback
 
  Select Case uMsg
  
  Case WM_GETMINMAXINFO
        
        Dim MinMax As MINMAXINFO
        
        'Retrieve default MinMax settings
        CopyMemoryToMinMaxInfo MinMax, lParam, Len(MinMax)
    
        'Specify new minimum size for window.
        If m_intMinFormWidth <> -1 Then
          MinMax.ptMinTrackSize.x = m_intMinFormWidth
        End If
        
        If m_intMinFormHeight <> -1 Then
          MinMax.ptMinTrackSize.y = m_intMinFormHeight
        End If
    
        'Specify new maximum size for window.
        If m_intMaxFormWidth <> -1 Then
          MinMax.ptMaxTrackSize.x = m_intMaxFormWidth
        End If
        
        If m_intMaxFormHeight <> -1 Then
          MinMax.ptMaxTrackSize.y = m_intMaxFormHeight
        End If
    
        'Copy local structure back.
        CopyMemoryFromMinMaxInfo lParam, MinMax, Len(MinMax)
    
  Case WM_SIZE
        Resize
        
  End Select
  
  'WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lParam)
  
End Sub

Public Sub CenterFormScreen()
Attribute CenterFormScreen.VB_Description = "Centers the form in the middle of the screen"
    
    On Error Resume Next
    UserControl.Parent.Move (Screen.Width - UserControl.Parent.Width) / 2, (Screen.Height - UserControl.Parent.Height) / 2, UserControl.Parent.Width, UserControl.Parent.Height

End Sub

Private Sub InitResizer()
Attribute InitResizer.VB_Description = "Initilizes the component. Invoke this method in your Form's Load event"

    Dim i As clsControlInfo
    
    On Error Resume Next
    
    For Each i In m_ControlTracker.ControlArray
                        
        If TypeOf i.ControlRef Is Line Then
            i.Height = i.ControlRef.Y2 - i.ControlRef.Y1
            i.Width = i.ControlRef.X2 - i.ControlRef.X1
            i.Left = i.ControlRef.X1
            i.Top = i.ControlRef.Y1
        Else
            i.Height = i.ControlRef.Height
            i.Width = i.ControlRef.Width
            i.Left = i.ControlRef.Left
            i.Top = i.ControlRef.Top
        End If
        
        i.ContainerHeight = i.ControlRef.Container.Height
        i.ContainerWidth = i.ControlRef.Container.Width
        
    Next i
    
End Sub

Public Sub SetControlStyle(ctrl As Object, Style As AnchorResizeStyle)
Attribute SetControlStyle.VB_Description = "Sets the anchor/resize style of the specified control"
    
    On Error GoTo ErrHandler
    
    m_ControlTracker.ControlArray(modUtilities.GetCtrlNameWithIndex(ctrl)).Style = Style
    
    Exit Sub
    
ErrHandler:
    If Err.Number = 5 Then
        'Passed control doesn't exist in our ControlArray collection. Add it with the specified style
        
        If modUtilities.IsSupportedControl(ctrl) Then
           
                m_ControlTracker.ControlArray.Add ctrl, Style, ctrl.Width, ctrl.Height, ctrl.Top, ctrl.Left, ctrl.Container.Width, ctrl.Container.Height
                
        End If
        
    End If
End Sub

Public Function GetControlStyle(ctrl As Object) As AnchorResizeStyle
Attribute GetControlStyle.VB_Description = "Returns the anchor/resize style of the specified control"
    
    On Error Resume Next
    GetControlStyle = m_ControlTracker.ControlArray(modUtilities.GetCtrlNameWithIndex(ctrl)).Style
    
    If Err.Number <> 0 Then
        GetControlStyle = None
    End If
End Function

Public Sub UpdateControlPositionSize(ctrl As Object)
Attribute UpdateControlPositionSize.VB_Description = "Updates the component's internal copy of the specified control's position and size. You must call this method after you modify the size or position of a control in order to maintain correct anchoring/resizing behaviour"
    
    On Error GoTo Err
    
    With m_ControlTracker.ControlArray(modUtilities.GetCtrlNameWithIndex(ctrl))
        .Height = ctrl.Height
        .Width = ctrl.Width
        .Top = ctrl.Top
        .Left = ctrl.Left
        .ContainerHeight = ctrl.Container.Height
        .ContainerWidth = ctrl.Container.Width
    End With

Err:
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = 465
    UserControl.Height = 450
End Sub

Public Sub About()
Attribute About.VB_Description = "Displays component information dialog box"
Attribute About.VB_UserMemId = -552
    On Error GoTo Err
    Load frmAbout
    frmAbout.picIcon = UserControl.Picture
    frmAbout.Show vbModal
    Unload frmAbout
    Set frmAbout = Nothing
    
Err:
    
End Sub

Public Property Get ControlTracker() As clsControlTracker
Attribute ControlTracker.VB_MemberFlags = "40"
    Set ControlTracker = m_ControlTracker
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets whether the control is enabled. When set to False, the control will not respond to Resize calls."
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal vNewValue As Boolean)
    UserControl.Enabled = vNewValue
End Property

Public Property Get ControlsSetupString() As String
Attribute ControlsSetupString.VB_ProcData.VB_Invoke_Property = "SetupControls"
Attribute ControlsSetupString.VB_MemberFlags = "640"
    ControlsSetupString = m_ControlTracker.Serialize
End Property

Public Property Let ControlsSetupString(ByVal vNewValue As String)
    If Ambient.UserMode Then
        Err.Raise Number:=382, Description:="Let/Set not supported at run time."
    Else
        PropertyChanged "ControlsSetupString"
    End If
End Property

Public Property Get MinFormWidth() As Integer
Attribute MinFormWidth.VB_Description = "Returns/sets the minimum resized width of the form (-1 to ignore)"
    MinFormWidth = m_intMinFormWidth
End Property

Public Property Let MinFormWidth(ByVal vNewValue As Integer)
    m_intMinFormWidth = vNewValue
End Property

Public Property Get MinFormHeight() As Integer
Attribute MinFormHeight.VB_Description = "Returns/sets the minimum resized height of the form (-1 to ignore)"
    MinFormHeight = m_intMinFormHeight
End Property

Public Property Let MinFormHeight(ByVal vNewValue As Integer)
    m_intMinFormHeight = vNewValue
End Property

Public Property Get MaxFormWidth() As Integer
Attribute MaxFormWidth.VB_Description = "Returns/sets the maximum resized width of the form (-1 to ignore)"
    MaxFormWidth = m_intMaxFormWidth
End Property

Public Property Let MaxFormWidth(ByVal vNewValue As Integer)
    m_intMaxFormWidth = vNewValue
End Property

Public Property Get MaxFormHeight() As Integer
Attribute MaxFormHeight.VB_Description = "Returns/sets the maximum resized height of the form (-1 to ignore)"
    MaxFormHeight = m_intMaxFormHeight
End Property

Public Property Let MaxFormHeight(ByVal vNewValue As Integer)
    m_intMaxFormHeight = vNewValue
End Property

Public Sub Resize()
Attribute Resize.VB_Description = "Performs resizing/anchoring operations for every control that has a style preset."
    If Not UserControl.Enabled Then Exit Sub
    
    Dim CurItem As clsControlInfo
    Dim blnRTL As Boolean
    
    Dim mForm As Object
    Set mForm = UserControl.Parent
    
    On Error Resume Next
    
    blnRTL = mForm.RightToLeft
        
    For Each CurItem In m_ControlTracker.ControlArray
        
        'Check if item exists in the form. (dynamic control may have been deleted)
        If CurItem.ControlRef Is Nothing Then
            m_ControlTracker.ControlArray.Remove CurItem.Key
            GoTo Continue
        End If
        If CurItem.Style <> AnchorResizeStyle.None Then
        
            If blnRTL And CheckBit(CurItem.Style, 16) Then
                If (CheckBit(CurItem.Style, 32 + 1) And Not CheckBit(CurItem.Style, 32 + 2)) Or _
                     (Not CheckBit(CurItem.Style, 32 + 1) And CheckBit(CurItem.Style, 32 + 2)) Then
                    CurItem.Style = CurItem.Style + (IIf(CheckBit(CurItem.Style, 32 + 1), 1, -1))
                End If
            End If
            
            If CheckBit(CurItem.Style, 1 + 2) Then
            
                If TypeOf CurItem.ControlRef Is Line Then
                    CurItem.ControlRef.X2 = CurItem.ControlRef.X1 + (CurItem.Width + CurItem.ControlRef.Container.Width - CurItem.ContainerWidth)
                Else
                    CurItem.ControlRef.Width = CurItem.Width + CurItem.ControlRef.Container.Width - CurItem.ContainerWidth
                End If
                
            ElseIf CheckBit(CurItem.Style, 2) Then
            
                If TypeOf CurItem.ControlRef Is Line Then
                    Dim oldLeft As Integer
                    oldLeft = CurItem.ControlRef.X1
                    CurItem.ControlRef.X1 = CurItem.Left + CurItem.ControlRef.Container.Width - CurItem.ContainerWidth
                    CurItem.ControlRef.X2 = CurItem.ControlRef.X2 + (CurItem.ControlRef.X1 - oldLeft)
                Else
                    CurItem.ControlRef.Left = CurItem.Left + CurItem.ControlRef.Container.Width - CurItem.ContainerWidth
                End If
                
            End If
            
            If CheckBit(CurItem.Style, 4 + 8) Then
            
                If TypeOf CurItem.ControlRef Is Line Then
                    CurItem.ControlRef.Y2 = CurItem.ControlRef.Y1 + (CurItem.Height + CurItem.ControlRef.Container.Height - CurItem.ContainerHeight)
                Else
                    CurItem.ControlRef.Height = CurItem.Height + CurItem.ControlRef.Container.Height - CurItem.ContainerHeight
                End If
                
            ElseIf CheckBit(CurItem.Style, 8) Then
            
                If TypeOf CurItem.ControlRef Is Line Then
                    Dim oldTop As Integer
                    oldTop = CurItem.ControlRef.Y1
                    CurItem.ControlRef.Y1 = CurItem.Top + CurItem.ControlRef.Container.Height - CurItem.ContainerHeight
                    CurItem.ControlRef.Y2 = CurItem.ControlRef.Y2 + (CurItem.ControlRef.Y1 - oldTop)
                Else
                    CurItem.ControlRef.Top = CurItem.Top + CurItem.ControlRef.Container.Height - CurItem.ContainerHeight
                End If
                
            End If
            
        End If
Continue:
        
    Next
    Err.Clear

End Sub


'======================================================================================================
'Subclass code - The programmer may call any of the following Subclass_??? routines

'Add a message to the table of those that will invoke a callback. You should Subclass_Start first and then add the messages
Private Sub Subclass_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
'Parameters:
  'lng_hWnd  - The handle of the window for which the uMsg is to be added to the callback table
  'uMsg      - The message number that will invoke a callback. NB Can also be ALL_MESSAGES, ie all messages will callback
  'When      - Whether the msg is to callback before, after or both with respect to the the default (previous) handler
  With sc_aSubData(zIdx(lng_hWnd))
    If When And eMsgWhen.MSG_BEFORE Then
      Call zAddMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
    End If
    If When And eMsgWhen.MSG_AFTER Then
      Call zAddMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
    End If
  End With
End Sub

'Delete a message from the table of those that will invoke a callback.
Private Sub Subclass_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
'Parameters:
  'lng_hWnd  - The handle of the window for which the uMsg is to be removed from the callback table
  'uMsg      - The message number that will be removed from the callback table. NB Can also be ALL_MESSAGES, ie all messages will callback
  'When      - Whether the msg is to be removed from the before, after or both callback tables
  With sc_aSubData(zIdx(lng_hWnd))
    If When And eMsgWhen.MSG_BEFORE Then
      Call zDelMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
    End If
    If When And eMsgWhen.MSG_AFTER Then
      Call zDelMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
    End If
  End With
End Sub

'Return whether we're running in the IDE.
Private Function Subclass_InIDE() As Boolean
  Debug.Assert zSetTrue(Subclass_InIDE)
End Function

'Start subclassing the passed window handle
Private Function Subclass_Start(ByVal lng_hWnd As Long) As Long
'Parameters:
  'lng_hWnd  - The handle of the window to be subclassed
'Returns;
  'The sc_aSubData() index
  Const CODE_LEN              As Long = 200                                             'Length of the machine code in bytes
  Const FUNC_CWP              As String = "CallWindowProcA"                             'We use CallWindowProc to call the original WndProc
  Const FUNC_EBM              As String = "EbMode"                                      'VBA's EbMode function allows the machine code thunk to know if the IDE has stopped or is on a breakpoint
  Const FUNC_SWL              As String = "SetWindowLongA"                              'SetWindowLongA allows the cSubclasser machine code thunk to unsubclass the subclasser itself if it detects via the EbMode function that the IDE has stopped
  Const MOD_USER              As String = "user32"                                      'Location of the SetWindowLongA & CallWindowProc functions
  Const MOD_VBA5              As String = "vba5"                                        'Location of the EbMode function if running VB5
  Const MOD_VBA6              As String = "vba6"                                        'Location of the EbMode function if running VB6
  Const PATCH_01              As Long = 18                                              'Code buffer offset to the location of the relative address to EbMode
  Const PATCH_02              As Long = 68                                              'Address of the previous WndProc
  Const PATCH_03              As Long = 78                                              'Relative address of SetWindowsLong
  Const PATCH_06              As Long = 116                                             'Address of the previous WndProc
  Const PATCH_07              As Long = 121                                             'Relative address of CallWindowProc
  Const PATCH_0A              As Long = 186                                             'Address of the owner object
  Static aBuf(1 To CODE_LEN)  As Byte                                                   'Static code buffer byte array
  Static pCWP                 As Long                                                   'Address of the CallWindowsProc
  Static pEbMode              As Long                                                   'Address of the EbMode IDE break/stop/running function
  Static pSWL                 As Long                                                   'Address of the SetWindowsLong function
  Dim i                       As Long                                                   'Loop index
  Dim j                       As Long                                                   'Loop index
  Dim nSubIdx                 As Long                                                   'Subclass data index
  Dim sHex                    As String                                                 'Hex code string
  
'If it's the first time through here..
  If aBuf(1) = 0 Then
  
'The hex pair machine code representation.
    sHex = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D00" & _
           "00005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D00" & _
           "0000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209" & _
           "C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90A4070000C3"

'Convert the string from hex pairs to bytes and store in the static machine code buffer
    i = 1
    Do While j < CODE_LEN
      j = j + 1
      aBuf(j) = Val("&H" & Mid$(sHex, i, 2))                                            'Convert a pair of hex characters to an eight-bit value and store in the static code buffer array
      i = i + 2
    Loop                                                                                'Next pair of hex characters
    
'Get API function addresses
    If Subclass_InIDE Then                                                              'If we're running in the VB IDE
      aBuf(16) = &H90                                                                   'Patch the code buffer to enable the IDE state code
      aBuf(17) = &H90                                                                   'Patch the code buffer to enable the IDE state code
      pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)                                           'Get the address of EbMode in vba6.dll
      If pEbMode = 0 Then                                                               'Found?
        pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)                                         'VB5 perhaps
      End If
    End If
    
    pCWP = zAddrFunc(MOD_USER, FUNC_CWP)                                                'Get the address of the CallWindowsProc function
    pSWL = zAddrFunc(MOD_USER, FUNC_SWL)                                                'Get the address of the SetWindowLongA function
    ReDim sc_aSubData(0 To 0) As tSubData                                               'Create the first sc_aSubData element
  Else
    nSubIdx = zIdx(lng_hWnd, True)
    If nSubIdx = -1 Then                                                                'If an sc_aSubData element isn't being re-cycled
      nSubIdx = UBound(sc_aSubData()) + 1                                               'Calculate the next element
      ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData                              'Create a new sc_aSubData element
    End If
    
    Subclass_Start = nSubIdx
  End If

  With sc_aSubData(nSubIdx)
    .hWnd = lng_hWnd                                                                    'Store the hWnd
    .nAddrSub = GlobalAlloc(GMEM_FIXED, CODE_LEN)                                       'Allocate memory for the machine code WndProc
    .nAddrOrig = SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrSub)                          'Set our WndProc in place
    Call RtlMoveMemory(ByVal .nAddrSub, aBuf(1), CODE_LEN)                              'Copy the machine code from the static byte array to the code array in sc_aSubData
    Call zPatchRel(.nAddrSub, PATCH_01, pEbMode)                                        'Patch the relative address to the VBA EbMode api function, whether we need to not.. hardly worth testing
    Call zPatchVal(.nAddrSub, PATCH_02, .nAddrOrig)                                     'Original WndProc address for CallWindowProc, call the original WndProc
    Call zPatchRel(.nAddrSub, PATCH_03, pSWL)                                           'Patch the relative address of the SetWindowLongA api function
    Call zPatchVal(.nAddrSub, PATCH_06, .nAddrOrig)                                     'Original WndProc address for SetWindowLongA, unsubclass on IDE stop
    Call zPatchRel(.nAddrSub, PATCH_07, pCWP)                                           'Patch the relative address of the CallWindowProc api function
    Call zPatchVal(.nAddrSub, PATCH_0A, ObjPtr(Me))                                     'Patch the address of this object instance into the static machine code buffer
  End With
End Function

'Stop all subclassing
Private Sub Subclass_StopAll()
  Dim i As Long
  
  i = UBound(sc_aSubData())                                                             'Get the upper bound of the subclass data array
  Do While i >= 0                                                                       'Iterate through each element
    With sc_aSubData(i)
      If .hWnd <> 0 Then                                                                'If not previously Subclass_Stop'd
        Call Subclass_Stop(.hWnd)                                                       'Subclass_Stop
      End If
    End With
    
    i = i - 1                                                                           'Next element
  Loop
End Sub

'Stop subclassing the passed window handle
Private Sub Subclass_Stop(ByVal lng_hWnd As Long)
'Parameters:
  'lng_hWnd  - The handle of the window to stop being subclassed
  With sc_aSubData(zIdx(lng_hWnd))
    Call SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrOrig)                                 'Restore the original WndProc
    Call zPatchVal(.nAddrSub, PATCH_05, 0)                                              'Patch the Table B entry count to ensure no further 'before' callbacks
    Call zPatchVal(.nAddrSub, PATCH_09, 0)                                              'Patch the Table A entry count to ensure no further 'after' callbacks
    Call GlobalFree(.nAddrSub)                                                          'Release the machine code memory
    .hWnd = 0                                                                           'Mark the sc_aSubData element as available for re-use
    .nMsgCntB = 0                                                                       'Clear the before table
    .nMsgCntA = 0                                                                       'Clear the after table
    Erase .aMsgTblB                                                                     'Erase the before table
    Erase .aMsgTblA                                                                     'Erase the after table
  End With
End Sub

'======================================================================================================
'These z??? routines are exclusively called by the Subclass_??? routines.

'Worker sub for Subclass_AddMsg
Private Sub zAddMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
  Dim nEntry  As Long                                                                   'Message table entry index
  Dim nOff1   As Long                                                                   'Machine code buffer offset 1
  Dim nOff2   As Long                                                                   'Machine code buffer offset 2
  
  If uMsg = ALL_MESSAGES Then                                                           'If all messages
    nMsgCnt = ALL_MESSAGES                                                              'Indicates that all messages will callback
  Else                                                                                  'Else a specific message number
    Do While nEntry < nMsgCnt                                                           'For each existing entry. NB will skip if nMsgCnt = 0
      nEntry = nEntry + 1
      
      If aMsgTbl(nEntry) = 0 Then                                                       'This msg table slot is a deleted entry
        aMsgTbl(nEntry) = uMsg                                                          'Re-use this entry
        Exit Sub                                                                        'Bail
      ElseIf aMsgTbl(nEntry) = uMsg Then                                                'The msg is already in the table!
        Exit Sub                                                                        'Bail
      End If
    Loop                                                                                'Next entry

    nMsgCnt = nMsgCnt + 1                                                               'New slot required, bump the table entry count
    ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long                                        'Bump the size of the table.
    aMsgTbl(nMsgCnt) = uMsg                                                             'Store the message number in the table
  End If

  If When = eMsgWhen.MSG_BEFORE Then                                                    'If before
    nOff1 = PATCH_04                                                                    'Offset to the Before table
    nOff2 = PATCH_05                                                                    'Offset to the Before table entry count
  Else                                                                                  'Else after
    nOff1 = PATCH_08                                                                    'Offset to the After table
    nOff2 = PATCH_09                                                                    'Offset to the After table entry count
  End If

  If uMsg <> ALL_MESSAGES Then
    Call zPatchVal(nAddr, nOff1, VarPtr(aMsgTbl(1)))                                    'Address of the msg table, has to be re-patched because Redim Preserve will move it in memory.
  End If
  Call zPatchVal(nAddr, nOff2, nMsgCnt)                                                 'Patch the appropriate table entry count
End Sub

'Return the memory address of the passed function in the passed dll
Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
  zAddrFunc = GetProcAddress(GetModuleHandleA(sDLL), sProc)
  Debug.Assert zAddrFunc                                                                'You may wish to comment out this line if you're using vb5 else the EbMode GetProcAddress will stop here everytime because we look for vba6.dll first
End Function

'Worker sub for Subclass_DelMsg
Private Sub zDelMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
  Dim nEntry As Long
  
  If uMsg = ALL_MESSAGES Then                                                           'If deleting all messages
    nMsgCnt = 0                                                                         'Message count is now zero
    If When = eMsgWhen.MSG_BEFORE Then                                                  'If before
      nEntry = PATCH_05                                                                 'Patch the before table message count location
    Else                                                                                'Else after
      nEntry = PATCH_09                                                                 'Patch the after table message count location
    End If
    Call zPatchVal(nAddr, nEntry, 0)                                                    'Patch the table message count to zero
  Else                                                                                  'Else deleteting a specific message
    Do While nEntry < nMsgCnt                                                           'For each table entry
      nEntry = nEntry + 1
      If aMsgTbl(nEntry) = uMsg Then                                                    'If this entry is the message we wish to delete
        aMsgTbl(nEntry) = 0                                                             'Mark the table slot as available
        Exit Do                                                                         'Bail
      End If
    Loop                                                                                'Next entry
  End If
End Sub

'Get the sc_aSubData() array index of the passed hWnd
Private Function zIdx(ByVal lng_hWnd As Long, Optional ByVal bAdd As Boolean = False) As Long
'Get the upper bound of sc_aSubData() - If you get an error here, you're probably Subclass_AddMsg-ing before Subclass_Start
  zIdx = UBound(sc_aSubData)
  Do While zIdx >= 0                                                                    'Iterate through the existing sc_aSubData() elements
    With sc_aSubData(zIdx)
      If .hWnd = lng_hWnd Then                                                          'If the hWnd of this element is the one we're looking for
        If Not bAdd Then                                                                'If we're searching not adding
          Exit Function                                                                 'Found
        End If
      ElseIf .hWnd = 0 Then                                                             'If this an element marked for reuse.
        If bAdd Then                                                                    'If we're adding
          Exit Function                                                                 'Re-use it
        End If
      End If
    End With
    zIdx = zIdx - 1                                                                     'Decrement the index
  Loop
  
  If Not bAdd Then
    Debug.Assert False                                                                  'hWnd not found, programmer error
  End If

'If we exit here, we're returning -1, no freed elements were found
End Function

'Patch the machine code buffer at the indicated offset with the relative address to the target address.
Private Sub zPatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)
  Call RtlMoveMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)
End Sub

'Patch the machine code buffer at the indicated offset with the passed value
Private Sub zPatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)
  Call RtlMoveMemory(ByVal nAddr + nOffset, nValue, 4)
End Sub

'Worker function for Subclass_InIDE
Private Function zSetTrue(ByRef bValue As Boolean) As Boolean
  zSetTrue = True
  bValue = True
End Function

Private Sub UserControl_Terminate()
    Set m_ControlTracker = Nothing
    
    On Error GoTo Catch
    
    'Stop all subclassing
    Call Subclass_StopAll
    
Catch:

End Sub

