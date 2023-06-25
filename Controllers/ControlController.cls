Option Explicit
'@Folder(Database7.Controllers)

'ControlController.cls

'Child of a ControlManager object.
'Holds ParentManager
'Attaches to ChildControl

'Raises all events to parent manager
'Passes its reference

'------Testing Shit-----'
Public Event Change(Index As Integer)


'------[ API definition for Enter/Exit/BeforeUpdate/AfterUpdate ]-------------
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

#If VBA7 And Win64 Then
    Private Declare PtrSafe Function ConnectToConnectionPoint Lib "shlwapi" Alias "#168" _
    (ByVal punk As stdole.IUnknown, _
    ByRef riidEvent As GUID, _
    ByVal fConnect As Long, _
    ByVal punkTarget As stdole.IUnknown, _
    ByRef pdwCookie As Long, _
    Optional ByVal ppcpOut As LongPtr) As Long
#Else
    Private Declare Function ConnectToConnectionPoint Lib "shlwapi" Alias "#168" _
                             (ByVal punk As stdole.IUnknown, _
                              ByRef riidEvent As GUID, _
                              ByVal fConnect As Long, _
                              ByVal punkTarget As stdole.IUnknown, _
                              ByRef pdwCookie As Long, _
                              Optional ByVal ppcpOut As Long) As Long
#End If

Private Cookie As Long                           'pdwCookie of ConnectToConnectionPoint
Private MyCtrl As Object


Private MyIndex As Integer                       'Index of the control
Private MyParent As ControllerManager            'Reference to clsBpca object
Private MyEnableEvents As Boolean
Private MyFontColor As Long                      'Initial Foreground Color
Private MyBackColor As Long                      'Initial Background Color
'@Ignore HungarianNotation
Private blnParent As Boolean


'------[ Objects of Event-Listener ]----------------------------
'@MemberAttribute VB_VarHelpID, -1
Private WithEvents MyCtrlLbl As Label
'@MemberAttribute VB_VarHelpID, -1
Private WithEvents MyCtrlCmd As CommandButton
'@MemberAttribute VB_VarHelpID, -1
Private WithEvents MyCtrlTxt As TextBox
'@MemberAttribute VB_VarHelpID, -1
Private WithEvents MyCtrlChk As CheckBox
'@MemberAttribute VB_VarHelpID, -1
Private WithEvents MyCtrlOpt As OptionButton
'@MemberAttribute VB_VarHelpID, -1
Private WithEvents MyCtrlTgl As ToggleButton
'@MemberAttribute VB_VarHelpID, -1
Private WithEvents MyCtrlCbo As ComboBox

'In the child class, Frame/MultiPage are added for [FakeExit].
'@MemberAttribute VB_VarHelpID, -1
Private WithEvents MyCtrlFrm As Form
'@MemberAttribute VB_VarHelpID, -1
Private WithEvents MyCtrlSubFrm As SubForm

'TODO: figure this out one day.
'Private MyCtrlMlt_OldPage As Integer
'Previous page number of Page_Change event

'---------------------------------------------------
'Raise Event(Enable or Not enable)
'@Ignore HungarianNotation
Private lngEventState As Long

'@Ignore HungarianNotation
Private blnEventChange As Boolean                '[ 1bit]
'@Ignore HungarianNotation
Private blnEventClick As Boolean                 '[ 2]
'@Ignore HungarianNotation
Private blnEventDblClick As Boolean              '[ 3]
'@Ignore HungarianNotation
Private blnEventKeyDown As Boolean               '[ 4]
'@Ignore HungarianNotation
Private blnEventKeyPress As Boolean              '[ 5]
'@Ignore HungarianNotation
Private blnEventKeyUp As Boolean                 '[ 6]
'@Ignore HungarianNotation
Private blnEventMouseMove As Boolean             '[ 7]
'@Ignore HungarianNotation
Private blnEventMouseDown As Boolean             '[ 8]
'@Ignore HungarianNotation
Private blnEventMouseUp As Boolean               '[ 9]
'@Ignore HungarianNotation
Private blnEventDropBtnClick As Boolean          '[10]
'@Ignore HungarianNotation
Private blnEventSpinDown As Boolean              '[11]
'@Ignore HungarianNotation
Private blnEventSpinUp As Boolean                '[12]
'@Ignore HungarianNotation
Private blnEventEnter As Boolean                 '[13]
'@Ignore HungarianNotation
Private blnEventExit As Boolean                  '[14]
'@Ignore HungarianNotation
Private blnEventBeforeUpdate As Boolean          '[15]
'@Ignore HungarianNotation
Private blnEventAfterUpdate As Boolean           '[16]
'@Ignore HungarianNotation
Private blnEventFakeExit As Boolean              '[17]

'=============================================
Private Sub Class_Initialize()
    blnParent = False
    MyIndex = 0
    MyEnableEvents = True
End Sub

'=====[ Event connection by ConnectToConnectionPoint ]=============
Private Sub ConnectEvent(ByVal Connect As Boolean)
    Dim IID_IDispatch As GUID

    ' GUID {00020400-0000-0000-C000000000000046}
    With IID_IDispatch
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With
    
    '@Ignore ObsoleteCallStatement
    Call ConnectToConnectionPoint _
         (Me, IID_IDispatch, Connect, MyCtrl, Cookie, 0&)
End Sub

'=============================================
Public Sub Clear()
    If (Cookie <> 0) Then
        'Destruction of the event connection
        'with MyCtrl by ConnectToConnectionPoint
        ConnectEvent False
        Cookie = 0
    End If
    
    Set MyCtrlLbl = Nothing
    Set MyCtrlCmd = Nothing
    Set MyCtrlTxt = Nothing
    Set MyCtrlChk = Nothing
    Set MyCtrlOpt = Nothing
    Set MyCtrlTgl = Nothing
    Set MyCtrlCbo = Nothing

    Set MyCtrl = Nothing
    Set MyParent = Nothing
    blnParent = False
    MyIndex = 0
End Sub

'=============================================
Public Property Let Item(NewCtrl As Object)
    Set MyCtrl = NewCtrl
    MyFontColor = MyCtrl.ForeColor               'Store the initial Foreground Color
    MyBackColor = MyCtrl.BackColor               'Store the initial Background Color
    
    '[EventSet] has been carried out earlier from [Item].
    If ((lngEventState And (BPCA2_EnterExit + BPCA2_BAUpdate)) = 0) Then
        'There are no instructions of [Enter,Exit,BeforeUpdate,AfterUpdate].
        Debug.Print "what the hell is this"
    Else
        'Event connection with MyCtrl by ConnectToConnectionPoint
        ConnectEvent True
    End If
    
    Select Case TypeName(NewCtrl)
    Case "Label": Set MyCtrlLbl = NewCtrl
    Case "CommandButton": Set MyCtrlCmd = NewCtrl
    Case "TextBox": Set MyCtrlTxt = NewCtrl
    Case "CheckBox": Set MyCtrlChk = NewCtrl
    Case "OptionButton": Set MyCtrlOpt = NewCtrl
    Case "ToggleButton": Set MyCtrlTgl = NewCtrl
    Case "ComboBox": Set MyCtrlCbo = NewCtrl
        'Todo look up reference code to figure this out to get oldpage to work
        'MyCtrlMlt_OldPage = MyCtrlMlt.Value
    End Select
End Property

Public Property Get Item() As Object
    Set Item = MyCtrl
End Property

'=============================================
Public Property Let Index(NewIndex As Integer)
    MyIndex = NewIndex
End Property

Public Property Get Index() As Integer
    Index = MyIndex
End Property

'=============================================
Public Property Let EnableEvents(NewEnableEvents As Boolean)
    MyEnableEvents = NewEnableEvents
End Property

Public Property Get EnableEvents() As Boolean
    EnableEvents = MyEnableEvents
End Property

'=============================================
Public Property Get InitFColor() As Long
    InitFColor = MyFontColor
End Property

'=============================================
Public Property Get InitBColor() As Long
    InitBColor = MyBackColor
End Property

'=============================================
Public Property Let Parent(NewParent As ControllerManager)
    blnParent = True
    Set MyParent = NewParent
End Property

Public Property Get Parent() As ControllerManager
    Set Parent = MyParent
End Property

'=============================================
Public Property Let EventSet(NewEvent As Long)
    'Break numerical value into a bit and substitute it for each flag

    lngEventState = NewEvent
    
    blnEventChange = ((NewEvent And BPCA2_Change) = BPCA2_Change)
    blnEventClick = ((NewEvent And BPCA2_Click) = BPCA2_Click)
    blnEventDblClick = ((NewEvent And BPCA2_DblClick) = BPCA2_DblClick)
    blnEventKeyDown = ((NewEvent And BPCA2_KeyDown) = BPCA2_KeyDown)
    blnEventKeyPress = ((NewEvent And BPCA2_KeyPress) = BPCA2_KeyPress)
    blnEventKeyUp = ((NewEvent And BPCA2_KeyUp) = BPCA2_KeyUp)
    blnEventMouseMove = ((NewEvent And BPCA2_MouseMove) = BPCA2_MouseMove)
    blnEventMouseDown = ((NewEvent And BPCA2_MouseDown) = BPCA2_MouseDown)
    blnEventMouseUp = ((NewEvent And BPCA2_MouseUp) = BPCA2_MouseUp)
    blnEventDropBtnClick = ((NewEvent And BPCA2_DropBtnClick) = BPCA2_DropBtnClick)
    blnEventSpinDown = ((NewEvent And BPCA2_SpinDown) = BPCA2_SpinDown)
    blnEventSpinUp = ((NewEvent And BPCA2_SpinUp) = BPCA2_SpinUp)
    
    blnEventEnter = ((NewEvent And BPCA2_Enter) = BPCA2_Enter)
    blnEventExit = ((NewEvent And BPCA2_Exit) = BPCA2_Exit)
    blnEventBeforeUpdate = ((NewEvent And BPCA2_BeforeUpdate) = BPCA2_BeforeUpdate)
    blnEventAfterUpdate = ((NewEvent And BPCA2_AfterUpdate) = BPCA2_AfterUpdate)
    
    blnEventFakeExit = ((NewEvent And BPCA2_FakeExit) = BPCA2_FakeExit)
End Property

'=============================================
' I return Variant array for the event raise classification
' that I received in EventSet property.
'  0: Long type, 1-16: Boolean type
'
'  [0] The numerical value that I received in EventSet
'  [1] Change, [2]Click, [3]DblClick
'  [4] KeyDown, [5]KeyPress, [6]KeyUp
'  [7] MouseMove, [8]MouseDown, [9]MouseUp
'  [10] DropBtnClick), [11]SpinDown, [12]SpinUp
'  [13] Enter, [14] Exit, [15] BeforeUpdate, [16] AfterUpdate
'  [17] FakeExit

Public Property Get EventState() As Variant
    EventState = Array(lngEventState, _
                       blnEventChange, blnEventClick, blnEventDblClick, _
                       blnEventKeyDown, blnEventKeyPress, blnEventKeyUp, _
                       blnEventMouseMove, blnEventMouseDown, blnEventMouseUp, _
                       blnEventDropBtnClick, blnEventSpinDown, blnEventSpinUp, _
                       blnEventEnter, blnEventExit, blnEventBeforeUpdate, blnEventAfterUpdate, _
                       blnEventFakeExit)
End Property

'@MemberAttribute VB_UserMemId, -2147384830
Public Sub HookEnter()
    'Attribute HookEnter.VB_UserMemId = -2147384830
    If (blnParent And blnEventEnter And MyEnableEvents) Then
        MyParent.RaiseEnter MyIndex
    End If
End Sub

'@MemberAttribute VB_UserMemId, -2147384829
Public Sub HookExit(ByVal Cancel As Boolean)
    'Attribute HookExit.VB_UserMemId = -2147384829
    If (blnParent And blnEventExit And MyEnableEvents) Then
        Select Case TypeName(MyCtrl)
        Case "Frame", "MultiPage"
            '@TODO: fix this
            'TODO: fix this someday
            MyParent.Exit4FakeExit MyIndex
        Case Else
            MyParent.RaiseExit MyIndex, Cancel
        End Select
    End If
End Sub

'@MemberAttribute VB_UserMemId, -2147384831
Public Sub HookBeforeUpdate(Cancel As Boolean)
    'Attribute HookBeforeUpdate.VB_UserMemId = -2147384831
    If (blnParent And blnEventBeforeUpdate And MyEnableEvents) Then
        MyParent.RaiseBeforeUpdate MyIndex, Cancel
    End If
End Sub

'@MemberAttribute VB_UserMemId, -2147384832
Public Sub HookAfterUpdate()
    'Attribute HookAfterUpdate.VB_UserMemId = -2147384832
    If (blnParent And blnEventAfterUpdate And MyEnableEvents) Then
        MyParent.RaiseAfterUpdate MyIndex
    End If
End Sub

Private Sub MyCtrlCbo_DblClick(Cancel As Integer)
    If (blnParent And blnEventDblClick And MyEnableEvents) Then
        MyParent.RaiseDblClick MyIndex, Cancel
    End If
End Sub

Private Sub MyCtrlChk_DblClick(Cancel As Integer)
    If (blnParent And blnEventDblClick And MyEnableEvents) Then
        MyParent.RaiseDblClick MyIndex, Cancel
    End If
End Sub

Private Sub MyCtrlCmd_DblClick(Cancel As Integer)
    If (blnParent And blnEventDblClick And MyEnableEvents) Then
        MyParent.RaiseDblClick MyIndex, Cancel
    End If
End Sub

Private Sub MyCtrlLbl_Click()
    If (blnParent And blnEventClick And MyEnableEvents) Then
        MyParent.RaiseClick MyIndex
    End If
End Sub

Private Sub MyCtrlLbl_DblClick(Cancel As Integer)
    If (blnParent And blnEventDblClick And MyEnableEvents) Then
        MyParent.RaiseDblClick MyIndex, Cancel
    End If
End Sub

'----[ Label ]--------------------------------------




Private Sub MyCtrlLbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (blnParent And blnEventMouseMove And MyEnableEvents) Then
        MyParent.RaiseMouseMove MyIndex, Button, Shift, X, Y
    End If
End Sub

Private Sub MyCtrlLbl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (blnParent And blnEventMouseDown And MyEnableEvents) Then
        MyParent.RaiseMouseDown MyIndex, Button, Shift, X, Y
    End If
End Sub

Private Sub MyCtrlLbl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (blnParent And blnEventMouseUp And MyEnableEvents) Then
        MyParent.RaiseMouseUp MyIndex, Button, Shift, X, Y
    End If
End Sub

'----[ CommandButton ]--------------------------------------
Private Sub MyCtrlCmd_Click()
    If (blnParent And blnEventClick And MyEnableEvents) Then
        MyParent.RaiseClick MyIndex
    End If
End Sub

Private Sub MyCtrlCmd_KeyDown(KeyCode As Integer, Shift As Integer)
    If (blnParent And blnEventKeyDown And MyEnableEvents) Then
        MyParent.RaiseKeyDown MyIndex, KeyCode, Shift
    End If
End Sub

Private Sub MyCtrlCmd_KeyPress(KeyAscii As Integer)
    If (blnParent And blnEventKeyPress And MyEnableEvents) Then
        MyParent.RaiseKeyPress MyIndex, KeyAscii
    End If
End Sub

Private Sub MyCtrlCmd_KeyUp(KeyCode As Integer, Shift As Integer)
    If (blnParent And blnEventKeyUp And MyEnableEvents) Then
        MyParent.RaiseKeyUp MyIndex, KeyCode, Shift
    End If
End Sub

Private Sub MyCtrlCmd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (blnParent And blnEventMouseMove And MyEnableEvents) Then
        MyParent.RaiseMouseMove MyIndex, Button, Shift, X, Y
    End If
End Sub

Private Sub MyCtrlCmd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (blnParent And blnEventMouseDown And MyEnableEvents) Then
        MyParent.RaiseMouseDown MyIndex, Button, Shift, X, Y
    End If
End Sub

Private Sub MyCtrlCmd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (blnParent And blnEventMouseUp And MyEnableEvents) Then
        MyParent.RaiseMouseUp MyIndex, Button, Shift, X, Y
    End If
End Sub

Private Sub MyCtrlOpt_DblClick(Cancel As Integer)
    If (blnParent And blnEventDblClick And MyEnableEvents) Then
        MyParent.RaiseDblClick MyIndex, Cancel
    End If
End Sub

Private Sub MyCtrlTgl_DblClick(Cancel As Integer)
    If (blnParent And blnEventDblClick And MyEnableEvents) Then
        MyParent.RaiseDblClick MyIndex, Cancel
    End If
End Sub

Private Sub MyCtrlTxt_Change()
    'Todo: Reset this after finished testing it.
    'If (blnParent And blnEventChange And MyEnableEvents) Then
        'MyParent.RaiseChange MyIndex
        Debug.Print "Change event fired from control controller."
        RaiseEvent Change(0)
    'End If
End Sub

'----[ TextBox ]--------------------------------------



Private Sub MyCtrlTxt_DblClick(Cancel As Integer)
    If (blnParent And blnEventDblClick And MyEnableEvents) Then
        MyParent.RaiseDblClick MyIndex, Cancel
    End If
End Sub

Private Sub MyCtrlTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    If (blnParent And blnEventKeyDown And MyEnableEvents) Then
        MyParent.RaiseKeyDown MyIndex, KeyCode, Shift
    End If
End Sub

Private Sub MyCtrlTxt_KeyPress(KeyAscii As Integer)
    If (blnParent And blnEventKeyPress And MyEnableEvents) Then
        MyParent.RaiseKeyPress MyIndex, KeyAscii
    End If
End Sub

Private Sub MyCtrlTxt_KeyUp(KeyCode As Integer, Shift As Integer)
    If (blnParent And blnEventKeyUp And MyEnableEvents) Then
        MyParent.RaiseKeyUp MyIndex, KeyCode, Shift
    End If
End Sub

Private Sub MyCtrlTxt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (blnParent And blnEventMouseMove And MyEnableEvents) Then
        MyParent.RaiseMouseMove MyIndex, Button, Shift, X, Y
    End If
End Sub

Private Sub MyCtrlTxt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (blnParent And blnEventMouseDown And MyEnableEvents) Then
        MyParent.RaiseMouseDown MyIndex, Button, Shift, X, Y
    End If
End Sub

Private Sub MyCtrlTxt_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (blnParent And blnEventMouseUp And MyEnableEvents) Then
        MyParent.RaiseMouseUp MyIndex, Button, Shift, X, Y
    End If
End Sub

'----[ CheckBox ]--------------------------------------
Private Sub MyCtrlChk_Change()
    If (blnParent And blnEventChange And MyEnableEvents) Then
        MyParent.RaiseChange MyIndex
    End If
End Sub

Private Sub MyCtrlChk_Click()
    If (blnParent And blnEventClick And MyEnableEvents) Then
        MyParent.RaiseClick MyIndex
    End If
End Sub

Private Sub MyCtrlChk_KeyDown(KeyCode As Integer, Shift As Integer)
    If (blnParent And blnEventKeyDown And MyEnableEvents) Then
        MyParent.RaiseKeyDown MyIndex, KeyCode, Shift
    End If
End Sub

Private Sub MyCtrlChk_KeyPress(KeyAscii As Integer)
    If (blnParent And blnEventKeyPress And MyEnableEvents) Then
        MyParent.RaiseKeyPress MyIndex, KeyAscii
    End If
End Sub

Private Sub MyCtrlChk_KeyUp(KeyCode As Integer, Shift As Integer)
    If (blnParent And blnEventKeyUp And MyEnableEvents) Then
        MyParent.RaiseKeyUp MyIndex, KeyCode, Shift
    End If
End Sub

Private Sub MyCtrlChk_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (blnParent And blnEventMouseMove And MyEnableEvents) Then
        MyParent.RaiseMouseMove MyIndex, Button, Shift, X, Y
    End If
End Sub

Private Sub MyCtrlChk_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (blnParent And blnEventMouseDown And MyEnableEvents) Then
        MyParent.RaiseMouseDown MyIndex, Button, Shift, X, Y
    End If
End Sub

Private Sub MyCtrlChk_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (blnParent And blnEventMouseUp And MyEnableEvents) Then
        MyParent.RaiseMouseUp MyIndex, Button, Shift, X, Y
    End If
End Sub

'----[ OptionButtton ]--------------------------------------
Private Sub MyCtrlOpt_Change()
    If (blnParent And blnEventChange And MyEnableEvents) Then
        MyParent.RaiseChange MyIndex
    End If
End Sub

Private Sub MyCtrlOpt_Click()
    If (blnParent And blnEventClick And MyEnableEvents) Then
        MyParent.RaiseClick MyIndex
    End If
End Sub

Private Sub MyCtrlOpt_KeyDown(KeyCode As Integer, Shift As Integer)
    If (blnParent And blnEventKeyDown And MyEnableEvents) Then
        MyParent.RaiseKeyDown MyIndex, KeyCode, Shift
    End If
End Sub

Private Sub MyCtrlOpt_KeyPress(KeyAscii As Integer)
    If (blnParent And blnEventKeyPress And MyEnableEvents) Then
        MyParent.RaiseKeyPress MyIndex, KeyAscii
    End If
End Sub

Private Sub MyCtrlOpt_KeyUp(KeyCode As Integer, Shift As Integer)
    If (blnParent And blnEventKeyUp And MyEnableEvents) Then
        MyParent.RaiseKeyUp MyIndex, KeyCode, Shift
    End If
End Sub

Private Sub MyCtrlOpt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (blnParent And blnEventMouseMove And MyEnableEvents) Then
        MyParent.RaiseMouseMove MyIndex, Button, Shift, X, Y
    End If
End Sub

Private Sub MyCtrlOpt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (blnParent And blnEventMouseDown And MyEnableEvents) Then
        MyParent.RaiseMouseDown MyIndex, Button, Shift, X, Y
    End If
End Sub

Private Sub MyCtrlOpt_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (blnParent And blnEventMouseUp And MyEnableEvents) Then
        MyParent.RaiseMouseUp MyIndex, Button, Shift, X, Y
    End If
End Sub

'----[ ToggleButtton ]--------------------------------------
Private Sub MyCtrlTgl_Change()
    If (blnParent And blnEventChange And MyEnableEvents) Then
        MyParent.RaiseChange MyIndex
    End If
End Sub

Private Sub MyCtrlTgl_Click()
    If (blnParent And blnEventClick And MyEnableEvents) Then
        MyParent.RaiseClick MyIndex
    End If
End Sub

Private Sub MyCtrlTgl_KeyDown(KeyCode As Integer, Shift As Integer)
    If (blnParent And blnEventKeyDown And MyEnableEvents) Then
        MyParent.RaiseKeyDown MyIndex, KeyCode, Shift
    End If
End Sub

Private Sub MyCtrlTgl_KeyPress(KeyAscii As Integer)
    If (blnParent And blnEventKeyPress And MyEnableEvents) Then
        MyParent.RaiseKeyPress MyIndex, KeyAscii
    End If
End Sub

Private Sub MyCtrlTgl_KeyUp(KeyCode As Integer, Shift As Integer)
    If (blnParent And blnEventKeyUp And MyEnableEvents) Then
        MyParent.RaiseKeyUp MyIndex, KeyCode, Shift
    End If
End Sub

Private Sub MyCtrlOTgl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (blnParent And blnEventMouseMove And MyEnableEvents) Then
        MyParent.RaiseMouseMove MyIndex, Button, Shift, X, Y
    End If
End Sub

Private Sub MyCtrlTgl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (blnParent And blnEventMouseDown And MyEnableEvents) Then
        MyParent.RaiseMouseDown MyIndex, Button, Shift, X, Y
    End If
End Sub

Private Sub MyCtrlTgl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (blnParent And blnEventMouseUp And MyEnableEvents) Then
        MyParent.RaiseMouseUp MyIndex, Button, Shift, X, Y
    End If
End Sub

'----[ ComboBox ]--------------------------------------
Private Sub MyCtrlCbo_Change()
    If (blnParent And blnEventChange And MyEnableEvents) Then
        MyParent.RaiseChange MyIndex
    End If
End Sub

Private Sub MyCtrlCbo_Click()
    If (blnParent And blnEventClick And MyEnableEvents) Then
        MyParent.RaiseClick MyIndex
    End If
End Sub

Private Sub MyCtrlCbo_KeyDown(KeyCode As Integer, Shift As Integer)
    If (blnParent And blnEventKeyDown And MyEnableEvents) Then
        MyParent.RaiseKeyDown MyIndex, KeyCode, Shift
    End If
End Sub

Private Sub MyCtrlCbo_KeyPress(KeyAscii As Integer)
    If (blnParent And blnEventKeyPress And MyEnableEvents) Then
        MyParent.RaiseKeyPress MyIndex, KeyAscii
    End If
End Sub

Private Sub MyCtrlCbo_KeyUp(KeyCode As Integer, Shift As Integer)
    If (blnParent And blnEventKeyUp And MyEnableEvents) Then
        MyParent.RaiseKeyUp MyIndex, KeyCode, Shift
    End If
End Sub

Private Sub MyCtrlCbo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (blnParent And blnEventMouseMove And MyEnableEvents) Then
        MyParent.RaiseMouseMove MyIndex, Button, Shift, X, Y
    End If
End Sub

Private Sub MyCtrlCbo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (blnParent And blnEventMouseDown And MyEnableEvents) Then
        MyParent.RaiseMouseDown MyIndex, Button, Shift, X, Y
    End If
End Sub

Private Sub MyCtrlCbo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (blnParent And blnEventMouseUp And MyEnableEvents) Then
        MyParent.RaiseMouseUp MyIndex, Button, Shift, X, Y
    End If
End Sub

'Private Sub MyCtrlCbo_DropButtonClick()
'    If (blnParent And blnEventDropBtnClick And MyEnableEvents) Then
'        MyParent.RaiseDropButtonClick MyIndex
'    End If
'End Sub

'----[ SpinButton ]--------------------------------------
' Private Sub MyCtrlSpn_Change()
'     If (blnParent And blnEventChange And MyEnableEvents) Then
'         MyParent.RaiseChange MyIndex
'     End If
' End Sub

' Private Sub MyCtrlSpn_KeyDown(KeyCode As Integer, Shift As Integer)
'     If (blnParent And blnEventKeyDown And MyEnableEvents) Then
'         MyParent.RaiseKeyDown MyIndex, KeyCode, Shift
'     End If
' End Sub

' Private Sub MyCtrlSpn_KeyPress(KeyAscii As Integer)
'     If (blnParent And blnEventKeyPress And MyEnableEvents) Then
'         MyParent.RaiseKeyPress MyIndex, KeyAscii
'     End If
' End Sub

' Private Sub MyCtrlSpn_KeyUp(KeyCode As Integer, Shift As Integer)
'     If (blnParent And blnEventKeyUp And MyEnableEvents) Then
'         MyParent.RaiseKeyUp MyIndex, KeyCode, Shift
'     End If
' End Sub

' Private Sub MyCtrlSpn_SpinDown()
'     If (blnParent And blnEventSpinDown And MyEnableEvents) Then
'         MyParent.RaiseSpinDown MyIndex
'     End If
' End Sub

' Private Sub MyCtrlSpn_SpinUp()
'     If (blnParent And blnEventSpinUp And MyEnableEvents) Then
'         MyParent.RaiseSpinUp MyIndex
'     End If
' End Sub

'----[ MultiPage ]--------------------------------------
'Private Sub MyCtrlMlt_Change()
'    If (blnParent And blnEventChange And MyEnableEvents) Then
'        MyParent.Change4FakeExit MyIndex, MyCtrlMlt_OldPage
'    End If
'
'    MyCtrlMlt_OldPage = MyCtrlMlt.Value
'End Sub
