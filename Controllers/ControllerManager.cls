Option Compare Database
Option Explicit
'@Folder(Database7.Controllers)
'ControlManager

'----Specification Notes-----
    'ControlManager objects are used to organize groups of controls.
    'Possibly connected to a single control to expose functionality

    'Add/remove/access ControlControllers
    'Add/remove/access Groups of Controllers ()

    'Need to define Control sourced events,
    'and child Manager sourced events.

    'Manager level events must pass to the next level (FormManager),
    '   Or pass reraised as a control level event along to the Manager,
    '   passing the ControlController along if the manager's in a group.
    '
    '   Do this in the Controller's or Manager's code automatically if the config
    '   routes it to a specific level.
    '   (Like MouseOver events talking to the Mainform controller.)


    '-----Variables-----
    'TODO Consider making ParentManagers
    '   The Manager would need to create a child manager to handle multiple parents.
    '   Whereas storing a reference to a single parent would be easier to propigate events.
    '   We'll see if it's neccessary.

    'ObserversManager          (FormManeger.ControllerGroups)
    'ParentGroupManager        (Not used atm.)

    'blnIsTopLevelManager    (Is set in Form Code.)

    'ParentManager           (Parent Manager if not top level)
    'TopLevelManager         (Manager of Main Form, passed down through constructors.)

    'blnHasChildControls
    'blnHasChildManagers

    'ChildControllerGroup    Controllers the object contains and manages.
    'ChildManagerGroups      Managers the object contains and manages

    '----- ------

    '-----Methods------
    'AddControl
    'AddControlGroup
'-----

'Private ChildController() As ControlController  'To pass Manager as controller arg, cause no OOP?

Private ParentManager As ControllerManager      'Parent Manager reference. | Maybe a group if multiple.

Private ChildControllers() As ControlController
Private ChildGUIControls As Collection

private ChildFormManagers() as ControllerManagers

Private MyTypeName As String
Private blnSameType As Boolean      '
Private blnIsRegistered As Boolean  'Registration Flag

'TODO: this would solve the issue of tracking controller objects when elevating a manager's event.

'-----[ Event definition ]-----------------------------------------------
Public Event Change(ByVal Index As Integer)
Public Event Click(ByVal Index As Integer)
Public Event DblClick(ByVal Index As Integer, ByVal Cancel As Boolean)
Public Event KeyDown(ByVal Index As Integer, ByVal KeyCode As Integer, ByVal Shift As Integer)
Public Event KeyPress(ByVal Index As Integer, ByVal KeyAscii As Integer)
Public Event KeyUp(ByVal Index As Integer, ByVal KeyCode As Integer, ByVal Shift As Integer)
Public Event MouseMove(ByVal Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Public Event MouseDown(ByVal Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Public Event MouseUp(ByVal Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)


'Because Enter/Exit/BeforeUpdate/AfterUpdate are events of MsForm.Control
'(it cannot treat in the class module), we hook those events by ConnectToConnectionPoint.
'Because the name of "Exit event" and "Exit statement" collide,
'I add "On" to the head of the name of the Enter/Exit event.
Public Event OnEnter(ByVal Index As Integer)
Public Event OnExit(ByVal Index As Integer, ByVal Cancel As Boolean)
Public Event AfterUpdate(ByVal Index As Integer)
Public Event BeforeUpdate(ByVal Index As Integer, ByVal Cancel As Boolean)

'An event to improve the issue that Exit event does not fire
'at the time of Frame/MutiPage escape.
'It detect Exit event of Frame/MultiPage and Change event of MultiPage.
'And it fire FakeExit event for ActiveControl at that time.
Public Event FakeExit(ByVal Index As Integer)


Private Sub Class_Initialize()
    Set ChildGUIControls = New Collection
    blnIsRegistered = False
    blnSameType = False
    MyTypeName = ""
End Sub

'==================================================================
Private Sub Class_Terminate()
    If (blnIsRegistered = True) Or (Not (ChildGUIControls Is Nothing)) Then
        Me.Clear
    End If
End Sub

'======== Property / Method definition =============================

'---( Add method )------------------------------------------------
'Add control to become array to collection.
Public Sub Add(ByVal NewCtrl As Object)
    If (blnIsRegistered = False) Then
        Select Case TypeName(NewCtrl)
            Case "Label", "CommandButton", "TextBox", "CheckBox", _
                 "OptionButton", "ToggleButton", "ComboBox"
                'Add to Collection
                childguicontrol.Add NewCtrl
            Case Else
                'do not add the non-supported control
        End Select
    Else
        'Because Rgst method is executed, I do not add it.
    End If
End Sub

'---( Clear method )-----------------------------
Public Sub Clear()
Dim i As Integer
Dim j As Integer

    If (blnRgst = True) Then
        ;For i = 1 To ChildGUIControls.Count
            ChildControllers(i).Clear
        Next i
        
        On Error Resume Next
        j = -1
        j = LBound(ChildFormManagers)
        On Error GoTo 0
        For i = 1 To j
            ChildFormManagers(i).Clear
        Next i
    End If
    
    Set ChildGUIControls = Nothing
    Erase ChildControllers

    Set colBpcaContainerGrp = Nothing
    Erase ChildFormManagers

    blnRgst = False
End Sub


'By MultiPage(Change),
'it fire FakeExit event of ActiveControl in the container control.
'OldPage: Previous page number of Page_Change event.
'TODO: Ug change4fake
Friend Sub Change4FakeExit(ByVal Index As Integer, ByVal OldPage As Integer)
'Dim Index2 As Integer
'Dim ContainerMlt As MSForms.MultiPage
'Dim ActCtrl As Object
'-----
    '    Set ContainerMlt = colBpcaContainerGrp(Index)
        
    '    On Error Resume Next
    '    Set ActCtrl = ContainerMlt.Pages(OldPage).ActiveControl
    '    On Error GoTo 0
        
    '    If Not (ActCtrl Is Nothing) Then
    '        Index2 = Me.getIndex(ActCtrl.Name)
    '        If (Index2 <> -1) Then
    '            RaiseEvent FakeExit(Index2)
    '        End If
    '    End If
        
    '    Set ContainerMlt = Nothing
'    Set ActCtrl = Nothing
'-----
End Sub


'By Frame/MultiPage(Exit),
'it fire FakeExit event of ActiveControl in the container control.
Friend Sub Exit4FakeExit(ByVal Index As Integer)
'Dim Index2 As Integer
'Dim ContainerMlt As MSForms.MultiPage
'Dim ContainerFra As MSForms.Frame
'Dim ActCtrl As Object
    
'------
    '    Select Case TypeName(colBpcaContainerGrp(Index))
    '      Case "MultiPage"
    '        Set ContainerMlt = colBpcaContainerGrp(Index)
    '        On Error Resume Next
    '        Set ActCtrl = ContainerMlt.Pages(ContainerMlt.Value).ActiveControl
    '        On Error GoTo 0
    '      Case "Frame"
    '        Set ContainerFra = colBpcaContainerGrp(Index)
    '        On Error Resume Next
    '        Set ActCtrl = ContainerFra.ActiveControl
    '        On Error GoTo 0
    '    End Select
        
    '    If Not (ActCtrl Is Nothing) Then
    '        Index2 = Me.getIndex(ActCtrl.Name)
    '        If (Index2 <> -1) Then
    '            RaiseEvent FakeExit(Index2)
    '        End If
    '    End If
        
    '    Set ContainerMlt = Nothing
    '    Set ContainerFra = Nothing
    '    Set ActCtrl = Nothing
'-----
End Sub
'===================================================================
' Event excitation method to be called from clsBpcaCh
'===================================================================
'Because the name of "Exit event" and "Exit statement" collide,
'I add "On" to the head of the name of the Enter/Exit event.

Friend Sub RaiseEnter(ByVal Index As Integer)
    RaiseEvent OnEnter(Index)
End Sub
  
Friend Sub RaiseExit(ByVal Index As Integer, ByVal Cancel As Boolean)
    RaiseEvent OnExit(Index, Cancel)
End Sub
  
Friend Sub RaiseAfterUpdate(ByVal Index As Integer)
    RaiseEvent AfterUpdate(Index)
End Sub
  
Friend Sub RaiseBeforeUpdate(ByVal Index As Integer, ByVal Cancel As Boolean)
    RaiseEvent BeforeUpdate(Index, Cancel)
End Sub

Friend Sub RaiseChange(ByVal Index As Integer)
    RaiseEvent Change(Index)
End Sub

Friend Sub RaiseClick(ByVal Index As Integer)
    RaiseEvent Click(Index)
End Sub

Friend Sub RaiseDblClick(ByVal Index As Integer, _
                         ByVal Cancel As Boolean)
    RaiseEvent DblClick(Index, Cancel)
End Sub

Friend Sub RaiseKeyDown(ByVal Index As Integer, _
                        ByVal KeyCode As Integer, _
                        ByVal Shift As Integer)
    RaiseEvent KeyDown(Index, KeyCode, Shift)
End Sub

Friend Sub RaiseKeyPress(ByVal Index As Integer, _
                         ByVal KeyAscii As Integer)
    RaiseEvent KeyPress(Index, KeyAscii)
End Sub

Friend Sub RaiseKeyUp(ByVal Index As Integer, _
                      ByVal KeyCode As Integer, _
                      ByVal Shift As Integer)
    RaiseEvent KeyUp(Index, KeyCode, Shift)
End Sub

Friend Sub RaiseMouseMove(ByVal Index As Integer, _
                          ByVal Button As Integer, _
                          ByVal Shift As Integer, _
                          ByVal X As Single, _
                          ByVal Y As Single)
    RaiseEvent MouseMove(Index, Button, Shift, X, Y)
End Sub

Friend Sub RaiseMouseDown(ByVal Index As Integer, _
                          ByVal Button As Integer, _
                          ByVal Shift As Integer, _
                          ByVal X As Single, _
                          ByVal Y As Single)
    RaiseEvent MouseDown(Index, Button, Shift, X, Y)
End Sub

Friend Sub RaiseMouseUp(ByVal Index As Integer, _
                        ByVal Button As Integer, _
                        ByVal Shift As Integer, _
                        ByVal X As Single, _
                        ByVal Y As Single)
    RaiseEvent MouseUp(Index, Button, Shift, X, Y)
End Sub




