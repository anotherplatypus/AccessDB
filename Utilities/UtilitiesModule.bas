Attribute VB_Name = "modBpcaConst"
Option Explicit

'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/
'_/ --- Breakthrough in the Pseudo-Control-Array  --- (Excel2000 or later)
'_/  [ clsBpca/clsBpcaCh Ver4.0 ] Event type constants
'_/
'_/     Control: Label, CommandButton, TextBox, CheckBox
'_/              OptionButton, ToggleButton, ComboBox, SpinButton
'_/
'_/     Event  : Enter, Exit, BeforeUpdate, AfterUpdate
'_/              Change, Click, DblClick, KeyDown, KeyPress, KeyUp
'_/              SpinDown, SpinUp, DropButtonClick
'_/              MouseMove, MouseDown, MouseUp, (FakeExit)
'_/
'_/     AddinBox(K.Tsunoda) CopyRight(C) 2004 Allrights Reserved.
'_/       [http://addinbox.sakura.ne.jp/Breakthrough_P-Ctrl_Arrays_Eng.htm ]
'_/       [Old Site : http://www.h3.dion.ne.jp/~sakatsu/Breakthrough_P-Ctrl_Arrays_Eng.htm ]
'_/
'_/     ----- Japanese version ----
'_/     22 Jun 2004 Ver1.0  1st version
'_/     23 Jun 2004 Ver1.1
'_/     10 Mar 2005 Ver1.2
'_/     15 Apr 2011 Ver1.3
'_/     22 Jul 2014 Ver1.4
'_/     11 Aug 2014 Ver2.0  Enter/Exit/BeforeUpdate/AfterUpdate are supported.
'_/     10 Oct 2016 Ver3.0  FakeExit is added.
'_/     13 Oct 2016 Ver3.1
'_/     (Ver3.2 - Ver3.4)   Replace "Japan Holiday module" for Calendar form.
'_/      1 Sep 2020 Ver4.0
'_/
'_/     ----- English version ----
'_/     24 Jul 2014 Ver1.4  1st version
'_/     11 Aug 2014 Ver2.0  Enter/Exit/BeforeUpdate/AfterUpdate are supported.
'_/     10 Oct 2016 Ver3.0  FakeExit is added.
'_/     13 Oct 2016 Ver3.1
'_/      1 Sep 2020 Ver4.0
'_/
'_/
'_/  A definition of this event constants is necessary to use clsBpca/clsBpcaCh class.
'_/  It works in x64.
'_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/


Public Const BPCA2_Change As Long = 1            '[ 1]bit(ON)
Public Const BPCA2_Click As Long = 2             '[ 2]
Public Const BPCA2_DblClick As Long = 4          '[ 3]
Public Const BPCA2_KeyDown As Long = 8           '[ 4]
Public Const BPCA2_KeyPress As Long = 16         '[ 5]
Public Const BPCA2_KeyUp As Long = 32            '[ 6]
Public Const BPCA2_MouseMove As Long = 64        '[ 7]
Public Const BPCA2_MouseDown As Long = 128       '[ 8]
Public Const BPCA2_MouseUp As Long = 256         '[ 9]
Public Const BPCA2_DropBtnClick As Long = 512    '[10]
Public Const BPCA2_SpinDown As Long = 1024       '[11]
Public Const BPCA2_SpinUp As Long = 2048         '[12]
Public Const BPCA2_Enter As Long = 4096          '[13]
Public Const BPCA2_Exit As Long = 8192           '[14]
Public Const BPCA2_BeforeUpdate As Long = 16384  '[15]
Public Const BPCA2_AfterUpdate As Long = 32768   '[16]
Public Const BPCA2_FakeExit As Long = 65536      '[17]

Public Const BPCA2_All As Long = &HFFFF&     '(65535 , 16bit All-ON , FakeExit is excluded)

Public Const BPCA2_KeyDU As Long = BPCA2_KeyDown + BPCA2_KeyUp
Public Const BPCA2_KeyDPU As Long = BPCA2_KeyDown + BPCA2_KeyPress + BPCA2_KeyUp
Public Const BPCA2_MouseDU As Long = BPCA2_MouseDown + BPCA2_MouseUp
Public Const BPCA2_MouseMDU As Long = BPCA2_MouseMove + BPCA2_MouseDown + BPCA2_MouseUp
Public Const BPCA2_SpinDU As Long = BPCA2_SpinDown + BPCA2_SpinUp
Public Const BPCA2_EnterExit As Long = BPCA2_Enter + BPCA2_Exit
Public Const BPCA2_BAUpdate As Long = BPCA2_BeforeUpdate + BPCA2_AfterUpdate

Public Const BPCA2_Except_MouseM As Long = BPCA2_All - BPCA2_MouseMove


'-- Enumeration list of Event type constants --
Public Enum BPCA_Event
    BPCA_Change = BPCA2_Change
    BPCA_Click = BPCA2_Click
    BPCA_DblClick = BPCA2_DblClick
    BPCA_KeyDown = BPCA2_KeyDown
    BPCA_KeyPress = BPCA2_KeyPress
    BPCA_KeyUp = BPCA2_KeyUp
    BPCA_MouseMove = BPCA2_MouseMove
    BPCA_MouseDown = BPCA2_MouseDown
    BPCA_MouseUp = BPCA2_MouseUp
    BPCA_DropBtnClick = BPCA2_DropBtnClick
    BPCA_SpinDown = BPCA2_SpinDown
    BPCA_SpinUp = BPCA2_SpinUp
    BPCA_Enter = BPCA2_Enter
    BPCA_Exit = BPCA2_Exit
    BPCA_BeforeUpdate = BPCA2_BeforeUpdate
    BPCA_AfterUpdate = BPCA2_AfterUpdate
    BPCA_FakeExit = BPCA2_FakeExit
    
    BPCA_All = BPCA2_All
    
    BPCA_KeyDU = BPCA2_KeyDU
    BPCA_KeyDPU = BPCA2_KeyDPU
    BPCA_MouseDU = BPCA2_MouseDU
    BPCA_MouseMDU = BPCA2_MouseMDU
    BPCA_SpinDU = BPCA2_SpinDU
    BPCA_EnterExit = BPCA2_EnterExit
    BPCA_BAUpdate = BPCA2_BAUpdate
    BPCA_Except_MouseM = BPCA2_Except_MouseM
End Enum

'------------( modBpcaConst )-----------------------------------------
