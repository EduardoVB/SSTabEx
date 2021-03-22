Attribute VB_Name = "mSubclass"
Option Explicit

' ======================================================================================
' Name:     vbAccelerator SSubTmr object
'           MSubClass.bas
' Author:   Steve McMahon (steve@vbaccelerator.com)
' Date:     25 June 1998
'
' Requires: None
'
' Copyright © 1998-2003 Steve McMahon for vbAccelerator
' --------------------------------------------------------------------------------------
' Visit vbAccelerator - advanced free source code for VB programmers
' http://vbaccelerator.com
' --------------------------------------------------------------------------------------
'
' The implementation of the Subclassing part of the SSubTmr object.
' Use this module + ISubClass.Cls to replace dependency on the DLL.
'
' Fixes:
' 23 Jan 03
' SPM: Fixed multiple attach/detach bug which resulted in incorrectly setting
' the message count.
' SPM: Refactored code
' SPM: Added automated detach on WM_DESTROY
' 27 Dec 99
' DetachMessage: Fixed typo in DetachMessage which removed more messages than it should
'   (Thanks to Vlad Vissoultchev <wqw@bora.exco.net>)
' DetachMessage: Fixed resource leak (very slight) due to failure to remove property
'   (Thanks to Andrew Smith <asmith2@optonline.net>)
' AttachMessage: Added extra error handlers
'
' ======================================================================================

' Note: it is a completely modified version.
' Date 16 Nov 2017
' It uses common controls subclass for better compatibility with other projects, but keeping the interface and idea of only sending to the windows procedure the messages that the developer wants to handle;
' allowing to preprocess, postprocess or consume any particular message.

' declares:
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function SetWindowSubclass Lib "comctl32" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
Private Declare Function DefSubclassProc Lib "comctl32" (ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function RemoveWindowSubclass Lib "comctl32" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function EbModeVBA5 Lib "vba5" Alias "EbMode" () As Long
Private Declare Function EbModeVBA6 Lib "vba6" Alias "EbMode" () As Long
Private Declare Function EbIsResettingVBA5 Lib "vba5" Alias "EbIsResetting" () As Long
Private Declare Function EbIsResettingVBA6 Lib "vba6" Alias "EbIsResetting" () As Long

Private Const GWL_WNDPROC = (-4)
Private Const WM_DESTROY = &H2
Private Const WM_NCDESTROY As Long = &H82&
Private Const WM_UAHDESTROYWINDOW As Long = &H90& 'Undocumented.

' SubTimer is independent of VBCore, so it hard codes error handling

Public Enum EErrorWindowProc
    eeBaseWindowProc = 13080 ' WindowProc
    eeCantSubclass           ' Can't subclass window
    eeAlreadyAttached        ' Message already handled by another class
    eeInvalidWindow          ' Invalid window
    eeNoExternalWindow       ' Can't modify external window
End Enum

Private m_iCurrentMessage As Long
Private m_f As Long

Private mPropsDatabaseChecked As Boolean
Private mUseLocalPropsDB As Boolean
Private mAddressOfWindowProc As Long

Public Property Get CurrentMessage() As Long
    CurrentMessage = m_iCurrentMessage
End Property

Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    
    If e > 1000 Then
        sSource = App.EXEName & ".WindowProc"
        Select Case e
        Case eeCantSubclass
            sText = "Can't subclass window"
        Case eeAlreadyAttached
            sText = "Message already handled by the same object"
        Case eeInvalidWindow
            sText = "Invalid window"
        Case eeNoExternalWindow
            sText = "Can't modify external window"
        End Select
        Err.Raise e Or vbObjectError, sSource, sText
    Else
        ' Raise standard Visual Basic error
        Err.Raise e, sSource
    End If
End Sub

Private Property Get MessageCount(ByVal hWnd As Long) As Long
    Dim sName As String
    
    sName = "C" & hWnd
    MessageCount = ThisGetProp(hWnd, sName)
    If MessageCount > 1000000 Then
        mUseLocalPropsDB = True
        MessageCount = ThisGetProp(hWnd, sName)
        If MessageCount > 1000000 Then
            MessageCount = 10
        End If
    End If
End Property

Private Property Let MessageCount(ByVal hWnd As Long, ByVal Count As Long)
    Dim sName As String
    
    m_f = 1
    sName = "C" & hWnd
    m_f = ThisSetProp(hWnd, sName, Count)
    If (Count = 0) Then
        ThisRemoveProp hWnd, sName
    End If
    'logMessage "Changed message count for " & Hex(hWnd) & " to " & count
End Property

Private Property Get MessageClassCount(ByVal hWnd As Long, ByVal iMsg As Long) As Long
    Dim sName As String
    
    sName = hWnd & "#" & iMsg & "C"
    MessageClassCount = ThisGetProp(hWnd, sName)
    If MessageClassCount > 1000000 Then
        mUseLocalPropsDB = True
        MessageClassCount = ThisGetProp(hWnd, sName)
        If MessageClassCount > 1000000 Then
            MessageClassCount = 10
        End If
    End If
    
End Property

Private Property Let MessageClassCount(ByVal hWnd As Long, ByVal iMsg As Long, ByVal Count As Long)
    Dim sName As String
    
    sName = hWnd & "#" & iMsg & "C"
    m_f = ThisSetProp(hWnd, sName, Count)
    If (Count = 0) Then
       ThisRemoveProp hWnd, sName
    End If
    'logMessage "Changed message count for " & Hex(hWnd) & " Message " & iMsg & " to " & count
End Property

Private Property Get MessageClass(ByVal hWnd As Long, ByVal iMsg As Long, ByVal Index As Long) As Long
    Dim sName As String
    sName = hWnd & "#" & iMsg & "#" & Index
    MessageClass = ThisGetProp(hWnd, sName)
End Property

Private Property Let MessageClass(ByVal hWnd As Long, ByVal iMsg As Long, ByVal Index As Long, ByVal classPtr As Long)
    Dim sName As String
    
    sName = hWnd & "#" & iMsg & "#" & Index
    m_f = ThisSetProp(hWnd, sName, classPtr)
    If (classPtr = 0) Then
       ThisRemoveProp hWnd, sName
    End If
    'logMessage "Changed message class for " & Hex(hWnd) & " Message " & iMsg & " Index " & index & " to " & Hex(classPtr)
End Property

Sub AttachMessage(iwp As ISubclass, ByVal hWnd As Long, ByVal iMsg As Long)
    Dim msgCount As Long
    Dim msgClassCount As Long
    Dim msgClass As Long
    Dim iLng As Long

'   If InIDE Then Exit Sub
    If Not mPropsDatabaseChecked Then
         CheckPropsDatabase
    End If
    
'    mUseLocalPropsDB = True
    
    ' --------------------------------------------------------------------
    ' 1) Validate window
    ' --------------------------------------------------------------------
    If IsWindow(hWnd) = False Then
       ErrRaise eeInvalidWindow
       Exit Sub
    End If
    If IsWindowLocal(hWnd) = False Then
       ErrRaise eeNoExternalWindow
       Exit Sub
    End If

    ' --------------------------------------------------------------------
    ' 2) Check if this class is already attached for this message:
    ' --------------------------------------------------------------------
    msgClassCount = MessageClassCount(hWnd, iMsg)
    If (msgClassCount > 0) Then
        For msgClass = 1 To msgClassCount
            iLng = MessageClass(hWnd, iMsg, msgClass)
            If iLng = 0 Then
                mUseLocalPropsDB = True
                iLng = MessageClass(hWnd, iMsg, msgClass)
                If iLng = 0 Then
                    Exit Sub
                End If
            End If
            If (iLng = ObjPtr(iwp)) Then
'                ErrRaise eeAlreadyAttached
                Exit Sub
            End If
        Next msgClass
    End If

    ' --------------------------------------------------------------------
    ' 3) Associate this class with this message for this window:
    ' --------------------------------------------------------------------
    MessageClassCount(hWnd, iMsg) = MessageClassCount(hWnd, iMsg) + 1
    If (m_f = 0) Then
        ' Failed, out of memory:
        ErrRaise 5
        Exit Sub
    End If
   
    ' --------------------------------------------------------------------
    ' 4) Associate the class pointer:
    ' --------------------------------------------------------------------
    MessageClass(hWnd, iMsg, MessageClassCount(hWnd, iMsg)) = ObjPtr(iwp)
    If (m_f = 0) Then
        ' Failed, out of memory:
        MessageClassCount(hWnd, iMsg) = MessageClassCount(hWnd, iMsg) - 1
        ErrRaise 5
        Exit Sub
    End If
    
    ' --------------------------------------------------------------------
    ' 5) Get the message count
    ' --------------------------------------------------------------------
    msgCount = MessageCount(hWnd)
    If msgCount = 0 Then
        
        ' Subclass window by installing window procedure
        If SetWindowSubclass(hWnd, AddressOf WindowProc, ObjPtr(iwp), 0&) = 0 Then
            ' remove class:
            MessageClass(hWnd, iMsg, MessageClassCount(hWnd, iMsg)) = 0
            ' remove class count:
            MessageClassCount(hWnd, iMsg) = MessageClassCount(hWnd, iMsg) - 1
            
            ErrRaise eeCantSubclass
            Exit Sub
        Else
            If mAddressOfWindowProc = 0 Then
                mAddressOfWindowProc = GetAddresOfProc(AddressOf WindowProc)
            End If
        End If
    End If
   
      
    ' Count this message
    MessageCount(hWnd) = MessageCount(hWnd) + 1
    If m_f = 0 Then
        ' SPM: Failed to set prop, windows properties database problem.
        ' Has to be out of memory
        
        ' remove class:
        MessageClass(hWnd, iMsg, MessageClassCount(hWnd, iMsg)) = 0
        ' remove class count contribution:
        MessageClassCount(hWnd, iMsg) = MessageClassCount(hWnd, iMsg) - 1
        
        ' If we haven't any messages on this window then remove the subclass:
        If (MessageCount(hWnd) = 0) Then
            ' put old window proc back again:
            RemoveWindowSubclass hWnd, mAddressOfWindowProc, ObjPtr(iwp)
        End If
        
        ' Raise the error:
        ErrRaise 5
        Exit Sub
    End If
End Sub

Sub DetachMessage(iwp As ISubclass, ByVal hWnd As Long, ByVal iMsg As Long)
    Dim msgClassCount As Long
    Dim msgClass As Long
    Dim msgClassIndex As Long
    Dim msgCount As Long
    Dim iLng As Long
    
    ' --------------------------------------------------------------------
    ' 1) Validate window
    ' --------------------------------------------------------------------
    If IsWindow(hWnd) = False Then
        ' for compatibility with the old version, we don't
        ' raise a message:
        ' ErrRaise eeInvalidWindow
        Exit Sub
    End If
    If IsWindowLocal(hWnd) = False Then
        ' for compatibility with the old version, we don't
        ' raise a message:
        ' ErrRaise eeNoExternalWindow
        Exit Sub
    End If
    
    ' --------------------------------------------------------------------
    ' 2) Check if this message is attached for this class:
    ' --------------------------------------------------------------------
    msgClassCount = MessageClassCount(hWnd, iMsg)
    If (msgClassCount > 0) Then
        msgClassIndex = 0
        For msgClass = 1 To msgClassCount
            iLng = MessageClass(hWnd, iMsg, msgClass)
            If iLng = 0 Then
                Exit For
            End If
            If (iLng = ObjPtr(iwp)) Then
                msgClassIndex = msgClass
                Exit For
            End If
        Next msgClass
        
        If (msgClassIndex = 0) Then
            ' fail silently
            Exit Sub
        Else
            ' remove this message class:
            
            ' a) Anything above this index has to be shifted up:
            For msgClass = msgClassIndex To msgClassCount - 1
                iLng = MessageClass(hWnd, iMsg, msgClass + 1)
                If iLng = 0 Then
                    Exit For
                End If
                MessageClass(hWnd, iMsg, msgClass) = iLng
            Next msgClass
            
            ' b) The message class at the end can be removed:
            MessageClass(hWnd, iMsg, msgClassCount) = 0
            
            ' c) Reduce the message class count:
            MessageClassCount(hWnd, iMsg) = MessageClassCount(hWnd, iMsg) - 1
        
        End If
       
    Else
        ' fail silently
        Exit Sub
    End If
   
    ' ---------------------------------------------------------------------
    ' 3) Reduce the message count:
    ' ---------------------------------------------------------------------
    msgCount = MessageCount(hWnd)
    If (msgCount = 1) Then
        ' remove the subclass:
        RemoveWindowSubclass hWnd, mAddressOfWindowProc, ObjPtr(iwp)
    End If
    MessageCount(hWnd) = MessageCount(hWnd) - 1
End Sub

Private Function WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
    Dim bCalled As Boolean
    Dim pSubClass As Long
    Dim iwp As ISubclass
    Dim iwpT As ISubclass
    Dim iIndex As Long
    Dim iHandled As Boolean
    Dim bConsume As Boolean
    Dim iResp As Long
    
    If IsResetting Then
        pClearUp hWnd, uIdSubclass
        Exit Function
    End If
    If IsWindow(hWnd) = 0 Then
        pClearUp hWnd, uIdSubclass
        Exit Function
    End If
    If InBreakMode Then
        WindowProc = DefSubclassProc(hWnd, iMsg, wParam, lParam)
        Exit Function
    End If
    
    ' SPM - in this version I am allowing more than one class to
    ' make a subclass to the same hWnd and Msg.  Why am I doing
    ' this?  Well say the class in question is a control, and it
    ' wants to subclass its container.  In this case, we want
    ' all instances of the control on the form to receive the
    ' form notification message.
     
    ' Get the number of instances for this msg/hWnd:
    bCalled = False
   
    If (MessageClassCount(hWnd, iMsg) > 0) Then
        iIndex = MessageClassCount(hWnd, iMsg)
        
        Do While (iIndex >= 1)
            pSubClass = MessageClass(hWnd, iMsg, iIndex)
            
            If (pSubClass = 0) Then
                ' Not handled by this instance
            Else
                iHandled = True
                ' Turn pointer into a reference:
                CopyMemory iwpT, pSubClass, 4
                Set iwp = iwpT
                CopyMemory iwpT, 0&, 4
                
                ' Store the current message, so the client can check it:
                m_iCurrentMessage = iMsg
                
                With iwp
                    ' Preprocess (only checked first time around):
                    On Error GoTo TheExit:
                    If (.MsgResponse(hWnd, iMsg) = emrPreprocess) Then
                        On Error GoTo 0
                        ' Consume (this message is always passed to all control
                        ' instances regardless of whether any single one of them
                        ' requests to consume it):
                        WindowProc = .WindowProc(hWnd, iMsg, wParam, lParam, bConsume)
                        
                        If Not bConsume Then
                            If (iIndex = 1) Then
                                If Not (bCalled) Then
                                    WindowProc = DefSubclassProc(hWnd, iMsg, wParam, ByVal lParam)
                                    bCalled = True
                                End If
                            End If
                        End If
                        On Error GoTo 0
                    Else
                        ' Consume (this message is always passed to all control
                        ' instances regardless of whether any single one of them
                        ' requests to consume it):
                        WindowProc = .WindowProc(hWnd, iMsg, wParam, lParam, bConsume)
                    End If
                End With
            End If
            
            iIndex = iIndex - 1
       Loop
       
       ' PostProcess (only check this the last time around):
        If Not (iwp Is Nothing) Then
            iResp = iwp.MsgResponse(hWnd, iMsg)
            If (iResp = emrPostProcess) Then
                If Not (bCalled) Then
                    WindowProc = DefSubclassProc(hWnd, iMsg, wParam, ByVal lParam)
                    bCalled = True
                End If
            End If
        End If
        
        If Not iHandled Then
            WindowProc = DefSubclassProc(hWnd, iMsg, wParam, ByVal lParam)
            If GetWindowLong(hWnd, GWL_WNDPROC) = mAddressOfWindowProc Then     ' if we are at the top of the subclassing chain, else we'll wait for the WM_DESTROY, WM_NCDESTROY and WM_UAHDESTROYWINDOW messages
                pClearUp hWnd, uIdSubclass
            End If
        End If
    Else
        ' Not handled:
        If (iMsg = WM_DESTROY) Or (iMsg = WM_NCDESTROY) Or (iMsg = WM_UAHDESTROYWINDOW) Then
            ' If WM_DESTROY isn't handled already, we should
            ' clear up any subclass
            If GetWindowLong(hWnd, GWL_WNDPROC) = mAddressOfWindowProc Then ' if we are at the top of the subclassing chain
                WindowProc = DefSubclassProc(hWnd, iMsg, wParam, ByVal lParam)
                pClearUp hWnd, uIdSubclass
            Else ' we are not a the top subclassing chain
                WindowProc = DefSubclassProc(hWnd, iMsg, wParam, ByVal lParam)  ' let's see if the other subclass unsubclass itself
                If GetWindowLong(hWnd, GWL_WNDPROC) = mAddressOfWindowProc Then ' it did
                    pClearUp hWnd, uIdSubclass
                Else
                    If (iMsg = WM_NCDESTROY) Or (iMsg = WM_UAHDESTROYWINDOW) Then ' in these cases we will unsubclass anyway, but for WM_DESTROY we will wait for the WM_NCDESTROY message
                        pClearUp hWnd, uIdSubclass
                    End If
                End If
            End If
        Else
            WindowProc = DefSubclassProc(hWnd, iMsg, wParam, ByVal lParam)
        End If
    End If
    
TheExit:
End Function

Public Function CallOldWindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    CallOldWindowProc = DefSubclassProc(hWnd, iMsg, wParam, lParam)
End Function

Private Function IsWindowLocal(ByVal hWnd As Long) As Boolean
    Dim idWnd As Long
    
    Call GetWindowThreadProcessId(hWnd, idWnd)
    IsWindowLocal = (idWnd = GetCurrentProcessId())
End Function

'Private Sub logMessage(ByVal sMsg As String)
'    Debug.Print sMsg
'End Sub


Private Sub pClearUp(ByVal hWnd As Long, uIdSubclass As Long)
    Dim msgCount As Long
    
    ' this is only called if you haven't explicitly cleared up
    ' your subclass from the caller.  You will get a minor
    ' resource leak as it does not clear up any message
    ' specific properties.
    msgCount = MessageCount(hWnd)
    If (msgCount > 0) Then
        ' remove the subclass:
        ' Unsubclass
        RemoveWindowSubclass hWnd, mAddressOfWindowProc, uIdSubclass
        ' remove the old window proc:
        MessageCount(hWnd) = 0
    End If
End Sub

Private Function ThisGetProp(ByVal hWnd As Long, ByVal lpString As String) As Long
#If UseOnlyLocalDB Then
    ThisGetProp = MyGetProp(hWnd, lpString)
#Else
    If mUseLocalPropsDB Then
        ThisGetProp = GetProp(hWnd, lpString)
        If ThisGetProp = 0 Then
            ThisGetProp = MyGetProp(hWnd, lpString)
        End If
    Else
        ThisGetProp = GetProp(hWnd, lpString)
    End If
#End If
End Function

Private Function ThisSetProp(ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
#If UseOnlyLocalDB Then
    If hData = 0 Then
        ThisSetProp = MyRemoveProp(hWnd, lpString)
    Else
        If MyGetProp(hWnd, lpString) <> 0 Then
            ThisSetProp = MyRemoveProp(hWnd, lpString)
            MySetProp hWnd, lpString, hData
        Else
            ThisSetProp = MySetProp(hWnd, lpString, hData)
        End If
    End If
#Else
    If mUseLocalPropsDB Then
        If hData = 0 Then
            ThisSetProp = MyRemoveProp(hWnd, lpString)
        Else
            If MyGetProp(hWnd, lpString) <> 0 Then
                ThisSetProp = MyRemoveProp(hWnd, lpString)
                MySetProp hWnd, lpString, hData
            Else
                ThisSetProp = MySetProp(hWnd, lpString, hData)
            End If
        End If
    Else
        If hData = 0 Then
            ThisSetProp = RemoveProp(hWnd, lpString)
            MyRemoveProp hWnd, lpString
        Else
            If GetProp(hWnd, lpString) <> 0 Then
                ThisSetProp = RemoveProp(hWnd, lpString)
                MyRemoveProp hWnd, lpString
                SetProp hWnd, lpString, hData
                MySetProp hWnd, lpString, hData
            Else
                ThisSetProp = SetProp(hWnd, lpString, hData)
                MySetProp hWnd, lpString, hData
            End If
        End If
    End If
#End If
End Function

Private Function ThisRemoveProp(ByVal hWnd As Long, ByVal lpString As String) As Long
#If UseOnlyLocalDB Then
    ThisRemoveProp = MyRemoveProp(hWnd, lpString)
#Else
    If mUseLocalPropsDB Then
        ThisRemoveProp = RemoveProp(hWnd, lpString)
        If ThisRemoveProp = 0 Then
            ThisRemoveProp = MyRemoveProp(hWnd, lpString)
        Else
            MyRemoveProp hWnd, lpString
        End If
    Else
        ThisRemoveProp = RemoveProp(hWnd, lpString)
        MyRemoveProp hWnd, lpString
    End If
#End If
End Function


Private Function InIDE() As Boolean
    Static sValue As Long
    
    If sValue = 0 Then
        On Error Resume Next
        Err.Clear
        Debug.Assert "a"
        If Err.Number = 13 Then
            sValue = 1
        Else
            sValue = 2
        End If
        Err.Clear
    End If
    InIDE = sValue = 1
End Function

Private Sub CheckPropsDatabase()
    Dim c As Long
    Dim iHwnd As Long
    Dim iRnd As Long
    
    iHwnd = GetDesktopWindow
    Randomize
    iRnd = Rnd * 10000
    
    For c = 1 To 1000
        SetProp iHwnd, "TestPDB" & CStr(c), c + iRnd
    Next c
    For c = 1 To 1000
        If GetProp(iHwnd, "TestPDB" & CStr(c)) <> (c + iRnd) Then
            mUseLocalPropsDB = True
            Exit For
        End If
    Next c
    For c = 1 To 1000
        RemoveProp iHwnd, "TestPDB" & CStr(c)
    Next c
    mPropsDatabaseChecked = True
End Sub

Public Function GetMessageName(nMsg As Long) As String
   Dim msg As String
   
   Select Case nMsg
      Case &H0: msg = "WM_NULL"
      Case &H1: msg = "WM_CREATE"
      Case &H2: msg = "WM_DESTROY"
      Case &H3: msg = "WM_MOVE"
      Case &H5: msg = "WM_SIZE"
      Case &H6: msg = "WM_ACTIVATE"
      Case &H7: msg = "WM_SETFOCUS"
      Case &H8: msg = "WM_KILLFOCUS"
      Case &HA: msg = "WM_ENABLE"
      Case &HB: msg = "WM_SETREDRAW"
      Case &HC: msg = "WM_SETTEXT"
      Case &HD: msg = "WM_GETTEXT"
      Case &HE: msg = "WM_GETTEXTLENGTH"
      Case &HF: msg = "WM_PAINT"
      Case &H10: msg = "WM_CLOSE"
      Case &H11: msg = "WM_QUERYENDSESSION"
      Case &H12: msg = "WM_QUIT"
      Case &H13: msg = "WM_QUERYOPEN"
      Case &H14: msg = "WM_ERASEBKGND"
      Case &H15: msg = "WM_SYSCOLORCHANGE"
      Case &H16: msg = "WM_ENDSESSION"
      Case &H18: msg = "WM_SHOWWINDOW"
      Case &H1A: msg = "WM_SETTINGCHANGE"
      Case &H1B: msg = "WM_DEVMODECHANGE"
      Case &H1C: msg = "WM_ACTIVATEAPP"
      Case &H1D: msg = "WM_FONTCHANGE"
      Case &H1E: msg = "WM_TIMECHANGE"
      Case &H1F: msg = "WM_CANCELMODE"
      Case &H20: msg = "WM_SETCURSOR"
      Case &H21: msg = "WM_MOUSEACTIVATE"
      Case &H22: msg = "WM_CHILDACTIVATE"
      Case &H23: msg = "WM_QUEUESYNC"
      Case &H24: msg = "WM_GETMINMAXINFO"
      Case &H26: msg = "WM_PAINTICON"
      Case &H27: msg = "WM_ICONERASEBKGND"
      Case &H28: msg = "WM_NEXTDLGCTL"
      Case &H2A: msg = "WM_SPOOLERSTATUS"
      Case &H2B: msg = "WM_DRAWITEM"
      Case &H2C: msg = "WM_MEASUREITEM"
      Case &H2D: msg = "WM_DELETEITEM"
      Case &H2E: msg = "WM_VKEYTOITEM"
      Case &H2F: msg = "WM_CHARTOITEM"
      Case &H30: msg = "WM_SETFONT"
      Case &H31: msg = "WM_GETFONT"
      Case &H32: msg = "WM_SETHOTKEY"
      Case &H33: msg = "WM_GETHOTKEY"
      Case &H37: msg = "WM_QUERYDRAGICON"
      Case &H39: msg = "WM_COMPAREITEM"
      Case &H3D: msg = "WM_GETOBJECT"
      Case &H41: msg = "WM_COMPACTING"
      Case &H44: msg = "WM_COMMNOTIFY"
      Case &H46: msg = "WM_WINDOWPOSCHANGING"
      Case &H47: msg = "WM_WINDOWPOSCHANGED"
      Case &H48: msg = "WM_POWER"
      Case &H4A: msg = "WM_COPYDATA"
      Case &H4B: msg = "WM_CANCELJOURNAL"
      Case &H4E: msg = "WM_NOTIFY"
      Case &H50: msg = "WM_INPUTLANGCHANGEREQUEST"
      Case &H51: msg = "WM_INPUTLANGCHANGE"
      Case &H52: msg = "WM_TCARD"
      Case &H53: msg = "WM_HELP"
      Case &H54: msg = "WM_USERCHANGED"
      Case &H55: msg = "WM_NOTIFYFORMAT"
      Case &H7B: msg = "WM_CONTEXTMENU"
      Case &H7C: msg = "WM_STYLECHANGING"
      Case &H7D: msg = "WM_STYLECHANGED"
      Case &H7E: msg = "WM_DISPLAYCHANGE"
      Case &H7F: msg = "WM_GETICON"
      Case &H80: msg = "WM_SETICON"
      Case &H81: msg = "WM_NCCREATE"
      Case &H82: msg = "WM_NCDESTROY"
      Case &H83: msg = "WM_NCCALCSIZE"
      Case &H84: msg = "WM_NCHITTEST"
      Case &H85: msg = "WM_NCPAINT"
      Case &H86: msg = "WM_NCACTIVATE"
      Case &H87: msg = "WM_GETDLGCODE"
      Case &H88: msg = "WM_SYNCPAINT"
      Case &HA0: msg = "WM_NCMOUSEMOVE"
      Case &HA1: msg = "WM_NCLBUTTONDOWN"
      Case &HA2: msg = "WM_NCLBUTTONUP"
      Case &HA3: msg = "WM_NCLBUTTONDBLCLK"
      Case &HA4: msg = "WM_NCRBUTTONDOWN"
      Case &HA5: msg = "WM_NCRBUTTONUP"
      Case &HA6: msg = "WM_NCRBUTTONDBLCLK"
      Case &HA7: msg = "WM_NCMBUTTONDOWN"
      Case &HA8: msg = "WM_NCMBUTTONUP"
      Case &HA9: msg = "WM_NCMBUTTONDBLCLK"
      Case &HAB: msg = "WM_NCXBUTTONDOWN"
      Case &HAC: msg = "WM_NCXBUTTONUP"
      Case &HAD: msg = "WM_NCXBUTTONDBLCLK"
      Case &HFF: msg = "WM_INPUT"
      Case &H100: msg = "WM_KEYDOWN"
      Case &H101: msg = "WM_KEYUP"
      Case &H102: msg = "WM_CHAR"
      Case &H103: msg = "WM_DEADCHAR"
      Case &H104: msg = "WM_SYSKEYDOWN"
      Case &H105: msg = "WM_SYSKEYUP"
      Case &H106: msg = "WM_SYSCHAR"
      Case &H107: msg = "WM_SYSDEADCHAR"
      Case &H108: msg = "WM_KEYLAST"
      Case &H10D: msg = "WM_IME_STARTCOMPOSITION"
      Case &H10E: msg = "WM_IME_ENDCOMPOSITION"
      Case &H10F: msg = "WM_IME_COMPOSITION"
      Case &H110: msg = "WM_INITDIALOG"
      Case &H111: msg = "WM_COMMAND"
      Case &H112: msg = "WM_SYSCOMMAND"
      Case &H113: msg = "WM_TIMER"
      Case &H114: msg = "WM_HSCROLL"
      Case &H115: msg = "WM_VSCROLL"
      Case &H116: msg = "WM_INITMENU"
      Case &H117: msg = "WM_INITMENUPOPUP"
      Case &H11F: msg = "WM_MENUSELECT"
      Case &H120: msg = "WM_MENUCHAR"
      Case &H121: msg = "WM_ENTERIDLE"
      Case &H122: msg = "WM_MENURBUTTONUP"
      Case &H123: msg = "WM_MENUDRAG"
      Case &H124: msg = "WM_MENUGETOBJECT"
      Case &H125: msg = "WM_UNINITMENUPOPUP"
      Case &H126: msg = "WM_MENUCOMMAND"
      Case &H127: msg = "WM_CHANGEUISTATE"
      Case &H128: msg = "WM_UPDATEUISTATE"
      Case &H129: msg = "WM_QUERYUISTATE"
      Case &H132: msg = "WM_CTLCOLORMSGBOX"
      Case &H133: msg = "WM_CTLCOLOREDIT"
      Case &H134: msg = "WM_CTLCOLORLISTBOX"
      Case &H135: msg = "WM_CTLCOLORBTN"
      Case &H136: msg = "WM_CTLCOLORDLG"
      Case &H137: msg = "WM_CTLCOLORSCROLLBAR"
      Case &H138: msg = "WM_CTLCOLORSTATIC"
      Case &H1E1: msg = "MN_GETHMENU"
'      Case &H200: msg = "WM_MOUSEFIRST"
      Case &H200: msg = "WM_MOUSEMOVE"
      Case &H201: msg = "WM_LBUTTONDOWN"
      Case &H202: msg = "WM_LBUTTONUP"
      Case &H203: msg = "WM_LBUTTONDBLCLK"
      Case &H204: msg = "WM_RBUTTONDOWN"
      Case &H205: msg = "WM_RBUTTONUP"
      Case &H206: msg = "WM_RBUTTONDBLCLK"
      Case &H207: msg = "WM_MBUTTONDOWN"
      Case &H208: msg = "WM_MBUTTONUP"
      Case &H209: msg = "WM_MBUTTONDBLCLK"
      Case &H20A: msg = "WM_MOUSEWHEEL"
      Case &H20B: msg = "WM_XBUTTONDOWN"
      Case &H20C: msg = "WM_XBUTTONUP"
      Case &H20D: msg = "WM_XBUTTONDBLCLK"
      Case &H210: msg = "WM_PARENTNOTIFY"
      Case &H211: msg = "WM_ENTERMENULOOP"
      Case &H212: msg = "WM_EXITMENULOOP"
      Case &H213: msg = "WM_NEXTMENU"
      Case &H214: msg = "WM_SIZING"
      Case &H215: msg = "WM_CAPTURECHANGED"
      Case &H216: msg = "WM_MOVING"
      Case &H218: msg = "WM_POWERBROADCAST"
      Case &H219: msg = "WM_DEVICECHANGE"
      Case &H220: msg = "WM_MDICREATE"
      Case &H221: msg = "WM_MDIDESTROY"
      Case &H222: msg = "WM_MDIACTIVATE"
      Case &H223: msg = "WM_MDIRESTORE"
      Case &H224: msg = "WM_MDINEXT"
      Case &H225: msg = "WM_MDIMAXIMIZE"
      Case &H226: msg = "WM_MDITILE"
      Case &H227: msg = "WM_MDICASCADE"
      Case &H228: msg = "WM_MDIICONARRANGE"
      Case &H229: msg = "WM_MDIGETACTIVE"
      Case &H230: msg = "WM_MDISETMENU"
      Case &H231: msg = "WM_ENTERSIZEMOVE"
      Case &H232: msg = "WM_EXITSIZEMOVE"
      Case &H233: msg = "WM_DROPFILES"
      Case &H234: msg = "WM_MDIREFRESHMENU"
      Case &H281: msg = "WM_IME_SETCONTEXT"
      Case &H282: msg = "WM_IME_NOTIFY"
      Case &H283: msg = "WM_IME_CONTROL"
      Case &H284: msg = "WM_IME_COMPOSITIONFULL"
      Case &H285: msg = "WM_IME_SELECT"
      Case &H286: msg = "WM_IME_CHAR"
      Case &H288: msg = "WM_IME_REQUEST"
      Case &H290: msg = "WM_IME_KEYDOWN"
      Case &H291: msg = "WM_IME_KEYUP"
      Case &H2A1: msg = "WM_MOUSEHOVER"
      Case &H2A3: msg = "WM_MOUSELEAVE"
      Case &H2A0: msg = "WM_NCMOUSEHOVER"
      Case &H2A2: msg = "WM_NCMOUSELEAVE"
      Case &H2B1: msg = "WM_WTSSESSION_CHANGE"
      Case &H2C0: msg = "WM_TABLET_FIRST"
      Case &H2DF: msg = "WM_TABLET_LAST"
      Case &H300: msg = "WM_CUT"
      Case &H301: msg = "WM_COPY"
      Case &H302: msg = "WM_PASTE"
      Case &H303: msg = "WM_CLEAR"
      Case &H304: msg = "WM_UNDO"
      Case &H305: msg = "WM_RENDERFORMAT"
      Case &H306: msg = "WM_RENDERALLFORMATS"
      Case &H307: msg = "WM_DESTROYCLIPBOARD"
      Case &H308: msg = "WM_DRAWCLIPBOARD"
      Case &H309: msg = "WM_PAINTCLIPBOARD"
      Case &H30A: msg = "WM_VSCROLLCLIPBOARD"
      Case &H30B: msg = "WM_SIZECLIPBOARD"
      Case &H30C: msg = "WM_ASKCBFORMATNAME"
      Case &H30D: msg = "WM_CHANGECBCHAIN"
      Case &H30E: msg = "WM_HSCROLLCLIPBOARD"
      Case &H30F: msg = "WM_QUERYNEWPALETTE"
      Case &H310: msg = "WM_PALETTEISCHANGING"
      Case &H311: msg = "WM_PALETTECHANGED"
      Case &H312: msg = "WM_HOTKEY"
      Case &H317: msg = "WM_PRINT"
      Case &H318: msg = "WM_PRINTCLIENT"
      Case &H319: msg = "WM_APPCOMMAND"
      Case &H31A: msg = "WM_THEMECHANGED"
      Case &H358: msg = "WM_HANDHELDFIRST"
      Case &H35F: msg = "WM_HANDHELDLAST"
      Case &H360: msg = "WM_AFXFIRST"
      Case &H37F: msg = "WM_AFXLAST"
      Case &H380: msg = "WM_PENWINFIRST"
      Case &H38F: msg = "WM_PENWINLAST"
      Case &H400: msg = "WM_USER"
      Case Else: msg = "&H" & Hex(nMsg)
   End Select
   GetMessageName = msg
End Function

Private Function GetAddresOfProc(nProcAddress As Long) As Long
    GetAddresOfProc = nProcAddress
End Function

'*** the three following functions determine IDE-States (Break and ShutDown)
Private Function InBreakMode() As Boolean
    Static InitDone As Boolean, VBAVersion As Long
    Const vbmRun& = 1, vbmBreak& = 2
    If Not InitDone Then
        InitDone = True
        VBAVersion = VBAEnvironment
    End If
    If VBAVersion = 5 Then InBreakMode = (EbModeVBA5 = vbmBreak)
    If VBAVersion = 6 Then InBreakMode = (EbModeVBA6 = vbmBreak)
End Function

Private Function IsResetting() As Boolean
    Static InitDone As Boolean, VBAVersion As Long, Result As Boolean
    If Not InitDone Then
        InitDone = True
        VBAVersion = VBAEnvironment
    End If
    If Not Result Then
        If VBAVersion = 5 Then Result = EbIsResettingVBA5
        If VBAVersion = 6 Then Result = EbIsResettingVBA6
    End If
    IsResetting = Result
End Function

Private Function VBAEnvironment() As Long
    Static Done As Boolean, Result As Long
    If Not Done Then
        Done = True
        If GetModuleHandle("vba5.dll") Then
            Result = 5
        ElseIf GetModuleHandle("vba6.dll") Then
            Result = 6
        End If
    End If
    VBAEnvironment = Result
End Function

