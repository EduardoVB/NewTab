Attribute VB_Name = "mBSSubclass"
Option Explicit

' This subclasser uses Windows Common Controls subclassing with the interface of vbAccelerator's subtimer (AttachMessage/DettachMessage) that has the ability of only sending to the subclassed objects the messages that they need to handle (and not all window messages)

#Const IDE_PROTECTION_ENABLED = 1 ' you can set it to 0 if you compile it into an OCX or EXE if you want, but anyway it shouldn't be neccesary since the IDE protection code won't get compiled anyway
' IDE protection watchs for project resetting (Stop button), compiling and UserControls zombie states to remove all installed subclasses when they happen
' This IDE protection is needed only when the subclasser runs in source code
' It does not use any thunk when compiled, but uses some ASM thunks for IDE protection (this code does not get into the compiled program)

#If IDE_PROTECTION_ENABLED Then
Private Type IMAGE_DATA_DIRECTORY
    VirtualAddress As Long
    Size As Long
End Type

Private Type IMAGE_OPTIONAL_HEADER32
    Magic As Integer
    MajorLinkerVersion As Byte
    MinorLinkerVersion As Byte
    SizeOfCode As Long
    SizeOfInitalizedData As Long
    SizeOfUninitalizedData As Long
    AddressOfEntryPoint As Long
    BaseOfCode As Long
    BaseOfData As Long
    ImageBase As Long
    SectionAlignment As Long
    FileAlignment As Long
    MajorOperatingSystemVer As Integer
    MinorOperatingSystemVer As Integer
    MajorImageVersion As Integer
    MinorImageVersion As Integer
    MajorSubsystemVersion As Integer
    MinorSubsystemVersion As Integer
    Reserved1 As Long
    SizeOfImage As Long
    SizeOfHeaders As Long
    CheckSum As Long
    Subsystem As Integer
    DllCharacteristics As Integer
    SizeOfStackReserve As Long
    SizeOfStackCommit As Long
    SizeOfHeapReserve As Long
    SizeOfHeapCommit As Long
    LoaderFlags As Long
    NumberOfRvaAndSizes As Long
    DataDirectory(15) As IMAGE_DATA_DIRECTORY
End Type

Private Type IMAGE_DOS_HEADER
    e_magic As Integer
    e_cblp As Integer
    e_cp As Integer
    e_crlc As Integer
    e_cparhdr As Integer
    e_minalloc As Integer
    e_maxalloc As Integer
    e_ss As Integer
    e_sp As Integer
    e_csum As Integer
    e_ip As Integer
    e_cs As Integer
    e_lfarlc As Integer
    e_onvo As Integer
    e_res(0 To 3) As Integer
    e_oemid As Integer
    e_oeminfo As Integer
    e_res2(0 To 9) As Integer
    e_lfanew As Long
End Type
 
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CallWindowProcW Lib "user32.dll" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByRef lpAddress As Long, ByVal dwSize As Long, ByVal flAllocType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32.dll" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Function VirtualProtect Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flNewProtect As Long, ByRef lpflOldProtect As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function EnumThreadWindows Lib "user32" (ByVal dwThreadId As Long, ByVal lpfn As Long, ByVal lParam As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long

'Private Const GW_HWNDFIRST As Long = 0
Private Const GW_HWNDNEXT As Long = 2
Private Const GW_CHILD As Long = 5

#End If

Public Enum EMsgResponse
    emrConsume = 0       ' Process instead of original WindowProc
    emrPreprocess = 1    ' Process before original WindowProc
    emrPostProcess = 2  ' Process after original WindowProc
End Enum

#Const USE_ONLY_LOCAL_DB = 0 ' whether not to use SetProp/GetProp/RemoveProp in the subclasser and instead just use a local collection (does not affects the IDE protection code use of SetProp/GetProp/RemoveProp)

' ====================================================================================== Original vbAccelerator's module headers:
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
' ====================================================================================== End of vbAccelerator's module headers

' Note: it is a completely modified version.
' Date: April 8, 2021
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
Public gSubclassWithSetWindowLong As Boolean

#If IDE_PROTECTION_ENABLED Then
Private mIDEProtectionInitialized As Boolean
Private mCodeWindowsToWatch As Collection
Private mObjSubclassed As Collection
Private mObjSubclassed_CallCount As Collection
Private mCodeWindowsSubclassedHwnds As Collection
Private mAddressOf_CodeWindowWindowProc As Long
Private mCompiling As Boolean
Private mIDEIsResetting As Boolean
Private mTimerFindCodeWindowsHandle As Long
Private mAllSubclassesRemoved As Boolean
Private mIDEMainHwnd As Long
#End If

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

Private Property Get OldWindowProc(ByVal hWnd As Long) As Long
    Dim sName As String
    
    sName = hWnd
    OldWindowProc = ThisGetProp(hWnd, sName)
End Property

Private Property Let OldWindowProc(ByVal hWnd As Long, ByVal lPtr As Long)
    Dim sName As String
    
    m_f = 1
    sName = hWnd
    m_f = ThisSetProp(hWnd, sName, lPtr)
    If (lPtr = 0) Then
        ThisRemoveProp hWnd, sName
    End If
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

Sub AttachMessage(iwp As IBSSubclass, ByVal hWnd As Long, ByVal iMsg As Long)
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

#If IDE_PROTECTION_ENABLED Then
    Dim iInIDE As Boolean
    
    Debug.Assert MakeTrue(iInIDE)
    If iInIDE Then
        Dim iStr As String
        
        InitializeIDEProtection
        iStr = TypeName(iwp)
        If Not CodeWindowToWatchExists(iStr) Then
            mCodeWindowsToWatch.Add iStr, iStr
            ' find if the code window of this object is open and in this case subclass it
            EnumThreadWindows App.ThreadID, AddressOf EnumCodeWindowsCallback, 0
        End If
        If Not mObjSubclassed Is Nothing Then
            On Error Resume Next
            iStr = CStr(ObjPtr(iwp))
            mObjSubclassed.Add ObjPtr(iwp), iStr
            iLng = 0
            iLng = mObjSubclassed_CallCount(iStr)
            If iLng = 0 Then
                mObjSubclassed_CallCount.Add 1, iStr
            Else
                mObjSubclassed_CallCount.Remove iStr
                mObjSubclassed_CallCount.Add iLng + 1, iStr
            End If
            On Error GoTo 0
        End If
        
    End If
#End If
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
        If gSubclassWithSetWindowLong Then
            Dim iProcOld As Long

            ' Subclass window by installing window procedure
            iProcOld = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WindowProcSWL)
            If iProcOld = 0 Then
               ' remove class:
               MessageClass(hWnd, iMsg, MessageClassCount(hWnd, iMsg)) = 0
               ' remove class count:
               MessageClassCount(hWnd, iMsg) = MessageClassCount(hWnd, iMsg) - 1
               
               ErrRaise eeCantSubclass
               Exit Sub
            End If
            
            ' Associate old procedure with handle
            OldWindowProc(hWnd) = iProcOld
            If m_f = 0 Then
               ' SPM: Failed to VBThisSetProp, windows properties database problem.
               ' Has to be out of memory.
               
               ' Put the old window proc back again:
               SetWindowLong hWnd, GWL_WNDPROC, iProcOld
               ' remove class:
               MessageClass(hWnd, iMsg, MessageClassCount(hWnd, iMsg)) = 0
               ' remove class count:
               MessageClassCount(hWnd, iMsg) = MessageClassCount(hWnd, iMsg) - 1
               
               ' Raise an error:
               ErrRaise 5
               Exit Sub
            End If
        Else
            If SetWindowSubclass(hWnd, AddressOf WindowProcSWS, ObjPtr(iwp), 0&) = 0 Then
                ' remove class:
                MessageClass(hWnd, iMsg, MessageClassCount(hWnd, iMsg)) = 0
                ' remove class count:
                MessageClassCount(hWnd, iMsg) = MessageClassCount(hWnd, iMsg) - 1
                
                ErrRaise eeCantSubclass
                Exit Sub
            Else
                If mAddressOfWindowProc = 0 Then
                    mAddressOfWindowProc = GetAddresOfProc(AddressOf WindowProcSWS)
                End If
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
            If gSubclassWithSetWindowLong Then

                iProcOld = OldWindowProc(hWnd)
                If Not (iProcOld = 0) Then
                   SetWindowLong hWnd, GWL_WNDPROC, iProcOld
                   OldWindowProc(hWnd) = 0
                End If
            Else
                RemoveWindowSubclass hWnd, mAddressOfWindowProc, ObjPtr(iwp)
            End If
        End If
        
        ' Raise the error:
        ErrRaise 5
        Exit Sub
    End If
End Sub

Sub DetachMessage(iwp As IBSSubclass, ByVal hWnd As Long, ByVal iMsg As Long)
    Dim msgClassCount As Long
    Dim msgClass As Long
    Dim msgClassIndex As Long
    Dim msgCount As Long
    Dim iLng As Long
    
    #If IDE_PROTECTION_ENABLED Then
        Dim iInIDE As Boolean
        
        'If mAllSubclassesRemoved Then Exit Sub
        Debug.Assert MakeTrue(iInIDE)
        If iInIDE Then
            If Not mObjSubclassed Is Nothing Then
                Dim iStr As String
                
                iLng = 0
                iStr = CStr(ObjPtr(iwp))
                On Error Resume Next
                iLng = mObjSubclassed_CallCount(iStr)
                'Debug.Print "DetachMessage " & iStr & ", " & hWnd & " Msg: " & GetMessageName(iMsg) & ": " & mObjSubclassed_CallCount(iStr)
                If iLng = 1 Then
                    mObjSubclassed.Remove iStr
                    mObjSubclassed_CallCount.Remove iStr
                Else
                    mObjSubclassed_CallCount.Remove iStr
                    mObjSubclassed_CallCount.Add iLng - 1, iStr
                End If
                On Error GoTo 0
                'Debug.Print mObjSubclassed.Count
                If Not mObjSubclassed Is Nothing Then
                    If mObjSubclassed.Count = 0 Then
                        RemoveAllSubclasses ' for code windows
                        TerminateIDEProtection
                    End If
                End If
            End If
        End If
    #End If
    
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
        If gSubclassWithSetWindowLong Then
            Dim iProcOld As Long
            
            iProcOld = OldWindowProc(hWnd)
            If Not (iProcOld = 0) Then
               SetWindowLong hWnd, GWL_WNDPROC, iProcOld
               OldWindowProc(hWnd) = 0
            End If
        Else
            RemoveWindowSubclass hWnd, mAddressOfWindowProc, ObjPtr(iwp)
        End If
    End If
    MessageCount(hWnd) = MessageCount(hWnd) - 1

End Sub

Private Function WindowProcSWS(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
    
    Dim bCalled As Boolean
    Dim pSubClass As Long
    Dim iwp As IBSSubclass
    Dim iwpT As IBSSubclass
    Dim iIndex As Long
    Dim iHandled As Boolean
    Dim bConsume As Boolean
    Dim iResp As Long
    Dim iInIDE As Boolean
    
#If IDE_PROTECTION_ENABLED Then
    If mCompiling Then
        Debug.Assert MakeTrue(iInIDE)
        If iInIDE Then
            RemoveAllSubclasses
            mCompiling = False
            Exit Function
        End If
    ElseIf mAllSubclassesRemoved Then
       pClearUp hWnd, uIdSubclass
       Exit Function
    End If
#End If
    
    If IsResetting Then ' this runs when it is compiled into an OCX or DLL but is running in the IDE
        pClearUp hWnd, uIdSubclass
'        #If IDE_PROTECTION_ENABLED Then
'            Debug.Assert MakeTrue(iInIDE)
'            If iInIDE Then
'                TerminateIDEProtection
'                RemoveAllSubclasses
'            End If
'        #End If
        Exit Function
    End If
    
    If IsWindow(hWnd) = 0 Then
        pClearUp hWnd, uIdSubclass
        Exit Function
    End If
    If InBreakMode Then
        WindowProcSWS = DefSubclassProc(hWnd, iMsg, wParam, lParam)
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
                        WindowProcSWS = .WindowProc(hWnd, iMsg, wParam, lParam, bConsume)
                        
                        If Not bConsume Then
                            If (iIndex = 1) Then
                                If Not (bCalled) Then
                                    WindowProcSWS = DefSubclassProc(hWnd, iMsg, wParam, ByVal lParam)
                                    bCalled = True
                                End If
                            End If
                        End If
                        On Error GoTo 0
                    Else
                        ' Consume (this message is always passed to all control
                        ' instances regardless of whether any single one of them
                        ' requests to consume it):
                        WindowProcSWS = .WindowProc(hWnd, iMsg, wParam, lParam, bConsume)
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
                    WindowProcSWS = DefSubclassProc(hWnd, iMsg, wParam, ByVal lParam)
                    bCalled = True
                End If
            End If
        End If
        
        If Not iHandled Then
            WindowProcSWS = DefSubclassProc(hWnd, iMsg, wParam, ByVal lParam)
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
                WindowProcSWS = DefSubclassProc(hWnd, iMsg, wParam, ByVal lParam)
                pClearUp hWnd, uIdSubclass
            Else ' we are not a the top subclassing chain
                WindowProcSWS = DefSubclassProc(hWnd, iMsg, wParam, ByVal lParam)  ' let's see if the other subclass unsubclass itself
                If GetWindowLong(hWnd, GWL_WNDPROC) = mAddressOfWindowProc Then ' it did
                    pClearUp hWnd, uIdSubclass
                Else
                    If (iMsg = WM_NCDESTROY) Or (iMsg = WM_UAHDESTROYWINDOW) Then ' in these cases we will unsubclass anyway, but for WM_DESTROY we will wait for the WM_NCDESTROY message
                        pClearUp hWnd, uIdSubclass
                    End If
                End If
            End If
        Else
            WindowProcSWS = DefSubclassProc(hWnd, iMsg, wParam, ByVal lParam)
        End If
    End If
    
TheExit:
    Err.Clear
End Function

Private Function WindowProcSWL(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    Dim bCalled As Boolean
    Dim pSubClass As Long
    Dim iwp As IBSSubclass
    Dim iwpT As IBSSubclass
    Dim iIndex As Long
    Dim iHandled As Boolean
    Dim bConsume As Boolean
    Dim iResp As Long
    Dim iInIDE As Boolean
    
#If IDE_PROTECTION_ENABLED Then
    If mCompiling Then
        Debug.Assert MakeTrue(iInIDE)
        If iInIDE Then
            RemoveAllSubclasses
            mCompiling = False
            Exit Function
        End If
    ElseIf mAllSubclassesRemoved Then
       pClearUp hWnd
       Exit Function
    End If
#End If
    
    If IsResetting Then ' this runs when it is compiled into an OCX or DLL but is running in the IDE
        pClearUp hWnd
'        #If IDE_PROTECTION_ENABLED Then
'            Debug.Assert MakeTrue(iInIDE)
'            If iInIDE Then
'                TerminateIDEProtection
'                RemoveAllSubclasses
'            End If
'        #End If
        Exit Function
    End If
    
    If IsWindow(hWnd) = 0 Then
        pClearUp hWnd
        Exit Function
    End If
    If InBreakMode Then
        WindowProcSWL = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
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
                        WindowProcSWL = .WindowProc(hWnd, iMsg, wParam, lParam, bConsume)
                        
                        If Not bConsume Then
                            If (iIndex = 1) Then
                                If Not (bCalled) Then
                                    WindowProcSWL = CallOldWindowProc(hWnd, iMsg, wParam, ByVal lParam)
                                    bCalled = True
                                End If
                            End If
                        End If
                        On Error GoTo 0
                    Else
                        ' Consume (this message is always passed to all control
                        ' instances regardless of whether any single one of them
                        ' requests to consume it):
                        WindowProcSWL = .WindowProc(hWnd, iMsg, wParam, lParam, bConsume)
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
                    WindowProcSWL = CallOldWindowProc(hWnd, iMsg, wParam, ByVal lParam)
                    bCalled = True
                End If
            End If
        End If
        
        If Not iHandled Then
            WindowProcSWL = CallOldWindowProc(hWnd, iMsg, wParam, ByVal lParam)
            If GetWindowLong(hWnd, GWL_WNDPROC) = mAddressOfWindowProc Then     ' if we are at the top of the subclassing chain, else we'll wait for the WM_DESTROY, WM_NCDESTROY and WM_UAHDESTROYWINDOW messages
                pClearUp hWnd
            End If
        End If
    Else
        ' Not handled:
        If (iMsg = WM_DESTROY) Or (iMsg = WM_NCDESTROY) Or (iMsg = WM_UAHDESTROYWINDOW) Then
            ' If WM_DESTROY isn't handled already, we should
            ' clear up any subclass
            If GetWindowLong(hWnd, GWL_WNDPROC) = mAddressOfWindowProc Then ' if we are at the top of the subclassing chain
                WindowProcSWL = CallOldWindowProc(hWnd, iMsg, wParam, ByVal lParam)
                pClearUp hWnd
            Else ' we are not a the top subclassing chain
                WindowProcSWL = CallOldWindowProc(hWnd, iMsg, wParam, ByVal lParam)  ' let's see if the other subclass unsubclass itself
                If GetWindowLong(hWnd, GWL_WNDPROC) = mAddressOfWindowProc Then ' it did
                    pClearUp hWnd
                Else
                    If (iMsg = WM_NCDESTROY) Or (iMsg = WM_UAHDESTROYWINDOW) Then ' in these cases we will unsubclass anyway, but for WM_DESTROY we will wait for the WM_NCDESTROY message
                        pClearUp hWnd
                    End If
                End If
            End If
        Else
            WindowProcSWL = CallOldWindowProc(hWnd, iMsg, wParam, ByVal lParam)
        End If
    End If
    
TheExit:
    Err.Clear
End Function

Public Function CallOldWindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If gSubclassWithSetWindowLong Then
        Dim iProcOld As Long
        
        iProcOld = OldWindowProc(hWnd)
        If Not (iProcOld = 0) Then
           CallOldWindowProc = CallWindowProc(iProcOld, hWnd, iMsg, wParam, lParam)
        End If
    Else
        CallOldWindowProc = DefSubclassProc(hWnd, iMsg, wParam, lParam)
    End If
End Function

Private Function IsWindowLocal(ByVal hWnd As Long) As Boolean
    Dim idWnd As Long
    
    Call GetWindowThreadProcessId(hWnd, idWnd)
    IsWindowLocal = (idWnd = GetCurrentProcessId())
End Function

'Private Sub logMessage(ByVal sMsg As String)
'    Debug.Print sMsg
'End Sub


Private Sub pClearUp(ByVal hWnd As Long, Optional uIdSubclass As Long)
    Dim msgCount As Long
    
    ' this is only called if you haven't explicitly cleared up
    ' your subclass from the caller.  You will get a minor
    ' resource leak as it does not clear up any message
    ' specific properties.
    msgCount = MessageCount(hWnd)
    If (msgCount > 0) Then
        ' remove the subclass:
        ' Unsubclass
        
        If gSubclassWithSetWindowLong Then
            Dim iProcOld As Long
            
            iProcOld = OldWindowProc(hWnd)
            If Not (iProcOld = 0) Then
               SetWindowLong hWnd, GWL_WNDPROC, iProcOld
               OldWindowProc(hWnd) = 0
            End If
        Else
            RemoveWindowSubclass hWnd, mAddressOfWindowProc, uIdSubclass
        End If
        ' remove the old window proc:
        MessageCount(hWnd) = 0
    End If
End Sub

Private Function ThisGetProp(ByVal hWnd As Long, ByVal lpString As String) As Long
#If USE_ONLY_LOCAL_DB Then
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
#If USE_ONLY_LOCAL_DB Then
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
#If USE_ONLY_LOCAL_DB Then
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


'Private Function InIDE() As Boolean
'    Static sValue As Long
'
'    If sValue = 0 Then
'        On Error Resume Next
'        Err.Clear
'        Debug.Assert "a"
'        If Err.Number = 13 Then
'            sValue = 1
'        Else
'            sValue = 2
'        End If
'        Err.Clear
'    End If
'    InIDE = sValue = 1
'End Function

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

Private Function GetAddresOfProc(nProcAddress As Long) As Long
    GetAddresOfProc = nProcAddress
End Function

'*** the three following functions determine IDE-States (Break and ShutDown)
Private Function InBreakMode() As Boolean
    Dim iInIDE As Boolean
    
    Debug.Assert MakeTrue(iInIDE)
    If iInIDE Then
        Static InitDone As Boolean, VBAVersion As Long
        Const vbmBreak& = 2
        If Not InitDone Then
            InitDone = True
            VBAVersion = VBAEnvironment
        End If
        If VBAVersion = 5 Then InBreakMode = (EbModeVBA5 = vbmBreak)
        If VBAVersion = 6 Then InBreakMode = (EbModeVBA6 = vbmBreak)
    End If
End Function

Private Function IsResetting() As Boolean
    Dim iInIDE As Boolean
    
    Debug.Assert MakeTrue(iInIDE)
    If iInIDE Then
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
    End If
End Function

Private Function VBAEnvironment() As Long
    Dim iInIDE As Boolean
    
    Debug.Assert MakeTrue(iInIDE)
    If iInIDE Then
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
    End If
End Function

#If False Then

Private Function GetMessageName(nMsg As Long) As String
   Dim Msg As String
   
   Select Case nMsg
      Case &H0: Msg = "WM_NULL"
      Case &H1: Msg = "WM_CREATE"
      Case &H2: Msg = "WM_DESTROY"
      Case &H3: Msg = "WM_MOVE"
      Case &H5: Msg = "WM_SIZE"
      Case &H6: Msg = "WM_ACTIVATE"
      Case &H7: Msg = "WM_SETFOCUS"
      Case &H8: Msg = "WM_KILLFOCUS"
      Case &HA: Msg = "WM_ENABLE"
      Case &HB: Msg = "WM_SETREDRAW"
      Case &HC: Msg = "WM_SETTEXT"
      Case &HD: Msg = "WM_GETTEXT"
      Case &HE: Msg = "WM_GETTEXTLENGTH"
      Case &HF: Msg = "WM_PAINT"
      Case &H10: Msg = "WM_CLOSE"
      Case &H11: Msg = "WM_QUERYENDSESSION"
      Case &H12: Msg = "WM_QUIT"
      Case &H13: Msg = "WM_QUERYOPEN"
      Case &H14: Msg = "WM_ERASEBKGND"
      Case &H15: Msg = "WM_SYSCOLORCHANGE"
      Case &H16: Msg = "WM_ENDSESSION"
      Case &H18: Msg = "WM_SHOWWINDOW"
      Case &H1A: Msg = "WM_SETTINGCHANGE"
      Case &H1B: Msg = "WM_DEVMODECHANGE"
      Case &H1C: Msg = "WM_ACTIVATEAPP"
      Case &H1D: Msg = "WM_FONTCHANGE"
      Case &H1E: Msg = "WM_TIMECHANGE"
      Case &H1F: Msg = "WM_CANCELMODE"
      Case &H20: Msg = "WM_SETCURSOR"
      Case &H21: Msg = "WM_MOUSEACTIVATE"
      Case &H22: Msg = "WM_CHILDACTIVATE"
      Case &H23: Msg = "WM_QUEUESYNC"
      Case &H24: Msg = "WM_GETMINMAXINFO"
      Case &H26: Msg = "WM_PAINTICON"
      Case &H27: Msg = "WM_ICONERASEBKGND"
      Case &H28: Msg = "WM_NEXTDLGCTL"
      Case &H2A: Msg = "WM_SPOOLERSTATUS"
      Case &H2B: Msg = "WM_DRAWITEM"
      Case &H2C: Msg = "WM_MEASUREITEM"
      Case &H2D: Msg = "WM_DELETEITEM"
      Case &H2E: Msg = "WM_VKEYTOITEM"
      Case &H2F: Msg = "WM_CHARTOITEM"
      Case &H30: Msg = "WM_SETFONT"
      Case &H31: Msg = "WM_GETFONT"
      Case &H32: Msg = "WM_SETHOTKEY"
      Case &H33: Msg = "WM_GETHOTKEY"
      Case &H37: Msg = "WM_QUERYDRAGICON"
      Case &H39: Msg = "WM_COMPAREITEM"
      Case &H3D: Msg = "WM_GETOBJECT"
      Case &H41: Msg = "WM_COMPACTING"
      Case &H44: Msg = "WM_COMMNOTIFY"
      Case &H46: Msg = "WM_WINDOWPOSCHANGING"
      Case &H47: Msg = "WM_WINDOWPOSCHANGED"
      Case &H48: Msg = "WM_POWER"
      Case &H4A: Msg = "WM_COPYDATA"
      Case &H4B: Msg = "WM_CANCELJOURNAL"
      Case &H4E: Msg = "WM_NOTIFY"
      Case &H50: Msg = "WM_INPUTLANGCHANGEREQUEST"
      Case &H51: Msg = "WM_INPUTLANGCHANGE"
      Case &H52: Msg = "WM_TCARD"
      Case &H53: Msg = "WM_HELP"
      Case &H54: Msg = "WM_USERCHANGED"
      Case &H55: Msg = "WM_NOTIFYFORMAT"
      Case &H7B: Msg = "WM_CONTEXTMENU"
      Case &H7C: Msg = "WM_STYLECHANGING"
      Case &H7D: Msg = "WM_STYLECHANGED"
      Case &H7E: Msg = "WM_DISPLAYCHANGE"
      Case &H7F: Msg = "WM_GETICON"
      Case &H80: Msg = "WM_SETICON"
      Case &H81: Msg = "WM_NCCREATE"
      Case &H82: Msg = "WM_NCDESTROY"
      Case &H83: Msg = "WM_NCCALCSIZE"
      Case &H84: Msg = "WM_NCHITTEST"
      Case &H85: Msg = "WM_NCPAINT"
      Case &H86: Msg = "WM_NCACTIVATE"
      Case &H87: Msg = "WM_GETDLGCODE"
      Case &H88: Msg = "WM_SYNCPAINT"
      Case &HA0: Msg = "WM_NCMOUSEMOVE"
      Case &HA1: Msg = "WM_NCLBUTTONDOWN"
      Case &HA2: Msg = "WM_NCLBUTTONUP"
      Case &HA3: Msg = "WM_NCLBUTTONDBLCLK"
      Case &HA4: Msg = "WM_NCRBUTTONDOWN"
      Case &HA5: Msg = "WM_NCRBUTTONUP"
      Case &HA6: Msg = "WM_NCRBUTTONDBLCLK"
      Case &HA7: Msg = "WM_NCMBUTTONDOWN"
      Case &HA8: Msg = "WM_NCMBUTTONUP"
      Case &HA9: Msg = "WM_NCMBUTTONDBLCLK"
      Case &HAB: Msg = "WM_NCXBUTTONDOWN"
      Case &HAC: Msg = "WM_NCXBUTTONUP"
      Case &HAD: Msg = "WM_NCXBUTTONDBLCLK"
      Case &HFF: Msg = "WM_INPUT"
      Case &H100: Msg = "WM_KEYDOWN"
      Case &H101: Msg = "WM_KEYUP"
      Case &H102: Msg = "WM_CHAR"
      Case &H103: Msg = "WM_DEADCHAR"
      Case &H104: Msg = "WM_SYSKEYDOWN"
      Case &H105: Msg = "WM_SYSKEYUP"
      Case &H106: Msg = "WM_SYSCHAR"
      Case &H107: Msg = "WM_SYSDEADCHAR"
      Case &H108: Msg = "WM_KEYLAST"
      Case &H10D: Msg = "WM_IME_STARTCOMPOSITION"
      Case &H10E: Msg = "WM_IME_ENDCOMPOSITION"
      Case &H10F: Msg = "WM_IME_COMPOSITION"
      Case &H110: Msg = "WM_INITDIALOG"
      Case &H111: Msg = "WM_COMMAND"
      Case &H112: Msg = "WM_SYSCOMMAND"
      Case &H113: Msg = "WM_TIMER"
      Case &H114: Msg = "WM_HSCROLL"
      Case &H115: Msg = "WM_VSCROLL"
      Case &H116: Msg = "WM_INITMENU"
      Case &H117: Msg = "WM_INITMENUPOPUP"
      Case &H11F: Msg = "WM_MENUSELECT"
      Case &H120: Msg = "WM_MENUCHAR"
      Case &H121: Msg = "WM_ENTERIDLE"
      Case &H122: Msg = "WM_MENURBUTTONUP"
      Case &H123: Msg = "WM_MENUDRAG"
      Case &H124: Msg = "WM_MENUGETOBJECT"
      Case &H125: Msg = "WM_UNINITMENUPOPUP"
      Case &H126: Msg = "WM_MENUCOMMAND"
      Case &H127: Msg = "WM_CHANGEUISTATE"
      Case &H128: Msg = "WM_UPDATEUISTATE"
      Case &H129: Msg = "WM_QUERYUISTATE"
      Case &H132: Msg = "WM_CTLCOLORMSGBOX"
      Case &H133: Msg = "WM_CTLCOLOREDIT"
      Case &H134: Msg = "WM_CTLCOLORLISTBOX"
      Case &H135: Msg = "WM_CTLCOLORBTN"
      Case &H136: Msg = "WM_CTLCOLORDLG"
      Case &H137: Msg = "WM_CTLCOLORSCROLLBAR"
      Case &H138: Msg = "WM_CTLCOLORSTATIC"
      Case &H1E1: Msg = "MN_GETHMENU"
'      Case &H200: msg = "WM_MOUSEFIRST"
      Case &H200: Msg = "WM_MOUSEMOVE"
      Case &H201: Msg = "WM_LBUTTONDOWN"
      Case &H202: Msg = "WM_LBUTTONUP"
      Case &H203: Msg = "WM_LBUTTONDBLCLK"
      Case &H204: Msg = "WM_RBUTTONDOWN"
      Case &H205: Msg = "WM_RBUTTONUP"
      Case &H206: Msg = "WM_RBUTTONDBLCLK"
      Case &H207: Msg = "WM_MBUTTONDOWN"
      Case &H208: Msg = "WM_MBUTTONUP"
      Case &H209: Msg = "WM_MBUTTONDBLCLK"
      Case &H20A: Msg = "WM_MOUSEWHEEL"
      Case &H20B: Msg = "WM_XBUTTONDOWN"
      Case &H20C: Msg = "WM_XBUTTONUP"
      Case &H20D: Msg = "WM_XBUTTONDBLCLK"
      Case &H210: Msg = "WM_PARENTNOTIFY"
      Case &H211: Msg = "WM_ENTERMENULOOP"
      Case &H212: Msg = "WM_EXITMENULOOP"
      Case &H213: Msg = "WM_NEXTMENU"
      Case &H214: Msg = "WM_SIZING"
      Case &H215: Msg = "WM_CAPTURECHANGED"
      Case &H216: Msg = "WM_MOVING"
      Case &H218: Msg = "WM_POWERBROADCAST"
      Case &H219: Msg = "WM_DEVICECHANGE"
      Case &H220: Msg = "WM_MDICREATE"
      Case &H221: Msg = "WM_MDIDESTROY"
      Case &H222: Msg = "WM_MDIACTIVATE"
      Case &H223: Msg = "WM_MDIRESTORE"
      Case &H224: Msg = "WM_MDINEXT"
      Case &H225: Msg = "WM_MDIMAXIMIZE"
      Case &H226: Msg = "WM_MDITILE"
      Case &H227: Msg = "WM_MDICASCADE"
      Case &H228: Msg = "WM_MDIICONARRANGE"
      Case &H229: Msg = "WM_MDIGETACTIVE"
      Case &H230: Msg = "WM_MDISETMENU"
      Case &H231: Msg = "WM_ENTERSIZEMOVE"
      Case &H232: Msg = "WM_EXITSIZEMOVE"
      Case &H233: Msg = "WM_DROPFILES"
      Case &H234: Msg = "WM_MDIREFRESHMENU"
      Case &H281: Msg = "WM_IME_SETCONTEXT"
      Case &H282: Msg = "WM_IME_NOTIFY"
      Case &H283: Msg = "WM_IME_CONTROL"
      Case &H284: Msg = "WM_IME_COMPOSITIONFULL"
      Case &H285: Msg = "WM_IME_SELECT"
      Case &H286: Msg = "WM_IME_CHAR"
      Case &H288: Msg = "WM_IME_REQUEST"
      Case &H290: Msg = "WM_IME_KEYDOWN"
      Case &H291: Msg = "WM_IME_KEYUP"
      Case &H2A1: Msg = "WM_MOUSEHOVER"
      Case &H2A3: Msg = "WM_MOUSELEAVE"
      Case &H2A0: Msg = "WM_NCMOUSEHOVER"
      Case &H2A2: Msg = "WM_NCMOUSELEAVE"
      Case &H2B1: Msg = "WM_WTSSESSION_CHANGE"
      Case &H2C0: Msg = "WM_TABLET_FIRST"
      Case &H2DF: Msg = "WM_TABLET_LAST"
      Case &H300: Msg = "WM_CUT"
      Case &H301: Msg = "WM_COPY"
      Case &H302: Msg = "WM_PASTE"
      Case &H303: Msg = "WM_CLEAR"
      Case &H304: Msg = "WM_UNDO"
      Case &H305: Msg = "WM_RENDERFORMAT"
      Case &H306: Msg = "WM_RENDERALLFORMATS"
      Case &H307: Msg = "WM_DESTROYCLIPBOARD"
      Case &H308: Msg = "WM_DRAWCLIPBOARD"
      Case &H309: Msg = "WM_PAINTCLIPBOARD"
      Case &H30A: Msg = "WM_VSCROLLCLIPBOARD"
      Case &H30B: Msg = "WM_SIZECLIPBOARD"
      Case &H30C: Msg = "WM_ASKCBFORMATNAME"
      Case &H30D: Msg = "WM_CHANGECBCHAIN"
      Case &H30E: Msg = "WM_HSCROLLCLIPBOARD"
      Case &H30F: Msg = "WM_QUERYNEWPALETTE"
      Case &H310: Msg = "WM_PALETTEISCHANGING"
      Case &H311: Msg = "WM_PALETTECHANGED"
      Case &H312: Msg = "WM_HOTKEY"
      Case &H317: Msg = "WM_PRINT"
      Case &H318: Msg = "WM_PRINTCLIENT"
      Case &H319: Msg = "WM_APPCOMMAND"
      Case &H31A: Msg = "WM_THEMECHANGED"
      Case &H358: Msg = "WM_HANDHELDFIRST"
      Case &H35F: Msg = "WM_HANDHELDLAST"
      Case &H360: Msg = "WM_AFXFIRST"
      Case &H37F: Msg = "WM_AFXLAST"
      Case &H380: Msg = "WM_PENWINFIRST"
      Case &H38F: Msg = "WM_PENWINLAST"
      Case &H400: Msg = "WM_USER"
      Case Else: Msg = "&H" & Hex$(nMsg)
   End Select
   GetMessageName = Msg
End Function

#End If

#If IDE_PROTECTION_ENABLED Then

Private Function GetIDEMainHwnd() As Long
    If mIDEMainHwnd = 0 Then EnumThreadWindows App.ThreadID, AddressOf EnumThreadProc_GetIDEMainWindow, 0&
    GetIDEMainHwnd = mIDEMainHwnd
End Function

Private Function MakeTrue(Value As Boolean) As Boolean
    MakeTrue = True
    Value = True
End Function

Private Function IsHwndOfCodeWindowWatched(nHwnd As Long) As Boolean
    Dim v As Variant
    Dim iCaption As String
    Dim iAppEXEName As String
    
    If Not mCodeWindowsToWatch Is Nothing Then
        iAppEXEName = App.EXEName
        iCaption = GetWindowCaption(nHwnd)
        For Each v In mCodeWindowsToWatch
            If Left$(iCaption, Len(iAppEXEName & " - " & v)) = iAppEXEName & " - " & v Then
                IsHwndOfCodeWindowWatched = True
                Exit Function
            ElseIf Left$(iCaption, Len(v) + 1) = v & " " Then
                IsHwndOfCodeWindowWatched = True
                Exit Function
            End If
        Next
    End If
End Function

Private Function GetWindowClassName(nHwnd As Long) As String
    Dim iClassName As String
    Dim iSize As Long
    
    If nHwnd = 0 Then Exit Function
    
    iClassName = Space$(64)
    iSize = GetClassName(nHwnd, iClassName, Len(iClassName))
    GetWindowClassName = Left$(iClassName, iSize)
End Function

Private Function GetWindowCaption(nHwnd As Long) As String
    Dim iWinCaption As String
    Dim iRet As Long
    
    iWinCaption = String$(255, 0)
    iRet = GetWindowText(nHwnd, iWinCaption, 255)
    If iRet > 0 Then
        GetWindowCaption = Left$(iWinCaption, iRet)
    End If
End Function

Private Function CodeWindowToWatchExists(nName As String) As Boolean
    Dim iStr As String
    
    On Error GoTo ErrH
    iStr = mCodeWindowsToWatch(nName)
    CodeWindowToWatchExists = True
    Exit Function
    
ErrH:
    Err.Clear
End Function

Private Sub RemoveAllSubclasses()
    Dim v As Variant
    Dim o As Object
    Dim iwp As IBSSubclass
    
    If Not mCodeWindowsSubclassedHwnds Is Nothing Then
        For Each v In mCodeWindowsSubclassedHwnds
            UnSubClassCodeWindow CLng(v)
        Next
    End If
    
    If Not mObjSubclassed Is Nothing Then
        On Error Resume Next
        For Each v In mObjSubclassed
            CopyMemory o, CLng(v), 4
            Set iwp = o
            CopyMemory o, 0&, 4
            iwp.UnsubclassIt
        Next
        On Error GoTo 0
    End If
    mAllSubclassesRemoved = True
End Sub

Private Sub InitializeIDEProtection()
    Dim iInIDE As Boolean
    
    Debug.Assert MakeTrue(iInIDE)
    If iInIDE Then
        Dim iHwndIDEMain As Long
        Dim iOldProtect As Long
        Const MEM_COMMIT As Long = &H1000
        Const PAGE_EXECUTE_READWRITE As Long = &H40
        Const MEM_RELEASE As Long = &H8000&
        
        If mIDEProtectionInitialized Then Exit Sub
        mIDEProtectionInitialized = True
        mAllSubclassesRemoved = False
        
        Set mCodeWindowsToWatch = New Collection
        Set mCodeWindowsSubclassedHwnds = New Collection
        Set mObjSubclassed = New Collection
        Set mObjSubclassed_CallCount = New Collection
        
        ' watch code windows related to subclassing
        mCodeWindowsToWatch.Add "mBSSubclass", "mBSSubclass"
        mCodeWindowsToWatch.Add "mBSPropsDB", "mBSPropsDB"
        mCodeWindowsToWatch.Add "IBSSubclass", "IBSSubclass"
        ' find if there is anyone already one open and subclass it
        EnumThreadWindows App.ThreadID, AddressOf EnumCodeWindowsCallback, 0
        
        ' https://www.vbarchiv.net/tipps/tipp_1852.html
        ' https://www.vbforums.com/showthread.php?832275
        
        iHwndIDEMain = GetIDEMainHwnd
        If iHwndIDEMain <> 0 Then
            ' EbProjectReset
            Dim iIATEntryAddress_EbProjectReset As Long
            Dim iFuncProcAddress_EbProjectReset As Long
            Dim iLocalProcAddress_EbProjectReset As Long
            Dim iASMThunkAddress_EbProjectReset As Long

            iLocalProcAddress_EbProjectReset = GetAddresOfProc(AddressOf IDEIsResetting)
            iFuncProcAddress_EbProjectReset = GetProp(iHwndIDEMain, "EbPR.ProAd")
            If iFuncProcAddress_EbProjectReset = 0 Then
                iFuncProcAddress_EbProjectReset = GetProcAddress(GetModuleHandle("vba6.dll"), "EbProjectReset")
                SetProp iHwndIDEMain, "EbPR.ProAd", iFuncProcAddress_EbProjectReset
            End If
            iIATEntryAddress_EbProjectReset = GetProp(iHwndIDEMain, "EbPR.IATAd")
            If iIATEntryAddress_EbProjectReset = 0 Then
                iIATEntryAddress_EbProjectReset = GetIATEntryAddress("vb6.exe", iFuncProcAddress_EbProjectReset)
                SetProp iHwndIDEMain, "EbPR.IATAd", iIATEntryAddress_EbProjectReset
            End If
            iASMThunkAddress_EbProjectReset = GetProp(iHwndIDEMain, "EbPR.Thunk")
            If iASMThunkAddress_EbProjectReset <> 0 Then
                VirtualFree iASMThunkAddress_EbProjectReset, 0, MEM_RELEASE
                RemoveProp iHwndIDEMain, "EbPR.Thunk"
            End If
            If (iIATEntryAddress_EbProjectReset <> 0) And (iFuncProcAddress_EbProjectReset <> 0) Then
                iASMThunkAddress_EbProjectReset = VirtualAlloc(ByVal 0, 20, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
                If iASMThunkAddress_EbProjectReset <> 0 Then
                    
                    ' general (same protection for the whole memory block)
                    iOldProtect = GetProp(iHwndIDEMain, "IAT.OrigProt")
                    If iOldProtect = 0 Then
                        VirtualProtect iIATEntryAddress_EbProjectReset, 4, PAGE_EXECUTE_READWRITE, iOldProtect
                        SetProp iHwndIDEMain, "IAT.OrigProt", iOldProtect
                        SetProp iHwndIDEMain, "IAT.ProtAdd", iIATEntryAddress_EbProjectReset
                    End If
                    ' end general
                    
                    SetProp iHwndIDEMain, "EbPR.Thunk", iASMThunkAddress_EbProjectReset
                    CopyMemory ByVal iIATEntryAddress_EbProjectReset, iASMThunkAddress_EbProjectReset, 4
                    
                    ' call the local procedure
                    WriteCall iASMThunkAddress_EbProjectReset, iLocalProcAddress_EbProjectReset
                    ' restore original IAT entry
                    WriteByte iASMThunkAddress_EbProjectReset, &HC7 ' MOV
                    WriteByte iASMThunkAddress_EbProjectReset, &H5
                    WriteLong iASMThunkAddress_EbProjectReset, iIATEntryAddress_EbProjectReset
                    WriteLong iASMThunkAddress_EbProjectReset, iFuncProcAddress_EbProjectReset
                    ' jump to the original function address
                    WriteJump iASMThunkAddress_EbProjectReset, iFuncProcAddress_EbProjectReset
                End If
            Else
                RemoveProp iHwndIDEMain, "EbPR.ProAd"
                RemoveProp iHwndIDEMain, "EbPR.IATAd"
            End If
            
            ' TipStartMakeExe
            Dim iIATEntryAddress_TipStartMakeExe As Long
            Dim iFuncProcAddress_TipStartMakeExe As Long
            Dim iLocalProcAddress_TipStartMakeExe As Long
            Dim iASMThunkAddress_TipStartMakeExe As Long
            
            iLocalProcAddress_TipStartMakeExe = GetAddresOfProc(AddressOf IDEAboutToMakeExe)
            iFuncProcAddress_TipStartMakeExe = GetProp(iHwndIDEMain, "TSME.ProAd")
            If iFuncProcAddress_TipStartMakeExe = 0 Then
                iFuncProcAddress_TipStartMakeExe = GetProcAddress(GetModuleHandle("vba6.dll"), "TipStartMakeExe")
                SetProp iHwndIDEMain, "TSME.ProAd", iFuncProcAddress_TipStartMakeExe
            End If
            iIATEntryAddress_TipStartMakeExe = GetProp(iHwndIDEMain, "TSME.IATAd")
            If iIATEntryAddress_TipStartMakeExe = 0 Then
                iIATEntryAddress_TipStartMakeExe = GetIATEntryAddress("vb6.exe", iFuncProcAddress_TipStartMakeExe)
                SetProp iHwndIDEMain, "TSME.IATAd", iIATEntryAddress_TipStartMakeExe
            End If
            iASMThunkAddress_TipStartMakeExe = GetProp(iHwndIDEMain, "TSME.Thunk")
            If iASMThunkAddress_TipStartMakeExe <> 0 Then
                VirtualFree iASMThunkAddress_TipStartMakeExe, 0, MEM_RELEASE
                RemoveProp iHwndIDEMain, "TSME.Thunk"
            End If
            If (iIATEntryAddress_TipStartMakeExe <> 0) And (iFuncProcAddress_TipStartMakeExe <> 0) Then
                iASMThunkAddress_TipStartMakeExe = VirtualAlloc(ByVal 0, 20, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
                If iASMThunkAddress_TipStartMakeExe <> 0 Then
                    SetProp iHwndIDEMain, "TSME.Thunk", iASMThunkAddress_TipStartMakeExe
                    CopyMemory ByVal iIATEntryAddress_TipStartMakeExe, iASMThunkAddress_TipStartMakeExe, 4
                    
                    ' call the local procedure
                    WriteCall iASMThunkAddress_TipStartMakeExe, iLocalProcAddress_TipStartMakeExe
                    ' restore original IAT entry
                    WriteByte iASMThunkAddress_TipStartMakeExe, &HC7 ' MOV
                    WriteByte iASMThunkAddress_TipStartMakeExe, &H5
                    WriteLong iASMThunkAddress_TipStartMakeExe, iIATEntryAddress_TipStartMakeExe
                    WriteLong iASMThunkAddress_TipStartMakeExe, iFuncProcAddress_TipStartMakeExe
                    ' jump to the original function address
                    WriteJump iASMThunkAddress_TipStartMakeExe, iFuncProcAddress_TipStartMakeExe
                End If
            Else
                RemoveProp iHwndIDEMain, "TSME.ProAd"
                RemoveProp iHwndIDEMain, "TSME.IATAd"
            End If
        
            ' EbShowCode
            Dim iIATEntryAddress_EbShowCode As Long
            Dim iFuncProcAddress_EbShowCode As Long
            Dim iLocalProcAddress_EbShowCode As Long
            Dim iASMThunkAddress_EbShowCode As Long
            
            iLocalProcAddress_EbShowCode = GetAddresOfProc(AddressOf IDECodeWindowShowing)
            iFuncProcAddress_EbShowCode = GetProp(iHwndIDEMain, "EbSC.ProAd")
            If iFuncProcAddress_EbShowCode = 0 Then
                iFuncProcAddress_EbShowCode = GetProcAddress(GetModuleHandle("vba6.dll"), "EbShowCode")
                SetProp iHwndIDEMain, "EbSC.ProAd", iFuncProcAddress_EbShowCode
            End If
            iIATEntryAddress_EbShowCode = GetProp(iHwndIDEMain, "EbSC.IATAd")
            If iIATEntryAddress_EbShowCode = 0 Then
                iIATEntryAddress_EbShowCode = GetIATEntryAddress("vb6.exe", iFuncProcAddress_EbShowCode)
                SetProp iHwndIDEMain, "EbSC.IATAd", iIATEntryAddress_EbShowCode
            End If
            iASMThunkAddress_EbShowCode = GetProp(iHwndIDEMain, "EbSC.Thunk")
            If iASMThunkAddress_EbShowCode <> 0 Then
                VirtualFree iASMThunkAddress_EbShowCode, 0, MEM_RELEASE
                RemoveProp iHwndIDEMain, "EbSC.Thunk"
            End If
            If (iIATEntryAddress_EbShowCode <> 0) And (iFuncProcAddress_EbShowCode <> 0) Then
                iASMThunkAddress_EbShowCode = VirtualAlloc(ByVal 0, 10, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
                If iASMThunkAddress_EbShowCode <> 0 Then
                    SetProp iHwndIDEMain, "EbSC.Thunk", iASMThunkAddress_EbShowCode
                    CopyMemory ByVal iIATEntryAddress_EbShowCode, iASMThunkAddress_EbShowCode, 4
                    
                    ' call the local procedure
                    WriteCall iASMThunkAddress_EbShowCode, iLocalProcAddress_EbShowCode
                    ' restore original IAT entry (no, we will keep it, the size of the thunk changed to 10.
                    'WriteByte iASMThunkAddress_EbShowCode, &HC7 ' MOV
                    'WriteByte iASMThunkAddress_EbShowCode, &H5
                    'WriteLong iASMThunkAddress_EbShowCode, iIATEntryAddress_EbShowCode
                    'WriteLong iASMThunkAddress_EbShowCode, iFuncProcAddress_EbShowCode
                    ' jump to the original function address
                    WriteJump iASMThunkAddress_EbShowCode, iFuncProcAddress_EbShowCode
                End If
                
            Else
                RemoveProp iHwndIDEMain, "EbSC.ProAd"
                RemoveProp iHwndIDEMain, "EbSC.IATAd"
            End If
        End If
    End If
End Sub

Private Sub TerminateIDEProtection_EbProjectReset()
    Dim iInIDE As Boolean
    
    If mIDEIsResetting Then Exit Sub
    
    Debug.Assert MakeTrue(iInIDE)
    If iInIDE Then
        Dim iHwndIDEMain As Long
        Const MEM_RELEASE As Long = &H8000&
        
        iHwndIDEMain = GetIDEMainHwnd
        If iHwndIDEMain <> 0 Then
            Dim iFuncProcAddress_EbProjectReset As Long
            Dim iIATEntryAddress_EbProjectReset As Long
            Dim iASMThunkAddress_EbProjectReset As Long
            
            iFuncProcAddress_EbProjectReset = GetProp(iHwndIDEMain, "EbPR.ProAd")
            iIATEntryAddress_EbProjectReset = GetProp(iHwndIDEMain, "EbPR.IATAd")
            iASMThunkAddress_EbProjectReset = GetProp(iHwndIDEMain, "EbPR.Thunk")

            If (iFuncProcAddress_EbProjectReset <> 0) And (iIATEntryAddress_EbProjectReset <> 0) And (iASMThunkAddress_EbProjectReset <> 0) Then
                RemoveProp iHwndIDEMain, "EbPR.ProAd"
                RemoveProp iHwndIDEMain, "EbPR.IATAd"
                RemoveProp iHwndIDEMain, "EbPR.Thunk"

                CopyMemory ByVal iIATEntryAddress_EbProjectReset, iFuncProcAddress_EbProjectReset, 4
                VirtualFree iASMThunkAddress_EbProjectReset, 0, MEM_RELEASE
            End If
        End If
    End If
End Sub

Private Sub TerminateIDEProtection_OtherFunctions()
    Dim iInIDE As Boolean
    
    Debug.Assert MakeTrue(iInIDE)
    If iInIDE Then
        Dim iHwndIDEMain As Long
        Const MEM_RELEASE As Long = &H8000&
        
        iHwndIDEMain = GetIDEMainHwnd
        If iHwndIDEMain <> 0 Then
            ' TipStartMakeExe
            Dim iFuncProcAddress_TipStartMakeExe As Long
            Dim iIATEntryAddress_TipStartMakeExe As Long
            Dim iASMThunkAddress_TipStartMakeExe As Long
            
            iFuncProcAddress_TipStartMakeExe = GetProp(iHwndIDEMain, "TSME.ProAd")
            iIATEntryAddress_TipStartMakeExe = GetProp(iHwndIDEMain, "TSME.IATAd")
            iASMThunkAddress_TipStartMakeExe = GetProp(iHwndIDEMain, "TSME.Thunk")
            
            If (iFuncProcAddress_TipStartMakeExe <> 0) And (iIATEntryAddress_TipStartMakeExe <> 0) And (iASMThunkAddress_TipStartMakeExe <> 0) Then
                RemoveProp iHwndIDEMain, "TSME.ProAd"
                RemoveProp iHwndIDEMain, "TSME.IATAd"
                RemoveProp iHwndIDEMain, "TSME.Thunk"
                
                CopyMemory ByVal iIATEntryAddress_TipStartMakeExe, iFuncProcAddress_TipStartMakeExe, 4
                VirtualFree iASMThunkAddress_TipStartMakeExe, 0, MEM_RELEASE
            End If
            
            ' EbShowCode
            Dim iFuncProcAddress_EbShowCode As Long
            Dim iIATEntryAddress_EbShowCode As Long
            Dim iASMThunkAddress_EbShowCode As Long
            
            iFuncProcAddress_EbShowCode = GetProp(iHwndIDEMain, "EbSC.ProAd")
            iIATEntryAddress_EbShowCode = GetProp(iHwndIDEMain, "EbSC.IATAd")
            iASMThunkAddress_EbShowCode = GetProp(iHwndIDEMain, "EbSC.Thunk")
            
            If (iFuncProcAddress_EbShowCode <> 0) And (iIATEntryAddress_EbShowCode <> 0) And (iASMThunkAddress_EbShowCode <> 0) Then
                RemoveProp iHwndIDEMain, "EbSC.ProAd"
                RemoveProp iHwndIDEMain, "EbSC.IATAd"
                RemoveProp iHwndIDEMain, "EbSC.Thunk"
                
                CopyMemory ByVal iIATEntryAddress_EbShowCode, iFuncProcAddress_EbShowCode, 4
                VirtualFree iASMThunkAddress_EbShowCode, 0, MEM_RELEASE
            End If
            
            
            ' restore protect
            Dim iOldProtect As Long
            
            iOldProtect = GetProp(iHwndIDEMain, "IAT.OrigProt")
            If iOldProtect <> 0 Then
                Dim iIATEntryAddress_EbProjectReset As Long
                
                iIATEntryAddress_EbProjectReset = GetProp(iHwndIDEMain, "EbPR.IATAd")
                iIATEntryAddress_TipStartMakeExe = GetProp(iHwndIDEMain, "TSME.IATAd")
                iIATEntryAddress_EbShowCode = GetProp(iHwndIDEMain, "EbSC.IATAd")
                If (iIATEntryAddress_EbProjectReset = 0) And (iIATEntryAddress_TipStartMakeExe = 0) And (iIATEntryAddress_EbShowCode = 0) Then   ' if all IAT entries have already been restored
                    iIATEntryAddress_EbProjectReset = GetProp(iHwndIDEMain, "IAT.ProtAdd") ' get the address to restore the protection, one entry address is for all memory block protection, that includes all other IAT entries that habe been replaced
                    If iIATEntryAddress_EbProjectReset <> 0 Then
                        VirtualProtect iIATEntryAddress_EbProjectReset, 4, iOldProtect, iOldProtect
                        RemoveProp iHwndIDEMain, "IAT.OrigProt"
                        RemoveProp iHwndIDEMain, "IAT.ProtAdd"
                        mIDEProtectionInitialized = False
                        Set mCodeWindowsSubclassedHwnds = Nothing
                        Set mCodeWindowsToWatch = Nothing
                        Set mObjSubclassed = Nothing
                        Set mObjSubclassed_CallCount = Nothing
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub TerminateIDEProtection()
    DestroyTimerFindCodeWindows
    TerminateIDEProtection_EbProjectReset
    TerminateIDEProtection_OtherFunctions
End Sub

Private Function GetIATEntryAddress(ByVal nModule As String, ByVal nLibFncAddr As Long) As Long
    Dim hmod As Long
    Dim lpIAT As Long
    Dim IATLen As Long
    Dim IATPos As Long
    Dim DOSHdr As IMAGE_DOS_HEADER
    Dim PEHdr As IMAGE_OPTIONAL_HEADER32
    Const IMAGE_NT_SIGNATURE As Long = &H4550
    
    hmod = GetModuleHandle(nModule)
    If hmod = 0 Then Exit Function
    
    If nLibFncAddr = 0 Then Exit Function
    CopyMemory DOSHdr, ByVal hmod, LenB(DOSHdr)
    CopyMemory PEHdr, ByVal UnsignedAdd(hmod, DOSHdr.e_lfanew), LenB(PEHdr)
    If PEHdr.Magic = IMAGE_NT_SIGNATURE Then
        lpIAT = PEHdr.DataDirectory(15).VirtualAddress + hmod
        IATLen = PEHdr.DataDirectory(15).Size
        IATPos = lpIAT
        Do Until CLongToULong(IATPos) >= CLongToULong(UnsignedAdd(lpIAT, IATLen))
            If DeRef(IATPos) = nLibFncAddr Then
'                VirtualProtect IATPos, 4, PAGE_EXECUTE_READWRITE, 0
 '               CopyMemory ByVal IATPos, NewAddr, 4
                GetIATEntryAddress = IATPos
                Exit Do
            End If
            IATPos = UnsignedAdd(IATPos, 4)
        Loop
    End If
End Function

Private Sub WriteJump(ByRef ASM As Long, ByRef Addr As Long)
    WriteByte ASM, &HE9
    WriteLong ASM, Addr - ASM - 4
End Sub

Private Sub WriteCall(ByRef ASM As Long, ByRef Addr As Long)
    WriteByte ASM, &HE8
    WriteLong ASM, Addr - ASM - 4
End Sub

Private Sub WriteLong(ByRef ASM As Long, ByRef Lng As Long)
    CopyMemory ByVal ASM, Lng, 4
    ASM = ASM + 4
End Sub

Private Sub WriteByte(ByRef ASM As Long, ByRef B As Byte)
    CopyMemory ByVal ASM, B, 1
    ASM = ASM + 1
End Sub

Private Function DeRef(ByVal Addr As Long) As Long
    CopyMemory DeRef, ByVal Addr, 4
End Function

Private Function UnsignedAdd(ByVal Start As Long, ByVal Incr As Long) As Long
    UnsignedAdd = ((Start Xor &H80000000) + Incr) Xor &H80000000
End Function

Private Function CLongToULong(ByVal Value As Long) As Double
    Const OFFSET_4 As Double = 4294967296#
    '
    If Value < 0 Then
        CLongToULong = Value + OFFSET_4
    Else
        CLongToULong = Value
    End If
End Function

Private Function EnumCodeWindowsCallback(ByVal nHwnd As Long, ByVal Param As Long) As Long
    Const cCodeWindowClassName = "VbaWindow"
    Dim iWindowClassName As String
    Dim iHwndChild As Long
    Dim iHwndChild2 As Long
    
    If IsWindowVisible(nHwnd) <> 0 Then
        iWindowClassName = GetWindowClassName(nHwnd)
        Select Case iWindowClassName
            Case cCodeWindowClassName ' SDI IDE
                If IsHwndOfCodeWindowWatched(nHwnd) Then
                    If Not IsCodeWindowSubclassed(nHwnd) Then
                        SubClassCodeWindow nHwnd
                    End If
                End If
            Case "wndclass_desked_gsk" ' MDI IDE
                iHwndChild = GetWindow(nHwnd, GW_CHILD)
                Do Until iHwndChild = 0
                    If GetWindowClassName(iHwndChild) = "MDIClient" Then
                        iHwndChild2 = GetWindow(iHwndChild, GW_CHILD)
                        Do Until iHwndChild2 = 0
                            If GetWindowClassName(iHwndChild2) = cCodeWindowClassName Then
                                If IsHwndOfCodeWindowWatched(iHwndChild2) Then
                                    If Not IsCodeWindowSubclassed(iHwndChild2) Then
                                        SubClassCodeWindow iHwndChild2
                                    End If
                                End If
                            End If
                            iHwndChild2 = GetWindow(iHwndChild2, GW_HWNDNEXT)
                        Loop
                    End If
                    iHwndChild = GetWindow(iHwndChild, GW_HWNDNEXT)
                Loop
        End Select
    End If
    EnumCodeWindowsCallback = 1
End Function

Private Function IsCodeWindowSubclassed(nHwnd As Long) As Boolean
    Dim n As Long
    
    On Error GoTo ErrH
    n = mCodeWindowsSubclassedHwnds(CStr(nHwnd))
    IsCodeWindowSubclassed = True
    Exit Function
    
ErrH:
    Err.Clear
End Function

Private Sub SubClassCodeWindow(ByVal nHwnd As Long)
    If Not mCodeWindowsSubclassedHwnds Is Nothing Then
        If mAddressOf_CodeWindowWindowProc = 0 Then
            mAddressOf_CodeWindowWindowProc = GetAddresOfProc(AddressOf CodeWindowWindowProc)
        End If
        SetWindowSubclass nHwnd, mAddressOf_CodeWindowWindowProc, nHwnd, 0&
        mCodeWindowsSubclassedHwnds.Add nHwnd, CStr(nHwnd)
    End If
End Sub

Private Sub UnSubClassCodeWindow(ByVal nHwnd As Long)
    If mAddressOf_CodeWindowWindowProc <> 0 Then
        RemoveWindowSubclass nHwnd, mAddressOf_CodeWindowWindowProc, nHwnd
        If Not mCodeWindowsSubclassedHwnds Is Nothing Then
            mCodeWindowsSubclassedHwnds.Remove CStr(nHwnd)
        End If
    End If
End Sub

Private Function CodeWindowWindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
    Const WM_KEYDOWN As Long = &H100
    Const WM_KEYUP As Long = &H101
    Dim iDo As Boolean
    
    If mAllSubclassesRemoved Then
        UnSubClassCodeWindow hWnd
    Else
        'Debug.Print GetMessageName(iMsg)
        Select Case iMsg
            Case WM_KEYDOWN, WM_KEYUP
                Select Case wParam
                    Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyPageDown, vbKeyPageUp, vbKeyHome, vbKeyEnd, vbKeyControl, vbKeyTab, vbKeyF2, vbKeyF3, vbKeyF5, vbKeyF8, vbKeyF9, vbKeyEscape
                    Case Else
                        If GetKeyState(vbKeyControl) >= 0 Then
                            iDo = True
                        Else
                            Select Case wParam
                                Case vbKeyS, vbKeyA, vbKeyI, vbKeyG, vbKeyJ, vbKeyL
                                Case Else
                                    iDo = True
                            End Select
                        End If
                        If iDo Then
                            RemoveAllSubclasses
                        End If
                End Select
            Case WM_DESTROY
                UnSubClassCodeWindow hWnd
        End Select
    End If
    CodeWindowWindowProc = DefSubclassProc(hWnd, iMsg, wParam, lParam)
End Function

Private Sub TimerFindCodeWindowsProc(ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long)
    mTimerFindCodeWindowsHandle = 0
    KillTimer 0, uElapse
    If Not mAllSubclassesRemoved Then
        ' find if a new code window that we are watching is open and subclass it
        EnumThreadWindows App.ThreadID, AddressOf EnumCodeWindowsCallback, 0
    End If
End Sub

Private Sub DestroyTimerFindCodeWindows()
    If mTimerFindCodeWindowsHandle <> 0 Then
        KillTimer 0, mTimerFindCodeWindowsHandle
        mTimerFindCodeWindowsHandle = 0
    End If
End Sub

Private Function EnumThreadProc_GetIDEMainWindow(ByVal lhWnd As Long, ByVal lParam As Long) As Long
    Dim iBuff As String * 255
    Dim iWinClass As String
    Dim iRet As Long
    
    iRet = GetClassName(lhWnd, iBuff, 255)
    
    If iRet > 0 Then
        iWinClass = Left$(iBuff, iRet)
    Else
        iWinClass = ""
    End If
    
    Select Case iWinClass
        Case "wndclass_desked_gsk"
            mIDEMainHwnd = lhWnd
            EnumThreadProc_GetIDEMainWindow = 0
            Exit Function
    End Select
    EnumThreadProc_GetIDEMainWindow = 1
End Function


Private Sub IDEIsResetting()
    ' Debug.Print "IDE is resetting " & Rnd
    
    TerminateIDEProtection_OtherFunctions
    
    mIDEIsResetting = True
    RemoveAllSubclasses
    mIDEIsResetting = False
End Sub

Private Sub IDEAboutToMakeExe()
    ' Debug.Print "IDE is about to compile " & Rnd
    
    ' Can't do much here, only to set a variable, otherwise the parameter(s) get mangled in the processor registers and VB crashes
    mCompiling = True
End Sub

Private Sub IDECodeWindowShowing()
    'Debug.Print "A code window is about to show up " & Rnd
    
    DestroyTimerFindCodeWindows
    If Not mAllSubclassesRemoved Then
        ' start a timer to look for the new window because here it is still no created
        mTimerFindCodeWindowsHandle = SetTimer(0, 0, 1, AddressOf TimerFindCodeWindowsProc)
    End If
End Sub

#End If


Public Function UniTextBoxWindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    Dim refObject As Object, unRefObj As Object
    Dim lProc As Long, bUnicode As Boolean
    
    CopyMemory unRefObj, GetProp(hWnd, "cUniTextBox"), 4&
    Set refObject = unRefObj
    CopyMemory unRefObj, 0&, 4&
    
    UniTextBoxWindowProc = refObject.DoWindowMsg(hWnd, uMsg, wParam, lParam, bUnicode, lProc)
    Set refObject = Nothing
    If lProc Then
        If bUnicode Then
            UniTextBoxWindowProc = CallWindowProcW(lProc, hWnd, uMsg, wParam, lParam)
        Else
            UniTextBoxWindowProc = CallWindowProc(lProc, hWnd, uMsg, wParam, lParam)
        End If
    End If
    
End Function

