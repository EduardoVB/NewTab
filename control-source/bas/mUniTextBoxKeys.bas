Attribute VB_Name = "mUniTextBoxKeys"
Option Explicit

' Inside a property page the frame takes the cursor keys as dialog navigation before a window that is
' not a VB control can see them, so an API edit box loses the focus on every arrow. A thread message
' hook hands those keys to the edit box itself as soon as they are retrieved, and turns the queued
' message into WM_NULL so nothing downstream acts on it again

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type T_MSG
    hWnd As LongPtr
    Message As Long
    wParam As LongPtr
    lParam As LongPtr
    time As Long
    pt As POINTAPI
End Type

Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExW" (ByVal idHook As Long, ByVal lpfn As LongPtr, ByVal hmod As LongPtr, ByVal dwThreadId As Long) As LongPtr
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As LongPtr) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As LongPtr, ByVal nCode As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare Function SendMessageW Lib "user32" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Declare Function GetPropW Lib "user32" (ByVal hWnd As LongPtr, ByVal lpString As LongPtr) As LongPtr
Private Declare Function GetClassNameW Lib "user32" (ByVal hWnd As LongPtr, ByVal lpClassName As LongPtr, ByVal nMaxCount As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As LongPtr)

Private Const WH_GETMESSAGE As Long = 3
Private Const HC_ACTION As Long = 0
Private Const PM_REMOVE As Long = 1
Private Const WM_NULL As Long = 0
Private Const WM_KEYDOWN As Long = &H100
Private Const WM_KEYUP As Long = &H101
Private Const VK_END As Long = &H23
Private Const VK_HOME As Long = &H24
Private Const VK_LEFT As Long = &H25
Private Const VK_DOWN As Long = &H28

Private mHook As LongPtr
Private mUsers As Long

Public Sub UniTextBoxKeysHookAdd()
    mUsers = mUsers + 1
    If mHook = 0 Then mHook = SetWindowsHookEx(WH_GETMESSAGE, AddressOf UniTextBoxKeysHookProc, 0, GetCurrentThreadId)
End Sub

Public Sub UniTextBoxKeysHookRemove()
    If mUsers > 0 Then mUsers = mUsers - 1
    If (mUsers = 0) And (mHook <> 0) Then
        UnhookWindowsHookEx mHook
        mHook = 0
    End If
End Sub

Private Function UniTextBoxKeysHookProc(ByVal nCode As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    Dim iMsg As T_MSG
    Dim iKey As Long
    
    If (nCode = HC_ACTION) And (wParam = PM_REMOVE) Then
        CopyMemory iMsg, ByVal lParam, LenB(iMsg)
        If (iMsg.Message = WM_KEYDOWN) Or (iMsg.Message = WM_KEYUP) Then
            iKey = CLng(iMsg.wParam)
            If (iKey >= VK_END) And (iKey <= VK_DOWN) Then
                If IsUniTextBoxEdit(iMsg.hWnd) Then
                    SendMessageW iMsg.hWnd, iMsg.Message, iMsg.wParam, iMsg.lParam
                    iMsg.Message = WM_NULL
                    CopyMemory ByVal lParam, iMsg, LenB(iMsg)
                End If
            End If
        End If
    End If
    UniTextBoxKeysHookProc = CallNextHookEx(mHook, nCode, wParam, lParam)
End Function

' The container of the edit box carries the same property, so the class tells the two apart
Private Function IsUniTextBoxEdit(ByVal nHwnd As LongPtr) As Boolean
    Dim iClass As String
    Dim iLen As Long
    
    If nHwnd = 0 Then Exit Function
    If GetPropW(nHwnd, StrPtr("cUniTextBox")) = 0 Then Exit Function
    iClass = Space$(16)
    iLen = GetClassNameW(nHwnd, StrPtr(iClass), Len(iClass))
    IsUniTextBoxEdit = (StrComp(Left$(iClass, iLen), "Edit", vbTextCompare) = 0)
End Function
