Attribute VB_Name = "mMenuWheel"
Option Explicit

' Windows menus do not scroll with the mouse wheel, and the menu loop does not pass the wheel through
' WH_MSGFILTER either. While a popup menu is open a low level mouse hook takes the wheel and turns each
' step into the Up or Down key the menu loop already understands

Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExW" (ByVal idHook As Long, ByVal lpfn As LongPtr, ByVal hmod As LongPtr, ByVal dwThreadId As Long) As LongPtr
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As LongPtr) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As LongPtr, ByVal nCode As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Declare Function PostMessageW Lib "user32" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long
Private Declare Function SystemParametersInfoW Lib "user32" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Long, ByVal fuWinIni As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As LongPtr)

Private Const WH_MOUSE_LL As Long = 14
Private Const HC_ACTION As Long = 0
Private Const WM_MOUSEWHEEL As Long = &H20A
Private Const WM_KEYDOWN As Long = &H100
Private Const WM_KEYUP As Long = &H101
Private Const VK_UP As Long = &H26
Private Const VK_DOWN As Long = &H28
Private Const WHEEL_DELTA As Long = 120
Private Const SPI_GETWHEELSCROLLLINES As Long = &H68

Private mHook As LongPtr
Private mTargetHwnd As LongPtr
Private mRemainder As Long
Private mLinesPerNotch As Long

' nTargetHwnd: any window of this thread. The menu loop takes the keys from the thread's queue,
' whichever window they are posted to
Public Sub MenuWheelHookInstall(ByVal nTargetHwnd As LongPtr)
    Dim iLines As Long
    
    If mHook <> 0 Then Exit Sub
    iLines = 3
    If SystemParametersInfoW(SPI_GETWHEELSCROLLLINES, 0, iLines, 0) = 0 Then iLines = 3
    ' -1 is "one page per notch", which a menu has no notion of
    If (iLines < 1) Or (iLines > 20) Then iLines = 10
    mLinesPerNotch = iLines
    mRemainder = 0
    mTargetHwnd = nTargetHwnd
    mHook = SetWindowsHookEx(WH_MOUSE_LL, AddressOf MenuWheelHookProc, App.hInstance, 0)
End Sub

Public Sub MenuWheelHookUninstall()
    If mHook <> 0 Then
        UnhookWindowsHookEx mHook
        mHook = 0
    End If
    mTargetHwnd = 0
End Sub

Private Function MenuWheelHookProc(ByVal nCode As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    Dim iDelta As Integer
    Dim iSteps As Long
    Dim iKey As Long
    Dim s As Long
    
    If (nCode = HC_ACTION) And (wParam = WM_MOUSEWHEEL) And (mTargetHwnd <> 0) Then
        ' MSLLHOOKSTRUCT: the delta is the high word of mouseData, which follows the 8 byte POINT
        CopyMemory iDelta, ByVal lParam + 10, 2
        mRemainder = mRemainder + iDelta
        iSteps = mRemainder \ WHEEL_DELTA
        mRemainder = mRemainder - iSteps * WHEEL_DELTA
        If iSteps > 0 Then
            iKey = VK_UP
        Else
            iKey = VK_DOWN
            iSteps = -iSteps
        End If
        For s = 1 To iSteps * mLinesPerNotch
            PostMessageW mTargetHwnd, WM_KEYDOWN, iKey, 0
            PostMessageW mTargetHwnd, WM_KEYUP, iKey, &HC0000000
        Next s
        MenuWheelHookProc = 1
        Exit Function
    End If
    MenuWheelHookProc = CallNextHookEx(mHook, nCode, wParam, lParam)
End Function
