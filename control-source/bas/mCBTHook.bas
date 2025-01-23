Attribute VB_Name = "mCBTHook"
Option Explicit

Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function SetWindowSubclass Lib "comctl32" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
Private Declare Function DefSubclassProc Lib "comctl32" (ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function RemoveWindowSubclass Lib "comctl32" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long) As Long

Private mCBTHook As Long
Private mAddressOfWindowProc As Long
Private mHwndSubclassed As Long
Private mNewTabControl As NewTab

Public Sub InstallCBTHook(nNewTabControl As NewTab)
    Const WH_CBT = 5
    
    If mCBTHook = 0 Then
        Set mNewTabControl = nNewTabControl
        mCBTHook = SetWindowsHookEx(WH_CBT, AddressOf FormCreationHookProc, App.hInstance, App.ThreadID)
    End If
End Sub

Public Sub UninstallCBTHook()
    If mCBTHook <> 0 Then
        UnhookWindowsHookEx mCBTHook
        mCBTHook = 0
        Set mNewTabControl = Nothing
    End If
End Sub

Private Function FormCreationHookProc(ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Const HCBT_CREATEWND = 3
    Const HCBT_DESTROYWND = 4
    Const HCBT_ACTIVATE = 5
    Const HCBT_SETFOCUS = 9
    
    If uMsg = HCBT_CREATEWND Then
        Dim iStr As String
        
        If mAddressOfWindowProc = 0 Then
            mAddressOfWindowProc = GetAddresOfProc(AddressOf WindowProc)
        End If
        If mHwndSubclassed = 0 Then
            iStr = GetWindowClassName(wParam)
            If (iStr = "ThunderFormDC") Or (iStr = "ThunderRT6FormDC") Then
                mHwndSubclassed = wParam
                SetWindowSubclass mHwndSubclassed, mAddressOfWindowProc, 1&, 0&
            End If
        End If
    ElseIf uMsg = HCBT_DESTROYWND Then
        iStr = GetWindowClassName(wParam)
        If (iStr = "ThunderFormDC") Or (iStr = "ThunderRT6FormDC") Then
            If Not mNewTabControl Is Nothing Then mNewTabControl.TDIFormClosing wParam
        End If
    ElseIf (uMsg = HCBT_ACTIVATE) Or (uMsg = HCBT_SETFOCUS) Then
        iStr = GetWindowClassName(wParam)
        If (iStr = "ThunderFormDC") Or (iStr = "ThunderRT6FormDC") Then
            mNewTabControl.TDIFocusForm wParam
        End If
    End If
End Function

Private Function WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
    Const WM_DESTROY As Long = &H2
    Const WM_NCDESTROY As Long = &H82&
    Const WM_CREATE As Long = &H1&
    Const GWL_EXSTYLE As Long = (-20)
    Const WS_EX_TOOLWINDOW As Long = &H80&
    
    If (iMsg = WM_CREATE) Or (iMsg = WM_DESTROY) Or (iMsg = WM_NCDESTROY) Then
        RemoveWindowSubclass hWnd, mAddressOfWindowProc, 1&
        If iMsg = WM_CREATE Then
            If Not mNewTabControl Is Nothing Then
                If mNewTabControl.IsParentEnabled Then
                    If WindowHasCaption(mHwndSubclassed) Then
                        If (GetWindowLong(mHwndSubclassed, GWL_EXSTYLE) And WS_EX_TOOLWINDOW) = 0 Then
                            mNewTabControl.TDIPutFormIntoTab mHwndSubclassed
                        End If
                    End If
                Else
                    mNewTabControl.ShowsModalForm
                End If
            End If
        End If
        mHwndSubclassed = 0
    End If
    WindowProc = DefSubclassProc(hWnd, iMsg, wParam, ByVal lParam)
End Function

Private Function GetAddresOfProc(nProcAddress As Long) As Long
    GetAddresOfProc = nProcAddress
End Function

Private Function GetWindowClassName(nHwnd As Long) As String
    Dim iClassName As String
    Dim iSize As Long
    
    If nHwnd = 0 Then Exit Function
    
    iClassName = Space$(64)
    iSize = GetClassName(nHwnd, iClassName, Len(iClassName))
    GetWindowClassName = Left$(iClassName, iSize)
End Function

Private Function WindowHasCaption(nHwnd As Long) As Boolean
    Const WS_CAPTION = &HC00000
    Const GWL_STYLE = (-16)
    
    WindowHasCaption = (GetWindowLong(nHwnd, GWL_STYLE) And WS_CAPTION) <> 0
End Function

