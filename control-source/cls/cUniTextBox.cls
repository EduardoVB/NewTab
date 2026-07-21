VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cUniTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' A simple callback class used to offer unicode textbox on the property page
' It is truly simple and only has a couple properties/methods, but works well

' Should you want to use this for your own projects, here is how you can set it up
' 1) Add a borderless picturebox to your form/propertypage/usercontrol
' 2) Set that picturebox backcolor to the color you want for your textbox
' 3) Declare this class WithEvents in your project
' 4) Call this class' CreateTextBox function, ensure you pass pixel sizes
' 5) Monitor the 3 public events below
' 6) Synchronize GotFocus/LostFocus with the picturebox to ensure tabbing is maintained
' 7) Declare a boolean value in your project and set its values as so:
'   -- When getting class GotFocus event:
'       If yourBoolean = False Then
'           set focus to your picturebox (i.e., Picture1.SetFocus)
'       End If
'   -- When getting class LostFocus event:
'       set yourBoolean = False
'   -- In your picturebox' GotFocus event:
'       If yourBoolean = False Then
'           set yourBoolean = True
'           call this class' SetFocus routine
'       End If
' 8) Add this procedure into a module & include appropriate APIs from that procedure
'        Public Function UniTextBoxWindowProc(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
'
'            Dim refObject As Object, unRefObj As Object
'            Dim lProc As LongPtr, bUnicode As Boolean
'
'            CopyMemory unRefObj, GetProp(hWnd, "cUniTextBox"), LenB(lProc)
'            Set refObject = unRefObj
'            CopyMemory unRefObj, 0&, LenB(lProc)
'
'            UniTextBoxWindowProc = refObject.DoWindowMsg(hWnd, uMsg, wParam, lParam, bUnicode, lProc)
'            Set refObject = Nothing
'            If lProc Then
'                If bUnicode Then
'                    UniTextBoxWindowProc = CallUniTextBoxWindowProcW(lProc, hWnd, uMsg, wParam, lParam)
'                Else
'                    UniTextBoxWindowProc = CallUniTextBoxWindowProc(lProc, hWnd, uMsg, wParam, lParam)
'                End If
'            End If
'
'        End Function
' 9) Use one class instance per textbox


Public Event Change()
Public Event GotFocus()
Public Event LostFocus()

Private Declare Function SetProp Lib "user32.dll" Alias "SetPropA" (ByVal hWnd As LongPtr, ByVal lpString As String, ByVal hData As LongPtr) As Long
'Private Declare Function GetProp Lib "user32.dll" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function CreateWindowEx Lib "user32.dll" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, ByRef lpParam As Any) As LongPtr
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByRef lParam As Any) As LongPtr
' GWL_WNDPROC holds a function address (LRESULT/LONG_PTR): SetWindowLong/GetWindowLong
' truncate to 32 bits even on 64-bit Windows, so the ...LongPtr variants must be used.
' NOTE: SetWindowLongPtrA/W only exist as real DLL symbols on 64-bit Windows;
' on 32-bit those names don't exist (they're C-side macros pointing at SetWindowLongA/W),
' so the Alias has to be conditioned on bitness.
#If Win64 Then
    Private Declare Function SetWindowLongPtr Lib "user32.dll" Alias "SetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    Private Declare Function GetWindowLongPtr Lib "user32.dll" Alias "GetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
#Else
    Private Declare Function SetWindowLongPtr Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    Private Declare Function GetWindowLongPtr Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
#End If

#If Win64 Then
    Private Declare Function SetWindowLongPtrW Lib "user32.dll" Alias "SetWindowLongPtrW" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    Private Declare Function GetWindowLongPtrW Lib "user32.dll" Alias "GetWindowLongPtrW" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
#Else
    Private Declare Function SetWindowLongPtrW Lib "user32.dll" Alias "SetWindowLongW" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    Private Declare Function GetWindowLongPtrW Lib "user32.dll" Alias "GetWindowLongW" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
#End If

Private Declare Function CreateWindowExW Lib "user32.dll" (ByVal dwExStyle As Long, ByVal lpClassName As LongPtr, ByVal lpWindowName As LongPtr, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, ByRef lpParam As Any) As LongPtr
Private Declare Function SendMessageW Lib "user32.dll" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByRef lParam As Any) As LongPtr

Private Const WM_GETFONT As Long = &H31
Private Const WM_SETFONT As Long = &H30
Private Const WM_GETTEXTLENGTH As Long = &HE
Private Const WM_GETTEXT As Long = &HD
Private Const WM_SETTEXT As Long = &HC

Private Declare Function DestroyWindow Lib "user32.dll" (ByVal hWnd As LongPtr) As Long
Private Declare Function GetDesktopWindow Lib "user32.dll" () As LongPtr
Private Declare Function IsWindowUnicode Lib "user32.dll" (ByVal hWnd As LongPtr) As Long
Private Declare Function SetFocusAPI Lib "user32.dll" Alias "SetFocus" (ByVal hWnd As LongPtr) As LongPtr

Private Const WM_KILLFOCUS As Long = &H8
Private Const WM_SETFOCUS As Long = &H7
Private Const EN_CHANGE As Long = &H300
Private Const WM_COMMAND As Long = &H111
Private Const GWL_WNDPROC As Long = -4
Private Const WM_DESTROY As Long = &H2
Private Const WS_CHILD As Long = &H40000000

Public Enum TextBoxStyleEnum
    ES_AUTOHSCROLL = &H80&
    ES_AUTOVSCROLL = &H40&
    ES_CENTER = &H1&
    ES_LEFT = &H0&
    ES_LOWERCASE = &H10&
    ES_MULTILINE = &H4&
    ES_NOHIDESEL = &H100&
    ES_NUMBER = &H2000&
    ES_PASSWORD = &H20&
    ES_READONLY = &H800&
    ES_RIGHT = &H2&
    ES_SUNKEN = &H4000&
    ES_UPPERCASE = &H8&
    ES_WANTRETURN = &H1000&
    WS_DISABLED = &H8000000
    WS_DLGFRAME = &H400000
    WS_HSCROLL = &H100000
    WS_THICKFRAME = &H40000
    WS_VISIBLE = &H10000000
    WS_VSCROLL = &H200000
End Enum
Public Enum TextBoxStyleExEnum
    WS_EX_CLIENTEDGE = &H200&
    WS_EX_LAYOUTRTL = &H400000
    WS_EX_LEFTSCROLLBAR = &H4000&
    WS_EX_NOPARENTNOTIFY = &H4&
    WS_EX_RIGHT = &H1000&
    WS_EX_RTLREADING = &H2000&
    WS_EX_STATICEDGE = &H20000
End Enum

Private m_Hwnd As LongPtr
Private m_Parent As LongPtr
Private m_Unicode As Boolean
Private m_WinProc As LongPtr
Private m_WinProcParent As LongPtr

Public Function CreateTextBox(ByVal ContainerHwnd As LongPtr, ByVal X As Long, ByVal Y As Long, _
                        ByVal Width As Long, ByVal Height As Long, _
                        ByVal Style As TextBoxStyleEnum, _
                        ByVal ExtendedStyle As TextBoxStyleExEnum, _
                        Optional ByVal Text As String = vbNullString) As LongPtr
                        
    pvUnSubclass True
    If m_Unicode Then
        m_Hwnd = CreateWindowExW(ExtendedStyle, StrPtr("Edit"), StrPtr(Text), Style Or WS_CHILD, X, Y, Width, Height, ContainerHwnd, 0&, App.hInstance, ByVal 0&)
    Else
        m_Hwnd = CreateWindowEx(ExtendedStyle, "Edit", Text, Style Or WS_CHILD, X, Y, Width, Height, ContainerHwnd, 0&, App.hInstance, ByVal 0&)
    End If
    If m_Hwnd Then
        m_Parent = ContainerHwnd
        SetProp m_Hwnd, "cUniTextBox", ObjPtr(Me)
        If m_Unicode Then
            SendMessageW m_Hwnd, WM_SETFONT, SendMessage(m_Parent, WM_GETFONT, 0&, ByVal 0&), ByVal 0&
            m_WinProc = GetWindowLongPtrW(m_Hwnd, GWL_WNDPROC)
            SetWindowLongPtrW m_Hwnd, GWL_WNDPROC, AddressOf UniTextBoxWindowProc
        Else
            SendMessage m_Hwnd, WM_SETFONT, SendMessage(m_Parent, WM_GETFONT, 0&, ByVal 0&), ByVal 0&
            m_WinProc = GetWindowLongPtr(m_Hwnd, GWL_WNDPROC)
            SetWindowLongPtr m_Hwnd, GWL_WNDPROC, AddressOf UniTextBoxWindowProc
        End If
        
        SetProp m_Parent, "cUniTextBox", ObjPtr(Me)
        m_WinProcParent = GetWindowLongPtr(m_Parent, GWL_WNDPROC)
        SetWindowLongPtr m_Parent, GWL_WNDPROC, AddressOf UniTextBoxWindowProc
        
        CreateTextBox = m_Hwnd
    End If

End Function

Private Sub pvUnSubclass(bDestroy As Boolean)
    If m_Hwnd Then
        If m_Unicode Then
            SetWindowLongPtrW m_Hwnd, GWL_WNDPROC, m_WinProc
        Else
            SetWindowLongPtr m_Hwnd, GWL_WNDPROC, m_WinProc
        End If
        If bDestroy Then DestroyWindow m_Hwnd
        m_Hwnd = 0&
        SetWindowLongPtr m_Parent, GWL_WNDPROC, m_WinProcParent
    End If

End Sub

Public Function DoWindowMsg(hWnd As LongPtr, uMsg As Long, wParam As LongPtr, lParam As LongPtr, Unicode As Boolean, wndProc As LongPtr) As LongPtr

    
    Select Case uMsg
    
        Case WM_SETFOCUS
            If hWnd = m_Hwnd Then RaiseEvent GotFocus
        Case WM_KILLFOCUS
            If hWnd = m_Hwnd Then RaiseEvent LostFocus
        Case WM_COMMAND
            If lParam = m_Hwnd And hWnd = m_Parent Then
                If (wParam And &H7FFF0000) \ &H10000 = EN_CHANGE Then RaiseEvent Change
            End If
        Case WM_DESTROY
            pvUnSubclass False
            
    End Select
    
    If hWnd = m_Hwnd Then
        Unicode = m_Unicode
        wndProc = m_WinProc         ' set to zero if eating the message & set return value appropriately
    Else
        Unicode = False
        wndProc = m_WinProcParent
    End If

End Function

Public Property Get Text() As String
    If m_Hwnd Then
        Dim lLength As Long
        If m_Unicode Then
            lLength = SendMessageW(m_Hwnd, WM_GETTEXTLENGTH, 0&, ByVal 0&)
            If lLength Then
                Text = String$(lLength, 0)
                SendMessageW m_Hwnd, WM_GETTEXT, lLength + 1&, ByVal StrPtr(Text)
            End If
        Else
            lLength = SendMessage(m_Hwnd, WM_GETTEXTLENGTH, 0&, ByVal 0&)
            If lLength Then
                Text = String$(lLength, 0)
                SendMessage m_Hwnd, WM_GETTEXT, lLength + 1&, ByVal StrPtr(Text)
            End If
        End If
    End If
End Property

Public Property Let Text(newVal As String)

    If m_Hwnd Then
        If m_Unicode Then
            SendMessageW m_Hwnd, WM_SETTEXT, 0&, ByVal StrPtr(newVal)
        Else
            SendMessage m_Hwnd, WM_SETTEXT, 0&, ByVal newVal
        End If
    End If
    
End Property


Public Sub SetFocus()
    If m_Hwnd Then SetFocusAPI m_Hwnd
End Sub

Private Sub Class_Initialize()
    m_Unicode = Not (IsWindowUnicode(GetDesktopWindow) = 0&)
End Sub

Private Sub Class_Terminate()
    pvUnSubclass True
End Sub
