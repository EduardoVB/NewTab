VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cToolTipEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Implements IBSSubclass

Private Const LF_FACESIZE = 32

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type LOGFONTW
    LFHeight As Long
    LFWidth As Long
    LFEscapement As Long
    LFOrientation As Long
    LFWeight As Long
    LFItalic As Byte
    LFUnderline As Byte
    LFStrikeOut As Byte
    LFCharset As Byte
    LFOutPrecision As Byte
    LFClipPrecision As Byte
    LFQuality As Byte
    LFPitchAndFamily As Byte
    LFFaceName(0 To ((LF_FACESIZE * 2) - 1)) As Byte
End Type

Private Type NONCLIENTMETRICSW
    cbSize As Long
    iBorderWidth As Long
    iScrollWidth As Long
    iScrollHeight As Long
    iCaptionWidth As Long
    iCaptionHeight As Long
    lfCaptionFont As LOGFONTW
    iSMCaptionWidth As Long
    iSMCaptionHeight As Long
    lfSMCaptionFont As LOGFONTW
    iMenuWidth As Long
    iMenuHeight As Long
    lfMenuFont As LOGFONTW
    lfStatusFont As LOGFONTW
    lfMessageFont As LOGFONTW
End Type

Private Type InitCommonControlsExType
    dwSize As Long 'size of this structure
    dwICC As Long 'flags indicating which classes to be initialized
End Type

Private Declare Function InitCommonControlsEx Lib "comctl32" (init As InitCommonControlsExType) As Long

Private Const ICC_BAR_CLASSES = &H4            ' toolbar, statusbar, trackbar, tooltips

Private Const WM_DESTROY As Long = &H2&
Private Const WM_WINDOWPOSCHANGED As Long = &H47&
Private Const WM_MOVE As Long = &H3&
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WS_EX_LAYOUTRTL = &H400000
Private Const WM_SETFONT As Long = &H30

Private Const MOUSEEVENTF_LEFTDOWN = &H2 ' Left button down
Private Const MOUSEEVENTF_LEFTUP = &H4 ' Left button up

Public Event Closed()
Public Event BeforeClose()

Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, Col As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_HWNDPARENT As Long = (-8)
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetMessageExtraInfo Lib "user32" () As Long
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dX As Long, ByVal dY As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Declare Function SystemParametersInfoW Lib "user32" (ByVal uAction As Long, ByVal uParam As Long, ByRef pvParam As Any, ByVal fWinIni As Long) As Long
Private Declare Sub CopyMemoryAny1 Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC& Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long)

Private Enum bsSystemFontConstants
   bsIconFont = 1
   bsCaptionFont = 2
   bsSmallCaptionFont = 3
   bsMenuFont = 4
   bsStatusAndTooltipFont = 5
   bsMsgBoxFont = 6
End Enum

''Windows API Constants
Private Const WM_USER = &H400
Private Const CW_USEDEFAULT = &H80000000

''Windows API Types
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

''Tooltip Window Constants
Private Const TTS_NOPREFIX = &H2
'Private Const TTF_TRANSPARENT = &H100
'Private Const TTF_CENTERTIP = &H2
Private Const TTM_ADDTOOLW = (WM_USER + 50)
'Private Const TTM_ACTIVATE = WM_USER + 1
Private Const TTM_UPDATETIPTEXTW = (WM_USER + 57)
Private Const TTM_SETMAXTIPWIDTH = (WM_USER + 24)
Private Const TTM_SETTIPBKCOLOR = (WM_USER + 19)
Private Const TTM_SETTIPTEXTCOLOR = (WM_USER + 20)
Private Const TTM_SETTITLEW = (WM_USER + 33)
Private Const TTS_BALLOON = &H40
Private Const TTS_ALWAYSTIP = &H1
Private Const TTF_SUBCLASS = &H10
'Private Const TTF_IDISHWND = &H1
'Private Const TTM_SETDelayTimeSeconds = (WM_USER + 3)
'Private Const TTDT_AUTOPOP = 2
'Private Const TTDT_INITIAL = 3
Private Const TTS_CLOSE As Long = &H80
Private Const TTM_TRACKPOSITION = (WM_USER + 18)
Private Const TTM_TRACKACTIVATE As Long = (WM_USER + 17)
Private Const TTF_ABSOLUTE As Long = &H80
Private Const TTF_TRACK As Long = &H20

Private Const TOOLTIPS_CLASSA = "tooltips_class32"

''Tooltip Window Types
'Private Type TOOLINFO
'    lSize As Long
'    lFlags As Long
'    hWnd As Long
'    lId As Long
'    lpRect As RECT
'    hInstance As Long
'    lpStr As String
'    lParam As Long
'End Type

Private Type TOOLINFOW
    cbSize As Long
    uFlags As Long
    hWnd As Long
    uId As Long
    RECT As RECT
    hInst As Long
    lpszText As Long
    lParam As Long
End Type

Public Enum vbExBalloonTooltipIconConstants
    vxTTNoIcon = 0
    vxTTIconInfo = 1
    vxTTIconWarning = 2
    vxTTIconError = 3
End Enum

Public Enum vbExBalloonTooltipStyleConstants
    vxTTStandard
    vxTTBalloon
End Enum

'local variable(s) to hold property value(s)
Private mBackColor As Variant
Private mForeColor As Variant
Private mTitle As String
Private mIcon As vbExBalloonTooltipIconConstants
Private mStyle As vbExBalloonTooltipStyleConstants
Private mTipText As String
Private mDelayTimeSeconds As Variant
Private mVisibleTimeSeconds As Variant
Private mPositionX As Variant
Private mPositionY As Variant
Private mCloseButton As Boolean
Private mWidth As Variant
Private mRestrictMouseMoveToTwips As Long
Private mTag As String
Private mRightToLeft As Boolean

'private data
Private mTTHwnd As Long ' hWnd of the tooltip
Private mParentHwnd As Long ' hWnd of the window the tooltip attached to
Private mTi As TOOLINFOW
Private mOldOwner As Long
Private mTTShown As Boolean
Private mAttached As Boolean
Private mShowTime As Single
Private mStartTime As Single
Private mMouseStartPosition As POINTAPI
Private mDelayTimeSecondsSng As Single
Private mVisibleTimeSecondsSng As Single
Private mParentOriginalRectPos As RECT
Private mTTOriginalRectPos As RECT
Private mFont As StdFont

Private WithEvents mTmrDelay As cTimer
Attribute mTmrDelay.VB_VarHelpID = -1
Private WithEvents mTmrClose As cTimer
Attribute mTmrClose.VB_VarHelpID = -1
Private WithEvents mTmrTrackMouse As cTimer
Attribute mTmrTrackMouse.VB_VarHelpID = -1

Private miFont As iFont

Public Property Let PositionX(ByVal nValue As Variant)
    If Not IsEmpty(nValue) Then
        mPositionX = Val(nValue)
    End If
End Property

Public Property Get PositionX() As Variant
   PositionX = mPositionX
End Property

Public Property Let PositionY(ByVal nValue As Variant)
    If Not IsEmpty(nValue) Then
        mPositionY = Val(nValue)
    End If
End Property

Public Property Get PositionY() As Variant
   PositionY = mPositionY
End Property

Public Property Let CloseButton(ByVal nValue As Boolean)
   mCloseButton = nValue
End Property

Public Property Get CloseButton() As Boolean
   CloseButton = mCloseButton
End Property

Public Property Let Width(ByVal nValue As Variant)
   If IsEmpty(nValue) Then Exit Property
   mWidth = Val(nValue)
End Property

Public Property Get Width() As Variant
   Width = mWidth
End Property

Public Property Let RestrictMouseMoveToTwips(ByVal nValue As Long)
   mRestrictMouseMoveToTwips = nValue
End Property

Public Property Get RestrictMouseMoveToTwips() As Long
   RestrictMouseMoveToTwips = mRestrictMouseMoveToTwips
End Property

Public Property Let Style(ByVal nValue As vbExBalloonTooltipStyleConstants)
   mStyle = nValue
End Property

Public Property Get Style() As vbExBalloonTooltipStyleConstants
   Style = mStyle
End Property

Public Function Create(ByVal nParentHwnd As Long) As Boolean
    Dim lWinStyle As Long
    Dim iAbsolutePosition As Boolean
    Dim iPt As POINTAPI
    Dim iLng As Long
    Dim iPositionX As Variant
    Dim iPositionY As Variant
    Dim initcc As InitCommonControlsExType
    
    If IsWindowVisible(nParentHwnd) = 0 Then Exit Function
    
    mParentHwnd = nParentHwnd
    
    initcc.dwSize = Len(initcc)
    initcc.dwICC = ICC_BAR_CLASSES
    InitCommonControlsEx initcc
        
    If Not IsEmpty(mPositionX) Or Not IsEmpty(mPositionY) Then
        iAbsolutePosition = True
        If IsEmpty(mPositionX) Or IsEmpty(mPositionY) Then
            GetCursorPos iPt
            ScreenToClient mParentHwnd, iPt
            If IsEmpty(mPositionX) Then
                iPositionX = iPt.X * Screen.TwipsPerPixelX
            Else
                iPositionX = mPositionX
            End If
            If IsEmpty(mPositionY) Then
                iPositionY = iPt.Y * Screen.TwipsPerPixelY
            Else
                iPositionY = mPositionY
            End If
        Else
            iPositionX = mPositionX
            iPositionY = mPositionY
        End If
    End If
    
    CloseTip
    
    lWinStyle = TTS_ALWAYSTIP Or TTS_NOPREFIX
    
    If mCloseButton Then
        lWinStyle = lWinStyle Or TTS_CLOSE
        
        If IsEmpty(mDelayTimeSeconds) Then
            mDelayTimeSecondsSng = 0 ' default
        Else
            mDelayTimeSecondsSng = mDelayTimeSeconds
        End If
        If IsEmpty(mVisibleTimeSeconds) Then
            mVisibleTimeSecondsSng = 120 ' default
        Else
            mVisibleTimeSecondsSng = mVisibleTimeSeconds
        End If
    Else
        If IsEmpty(mDelayTimeSeconds) Then
            mDelayTimeSecondsSng = 0.5 ' default
        Else
            mDelayTimeSecondsSng = mDelayTimeSeconds
        End If
        If IsEmpty(mVisibleTimeSeconds) Then
            mVisibleTimeSecondsSng = 15 ' default
        Else
            mVisibleTimeSecondsSng = mVisibleTimeSeconds
        End If
    End If
    
    ''create baloon style if desired
    If mStyle = vxTTBalloon Then lWinStyle = lWinStyle Or TTS_BALLOON
    
    mTTHwnd = CreateWindowEx(IIf(mRightToLeft, WS_EX_LAYOUTRTL, 0&), _
        TOOLTIPS_CLASSA, _
        vbNullString, _
        lWinStyle, _
        CW_USEDEFAULT, _
        CW_USEDEFAULT, _
        CW_USEDEFAULT, _
        CW_USEDEFAULT, _
        0&, _
        0&, _
        App.hInstance, _
        0&)
    
    Set mFont = GetSystemFont(bsStatusAndTooltipFont)
    If Not mFont Is Nothing Then
      Set miFont = mFont
      SendMessageLong mTTHwnd, WM_SETFONT, miFont.hFont, 1
    End If
               
    If Not IsEmpty(mWidth) Then
        If mWidth > 0 Then
            SendMessageLong mTTHwnd, TTM_SETMAXTIPWIDTH, 0, CLng(mWidth / Screen.TwipsPerPixelX)
        Else
            SendMessageLong mTTHwnd, TTM_SETMAXTIPWIDTH, 0, Screen.Width / Screen.TwipsPerPixelX / 2 ' to enable multiline and limit the width to half the screen size
        End If
    Else
        SendMessageLong mTTHwnd, TTM_SETMAXTIPWIDTH, 0, Screen.Width / Screen.TwipsPerPixelX / 2 ' to enable multiline and limit the width to half the screen size
    End If
               
    ''now set our tooltip info structure
    With mTi
        If iAbsolutePosition Then
           .uFlags = TTF_TRACK Or TTF_ABSOLUTE
        Else
           .uFlags = TTF_SUBCLASS  'Or TTF_TRACK Or TTF_ABSOLUTE
        End If
        
        ''set the hWnd prop to our parent control's hWnd
        .hWnd = mParentHwnd
        .uId = mParentHwnd  '0
        .hInst = App.hInstance
        '.lpstr = ALREADY SET
        '.lpRect = lpRect
        .cbSize = Len(mTi)
    End With
    
    ''add the tooltip structure
    SendMessage mTTHwnd, TTM_ADDTOOLW, 0&, mTi
    
    ''if we want a title or we want an icon
    If mTitle <> "" Then
        SendMessage mTTHwnd, TTM_SETTITLEW, CLng(mIcon), ByVal StrPtr(mTitle)
    Else
         If (mIcon <> vxTTNoIcon) Or mCloseButton Then
            SendMessage mTTHwnd, TTM_SETTITLEW, CLng(mIcon), ByVal StrPtr(" ")
         End If
    End If
    
    If Not IsEmpty(mBackColor) Then
        TranslateColor CLng(mBackColor), 0&, iLng
        SendMessage mTTHwnd, TTM_SETTIPBKCOLOR, iLng, 0&
    End If
    If Not IsEmpty(mForeColor) Then
        TranslateColor CLng(mForeColor), 0&, iLng
        SendMessage mTTHwnd, TTM_SETTIPTEXTCOLOR, iLng, 0&
    End If
    
    If iAbsolutePosition Then
        iPt.X = iPositionX / Screen.TwipsPerPixelX
        iPt.Y = iPositionY / Screen.TwipsPerPixelY
        ClientToScreen mParentHwnd, iPt
        SendMessageLong mTTHwnd, TTM_TRACKPOSITION, 0&, MakeLong(iPt.X, iPt.Y)
    End If
    
    mOldOwner = SetOwner(mTTHwnd, mParentHwnd)
    
    AttachMessage Me, mTTHwnd, WM_WINDOWPOSCHANGED
    AttachMessage Me, mTTHwnd, WM_DESTROY
    AttachMessage Me, mTTHwnd, WM_LBUTTONDOWN
    AttachMessage Me, mParentHwnd, WM_WINDOWPOSCHANGED
    AttachMessage Me, mParentHwnd, WM_DESTROY
    AttachMessage Me, mParentHwnd, WM_MOVE
    
    mAttached = True
    
    If mDelayTimeSecondsSng = 0 Then
'        SendMessage mTTHwnd, TTM_TRACKACTIVATE, 1&, mTi
        Set mTmrDelay = New cTimer
        mTmrDelay.Interval = 1
        mStartTime = Timer - 1
    Else
        Set mTmrDelay = New cTimer
        mTmrDelay.Interval = 100
        mStartTime = Timer
    End If
    
    If mRestrictMouseMoveToTwips > 0 Then
        If Not (mCloseButton And (mRestrictMouseMoveToTwips = 1000)) Then
            GetCursorPos mMouseStartPosition
            Set mTmrTrackMouse = New cTimer
            mTmrTrackMouse.Interval = 100
        End If
    End If
    
    GetWindowRect mParentHwnd, mParentOriginalRectPos
    
    Create = mTTHwnd <> 0
End Function

Private Function MakeLong(ByVal wLow As Integer, ByVal wHigh As Integer) As Long
    MakeLong = wHigh * &H10000 + wLow
End Function

Public Property Let Icon(ByVal nValue As vbExBalloonTooltipIconConstants)
   mIcon = nValue
   If mTTHwnd <> 0 And mTitle <> Empty And mIcon <> vxTTNoIcon Then
      SendMessage mTTHwnd, TTM_SETTITLEW, CLng(mIcon), ByVal StrPtr(mTitle)
   End If
End Property

Public Property Get Icon() As vbExBalloonTooltipIconConstants
   Icon = mIcon
End Property

Public Property Let BackColor(ByVal nValue As Variant)
   If Not IsNumeric(nValue) Then Exit Property
   mBackColor = nValue
   If mTTHwnd <> 0 Then
      SendMessage mTTHwnd, TTM_SETTIPBKCOLOR, mBackColor, 0&
   End If
End Property

Public Property Get BackColor() As Variant
   BackColor = mBackColor
End Property

Public Property Let ForeColor(ByVal nValue As Variant)
   If Not IsNumeric(nValue) Then Exit Property
   mForeColor = nValue
   If mTTHwnd <> 0 Then
      SendMessage mTTHwnd, TTM_SETTIPTEXTCOLOR, mForeColor, 0&
   End If
End Property

Public Property Get ForeColor() As Variant
   ForeColor = mForeColor
End Property

Public Property Let Title(ByVal nValue As String)
   mTitle = nValue
   If mTTHwnd <> 0 And mTitle <> Empty And mIcon <> vxTTNoIcon Then
      SendMessage mTTHwnd, TTM_SETTITLEW, CLng(mIcon), ByVal StrPtr(mTitle)
   End If
End Property

Public Property Get Title() As String
   Title = mTitle
End Property

Public Property Let TipText(ByRef nValue As String)
   mTipText = nValue
   mTi.lpszText = StrPtr(mTipText)
   If mTTHwnd <> 0 Then
      SendMessage mTTHwnd, TTM_UPDATETIPTEXTW, 0&, mTi
   End If
End Property

Public Property Get TipText() As String
   TipText = mTipText
End Property

Private Sub Class_Terminate()
   CloseTip
End Sub

Public Sub CloseTip()
    If mTTHwnd <> 0 Then
        RaiseEvent BeforeClose
        Detach
        SetOwner mTTHwnd, mOldOwner
        DestroyWindow mTTHwnd
        mTTHwnd = 0
        RaiseEvent Closed
    End If

    DestroyTimers
    Set mFont = Nothing
    Set miFont = Nothing
End Sub

Private Sub Detach()
    If mAttached Then
        DetachMessage Me, mTTHwnd, WM_DESTROY
        DetachMessage Me, mTTHwnd, WM_WINDOWPOSCHANGED
        DetachMessage Me, mTTHwnd, WM_LBUTTONDOWN
        DetachMessage Me, mParentHwnd, WM_WINDOWPOSCHANGED
        DetachMessage Me, mParentHwnd, WM_DESTROY
        DetachMessage Me, mParentHwnd, WM_MOVE
        mAttached = False
    End If
End Sub

Private Sub DestroyTimers()
    If Not mTmrDelay Is Nothing Then
        mTmrDelay.Interval = 0
        Set mTmrDelay = Nothing
    End If
    If Not mTmrClose Is Nothing Then
        mTmrClose.Interval = 0
        Set mTmrClose = Nothing
    End If
    If Not mTmrTrackMouse Is Nothing Then
        mTmrTrackMouse.Interval = 0
        Set mTmrTrackMouse = Nothing
    End If
End Sub

Public Property Get VisibleTimeSeconds() As Variant
   VisibleTimeSeconds = mVisibleTimeSeconds
End Property

Public Property Let VisibleTimeSeconds(ByVal nValue As Variant)
   If IsEmpty(nValue) Then Exit Property
   mVisibleTimeSeconds = nValue
End Property

Public Property Get DelayTimeSeconds() As Variant
   DelayTimeSeconds = mDelayTimeSeconds
End Property

Public Property Let DelayTimeSeconds(ByVal nValue As Variant)
   If IsEmpty(nValue) Then Exit Property
   mDelayTimeSeconds = nValue
End Property

Public Property Get ToolTipHwnd() As Long
    ToolTipHwnd = mTTHwnd
End Property

Public Property Get ParentHwnd() As Long
    ParentHwnd = mParentHwnd
End Property

Private Function SetOwner(ByVal HwndWindow As Long, ByVal HwndofOwner As Long) As Long
    SetOwner = SetWindowLong(HwndWindow, GWL_HWNDPARENT, HwndofOwner)
End Function

Private Function IBSSubclass_MsgResponse(ByVal hWnd As Long, ByVal iMsg As Long) As EMsgResponse
    IBSSubclass_MsgResponse = emrPreprocess
End Function

Private Sub IBSSubclass_UnsubclassIt()

End Sub

Private Function IBSSubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByRef wParam As Long, ByRef lParam As Long, ByRef bConsume As Boolean) As Long
    Dim iRect As RECT
    
    Select Case iMsg
        Case WM_DESTROY
            CloseTip
        Case WM_WINDOWPOSCHANGED
            If hWnd = mTTHwnd Then
                If IsWindowVisible(mTTHwnd) = 0 Then
                    If mTTShown Then
                        CloseTip
                        mTTShown = False
                    End If
                Else
                    If Not mTTShown Then
                        mTTShown = True
                        Set mTmrClose = New cTimer
                        mTmrClose.Interval = 100
                        mShowTime = Timer
                        GetWindowRect mTTHwnd, mTTOriginalRectPos
                    End If
                End If
            Else ' parent
                If IsWindowVisible(mParentHwnd) = 0 Then
                    CloseTip
                    mTTShown = False
                End If
            End If
        Case WM_LBUTTONDOWN
            mTTShown = False
            CloseTip
            mouse_event MOUSEEVENTF_LEFTDOWN, 0&, 0&, 0&, GetMessageExtraInfo()
        Case WM_MOVE
            GetWindowRect mParentHwnd, iRect
            MoveWindow mTTHwnd, mTTOriginalRectPos.Left + iRect.Left - mParentOriginalRectPos.Left, mTTOriginalRectPos.Top + iRect.Top - mParentOriginalRectPos.Top, mTTOriginalRectPos.Right - mTTOriginalRectPos.Left, mTTOriginalRectPos.Bottom - mTTOriginalRectPos.Top, 1&
    End Select
    
'    IBSSubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
End Function

Private Sub mTmrClose_ThatTime()
    If (Timer - mShowTime) >= mVisibleTimeSecondsSng Then
        mTmrClose.Interval = 0
        Set mTmrClose = Nothing
        CloseTip
    End If
End Sub

Private Sub mTmrDelay_ThatTime()
    If (Timer - mStartTime) >= mDelayTimeSecondsSng Then
        mTmrDelay.Interval = 0
        Set mTmrDelay = Nothing
        If mRestrictMouseMoveToTwips > 0 Then mTmrTrackMouse_ThatTime
        If mTTHwnd <> 0 Then
            SendMessage mTTHwnd, TTM_TRACKACTIVATE, 1&, mTi
        End If
    End If
End Sub

Public Sub Reset()
    If mTTShown Then
        mShowTime = Timer
    End If
End Sub

Private Sub mTmrTrackMouse_ThatTime()
    Dim iMpos As POINTAPI
    Dim iPixX As Long
    Dim iPixY As Long
    Dim iClose As Boolean
    
    GetCursorPos iMpos
    iPixX = mRestrictMouseMoveToTwips / Screen.TwipsPerPixelX
    iPixY = mRestrictMouseMoveToTwips / Screen.TwipsPerPixelY
    
    If iPixX = 0 Then
        iClose = True
    End If
    If Abs(iMpos.X - mMouseStartPosition.X) > iPixX Then
        iClose = True
    End If
    If Abs(iMpos.Y - mMouseStartPosition.Y) > iPixY Then
        iClose = True
    End If
    
    If iClose Then
        If Not mTmrTrackMouse Is Nothing Then
            mTmrTrackMouse.Interval = 0
            Set mTmrTrackMouse = Nothing
        End If
        CloseTip
    End If
End Sub

Public Sub RaiseEventClosed()
Attribute RaiseEventClosed.VB_MemberFlags = "40"
    RaiseEvent Closed
End Sub

Public Sub RaiseEventBeforeClose()
    RaiseEvent BeforeClose
End Sub


Public Property Let Tag(nValue As String)
    mTag = nValue
End Property

Public Property Get Tag() As String
    Tag = mTag
End Property


Public Property Get RightToLeft() As Boolean
    RightToLeft = mRightToLeft
End Property

Public Property Let RightToLeft(nValue As Boolean)
    mRightToLeft = nValue
End Property

Private Function GetSystemFont(nSystemFont As bsSystemFontConstants) As StdFont
    Dim iLF As LOGFONTW
    Dim iNcm As NONCLIENTMETRICSW
    Dim iILf As LOGFONTW
    Dim iRet As Long
    Const SPI_GETNONCLIENTMETRICS = 41
    Const SPI_GETICONTITLELOGFONT = 31
    
    
'    iNcm.cbSize = 340
'    iNcm.cbSize = 500
    iNcm.cbSize = LenB(iNcm)
    iRet = SystemParametersInfoW(SPI_GETNONCLIENTMETRICS, iNcm.cbSize, iNcm, 0)
    If (iRet = 0) Then Exit Function
    
    Select Case nSystemFont
        Case bsCaptionFont
            CopyMemoryAny1 iLF, iNcm.lfCaptionFont, LenB(iNcm.lfCaptionFont)
        Case bsIconFont
            iRet = SystemParametersInfoW(SPI_GETICONTITLELOGFONT, LenB(iILf), iILf, 0)
            If (iRet <> 0) Then
                CopyMemoryAny1 iLF, iILf, LenB(iILf)
            End If
        Case bsMenuFont
            CopyMemoryAny1 iLF, iNcm.lfMenuFont, LenB(iNcm.lfMenuFont)
        Case bsMsgBoxFont
            CopyMemoryAny1 iLF, iNcm.lfMessageFont, LenB(iNcm.lfMessageFont)
        Case bsSmallCaptionFont
            CopyMemoryAny1 iLF, iNcm.lfSMCaptionFont, LenB(iNcm.lfSMCaptionFont)
        Case bsStatusAndTooltipFont
            CopyMemoryAny1 iLF, iNcm.lfStatusFont, LenB(iNcm.lfStatusFont)
        Case Else
            Exit Function
    End Select
    
    Set GetSystemFont = LogFontToStdFont(iLF)
End Function

Private Function LogFontToStdFont(LF As LOGFONTW, Optional nPrinterFont As Boolean) As iFont
    Dim iFontName As String
    Dim iDPIY As Single
    Dim iDC As Long
    Const LOGPIXELSY As Long = 90
    
    Set LogFontToStdFont = New StdFont
    
    If LF.LFHeight = 0 Then Exit Function
    
    If nPrinterFont Then
        iDPIY = GetDeviceCaps(Printer.hDC, LOGPIXELSY)
    Else
        iDC = GetDC(0)
        iDPIY = GetDeviceCaps(iDC, LOGPIXELSY)
        ReleaseDC 0, iDC
    End If
    
    iFontName = LF.LFFaceName
    If Len(iFontName) > 0 Then
        If InStr(iFontName, Chr$(0)) > 0 Then
            iFontName = Left$(iFontName, InStr(iFontName, Chr$(0)) - 1)
        End If
    End If
    
    If iFontName <> "" Then
        LogFontToStdFont.Name = iFontName
    Else
        LogFontToStdFont.Name = "Arial"
    End If
    
    Select Case LF.LFHeight
        Case Is < 0
            LogFontToStdFont.Size = -LF.LFHeight / iDPIY * 72
        Case Is > 0
            LogFontToStdFont.Size = LF.LFHeight / iDPIY * 72 * 0.8 ' lF.lfHeight / Screen.TwipsPerPixelY / 2.777777777
        Case Else
            LogFontToStdFont.Size = 12
    End Select
    
    If LF.LFWeight > 1000 Then
        LogFontToStdFont.Weight = 400
        If LogFontToStdFont.Size > 20 Then LogFontToStdFont.Size = 12
    Else
        LogFontToStdFont.Weight = LF.LFWeight
    End If
    
    LogFontToStdFont.Italic = LF.LFItalic
    LogFontToStdFont.Strikethrough = LF.LFStrikeOut
    LogFontToStdFont.Underline = LF.LFUnderline
    LogFontToStdFont.Charset = LF.LFCharset
End Function

