VERSION 5.00
Begin VB.UserControl NewTab 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3840
   ControlContainer=   -1  'True
   LockControls    =   -1  'True
   PropertyPages   =   "ctlNewTab.ctx":0000
   ScaleHeight     =   2880
   ScaleWidth      =   3840
   ToolboxBitmap   =   "ctlNewTab.ctx":0079
   Begin VB.PictureBox picTDIFormContainer 
      BorderStyle     =   0  'None
      Height          =   612
      Index           =   0
      Left            =   3120
      ScaleHeight     =   612
      ScaleWidth      =   492
      TabIndex        =   9
      Top             =   240
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.Timer tmrShowTabTTT 
      Enabled         =   0   'False
      Interval        =   900
      Left            =   1560
      Top             =   1560
   End
   Begin VB.PictureBox picAuxIconFont 
      BorderStyle     =   0  'None
      Height          =   492
      Left            =   2400
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   61
      TabIndex        =   7
      Top             =   1920
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.Timer tmrPreHighlightIcon 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   1560
      Top             =   2280
   End
   Begin VB.Timer tmrHighlightIcon 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1200
      Top             =   2280
   End
   Begin VB.Timer tmrCheckTabDrag 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   1200
      Top             =   1920
   End
   Begin VB.Timer tmrTabDragging 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1200
      Top             =   1560
   End
   Begin VB.PictureBox picAux2 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   624
      Left            =   1920
      ScaleHeight     =   52
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   76
      TabIndex        =   6
      Top             =   684
      Visible         =   0   'False
      Width           =   912
   End
   Begin VB.PictureBox picCover 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   492
      Left            =   1440
      ScaleHeight     =   492
      ScaleWidth      =   732
      TabIndex        =   5
      Top             =   1920
      Visible         =   0   'False
      Width           =   732
      Begin VB.Timer tmrTDIIconColor 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   120
         Top             =   0
      End
   End
   Begin VB.Timer tmrTabTransition 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   792
      Top             =   2268
   End
   Begin VB.Timer tmrRestoreDropMode 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   792
      Top             =   1908
   End
   Begin VB.Timer tmrCheckDuplicationByIDEPaste 
      Interval        =   1
      Left            =   792
      Top             =   1548
   End
   Begin VB.Timer tmrHighlightEffect 
      Enabled         =   0   'False
      Interval        =   15
      Left            =   396
      Top             =   2268
   End
   Begin VB.Timer tmrSubclassControls 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   396
      Top             =   1908
   End
   Begin VB.Timer tmrCancelDoubleClick 
      Enabled         =   0   'False
      Interval        =   350
      Left            =   396
      Top             =   1548
   End
   Begin VB.Timer tmrCheckContainedControlsAdditionDesignTime 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   36
      Top             =   2268
   End
   Begin VB.PictureBox picInactiveBodyThemed 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   624
      Left            =   972
      ScaleHeight     =   52
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   76
      TabIndex        =   4
      Top             =   684
      Visible         =   0   'False
      Width           =   912
   End
   Begin VB.PictureBox picBodyThemed 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   624
      Left            =   0
      ScaleHeight     =   52
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   76
      TabIndex        =   3
      Top             =   684
      Visible         =   0   'False
      Width           =   912
   End
   Begin VB.PictureBox picAux 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   624
      Left            =   1944
      ScaleHeight     =   52
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   76
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   912
   End
   Begin VB.Timer tmrDraw 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   36
      Top             =   1908
   End
   Begin VB.Timer tmrTabMouseLeave 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   36
      Top             =   1548
   End
   Begin VB.PictureBox picRotate 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   624
      Left            =   972
      ScaleHeight     =   52
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   76
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   912
   End
   Begin VB.PictureBox picDraw 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   624
      Left            =   0
      ScaleHeight     =   52
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   76
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   912
   End
   Begin VB.Label lblTDILabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   192
      Left            =   468
      TabIndex        =   8
      Top             =   1080
      Visible         =   0   'False
      Width           =   3156
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "NewTab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "NewTab: tabbed control"
Option Explicit

' Set the line below to 0 to disable the IDE protection (with IDE protection some features are be lost in the IDE-uncompiled when it is set to 1, like changing tabs with a click at design time)
#Const NOSUBCLASSINIDE = 1

Implements IBSSubclass

Private Type STYLESTRUCT
   styleOld As Long
   styleNew As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type WINDOWPOS
   hWnd As Long
   hWndInsertAfter As Long
   X As Long
   Y As Long
   cx As Long
   cy As Long
   Flags As Long
End Type

'Bitmap type used to store Bitmap Data
Private Type BITMAP
  bmType As Long
  bmWidth As Long
  bmHeight As Long
  bmWidthBytes As Long
  bmPlanes As Integer
  bmBitsPixel As Integer
  bmBits As Long
End Type

Private Type PAINTSTRUCT
    hDC                     As Long
    fErase                  As Long
    rcPaint                 As RECT
    fRestore                As Long
    fIncUpdate              As Long
    rgbReserved(1 To 32)    As Byte
End Type

Private Type T_MSG
    hWnd As Long
    Message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Private Type tagHIGHCONTRAST
    cbSize As Long
    dwFlags As Long
    lpszDefaultScheme As Long
End Type

Private Type XFORM
    eM11 As Single
    eM12 As Single
    eM21 As Single
    eM22 As Single
    eDx As Single
    eDy As Single
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type
 
Private Type PicBmp
    Size As Long
    Type As Long
    hBmp As Long
    hPal As Long
    Reserved As Long
End Type
 
Private Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthW" (ByVal hWnd As Long) As Long
Private Declare Function DefWindowProcW Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SetGraphicsMode Lib "gdi32" (ByVal hDC As Long, ByVal iMode As Long) As Long
Private Declare Function SetWorldTransform Lib "gdi32" (ByVal hDC As Long, lpXform As XFORM) As Long
Private Declare Function GetWorldTransform Lib "gdi32" (ByVal hDC As Long, lpXform As XFORM) As Long

Private Const GM_ADVANCED = 2

Private Declare Function Polygon Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Sub ClipCursor Lib "user32" (lpRect As Any)
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function GetCursor Lib "user32" () As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As tagHIGHCONTRAST, ByVal fuWinIni As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Private Const SPI_GETHIGHCONTRAST As Long = 66
Private Const HCF_HIGHCONTRASTON As Long = 1

Private Declare Function SetLayout Lib "gdi32" (ByVal hDC As Long, ByVal dwLayout As Long) As Long
Private Const LAYOUT_RTL = &H1                               ' Right to left
Private Const LAYOUT_BITMAPORIENTATIONPRESERVED = &H8

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long

Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function ValidateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function RevokeDragDrop Lib "ole32" (ByVal hWnd As Long) As Long

Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dX As Long, ByVal dY As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Declare Function GetMessageExtraInfo Lib "user32" () As Long
Private Const MOUSEEVENTF_LEFTDOWN = &H2 ' Left button down
Private Const MOUSEEVENTF_LEFTUP = &H4 ' Left button up

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Const VK_LBUTTON = &H1
Private Const VK_RBUTTON = &H2
Private Const SM_SWAPBUTTON = 23&

Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As T_MSG, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Private Const PM_REMOVE = &H1

Private Declare Function GetUpdateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function BeginPaint Lib "user32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function EndPaint Lib "user32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long

Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Const RDW_ALLCHILDREN = &H80
Private Const RDW_INTERNALPAINT = &H2
Private Const RDW_INVALIDATE = &H1
Private Const RDW_UPDATENOW = &H100

Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long

Private Const WM_SYSCOLORCHANGE As Long = &H15
Private Const WM_THEMECHANGED As Long = &H31A
Private Const WM_PAINT As Long = &HF
Private Const WM_MOVE As Long = &H3&
Private Const WM_MOUSEACTIVATE As Long = &H21
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_LBUTTONUP As Long = &H202
Private Const WM_SETFOCUS As Long = &H7
Private Const WM_SETREDRAW As Long = &HB&
Private Const WM_USER As Long = &H400
Private Const WM_DRAW As Long = WM_USER + 10 ' custom message
Private Const WM_INIT As Long = WM_USER + 11 ' custom message
Private Const WM_LBUTTONDBLCLK As Long = &H203&
Private Const WM_MOUSEMOVE As Long = &H200
Private Const WM_PRINTCLIENT As Long = &H318
Private Const WM_NCACTIVATE As Long = &H86&
Private Const WM_WINDOWPOSCHANGING = &H46&
Private Const WM_GETDPISCALEDSIZE As Long = &H2E4&
Private Const WM_SETCURSOR As Long = &H20
Private Const WM_ACTIVATE As Long = &H6
Private Const WM_ACTIVATEAPP As Long = &H1C&
Private Const WM_STYLECHANGING = &H7C&
Private Const WM_DESTROY As Long = &H2&
Private Const WM_SHOWWINDOW As Long = &H18&

'Private Const MA_NOACTIVATEANDEAT As Long = &H4
Private Const WM_MOUSELEAVE As Long = &H2A3

Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetAncestor Lib "user32.dll" (ByVal hWnd As Long, ByVal gaFlags As Long) As Long
Private Const GA_ROOT = 2

Private Declare Function DrawTextW Lib "user32" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function GetObjectA Lib "gdi32.dll" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal crTransparent As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, Col As Long) As Long
Private Declare Sub ColorRGBToHLS Lib "shlwapi" (ByVal clrRGB As Long, ByRef pwHue As Integer, ByRef pwLuminance As Integer, ByRef pwSaturation As Integer)
Private Declare Function ColorHLSToRGB Lib "shlwapi" (ByVal wHue As Integer, ByVal wLuminance As Integer, ByVal wSaturation As Integer) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

'Draw Text Constants
Private Const DT_CALCRECT = &H400&
Private Const DT_LEFT = &H0&
Private Const DT_CENTER = &H1&
Private Const DT_RIGHT = &H2&
Private Const DT_SINGLELINE = &H20&
Private Const DT_VCENTER = &H4&
Private Const DT_END_ELLIPSIS = &H8000&
Private Const DT_MODIFYSTRING = &H10000
Private Const DT_WORDBREAK = &H10&
Private Const DT_RTLREADING As Long = &H20000

Private Declare Function OpenThemeData Lib "uxtheme" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme" (ByVal hTheme As Long) As Long
Private Declare Function DrawThemeBackground Lib "uxtheme" (ByVal hTheme As Long, ByVal lHDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As RECT, pClipRect As RECT) As Long

Private Const TABP_TABITEM = 1
Private Const TABP_TABITEMLEFTEDGE = 2
Private Const TABP_TABITEMRightEDGE = 3
'Private Const TABP_TABITEMBOTHEDGE = 4
'Private Const TABP_TOPTABITEM = 5
'Private Const TABP_TOPTABITEMLEFTEDGE = 6
'Private Const TABP_TOPTABITEMRIGHTEDGE = 7
'Private Const TABP_TOPTABITEMBOTHEDGE = 8
Private Const TABP_PANE = 9
'Private Const TABP_BODY = 10

Private Const TIS_NORMAL = 1
Private Const TIS_HOT = 2
Private Const TIS_SELECTED = 3
Private Const TIS_DISABLED = 4
Private Const TIS_FOCUSED = 5

Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function PlgBlt Lib "gdi32" (ByVal hdcDest As Long, lpPoint As POINTAPI, ByVal hdcSrc As Long, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hbmMask As Long, ByVal xMask As Long, ByVal yMask As Long) As Long

Private Const HALFTONE = 4
Private Type COLORADJUSTMENT
        caSize As Integer
        caFlags As Integer
        caIlluminantIndex As Integer
        caRedGamma As Integer
        caGreenGamma As Integer
        caBlueGamma As Integer
        caReferenceBlack As Integer
        caReferenceWhite As Integer
        caContrast As Integer
        caBrightness As Integer
        caColorfulness As Integer
        caRedGreenTint As Integer
End Type

Private Declare Function GetColorAdjustment Lib "gdi32" (ByVal hDC As Long, lpca As COLORADJUSTMENT) As Long
Private Declare Function SetColorAdjustment Lib "gdi32" (ByVal hDC As Long, lpca As COLORADJUSTMENT) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long

Private Const IDC_HAND = 32649&

Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As Long, ByVal hwndChildAfter As Long, ByVal lpszClass As String, ByVal lpszCaption As String) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long


Private Const cAuxTransparentColor As Long = &HFF01FF ' Not the MaskColor, but another transparent color for internal operations
Private Const cTabIconDistanceToCaption As Long = 7
Private Const cIconClickExtend As Long = 5 ' extend the click area a bit for catching the click when the user does not click exactly on the icon area but very near

Private Enum NTMouseButtonsConstants
    ntMBLeft = 1&
    ntMBRight = 2&
End Enum

Private Enum NTRotatePicDirectionConstants
    nt90DegreesClockWise = 0
    nt90DegreesCounterClockWise = 1
    ntFlipVertical = 2
    ntFlipHorizontal = 3
End Enum

Private Enum NTHighlightGradientConstants
    ntGradientNone = 0
    ntGradientPlain = 1
    ntGradientSimple = 3
    ntGradientDouble = 4
End Enum

Private Enum NTHighlightIntensityConstants
    ntHighlightIntensityStrong = 0
    ntHighlightIntensityLight = 1
End Enum

Private Enum NTCornerPositionConstants
    ntCornerTopleft = 0
    ntCornerTopRight = 1
    ntCornerBottomLeft = 2
    ntCornerBottomRight = 3
End Enum

' Public Enums
Public Enum NTTabOrientationConstants
    ssTabOrientationTop = 0
    ssTabOrientationBottom = 1
    ssTabOrientationLeft = 2
    ssTabOrientationRight = 3
    ntTabOrientationLeftHorizontal = 4
    ntTabOrientationRightHorizontal = 5
End Enum

Public Enum NTMousePointerConstants
    ssDefault = 0
    ssArrow = 1
    ssCross = 2
    ssIBeam = 3
    ssIcon = 4
    ssSize = 5
    ssSizeNESW = 6
    ssSizeNS = 7
    ssSizeNWSE = 8
    ssSizeEW = 9
    ssUpArrow = 10
    ssHourglass = 11
    ntNoDrop = 12
    ssArrowHourglass = 13
    ssArrowQuestion = 14
    ssSizeAll = 15
    ssCustom = 99
End Enum

Public Enum NTOLEDropConstants
    ssOLEDropNone = 0
    ssOLEDropManual = 1
End Enum

Public Enum NTStyleConstants
    ssStyleTabbedDialog = 0
    ssStylePropertyPage = 1
    ntStyleTabStrip = 2
    ntStyleFlat = 3
    ntStyleWindows = 4
End Enum

Public Enum NTAutoYesNoConstants
    ntNo = 0
    ntYes = 1
    ntYNAuto = 2
End Enum

Public Enum NTTabWidthStyleConstants
    ntTWTabStripEmulation = 0
    ntTWTabCaptionWidth = 1
    ntTWFixed = 2
    ntTWAuto = 3
    ntTWStretchToFill = 4
    ntTWTabCaptionWidthFillRows = 5
End Enum

Public Enum NTTabAppearanceConstants
    ntTAAuto = 0
    ntTATabbedDialog = 1
    ntTATabbedDialogRounded = 2
    ntTAPropertyPage = 3
    ntTAPropertyPageRounded = 4
    ntTAFlat = 5
End Enum

Public Enum NTIconAlignmentConstants
    ntIconAlignBeforeCaption = 0
    ntIconAlignCenteredBeforeCaption = 1
    ntIconAlignAfterCaption = 2
    ntIconAlignCenteredAfterCaption = 3
    ntIconAlignStart = 4
    ntIconAlignEnd = 5
    ntIconAlignCenteredOnTab = 6
    ntIconAlignAtTop = 7
    ntIconAlignAtBottom = 8
End Enum

Public Enum NTAutoRelocateControlsConstants
    ntRelocateNever = 0
    ntRelocateAlways = 1
    ntRelocateOnTabOrientationChange = 2
End Enum

Public Enum NTBackStyleConstants
    ntTransparent = 0
    ntOpaque = 1
    ntOpaqueSelectedTab = 2
End Enum

Public Enum NTTabTransitionConstants
    ntTransitionImmediate = 0
    ntTransitionFaster = 1
    ntTransitionFast = 2
    ntTransitionMedium = 3
    ntTransitionSlow = 4
    ntTransitionSlower = 5
End Enum

Public Enum NTHighlightModeFlagsConstants
    ntHLAuto = 0
    ntHLNone = 1
    ntHLBackgroundGradient = 2
    ntHLBackgroundDoubleGradient = 3
    ntHLBackgroundPlain = 4
    ntHLBackgroundTypeFilter = 7
    ntHLBackgroundLight = 8
    ntHLCaptionBold = 16
    ntHLCaptionUnderlined = 32 '*
    ntHLFlatBar = 64
    ntHLFlatBarGrip = 128
    ntHLFlatBarWithGrip = 196
    ntHLExtraHeight = 256
    ntHLFlatDrawBorder = 512
    ntHLAllFlags = 1023
End Enum

Public Enum NTFlatBorderModeConstants
    ntBorderTabs = 0
    ntBorderSelectedTab = 1
End Enum

Public Enum NTFlatBarPositionConstants
    ntBarPositionTop = 0
    ntBarPositionBottom = 1
End Enum

Public Enum NTTDINewTabTypeConstants
    ntDefaultTab = 0
    ntNewTabByClickingIcon = 1
    ntLastTabClosed = 2
    ntTDIForm = 3
End Enum

Public Enum NTSubclassingMethodConstants
    ntSMSetWindowSubclass = 0
    ntSMSetWindowLong = 1
    ntSMDisabled = 2
    ntSM_SWSOnlyUserControl = 3
    ntSM_SWLOnlyUserControl = 4
End Enum

Public Enum NTTDIModeConstants
    ntTDIModeNone = 0
    ntTDIModeControls = 1
    ntTDIModeForms = 2
End Enum

Public Enum NTFindTabMethodConstants
    ntFindCaption = 0
    ntFindOriginalIndex = 1
    ntFindData = 2
    ntFindTag = 3
    ntFindPartialCaption = 4
    ntFindPartialCaptionReverse = 5
    ntFindKey = 6
End Enum

Public Enum NTTabCustomColorIDConstants
    ntCCBackColorTab = 0
    ntCCBackColorSelectedTab = 1
    ntCCHighlightColor = 2
    ntCCHighlightColorSelectedTab = 3
    ntCCFlatBarColorInactive = 4
    ntCCFlatBarColorHighlight = 5
    ntCCFlatBarColorSelectedTab = 6
    ntCCFlatTabBorderColorHighlight = 7
    ntCCFlatTabBorderColorSelectedTab = 8
    ntCCForeColor = 9
    ntCCForeColorHighlighted = 10
    ntCCForeColorSelectedTab = 11
    ntCCIconColor = 12
    ntCCIconColorSelectedTab = 13
    ntCCIconColorMouseHover = 14
    ntCCIconColorMouseHoverSelectedTab = 15
    ntCCIconColorTabHighlighted = 16
End Enum

Public Enum NTAlignmentConstants
    ntAlignmentLeft = 0
    ntAlignmentRight = 1
    ntAlignmentCenter = 2
End Enum

' Events
' Original
Event Click(ByVal PreviousTab As Integer)
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Attribute Click.VB_UserMemId = -600
Event DblClick()
Attribute DblClick.VB_Description = "Occurs when you press and release a mouse button and then press and release it again over an object."
Attribute DblClick.VB_UserMemId = -601
Attribute DblClick.VB_MemberFlags = "200"
Event KeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer)
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Attribute KeyDown.VB_UserMemId = -602
Event KeyPress(ByVal KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Attribute KeyPress.VB_UserMemId = -603
Event KeyUp(ByVal KeyCode As Integer, ByVal Shift As Integer)
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Attribute KeyUp.VB_UserMemId = -604
Event MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Attribute MouseDown.VB_UserMemId = -605
Event MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Attribute MouseMove.VB_UserMemId = -606
Event MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Attribute MouseUp.VB_UserMemId = -607
Event OLECompleteDrag(Effect As Long)
Attribute OLECompleteDrag.VB_Description = "Occurs when a source component is dropped onto a target component, informing the source component that a drag action was either performed or canceled."
Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute OLEDragDrop.VB_Description = "Occurs when a source component is dropped onto a target component when the source component determines that a drop can occur."
Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
Attribute OLEDragOver.VB_Description = "Occurs when one component is dragged over another."
Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Attribute OLEGiveFeedback.VB_Description = "Occurs after every OLEDragOver event."
Event OLESetData(Data As DataObject, DataFormat As Integer)
Attribute OLESetData.VB_Description = "Occurs on an source component when a target component performs the GetData method on the sources DataObject object, but the data for the specified format has not yet been loaded."
Event OLEStartDrag(Data As DataObject, AllowedEffects As Long)
Attribute OLEStartDrag.VB_Description = "Occurs when a component's OLEDrag method is performed, or when a component initiates an OLE drag/drop operation when the OLEDragMode property is set to Automatic."

' Added
Event BeforeClick(ByVal CurrentSelectedTab As Long, ByRef NewSelectedTab As Long, ByRef Cancel As Boolean)
Event ChangeControlBackColor(ByVal ControlName As String, ByVal ControlTypeName As String, ByRef Cancel As Boolean)
Attribute ChangeControlBackColor.VB_Description = "When the ChangeControlsBackColor property is set to True, it allows you to individually determine which controls will (or will not) change their BackColor.\r\nThis event is raised for each control on the current tab, before the tab is painted."
Event ChangeControlForeColor(ByVal ControlName As String, ByVal ControlTypeName As String, ByRef Cancel As Boolean)
Attribute ChangeControlForeColor.VB_Description = "When the ChangeControlsBackColor property is set to True, it allows you to individually determine which controls will (or will not) change their ForeColor.\r\nThis event is raised for each control on the current tab, before the tab is painted."
Event RowsChange()
Attribute RowsChange.VB_Description = "Occurs when the Rows property changes its value."
Event Resize()
Attribute Resize.VB_Description = "Occurs when the tab body changes its size."
Event TabMouseEnter(ByVal nTab As Long)
Attribute TabMouseEnter.VB_Description = "Occurs when the mouse starts hovering on a tab."
Event TabMouseLeave(ByVal nTab As Long)
Attribute TabMouseLeave.VB_Description = "Occurs when the mouse ends hovering on a tab."
Event TabRightClick(ByVal nTab As Long, ByVal Shift As Long, ByVal X As Single, ByVal Y As Single)
Attribute TabRightClick.VB_Description = "Occurs when a click with the right mouse button takes places over a tab."
Event TabChange()
Attribute TabChange.VB_Description = "Occurs after the current tab has already changed."
Event IconClick(ByVal nTab As Long, ByRef ForwardClickToTab As Boolean)
Attribute IconClick.VB_Description = "Occurs when the icon of a tab is clicked (it doesn't work with pictures)."
Event BeforeTabReorder(ByVal CurrentIndex As Long, ByRef NewIndex As Long, ByRef Cancel As Boolean)
Attribute BeforeTabReorder.VB_Description = "Occurs when before a tab is changed from one position to another. The action can be canceled using the Cancel parameter or the new position can be altered from the NewIndex parameter."
Event TabReordered(ByVal CurrentIndex As Long, ByVal OldIndex As Long)
Attribute TabReordered.VB_Description = "Occurs when a tab changed its position."
Event IconMouseEnter(ByVal nTab As Long)
Attribute IconMouseEnter.VB_Description = "Occurs when the mouse enters hovering over a tab icon (not picture)."
Event IconMouseLeave(ByVal nTab As Long)
Attribute IconMouseLeave.VB_Description = "Occurs when the mouse goes out after hovering over a tab icon (not picture)."
Event TDIBeforeNewTab(ByVal TabType As NTTDINewTabTypeConstants, ByVal TabNumber As Long, ByRef TabCaption As String, ByRef LoadControls As Boolean, ByRef ShowTabCloseButton As Boolean, ByRef Cancel As Boolean)
Event TDINewTabAdded(ByVal TabNumber As Long)
Attribute TDINewTabAdded.VB_Description = "When in TDI mode, it occurs after a new tab was opened."
Event TDIBeforeClosingTab(ByVal TabNumber As Long, ByVal IsLastTab As Boolean, ByRef OpenNewOnLastClosed As Boolean, ByRef UnloadControls As Boolean, ByRef Cancel As Boolean)
Attribute TDIBeforeClosingTab.VB_Description = "When in TDI mode, it occurs before closing a tab."
Event TDITabClosed(ByVal TabNumber As Long, ByVal IsLastTab As Boolean)
Attribute TDITabClosed.VB_Description = "When in TDI mode, it occurs after a tab was closed."
Event TDIFormsShowTabIcon(ByVal TabNumber As Long, ByVal FormCaption As String, ByVal FormhWnd As Long, ByRef IconPicture As Object, ByRef IconLeftOffset As Long, ByRef IconTopOffset As Long, ByRef IconWidth As Long, ByRef IconHeight As Long, UseIconFont As Boolean, ByRef IconFont As Object, IconChar As Variant, IconColor As Long, IconColorHovered As Long, IconColorSelected As Long, NoIcon As Boolean)
Attribute TDIFormsShowTabIcon.VB_Description = "In TDI mode forms, it allows full control  for the icon that is displayed on the tab, including the icon itself and the size and position. It is also possible to use an icon font for the icon."

Private Type T_TabData
    ' Properties
    Caption As String
    Enabled As Boolean
    Visible As Boolean
    Picture As StdPicture
    Pic16 As StdPicture
    Pic20 As StdPicture
    Pic24 As StdPicture
    ToolTipText As String
    Controls As Collection
    ' Run time data
    TabRect As RECT
    PicToUse As StdPicture
    PicToUseSet As Boolean
    PicDisabled As StdPicture
    PicDisabledSet As Boolean
    Hovered As Boolean
    Selected As Boolean
    LeftTab As Boolean
    RightTab As Boolean
    TopTab As Boolean
    IconAndCaptionWidth As Long
    Row As Long
    RowPos As Long
    PosH As Long
    Width As Long
    IconFont As StdFont
    IconFontName As String
    DoNotUseIconFont As Boolean
    IconChar As Long
    IconLeftOffset As Long
    IconTopOffset As Long
    IconRect As RECT
    Tag As Variant
    Data As Long
    TDITabNumber As Long
    OriginalIndex As Long
    FixedWidth As Long
    CustomColors As New cTabColors
    Key As String
End Type

Private Type TDIFormIconCustomData
    Icon As StdPicture
    Left As Long
    Top As Long
    Width As Long
    Height As Long
    hWnd As Long
    TabPosString As String
    UseIconFont As Boolean
    IconFont As StdFont
    IconCharStr As String
    IconColor As Long
    IconColorMouseHover As Long
    IconColorSelected As Long
    LeftOffset As Long
    TopOffset As Long
End Type

Private Const cRowPerspectiveSpace As Long = 150&  ' in Twips

' Variables for properties
' Original
Private mBackColor As Long
Private WithEvents mFont As StdFont
Attribute mFont.VB_VarHelpID = -1
Private mEnabled As Boolean
Private mForeColor As Long
Private mUserControlHwnd As Long ' read only
Private mTab As Long
Private mTabsPerRow As Long
Private mTabs As Long
Private mRows As Long ' read only
Private mTabOrientation As NTTabOrientationConstants
Private mShowFocusRect As Boolean
Private mWordWrap As Boolean
Private mStyle As NTStyleConstants
Private mTabHeight As Single ' internally in Himetric
Private mTabMaxWidth As Single ' internally in Himetric
Private mMousePointer As NTMousePointerConstants
Private mMouseIcon As StdPicture
Private mOLEDropMode As NTOLEDropConstants
Private mTabData() As T_TabData
Private mRightToLeft As Boolean

' Added
Private mBackColorTabs As Long
Private mBackColorSelectedTab As Long
Private mForeColorSelectedTab As Long
Private mForeColorHighlighted As Long
Private mIconColor As Long
Private mIconColorSelectedTab As Long
Private mIconColorMouseHover As Long
Private mIconColorMouseHoverSelectedTab As Long
Private mIconColorTabHighlighted As Long
Private mHighlightColor As Long
Private mHighlightColorSelectedTab As Long
Private mFlatBarColorHighlight As Long
Private mFlatBarColorSelectedTab As Long
Private mFlatBarColorInactive As Long
Private mFlatTabsSeparationLineColor As Long
Private mFlatBodySeparationLineColor As Long
Private mFlatBorderColor As Long
Private mFlatTabBorderColorHighlight As Long
Private mFlatTabBorderColorSelectedTab As Long

Private mMaskColor As Long
Private mUseMaskColor As Boolean
Private mHighlightTabExtraHeight As Single ' internally  in Himetric
Private mHighlightEffect As Boolean
Private mVisualStyles As Boolean
Private mShowDisabledState As Boolean
Private mClientRect As RECT ' internally in Pixels, red only
Private mChangeControlsBackColor As Boolean
Private mChangeControlsForeColor As Boolean
Private mTabMinWidth As Single ' internally in Himetric
Private mTabWidthStyle As NTTabWidthStyleConstants
Private mShowRowsInPerspective As NTAutoYesNoConstants
Private mTabSeparation As Long
Private mTabAppearance As NTTabAppearanceConstants
Private mRedraw As Boolean
Private mIconAlignment As NTIconAlignmentConstants
Private mAutoRelocateControls As NTAutoRelocateControlsConstants
Private mEndOfTabs As Single
Private mSoftEdges As Boolean
Private mMinSizeNeeded As Single
Private mHandleHighContrastTheme As Boolean
Private mBackStyle As NTBackStyleConstants
Private mAutoTabHeight As Boolean
Private mOLEDropOnOtherTabs As Boolean
Private mTabTransition As NTTabTransitionConstants
Private mFlatRoundnessTop As Long
Private mFlatRoundnessBottom As Long
Private mFlatRoundnessTabs As Long
Private mTabMousePointerHand As Boolean
Private mHighlightMode As Long
Private mHighlightModeSelectedTab As Long
Private mFlatBorderMode As NTFlatBorderModeConstants
Private mFlatBarHeight As Long
Private mFlatBarGripHeight As Long
Private WithEvents mThemesCollection As NewTabThemes
Attribute mThemesCollection.VB_VarHelpID = -1
Private mCurrentThemeName As String
Private mCanReorderTabs As Boolean
Private mTDIMode As NTTDIModeConstants
Private mFlatBarPosition As NTFlatBarPositionConstants
Private mFlatBodySeparationLineHeight As Long
Private mSubclassingMethod As NTSubclassingMethodConstants
Private mOnlySubclassUserControl As Boolean
Private mTabsEndFreeSpace As Long
Private mUserControlSizeCorrectionsCounter_ScaleWidthNotToResize As Long
Private mUserControlSizeCorrectionsCounter_ScaleHeightNotToResize As Long
Private mCaptionAlignment As NTAlignmentConstants
Private mIsTDIFormsSampleInPropertyPage As Boolean

' Variables
Private mClientStart As Long ' in Pixels
Private mClientHeight As Long ' in Pixels
Private mClientWidth As Long ' in Pixels
Private mScaleWidth As Long
Private mScaleHeight As Long
Private mHasFocus As Boolean
Private mFormIsActive As Boolean
Private mDrawing As Boolean
Private mTabUnderMouse As Long
Private mAmbientUserMode As Boolean
Private mThereAreTabsToolTipTexts As Boolean
Private mDefaultTabHeight As Single  ' in Himetric
Private mPropertiesReady As Boolean
Private mButtonFace_H As Integer
Private mButtonFace_L As Integer
Private mButtonFace_S As Integer
Private mBodyThemedReady As Boolean
Private mInactiveBodyThemedReady As Boolean
Private mClientWidth_Prev As Long
Private mClientHeight_Prev As Long
Private mTheme As Long
Private mControlIsThemed As Boolean
Private mTabSeparation2 As Long
Private mThemeExtraDataAlreadySet As Boolean
Private mParentControlsTabStop As Collection
Private mParentControlsUseMnemonic As Collection
Private mContainedControlsThatAreContainers As Collection
Private mSubclassedControlsForPaintingHwnds As Collection
Private mSubclassedFramesHwnds As Collection
Private mSubclassedControlsForMoveHwnds As Collection
Private mTDISubclassedFormsHwnds As Collection
Private mTabStopsInitialized As Boolean
Private mAccessKeys As String
Private mAccessKeysSet As Boolean
Private mBlendDisablePicWithBackColorTabs_NotThemed As Boolean
Private mBlendDisablePicWithBackColorTabs_Themed As Boolean
Private mSubclassControlsPaintingPending As Boolean
Private mRepaintSubclassedControls As Boolean
Private mFormHwnd As Long
Private mBtnDown As Boolean
Private mTabAppearance2 As NTTabAppearanceConstants
Private mAppearanceIsPP As Boolean
Private mAppearanceIsFlat As Boolean
Private mNoActivate As Boolean
Private mCanPostDrawMessage As Boolean
Private mDrawMessagePosted As Boolean
Private mNeedToDraw As Boolean
Private mRows_Prev As Long
Private mChangedControlsBackColor As Boolean
Private mChangedControlsForeColor As Boolean
Private mLastContainedControlsString As String
Private mLastContainedControlsCount As Long
Private mLastContainedControlsPositionsStr As String
Private mBodyReset As Boolean
Private mSubclassed As Boolean
Private mClientStart_Prev As Long
Private mTabOrientation_Prev As NTTabOrientationConstants
Private WithEvents mForm As Form
Attribute mForm.VB_VarHelpID = -1
Private mFirstDraw As Boolean
Private mUserControlShown As Boolean
Private mClientRect_Prev As RECT
Private mEnsuringDrawn As Boolean
Private mDPIX As Long
Private mDPIY As Long
Private mXCorrection As Single
Private mYCorrection As Single
Private mHighlightEffectColors_Strong(10) As Long
Private mHighlightEffect_Step As Long
Private mGlowColor_Sel_Bk As Long
Private mGlowColor_Sel_Light As Long
Private mHighlightGradient As NTHighlightGradientConstants
Private mHighlightGradientSelectedTab As NTHighlightGradientConstants
Private mHighlightIntensity As NTHighlightIntensityConstants
Private mHighlightIntensitySelectedTab As NTHighlightIntensityConstants
Private mHighlightFlatBar As Boolean
Private mHighlightFlatBarWithGrip As Boolean
Private mHighlightFlatBarSelectedTab As Boolean
Private mHighlightFlatBarWithGripSelectedTab As Boolean
Private mHighlightAddExtraHeight As Boolean
Private mHighlightAddExtraHeightSelectedTab As Boolean
Private mHighlightFlatDrawBorder As Boolean
Private mHighlightFlatDrawBorderSelectedTab As Boolean
Private mHighlightCaptionBold As Boolean
Private mHighlightCaptionBoldSelectedTab As Boolean
Private mHighlightCaptionUnderlined As Boolean
Private mHighlightCaptionUnderlinedSelectedTab As Boolean
Private mCurrentTheme As NewTabTheme
Private mVisibleTabs As Long
Private mMouseX As Single
Private mMouseX2 As Single
Private mMouseY As Single
Private mMouseY2 As Single
Private mDraggingATab As Boolean
Private mTabChangedFromAnotherRow As Boolean
Private mProcessingTabChange As Boolean
Private mMouseIsOverIcon As Boolean
Private mMouseIsOverIcon_Tab As Long
Private mChangingSelectedTab As Boolean
Private mControlJustAdded As Boolean
Private mTDILastTabNumber As Long
Private mTDIControlNames() As String
Private mTDIControlNames_Count As Long
Private mInIDE As Boolean
Private mTDIIconColorMouseHover As Long
Private mTDIChangingTabCount As Boolean
Private mSettingTDIMode As Boolean
Private mTDIAddingNewTab As Boolean
Private mTDIClosingATab As Boolean
Private mSetAutoTabHeightPending As Boolean
Private mDPIScale As Single
Private mFlatBarGripHeightDPIScaled As Long
Private mFlatBarHeightDPIScaled As Long
Private mFlatBodySeparationLineHeightDPIScaled As Long
Private mFlatRoundnessTopDPIScaled As Long
Private mFlatRoundnessBottomDPIScaled As Long
Private mFlatRoundnessTabsDPIScaled As Long
Private mTabSeparationDPIScaled As Long
Private mTabIconDistanceToCaptionDPIScaled As Long
Private mIconClickExtendDPIScaled As Long
Private mMovingATab As Boolean
Private mPreviousTabBeforeDragging As Long
Private mToolTipEx As cToolTipEx
Private mTabsAreRotatedButCaptionsAreHorizontal As Boolean
Private mShowsRowsPerspective2 As Boolean
Private mResizeEventPending As Boolean
Private mTDIAddingNewTabForForm As Boolean
Private mTDIModeFormsUnhooked As Boolean
Private mTDIModeFormsFormData_FormHwnd() As Long
Private mTDIModeFormsFormData_OldParentHwnd() As Long
Private mTDIModeFormsFormData_FormIcon() As StdPicture
Private mTDIFormIconCustomData() As TDIFormIconCustomData
Private mTDIFormWithoutCloseButton() As Boolean
Private mShowingModalForm As Boolean
Private mAppDeactivated As Boolean
Private mUserControlSizeCorrectionsCounter As Long
Private mBackColorTabs_SavedWhileVisualStyles As Long
Private mBackColorSelectedTab_SavedWhileVisualStyles As Long
Private mBackColorTabsSavingWhileVisualStyles As Boolean
Private mAlignmentDTFlag As Long
Private mTDIModeAlreadySet As Boolean

Private mHighContrastThemeOn As Boolean
Private mHandleHighContrastTheme_OrigForeColor As Long
Private mHandleHighContrastTheme_OrigBackColorTabs As Long
Private mHandleHighContrastTheme_OrigForeColorSelectedTab As Long
Private mHandleHighContrastTheme_OrigForeColorHighlighted As Long
Private mHandleHighContrastTheme_OrigBackColorSelectedTab As Long
Private mHandleHighContrastTheme_OrigIconColor As Long
Private mHandleHighContrastTheme_OrigIconColorSelectedTab As Long
Private mHandleHighContrastTheme_OrigIconColorMouseHover As Long
Private mHandleHighContrastTheme_OrigIconColorMouseHoverSelectedTab As Long
Private mHandleHighContrastTheme_OrigIconColorTabHighlighted As Long
Private mHandleHighContrastTheme_OrigFlatTabBorderColorHighlight As Long
Private mHandleHighContrastTheme_OrigFlatTabBorderColorSelectedTab As Long

Private mBackColorSelectedTab_IsAutomatic As Boolean
Private mFlatBarColorHighlight_IsAutomatic As Boolean
Private mFlatBarColorHighlight_ColorAutomatic As Long
Private mHighlightColor_IsAutomatic As Boolean
Private mHighlightColor_ColorAutomatic As Long
Private mHighlightColorSelectedTab_IsAutomatic As Boolean
Private mHighlightColorSelectedTab_ColorAutomatic As Long
Private mFlatBarColorInactive_IsAutomatic As Boolean
Private mFlatBarColorInactive_ColorAutomatic As Long
Private mFlatTabsSeparationLineColor_IsAutomatic As Boolean
Private mFlatTabsSeparationLineColor_ColorAutomatic As Long
Private mFlatBodySeparationLineColor_IsAutomatic As Boolean
Private mFlatBodySeparationLineColor_ColorAutomatic As Long
Private mFlatBorderColor_IsAutomatic As Boolean
Private mFlatBorderColor_ColorAutomatic As Long

Private mBackColorIsFromAmbient As Boolean
Private mForeColorIsFromAmbient As Boolean
Private mBackColorTabsIsFromAmbient As Boolean
Private mIconColorIsFromAmbient As Boolean
Private mLeftOffsetToHide As Long
Private mLeftThresholdHided As Long
Private mPendingLeftOffset As Long
Private mUserControlTerminated As Boolean
Private mFlatBarHighlightEffectColors(10) As Long
Private mHighlightEffectColors_Light(10) As Long
Private mTabTransition_Step As Long
Private mFlatRoundnessTop2 As Long
Private mFlatRoundnessTabs2 As Long
Private mRightMostTabsRightPos() As Long
Private mTabWidthStyle2 As NTTabWidthStyleConstants
Private mCurrentMousePointerIsHand As Boolean
Private mHandIconHandle As Long
Private mDefaultIconFont As StdFont
Private mNoTabVisible As Boolean
Private mReSelTab As Boolean
Private WithEvents mTabIconFontsEventsHandler As cFontEventHandlers
Attribute mTabIconFontsEventsHandler.VB_VarHelpID = -1
Private mChangingHighContrastTheme As Boolean
Private mControlsForeColor_PrevColor As Long
Private mControlsBackColor_PrevColor As Long
Private mSettingBackColorSelectedTabFromBackColorTabs As Boolean

' Colors
Private m3DDKShadow As Long
Private m3DHighlight As Long
Private m3DShadow As Long
Private m3DDKShadow_Sel As Long
Private m3DHighlight_Sel As Long
Private m3DShadow_Sel As Long
Private mBackColorTabsDisabled As Long
Private mBackColorSelectedTabDisabled As Long
Private mGrayText As Long
Private mGrayText_Sel As Long
Private mGlowColor As Long
Private mGlowColor_Sel As Long
Private mBackColorTabs_R As Long
Private mBackColorTabs_G As Long
Private mBackColorTabs_B As Long
Private mBackColorSelectedTab_R As Long
Private mBackColorSelectedTab_G As Long
Private mBackColorSelectedTab_B As Long
Private mFlatBarGlowColor As Long
'Private mFlatGlowColor As Long
'Private mFlatGlowColor_Sel As Long

Private m3DShadowH As Long
Private m3DShadowV As Long
Private m3DShadowH_Sel As Long
Private m3DShadowV_Sel As Long
Private m3DHighlightH As Long
Private m3DHighlightV As Long
Private m3DHighlightH_Sel As Long
Private m3DHighlightV_Sel As Long
Private mBackColorTabs2 As Long
Private mBackColorSelectedTab2 As Long

' Themed extra data
Private mThemedInactiveReferenceBackColorTabs As Long
Private mThemedInactiveReferenceBackColorTabs_H As Integer
Private mThemedInactiveReferenceBackColorTabs_L As Integer
Private mThemedInactiveReferenceBackColorTabs_S As Integer
Private mThemedBodyReferenceTopBackColor As Long
Private mTABITEM_TopLeftCornerTransparencyMask(5) As Long
Private mTABITEM_TopRightCornerTransparencyMask(5) As Long
Private mTABITEMRIGHTEDGE_RightSideTransparencyMask(5) As Long
Private mThemedBodyBottomShadowPixels As Long
Private mThemedBodyRightShadowPixels As Long
Private mThemedBodyBackColor_R As Long
Private mThemedBodyBackColor_G As Long
Private mThemedBodyBackColor_B As Long


' Properties

' Returns/sets the background color.
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Colors"
Attribute BackColor.VB_UserMemId = -501
Attribute BackColor.VB_MemberFlags = "c"
    BackColor = mBackColor
End Property

Public Property Let BackColor(ByVal nValue As OLE_COLOR)
    If nValue <> mBackColor Then
        If Not IsValidOLE_COLOR(nValue) Then RaiseError 380, TypeName(Me): Exit Property
        mBackColorIsFromAmbient = (nValue = Ambient.BackColor)
        mBackColor = nValue
        SetPropertyChanged "BackColor"
        UserControl.BackColor = mBackColor
        ResetCachedThemeImages
        DrawDelayed
    End If
End Property


' Returns a Font object.
Public Property Get Font() As StdFont
Attribute Font.VB_Description = "Returns/sets the Font that will be used to draw the tab captions."
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute Font.VB_UserMemId = -512
    Set Font = mFont
End Property

Public Property Let Font(ByVal nValue As StdFont)
    Set Font = nValue
End Property

Public Property Set Font(ByVal nValue As StdFont)
    If Not nValue Is mFont Then
        Set mFont = nValue
        SetPropertyChanged "Font"
        SetFont
        mSetAutoTabHeightPending = True
        DrawDelayed
    End If
End Property


Public Property Get IconFont() As StdFont
Attribute IconFont.VB_Description = "Returns/sets the Font that will be used to draw the icon of the currently selected tab."
Attribute IconFont.VB_ProcData.VB_Invoke_Property = ";Icon"
    Set IconFont = TabIconFont(mTab)
End Property

Public Property Let IconFont(ByVal nValue As StdFont)
    Set TabIconFont(mTab) = nValue
End Property

Public Property Set IconFont(ByVal nValue As StdFont)
    Set TabIconFont(mTab) = nValue
End Property


Public Property Get TabIconFont(ByVal Index As Long) As StdFont
Attribute TabIconFont.VB_Description = "Returns/sets the Font that will be used to draw the icon in the tab pointed by the Index parameter."
Attribute TabIconFont.VB_ProcData.VB_Invoke_Property = ";Font"
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 381, TypeName(Me) ' invalid property array index
        Exit Property
    End If
    If mTabData(Index).IconFont Is Nothing Then
        Set mTabData(Index).IconFont = CloneFont(mDefaultIconFont)
        mTabIconFontsEventsHandler.AddFont mTabData(Index).IconFont, CLng(Index)
    End If
    Set TabIconFont = mTabData(Index).IconFont
End Property

Public Property Let TabIconFont(ByVal Index As Long, ByVal nValue As StdFont)
    Set TabIconFont(Index) = nValue
End Property

Public Property Set TabIconFont(ByVal Index As Long, ByVal nValue As StdFont)
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 381, TypeName(Me) ' invalid property array index
        Exit Property
    End If
    If Not nValue Is mTabData(Index).IconFont Then
        If Not mTabData(Index).IconFont Is Nothing Then mTabIconFontsEventsHandler.RemoveFont mTabData(Index).IconFont, CLng(Index)
        'If FontsAreEqual(nValue, mDefaultIconFont) Then Set nValue = Nothing
        Set mTabData(Index).IconFont = nValue
        If Not mTabData(Index).IconFont Is Nothing Then
            mTabIconFontsEventsHandler.AddFont mTabData(Index).IconFont, CLng(Index)
            mTabData(Index).IconFontName = nValue.Name
        Else
            mTabData(Index).IconFontName = ""
        End If
        mTabData(Index).DoNotUseIconFont = False
        SetPropertyChanged "IconFont"
        mSetAutoTabHeightPending = True
        DrawDelayed
    End If
End Property


Private Sub SetFont()
    On Error Resume Next
    If mFont Is Nothing Then
        Set mFont = Ambient.Font
    End If
    If mFont Is Nothing Then
        Set mFont = UserControl.Font
    End If
    Set UserControl.Font = mFont
    Set picDraw.Font = CloneFont(mFont)
    Set picAux.Font = picDraw.Font
    Err.Clear
End Sub

' Determines if the control is enabled.
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether a form or control can respond to user-generated events."
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
Attribute Enabled.VB_UserMemId = -514
    Enabled = mEnabled
End Property

Public Property Let Enabled(ByVal nValue As Boolean)
    Dim iRedraw As Boolean
    Dim iWv As Boolean
    
    If nValue <> mEnabled Then
        mEnabled = nValue
        UserControl.Enabled = mEnabled Or (Not mAmbientUserMode)
        SetPropertyChanged "Enabled"
        If mChangeControlsBackColor Then
            If mShowDisabledState Then
                mBodyReset = True
                iWv = IsWindowVisible(mUserControlHwnd) <> 0
                If iWv Then SendMessage mUserControlHwnd, WM_SETREDRAW, False, 0&
                If mEnabled Then
                    SetControlsBackColor mBackColorSelectedTab, mBackColorSelectedTabDisabled
                Else
                    SetControlsBackColor mBackColorSelectedTabDisabled, mBackColorSelectedTab
                End If
                If iWv Then
                    SendMessage mUserControlHwnd, WM_SETREDRAW, True, 0&
                    iRedraw = True
                End If
            End If
        End If
        mSubclassControlsPaintingPending = True
        mRepaintSubclassedControls = True
        DrawDelayed
        If iRedraw Then
            RedrawWindow mUserControlHwnd, ByVal 0&, 0&, RDW_INVALIDATE Or RDW_ALLCHILDREN
        End If
    End If
End Property

            
' Returns/sets the text color.
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the color used to draw the captions of inactive tabs. This setting in inherited by other foreground color properties if they are not set specifically."
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Colors"
Attribute ForeColor.VB_UserMemId = -513
Attribute ForeColor.VB_MemberFlags = "c"
    If mAmbientUserMode And mHandleHighContrastTheme And mHighContrastThemeOn Then
        ForeColor = mHandleHighContrastTheme_OrigForeColor
    Else
        ForeColor = mForeColor
    End If
End Property

Public Property Let ForeColor(ByVal nValue As OLE_COLOR)
    Dim iPrev As Long
    
    If nValue <> mForeColor Then
        If Not IsValidOLE_COLOR(nValue) Then RaiseError 380, TypeName(Me): Exit Property
        mForeColorIsFromAmbient = (nValue = Ambient.ForeColor)
        If mAmbientUserMode And mHandleHighContrastTheme And mHighContrastThemeOn Then
            mHandleHighContrastTheme_OrigForeColor = nValue
        Else
            iPrev = mForeColor
            mForeColor = nValue
            SetPropertyChanged "ForeColor"
            If ForeColorSelectedTab = iPrev Then
                ForeColorSelectedTab = nValue
            End If
            If ForeColorHighlighted = iPrev Then
                ForeColorHighlighted = nValue
            End If
            If IconColor = iPrev Then
                IconColor = nValue
            End If
            If FlatTabBorderColorHighlight = iPrev Then
                FlatTabBorderColorHighlight = nValue
            End If
            If FlatTabBorderColorSelectedTab = iPrev Then
                FlatTabBorderColorSelectedTab = nValue
            End If
            DrawDelayed
        End If
    End If
End Property


Public Property Get ForeColorSelectedTab() As OLE_COLOR
Attribute ForeColorSelectedTab.VB_Description = "Returns/sets the color used to draw the selected tab caption. "
Attribute ForeColorSelectedTab.VB_ProcData.VB_Invoke_Property = ";Colors"
    If mAmbientUserMode And mHandleHighContrastTheme And mHighContrastThemeOn Then
        ForeColorSelectedTab = mHandleHighContrastTheme_OrigForeColorSelectedTab
    Else
        ForeColorSelectedTab = mForeColorSelectedTab
    End If
End Property

Public Property Let ForeColorSelectedTab(ByVal nValue As OLE_COLOR)
    Dim iPrev As Long
    
    If nValue <> mForeColorSelectedTab Then
        If Not IsValidOLE_COLOR(nValue) Then RaiseError 380, TypeName(Me): Exit Property
        If mAmbientUserMode And mHandleHighContrastTheme And mHighContrastThemeOn Then
            mHandleHighContrastTheme_OrigForeColorSelectedTab = nValue
        Else
            iPrev = mForeColorSelectedTab
            mForeColorSelectedTab = nValue
            If IconColorSelectedTab = iPrev Then
                IconColorSelectedTab = nValue
            End If
            If mChangeControlsForeColor Then
                SetControlsProperForeColor IIf(mControlsForeColor_PrevColor = -1, iPrev, mControlsForeColor_PrevColor)
            End If
            If mTDIMode <> ntTDIModeNone Then
                If Not mAmbientUserMode Then lblTDILabel.ForeColor = mForeColorSelectedTab
            End If
            SetPropertyChanged "ForeColorSelectedTab"
            DrawDelayed
        End If
    End If
End Property

Private Sub SetControlsProperForeColor(Optional nPrevColor As Long = -1)
    Dim iForeColorSelectedTab As Long
    
    If IsEmpty(mTabData(mTab).CustomColors.ForeColorSelectedTab) Then
        iForeColorSelectedTab = mForeColorSelectedTab
    Else
        iForeColorSelectedTab = mTabData(mTab).CustomColors.ForeColorSelectedTab
    End If
    
    If ColorsHaveEnoughContrast(mBackColorSelectedTab, iForeColorSelectedTab) Then
        SetControlsForeColor iForeColorSelectedTab, nPrevColor
    ElseIf ColorsHaveEnoughContrast(mBackColorSelectedTab, mForeColor) Then
        SetControlsForeColor mForeColor, nPrevColor
    ElseIf ColorsHaveEnoughContrast(mBackColorSelectedTab, Ambient.ForeColor) Then
        SetControlsForeColor Ambient.ForeColor, nPrevColor
    ElseIf ColorsHaveEnoughContrast(mBackColorSelectedTab, iForeColorSelectedTab, 80) Then
        SetControlsForeColor iForeColorSelectedTab, nPrevColor
    ElseIf ColorsHaveEnoughContrast(mBackColorSelectedTab, mForeColor, 80) Then
        SetControlsForeColor mForeColor, nPrevColor
    ElseIf ColorsHaveEnoughContrast(mBackColorSelectedTab, Ambient.ForeColor, 80) Then
        SetControlsForeColor Ambient.ForeColor, nPrevColor
    End If
End Sub

Public Property Get ForeColorHighlighted() As OLE_COLOR
Attribute ForeColorHighlighted.VB_Description = "Returns/sets the color used to draw the captions of tabs when the are highlighted."
Attribute ForeColorHighlighted.VB_ProcData.VB_Invoke_Property = ";Colors"
    If mAmbientUserMode And mHandleHighContrastTheme And mHighContrastThemeOn Then
        ForeColorHighlighted = mHandleHighContrastTheme_OrigForeColorHighlighted
    Else
        ForeColorHighlighted = mForeColorHighlighted
    End If
End Property

Public Property Let ForeColorHighlighted(ByVal nValue As OLE_COLOR)
    If nValue <> mForeColorHighlighted Then
        If Not IsValidOLE_COLOR(nValue) Then RaiseError 380, TypeName(Me): Exit Property
        If mAmbientUserMode And mHandleHighContrastTheme And mHighContrastThemeOn Then
            mHandleHighContrastTheme_OrigForeColorHighlighted = nValue
        Else
            If IconColorTabHighlighted = ForeColorHighlighted Then
                IconColorTabHighlighted = nValue
            End If
            mForeColorHighlighted = nValue
            SetPropertyChanged "ForeColorHighlighted"
            'DrawDelayed
        End If
    End If
End Property


Public Property Get FlatTabBorderColorHighlight() As OLE_COLOR
Attribute FlatTabBorderColorHighlight.VB_Description = "Returns/sets the color that is used to highlight the tabs with a border when the mouse hovers on them, and the HighlightMode property has the corresponding flag ntHLFlatDrawBorder."
Attribute FlatTabBorderColorHighlight.VB_ProcData.VB_Invoke_Property = ";Colors"
    If mAmbientUserMode And mHandleHighContrastTheme And mHighContrastThemeOn Then
        FlatTabBorderColorHighlight = mHandleHighContrastTheme_OrigFlatTabBorderColorHighlight
    Else
        FlatTabBorderColorHighlight = mFlatTabBorderColorHighlight
    End If
End Property

Public Property Let FlatTabBorderColorHighlight(ByVal nValue As OLE_COLOR)
    If nValue <> mFlatTabBorderColorHighlight Then
        If Not IsValidOLE_COLOR(nValue) Then RaiseError 380, TypeName(Me): Exit Property
        If mAmbientUserMode And mHandleHighContrastTheme And mHighContrastThemeOn Then
            mHandleHighContrastTheme_OrigFlatTabBorderColorHighlight = nValue
        Else
            mFlatTabBorderColorHighlight = nValue
            SetPropertyChanged "FlatTabBorderColorHighlight"
            'DrawDelayed
        End If
    End If
End Property


Public Property Get FlatTabBorderColorSelectedTab() As OLE_COLOR
Attribute FlatTabBorderColorSelectedTab.VB_Description = "Returns/sets the color that is used to highlight the active tab with a border when the HighlightModeSelectedTab property has the corresponding flag ntHLFlatDrawBorder."
Attribute FlatTabBorderColorSelectedTab.VB_ProcData.VB_Invoke_Property = ";Colors"
    If mAmbientUserMode And mHandleHighContrastTheme And mHighContrastThemeOn Then
        FlatTabBorderColorSelectedTab = mHandleHighContrastTheme_OrigFlatTabBorderColorSelectedTab
    Else
        FlatTabBorderColorSelectedTab = mFlatTabBorderColorSelectedTab
    End If
End Property

Public Property Let FlatTabBorderColorSelectedTab(ByVal nValue As OLE_COLOR)
    If nValue <> mFlatTabBorderColorSelectedTab Then
        If Not IsValidOLE_COLOR(nValue) Then RaiseError 380, TypeName(Me): Exit Property
        If mAmbientUserMode And mHandleHighContrastTheme And mHighContrastThemeOn Then
            mHandleHighContrastTheme_OrigFlatTabBorderColorSelectedTab = nValue
        Else
            mFlatTabBorderColorSelectedTab = nValue
            SetPropertyChanged "FlatTabBorderColorSelectedTab"
            DrawDelayed
        End If
    End If
End Property


' Returns/sets the text displayed in the active tab.
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in the active tab."
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Caption.VB_UserMemId = -518
Attribute Caption.VB_MemberFlags = "c"
    Caption = mTabData(mTab).Caption
End Property

Public Property Let Caption(ByVal nValue As String)
    TabCaption(mTab) = nValue
End Property


' Returns the Windows handle of the control.
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns the Windows handle of the control."
Attribute hWnd.VB_UserMemId = -515
Attribute hWnd.VB_MemberFlags = "400"
    hWnd = mUserControlHwnd
End Property


' Returns/sets the number of tabs to appear on each row.
Public Property Get TabsPerRow() As Long
Attribute TabsPerRow.VB_Description = "Returns/sets the number of tabs to appear on each row. It only works in some styles."
Attribute TabsPerRow.VB_ProcData.VB_Invoke_Property = ";Tabs"
    TabsPerRow = mTabsPerRow
End Property

Public Property Let TabsPerRow(ByVal nValue As Long)
    If (nValue < 1) Then
        RaiseError 380, TypeName(Me) ' invalid property value
        Exit Property
    End If
    
    If nValue <> mTabsPerRow Then
        mTabsPerRow = nValue
        SetPropertyChanged "TabsPerRow"
        DrawDelayed
    End If
End Property


' Returns/sets the number of tabs.
Public Property Get Tabs() As Long
Attribute Tabs.VB_Description = "Returns/sets the number of tabs."
Attribute Tabs.VB_ProcData.VB_Invoke_Property = ";Tabs"
    Tabs = mTabs
End Property

Public Property Let Tabs(ByVal nValue As Long)
    Dim c As Long
    
    If (nValue < 1) Or (nValue > 1000) Then
        RaiseError 380, TypeName(Me) ' invalid property value
        Exit Property
    ElseIf mTDIMode <> ntTDIModeNone Then
        If Not mAmbientUserMode Then
            If Not mTDIChangingTabCount Then
                RaiseError 1380, TypeName(Me), "Can't change Tabs in TDI mode"
                Exit Property
            End If
        End If
    End If
    
    If nValue <> mTabs Then
        SetPropertyChanged "Tabs"
        mMouseIsOverIcon = False
        mMouseIsOverIcon_Tab = -1
        If mTabUnderMouse > -1 Then
            tmrTabMouseLeave.Enabled = False
            RaiseEvent_TabMouseLeave (mTabUnderMouse)
        End If
        mTabUnderMouse = -1
        For c = 0 To mTabs - 1
            mTabData(c).Hovered = False
        Next
        If tmrHighlightEffect.Enabled Then
            tmrHighlightEffect.Enabled = False
        End If
        mHighlightEffect_Step = 0
        If mHighlightIntensity = ntHighlightIntensityStrong Then
            mGlowColor = mHighlightEffectColors_Strong(10)
        Else
            mGlowColor = mHighlightEffectColors_Light(10)
        End If
        If mTabUnderMouse > -1 Then
            If Not IsEmpty(mTabData(mTabUnderMouse).CustomColors.HighlightColor) Then
                If mHighlightIntensity = ntHighlightIntensityStrong Then
                    mTabData(mTabUnderMouse).CustomColors.GlowColor = mTabData(mTabUnderMouse).CustomColors.HighlightEffectColors_Strong(10)
                Else
                    mTabData(mTabUnderMouse).CustomColors.GlowColor = mTabData(mTabUnderMouse).CustomColors.HighlightEffectColors_Light(10)
                End If
            End If
        End If
        If mTabs > nValue Then
            For c = nValue To mTabs - 1
                If mTabData(c).Controls.Count > 0 Then
                    On Error Resume Next
                    Err.Clear
                    Err.Raise 380  '  invalid property value
                    Dim iStr As String
                    iStr = Err.Description
                    On Error GoTo 0
                    RaiseError 380, TypeName(Me), iStr & ". Tab " & CStr(c) & " has controls, can't remove tabs with controls. Remove the contained controls first."
                    Exit Property
                ElseIf Not mTabData(c).IconFont Is Nothing Then
                    mTabIconFontsEventsHandler.RemoveFont mTabData(c).IconFont, c
                End If
            Next c
        End If
        If UBound(mTabData) = -1 Then
            ReDim mTabData(-1 To nValue - 1)
        Else
            ReDim Preserve mTabData(-1 To nValue - 1)
            If mTabs < nValue Then
                For c = mTabs To nValue - 1
                    mTabData(c).OriginalIndex = c
                Next
            End If
        End If
        If mTabs < nValue Then
            For c = mTabs To nValue - 1
                Set mTabData(c).Controls = New Collection
                mTabData(c).Enabled = True
                mTabData(c).Visible = True
                mTabData(c).Caption = "Tab " & CStr(c)
            Next
        End If
        mTabs = nValue
        If mTab > (mTabs - 1) Then
            mTab = (mTabs - 1)
        End If
        DrawDelayed
    End If
End Property


' Returns the number of rows of tabs.
Public Property Get Rows() As Long
Attribute Rows.VB_Description = "Returns the number of rows of tabs."
Attribute Rows.VB_ProcData.VB_Invoke_Property = ";Data"
Attribute Rows.VB_UserMemId = 0
Attribute Rows.VB_MemberFlags = "400"
    Rows = mRows
End Property


' included for backward compatibility
Public Property Get TabSel() As Long
Attribute TabSel.VB_Description = "Returns/sets the active tab number."
Attribute TabSel.VB_ProcData.VB_Invoke_Property = ";Tabs"
Attribute TabSel.VB_MemberFlags = "40"
    TabSel = mTab
End Property

Public Property Let TabSel(ByVal nValue As Long)
    Dim iPrev As Long
    Dim iPrev2 As Long
    Dim iCancel As Boolean
    Dim iWv As Boolean
    Dim c As Long
    Dim iDo As Boolean
    
    If (nValue < 0) Or (nValue >= mTabs) Then
        RaiseError 380, TypeName(Me) ' invalid property value
        Exit Property
    End If
    If Not mTabData(nValue).Visible Then
        RaiseError 380, TypeName(Me), "Invalid property value. Non-visible tabs can't be set as active/current."
        Exit Property
    End If
    
    If (nValue <> mTab) Or mReSelTab Then
        iDo = True
        If mTDIMode <> ntTDIModeNone Then
            If mTabData(nValue).Data = -1 Then
                If Not (mTDIClosingATab Or mTDIAddingNewTab) Then
                    If mAmbientUserMode Then TDIAddNewTab
                    iDo = False
                End If
            End If
        End If
        If iDo Then
            RaiseEvent BeforeClick(mTab, nValue, iCancel)
            If (nValue < 0) Or (nValue > (mTabs - 1)) Then nValue = mTab
        End If
        If (Not iCancel) And (nValue <> mTab) Then
            mChangingSelectedTab = True
            If Not mMovingATab Then
                If IsWindowVisible(mUserControlHwnd) <> 0 Then
                    If mTabTransition <> ntTransitionImmediate Then ShowPicCover
                End If
            End If
            iPrev = mTab
            mTab = nValue
            SetPropertyChanged "TabSel"
            If (iPrev >= 0) And (iPrev <= UBound(mTabData)) Then
                mTabData(iPrev).Selected = False
            End If
            If (mTab > -1) And (mTab < mTabs) Then
                mTabData(mTab).Selected = True
            End If
            iPrev2 = iPrev
            RaiseEvent Click(iPrev)
            iWv = IsWindowVisible(mUserControlHwnd) <> 0
            If iWv Then SendMessage mUserControlHwnd, WM_SETREDRAW, False, 0&
            SetVisibleControls iPrev2
            If iWv Then SendMessage mUserControlHwnd, WM_SETREDRAW, True, 0&
            mSubclassControlsPaintingPending = True
            If tmrHighlightEffect.Enabled Then
                tmrHighlightEffect.Enabled = False
            End If
            If (Not IsEmpty(mTabData(mTab).CustomColors.BackColorSelectedTab)) Or (Not IsEmpty(mTabData(mTab).CustomColors.HighlightColorSelectedTab)) Or (Not IsEmpty(mTabData(iPrev2).CustomColors.BackColorSelectedTab)) Or (Not IsEmpty(mTabData(iPrev2).CustomColors.HighlightColorSelectedTab)) Then
                SetColors
            ElseIf mAppearanceIsFlat And ((Not IsEmpty(mTabData(mTab).CustomColors.FlatBarColorSelectedTab)) Or (Not IsEmpty(mTabData(iPrev2).CustomColors.FlatBarColorSelectedTab))) Then
                SetColors
            End If
            If (Not IsEmpty(mTabData(mTab).CustomColors.ForeColorSelectedTab)) Or (Not IsEmpty(mTabData(iPrev2).CustomColors.ForeColorSelectedTab)) Then
                SetControlsProperForeColor
            End If
            If (Not IsEmpty(mTabData(mTab).CustomColors.BackColorSelectedTab)) Or (Not IsEmpty(mTabData(iPrev2).CustomColors.BackColorSelectedTab)) Then
                If IsEmpty(mTabData(mTab).CustomColors.BackColorSelectedTab) Then
                    SetControlsBackColor mBackColorSelectedTab
                Else
                    SetControlsBackColor mTabData(mTab).CustomColors.BackColorSelectedTab
                End If
            End If
            Draw
            If iWv Then RedrawWindow mUserControlHwnd, ByVal 0&, 0&, RDW_INVALIDATE Or RDW_ALLCHILDREN
            mChangingSelectedTab = False
            RaiseEvent TabChange
            
            If mTabUnderMouse > -1 Then
                tmrTabMouseLeave.Enabled = False
                RaiseEvent_TabMouseLeave (mTabUnderMouse)
            End If
            mTabUnderMouse = -1
            For c = 0 To mTabs - 1
                mTabData(c).Hovered = False
            Next
        End If
    End If
End Property


' Returns/sets a value that determines which side of the control the tabs will appear.
Public Property Get TabOrientation() As NTTabOrientationConstants
Attribute TabOrientation.VB_Description = "Returns/sets a value that determines on which side of the control the tabs will appear."
Attribute TabOrientation.VB_ProcData.VB_Invoke_Property = ";Appearance"
    TabOrientation = mTabOrientation
End Property

Public Property Let TabOrientation(ByVal nValue As NTTabOrientationConstants)
    If nValue < 0 Or nValue > 5 Then
        RaiseError 380, TypeName(Me) ' invalid property value
        Exit Property
    End If
    If mTDIMode <> ntTDIModeNone Then
        If (nValue <> ssTabOrientationTop) And (nValue <> ssTabOrientationBottom) Then
            RaiseError 1383, TypeName(Me), "In TDI mode only TabOrientation top and bottom are available."
            Exit Property
        End If
    End If
    If mRightToLeft Then
        If (nValue <> ssTabOrientationTop) Then
            RaiseError 1384, TypeName(Me), "RightToLeft is only available for Top orientation."
            Exit Property
        End If
    End If
    If nValue <> mTabOrientation Then
        mTabOrientation = nValue
        SetPropertyChanged "TabOrientation"
        ResetCachedThemeImages
        DrawDelayed
    End If
End Property


Public Property Get IconAlignment() As NTIconAlignmentConstants
Attribute IconAlignment.VB_Description = "Returns/sets the alignment of the icon (or picture) with respect of the tab caption."
Attribute IconAlignment.VB_ProcData.VB_Invoke_Property = ";Icon"
    IconAlignment = mIconAlignment
End Property

Public Property Let IconAlignment(ByVal nValue As NTIconAlignmentConstants)
    If nValue < ntIconAlignBeforeCaption Or nValue > ntIconAlignAtBottom Then
        RaiseError 380, TypeName(Me) ' invalid property value
        Exit Property
    End If
    If mTDIMode <> ntTDIModeNone Then
        RaiseError 1386, TypeName(Me), "Property not available in TDI mode."
        Exit Property
    End If
    If nValue <> mIconAlignment Then
        mIconAlignment = nValue
        mSetAutoTabHeightPending = True
        SetPropertyChanged "IconAlignment"
        DrawDelayed
    End If
End Property


' Specifies a bitmap to display on the current tab.
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Specifies a bitmap or icon to display on the current tab."
Attribute Picture.VB_ProcData.VB_Invoke_Property = ";Icon"
    Set Picture = TabPicture(mTab)
End Property

Public Property Let Picture(ByVal nValue As Picture)
    If Not nValue Is Nothing Then If nValue.Handle = 0 Then Set nValue = Nothing
    Set TabPicture(mTab) = nValue
End Property

Public Property Set Picture(ByVal nValue As Picture)
    If Not nValue Is Nothing Then If nValue.Handle = 0 Then Set nValue = Nothing
    Set TabPicture(mTab) = nValue
End Property


Public Property Get Pic16() As Picture
Attribute Pic16.VB_Description = "Specifies a bitmap to display on the current tab at 96 DPI, when the application is DPI aware."
Attribute Pic16.VB_ProcData.VB_Invoke_Property = ";Icon"
    Set Pic16 = TabPic16(mTab)
End Property

Public Property Let Pic16(ByVal nValue As Picture)
    If Not nValue Is Nothing Then If nValue.Handle = 0 Then Set nValue = Nothing
    Set TabPic16(mTab) = nValue
End Property

Public Property Set Pic16(ByVal nValue As Picture)
    If Not nValue Is Nothing Then If nValue.Handle = 0 Then Set nValue = Nothing
    Set TabPic16(mTab) = nValue
End Property


Public Property Get Pic20() As Picture
Attribute Pic20.VB_Description = "Specifies a bitmap to display on the current tab at 120 DPI, when the application is DPI aware."
Attribute Pic20.VB_ProcData.VB_Invoke_Property = ";Icon"
    Set Pic20 = TabPic20(mTab)
End Property

Public Property Let Pic20(ByVal nValue As Picture)
    If Not nValue Is Nothing Then If nValue.Handle = 0 Then Set nValue = Nothing
    Set TabPic20(mTab) = nValue
End Property

Public Property Set Pic20(ByVal nValue As Picture)
    If Not nValue Is Nothing Then If nValue.Handle = 0 Then Set nValue = Nothing
    Set TabPic20(mTab) = nValue
End Property


Public Property Get Pic24() As Picture
Attribute Pic24.VB_Description = "Specifies a bitmap to display on the current tab at 144 DPI, when the application is DPI aware."
Attribute Pic24.VB_ProcData.VB_Invoke_Property = ";Icon"
    Set Pic24 = TabPic24(mTab)
End Property

Public Property Let Pic24(ByVal nValue As Picture)
    If Not nValue Is Nothing Then If nValue.Handle = 0 Then Set nValue = Nothing
    Set TabPic24(mTab) = nValue
End Property

Public Property Set Pic24(ByVal nValue As Picture)
    If Not nValue Is Nothing Then If nValue.Handle = 0 Then Set nValue = Nothing
    Set TabPic24(mTab) = nValue
End Property


Public Property Get IconCharHex() As String
Attribute IconCharHex.VB_Description = "Returns/sets a string representing the hexadecimal value of the character that will be used as the icon in the currently selected tab."
Attribute IconCharHex.VB_ProcData.VB_Invoke_Property = "pagNewTabIcons;Icon"
    IconCharHex = TabIconCharHex(mTab)
End Property

Public Property Let IconCharHex(ByVal nValue As String)
    TabIconCharHex(mTab) = nValue
End Property


Public Property Get IconLeftOffset() As Long
Attribute IconLeftOffset.VB_Description = "Returns/sets the value in pixels of the offset for the left position when drawing the icon of the currently selected tab. It can be negative."
Attribute IconLeftOffset.VB_ProcData.VB_Invoke_Property = ";Icon"
    IconLeftOffset = TabIconLeftOffset(mTab)
End Property

Public Property Let IconLeftOffset(ByVal nValue As Long)
    TabIconLeftOffset(mTab) = nValue
End Property


Public Property Get IconTopOffset() As Long
Attribute IconTopOffset.VB_Description = "Returns/sets the value in pixels of the offset for the top position when drawing the icon of the currently selected tab. It can be negative."
Attribute IconTopOffset.VB_ProcData.VB_Invoke_Property = ";Icon"
    IconTopOffset = TabIconTopOffset(mTab)
End Property

Public Property Let IconTopOffset(ByVal nValue As Long)
    TabIconTopOffset(mTab) = nValue
End Property


' Determines whether a focus rectangle will be drawn in the caption when the control has the focus.
Public Property Get ShowFocusRect() As Boolean
Attribute ShowFocusRect.VB_Description = "Determines whether a focus rectangle will be drawn in the caption when the control has the focus."
Attribute ShowFocusRect.VB_ProcData.VB_Invoke_Property = ";Behavior"
    ShowFocusRect = mShowFocusRect
End Property

Public Property Let ShowFocusRect(ByVal nValue As Boolean)
    If nValue <> mShowFocusRect Then
        mShowFocusRect = nValue
        SetPropertyChanged "ShowFocusRect"
        DrawDelayed
    End If
End Property


' Determines whether text in the caption of each tab will wrap to the next line if it is too long.
Public Property Get WordWrap() As Boolean
Attribute WordWrap.VB_Description = "Determines whether text in the caption of each tab will wrap to the next line if it is too long."
Attribute WordWrap.VB_ProcData.VB_Invoke_Property = ";Behavior"
    WordWrap = mWordWrap
End Property

Public Property Let WordWrap(ByVal nValue As Boolean)
    If nValue <> mWordWrap Then
        mWordWrap = nValue
        SetPropertyChanged "WordWrap"
        DrawDelayed
    End If
End Property


' Returns/sets the style of the tabs.
Public Property Get Style() As NTStyleConstants
Attribute Style.VB_Description = "Returns/sets the style of the tabs."
Attribute Style.VB_ProcData.VB_Invoke_Property = ";Tabs"
    Style = mStyle
End Property

Public Property Let Style(ByVal nValue As NTStyleConstants)
    Dim iStyle As NTStyleConstants
    
    If nValue < ssStyleTabbedDialog Or nValue > ntStyleWindows Then
        RaiseError 380, TypeName(Me) ' invalid property value
        Exit Property
    End If
    iStyle = nValue
    If iStyle <> mStyle Then
        mStyle = iStyle
        SetPropertyChanged "Style"
        If mStyle <> ntStyleWindows Then VisualStyles = False
        If mBackColorSelectedTab_IsAutomatic Then mBackColorSelectedTab = GetAutomaticBackColorSelectedTab
        If mTabAppearance <> ntTAAuto Then
            mAppearanceIsFlat = (mTabAppearance = ntTAFlat)
        Else
            mAppearanceIsFlat = mStyle = ntStyleFlat
        End If
        SetHighlightMode
        mSetAutoTabHeightPending = True
        mShowFocusRect = (mStyle = ssStyleTabbedDialog) Or (mStyle = ssStylePropertyPage) Or (mStyle = ntStyleTabStrip)
        If mStyle = ntStyleWindows Then VisualStyles = True
        DrawDelayed
    End If
End Property


' Returns/sets the height of the tabs.
Public Property Get TabHeight() As Single
Attribute TabHeight.VB_Description = "Returns/sets the height of tabs."
Attribute TabHeight.VB_ProcData.VB_Invoke_Property = ";Tabs"
    TabHeight = FixRoundingError(ToContainerSizeY(mTabHeight, vbHimetric))
End Property

Public Property Let TabHeight(ByVal nValue As Single)
    Dim iValue As Single
    
    iValue = FromContainerSizeY(nValue, vbHimetric)
    If (iValue < 1) Or (pScaleY(iValue, vbHimetric, vbTwips) > UserControl.Height) Then
        'RaiseError 380, TypeName(Me) ' invalid property value
        'Exit Property
        iValue = pScaleY(UserControl.Height, vbTwips, vbHimetric)
    End If
    If pScaleY(iValue, vbHimetric, vbPixels) < 1 Then iValue = pScaleY(1, vbPixels, vbHimetric)
    If Round(iValue * 10000) <> Round(mTabHeight * 10000) Then
        If Abs(Round(iValue) - Round(mTabHeight)) > 1 Then
            If Round(iValue) <> Round(mDefaultTabHeight) Then
                mAutoTabHeight = False
                SetPropertyChanged "AutoTabHeight"
            End If
        End If
        mTabHeight = iValue
        If mHighlightTabExtraHeight > mTabHeight Then
            HighlightTabExtraHeight = mTabHeight
        End If
        SetPropertyChanged "TabHeight"
        ResetCachedThemeImages
        DrawDelayed
    End If
End Property


' Returns/sets the maximum width of each tab.
Public Property Get TabMaxWidth() As Single
Attribute TabMaxWidth.VB_Description = "Returns/sets the maximum width of each tab."
Attribute TabMaxWidth.VB_ProcData.VB_Invoke_Property = ";Tabs"
    TabMaxWidth = FixRoundingError(ToContainerSizeX(mTabMaxWidth, vbHimetric))
End Property

Public Property Let TabMaxWidth(ByVal nValue As Single)
    Dim iValue As Single
    
    iValue = FromContainerSizeX(nValue, vbHimetric)
    If ((iValue < pScaleX(10, vbPixels, vbHimetric)) And Not iValue = 0) Or (pScaleX(iValue, vbHimetric, vbTwips) > IIf(UserControl.Width > 3000, UserControl.Width, 3000)) Then
        RaiseError 380, TypeName(Me) ' invalid property value
        Exit Property
    End If
    If Round(iValue * 10000) <> Round(mTabMaxWidth * 10000) Then
        mTabMaxWidth = iValue
        If mTabMaxWidth <> 0 Then
            If mTabMaxWidth < mTabMinWidth Then
                TabMinWidth = ToContainerSizeY(mTabMaxWidth, vbHimetric)
            End If
        End If
        SetPropertyChanged "TabMaxWidth"
        DrawDelayed
    End If
End Property


' Returns/sets the minimun width of each tab.
Public Property Get TabMinWidth() As Single
Attribute TabMinWidth.VB_Description = "Returns/sets the minimun width of each tab."
Attribute TabMinWidth.VB_ProcData.VB_Invoke_Property = ";Tabs"
    TabMinWidth = FixRoundingError(ToContainerSizeX(mTabMinWidth, vbHimetric))
End Property

Public Property Let TabMinWidth(ByVal nValue As Single)
    Dim iValue As Single
    
    iValue = FromContainerSizeX(nValue, vbHimetric)
    If (iValue < 0) Or (pScaleX(iValue, vbHimetric, vbTwips) > IIf(UserControl.Width > 3000, UserControl.Width, 3000)) Then
        RaiseError 380, TypeName(Me) ' invalid property value
        Exit Property
    End If
    If Round(iValue * 10000) <> Round(mTabMinWidth * 10000) Then
        mTabMinWidth = iValue
        If (mTabMinWidth > mTabMaxWidth) And (mTabMaxWidth <> 0) Then
            TabMaxWidth = ToContainerSizeY(mTabMinWidth, vbHimetric)
        End If
        SetPropertyChanged "TabMinWidth"
        DrawDelayed
    End If
End Property


Public Property Get TabWidthStyle() As NTTabWidthStyleConstants
Attribute TabWidthStyle.VB_Description = "Returns or sets a value that determines the justification or width of the tabs."
Attribute TabWidthStyle.VB_ProcData.VB_Invoke_Property = ";Tabs"
    TabWidthStyle = mTabWidthStyle
End Property

Public Property Let TabWidthStyle(ByVal nValue As NTTabWidthStyleConstants)
    If nValue < ntTWTabStripEmulation Or nValue > ntTWTabCaptionWidthFillRows Then
        RaiseError 380, TypeName(Me) ' invalid property value
        Exit Property
    End If
    If mTabWidthStyle <> nValue Then
        mTabWidthStyle = nValue
        SetPropertyChanged "TabWidthStyle"
        DrawDelayed
    End If
End Property


Public Property Get TabAppearance() As NTTabAppearanceConstants
Attribute TabAppearance.VB_Description = "Returns/sets a value that determines the appearance of the tabs."
Attribute TabAppearance.VB_ProcData.VB_Invoke_Property = ";Appearance"
    TabAppearance = mTabAppearance
End Property

Public Property Let TabAppearance(ByVal nValue As NTTabAppearanceConstants)
    If nValue < 0 Or nValue > 6 Then
        RaiseError 380, TypeName(Me) ' invalid property value
        Exit Property
    End If
    If mTabAppearance <> nValue Then
        mTabAppearance = nValue
        If mTabAppearance <> ntTAAuto Then
            mAppearanceIsFlat = (mTabAppearance = ntTAFlat)
        Else
            mAppearanceIsFlat = mStyle = ntStyleFlat
        End If
        SetPropertyChanged "TabAppearance"
        ResetCachedThemeImages
        DrawDelayed
    End If
End Property



' Returns/sets the type of mouse pointer displayed when over the control.
Public Property Get MousePointer() As NTMousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over the control."
Attribute MousePointer.VB_ProcData.VB_Invoke_Property = ";Misc"
    MousePointer = mMousePointer
End Property

Public Property Let MousePointer(ByVal nValue As NTMousePointerConstants)
    Select Case nValue
        Case Is < 0, 16 To 98, Is > 99
            RaiseError 380, TypeName(Me) ' invalid property value
            Exit Property
    End Select
    If nValue <> mMousePointer Then
        mMousePointer = nValue
        UserControl.MousePointer = mMousePointer
        SetPropertyChanged "MousePointer"
    End If
End Property


' Returns/sets the icon used as the mouse pointer when the MousePointer property is set to 99 (custom).
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Returns/sets the icon used as the mouse pointer when the MousePointer property is set to 99 (custom)."
Attribute MouseIcon.VB_ProcData.VB_Invoke_Property = ";Misc"
    Set MouseIcon = mMouseIcon
End Property

Public Property Let MouseIcon(ByVal nValue As Picture)
    If Not nValue Is Nothing Then If nValue.Handle = 0 Then Set nValue = Nothing
    If Not nValue Is mMouseIcon Then
        Set mMouseIcon = nValue
        SetPropertyChanged "MouseIcon"
    End If
End Property

Public Property Set MouseIcon(ByVal nValue As Picture)
    If Not nValue Is Nothing Then If nValue.Handle = 0 Then Set nValue = Nothing
    If Not nValue Is mMouseIcon Then
        Set mMouseIcon = nValue
        Set UserControl.MouseIcon = mMouseIcon
        SetPropertyChanged "MouseIcon"
    End If
End Property


' Returns/Sets whether this control can act as an OLE drop target.
Public Property Get OLEDropMode() As NTOLEDropConstants
Attribute OLEDropMode.VB_Description = "Returns/sets how a target component handles drop operations."
Attribute OLEDropMode.VB_ProcData.VB_Invoke_Property = ";Behavior"
    OLEDropMode = mOLEDropMode
End Property

Public Property Let OLEDropMode(ByVal nValue As NTOLEDropConstants)
    Const DRAGDROP_E_ALREADYREGISTERED As Long = &H80040101
    
    If nValue < 0 Or nValue > 1 Then
        RaiseError 380, TypeName(Me) ' invalid property value
        Exit Property
    End If
    If nValue <> mOLEDropMode Then
        mOLEDropMode = nValue
        
        On Error Resume Next
        UserControl.OLEDropMode = mOLEDropMode
        If Err.Number = DRAGDROP_E_ALREADYREGISTERED Then
            RevokeDragDrop UserControl.hWnd
            UserControl.OLEDropMode = mOLEDropMode
        End If
        On Error GoTo 0
        
        SetPropertyChanged "OLEDropMode"
    End If
End Property


' Returns the picture displayed on the specified tab.
Public Property Get TabPicture(ByVal Index As Long) As Picture
Attribute TabPicture.VB_Description = "Returns/sets the picture to be displayed on the specified tab."
Attribute TabPicture.VB_ProcData.VB_Invoke_Property = ";Icon"
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 381, TypeName(Me) ' invalid property array index
        Exit Property
    End If
    Set TabPicture = mTabData(Index).Picture
End Property

Public Property Let TabPicture(ByVal Index As Long, ByVal nValue As Picture)
    If Not nValue Is Nothing Then If nValue.Handle = 0 Then Set nValue = Nothing
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 381, TypeName(Me) ' invalid property array index
        Exit Property
    End If
    Set TabPicture(Index) = nValue
End Property

Public Property Set TabPicture(ByVal Index As Long, ByVal nValue As Picture)
    If Not nValue Is Nothing Then If nValue.Handle = 0 Then Set nValue = Nothing
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 381, TypeName(Me) ' invalid property array index
        Exit Property
    End If
    If Not nValue Is mTabData(Index).Picture Then
        Set mTabData(Index).Picture = nValue
        mTabData(Index).PicToUseSet = False
        mTabData(Index).PicDisabledSet = False
        SetPropertyChanged "TabPicture"
        mSetAutoTabHeightPending = True
        DrawDelayed
    End If
End Property


Public Property Get TabPic16(ByVal Index As Long) As Picture
Attribute TabPic16.VB_Description = "Specifies a bitmap to display on the specified tab at 96 DPI, when the application is DPI aware."
Attribute TabPic16.VB_ProcData.VB_Invoke_Property = ";Icon"
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 381, TypeName(Me) ' invalid property array index
        Exit Property
    End If
    Set TabPic16 = mTabData(Index).Pic16
End Property

Public Property Let TabPic16(ByVal Index As Long, ByVal nValue As Picture)
    If Not nValue Is Nothing Then If nValue.Handle = 0 Then Set nValue = Nothing
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 381, TypeName(Me) ' invalid property array index
        Exit Property
    End If
    Set TabPic16(Index) = nValue
End Property

Public Property Set TabPic16(ByVal Index As Long, ByVal nValue As Picture)
    If Not nValue Is Nothing Then If nValue.Handle = 0 Then Set nValue = Nothing
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 381, TypeName(Me) ' invalid property array index
        Exit Property
    End If
    If Not nValue Is mTabData(Index).Pic16 Then
        Set mTabData(Index).Pic16 = nValue
        mTabData(Index).PicToUseSet = False
        mTabData(Index).PicDisabledSet = False
        SetPropertyChanged "TabPic16"
        mSetAutoTabHeightPending = True
        DrawDelayed
    End If
End Property


Public Property Get TabPic20(ByVal Index As Long) As Picture
Attribute TabPic20.VB_Description = "Specifies a bitmap to display on the specified tab at 120 DPI, when the application is DPI aware."
Attribute TabPic20.VB_ProcData.VB_Invoke_Property = ";Icon"
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 381, TypeName(Me) ' invalid property array index
        Exit Property
    End If
    Set TabPic20 = mTabData(Index).Pic20
End Property

Public Property Let TabPic20(ByVal Index As Long, ByVal nValue As Picture)
    If Not nValue Is Nothing Then If nValue.Handle = 0 Then Set nValue = Nothing
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 381, TypeName(Me) ' invalid property array index
        Exit Property
    End If
    Set TabPic20(Index) = nValue
End Property

Public Property Set TabPic20(ByVal Index As Long, ByVal nValue As Picture)
    If Not nValue Is Nothing Then If nValue.Handle = 0 Then Set nValue = Nothing
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 381, TypeName(Me) ' invalid property array index
        Exit Property
    End If
    If Not nValue Is mTabData(Index).Pic20 Then
        Set mTabData(Index).Pic20 = nValue
        mTabData(Index).PicToUseSet = False
        mTabData(Index).PicDisabledSet = False
        SetPropertyChanged "TabPic20"
        mSetAutoTabHeightPending = True
        DrawDelayed
    End If
End Property


Public Property Get TabPic24(ByVal Index As Long) As Picture
Attribute TabPic24.VB_Description = "Specifies a bitmap to display on the specified tab at 144 DPI, when the application is DPI aware."
Attribute TabPic24.VB_ProcData.VB_Invoke_Property = ";Icon"
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 381, TypeName(Me) ' invalid property array index
        Exit Property
    End If
    Set TabPic24 = mTabData(Index).Pic24
End Property

Public Property Let TabPic24(ByVal Index As Long, ByVal nValue As Picture)
    If Not nValue Is Nothing Then If nValue.Handle = 0 Then Set nValue = Nothing
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 381, TypeName(Me) ' invalid property array index
        Exit Property
    End If
    Set TabPic24(Index) = nValue
End Property

Public Property Set TabPic24(ByVal Index As Long, ByVal nValue As Picture)
    If Not nValue Is Nothing Then If nValue.Handle = 0 Then Set nValue = Nothing
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 381, TypeName(Me) ' invalid property array index
        Exit Property
    End If
    If Not nValue Is mTabData(Index).Pic24 Then
        Set mTabData(Index).Pic24 = nValue
        mTabData(Index).PicToUseSet = False
        mTabData(Index).PicDisabledSet = False
        SetPropertyChanged "TabPic24"
        mSetAutoTabHeightPending = True
        DrawDelayed
    End If
End Property


Public Property Get TabIconCharHex(ByVal Index As Long) As String
Attribute TabIconCharHex.VB_Description = "Returns/sets a string representing the hexadecimal value of the character that will be used as the icon for the tab pointed by the Index parameter."
Attribute TabIconCharHex.VB_ProcData.VB_Invoke_Property = ";Icon"
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 381, TypeName(Me) ' invalid property array index
        Exit Property
    End If
    If mTabData(Index).IconChar <> 0 Then
        TabIconCharHex = "&H" & Hex$(mTabData(Index).IconChar)
    Else
        If Not mAmbientUserMode Then
            TabIconCharHex = "&H[Font Hex code here]"
        End If
    End If
End Property

Public Property Let TabIconCharHex(ByVal Index As Long, ByVal nValue As String)
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 381, TypeName(Me) ' invalid property array index
        Exit Property
    End If
    SetPropertyChanged "TabIconCharHex"
    SetPropertyChanged "IconCharHex"
    If Left$(nValue, 2) <> "&H" Then
        nValue = "&H" & nValue
    End If
    If Right$(nValue, 1) <> "&" Then
        nValue = nValue & "&"
    End If
    
    If Val(nValue) <> mTabData(Index).IconChar Then
        mTabData(Index).IconChar = Val(nValue)
        mSetAutoTabHeightPending = True
    End If
    mTabData(Index).DoNotUseIconFont = False
    DrawDelayed
End Property


Public Property Get TabIconLeftOffset(ByVal Index As Long) As Long
Attribute TabIconLeftOffset.VB_Description = "Returns/sets the value in pixels of the offset for the left position when drawing the icon of the tab pointed by the Index parameter. It can be negative."
Attribute TabIconLeftOffset.VB_ProcData.VB_Invoke_Property = ";Font"
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 381, TypeName(Me) ' invalid property array index
        Exit Property
    End If
    TabIconLeftOffset = mTabData(Index).IconLeftOffset
End Property

Public Property Let TabIconLeftOffset(ByVal Index As Long, ByVal nValue As Long)
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 381, TypeName(Me) ' invalid property array index
        Exit Property
    End If
    SetPropertyChanged "TabIconLeftOffset"
    SetPropertyChanged "IconLeftOffset"
    If Val(nValue) <> mTabData(Index).IconLeftOffset Then
        mTabData(Index).IconLeftOffset = Val(nValue)
        mSetAutoTabHeightPending = True
    End If
    DrawDelayed
End Property


Public Property Get TabIconTopOffset(ByVal Index As Long) As Long
Attribute TabIconTopOffset.VB_Description = "Returns/sets the value in pixels of the offset for the top position when drawing the icon of the tab pointed by the Index parameter. It can be negative."
Attribute TabIconTopOffset.VB_ProcData.VB_Invoke_Property = ";Font"
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 381, TypeName(Me) ' invalid property array index
        Exit Property
    End If
    TabIconTopOffset = mTabData(Index).IconTopOffset
End Property

Public Property Let TabIconTopOffset(ByVal Index As Long, ByVal nValue As Long)
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 381, TypeName(Me) ' invalid property array index
        Exit Property
    End If
    SetPropertyChanged "TabIconTopOffset"
    SetPropertyChanged "IconTopOffset"
    If Val(nValue) <> mTabData(Index).IconTopOffset Then
        mTabData(Index).IconTopOffset = Val(nValue)
        mSetAutoTabHeightPending = True
    End If
    DrawDelayed
End Property


' Determines if the specified tab is visible.
Public Property Get TabVisible(ByVal Index As Long) As Boolean
Attribute TabVisible.VB_Description = "Determines if the specified tab is visible."
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 381, TypeName(Me) ' invalid property array index
        Exit Property
    End If
    TabVisible = mTabData(Index).Visible
End Property

Public Property Let TabVisible(ByVal Index As Long, ByVal nValue As Boolean)
    Dim c As Long
    Dim iSetSelectedTab As Boolean
    
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 381, TypeName(Me) ' invalid property array index
        Exit Property
    End If
    If nValue <> mTabData(Index).Visible Then
        If mNoTabVisible Then
            If nValue Then
                iSetSelectedTab = True
            End If
        End If
        If mTab = Index Then
            c = mTab - 1
            Do Until c < 0
                If mTabData(c).Visible And mTabData(c).Enabled Then
                    Exit Do
                End If
                c = c - 1
            Loop
            If c = -1 Then
                c = mTab + 1
                Do Until c = mTabs
                    If mTabData(c).Visible And mTabData(c).Enabled Then
                        Exit Do
                    End If
                    c = c + 1
                Loop
            End If
            If (c < 0) Or (c > (mTabs - 1)) Then
                c = mTab - 1
                Do Until c < 0
                    If mTabData(c).Visible Then
                        Exit Do
                    End If
                    c = c - 1
                Loop
                If c = -1 Then
                    c = mTab + 1
                    Do Until c = mTabs
                        If mTabData(c).Visible Then
                            Exit Do
                        End If
                        c = c + 1
                    Loop
                End If
            End If
            If (c > -1) And (c < mTabs) Then
                TabSel = c
                If mTab = c Then ' the change could had been canceled through the BeforeClick event, in that case TabSel woudn't change
                    mTabData(Index).Visible = nValue
                    mTabData(Index).Selected = False
                End If
            Else
        '        mTab = -1
                mNoTabVisible = True
                mTabData(Index).Visible = nValue
                mTabData(Index).Selected = False
                HideAllContainedControls
            End If
            If iSetSelectedTab Then
                mReSelTab = True
                TabSel = Index
                mReSelTab = False
            End If
        Else
            mTabData(Index).Visible = nValue
            If (mTab < 0) Or (mTab > (mTabs - 1)) Then
                TabSel = Index
                mTabData(Index).Selected = True
            End If
        End If
        mAccessKeysSet = False
        SetPropertyChanged "TabVisible"
        mBodyReset = True
        DrawDelayed
    End If
End Property


' Determines if the specified tab is enabled.
Public Property Get TabEnabled(ByVal Index As Long) As Boolean
Attribute TabEnabled.VB_Description = "Determines if the specified tab is enabled."
Attribute TabEnabled.VB_ProcData.VB_Invoke_Property = ";Appearance"
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 381, TypeName(Me) ' invalid property array index
        Exit Property
    End If
    TabEnabled = mTabData(Index).Enabled
End Property

Public Property Let TabEnabled(ByVal Index As Long, ByVal nValue As Boolean)
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 381, TypeName(Me) ' invalid property array index
        Exit Property
    End If
    If nValue <> mTabData(Index).Enabled Then
        mTabData(Index).Enabled = nValue
        SetPropertyChanged "TabEnabled"
        mAccessKeysSet = False
        DrawDelayed
    End If
End Property


' Returns the text displayed on the specified tab.
Public Property Get TabCaption(ByVal Index As Long) As String
Attribute TabCaption.VB_Description = "Returns/sets  the text displayed on the specified tab."
Attribute TabCaption.VB_ProcData.VB_Invoke_Property = ";Appearance"
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 381, TypeName(Me) ' invalid property array index
        Exit Property
    End If
    TabCaption = mTabData(Index).Caption
End Property

Public Property Let TabCaption(ByVal Index As Long, ByVal nValue As String)
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 381, TypeName(Me) ' invalid property array index
        Exit Property
    End If
    If nValue <> mTabData(Index).Caption Then
        mTabData(Index).Caption = nValue
        SetPropertyChanged "TabCaption"
        mAccessKeysSet = False
        DrawDelayed
    End If
End Property


Public Property Get TabToolTipText(ByVal Index As Long) As String
Attribute TabToolTipText.VB_Description = "Returns/sets the text that will be shown as tooltip text when the mouse pointer is over the specified tab."
Attribute TabToolTipText.VB_ProcData.VB_Invoke_Property = ";Misc"
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 381, TypeName(Me) ' invalid property array index
        Exit Property
    End If
    TabToolTipText = mTabData(Index).ToolTipText
End Property

Public Property Let TabToolTipText(ByVal Index As Long, ByVal nValue As String)
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 381, TypeName(Me) ' invalid property array index
        Exit Property
    End If
    If nValue <> mTabData(Index).ToolTipText Then
        mTabData(Index).ToolTipText = nValue
        CheckIfThereAreTabsToolTipTexts
        If Index = mTabUnderMouse Then
            If mTabData(Index).ToolTipText <> "" Then
                ShowTabTTT mTabUnderMouse
            Else
                tmrShowTabTTT.Enabled = False
                Set mToolTipEx = Nothing
            End If
        End If
        SetPropertyChanged "TabToolTipText"
    End If
End Property


Public Property Get MaskColor() As OLE_COLOR
Attribute MaskColor.VB_Description = "Returns/sets a color in the tabs pictures to be a mask (that is, transparent)."
Attribute MaskColor.VB_ProcData.VB_Invoke_Property = ";Colors"
    MaskColor = mMaskColor
End Property

Public Property Let MaskColor(ByVal nValue As OLE_COLOR)
    If nValue <> mMaskColor Then
        If Not IsValidOLE_COLOR(nValue) Then RaiseError 380, TypeName(Me): Exit Property
        mMaskColor = nValue
        SetPropertyChanged "MaskColor"
        DrawDelayed
    End If
End Property


Public Property Get HighlightTabExtraHeight() As Single
Attribute HighlightTabExtraHeight.VB_Description = "Returns/sets a value that determines the extra height that a tab will have when it is highlighted. This value is in units of the container."
Attribute HighlightTabExtraHeight.VB_ProcData.VB_Invoke_Property = ";Appearance"
    HighlightTabExtraHeight = FixRoundingError(ToContainerSizeY(mHighlightTabExtraHeight, vbHimetric))
End Property

Public Property Let HighlightTabExtraHeight(ByVal nValue As Single)
    Dim iValue As Single
    
    iValue = FromContainerSizeY(nValue, vbHimetric)
    If iValue < 0 Then
        RaiseError 380, TypeName(Me) ' invalid property value
        Exit Property
    End If
    If iValue > mTabHeight Then
        iValue = mTabHeight 'limit
    End If
    If Round(iValue * 10000) <> Round(mHighlightTabExtraHeight * 10000) Then
        mHighlightTabExtraHeight = iValue
        SetPropertyChanged "HighlightTabExtraHeight"
        DrawDelayed
    End If
End Property


Public Property Get HighlightEffect() As Boolean
Attribute HighlightEffect.VB_Description = "Returns/sets a value that determines whether tabs will display a progressive effect when they are highlighted due to hovering over them."
Attribute HighlightEffect.VB_ProcData.VB_Invoke_Property = ";Appearance"
    HighlightEffect = mHighlightEffect
End Property

Public Property Let HighlightEffect(ByVal nValue As Boolean)
    If nValue <> mHighlightEffect Then
        mHighlightEffect = nValue
        SetPropertyChanged "HighlightEffect"
    End If
End Property


Public Property Get VisualStyles() As Boolean
Attribute VisualStyles.VB_Description = "Returns/sets a value that determines whether the appearance of the control will use Windows visual styles-"
Attribute VisualStyles.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute VisualStyles.VB_MemberFlags = "40"
    VisualStyles = mVisualStyles
End Property

Public Property Let VisualStyles(ByVal nValue As Boolean)
    Dim iWv As Boolean
    
    If nValue <> mVisualStyles Then
        If nValue Then
            mBackColorTabs_SavedWhileVisualStyles = mBackColorTabs
            mBackColorSelectedTab_SavedWhileVisualStyles = mBackColorSelectedTab
            BackColorTabs = vbButtonFace
            BackColorSelectedTab = vbButtonFace
            mBackColorTabsSavingWhileVisualStyles = True
        Else
            If mBackColorTabsSavingWhileVisualStyles Then
                mBackColorTabsSavingWhileVisualStyles = False
                BackColorTabs = mBackColorTabs_SavedWhileVisualStyles
                BackColorSelectedTab = mBackColorSelectedTab_SavedWhileVisualStyles
            End If
        End If
        mVisualStyles = nValue
        SetPropertyChanged "VisualStyles"
        mSubclassControlsPaintingPending = True
        mRepaintSubclassedControls = True
        iWv = IsWindowVisible(mUserControlHwnd) <> 0
        If iWv Then SendMessage mUserControlHwnd, WM_SETREDRAW, False, 0&
        Draw
        If iWv Then SendMessage mUserControlHwnd, WM_SETREDRAW, True, 0&
        If iWv Then RedrawWindow mUserControlHwnd, ByVal 0&, 0&, RDW_INVALIDATE Or RDW_ALLCHILDREN
        'If mVisualStyles Then Style = ntStyleWindows
    End If
End Property


Public Property Get BackColorTabs() As OLE_COLOR
Attribute BackColorTabs.VB_Description = "Returns/sets the background color of the tabs."
Attribute BackColorTabs.VB_ProcData.VB_Invoke_Property = ";Colors"
    If mAmbientUserMode And mHandleHighContrastTheme And mHighContrastThemeOn Then
        BackColorTabs = mHandleHighContrastTheme_OrigBackColorTabs
    Else
        If mBackColorTabsSavingWhileVisualStyles Then
            BackColorTabs = mBackColorTabs_SavedWhileVisualStyles
        Else
            BackColorTabs = mBackColorTabs
        End If
    End If
End Property

Public Property Let BackColorTabs(ByVal nValue As OLE_COLOR)
    Dim iWv As Boolean
    Dim iPrev As Long
    
    If nValue <> IIf(mBackColorTabsSavingWhileVisualStyles, mBackColorTabs_SavedWhileVisualStyles, mBackColorTabs) Then
        If Not IsValidOLE_COLOR(nValue) Then RaiseError 380, TypeName(Me): Exit Property
        mBackColorTabsIsFromAmbient = (nValue = Ambient.BackColor)
        If mBackColorTabsSavingWhileVisualStyles Then
            mBackColorTabs_SavedWhileVisualStyles = nValue
        Else
            iPrev = mBackColorTabs
            If mAmbientUserMode And mHandleHighContrastTheme And mHighContrastThemeOn Then
                mHandleHighContrastTheme_OrigBackColorTabs = nValue
                If (mBackColorSelectedTab = iPrev) And (mBackColorSelectedTab <> nValue) Then
                    mSettingBackColorSelectedTabFromBackColorTabs = True
                    BackColorSelectedTab = nValue
                    mSettingBackColorSelectedTabFromBackColorTabs = False
                End If
            Else
                mBackColorTabs = nValue
                SetPropertyChanged "BackColorTabs"
                SetColors
                iWv = IsWindowVisible(mUserControlHwnd) <> 0
                If iWv Then SendMessage mUserControlHwnd, WM_SETREDRAW, False, 0&
                If ((mBackColorSelectedTab = iPrev) Or mBackColorSelectedTab_IsAutomatic) And (mBackColorSelectedTab <> nValue) And (mBackStyle = ntOpaque) Then
                    mSettingBackColorSelectedTabFromBackColorTabs = True
                    BackColorSelectedTab = nValue
                    mSettingBackColorSelectedTabFromBackColorTabs = False
                Else
                    Draw
                End If
                If iWv Then SendMessage mUserControlHwnd, WM_SETREDRAW, True, 0&
                If iWv Then RedrawWindow mUserControlHwnd, ByVal 0&, 0&, RDW_INVALIDATE Or RDW_ALLCHILDREN
            End If
        End If
    End If
End Property


Public Property Get BackColorSelectedTab() As OLE_COLOR
Attribute BackColorSelectedTab.VB_Description = "Returns/sets the color of the active tab including the tab body."
Attribute BackColorSelectedTab.VB_ProcData.VB_Invoke_Property = ";Colors"
    If mAmbientUserMode And mHandleHighContrastTheme And mHighContrastThemeOn Then
        BackColorSelectedTab = mHandleHighContrastTheme_OrigBackColorSelectedTab
    Else
        If mBackColorTabsSavingWhileVisualStyles Then
            BackColorSelectedTab = mBackColorSelectedTab_SavedWhileVisualStyles
        Else
            BackColorSelectedTab = mBackColorSelectedTab
        End If
    End If
End Property

Public Property Let BackColorSelectedTab(ByVal nValue As OLE_COLOR)
    Dim iPrev As Long
    Dim iWv As Boolean
    
    If nValue <> IIf(mBackColorTabsSavingWhileVisualStyles, mBackColorSelectedTab_SavedWhileVisualStyles, mBackColorSelectedTab) Then
        If Not IsValidOLE_COLOR(nValue) Then RaiseError 380, TypeName(Me): Exit Property
        If Not mChangingHighContrastTheme Then
            'mBackColorSelectedTab_IsAutomatic = (nValue = BackColorTabs) Or (nValue = GetAutomaticBackColorSelectedTab)
            mBackColorSelectedTab_IsAutomatic = (nValue = -1) Or (nValue = GetAutomaticBackColorSelectedTab) Or mSettingBackColorSelectedTabFromBackColorTabs
            If mBackColorSelectedTab_IsAutomatic Then nValue = GetAutomaticBackColorSelectedTab
        End If
        If mAmbientUserMode And mHandleHighContrastTheme And mHighContrastThemeOn And (Not mChangingHighContrastTheme) Then
            mHandleHighContrastTheme_OrigBackColorSelectedTab = nValue
        Else
            If mBackColorTabsSavingWhileVisualStyles And Not (mHighContrastThemeOn Or mChangingHighContrastTheme) Then
                mBackColorSelectedTab_SavedWhileVisualStyles = nValue
            Else
                If Enabled Or Not mShowDisabledState Then
                    iPrev = mBackColorSelectedTab
                Else
                    iPrev = mBackColorSelectedTabDisabled
                End If
                mBackColorSelectedTab = nValue
                SetPropertyChanged "BackColorSelectedTab"
                SetColors
                iWv = IsWindowVisible(mUserControlHwnd) <> 0
                If iWv Then SendMessage mUserControlHwnd, WM_SETREDRAW, False, 0&
                If mChangeControlsBackColor Then
                    SetControlsBackColor IIf((Not Enabled) And mShowDisabledState, mBackColorSelectedTabDisabled, mBackColorSelectedTab), iPrev
                    If mChangeControlsForeColor Then
                        SetControlsProperForeColor
                    End If
                End If
                mSubclassControlsPaintingPending = True
                mRepaintSubclassedControls = True
                mBodyReset = True
                SubclassControlsPainting
                Draw
                If iWv Then SendMessage mUserControlHwnd, WM_SETREDRAW, True, 0&
                If iWv Then RedrawWindow mUserControlHwnd, ByVal 0&, 0&, RDW_INVALIDATE Or RDW_ALLCHILDREN
            End If
        End If
    End If
End Property


Public Property Get FlatBarColorSelectedTab() As OLE_COLOR
Attribute FlatBarColorSelectedTab.VB_Description = "Returns/sets the color of the bar when a tab is selected in flat style."
Attribute FlatBarColorSelectedTab.VB_ProcData.VB_Invoke_Property = ";Colors"
    FlatBarColorSelectedTab = mFlatBarColorSelectedTab
End Property

Public Property Let FlatBarColorSelectedTab(ByVal nValue As OLE_COLOR)
    If nValue <> mFlatBarColorSelectedTab Then
        If Not IsValidOLE_COLOR(nValue) Then RaiseError 380, TypeName(Me): Exit Property
        mFlatBarColorSelectedTab = nValue
        SetPropertyChanged "FlatBarColorSelectedTab"
        SetColors
        DrawDelayed
    End If
End Property


Public Property Get FlatBarColorHighlight() As OLE_COLOR
Attribute FlatBarColorHighlight.VB_Description = "Returns/sets the color of the bar when a tab is highlighted in flat style."
Attribute FlatBarColorHighlight.VB_ProcData.VB_Invoke_Property = ";Colors"
    FlatBarColorHighlight = mFlatBarColorHighlight
End Property

Public Property Let FlatBarColorHighlight(ByVal nValue As OLE_COLOR)
    If nValue <> mFlatBarColorHighlight Then
        If nValue = -1 Then
            nValue = mFlatBarColorHighlight_ColorAutomatic
        End If
        If Not IsValidOLE_COLOR(nValue) Then RaiseError 380, TypeName(Me): Exit Property
        mFlatBarColorHighlight = nValue
        mFlatBarColorHighlight_IsAutomatic = (nValue = mFlatBarColorHighlight_ColorAutomatic)
        SetPropertyChanged "FlatBarColorHighlight"
        SetColors
        DrawDelayed
    End If
End Property


Public Property Get FlatBarColorInactive() As OLE_COLOR
Attribute FlatBarColorInactive.VB_Description = "Returns/sets the color of the bar when a tab is inactive in flat style."
Attribute FlatBarColorInactive.VB_ProcData.VB_Invoke_Property = ";Colors"
    FlatBarColorInactive = mFlatBarColorInactive
End Property

Public Property Let FlatBarColorInactive(ByVal nValue As OLE_COLOR)
    If nValue <> mFlatBarColorInactive Then
        If nValue = -1 Then
            nValue = mFlatBarColorInactive_ColorAutomatic
        End If
        If Not IsValidOLE_COLOR(nValue) Then RaiseError 380, TypeName(Me): Exit Property
        mFlatBarColorInactive = nValue
        mFlatBarColorInactive_IsAutomatic = (nValue = mFlatBarColorInactive_ColorAutomatic)
        SetPropertyChanged "FlatBarColorInactive"
        SetColors
        DrawDelayed
    End If
End Property


Public Property Get FlatTabsSeparationLineColor() As OLE_COLOR
Attribute FlatTabsSeparationLineColor.VB_Description = "Returns/sets the color of the separation line between tabs in flat style."
Attribute FlatTabsSeparationLineColor.VB_ProcData.VB_Invoke_Property = ";Colors"
    FlatTabsSeparationLineColor = mFlatTabsSeparationLineColor
End Property

Public Property Let FlatTabsSeparationLineColor(ByVal nValue As OLE_COLOR)
    If nValue <> mFlatTabsSeparationLineColor Then
        If nValue = -1 Then
            nValue = mFlatTabsSeparationLineColor_ColorAutomatic
        End If
        If Not IsValidOLE_COLOR(nValue) Then RaiseError 380, TypeName(Me): Exit Property
        mFlatTabsSeparationLineColor = nValue
        mFlatTabsSeparationLineColor_IsAutomatic = (nValue = mFlatTabsSeparationLineColor_ColorAutomatic)
        SetPropertyChanged "FlatTabsSeparationLineColor"
        SetColors
        DrawDelayed
    End If
End Property


Public Property Get FlatBodySeparationLineColor() As OLE_COLOR
Attribute FlatBodySeparationLineColor.VB_Description = "Returns/sets the color of the separation line between the tabs and the body in flat style."
Attribute FlatBodySeparationLineColor.VB_ProcData.VB_Invoke_Property = ";Colors"
    FlatBodySeparationLineColor = mFlatBodySeparationLineColor
End Property

Public Property Let FlatBodySeparationLineColor(ByVal nValue As OLE_COLOR)
    If nValue <> mFlatBodySeparationLineColor Then
        If nValue = -1 Then
            nValue = mFlatBodySeparationLineColor_ColorAutomatic
        End If
        If Not IsValidOLE_COLOR(nValue) Then RaiseError 380, TypeName(Me): Exit Property
        mFlatBodySeparationLineColor = nValue
        mFlatBodySeparationLineColor_IsAutomatic = (nValue = mFlatBodySeparationLineColor_ColorAutomatic)
        SetPropertyChanged "FlatBodySeparationLineColor"
        SetColors
        DrawDelayed
    End If
End Property


Public Property Get FlatBorderColor() As OLE_COLOR
Attribute FlatBorderColor.VB_Description = "Returns/sets the color of the border in flat style."
Attribute FlatBorderColor.VB_ProcData.VB_Invoke_Property = ";Colors"
    FlatBorderColor = mFlatBorderColor
End Property

Public Property Let FlatBorderColor(ByVal nValue As OLE_COLOR)
    If nValue <> mFlatBorderColor Then
        If nValue = -1 Then
            nValue = mFlatBorderColor_ColorAutomatic
        End If
        If Not IsValidOLE_COLOR(nValue) Then RaiseError 380, TypeName(Me): Exit Property
        mFlatBorderColor = nValue
        mFlatBorderColor_IsAutomatic = (nValue = mFlatBorderColor_ColorAutomatic)
        SetPropertyChanged "FlatBorderColor"
        SetColors
        DrawDelayed
    End If
End Property


Public Property Get HighlightColor() As OLE_COLOR
Attribute HighlightColor.VB_Description = "Returns/sets the color used to highlight an inactive tab when the user hovers over it."
Attribute HighlightColor.VB_ProcData.VB_Invoke_Property = ";Colors"
    HighlightColor = mHighlightColor
End Property

Public Property Let HighlightColor(ByVal nValue As OLE_COLOR)
    If nValue <> mHighlightColor Then
        If nValue = -1 Then
            nValue = mHighlightColor_ColorAutomatic
        End If
        If Not IsValidOLE_COLOR(nValue) Then RaiseError 380, TypeName(Me): Exit Property
        mHighlightColor = nValue
        mHighlightColor_IsAutomatic = (nValue = mHighlightColor_ColorAutomatic)
        SetPropertyChanged "HighlightColor"
        SetColors
        DrawDelayed
    End If
End Property


Public Property Get HighlightColorSelectedTab() As OLE_COLOR
Attribute HighlightColorSelectedTab.VB_Description = "Returns/sets the color used to highlight the selected tab."
Attribute HighlightColorSelectedTab.VB_ProcData.VB_Invoke_Property = ";Colors"
    HighlightColorSelectedTab = mHighlightColorSelectedTab
End Property

Public Property Let HighlightColorSelectedTab(ByVal nValue As OLE_COLOR)
    If nValue <> mHighlightColorSelectedTab Then
        If nValue = -1 Then
            nValue = mHighlightColorSelectedTab_ColorAutomatic
        End If
        If Not IsValidOLE_COLOR(nValue) Then RaiseError 380, TypeName(Me): Exit Property
        mHighlightColorSelectedTab = nValue
        mHighlightColorSelectedTab_IsAutomatic = (nValue = mHighlightColorSelectedTab_ColorAutomatic)
        SetPropertyChanged "HighlightColorSelectedTab"
        SetColors
        DrawDelayed
    End If
End Property


Public Property Get IconColor() As OLE_COLOR
Attribute IconColor.VB_Description = "Returns/sets the color of the icon."
Attribute IconColor.VB_ProcData.VB_Invoke_Property = ";Colors"
    If mAmbientUserMode And mHandleHighContrastTheme And mHighContrastThemeOn Then
        IconColor = mHandleHighContrastTheme_OrigIconColor
    Else
        IconColor = mIconColor
    End If
End Property

Public Property Let IconColor(ByVal nValue As OLE_COLOR)
    Dim iPrev As Long
    
    If nValue <> mIconColor Then
        If Not IsValidOLE_COLOR(nValue) Then RaiseError 380, TypeName(Me): Exit Property
        mIconColorIsFromAmbient = (nValue = Ambient.ForeColor)
        If mAmbientUserMode And mHandleHighContrastTheme And mHighContrastThemeOn Then
            mHandleHighContrastTheme_OrigIconColor = nValue
        Else
            iPrev = mIconColor
            mIconColor = nValue
            SetPropertyChanged "IconColor"
            If IconColorSelectedTab = iPrev Then
                IconColorSelectedTab = nValue
            End If
            If IconColorMouseHover = iPrev Then
                IconColorMouseHover = nValue
            End If
            If IconColorMouseHoverSelectedTab = iPrev Then
                IconColorMouseHoverSelectedTab = nValue
            End If
            If IconColorTabHighlighted = iPrev Then
                IconColorTabHighlighted = nValue
            End If
            If IconColor = iPrev Then
                IconColor = nValue
            End If
            DrawDelayed
        End If
    End If
End Property


Public Property Get IconColorSelectedTab() As OLE_COLOR
Attribute IconColorSelectedTab.VB_Description = "Returns/sets the color used to draw the icon when the tab is seleted."
Attribute IconColorSelectedTab.VB_ProcData.VB_Invoke_Property = ";Colors"
    If mAmbientUserMode And mHandleHighContrastTheme And mHighContrastThemeOn Then
        IconColorSelectedTab = mHandleHighContrastTheme_OrigIconColorSelectedTab
    Else
        IconColorSelectedTab = mIconColorSelectedTab
    End If
End Property

Public Property Let IconColorSelectedTab(ByVal nValue As OLE_COLOR)
    If nValue <> mIconColorSelectedTab Then
        If Not IsValidOLE_COLOR(nValue) Then RaiseError 380, TypeName(Me): Exit Property
        If mAmbientUserMode And mHandleHighContrastTheme And mHighContrastThemeOn Then
            mHandleHighContrastTheme_OrigIconColorSelectedTab = nValue
        Else
            mIconColorSelectedTab = nValue
            SetPropertyChanged "IconColorSelectedTab"
            DrawDelayed
        End If
    End If
End Property


Public Property Get IconColorMouseHover() As OLE_COLOR
Attribute IconColorMouseHover.VB_Description = "Returns/sets the color that the icon will show when the mouse hovers it, for the non active tabs."
Attribute IconColorMouseHover.VB_ProcData.VB_Invoke_Property = ";Colors"
    If mAmbientUserMode And mHandleHighContrastTheme And mHighContrastThemeOn Then
        IconColorMouseHover = mHandleHighContrastTheme_OrigIconColorMouseHover
    Else
        IconColorMouseHover = mIconColorMouseHover
    End If
End Property

Public Property Let IconColorMouseHover(ByVal nValue As OLE_COLOR)
    If nValue <> mIconColorMouseHover Then
        If Not IsValidOLE_COLOR(nValue) Then RaiseError 380, TypeName(Me): Exit Property
        If mAmbientUserMode And mHandleHighContrastTheme And mHighContrastThemeOn Then
            mHandleHighContrastTheme_OrigIconColorMouseHover = nValue
        Else
            mIconColorMouseHover = nValue
            If mTDIMode <> ntTDIModeNone Then
                mTDIIconColorMouseHover = mIconColorMouseHover
            End If
            SetPropertyChanged "IconColorMouseHover"
            DrawDelayed
        End If
    End If
End Property


Public Property Get IconColorMouseHoverSelectedTab() As OLE_COLOR
Attribute IconColorMouseHoverSelectedTab.VB_Description = "Returns/sets the color that the icon will show when the mouse hovers it, for the active tab."
Attribute IconColorMouseHoverSelectedTab.VB_ProcData.VB_Invoke_Property = ";Colors"
    If mAmbientUserMode And mHandleHighContrastTheme And mHighContrastThemeOn Then
        IconColorMouseHoverSelectedTab = mHandleHighContrastTheme_OrigIconColorMouseHoverSelectedTab
    Else
        IconColorMouseHoverSelectedTab = mIconColorMouseHoverSelectedTab
    End If
End Property

Public Property Let IconColorMouseHoverSelectedTab(ByVal nValue As OLE_COLOR)
    If nValue <> mIconColorMouseHoverSelectedTab Then
        If Not IsValidOLE_COLOR(nValue) Then RaiseError 380, TypeName(Me): Exit Property
        If mAmbientUserMode And mHandleHighContrastTheme And mHighContrastThemeOn Then
            mHandleHighContrastTheme_OrigIconColorMouseHoverSelectedTab = nValue
        Else
            mIconColorMouseHoverSelectedTab = nValue
            SetPropertyChanged "IconColorMouseHoverSelectedTab"
            DrawDelayed
        End If
    End If
End Property


Public Property Get IconColorTabHighlighted() As OLE_COLOR
Attribute IconColorTabHighlighted.VB_Description = "Returns/sets the color used to draw the icon when the tab is highlighted (not the icon itself)."
Attribute IconColorTabHighlighted.VB_ProcData.VB_Invoke_Property = ";Colors"
    If mAmbientUserMode And mHandleHighContrastTheme And mHighContrastThemeOn Then
        IconColorTabHighlighted = mHandleHighContrastTheme_OrigIconColorTabHighlighted
    Else
        IconColorTabHighlighted = mIconColorTabHighlighted
    End If
End Property

Public Property Let IconColorTabHighlighted(ByVal nValue As OLE_COLOR)
    If nValue <> mIconColorTabHighlighted Then
        If Not IsValidOLE_COLOR(nValue) Then RaiseError 380, TypeName(Me): Exit Property
        If mAmbientUserMode And mHandleHighContrastTheme And mHighContrastThemeOn Then
            mHandleHighContrastTheme_OrigIconColorTabHighlighted = nValue
        Else
            mIconColorTabHighlighted = nValue
            SetPropertyChanged "IconColorTabHighlighted"
        End If
    End If
End Property


Public Property Get ShowDisabledState() As Boolean
Attribute ShowDisabledState.VB_Description = "Returns/sets a value that determines if the tabs color will be darkened when the control is disabled."
Attribute ShowDisabledState.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute ShowDisabledState.VB_MemberFlags = "400"
    ShowDisabledState = mShowDisabledState
End Property

Public Property Let ShowDisabledState(ByVal nValue As Boolean)
    If nValue <> mShowDisabledState Then
        mShowDisabledState = nValue
        SetPropertyChanged "ShowDisabledState"
        mBodyReset = True
        DrawDelayed
        If mChangeControlsBackColor Then
            If mEnabled Or Not mShowDisabledState Then
                SetControlsBackColor mBackColorSelectedTab, mBackColorSelectedTabDisabled
            Else
                SetControlsBackColor mBackColorSelectedTabDisabled, mBackColorSelectedTab
            End If
        End If
    End If
End Property


Public Property Get Redraw() As Boolean
Attribute Redraw.VB_Description = "Returns/sets a value that determines if the drawing of the control is enabled."
Attribute Redraw.VB_MemberFlags = "400"
    Redraw = mRedraw
End Property

Public Property Let Redraw(ByVal nValue As Boolean)
    If nValue <> mRedraw Then
        mRedraw = nValue
        If mRedraw Then
            If mNeedToDraw Or mDrawMessagePosted Or tmrDraw.Enabled Then
                Draw
            End If
        End If
    End If
End Property


Public Property Get UseMaskColor() As Boolean
Attribute UseMaskColor.VB_Description = "Returns/sets a value that determines whether the color assigned in the MaskColor property is used as a mask. (That is, used to create transparent regions.)"
Attribute UseMaskColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    UseMaskColor = mUseMaskColor
End Property

Public Property Let UseMaskColor(ByVal nValue As Boolean)
    If nValue <> mUseMaskColor Then
        mUseMaskColor = nValue
        SetPropertyChanged "UseMaskColor"
        DrawDelayed
    End If
End Property


Public Property Get TabTransition() As NTTabTransitionConstants
Attribute TabTransition.VB_Description = "Returns/sets a value that determines whether the transition between tabs will be with an effect that smooths the transition (and its speed)."
Attribute TabTransition.VB_ProcData.VB_Invoke_Property = ";Behavior"
    TabTransition = mTabTransition
End Property

Public Property Let TabTransition(ByVal nValue As NTTabTransitionConstants)
    If nValue <> mTabTransition Then
        If (nValue < ntTransitionImmediate) Or (nValue > ntTransitionSlower) Then
            RaiseError 380, TypeName(Me) ' invalid property value
            Exit Property
        End If
        mTabTransition = nValue
        SetPropertyChanged "TabTransition"
    End If
End Property


Public Property Get FlatRoundnessTop() As Long
Attribute FlatRoundnessTop.VB_Description = "Returns/sets the size in pixels of the roundness of the top corners."
Attribute FlatRoundnessTop.VB_ProcData.VB_Invoke_Property = ";Appearance"
    FlatRoundnessTop = mFlatRoundnessTop
End Property

Public Property Let FlatRoundnessTop(ByVal nValue As Long)
    If nValue <> mFlatRoundnessTop Then
        If (nValue < 0) Or (nValue > 100) Then
            RaiseError 380, TypeName(Me) ' invalid property value
            Exit Property
        End If
        mFlatRoundnessTop = nValue
        mFlatRoundnessTopDPIScaled = mFlatRoundnessTop * mDPIScale
        SetPropertyChanged "FlatRoundnessTop"
        DrawDelayed
    End If
End Property


Public Property Get FlatRoundnessBottom() As Long
Attribute FlatRoundnessBottom.VB_Description = "Returns/sets the size in pixels of the roundness of the bottom corners."
Attribute FlatRoundnessBottom.VB_ProcData.VB_Invoke_Property = ";Appearance"
    FlatRoundnessBottom = mFlatRoundnessBottom
End Property

Public Property Let FlatRoundnessBottom(ByVal nValue As Long)
    If nValue <> mFlatRoundnessBottom Then
        If (nValue < 0) Or (nValue > 100) Then
            RaiseError 380, TypeName(Me) ' invalid property value
            Exit Property
        End If
        mFlatRoundnessBottom = nValue
        mFlatRoundnessBottomDPIScaled = mFlatRoundnessBottom * mDPIScale
        SetPropertyChanged "FlatRoundnessBottom"
        DrawDelayed
    End If
End Property


Public Property Get FlatRoundnessTabs() As Long
Attribute FlatRoundnessTabs.VB_Description = "Returns/sets the size in pixels of the roundness of the tabs corners."
Attribute FlatRoundnessTabs.VB_ProcData.VB_Invoke_Property = ";Appearance"
    FlatRoundnessTabs = mFlatRoundnessTabs
End Property

Public Property Let FlatRoundnessTabs(ByVal nValue As Long)
    If nValue <> mFlatRoundnessTabs Then
        If (nValue < 0) Or (nValue > 100) Then
            RaiseError 380, TypeName(Me) ' invalid property value
            Exit Property
        End If
        mFlatRoundnessTabs = nValue
        mFlatRoundnessTabsDPIScaled = mFlatRoundnessTabs * mDPIScale
        SetPropertyChanged "FlatRoundnessTabs"
        DrawDelayed
    End If
End Property


Public Property Get HighlightMode() As Long
Attribute HighlightMode.VB_Description = "Returns/sets the mode that the inactive tabs are highlighted when the mouse hovers over them."
Attribute HighlightMode.VB_ProcData.VB_Invoke_Property = "pagHighlightMode;Appearance"
    HighlightMode = mHighlightMode
End Property

Public Property Let HighlightMode(ByVal nValue As Long)
    If nValue <> mHighlightMode Then
'        If (nValue < ntHLAuto) Or (nValue > ntHLAllFlags) Then
'            RaiseError 380, TypeName(Me) ' invalid property value
'            Exit Property
'        End If
        mHighlightMode = nValue
        mSetAutoTabHeightPending = True
        SetPropertyChanged "HighlightMode"
        SetHighlightMode
        DrawDelayed
    End If
End Property


Public Property Get HighlightModeSelectedTab() As Long
Attribute HighlightModeSelectedTab.VB_Description = "Returns/sets the mode that the selected tab is highlighted (it is always highlighted)."
Attribute HighlightModeSelectedTab.VB_ProcData.VB_Invoke_Property = "pagHighlightMode;Appearance"
    HighlightModeSelectedTab = mHighlightModeSelectedTab
End Property

Public Property Let HighlightModeSelectedTab(ByVal nValue As Long)
    If nValue <> mHighlightModeSelectedTab Then
'        If (nValue < ntHLAuto) Or (nValue > ntHLAllFlags) Then
'            RaiseError 380, TypeName(Me) ' invalid property value
'            Exit Property
'        End If
        mHighlightModeSelectedTab = nValue
        mSetAutoTabHeightPending = True
        SetPropertyChanged "HighlightModeSelectedTab"
        SetHighlightMode
        DrawDelayed
    End If
End Property


Public Property Get FlatBorderMode() As NTFlatBorderModeConstants
Attribute FlatBorderMode.VB_Description = "Returns/sets the way the border will be drawn in flat style. It may be around the selected tab or all the control."
Attribute FlatBorderMode.VB_ProcData.VB_Invoke_Property = ";Appearance"
    FlatBorderMode = mFlatBorderMode
End Property

Public Property Let FlatBorderMode(ByVal nValue As NTFlatBorderModeConstants)
    If nValue <> mFlatBorderMode Then
        If (nValue < ntBorderTabs) Or (nValue > ntBorderSelectedTab) Then
            RaiseError 380, TypeName(Me) ' invalid property value
            Exit Property
        End If
        mFlatBorderMode = nValue
        SetPropertyChanged "FlatBorderMode"
        SetHighlightMode
        DrawDelayed
    End If
End Property


Public Property Get FlatBarHeight() As Long
Attribute FlatBarHeight.VB_Description = "Returns/sets the height in pixels of the bar in flat style."
Attribute FlatBarHeight.VB_ProcData.VB_Invoke_Property = ";Appearance"
    FlatBarHeight = mFlatBarHeight
End Property

Public Property Let FlatBarHeight(ByVal nValue As Long)
    If nValue <> mFlatBarHeight Then
        If (nValue < 0) Or (nValue > 15) Then
            RaiseError 380, TypeName(Me) ' invalid property value
            Exit Property
        End If
        mFlatBarHeight = nValue
        mFlatBarHeightDPIScaled = mFlatBarHeight * mDPIScale
        mSetAutoTabHeightPending = True
        SetPropertyChanged "FlatBarHeight"
        DrawDelayed
    End If
End Property


Public Property Get FlatBarGripHeight() As Long
Attribute FlatBarGripHeight.VB_Description = "Returns/sets a value in pixels that determines the height of a grip for the highlight bar in the flat style. A negative value defines a notch instead."
Attribute FlatBarGripHeight.VB_ProcData.VB_Invoke_Property = ";Appearance"
    FlatBarGripHeight = mFlatBarGripHeight
End Property

Public Property Let FlatBarGripHeight(ByVal nValue As Long)
    If nValue <> mFlatBarGripHeight Then
        If (nValue < -50) Or (nValue > 50) Then
            RaiseError 380, TypeName(Me) ' invalid property value
            Exit Property
        End If
        mFlatBarGripHeight = nValue
        mFlatBarGripHeightDPIScaled = mFlatBarGripHeight * mDPIScale
        mSetAutoTabHeightPending = True
        SetPropertyChanged "FlatBarGripHeight"
        DrawDelayed
    End If
End Property


Public Property Get FlatBodySeparationLineHeight() As Long
Attribute FlatBodySeparationLineHeight.VB_Description = "Returns/sets the height in pixels of the separation line between the tabs and the body in flat style."
Attribute FlatBodySeparationLineHeight.VB_ProcData.VB_Invoke_Property = ";Appearance"
    FlatBodySeparationLineHeight = mFlatBodySeparationLineHeight
End Property

Public Property Let FlatBodySeparationLineHeight(ByVal nValue As Long)
    If nValue <> mFlatBodySeparationLineHeight Then
        If (nValue < 0) Or (nValue > 50) Then
            RaiseError 380, TypeName(Me) ' invalid property value
            Exit Property
        End If
        mFlatBodySeparationLineHeight = nValue
        mFlatBodySeparationLineHeightDPIScaled = mFlatBodySeparationLineHeight * mDPIScale
        SetPropertyChanged "FlatBodySeparationLineHeight"
        DrawDelayed
    End If
End Property


Public Property Get FlatBarPosition() As NTFlatBarPositionConstants
Attribute FlatBarPosition.VB_Description = "Returns/sets the position of the bar in flat style."
Attribute FlatBarPosition.VB_ProcData.VB_Invoke_Property = ";Appearance"
    FlatBarPosition = mFlatBarPosition
End Property

Public Property Let FlatBarPosition(ByVal nValue As NTFlatBarPositionConstants)
    If nValue <> mFlatBarPosition Then
        If (nValue < ntBarPositionTop) Or (nValue > ntBarPositionBottom) Then
            RaiseError 380, TypeName(Me) ' invalid property value
            Exit Property
        End If
        mFlatBarPosition = nValue
        mSetAutoTabHeightPending = True
        SetPropertyChanged "FlatBarPosition"
        DrawDelayed
    End If
End Property


Public Property Get ShowRowsInPerspective() As NTAutoYesNoConstants
Attribute ShowRowsInPerspective.VB_Description = "Returns/sets a value that determines when the control has more that one row of tabs, if they will be drawn changing the horizontal position on each row."
Attribute ShowRowsInPerspective.VB_ProcData.VB_Invoke_Property = ";Tabs"
    ShowRowsInPerspective = mShowRowsInPerspective
End Property

Public Property Let ShowRowsInPerspective(ByVal nValue As NTAutoYesNoConstants)
    Dim iValue As NTAutoYesNoConstants
    
    iValue = nValue
    If (iValue <> ntNo) And (iValue <> ntYNAuto) Then
        iValue = ntYes
    End If
    If iValue <> mShowRowsInPerspective Then
        mShowRowsInPerspective = iValue
        SetPropertyChanged "ShowRowsInPerspective"
        ResetCachedThemeImages
        DrawDelayed
    End If
End Property


Public Property Get TabSeparation() As Long
Attribute TabSeparation.VB_Description = "Returns/sets the number of pixels of separation between tabs."
Attribute TabSeparation.VB_ProcData.VB_Invoke_Property = ";Appearance"
    TabSeparation = mTabSeparation
End Property

Public Property Let TabSeparation(ByVal nValue As Long)
    If nValue < 0 Or nValue > 20 Then
        RaiseError 380, TypeName(Me) ' invalid property value
        Exit Property
    End If
    If nValue <> mTabSeparation Then
        mTabSeparation = nValue
        mTabSeparationDPIScaled = mTabSeparation * mDPIScale
        SetPropertyChanged "TabSeparation"
        ResetCachedThemeImages
        DrawDelayed
    End If
End Property

Public Property Get ChangeControlsBackColor() As Boolean
Attribute ChangeControlsBackColor.VB_Description = "Returns/sets a value that determines if the BackColor property value of the contained controls will be changed according to BackColorSelectedTab property value (applied only when they are the same)."
Attribute ChangeControlsBackColor.VB_ProcData.VB_Invoke_Property = ";Behavior"
    ChangeControlsBackColor = mChangeControlsBackColor
End Property

Public Property Let ChangeControlsBackColor(ByVal nValue As Boolean)
    Dim iWv As Boolean
    
    If nValue <> mChangeControlsBackColor Then
        mChangeControlsBackColor = nValue
        SetPropertyChanged "ChangeControlsBackColor"
        iWv = IsWindowVisible(mUserControlHwnd) <> 0
        If iWv Then SendMessage mUserControlHwnd, WM_SETREDRAW, False, 0&
        If Not mChangeControlsBackColor Then
            SetControlsBackColor vbButtonFace, IIf(mEnabled Or Not mShowDisabledState, mBackColorSelectedTab, mBackColorSelectedTabDisabled)
        Else
            SetControlsBackColor IIf(mEnabled Or Not mShowDisabledState, mBackColorSelectedTab, mBackColorSelectedTabDisabled)
        End If
        mSubclassControlsPaintingPending = True
        mRepaintSubclassedControls = True
        SubclassControlsPainting
        Draw
        If iWv Then SendMessage mUserControlHwnd, WM_SETREDRAW, True, 0&
        If iWv Then RedrawWindow mUserControlHwnd, ByVal 0&, 0&, RDW_INVALIDATE Or RDW_ALLCHILDREN
    End If
End Property


Public Property Get ChangeControlsForeColor() As Boolean
Attribute ChangeControlsForeColor.VB_Description = "Returns/sets a value that determines if the ForeColor property value of the contained controls will be changed according to ForeColor property value of the tab control (applied only when they are the same)."
Attribute ChangeControlsForeColor.VB_ProcData.VB_Invoke_Property = ";Behavior"
    ChangeControlsForeColor = mChangeControlsForeColor
End Property

Public Property Let ChangeControlsForeColor(ByVal nValue As Boolean)
    Dim iWv As Boolean
    
    If nValue <> mChangeControlsForeColor Then
        mChangeControlsForeColor = nValue
        SetPropertyChanged "ChangeControlsForeColor"
        iWv = IsWindowVisible(mUserControlHwnd) <> 0
        If iWv Then SendMessage mUserControlHwnd, WM_SETREDRAW, False, 0&
        If Not mChangeControlsForeColor Then
            SetControlsForeColor vbButtonText, IIf(mControlsForeColor_PrevColor = -1, mForeColorSelectedTab, mControlsForeColor_PrevColor)
        Else
            SetControlsProperForeColor
        End If
        mSubclassControlsPaintingPending = True
        mRepaintSubclassedControls = True
        SubclassControlsPainting
        Draw
        If iWv Then SendMessage mUserControlHwnd, WM_SETREDRAW, True, 0&
        If iWv Then RedrawWindow mUserControlHwnd, ByVal 0&, 0&, RDW_INVALIDATE Or RDW_ALLCHILDREN
    End If
End Property


Public Property Get AutoRelocateControls() As NTAutoRelocateControlsConstants
Attribute AutoRelocateControls.VB_Description = "Returns/sets a value that determines if the contained controls will be automatically relocated when the tab body changes."
Attribute AutoRelocateControls.VB_ProcData.VB_Invoke_Property = ";Behavior"
    AutoRelocateControls = mAutoRelocateControls
End Property

Public Property Let AutoRelocateControls(ByVal nValue As NTAutoRelocateControlsConstants)
    If (nValue < 0) Or (nValue > 2) Then
        RaiseError 380, TypeName(Me) ' invalid property value
        Exit Property
    End If
    If nValue <> mAutoRelocateControls Then
        mAutoRelocateControls = nValue
        SetPropertyChanged "AutoRelocateControls"
    End If
End Property


Public Property Get SoftEdges() As Boolean
Attribute SoftEdges.VB_Description = "Returns/sets a value that determines if the edges will be displayed with less contrast."
Attribute SoftEdges.VB_ProcData.VB_Invoke_Property = ";Appearance"
    SoftEdges = mSoftEdges
End Property

Public Property Let SoftEdges(ByVal nValue As Boolean)
    If nValue <> mSoftEdges Then
        mSoftEdges = nValue
        SetPropertyChanged "SoftEdges"
        SetColors
        DrawDelayed
    End If
End Property


Public Property Get RightToLeft() As Boolean
Attribute RightToLeft.VB_Description = "Returns/Sets the text display direction and control visual appearance on a bidirectional system."
Attribute RightToLeft.VB_ProcData.VB_Invoke_Property = ";Behavior"
    RightToLeft = mRightToLeft
End Property

Public Property Let RightToLeft(ByVal nValue As Boolean)
    If nValue Then
        If mTabOrientation <> ssTabOrientationTop Then
            RaiseError 1384, TypeName(Me), "RightToLeft is only available for Top orientation."
            Exit Property
        End If
    End If
    
    If nValue <> mRightToLeft Then
        mRightToLeft = nValue
        If mRightToLeft Then
            SetLayout GetDC(picDraw.hWnd), LAYOUT_RTL
            SetLayout GetDC(picAux.hWnd), LAYOUT_RTL
        Else
            SetLayout GetDC(picDraw.hWnd), 0
            SetLayout GetDC(picAux.hWnd), 0
        End If
        SetScaleMode
        SetPropertyChanged "RightToLeft"
        DrawDelayed
    End If
End Property


Public Property Get BackStyle() As NTBackStyleConstants
Attribute BackStyle.VB_Description = "Returns/sets the background style, opaque or transparent."
Attribute BackStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BackStyle = mBackStyle
End Property

Public Property Let BackStyle(ByVal nValue As NTBackStyleConstants)
    If nValue < ntTransparent Or nValue > ntOpaqueSelectedTab Then
        RaiseError 380, TypeName(Me) ' invalid property value
        Exit Property
    End If
    If nValue <> mBackStyle Then
        mBackStyle = nValue
        SetPropertyChanged "BackStyle"
        'ResetCachedThemeImages
        Draw
    End If
End Property


Public Property Get AutoTabHeight() As Boolean
Attribute AutoTabHeight.VB_Description = "Returns/sets a value that determines if the tab height is set automatically according to the font (and pictures)."
Attribute AutoTabHeight.VB_ProcData.VB_Invoke_Property = ";Tabs"
    AutoTabHeight = mAutoTabHeight
End Property

Public Property Let AutoTabHeight(ByVal nValue As Boolean)
    If nValue <> mAutoTabHeight Then
        mAutoTabHeight = nValue
        SetPropertyChanged "AutoTabHeight"
        mSetAutoTabHeightPending = True
        DrawDelayed
    End If
End Property


Private Function IBSSubclass_MsgResponse(ByVal hWnd As Long, ByVal iMsg As Long) As Long
    Select Case iMsg
        Case WM_PAINT, WM_PRINTCLIENT, WM_MOUSELEAVE
            IBSSubclass_MsgResponse = emrConsume
        Case WM_LBUTTONDOWN, WM_LBUTTONUP, WM_MOUSEACTIVATE, WM_SETFOCUS, WM_LBUTTONDBLCLK, WM_MOVE, WM_WINDOWPOSCHANGING, WM_SETCURSOR, WM_MOUSEMOVE, WM_ACTIVATE, WM_STYLECHANGING
            IBSSubclass_MsgResponse = emrPreprocess
        Case Else
            IBSSubclass_MsgResponse = emrPostProcess
    End Select
End Function

Private Sub IBSSubclass_UnsubclassIt()
    If mSubclassed Then
        ' The IDE protection was fired
        DoTerminate
           
        'If (Not mAmbientUserMode) Then
            ' The following emulates the zombie state (UserControl hatched/disabled), in case it didn't actually happened by VB.
            ' Because the control anyway will be unclickable on the IDE any more without the subclassing.
            ' The developer needs to close the form and open it again to restore the functionality.
            UserControl.FillStyle = 5
'            UserControl.DrawWidth = 30
'            UserControl.FillColor = vbRed
            UserControl.Line (0, 0)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1), , B
            UserControl.FillStyle = 1
            UserControl.Enabled = False
        'End If
    End If
End Sub

Private Function IBSSubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, wParam As Long, lParam As Long, bConsume As Boolean) As Long
    Dim iTab As Long
    Const WA_INACTIVE As Long = 0
    Const WA_ACTIVE As Long = 1
    Dim iHwnd As Long
    
    Select Case iMsg
        Case WM_WINDOWPOSCHANGING ' invisible controls, to prevent being moved to the visible space if they are moved by code. Unfortunately the same can't be done to Labels and other windowless controls. But at least the protection acts on windowed controls.
            Dim iwp As WINDOWPOS
            
            CopyMemory iwp, ByVal lParam, Len(iwp)
            If iwp.X > -mLeftThresholdHided \ Screen.TwipsPerPixelX Then
                iwp.X = iwp.X - mLeftOffsetToHide \ Screen.TwipsPerPixelX
                CopyMemory ByVal lParam, iwp, Len(iwp)
            End If
            
        Case WM_NCACTIVATE ' need to update the focus rect
            mFormIsActive = (wParam <> 0)
            If mHasFocus Then
                PostDrawMessage
            End If
        Case WM_PRINTCLIENT, WM_MOUSELEAVE ' fixes frames paint bug in XP
            IBSSubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
        Case WM_SYSCOLORCHANGE, WM_THEMECHANGED ' they are form's messages
            SetButtonFaceColor
            SetColors
            mThemeExtraDataAlreadySet = False
            SetThemeExtraData
            ResetCachedThemeImages
            If mHandleHighContrastTheme Then CheckHighContrastTheme
            Draw
        Case WM_SETFOCUS
            If mNoActivate Then
                bConsume = True
                IBSSubclass_WindowProc = 0
                SetFocusAPI wParam
                mNoActivate = False
            End If
        Case WM_DRAW
            Draw
        Case WM_INIT
            If Not mTabStopsInitialized Then
                StoreControlsTabStop True
                mTabStopsInitialized = True
            End If
        Case WM_MOUSEACTIVATE ' UserControl message, only at run time (Ambient.UserMode),, to avoid taking the focus when the tab control is clicked in a non-clickable part (outside a tab).
            If mTabUnderMouse = -1 Then
                Dim iPt2 As POINTAPI
                
                GetCursorPos iPt2
                iHwnd = WindowFromPoint(iPt2.X, iPt2.Y)
                If iHwnd = mUserControlHwnd Then
                    mNoActivate = True
                End If
            End If
        Case WM_LBUTTONDBLCLK
            If Not MouseIsOverAContainedControl Then
                iTab = mTab
                Call ProcessMouseMove(vbLeftButton, 0, (lParam And &HFFFF&) * Screen_TwipsPerPixelX, (lParam \ &H10000 And &HFFFF&) * Screen_TwipsPerPixelX)
                Call UserControl_MouseDown(vbLeftButton, 0, (lParam And &HFFFF&) * Screen_TwipsPerPixelX, (lParam \ &H10000 And &HFFFF&) * Screen_TwipsPerPixelX)
                If mTab <> iTab Then
                    bConsume = True
                    IBSSubclass_WindowProc = 0
                    tmrCancelDoubleClick.Enabled = True
                End If
            End If
            If tmrCancelDoubleClick.Enabled Then
                bConsume = True
                mouse_event MOUSEEVENTF_LEFTDOWN, 0&, 0&, 0&, GetMessageExtraInfo()
                mouse_event MOUSEEVENTF_LEFTUP, 0&, 0&, 0&, GetMessageExtraInfo()
            End If
        Case WM_MOUSEMOVE
            Call ProcessMouseMove(vbLeftButton, 0, (lParam And &HFFFF&) * Screen_TwipsPerPixelX, (lParam \ &H10000 And &HFFFF&) * Screen_TwipsPerPixelX)
        Case WM_LBUTTONDOWN ' UserControl message, only in design mode (Not Ambient.UserMode), to provide change of selected tab by clicking at design time
            If Not MouseIsOverAContainedControl Then
                iTab = mTab
                Call ProcessMouseMove(vbLeftButton, 0, (lParam And &HFFFF&) * Screen_TwipsPerPixelX, (lParam \ &H10000 And &HFFFF&) * Screen_TwipsPerPixelX)
                Call UserControl_MouseDown(vbLeftButton, 0, (lParam And &HFFFF&) * Screen_TwipsPerPixelX, (lParam \ &H10000 And &HFFFF&) * Screen_TwipsPerPixelX)
                If mTab <> iTab Then
                    bConsume = True
                    IBSSubclass_WindowProc = 0
                    mBtnDown = True
                    'tmrCancelDoubleClick.Enabled = True
                End If
            End If
            If mChangeControlsBackColor And ((mBackColorTabs <> vbButtonFace) Or mControlIsThemed) Then
                mLastContainedControlsCount = UserControlContainedControlsCount
                If Not mAmbientUserMode Then tmrCheckContainedControlsAdditionDesignTime.Enabled = True
            End If
        Case WM_LBUTTONUP ' UserControl message, only in design mode (Not Ambient.UserMode). To avoid the IDE to start dragging the control on mouse down when the developer clicks to change the selected tab
            If mBtnDown Then
                mBtnDown = False
                SendMessage hWnd, WM_LBUTTONDOWN, wParam, lParam
            End If
        Case WM_MOVE
            RedrawWindow hWnd, ByVal 0, 0, RDW_INVALIDATE Or RDW_ALLCHILDREN
        Case WM_PAINT ' contained controls paint messages, when the control is themed and ChangeControlsBackColor = True (only at run time, Ambient.UserMode)
            
            Dim iUpdateRect As RECT
            Dim iControlRect As RECT
            Dim iDestDC As Long
            Dim iWidth As Long
            Dim iHeight As Long
            Dim iTempDC As Long
            Dim iTempBmp As Long
            Dim iPs As PAINTSTRUCT
            Dim iBKColor As Long
            Dim iPt As POINTAPI
            Dim iBrush As Long
            Dim iTop As Long
            Dim iLeft As Long
            Dim iColor As Long
            Dim iFillRect As RECT
            
            If GetUpdateRect(hWnd, iUpdateRect, 0&) <> 0& Then
                Call BeginPaint(hWnd, iPs)
                
                iDestDC = iPs.hDC
                GetWindowRect hWnd, iControlRect
                
                iPt.X = iControlRect.Left + iPs.rcPaint.Left
                iPt.Y = iControlRect.Top + iPs.rcPaint.Top
                ScreenToClient hWnd, iPt
                iControlRect.Left = iControlRect.Left - iPt.X
                iControlRect.Top = iControlRect.Top - iPt.Y
                
                iTempDC = CreateCompatibleDC(iDestDC)
                iTempBmp = CreateCompatibleBitmap(iDestDC, iControlRect.Right - iControlRect.Left, iControlRect.Bottom - iControlRect.Top)
                DeleteObject SelectObject(iTempDC, iTempBmp)
                
                CallOldWindowProc hWnd, iMsg, iTempDC, lParam
                
                iWidth = iControlRect.Right - iControlRect.Left
                iHeight = iControlRect.Bottom - iControlRect.Top
                
                iPt.X = iControlRect.Left + iPs.rcPaint.Left
                iPt.Y = iControlRect.Top + iPs.rcPaint.Top
                ScreenToClient mUserControlHwnd, iPt
                
                
                If mChangeControlsBackColor Then
                    If mShowDisabledState And (Not mEnabled) Then
                        iColor = mBackColorSelectedTabDisabled
                    Else
                        iColor = mBackColorSelectedTab
                    End If
                Else
                    iColor = vbButtonFace
                End If
                TranslateColor iColor, 0&, iBKColor
                
                ' set the part of the update rect of the control that must be painted with the backgroung bitmap because is inside the tab body
                If iPt.Y < mClientRect.Top Then
                    iHeight = iHeight - (mClientRect.Top - 1 - iPt.Y)
                    iTop = (mClientRect.Top - 1 - iPt.Y)
                    iPt.Y = mClientRect.Top - 1
                    If (mClientRect.Top + iHeight - 2) > mClientRect.Bottom Then
                        iHeight = mClientRect.Bottom - mClientRect.Top + 2
                    End If
                ElseIf iPt.Y + iHeight > mClientRect.Bottom Then
                    iHeight = mClientRect.Bottom - iPt.Y
                    iTop = 0
                End If
                
                If iPt.X < mClientRect.Left Then
                    iWidth = iWidth - (mClientRect.Left - iPt.X)
                    iLeft = (mClientRect.Left - 1 - iPt.X)
                    iPt.X = mClientRect.Left - 1
                    If (mClientRect.Left + iWidth - 2) > mClientRect.Right Then
                        iWidth = mClientRect.Right - mClientRect.Left + 2
                    End If
                ElseIf iPt.X + iWidth > mClientRect.Right Then
                    iWidth = mClientRect.Right - iPt.X
                    iLeft = 0
                End If
                
                ' iLeft and iTop: from where to paint into the control in coordinates of the control
                ' iWidth and iHeight: the size of the image to be painted into the control
                ' iPt.X and iPt.Y: the position in the UserControl from where to take the image to be painted, in coordinales of the UserControl
                
                'the rest of the update rect that was not painted must be filled with the tab backcolor (if there are parts that are outside the tab body)
                
                If iTop > iPs.rcPaint.Top Then  ' there is a space over the painted region that must be filled
                    iFillRect = iPs.rcPaint
                    iFillRect.Bottom = iTop + 1
                    If iFillRect.Bottom > iFillRect.Top Then
                        iBrush = CreateSolidBrush(iBKColor)
                        FillRect iDestDC, iFillRect, iBrush
                        DeleteObject iBrush
                    End If
                End If
                If iLeft > iPs.rcPaint.Left Then   ' there is a space over the painted region that must be filled
                    iFillRect = iPs.rcPaint
                    iFillRect.Right = iLeft + 1
                    If iFillRect.Right > iFillRect.Left Then
                        iBrush = CreateSolidBrush(iBKColor)
                        FillRect iDestDC, iFillRect, iBrush
                        DeleteObject iBrush
                    End If
                End If
                If (iTop + iHeight) < iPs.rcPaint.Bottom Then
                    iFillRect = iPs.rcPaint
                    iFillRect.Top = (iTop + iHeight)
                    If iFillRect.Bottom > iFillRect.Top Then
                        iBrush = CreateSolidBrush(iBKColor)
                        FillRect iDestDC, iFillRect, iBrush
                        DeleteObject iBrush
                    End If
                End If
                If (iLeft + iWidth) < iPs.rcPaint.Right Then
                    iFillRect = iPs.rcPaint
                    iFillRect.Left = (iLeft + iWidth)
                    If iFillRect.Right > iFillRect.Left Then
                        iBrush = CreateSolidBrush(iBKColor)
                        FillRect iDestDC, iFillRect, iBrush
                        DeleteObject iBrush
                    End If
                End If

                If (iHeight > 0) And (iWidth > 0) Then
                    BitBlt iDestDC, iLeft, iTop, iWidth, iHeight, UserControl.hDC, iPt.X, iPt.Y, vbSrcCopy
                End If
                TransparentBlt iDestDC, iPs.rcPaint.Left, iPs.rcPaint.Top, iPs.rcPaint.Right - iPs.rcPaint.Left, iPs.rcPaint.Bottom - iPs.rcPaint.Top, iTempDC, iPs.rcPaint.Left, iPs.rcPaint.Top, iPs.rcPaint.Right - iPs.rcPaint.Left, iPs.rcPaint.Bottom - iPs.rcPaint.Top, iBKColor
                DeleteDC iTempDC
                DeleteObject iTempBmp
                Call EndPaint(hWnd, iPs)
                IBSSubclass_WindowProc = 0
            Else
                IBSSubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
            End If
        Case WM_GETDPISCALEDSIZE
            Dim iPrev As Long
            
            iPrev = mLeftOffsetToHide
            SetLeftOffsetToHide Int(1440 / wParam)
            If mLeftOffsetToHide <> iPrev Then
                mPendingLeftOffset = iPrev - mLeftOffsetToHide
                DoPendingLeftOffset
            End If
        Case WM_SETCURSOR
            If mCurrentMousePointerIsHand Then
                bConsume = True
                IBSSubclass_WindowProc = 1
                If GetCursor <> IDC_HAND Then
                    SetCursor mHandIconHandle
                End If
            End If
        Case WM_ACTIVATE
            If wParam = WA_INACTIVE Then
                If WindowIsTDIChild(lParam) Then
                    PostMessage hWnd, WM_NCACTIVATE, WA_ACTIVE, 0&
                End If
            End If
        Case WM_ACTIVATEAPP
            If wParam = WA_INACTIVE Then
                PostMessage hWnd, WM_NCACTIVATE, WA_INACTIVE, 0&
                mAppDeactivated = True
            End If
        Case WM_STYLECHANGING
            Dim iST As STYLESTRUCT
            Const WS_CAPTION As Long = &HC00000
            Const WS_THICKFRAME As Long = &H40000
            Dim iFormCaption As String
            
            CopyMemory iST, ByVal lParam, Len(iST)
            iST.styleNew = iST.styleNew And Not (WS_CAPTION Or WS_THICKFRAME)
            CopyMemory ByVal lParam, iST, Len(iST)
            
            iFormCaption = GetWindowCaption(hWnd)
            iTab = TDIGetTabByFormHwnd(hWnd)
            If iTab > -1 Then
                TabCaption(iTab) = iFormCaption
            End If
        Case WM_SHOWWINDOW
            PostDrawMessage
        Case WM_DESTROY
            iHwnd = 0
            On Error Resume Next
            iHwnd = mTDISubclassedFormsHwnds(CStr(hWnd))
            On Error GoTo 0
            If iHwnd <> 0 Then
                DetachMessage Me, hWnd, WM_STYLECHANGING
                DetachMessage Me, hWnd, WM_DESTROY
                DetachMessage Me, hWnd, WM_SHOWWINDOW
                mTDISubclassedFormsHwnds.Remove CStr(hWnd)
            End If
    End Select
End Function

Private Function WindowIsTDIChild(ByVal nHwnd As Long) As Boolean
    Dim c As Long
    
    For c = 1 To mTabs - 1
        If mTDIModeFormsFormData_FormHwnd(c) = nHwnd Then
            WindowIsTDIChild = True
            Exit Function
        End If
    Next
End Function

Private Sub mFont_FontChanged(ByVal PropertyName As String)
    If Not mDrawing Then
        If PropertyName = "Bold" Then
            picDraw.Font.Bold = mFont.Bold
        ElseIf PropertyName = "Charset" Then
            picDraw.Font.Charset = mFont.Charset
        ElseIf PropertyName = "Italic" Then
            picDraw.Font.Italic = mFont.Italic
        ElseIf PropertyName = "Name" Then
            picDraw.Font.Name = mFont.Name
        ElseIf PropertyName = "Size" Then
            picDraw.Font.Size = mFont.Size
        ElseIf PropertyName = "Strikethrough" Then
            picDraw.Font.Strikethrough = mFont.Strikethrough
        ElseIf PropertyName = "Underline" Then
            picDraw.Font.Underline = mFont.Underline
        ElseIf PropertyName = "Weight" Then
            picDraw.Font.Weight = mFont.Weight
        End If
        mSetAutoTabHeightPending = True
        DrawDelayed
    End If
End Sub

Private Sub mForm_Load()
    UserControl_Show
End Sub

Private Sub mTabIconFontsEventsHandler_FontChanged(ByVal PropertyName As String)
    If mAutoTabHeight Then
        If mPropertiesReady Then
            mSetAutoTabHeightPending = True
            DrawDelayed
        End If
    End If
End Sub

Private Sub mThemesCollection_ThemeRemoved()
    SetPropertyChanged
End Sub

Private Sub mThemesCollection_ThemeRenamed()
    SetPropertyChanged
End Sub


Private Sub tmrCancelDoubleClick_Timer()
    tmrCancelDoubleClick.Enabled = False
End Sub

Private Sub tmrCheckContainedControlsAdditionDesignTime_Timer()
    If IsMouseButtonPressed(ntMBLeft) Then Exit Sub
    If mBackStyle = ntOpaque Then tmrCheckContainedControlsAdditionDesignTime.Enabled = False
    
    If UserControlContainedControlsCount <> mLastContainedControlsCount Then
        mLastContainedControlsCount = UserControlContainedControlsCount
        SetControlsBackColor mBackColorSelectedTab
        SetControlsProperForeColor
        If mControlIsThemed Or (mBackStyle = ntTransparent) Then
            mSubclassControlsPaintingPending = True
            RedrawWindow mUserControlHwnd, ByVal 0&, 0&, RDW_INVALIDATE Or RDW_ALLCHILDREN
            Draw
        End If
    ElseIf (Not mAmbientUserMode) And (mBackStyle = ntTransparent) Then
        Dim iStr As String
        
        iStr = GetContainedControlsPositionsStr
        If iStr <> mLastContainedControlsPositionsStr Then
            mLastContainedControlsPositionsStr = iStr
            mSubclassControlsPaintingPending = True
            RedrawWindow mUserControlHwnd, ByVal 0&, 0&, RDW_INVALIDATE Or RDW_ALLCHILDREN
            Draw
        End If
    End If
End Sub

Private Function GetContainedControlsPositionsStr() As String
    Dim iCtl As Object
    Dim iLeft As Long
    Dim iWidth As Long
    
    On Error Resume Next
    For Each iCtl In UserControlContainedControls
        iLeft = -mLeftOffsetToHide
        iLeft = iCtl.Left
        If iLeft > -mLeftOffsetToHide Then
            iWidth = -1
            iWidth = iCtl.Width
            If iWidth <> -1 Then
                GetContainedControlsPositionsStr = GetContainedControlsPositionsStr & CStr(iLeft) & "," & CStr(iCtl.Top) & "," & CStr(iWidth) & "," & CStr(iCtl.Height) & "|"
            End If
        End If
    Next
    'On Error GoTo 0
    
End Function
    

Private Function IsMouseButtonPressed(ByVal nButton As NTMouseButtonsConstants) As Boolean
    Dim iButton As Long
    
    iButton = nButton
    If GetSystemMetrics(SM_SWAPBUTTON) <> 0 Then
        If nButton = ntMBLeft Then
            iButton = VK_RBUTTON
        ElseIf nButton = ntMBRight Then
            iButton = VK_LBUTTON
        End If
    End If
    IsMouseButtonPressed = GetAsyncKeyState(iButton) <> 0
End Function

Private Sub tmrCheckDuplicationByIDEPaste_Timer()
    If (Not mAmbientUserMode) Then
        If Not IsMsgBoxShown Then
            tmrCheckDuplicationByIDEPaste.Enabled = False
            CheckContainedControlsConsistency
        End If
    Else
        tmrCheckDuplicationByIDEPaste.Enabled = False
    End If
End Sub

Private Sub tmrCheckTabDrag_Timer()
    tmrCheckTabDrag.Enabled = False
End Sub

Private Sub tmrDraw_Timer()
    If mRedraw = False Then tmrDraw.Enabled = False
    Draw
End Sub

Private Sub tmrHighlightIcon_Timer()
    If Not mMouseIsOverIcon Then
        tmrHighlightIcon.Enabled = False
        Draw
    End If
End Sub

Private Sub tmrPreHighlightIcon_Timer()
    tmrPreHighlightIcon.Enabled = False
    If mMouseIsOverIcon Then
        Draw
    Else
        tmrHighlightIcon.Enabled = False
    End If
End Sub

Private Sub tmrRestoreDropMode_Timer()
    Dim t As Long
    Dim iPt As POINTAPI
    
    GetCursorPos iPt
    ScreenToClient mUserControlHwnd, iPt
    
    t = GetTabAtXY(iPt.X * Screen_TwipsPerPixelX, iPt.Y * Screen_TwipsPerPixely)
    If t = mTab Then
        UserControl.OLEDropMode = ssOLEDropManual
        tmrRestoreDropMode.Enabled = False
    End If
End Sub

Private Sub tmrShowTabTTT_Timer()
    Static sFormHwnd As Long
    
    tmrShowTabTTT.Enabled = False
    If tmrShowTabTTT.Tag <> mTabUnderMouse Then Exit Sub 'a protection, just in case
    Set mToolTipEx = New cToolTipEx
    mToolTipEx.TipText = mTabData(mTabUnderMouse).ToolTipText
    mToolTipEx.Style = vxTTStandard
    mToolTipEx.CloseButton = False
    mToolTipEx.DelayTimeSeconds = 0
    mToolTipEx.RightToLeft = mRightToLeft
    If sFormHwnd = 0 Then sFormHwnd = GetAncestor(UserControl.ContainerHwnd, GA_ROOT)
    mToolTipEx.Create sFormHwnd
End Sub

Private Sub tmrTabDragging_Timer()
    Draw
End Sub

Private Sub tmrTabTransition_Timer()
    Dim iValue As Long
    Const LWA_ALPHA = &H2&
    Dim iStepValue As Long
    
    If mTabTransition = ntTransitionSlower Then
        iStepValue = 5
    ElseIf mTabTransition = ntTransitionSlow Then
        iStepValue = 10
    ElseIf mTabTransition = ntTransitionFast Then
        iStepValue = 20
    ElseIf mTabTransition = ntTransitionFaster Then
        iStepValue = 25
    Else
        iStepValue = 15 ' normal
    End If
    
    If mTabTransition_Step * iStepValue > 255 Then
        HidePicCover
        tmrTabTransition.Enabled = False
    Else
        mTabTransition_Step = mTabTransition_Step + 1
        iValue = 255 - mTabTransition_Step * iStepValue
        If iValue < 0 Then iValue = 0
        SetLayeredWindowAttributes picCover.hWnd, 0, iValue, LWA_ALPHA
    End If
End Sub

Private Sub tmrSubclassControls_Timer()
    tmrSubclassControls.Enabled = False
    SubclassControlsPainting
End Sub

Private Sub tmrHighlightEffect_Timer()
    mHighlightEffect_Step = mHighlightEffect_Step + 2
    If mHighlightEffect_Step >= 10 Then mHighlightEffect_Step = 10
    If mHighlightIntensity = ntHighlightIntensityStrong Then
        mGlowColor = mHighlightEffectColors_Strong(mHighlightEffect_Step)
    Else
        mGlowColor = mHighlightEffectColors_Light(mHighlightEffect_Step)
    End If
    If mTabUnderMouse > -1 Then
        If Not IsEmpty(mTabData(mTabUnderMouse).CustomColors.HighlightColor) Then
            If mHighlightIntensity = ntHighlightIntensityStrong Then
                mTabData(mTabUnderMouse).CustomColors.GlowColor = mTabData(mTabUnderMouse).CustomColors.HighlightEffectColors_Strong(mHighlightEffect_Step)
            Else
                mTabData(mTabUnderMouse).CustomColors.GlowColor = mTabData(mTabUnderMouse).CustomColors.HighlightEffectColors_Light(mHighlightEffect_Step)
            End If
        End If
    End If
    mFlatBarGlowColor = mFlatBarHighlightEffectColors(mHighlightEffect_Step)
    If (Not IsEmpty(mTabData(mTabUnderMouse).CustomColors.FlatBarColorInactive)) Or (Not IsEmpty(mTabData(mTabUnderMouse).CustomColors.FlatBarColorHighlight)) Then
        mTabData(mTabUnderMouse).CustomColors.FlatBarGlowColor = mTabData(mTabUnderMouse).CustomColors.FlatBarHighlightEffectColors(mHighlightEffect_Step)
    End If
    Draw
    If mHighlightEffect_Step = 10 Then
        tmrHighlightEffect.Enabled = False
        If mHighlightIntensity = ntHighlightIntensityStrong Then
            mGlowColor = mHighlightEffectColors_Strong(10)
        Else
            mGlowColor = mHighlightEffectColors_Light(10)
        End If
        If mTabUnderMouse > -1 Then
            If Not IsEmpty(mTabData(mTabUnderMouse).CustomColors.HighlightColor) Then
                If mHighlightIntensity = ntHighlightIntensityStrong Then
                    mTabData(mTabUnderMouse).CustomColors.GlowColor = mTabData(mTabUnderMouse).CustomColors.HighlightEffectColors_Strong(10)
                Else
                    mTabData(mTabUnderMouse).CustomColors.GlowColor = mTabData(mTabUnderMouse).CustomColors.HighlightEffectColors_Light(10)
                End If
            End If
        End If
    End If
End Sub

Private Sub tmrTabMouseLeave_Timer()
    Dim iPt As POINTAPI
    Dim iHwnd As Long
    
    'If mTabUnderMouse = -1 Then Exit Sub
    GetCursorPos iPt
    iHwnd = WindowFromPoint(iPt.X, iPt.Y)
    If iHwnd <> mUserControlHwnd Then
        tmrTabMouseLeave.Enabled = False
        RaiseEvent_TabMouseLeave (mTabUnderMouse)
        mTabUnderMouse = -1
    End If
End Sub

Private Sub tmrTDIIconColor_Timer()
    tmrTDIIconColor.Enabled = False
    mIconColorMouseHover = mTDIIconColorMouseHover
    mIconColorMouseHoverSelectedTab = mTDIIconColorMouseHover
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    Dim iChr As String
    Dim iPos As Long
    
    iChr = LCase$(Chr$(KeyAscii))
    iPos = InStr(mTab + 2, mAccessKeys, iChr)
    If iPos = 0 Then
        iPos = InStr(mAccessKeys, iChr)
    End If
    If iPos > 0 Then
        TabSel = iPos - 1
    End If
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
    If PropertyName = "ScaleUnits" Then
        SetPropertyChanged "TabHeight"
        SetPropertyChanged "TabMaxWidth"
        SetPropertyChanged "TabMinWidth"
        SetPropertyChanged "HighlightTabExtraHeight"
    ElseIf PropertyName = "BackColor" Then
        If mBackColorIsFromAmbient Then BackColor = Ambient.BackColor
        If mBackColorTabsIsFromAmbient Then
            BackColorTabs = Ambient.BackColor
            If mBackColorSelectedTab_IsAutomatic Then BackColorSelectedTab = BackColorTabs
        End If
    ElseIf PropertyName = "ForeColor" Then
        If mForeColorIsFromAmbient Then ForeColor = Ambient.ForeColor
        If mIconColorIsFromAmbient Then IconColor = Ambient.ForeColor
    End If
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_GotFocus()
    If Not mHasFocus Then
        mHasFocus = True
        'PostDrawMessage
        tmrDraw.Enabled = True
    End If
End Sub

Friend Sub StoreVisibleControlsInSelectedTab()
    Dim iCtl As Object
    Dim iCtlName As String
    
    On Error Resume Next
    Set mTabData(mTab).Controls = New Collection
    For Each iCtl In UserControlContainedControls
        If TypeName(iCtl) = "Line" Then
            Err.Clear
            If iCtl.X1 > -mLeftThresholdHided Then
                If Err.Number = 0 Then
                    iCtlName = ControlName(iCtl)
                    mTabData(mTab).Controls.Add iCtlName, iCtlName
                End If
            End If
        Else
            Err.Clear
            If iCtl.Left > -mLeftThresholdHided Then
                If Err.Number = 0 Then
                    iCtlName = ControlName(iCtl)
                    mTabData(mTab).Controls.Add iCtlName, iCtlName
                End If
            End If
        End If
    Next
    Err.Clear
End Sub

Private Sub UserControl_Hide()
    Dim c As Long
    Const WM_SYSCOMMAND As Long = &H112
    Const SC_CLOSE As Long = &HF060&
    Const SW_HIDE As Long = 0
    Const WM_CLOSE As Long = &H10&
    Dim iHwnd As Long
    Dim iLng As Long
    
    mBodyReset = True
    If (mTDIMode = ntTDIModeForms) And mAmbientUserMode Then
        iLng = -1
        On Error Resume Next
        iLng = LBound(mTDIModeFormsFormData_FormHwnd)
        On Error GoTo 0
        If iLng > -1 Then
            For c = mTabs - 1 To 0 Step -1
                iHwnd = mTDIModeFormsFormData_FormHwnd(mTabData(c).Data)
                If iHwnd <> 0 Then
                    If IsWindow(iHwnd) <> 0 Then
                        ShowWindow iHwnd, SW_HIDE
                        SetParent iHwnd, mTDIModeFormsFormData_OldParentHwnd(mTabData(c).Data)
                        If mTDIFormWithoutCloseButton(mTabData(c).Data) Then
                            PostMessage iHwnd, WM_CLOSE, 0, 0
                        Else
                            SendMessage iHwnd, WM_SYSCOMMAND, SC_CLOSE, 0&
                        End If
                    End If
                End If
            Next
        End If
        For c = picTDIFormContainer.LBound + 1 To picTDIFormContainer.UBound
            Unload picTDIFormContainer(c)
        Next
        UninstallCBTHook
        mTDIModeFormsUnhooked = True
    End If
End Sub

Private Sub UserControl_Initialize()
    Debug.Assert MakeTrue(mInIDE)
    mTabUnderMouse = -1
    Set mParentControlsTabStop = New Collection
    Set mParentControlsUseMnemonic = New Collection
    Set mContainedControlsThatAreContainers = New Collection
    Set mSubclassedControlsForPaintingHwnds = New Collection
    Set mSubclassedFramesHwnds = New Collection
    Set mSubclassedControlsForMoveHwnds = New Collection
    Set mTDISubclassedFormsHwnds = New Collection
    Set mTabIconFontsEventsHandler = New cFontEventHandlers
    mRedraw = True
    mTabOrientation_Prev = -1
    SetDPI
    mMouseIsOverIcon_Tab = -1
    mTabIconDistanceToCaptionDPIScaled = cTabIconDistanceToCaption * mDPIScale
    mIconClickExtendDPIScaled = cIconClickExtend * mDPIScale
    mControlsForeColor_PrevColor = -1
    mControlsBackColor_PrevColor = -1
End Sub

Private Sub UserControl_InitProperties()
    Dim c As Long
    
    On Error Resume Next
    mUserControlHwnd = UserControl.hWnd
    mAmbientUserMode = Ambient.UserMode
    mDefaultTabHeight = pScaleY(cPropDef_TabHeight, vbTwips, vbHimetric)
    If mDefaultTabHeight = 0 Then
        mDefaultTabHeight = 419.8055
    End If
    If mAmbientUserMode Then
        If TypeOf UserControl.Parent Is Form Then
           Set mForm = UserControl.Parent
        End If
    End If
    On Error GoTo 0
    
    mTabs = 3
    ReDim mTabData(-1 To mTabs - 1)
    mTab = 0
    
    mTabsEndFreeSpace = cPropDef_TabsEndFreeSpace
    mSubclassingMethod = cPropDef_SubclassingMethod
    mChangeControlsBackColor = cPropDef_ChangeControlsBackColor: PropertyChanged "ChangeControlsBackColor"
    mChangeControlsForeColor = cPropDef_ChangeControlsForeColor: PropertyChanged "ChangeControlsForeColor"
    Set mFont = Ambient.Font
    
    SetDefaultPropertyValuesForThemedProperties
    
    mTabsPerRow = cPropDef_TabsPerRow
    mIconColorMouseHover = Ambient.ForeColor
    mIconColorMouseHoverSelectedTab = Ambient.ForeColor
    mRightToLeft = Ambient.RightToLeft
    SetScaleMode
    mEnabled = True
    Set mDefaultIconFont = New StdFont
    mDefaultIconFont.Name = cPropDef_IconFontName
    mDefaultIconFont.Size = cPropDef_IconFontSize
    For c = 0 To mTabs - 1
        Set mTabData(c).Controls = New Collection
        mTabData(c).Enabled = True
        mTabData(c).Visible = True
        mTabData(c).Caption = "Tab " & CStr(c)
        mTabData(c).OriginalIndex = c
    Next c
    mTabData(mTab).Selected = True
    mUseMaskColor = cPropDef_UseMaskColor
    mTabHeight = mDefaultTabHeight
    mIconAlignment = cPropDef_IconAlignment
    mHandleHighContrastTheme = cPropDef_HandleHighContrastThem
    mAutoTabHeight = cPropDef_AutoTabHeight
    mOLEDropOnOtherTabs = cPropDef_OLEDropOnOtherTabs
    mCanReorderTabs = cPropDef_CanReorderTabs
    mTDIMode = cPropDef_TDIMode
    mTabTransition = cPropDef_TabTransition
    mCaptionAlignment = cPropDef_CaptionAlignment: PropertyChanged "CaptionAlignment"
    If mCaptionAlignment = ntAlignmentLeft Then
        mAlignmentDTFlag = DT_LEFT
    ElseIf mCaptionAlignment = ntAlignmentRight Then
        mAlignmentDTFlag = DT_RIGHT
    Else
        mAlignmentDTFlag = DT_CENTER
    End If

    If mHandleHighContrastTheme Then CheckHighContrastTheme
    mPropertiesReady = True
    
    mSubclassed = mSubclassingMethod <> ntSMDisabled
#If NOSUBCLASSINIDE Then
    If mInIDE Then
        mSubclassed = False
    End If
#End If
    
    If mSubclassed Then
        gSubclassWithSetWindowLong = (mSubclassingMethod = ntSMSetWindowLong) Or (mSubclassingMethod = ntSM_SWLOnlyUserControl)
        mOnlySubclassUserControl = (mSubclassingMethod = ntSM_SWSOnlyUserControl) Or (mSubclassingMethod = ntSM_SWLOnlyUserControl)
        SubclassUserControl
    Else
        mFormIsActive = True
    End If
    UserControl.Size 2500, 1700
    
    If mAmbientUserMode Then
        mHandIconHandle = LoadCursor(ByVal 0&, IDC_HAND)
    End If
    mControlJustAdded = True
End Sub

Private Sub SubclassUserControl()
    If mAmbientUserMode Then
        AttachMessage Me, mUserControlHwnd, WM_MOUSEACTIVATE
        AttachMessage Me, mUserControlHwnd, WM_SETFOCUS
        AttachMessage Me, mUserControlHwnd, WM_DRAW
        AttachMessage Me, mUserControlHwnd, WM_INIT
        AttachMessage Me, mUserControlHwnd, WM_SETCURSOR
        PostMessage mUserControlHwnd, WM_INIT, 0&, 0&
        mCanPostDrawMessage = True
    Else
        AttachMessage Me, mUserControlHwnd, WM_LBUTTONDOWN
        AttachMessage Me, mUserControlHwnd, WM_LBUTTONUP
        AttachMessage Me, mUserControlHwnd, WM_LBUTTONDBLCLK
        AttachMessage Me, mUserControlHwnd, WM_MOUSEMOVE
    End If
End Sub

Private Sub SubclassForm()
    If (mFormHwnd <> 0) And (Not mOnlySubclassUserControl) Then
        AttachMessage Me, mFormHwnd, WM_SYSCOLORCHANGE
        AttachMessage Me, mFormHwnd, WM_THEMECHANGED
        AttachMessage Me, mFormHwnd, WM_NCACTIVATE
        AttachMessage Me, mFormHwnd, WM_GETDPISCALEDSIZE
        If mTDIMode = ntTDIModeForms Then
            AttachMessage Me, mFormHwnd, WM_ACTIVATE
            AttachMessage Me, mFormHwnd, WM_ACTIVATEAPP
        End If
    End If
End Sub

Friend Sub SetDefaultPropertyValuesForThemedProperties(Optional nSetControlsColors As Boolean)
    Dim iBackColor_Prev As Long
    Dim iForeColor_Prev As Long
    
    iBackColor_Prev = IIf(mEnabled Or Not mShowDisabledState, mBackColorSelectedTab, mBackColorSelectedTabDisabled)
    iForeColor_Prev = mForeColorSelectedTab
    
    mBackColor = Ambient.BackColor: PropertyChanged "BackColor"
    mForeColor = Ambient.ForeColor: PropertyChanged "ForeColor"
    mForeColorSelectedTab = Ambient.ForeColor: PropertyChanged "ForeColorSelectedTab"
    mForeColorHighlighted = Ambient.ForeColor: PropertyChanged "ForeColorHighlighted"
    mFlatTabBorderColorHighlight = Ambient.ForeColor: PropertyChanged "FlatTabBorderColorHighlight"
    mFlatTabBorderColorSelectedTab = Ambient.ForeColor: PropertyChanged "FlatTabBorderColorSelectedTab"
    mBackColorTabs = Ambient.BackColor: PropertyChanged "BackColorTabs"
    mIconColorSelectedTab = Ambient.ForeColor: PropertyChanged "IconColorSelectedTab"
    mIconColorTabHighlighted = Ambient.ForeColor: PropertyChanged "IconColorTabHighlighted"
    
    mBackColorIsFromAmbient = True
    mForeColorIsFromAmbient = True
    mIconColorIsFromAmbient = True
    mBackColorTabsIsFromAmbient = True
    mStyle = cPropDef_Style
    If Not nSetControlsColors Then
        If cPropDef_Style = ntStyleWindows Then
            If ((Ambient.BackColor = vbButtonFace) And (Ambient.ForeColor = vbButtonText)) Or (TypeName(Parent) = "MDIForm") Then
                mStyle = cPropDef_Style
            Else
                mStyle = ntStyleFlat
            End If
        Else
            mStyle = cPropDef_Style
        End If
    End If
    PropertyChanged "Style"
    mVisualStyles = (mStyle = ntStyleWindows) 'And IsThemed
    mBackColorSelectedTab_IsAutomatic = True: mBackColorSelectedTab = GetAutomaticBackColorSelectedTab: PropertyChanged "BackColorSelectedTab"
    mWordWrap = cPropDef_WordWrap: PropertyChanged "WordWrap"
    mMaskColor = cPropDef_MaskColor: PropertyChanged "MaskColor"
    mShowFocusRect = cPropDef_ShowFocusRect: PropertyChanged "ShowFocusRect"
    mShowDisabledState = cPropDef_ShowDisabledState: PropertyChanged "ShowDisabledState"
    mHighlightEffect = cPropDef_HighlightEffect: PropertyChanged "HighlightEffect"
    mTabWidthStyle = cPropDef_TabWidthStyle: PropertyChanged "TabWidthStyle"
    mShowRowsInPerspective = cPropDef_ShowRowsInPerspective: PropertyChanged "ShowRowsInPerspective"
    mTabSeparation = cPropDef_TabSeparation: PropertyChanged "TabSeparation"
    mTabSeparationDPIScaled = mTabSeparation * mDPIScale
    mTabAppearance = cPropDef_TabAppearance: PropertyChanged "TabAppearance"
    mAutoRelocateControls = cPropDef_AutoRelocateControls: PropertyChanged "AutoRelocateControls"
    mSoftEdges = cPropDef_SoftEdges: PropertyChanged "SoftEdges"
    mBackStyle = cPropDef_BackStyle: PropertyChanged "BackStyle"
    mFlatBarColorSelectedTab = cPropDef_FlatBarColorSelectedTab: PropertyChanged "FlatBarColorSelectedTab"
    
    mFlatBarColorHighlight_IsAutomatic = True: PropertyChanged "FlatBarColorHighlight"
    mFlatBarColorInactive_IsAutomatic = True: PropertyChanged "FlatBarColorInactive"
    mFlatTabsSeparationLineColor_IsAutomatic = True: PropertyChanged "FlatTabsSeparationLineColor"
    mFlatBodySeparationLineColor_IsAutomatic = True: PropertyChanged "FlatBodySeparationLineColor"
    mFlatBorderColor_IsAutomatic = True: PropertyChanged "FlatBorderColor"
    mHighlightColor_IsAutomatic = True: PropertyChanged "HighlightColor"
    mHighlightColorSelectedTab_IsAutomatic = True: PropertyChanged "HighlightColorSelectedTab"
    mFlatRoundnessTop = cPropDef_FlatRoundnessTop: PropertyChanged "FlatRoundnessTop"
    mFlatRoundnessTopDPIScaled = mFlatRoundnessTop * mDPIScale
    mFlatRoundnessBottom = cPropDef_FlatRoundnessBottom: PropertyChanged "FlatRoundnessBottom"
    mFlatRoundnessBottomDPIScaled = mFlatRoundnessBottom * mDPIScale
    mFlatRoundnessTabs = cPropDef_FlatRoundnessTabs: PropertyChanged "FlatRoundnessTabs"
    mFlatRoundnessTabsDPIScaled = mFlatRoundnessTabs * mDPIScale
    mTabMousePointerHand = cPropDef_TabMousePointerHand: PropertyChanged "TabMousePointerHand"
    mIconColor = Ambient.ForeColor: PropertyChanged "IconColor"
    mHighlightMode = cPropDef_HighlightMode: PropertyChanged "HighlightMode"
    mHighlightModeSelectedTab = cPropDef_HighlightModeSelectedTab: PropertyChanged "HighlightModeSelectedTab"
    mFlatBorderMode = cPropDef_FlatBorderMode: PropertyChanged "FlatBorderMode"
    mFlatBarHeight = cPropDef_FlatBarHeight: PropertyChanged "FlatBarHeight"
    mFlatBarHeightDPIScaled = mFlatBarHeight * mDPIScale
    mFlatBarGripHeight = cPropDef_FlatBarGripHeight: PropertyChanged "FlatBarGripHeight"
    mFlatBarGripHeightDPIScaled = mFlatBarGripHeight * mDPIScale
    mFlatBarPosition = cPropDef_FlatBarPosition: PropertyChanged "FlatBarPosition"
    mHighlightTabExtraHeight = Round(ScaleY(cPropDef_HighlightTabExtraHeight, vbTwips, vbHimetric)): PropertyChanged "HighlightTabExtraHeight"
    mFlatBodySeparationLineHeight = cPropDef_FlatBodySeparationLineHeight: PropertyChanged "FlatBodySeparationLineHeight"
    mFlatBodySeparationLineHeightDPIScaled = mFlatBodySeparationLineHeight * mDPIScale
    mTabMaxWidth = cPropDef_TabMaxWidth: PropertyChanged "TabMaxWidth"
    mTabMinWidth = cPropDef_TabMinWidth: PropertyChanged "TabMinWidth"
    
    SetFont
    If mTabAppearance <> ntTAAuto Then
        mAppearanceIsFlat = (mTabAppearance = ntTAFlat)
    Else
        mAppearanceIsFlat = mStyle = ntStyleFlat
    End If
    SetHighlightMode
    mSetAutoTabHeightPending = True
    SetButtonFaceColor
    SetColors
    UserControl.BackColor = mBackColor
    If nSetControlsColors Then
        SetControlsBackColor IIf(mEnabled Or Not mShowDisabledState, mBackColorSelectedTab, mBackColorSelectedTabDisabled), iBackColor_Prev
        SetControlsProperForeColor IIf(mControlsForeColor_PrevColor = -1, iForeColor_Prev, mControlsForeColor_PrevColor)
    End If
    mNeedToDraw = True
End Sub

Friend Property Let ControlJustAdded(nValue As Boolean)
    mControlJustAdded = nValue
End Property

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim t As Long
    Dim iAgain As Boolean
    Dim iRow As Long
    Dim iHandled As Boolean
    
    RaiseEvent KeyDown(KeyCode, Shift)
    If KeyCode = vbKeyRight Then
        If (mTabOrientation = ssTabOrientationTop) Or (mTabOrientation = ssTabOrientationBottom) Then
            t = mTab + 1
            If t = mTabs Then t = 0
            Do Until mTabData(t).Enabled And mTabData(t).Visible
                t = t + 1
                If t = mTabs Then
                    If iAgain Then Exit Sub
                    t = 0
                    iAgain = True
                End If
            Loop
            TabSel = t
            iHandled = True
        ElseIf (mTabOrientation = ssTabOrientationLeft) Or (mTabOrientation = ntTabOrientationLeftHorizontal) Or (mTabOrientation = ssTabOrientationRight) Or (mTabOrientation = ntTabOrientationRightHorizontal) Then
            If (mRows > 1) And ((mTabWidthStyle2 = ntTWFixed) Or (mTabWidthStyle2 = ntTWTabCaptionWidth)) Then
                If (mTabOrientation = ssTabOrientationLeft) Or (mTabOrientation = ntTabOrientationLeftHorizontal) Then
                    iRow = 0
                Else
                    iRow = mTabData(mTab).RowPos - 1
                End If
                Do Until (mTabData(t).PosH = mTabData(mTab).PosH) And (mTabData(t).RowPos = iRow)
                    t = t - 1
                    If t = -1 Then t = mTabs - 1
                    Do Until mTabData(t).Enabled And mTabData(t).Visible
                        t = t - 1
                        If t = -1 Then
                            If iAgain Then Exit Sub
                            t = mTabs - 1
                            iAgain = True
                        End If
                    Loop
                Loop
                TabSel = t
                iHandled = True
            End If
        End If
        If Not iHandled Then
            SetFocusToNextControlInSameContainer True
        End If
    ElseIf KeyCode = vbKeyLeft Then
        If (mTabOrientation = ssTabOrientationTop) Or (mTabOrientation = ssTabOrientationBottom) Then
            t = mTab - 1
            If t = -1 Then t = mTabs - 1
            Do Until mTabData(t).Enabled And mTabData(t).Visible
                t = t - 1
                If t = -1 Then
                    If iAgain Then Exit Sub
                    t = mTabs - 1
                    iAgain = True
                End If
            Loop
            TabSel = t
            iHandled = True
        ElseIf (mTabOrientation = ssTabOrientationLeft) Or (mTabOrientation = ntTabOrientationLeftHorizontal) Or (mTabOrientation = ssTabOrientationRight) Or (mTabOrientation = ntTabOrientationRightHorizontal) Then
            If (mRows > 1) And ((mTabWidthStyle2 = ntTWFixed) Or (mTabWidthStyle2 = ntTWTabCaptionWidth)) Then
                If (mTabOrientation = ssTabOrientationLeft) Or (mTabOrientation = ntTabOrientationLeftHorizontal) Then
                    iRow = mTabData(mTab).RowPos - 1
                Else
                    iRow = 0
                End If
                Do Until (mTabData(t).PosH = mTabData(mTab).PosH) And (mTabData(t).RowPos = iRow)
                    t = t - 1
                    If t = -1 Then t = mTabs - 1
                    Do Until mTabData(t).Enabled And mTabData(t).Visible
                        t = t - 1
                        If t = -1 Then
                            If iAgain Then Exit Sub
                            t = mTabs - 1
                            iAgain = True
                        End If
                    Loop
                Loop
                TabSel = t
                iHandled = True
            End If
        End If
        If Not iHandled Then
            SetFocusToNextControlInSameContainer False
        End If
    ElseIf (KeyCode = vbKeyDown And ((Shift And vbCtrlMask) = 0)) Then
        If (mTabOrientation = ssTabOrientationTop) Or (mTabOrientation = ssTabOrientationBottom) Then
            If (mRows > 1) And ((mTabWidthStyle2 = ntTWFixed) Or (mTabWidthStyle2 = ntTWTabCaptionWidth)) Then
                t = mTab - 1
                If t = -1 Then t = mTabs - 1
                If mTabOrientation = ssTabOrientationTop Then
                    iRow = 0
                Else
                    iRow = mTabData(mTab).RowPos - 1
                End If
                Do Until (mTabData(t).PosH = mTabData(mTab).PosH) And (mTabData(t).RowPos = iRow)
                    t = t - 1
                    If t = -1 Then t = mTabs - 1
                    Do Until mTabData(t).Enabled And mTabData(t).Visible
                        t = t - 1
                        If t = -1 Then
                            If iAgain Then Exit Sub
                            t = mTabs - 1
                            iAgain = True
                        End If
                    Loop
                Loop
                TabSel = t
                iHandled = True
            End If
        End If
        If Not iHandled Then
            If (mTabOrientation = ssTabOrientationLeft) Or (mTabOrientation = ntTabOrientationLeftHorizontal) Then
                t = mTab - 1
                If t = -1 Then t = mTabs - 1
                Do Until mTabData(t).Enabled And mTabData(t).Visible
                    t = t - 1
                    If t = -1 Then
                        If iAgain Then Exit Sub
                        t = mTabs - 1
                        iAgain = True
                    End If
                Loop
                TabSel = t
                iHandled = True
            ElseIf (mTabOrientation = ssTabOrientationRight) Or (mTabOrientation = ntTabOrientationRightHorizontal) Then
                t = mTab + 1
                If t = mTabs Then t = 0
                Do Until mTabData(t).Enabled And mTabData(t).Visible
                    t = t + 1
                    If t = mTabs Then
                        If iAgain Then Exit Sub
                        t = 0
                        iAgain = True
                    End If
                Loop
                TabSel = t
                iHandled = True
            End If
        End If
        If Not iHandled Then
            SetFocusToNextControlInSameContainer True
        End If
    ElseIf (KeyCode = vbKeyUp And ((Shift And vbCtrlMask) = 0)) Then
        If (mTabOrientation = ssTabOrientationTop) Or (mTabOrientation = ssTabOrientationBottom) Then
            If (mRows > 1) And ((mTabWidthStyle2 = ntTWFixed) Or (mTabWidthStyle2 = ntTWTabCaptionWidth)) Then
                t = mTab - 1
                If t = -1 Then t = mTabs - 1
                If mTabOrientation = ssTabOrientationTop Then
                    iRow = mTabData(mTab).RowPos - 1
                Else
                    iRow = 0
                End If
                Do Until (mTabData(t).PosH = mTabData(mTab).PosH) And (mTabData(t).RowPos = iRow)
                    t = t - 1
                    If t = -1 Then t = mTabs - 1
                    Do Until mTabData(t).Enabled And mTabData(t).Visible
                        t = t - 1
                        If t = -1 Then
                            If iAgain Then Exit Sub
                            t = mTabs - 1
                            iAgain = True
                        End If
                    Loop
                Loop
                TabSel = t
                iHandled = True
            End If
        End If
        If Not iHandled Then
            If (mTabOrientation = ssTabOrientationLeft) Or (mTabOrientation = ntTabOrientationLeftHorizontal) Then
                t = mTab + 1
                If t = mTabs Then t = 0
                Do Until mTabData(t).Enabled And mTabData(t).Visible
                    t = t + 1
                    If t = mTabs Then
                        If iAgain Then Exit Sub
                        t = 0
                        iAgain = True
                    End If
                Loop
                TabSel = t
                iHandled = True
            ElseIf (mTabOrientation = ssTabOrientationRight) Or (mTabOrientation = ntTabOrientationRightHorizontal) Then
                t = mTab - 1
                If t = -1 Then t = mTabs - 1
                Do Until mTabData(t).Enabled And mTabData(t).Visible
                    t = t - 1
                    If t = -1 Then
                        If iAgain Then Exit Sub
                        t = mTabs - 1
                        iAgain = True
                    End If
                Loop
                TabSel = t
                iHandled = True
            End If
        End If
        If Not iHandled Then
            SetFocusToNextControlInSameContainer False
        End If
    ElseIf KeyCode = vbKeyTab And ((Shift And vbCtrlMask) > 0) And ((Shift And vbShiftMask) = 0) Then
        t = mTab + 1
        If t = mTabs Then t = 0
        Do Until mTabData(t).Enabled And mTabData(t).Visible
            t = t + 1
            If t = mTabs Then
                If iAgain Then Exit Sub
                t = 0
                iAgain = True
            End If
        Loop
        TabSel = t
    ElseIf KeyCode = vbKeyTab And ((Shift And vbCtrlMask) > 0) And ((Shift And vbShiftMask) > 0) Then
        t = mTab - 1
        If t = -1 Then t = mTabs - 1
        Do Until mTabData(t).Enabled And mTabData(t).Visible
            t = t - 1
            If t = -1 Then
                If iAgain Then Exit Sub
                t = mTabs - 1
                iAgain = True
            End If
        Loop
        TabSel = t
    ElseIf (KeyCode = vbKeyPageDown And ((Shift And vbCtrlMask) > 0)) Then
        If (mTabOrientation = ssTabOrientationTop) Or (mTabOrientation = ssTabOrientationBottom) Or (mTabOrientation = ssTabOrientationRight) Or (mTabOrientation = ntTabOrientationRightHorizontal) Then
            t = mTab + 1
            If t = mTabs Then t = 0
            Do Until mTabData(t).Enabled And mTabData(t).Visible
                t = t + 1
                If t = mTabs Then
                    If iAgain Then Exit Sub
                    t = 0
                    iAgain = True
                End If
            Loop
            TabSel = t
        Else
            t = mTab - 1
            If t = -1 Then t = mTabs - 1
            Do Until mTabData(t).Enabled And mTabData(t).Visible
                t = t - 1
                If t = -1 Then
                    If iAgain Then Exit Sub
                    t = mTabs - 1
                    iAgain = True
                End If
            Loop
            TabSel = t
        End If
    ElseIf KeyCode = vbKeyPageUp And ((Shift And vbCtrlMask) > 0) Then
        If (mTabOrientation = ssTabOrientationTop) Or (mTabOrientation = ssTabOrientationBottom) Or (mTabOrientation = ssTabOrientationRight) Or (mTabOrientation = ntTabOrientationRightHorizontal) Then
            t = mTab - 1
            If t = -1 Then t = mTabs - 1
            Do Until mTabData(t).Enabled And mTabData(t).Visible
                t = t - 1
                If t = -1 Then
                    If iAgain Then Exit Sub
                    t = mTabs - 1
                    iAgain = True
                End If
            Loop
            TabSel = t
        Else
            t = mTab + 1
            If t = mTabs Then t = 0
            Do Until mTabData(t).Enabled And mTabData(t).Visible
                t = t + 1
                If t = mTabs Then
                    If iAgain Then Exit Sub
                    t = 0
                    iAgain = True
                End If
            Loop
            TabSel = t
        End If
    End If
End Sub

Private Sub SetFocusToNextControlInSameContainer(nForward As Boolean)
    Dim iContainerUsr As Object
    Dim iContainerCtl As Object
    Dim iControls As Object
    Dim iHwnds() As Long
    Dim iTabIndexes() As Long
    Dim iCtl As Object
    Dim iTi As Long
    Dim iHwnd As Long
    Dim iEnabled As Boolean
    Dim iVisible As Boolean
    Dim iCount As Long
    Dim iUb As Long
    Dim iTiUsr As Long
    Dim c As Long
    
    On Error Resume Next
    Set iContainerUsr = UserControl.Extender.Container
    If iContainerUsr Is Nothing Then GoTo Exit_Sub
    
    Set iControls = UserControl.Parent.Controls
    If iControls Is Nothing Then GoTo Exit_Sub
    
    iTiUsr = -1
    iTiUsr = UserControl.Extender.TabIndex
    If iTiUsr = -1 Then GoTo Exit_Sub
    
    ReDim iHwnds(100)
    ReDim iTabIndexes(100)
    iUb = 100
    iCount = 0
    
    For Each iCtl In iControls
        Set iContainerCtl = Nothing
        Set iContainerCtl = iCtl.Container
        If iContainerCtl Is iContainerUsr Then
            iTi = -1
            iHwnd = 0
            iEnabled = False
            iVisible = False
            
            iTi = iCtl.TabIndex
            If iTi > -1 Then
                iHwnd = GetControlHwnd(iCtl)
                If iHwnd > 0 Then
                    iEnabled = iCtl.Enabled
                    iVisible = iCtl.Visible
                    If iEnabled And iVisible Then
                        iCount = iCount + 1
                        If (iCount - 1) > iUb Then
                            iUb = iUb + 100
                            ReDim Preserve iHwnds(iUb)
                            ReDim Preserve iTabIndexes(iUb)
                        End If
                        iHwnds(iCount - 1) = iHwnd
                        iTabIndexes(iCount - 1) = iTi
                    End If
                End If
            End If
        End If
    Next
    
    If iCount > 1 Then ' 1 means that the UserControl is the only control in the container, so there is no other control to focus
        ReDim Preserve iHwnds(iCount - 1)
        ReDim Preserve iTabIndexes(iCount - 1)
        
        ' Bubble sort
        Dim s As Long
        Dim iChanged As Boolean

        s = UBound(iTabIndexes)
        Do
            iChanged = False
            For c = 0 To s - 1
                If iTabIndexes(c) > iTabIndexes(c + 1) Then
                    iTi = iTabIndexes(c)
                    iHwnd = iHwnds(c)
                    iTabIndexes(c) = iTabIndexes(c + 1)
                    iHwnds(c) = iHwnds(c + 1)
                    iTabIndexes(c + 1) = iTi
                    iHwnds(c + 1) = iHwnd
                    iChanged = True
                End If
            Next c
            s = s - 1
        Loop While iChanged
        
        For c = 0 To UBound(iTabIndexes)
            If iTabIndexes(c) = iTiUsr Then
                If nForward Then
                    If c = UBound(iTabIndexes) Then
                        iHwnd = iHwnds(0)
                    Else
                        iHwnd = iHwnds(c + 1)
                    End If
                Else
                    If c = 0 Then
                        iHwnd = iHwnds(UBound(iTabIndexes))
                    Else
                        iHwnd = iHwnds(c - 1)
                    End If
                End If
                SetFocusAPI iHwnd
            End If
        Next c
    End If
    
Exit_Sub:
    Err.Clear
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    If mDraggingATab Then
        If KeyAscii = vbKeyEscape Then
            If mPreviousTabBeforeDragging <> mTab Then
                MoveTab mTab, mPreviousTabBeforeDragging
            End If
            DraggingATab = False
            Draw
        End If
    End If
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_LostFocus()
    mHasFocus = False
    PostDrawMessage
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim iX As Single
    Dim iY As Single
    Dim iXp As Single
    Dim iYp As Single
    Dim iTabIconClicked As Boolean
    Dim iForwardClickToTab As Boolean
    
    iX = X * mXCorrection
    iY = Y * mYCorrection
    
    RaiseEvent MouseDown(Button, Shift, iX, iY)
    
    If Button = 1 Then
        If mTabUnderMouse > -1 Then
            iXp = pScaleX(iX, vbTwips, vbPixels)
            iYp = pScaleX(iY, vbTwips, vbPixels)
            If (iXp + mIconClickExtendDPIScaled) >= mTabData(mTabUnderMouse).IconRect.Left Then
                If (iXp - mIconClickExtendDPIScaled) <= mTabData(mTabUnderMouse).IconRect.Right Then
                    If (iYp + mIconClickExtendDPIScaled) >= mTabData(mTabUnderMouse).IconRect.Top Then
                        If (iYp - mIconClickExtendDPIScaled) <= mTabData(mTabUnderMouse).IconRect.Bottom Then
                            If (mTDIMode <> ntTDIModeNone) And Not ((mTDIMode = ntTDIModeForms) And (mTabUnderMouse = FindTab(0, ntFindOriginalIndex))) Then
                                HandleTabTDIEvents
                            Else
                                iForwardClickToTab = True
                                RaiseEvent IconClick(mTabUnderMouse, iForwardClickToTab)
                            End If
                            iTabIconClicked = True
                        End If
                    End If
                End If
            End If
            If (Not iTabIconClicked) Or iForwardClickToTab Then
                If mTabData(mTabUnderMouse).Enabled Then
                    If mTab <> mTabUnderMouse Then
                        mHasFocus = True
                        mTabChangedFromAnotherRow = (mTabData(mTabUnderMouse).RowPos <> (mRows - 1))
                        mProcessingTabChange = True
                        TabSel = mTabUnderMouse
                        mProcessingTabChange = False
                    End If
                End If
            End If
        End If
        If CanReorderTabsEffective Then
            tmrCheckTabDrag.Enabled = False
            If Not (mTabChangedFromAnotherRow Or mProcessingTabChange Or (IIf(mTDIMode, mVisibleTabs < 3, mVisibleTabs < 2))) Then
                tmrCheckTabDrag.Enabled = True
                mMouseX = ScaleX(iX, vbTwips, vbPixels)
                mMouseY = ScaleY(iY, vbTwips, vbPixels)
            End If
        End If
    ElseIf (Button = vbMiddleButton) And (mTDIMode <> ntTDIModeNone) Then
        HandleTabTDIEvents
    ElseIf CanReorderTabsEffective Then
        tmrCheckTabDrag.Enabled = True
        mMouseX = 0
        mMouseY = 0
    End If
End Sub

Private Sub HandleTabTDIEvents()
    Dim iOpenAnother As Boolean
    Dim iIsLastTab As Boolean
    Dim iTabCaption As String
    Dim iTabNumber As Long
    Dim iCancel As Boolean
    Dim iLoadTabControls As Boolean
    Dim iUnloadTabControls As Boolean
    Const WM_SYSCOMMAND As Long = &H112
    Const SC_CLOSE As Long = &HF060&
    Const SW_HIDE As Long = 0
    Const WM_CLOSE As Long = &H10&
    Dim iHwnd As Long
    Dim iTabUnderMouse As Long
    Dim iShowTabCloseButton As Boolean
    Dim iDo As Boolean
    
    If mTabUnderMouse > -1 Then
        iDo = (mTabData(mTabUnderMouse).IconChar <> 0)
        If iDo And (mTDIMode = ntTDIModeForms) Then
            iDo = Not mTDIFormWithoutCloseButton(mTabData(mTabUnderMouse).Data)
        End If
        If iDo Then
            iDo = Not ((mTDIMode = ntTDIModeForms) And (mTabUnderMouse = FindTab(0, ntFindOriginalIndex)))
        End If
    End If
    If Not iDo Then Exit Sub
    
    iTabUnderMouse = mTabUnderMouse
    If mTabData(iTabUnderMouse).Data = -1 Then
        If mAmbientUserMode Then TDIAddNewTab
    Else
        iOpenAnother = (mTDIMode = ntTDIModeControls)
        iIsLastTab = mVisibleTabs = 2
        iTabNumber = mTabData(iTabUnderMouse).TDITabNumber
        iUnloadTabControls = True
        RaiseEvent TDIBeforeClosingTab(iTabNumber, iIsLastTab, iOpenAnother, iUnloadTabControls, iCancel)
        If iOpenAnother Then If (mTDIMode = ntTDIModeForms) Then iOpenAnother = False
        If mAmbientUserMode And (Not iCancel) Then
            If Not iIsLastTab Then
                iOpenAnother = False
            End If
            Redraw = False
            mTDIClosingATab = True
            TabVisible(iTabUnderMouse) = False
            mTDIClosingATab = False
            If mTDIMode = ntTDIModeControls Then
                If iUnloadTabControls Then
                    TDIUnloadTabControls iTabNumber
                End If
            ElseIf mTDIMode = ntTDIModeForms Then
                iHwnd = mTDIModeFormsFormData_FormHwnd(mTabData(iTabUnderMouse).Data)
                If IsWindow(iHwnd) <> 0 Then
                    ShowWindow iHwnd, SW_HIDE
                    SetParent iHwnd, mTDIModeFormsFormData_OldParentHwnd(mTabData(iTabUnderMouse).Data)
                    If mTDIFormWithoutCloseButton(mTabData(iTabUnderMouse).Data) Then
                        PostMessage iHwnd, WM_CLOSE, 0, 0
                    Else
                        SendMessage iHwnd, WM_SYSCOMMAND, SC_CLOSE, 0&
                    End If
                End If
                If IsWindow(iHwnd) <> 0 Then
                    iCancel = True
                    mTDIClosingATab = True
                    TabVisible(iTabUnderMouse) = True
                    mTDIClosingATab = False
                    If (mTabData(iTabUnderMouse).Data >= picTDIFormContainer.LBound) And (mTabData(iTabUnderMouse).Data <= picTDIFormContainer.UBound) Then
                        SetParent iHwnd, picTDIFormContainer(mTabData(iTabUnderMouse).Data).hWnd
                    End If
                End If
            End If
            If Not iCancel Then
                RaiseEvent TDITabClosed(iTabNumber, iIsLastTab)
                If iOpenAnother Then
                    mTDILastTabNumber = mTDILastTabNumber + 1
                    iTabCaption = "Default tab"
                    iLoadTabControls = True
                    iShowTabCloseButton = True
                    RaiseEvent TDIBeforeNewTab(ntLastTabClosed, mTDILastTabNumber, iTabCaption, iLoadTabControls, iShowTabCloseButton, False)
                    TDIPrepareNewTab iTabCaption, iLoadTabControls, , , iShowTabCloseButton
                End If
            End If
            Redraw = True
        End If
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim iX As Single
    Dim iY As Single
    Dim iXp As Single
    Dim iYp As Single
    Dim iCol As Long
    Dim iMouseOverIcon_Tab_Prev As Long
    Dim iBool As Boolean
    Dim i As Long
    Dim iTmpRC As RECT
    Static sXLast As Single
    Static sYLast As Single
    
    If mTDIAddingNewTab Then Exit Sub
    
    iX = X * mXCorrection
    iY = Y * mYCorrection
    
    ProcessMouseMove Button, Shift, iX, iY
    RaiseEvent MouseMove(Button, Shift, iX, iY)
    
    If mTabUnderMouse = mTab Then
        If IsEmpty(mTabData(mTabUnderMouse).CustomColors.IconColorSelectedTab) Then
            iCol = mIconColorSelectedTab
        Else
            iCol = mTabData(mTabUnderMouse).CustomColors.IconColorSelectedTab
        End If
        If IsEmpty(mTabData(mTabUnderMouse).CustomColors.IconColorMouseHoverSelectedTab) Then
            iBool = (mIconColorMouseHoverSelectedTab <> iCol)
        Else
            iBool = (mTabData(mTabUnderMouse).CustomColors.IconColorMouseHoverSelectedTab <> iCol)
        End If
    Else
        If IsEmpty(mTabData(mTabUnderMouse).CustomColors.IconColor) Then
            iCol = mIconColor
        Else
            iCol = mTabData(mTabUnderMouse).CustomColors.IconColor
        End If
        If IsEmpty(mTabData(mTabUnderMouse).CustomColors.IconColorMouseHover) Then
            iBool = (mIconColorMouseHover <> iCol)
        Else
            iBool = (mTabData(mTabUnderMouse).CustomColors.IconColorMouseHover <> iCol)
        End If
    End If
    If iBool Then
        Static sTabOverIcon_Last As Long
        Dim iImo As Boolean
        Dim iTum As Long
        
        If mTabUnderMouse > -1 Then
            iImo = mMouseIsOverIcon_Tab = mTabUnderMouse
            If mTabMousePointerHand Then
                If iImo Then
                    mCurrentMousePointerIsHand = True
                Else
                    mCurrentMousePointerIsHand = (mTabUnderMouse <> mTab)
                End If
                If mCurrentMousePointerIsHand Then
                    If GetCursor <> IDC_HAND Then
                        SetCursor mHandIconHandle
                    End If
                End If
            End If
            If iImo Then
                iTum = mTabUnderMouse
            Else
                iTum = -1
            End If
            If iTum <> sTabOverIcon_Last Then
                Draw
                sTabOverIcon_Last = iTum
            End If
        Else
            If mTabMousePointerHand Then mCurrentMousePointerIsHand = False
            sTabOverIcon_Last = -1
        End If
    End If

    If CanReorderTabsEffective Then
        If Not (mTabChangedFromAnotherRow Or mProcessingTabChange) Then
            If Not tmrCheckTabDrag.Enabled Then
                If (mMouseX <> 0) Or (mMouseY <> 0) Then
                    If Button = 1 Then
                        mMouseX2 = ScaleX(iX, vbTwips, vbPixels)
                        mMouseY2 = ScaleY(iY, vbTwips, vbPixels)
                        If tmrTabDragging.Enabled Then DrawDelayed
                    Else
                        mMouseX2 = 0
                        mMouseY2 = 0
                    End If
                    DraggingATab = (Not mChangingSelectedTab) And ((mMouseX <> 0 And mMouseX2 <> 0) Or (mMouseY <> 0 And mMouseY2 <> 0))
                    
                    If mRows = 1 Then
                        If DraggingATab Then
                            If (Abs(sXLast - iX) > (Screen.TwipsPerPixelX * (mTabData(mTab).TabRect.Right - mTabData(mTab).TabRect.Left) * 0.4)) Or (Abs(sYLast - iY) > (Screen.TwipsPerPixelY * (mTabData(mTab).TabRect.Bottom - mTabData(mTab).TabRect.Top) * 0.4)) Then
                                i = GetTabAtDropPoint
                                sXLast = iX
                                sYLast = iY
                                If i > -1 Then
                                    If mTDIMode = ntTDIModeControls Then
                                        If i = mTabs - 1 Then
                                            i = i - 1
                                        End If
                                    ElseIf mTDIMode = ntTDIModeForms Then
                                        If i = 0 Then i = 1
                                    End If
                                    If mTab <> i Then
                                        If (mTab > -1) And (i > -1) Then
                                            'mMouseX = mMouseX - mTabData(mTab).TabRect.Left + mTabData(i).TabRect.Left
                                            'mMouseY = mMouseY - mTabData(mTab).TabRect.Bottom + mTabData(i).TabRect.Bottom
                                            iTmpRC = mTabData(mTab).TabRect
                                            MoveTab mTab, i
                                            mMouseX = mMouseX - iTmpRC.Left + mTabData(mTab).TabRect.Left
                                            mMouseY = mMouseY - iTmpRC.Bottom + mTabData(mTab).TabRect.Bottom
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    If Not mMouseIsOverIcon Then
        If mTabUnderMouse > -1 Then
            If mTabData(mTabUnderMouse).IconChar <> 0 Then
                iXp = pScaleX(iX, vbTwips, vbPixels)
                iYp = pScaleX(iY, vbTwips, vbPixels)
                'Parent.Caption = iXp & "  " & mTabData(mTabUnderMouse).IconRect.Left & "  " & mTabData(mTabUnderMouse).IconRect.Right
                'Parent.Caption = iYp & "  " & mTabData(mTabUnderMouse).IconRect.Top & "  " & mTabData(mTabUnderMouse).IconRect.Bottom
                If (iXp + mIconClickExtendDPIScaled) >= mTabData(mTabUnderMouse).IconRect.Left Then
                    If (iXp - mIconClickExtendDPIScaled) <= mTabData(mTabUnderMouse).IconRect.Right Then
                        If (iYp + mIconClickExtendDPIScaled) >= mTabData(mTabUnderMouse).IconRect.Top Then
                            If (iYp - mIconClickExtendDPIScaled) <= mTabData(mTabUnderMouse).IconRect.Bottom Then
                                If Not ((mTDIMode = ntTDIModeForms) And (mTabUnderMouse = FindTab(0, ntFindOriginalIndex))) Then
                                    mMouseIsOverIcon = True
                                    mMouseIsOverIcon_Tab = mTabUnderMouse
                                    RaiseEvent IconMouseEnter(mTabUnderMouse)
                                    tmrHighlightIcon.Enabled = False
                                    tmrHighlightIcon.Enabled = True
                                    tmrPreHighlightIcon.Enabled = False
                                    tmrPreHighlightIcon.Enabled = True
                                    tmrHighlightIcon.Tag = mMouseIsOverIcon_Tab
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Else
        mMouseIsOverIcon = False
        iMouseOverIcon_Tab_Prev = mMouseIsOverIcon_Tab
        mMouseIsOverIcon_Tab = -1
        If mTabUnderMouse > -1 Then
            If mTabData(mTabUnderMouse).IconChar <> 0 Then
                iXp = pScaleX(iX, vbTwips, vbPixels)
                iYp = pScaleX(iY, vbTwips, vbPixels)
                If (iXp + mIconClickExtendDPIScaled) >= mTabData(mTabUnderMouse).IconRect.Left Then
                    If (iXp - mIconClickExtendDPIScaled) <= mTabData(mTabUnderMouse).IconRect.Right Then
                        If (iYp + mIconClickExtendDPIScaled) >= mTabData(mTabUnderMouse).IconRect.Top Then
                            If (iYp - mIconClickExtendDPIScaled) <= mTabData(mTabUnderMouse).IconRect.Bottom Then
                                mMouseIsOverIcon = True
                                mMouseIsOverIcon_Tab = mTabUnderMouse
                            End If
                        End If
                    End If
                End If
            End If
        End If
        If Not mMouseIsOverIcon Then
            RaiseEvent IconMouseLeave(iMouseOverIcon_Tab_Prev)
        End If
    End If
End Sub

Private Function GetTabAtXY(X As Single, Y As Single) As Long
    Dim t As Long
    Dim iX As Long
    Dim iY As Long
    
    iX = pScaleX(X, vbTwips, vbPixels)
    iY = pScaleX(Y, vbTwips, vbPixels)
    
    GetTabAtXY = mTab
    For t = 0 To mTabs - 1
        With mTabData(t).TabRect
            If iX >= .Left Then
                If iX <= .Right Then
                    If iY >= .Top Then
                        If iY <= .Bottom Then
                            GetTabAtXY = t
                            Exit For
                        End If
                    End If
                End If
            End If
        End With
    Next
End Function

Private Sub ProcessMouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim t As Long
    Dim iX As Long
    Dim iY As Long
    
    iX = pScaleX(X, vbTwips, vbPixels)
    iY = pScaleX(Y, vbTwips, vbPixels)
    If (Not mAmbientUserMode) And mRightToLeft Then iX = mScaleWidth - iX
    
    ' first check for the active tab, because in some cases it is bigger and can overlap surrounding tabs
    If (mTab > -1) And (mTab < mTabs) Then
        With mTabData(mTab).TabRect
            If iX >= .Left Then
                If iX <= .Right Then
                    If iY >= .Top Then
                        If iY <= .Bottom Then
                            If mTab <> mTabUnderMouse Then
                                If mTabUnderMouse > -1 Then
                                    RaiseEvent_TabMouseLeave (mTabUnderMouse)
                                End If
                                RaiseEvent_TabMouseEnter (mTab)
                                mTabUnderMouse = mTab
                                tmrTabMouseLeave.Enabled = False
                                tmrTabMouseLeave.Enabled = True
                            End If
                            Exit Sub
                        End If
                    End If
                End If
            End If
        End With
    End If
    
    For t = 0 To mTabs - 1
        If t <> mTab Then
            If mTabData(t).Visible And mTabData(t).Enabled Then
                With mTabData(t).TabRect
                    If iX >= .Left Then
                        If iX <= .Right Then
                            If iY >= .Top Then
                                If iY <= .Bottom Then
                                    If t <> mTabUnderMouse Then
                                        If mTabUnderMouse > -1 Then
                                            RaiseEvent_TabMouseLeave (mTabUnderMouse)
                                        End If
                                        RaiseEvent_TabMouseEnter (t)
                                        mTabUnderMouse = t
                                        tmrTabMouseLeave.Enabled = False
                                        tmrTabMouseLeave.Enabled = True
                                    End If
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                End With
            End If
        End If
    Next t
    If mTabUnderMouse > -1 Then
        tmrTabMouseLeave.Enabled = False
        RaiseEvent_TabMouseLeave (mTabUnderMouse)
    End If
    mTabUnderMouse = -1
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim iX As Single
    Dim iY As Single
    Dim i As Long
    
    iX = X * mXCorrection
    iY = Y * mYCorrection
    
    RaiseEvent MouseUp(Button, Shift, iX, iY)
    If mTabUnderMouse > -1 Then
        If Button = 2 Then
            RaiseEvent TabRightClick(mTabUnderMouse, Shift, iX, iY)
        End If
    End If
    
    If CanReorderTabsEffective Then
        If DraggingATab Then
            If mRows = 1 Then
                DraggingATab = False
                Draw
            Else
                If Not tmrCheckTabDrag.Enabled Then
                    tmrCheckTabDrag.Enabled = False
                    i = GetTabAtDropPoint
                    If i > -1 Then
                        If mTDIMode = ntTDIModeControls Then
                            If i = mTabs - 1 Then
                                i = i - 1
                            End If
                        ElseIf mTDIMode = ntTDIModeForms Then
                            If i = 0 Then i = 1
                        End If
                        If (mTab > -1) And (i > -1) Then
                            MoveTab mTab, i
                        End If
                    End If
                    DraggingATab = False
                    Draw
                End If
            End If
        End If
        tmrCheckTabDrag.Enabled = False
    End If
    mTabChangedFromAnotherRow = False
End Sub

Private Function GetTabAtDropPoint() As Long
    Dim c As Long
    Dim X As Single
    Dim Y  As Single
    
    X = mMouseX2 - mMouseX + mTabData(mTab).TabRect.Left + (mTabData(mTab).TabRect.Right - mTabData(mTab).TabRect.Left) / 2
    Y = mMouseY2 - mMouseY + mTabData(mTab).TabRect.Top + (mTabData(mTab).TabRect.Bottom - mTabData(mTab).TabRect.Top) / 2
    
    For c = 0 To mTabs - 1
        If mTabData(c).Visible Then
            If mTabData(c).TabRect.Top <= Y Then
                If mTabData(c).TabRect.Bottom >= Y Then
                    If IIf(mTabData(c).LeftTab, True, (mTabData(c).TabRect.Left + (mTabData(c).TabRect.Right - mTabData(c).TabRect.Left) / 2) <= X) Then
                        If IIf(mTabData(c).RightTab, True, (mTabData(c).TabRect.Right + (mTabData(c).TabRect.Right - mTabData(c).TabRect.Left) / 2) >= X) Then
                            GetTabAtDropPoint = c
                            If mTabData(c).LeftTab And (mMouseX2 - mMouseX < mTabData(c).TabRect.Left) Then
                                GetTabAtDropPoint = GetTabAtDropPoint - 1
                            End If
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Next
    GetTabAtDropPoint = -1
End Function

Private Sub UserControl_OLECompleteDrag(Effect As Long)
    RaiseEvent OLECompleteDrag(Effect)
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim t As Long
    
    UserControl.OLEDropMode = ssOLEDropManual
    tmrRestoreDropMode.Enabled = False
    
    If Not mOLEDropOnOtherTabs Then
        t = GetTabAtXY(X, Y)
        If t = mTab Then
            RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)
        End If
    Else
        RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)
    End If
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    Dim t As Long
    
    If Not mOLEDropOnOtherTabs Then
        t = GetTabAtXY(X, Y)
        If t <> mTab Then
            UserControl.OLEDropMode = ssOLEDropNone
            tmrRestoreDropMode.Enabled = True
        Else
            UserControl.OLEDropMode = ssOLEDropManual
            RaiseEvent OLEDragOver(Data, Effect, Button, Shift, X, Y, State)
        End If
    End If
End Sub

Private Sub UserControl_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub UserControl_OLESetData(Data As DataObject, DataFormat As Integer)
    RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub UserControl_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Dim c As Long
    Dim c2 As Long
    Dim iStr As String
    Dim iStr2 As String
    Dim iAllCtlNames As Collection
    Dim iLeftOffsetToHideWhenSaved As Long
    Dim iUpgradingFromSSTab As Boolean
    Dim iBytes() As Byte
    Dim iTheme As NewTabTheme
    Dim iControlVersion As Long
    
    On Error Resume Next
    mUserControlHwnd = UserControl.hWnd
    mAmbientUserMode = Ambient.UserMode
    mDefaultTabHeight = pScaleY(cPropDef_TabHeight, vbTwips, vbHimetric)
    If mDefaultTabHeight = 0 Then
        mDefaultTabHeight = 419.8055
    End If
    If mAmbientUserMode Then
        If TypeOf UserControl.Parent Is Form Then
            Set mForm = UserControl.Parent
        End If
    End If
    On Error GoTo 0
    
    mTabsEndFreeSpace = PropBag.ReadProperty("TabsEndFreeSpace", cPropDef_TabsEndFreeSpace)
    mSubclassingMethod = PropBag.ReadProperty("SubclassingMethod", cPropDef_SubclassingMethod)
    iLeftOffsetToHideWhenSaved = PropBag.ReadProperty("LeftOffsetToHideWhenSaved", 75000)
    If iLeftOffsetToHideWhenSaved <> mLeftOffsetToHide Then
        mPendingLeftOffset = iLeftOffsetToHideWhenSaved - mLeftOffsetToHide
    End If
    mTabs = PropBag.ReadProperty("Tabs", 3)
    mBackColor = PropBag.ReadProperty("BackColor", Ambient.BackColor)
    If mBackColor = Ambient.BackColor Then mBackColorIsFromAmbient = True
    mForeColor = PropBag.ReadProperty("ForeColor", Ambient.ForeColor)
    mForeColorSelectedTab = PropBag.ReadProperty("ForeColorSelectedTab", -1)
    If mForeColorSelectedTab = -1 Then
        mForeColorSelectedTab = PropBag.ReadProperty("ForeColorTabSel", mForeColor)
    End If
    mForeColorHighlighted = PropBag.ReadProperty("ForeColorHighlighted", mForeColor)
    mFlatTabBorderColorHighlight = PropBag.ReadProperty("FlatTabBorderColorHighlight", mForeColor)
    mFlatTabBorderColorSelectedTab = PropBag.ReadProperty("FlatTabBorderColorSelectedTab", -1)
    If mFlatTabBorderColorSelectedTab = -1 Then
        mFlatTabBorderColorSelectedTab = PropBag.ReadProperty("FlatTabBorderColorTabSel", mForeColor)
    End If
    If mForeColor = Ambient.ForeColor Then mForeColorIsFromAmbient = True
    Set mFont = PropBag.ReadProperty("Font", Ambient.Font)
    mEnabled = PropBag.ReadProperty("Enabled", True)
    mTabsPerRow = PropBag.ReadProperty("TabsPerRow", cPropDef_TabsPerRow)
    If mTabsPerRow < 1 Then mTabsPerRow = cPropDef_TabsPerRow
    mTab = PropBag.ReadProperty("Tab", 0)
    mTabOrientation = PropBag.ReadProperty("TabOrientation", ssTabOrientationTop)
    mShowFocusRect = PropBag.ReadProperty("ShowFocusRect", cPropDef_ShowFocusRect)
    mWordWrap = PropBag.ReadProperty("WordWrap", cPropDef_WordWrap)
    mStyle = PropBag.ReadProperty("Style", cPropDef_Style)
    mTabHeight = PropBag.ReadProperty("TabHeight", mDefaultTabHeight)    ' in Himetric, for compatibility with the original SSTab
    If pScaleY(mTabHeight, vbHimetric, vbPixels) < 1 Then mTabHeight = pScaleY(1, vbPixels, vbHimetric)
    mTabMaxWidth = PropBag.ReadProperty("TabMaxWidth", cPropDef_TabMaxWidth)  ' in Himetric, for compatibility with the original SSTab
    mTabMinWidth = PropBag.ReadProperty("TabMinWidth", cPropDef_TabMinWidth)  ' in Himetric
    mMousePointer = PropBag.ReadProperty("MousePointer", ssDefault)
    Set mMouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    mOLEDropMode = PropBag.ReadProperty("OLEDropMode", ssOLEDropNone)
    mMaskColor = PropBag.ReadProperty("MaskColor", cPropDef_MaskColor)
    mUseMaskColor = PropBag.ReadProperty("UseMaskColor", cPropDef_UseMaskColor)
    mHighlightTabExtraHeight = PropBag.ReadProperty("HighlightTabExtraHeight", Round(ScaleY(cPropDef_HighlightTabExtraHeight, vbTwips, vbHimetric)))
    If mHighlightTabExtraHeight < 0 Then mHighlightTabExtraHeight = 0
    mHighlightEffect = PropBag.ReadProperty("HighlightEffect", cPropDef_HighlightEffect)
    mShowRowsInPerspective = PropBag.ReadProperty("ShowRowsInPerspective", cPropDef_ShowRowsInPerspective)
    mTabWidthStyle = PropBag.ReadProperty("TabWidthStyle", cPropDef_TabWidthStyle)
    mBackColorTabs = PropBag.ReadProperty("BackColorTabs", Ambient.BackColor)
    If mBackColorTabs = Ambient.BackColor Then mBackColorTabsIsFromAmbient = True
    mBackColorSelectedTab = PropBag.ReadProperty("BackColorTabSel", -1)
    If mBackColorSelectedTab = -1 Then
        mBackColorSelectedTab = PropBag.ReadProperty("BackColorSelectedTab", -1)
    End If
    If mBackColorSelectedTab = -1 Then mBackColorSelectedTab_IsAutomatic = True
    If mBackColorSelectedTab_IsAutomatic Then mBackColorSelectedTab = GetAutomaticBackColorSelectedTab
    mFlatBarColorHighlight = PropBag.ReadProperty("FlatBarColorHighlight", -1)
    If mFlatBarColorHighlight = -1 Then mFlatBarColorHighlight_IsAutomatic = True: mFlatBarColorHighlight = 0
    mFlatBarColorInactive = PropBag.ReadProperty("FlatBarColorInactive", -1)
    If mFlatBarColorInactive = -1 Then mFlatBarColorInactive_IsAutomatic = True: mFlatBarColorInactive = 0
    mFlatTabsSeparationLineColor = PropBag.ReadProperty("FlatTabsSeparationLineColor", -1)
    If mFlatTabsSeparationLineColor = -1 Then mFlatTabsSeparationLineColor_IsAutomatic = True: mFlatTabsSeparationLineColor = 0
    mFlatBodySeparationLineColor = PropBag.ReadProperty("FlatBodySeparationLineColor", -1)
    If mFlatBodySeparationLineColor = -1 Then mFlatBodySeparationLineColor_IsAutomatic = True: mFlatBodySeparationLineColor = 0
    mFlatBorderColor = PropBag.ReadProperty("FlatBorderColor", -1)
    If mFlatBorderColor = -1 Then mFlatBorderColor_IsAutomatic = True: mFlatBorderColor = 0
    mHighlightColor = PropBag.ReadProperty("HighlightColor", -1)
    If mHighlightColor = -1 Then mHighlightColor_IsAutomatic = True: mHighlightColor = 0
    mHighlightColorSelectedTab = PropBag.ReadProperty("HighlightColorTabSel", -1)
    If mHighlightColorSelectedTab = -1 Then
        mHighlightColorSelectedTab = PropBag.ReadProperty("HighlightColorSelectedTab", -1)
    End If
    If mHighlightColorSelectedTab = -1 Then mHighlightColorSelectedTab_IsAutomatic = True: mHighlightColorSelectedTab = 0
    mSoftEdges = PropBag.ReadProperty("SoftEdges", cPropDef_SoftEdges)
    mChangeControlsBackColor = PropBag.ReadProperty("ChangeControlsBackColor", cPropDef_ChangeControlsBackColor)
    mChangeControlsForeColor = PropBag.ReadProperty("ChangeControlsForeColor", cPropDef_ChangeControlsForeColor)
    mTabTransition = PropBag.ReadProperty("TabTransition", cPropDef_TabTransition)
    mHighlightMode = PropBag.ReadProperty("HighlightMode", cPropDef_HighlightMode)
    If Not iUpgradingFromSSTab Then
        mHighlightModeSelectedTab = PropBag.ReadProperty("HighlightModeSelectedTab", -1)
        If mHighlightModeSelectedTab = -1 Then
            mHighlightModeSelectedTab = PropBag.ReadProperty("HighlightModeTabSel", cPropDef_HighlightModeSelectedTab)
        End If
    End If
    mCaptionAlignment = PropBag.ReadProperty("CaptionAlignment", cPropDef_CaptionAlignment)
    If mCaptionAlignment = ntAlignmentLeft Then
        mAlignmentDTFlag = DT_LEFT
    ElseIf mCaptionAlignment = ntAlignmentRight Then
        mAlignmentDTFlag = DT_RIGHT
    Else
        mAlignmentDTFlag = DT_CENTER
    End If
    If PropBag.ReadProperty("_Version", 0) <> 0 Then
        ' upgrading from SSTab
        If (mStyle <> ssStyleTabbedDialog) And (mStyle <> ssStylePropertyPage) Then
            mStyle = ssStyleTabbedDialog
        End If
        mBackColorTabs = vbButtonFace
        mBackColorTabsIsFromAmbient = False
        mBackColorSelectedTab = vbButtonFace
        mForeColor = vbButtonText
        mForeColorSelectedTab = vbButtonText
        mForeColorHighlighted = vbButtonText
        mFlatTabBorderColorHighlight = vbButtonText
        mFlatTabBorderColorSelectedTab = vbButtonText
        mSoftEdges = False
        mShowFocusRect = True
        mChangeControlsBackColor = False
        mChangeControlsForeColor = False
        mTabTransition = ntTransitionImmediate
        mHighlightEffect = False
        mHighlightMode = ntHLNone
        mHighlightModeSelectedTab = ntHLCaptionBold
        iUpgradingFromSSTab = True
    ElseIf PropBag.ReadProperty("Themed", False) Then
        ' upgrading from SSTab Ex
        mStyle = ntStyleWindows
        If mShowRowsInPerspective = ntYNAuto Then
            mShowRowsInPerspective = ntYes
        End If
        If mTabWidthStyle = ntTWAuto Then
            mTabWidthStyle = ntTWFixed
        End If
    End If
    mVisualStyles = (mStyle = ntStyleWindows) 'And IsThemed
    mShowDisabledState = PropBag.ReadProperty("ShowDisabledState", cPropDef_ShowDisabledState)
    mTabSeparation = PropBag.ReadProperty("TabSeparation", cPropDef_TabSeparation)
    mTabSeparationDPIScaled = mTabSeparation * mDPIScale
    mTabAppearance = PropBag.ReadProperty("TabAppearance", cPropDef_TabAppearance)
    mIconAlignment = PropBag.ReadProperty("IconAlignment", IIf(iUpgradingFromSSTab, ntIconAlignBeforeCaption, cPropDef_IconAlignment))
    mAutoRelocateControls = PropBag.ReadProperty("AutoRelocateControls", cPropDef_AutoRelocateControls)
    mRightToLeft = PropBag.ReadProperty("RightToLeft", Ambient.RightToLeft)
    If mRightToLeft Then
        SetLayout GetDC(picDraw.hWnd), LAYOUT_RTL
        SetLayout GetDC(picAux.hWnd), LAYOUT_RTL
    End If
    SetScaleMode
    mHandleHighContrastTheme = PropBag.ReadProperty("HandleHighContrastTheme", cPropDef_HandleHighContrastThem)
    mBackStyle = PropBag.ReadProperty("BackStyle", cPropDef_BackStyle)
    mAutoTabHeight = PropBag.ReadProperty("AutoTabHeight", False) ' Defaults to False for backward compatibility with SSTab
    mOLEDropOnOtherTabs = PropBag.ReadProperty("OLEDropOnOtherTabs", cPropDef_OLEDropOnOtherTabs)
    mFlatBarColorSelectedTab = PropBag.ReadProperty("FlatBarColorSelectedTab", -1)
    If mFlatBarColorSelectedTab = -1 Then
        mFlatBarColorSelectedTab = PropBag.ReadProperty("FlatBarColorTabSel", cPropDef_FlatBarColorSelectedTab)
    End If
    mFlatRoundnessTop = PropBag.ReadProperty("FlatRoundnessTop", cPropDef_FlatRoundnessTop)
    mFlatRoundnessTopDPIScaled = mFlatRoundnessTop * mDPIScale
    mFlatRoundnessBottom = PropBag.ReadProperty("FlatRoundnessBottom", cPropDef_FlatRoundnessBottom)
    mFlatRoundnessBottomDPIScaled = mFlatRoundnessBottom * mDPIScale
    mFlatRoundnessTabs = PropBag.ReadProperty("FlatRoundnessTabs", cPropDef_FlatRoundnessTabs)
    mFlatRoundnessTabsDPIScaled = mFlatRoundnessTabs * mDPIScale
    mTabMousePointerHand = PropBag.ReadProperty("TabMousePointerHand", cPropDef_TabMousePointerHand)
    mIconColor = PropBag.ReadProperty("IconColor", mForeColor)
    If mIconColor = Ambient.ForeColor Then mIconColorIsFromAmbient = True
    mIconColorSelectedTab = PropBag.ReadProperty("IconColorSelectedTab", -1)
    If mIconColorSelectedTab = -1 Then
        mIconColorSelectedTab = PropBag.ReadProperty("IconColorTabSel", mIconColor)
    End If
    mIconColorMouseHover = PropBag.ReadProperty("IconColorMouseHover", mIconColor)
    mIconColorMouseHoverSelectedTab = PropBag.ReadProperty("IconColorMouseHoverSelectedTab", -1)
    If mIconColorMouseHoverSelectedTab = -1 Then
        mIconColorMouseHoverSelectedTab = PropBag.ReadProperty("IconColorMouseHoverTabSel", mIconColor)
    End If
    mIconColorTabHighlighted = PropBag.ReadProperty("IconColorTabHighlighted", mIconColor)
    mFlatBorderMode = PropBag.ReadProperty("FlatBorderMode", cPropDef_FlatBorderMode)
    mFlatBarHeight = PropBag.ReadProperty("FlatBarHeight", cPropDef_FlatBarHeight)
    mFlatBarHeightDPIScaled = mFlatBarHeight * mDPIScale
    mFlatBarGripHeight = PropBag.ReadProperty("FlatBarGripHeight", cPropDef_FlatBarGripHeight)
    mFlatBarGripHeightDPIScaled = mFlatBarGripHeight * mDPIScale
    mFlatBarPosition = PropBag.ReadProperty("FlatBarPosition", cPropDef_FlatBarPosition)
    mCanReorderTabs = PropBag.ReadProperty("CanReorderTabs", cPropDef_CanReorderTabs)
    mTDIMode = PropBag.ReadProperty("TDIMode", cPropDef_TDIMode)
    If mTDIMode = -1 Then mTDIMode = ntTDIModeControls
    If (mTDIMode <> ntTDIModeNone) Then mTDIIconColorMouseHover = mIconColorMouseHover
    mFlatBodySeparationLineHeight = PropBag.ReadProperty("FlatBodySeparationLineHeight", cPropDef_FlatBodySeparationLineHeight)
    mFlatBodySeparationLineHeightDPIScaled = mFlatBodySeparationLineHeight * mDPIScale
    mUserControlSizeCorrectionsCounter_ScaleWidthNotToResize = PropBag.ReadProperty("UserControlSizeCorrectionsCounter_ScaleWidthNotToResize", 0)
    mUserControlSizeCorrectionsCounter_ScaleHeightNotToResize = PropBag.ReadProperty("UserControlSizeCorrectionsCounter_ScaleHeightNotToResize", 0)
    
    Set UserControl.MouseIcon = mMouseIcon
    UserControl.MousePointer = mMousePointer
    
    If mFont Is Nothing Then
        Set mFont = Ambient.Font
    End If
    If mFont Is Nothing Then
        Set mFont = UserControl.Font
    End If
    Set mDefaultIconFont = New StdFont
    mDefaultIconFont.Name = cPropDef_IconFontName
    mDefaultIconFont.Size = cPropDef_IconFontSize
    
    UserControl.Enabled = mEnabled Or (Not mAmbientUserMode)
    
    If mTabs = 0 Then
        ReDim mTabData(-1 To -1)
        mTab = -1
    Else
        ReDim mTabData(-1 To mTabs - 1)
        For c = 0 To mTabs - 1
            mTabData(c).OriginalIndex = c
        Next
    End If
    Set iAllCtlNames = New Collection
    mNoTabVisible = True
    iControlVersion = Val(PropBag.ReadProperty("ControlVersion", 0))
    For c = 0 To mTabs - 1
        Set mTabData(c).Controls = New Collection
        Set mTabData(c).Picture = PropBag.ReadProperty("TabPicture(" & CStr(c) & ")", Nothing)
        If Not mTabData(c).Picture Is Nothing Then
            If mTabData(c).Picture.Handle = 0 Then Set mTabData(c).Picture = Nothing
        End If
        Set mTabData(c).Pic16 = PropBag.ReadProperty("TabPic16(" & CStr(c) & ")", Nothing)
        If Not mTabData(c).Pic16 Is Nothing Then
            If mTabData(c).Pic16.Handle = 0 Then Set mTabData(c).Pic16 = Nothing
        End If
        Set mTabData(c).Pic20 = PropBag.ReadProperty("TabPic20(" & CStr(c) & ")", Nothing)
        If Not mTabData(c).Pic20 Is Nothing Then
            If mTabData(c).Pic20.Handle = 0 Then Set mTabData(c).Pic20 = Nothing
        End If
        Set mTabData(c).Pic24 = PropBag.ReadProperty("TabPic24(" & CStr(c) & ")", Nothing)
        If Not mTabData(c).Pic24 Is Nothing Then
            If mTabData(c).Pic24.Handle = 0 Then Set mTabData(c).Pic24 = Nothing
        End If
        mTabData(c).IconChar = PropBag.ReadProperty("TabIconChar(" & CStr(c) & ")", 0)
        mTabData(c).IconLeftOffset = PropBag.ReadProperty("TabIconLeftOffset(" & CStr(c) & ")", 0)
        mTabData(c).IconTopOffset = PropBag.ReadProperty("TabIconTopOffset(" & CStr(c) & ")", 0)
        If iControlVersion < 9 Then
            mTabData(c).Caption = PropBag.ReadProperty("TabCaption(" & CStr(c) & ")", "")
        Else
            mTabData(c).Caption = PropBag.ReadProperty("TabCaption(" & CStr(c) & ")", "Tab " & CStr(c))
        End If
        mTabData(c).ToolTipText = PropBag.ReadProperty("TabToolTipText(" & CStr(c) & ")", "")
        For c2 = 0 To PropBag.ReadProperty("Tab(" & c & ").ControlCount", 0) - 1
            iStr = PropBag.ReadProperty("Tab(" & c & ").Control(" & c2 & ")", "")
            If iStr <> "" Then
                iStr2 = ""
                On Error Resume Next
                iStr2 = iAllCtlNames(iStr)
                On Error GoTo 0
                If iStr2 = "" Then
                    mTabData(c).Controls.Add iStr, iStr
                    iAllCtlNames.Add iStr, iStr
                End If
            End If
        Next
        mTabData(c).Enabled = PropBag.ReadProperty("TabEnabled(" & CStr(c) & ")", True)
        mTabData(c).Visible = PropBag.ReadProperty("TabVisible(" & CStr(c) & ")", True)
        If mTabData(c).Visible Then mNoTabVisible = False
        Set mTabData(c).IconFont = PropBag.ReadProperty("IconFont(" & CStr(c) & ")", Nothing)
        mTabData(c).IconFontName = PropBag.ReadProperty("IconFontName(" & CStr(c) & ")", "")
        If mTabData(c).IconFontName <> "" Then
            If mTabData(c).IconFontName <> mTabData(c).IconFont.Name Then
                mTabData(c).DoNotUseIconFont = True
            End If
        End If
        If Not mTabData(c).IconFont Is Nothing Then mTabIconFontsEventsHandler.AddFont mTabData(c).IconFont, c
        mTabData(c).FixedWidth = PropBag.ReadProperty("TabFixedWidth(" & CStr(c) & ")", 0)
        mTabData(c).Tag = PropBag.ReadProperty("TabTag(" & CStr(c) & ")", "")
        mTabData(c).Data = PropBag.ReadProperty("TabData(" & CStr(c) & ")", 0)
        mTabData(c).CustomColors.Deserialize PropBag.ReadProperty("TabCustomColors(" & CStr(c) & ")", "")
        mTabData(c).Key = PropBag.ReadProperty("TabKey(" & CStr(c) & ")", "")
    Next c
    mTabData(mTab).Selected = True
    
    c2 = 1
    iBytes = PropBag.ReadProperty("Theme(" & CStr(c2) & ")", "")
    Do Until UBound(iBytes) = -1
        If mThemesCollection Is Nothing Then Set mThemesCollection = New NewTabThemes
        Set iTheme = New NewTabTheme
        iTheme.Deserialize iBytes
        mThemesCollection.Add iTheme
        c2 = c2 + 1
        iBytes = PropBag.ReadProperty("Theme(" & CStr(c2) & ")", "")
    Loop
    
    If mHandleHighContrastTheme Then CheckHighContrastTheme
    SetFont
    If mTabAppearance <> ntTAAuto Then
        mAppearanceIsFlat = (mTabAppearance = ntTAFlat)
    Else
        mAppearanceIsFlat = mStyle = ntStyleFlat
    End If
    SetHighlightMode
    mSetAutoTabHeightPending = True
    SetButtonFaceColor
    SetColors
    UserControl.BackColor = mBackColor
    CheckIfThereAreTabsToolTipTexts
    UserControl.OLEDropMode = mOLEDropMode
    
    mControlJustAdded = PropBag.ReadProperty("ControlJustAdded", True)
    
    mSubclassed = mSubclassingMethod <> ntSMDisabled
#If NOSUBCLASSINIDE Then
    If mInIDE Then
        mSubclassed = False
    End If
#End If
    
    If mSubclassed Then
        gSubclassWithSetWindowLong = (mSubclassingMethod = ntSMSetWindowLong) Or (mSubclassingMethod = ntSM_SWLOnlyUserControl)
        mOnlySubclassUserControl = (mSubclassingMethod = ntSM_SWSOnlyUserControl) Or (mSubclassingMethod = ntSM_SWLOnlyUserControl)
        SubclassUserControl
    Else
        mFormIsActive = True
    End If
    mPropertiesReady = True
    If mTDIMode = ntTDIModeForms Then ReDim mTDIFormWithoutCloseButton(0)
    
    PostDrawMessage
    If tmrDraw.Enabled Then
        Draw
    End If
    
    If mAmbientUserMode Then
        mHandIconHandle = LoadCursor(ByVal 0&, IDC_HAND)
    End If
    
    If mAmbientUserMode And (mTDIMode = ntTDIModeForms) Then
        If mSettingTDIMode Then Exit Sub
        If Not mTDIModeAlreadySet Then
            mTDIModeAlreadySet = True
            TDISetTDIMode
        End If
    End If
End Sub

Private Sub UserControl_Resize()
    ResetCachedThemeImages
    If mAmbientUserMode Then
        PostDrawMessage
    Else
        tmrDraw.Enabled = True
    End If
    If mResizeEventPending Then RaiseEvent_Resize
End Sub

Private Sub UserControl_Show()
    If mUserControlTerminated Then Exit Sub
    If mResizeEventPending Then RaiseEvent_Resize
    If mAmbientUserMode And (mTDIMode <> ntTDIModeNone) Then
        If mSettingTDIMode Then Exit Sub
        If IsWindowVisible(mUserControlHwnd) <> 0 Then
            If Not mTDIModeAlreadySet Then
                mTDIModeAlreadySet = True
                TDISetTDIMode
            End If
        End If
    End If
    
    If mUserControlShown Then
        Exit Sub
    End If
    
    If mTDIModeFormsUnhooked Then
        If mTDIMode = ntTDIModeForms Then
            If mAmbientUserMode Then
                InstallCBTHook Me
            End If
        End If
        mTDIModeFormsUnhooked = False
    End If
    
    If mPendingLeftOffset <> 0 Then
        DoPendingLeftOffset
    End If
    
    If mAmbientUserMode And mSubclassed Then
        If (mFormHwnd = 0) Then
            mFormHwnd = GetAncestor(UserControl.ContainerHwnd, GA_ROOT)
            mFormIsActive = GetForegroundWindow = mFormHwnd
            SubclassForm
        End If
        
        Dim iAuxLeft As Long
        Dim iHwnd As Long
        Dim c As Long
        Dim iCtlName As String
        Dim iCtl As Object
        Dim iIsLine As Boolean
        
        On Error Resume Next
        If mSubclassedControlsForMoveHwnds.Count > 0 Then
            For c = 1 To mSubclassedControlsForMoveHwnds.Count
                iHwnd = mSubclassedControlsForMoveHwnds(c)
                DetachMessage Me, iHwnd, WM_WINDOWPOSCHANGING
            Next c
            Set mSubclassedControlsForMoveHwnds = New Collection
        End If
    
        For Each iCtl In UserControlContainedControls
            iAuxLeft = 0
            iIsLine = False
            If TypeName(iCtl) = "Line" Then
                iAuxLeft = iCtl.X1
                iIsLine = True
            Else
                iAuxLeft = iCtl.Left
            End If
            If iAuxLeft >= -mLeftThresholdHided Then
                iCtlName = ControlName(iCtl)
                If Not ControlIsInTab(iCtlName, mTab) Then
                    If iIsLine Then
                        iCtl.X1 = iCtl.X1 - mLeftOffsetToHide
                        iCtl.X2 = iCtl.X2 - mLeftOffsetToHide
                    Else
                        iCtl.Left = iCtl.Left - mLeftOffsetToHide
                    End If
                    iAuxLeft = iAuxLeft - mLeftOffsetToHide
                End If
            End If
            If Not mOnlySubclassUserControl Then
                If iAuxLeft < -mLeftThresholdHided Then
                    iHwnd = 0
                    iHwnd = GetControlHwnd(iCtl)
                    If iHwnd <> 0 Then
                        mSubclassedControlsForMoveHwnds.Add iHwnd
                        AttachMessage Me, iHwnd, WM_WINDOWPOSCHANGING
                    End If
                End If
            End If
        Next
        On Error GoTo 0
    End If
    
    If mChangeControlsBackColor Then
        If Not mChangedControlsBackColor Then
            SetControlsBackColor mBackColorSelectedTab
            mChangedControlsBackColor = True
        End If
    End If
    If mChangeControlsForeColor Then
        If Not mChangedControlsForeColor Then
            SetControlsProperForeColor
            mChangedControlsForeColor = True
        End If
    End If
    
    If mAmbientUserMode Then
        If Not mTabStopsInitialized Then
            StoreControlsTabStop True
            mTabStopsInitialized = True
        End If
        If mForm Is Nothing Then SubclassControlsPainting
    Else
        HideAllContainedControls
        MakeContainedControlsInSelTabVisible
        If Not IsMsgBoxShown Then CheckContainedControlsConsistency
    End If
    mUserControlShown = True
    SubclassControlsPainting
    If (Not mFirstDraw) Or mDrawMessagePosted Then
        Draw
        mFirstDraw = True
    End If
    RaiseEvent TabChange
End Sub

Private Sub DoPendingLeftOffset()
    Dim iCtl As Object
    Dim iIsLine As Boolean
    Dim iAuxLeft As Long
    
    If mPendingLeftOffset <> 0 Then
        For Each iCtl In UserControlContainedControls
            iAuxLeft = 0
            iIsLine = False
            On Error Resume Next
            If TypeName(iCtl) = "Line" Then
                iAuxLeft = iCtl.X1
                iIsLine = True
            Else
                iAuxLeft = iCtl.Left
            End If
            On Error GoTo 0
            If iAuxLeft < -mLeftThresholdHided Then
                If iIsLine Then
                    iCtl.X1 = iCtl.X1 + mPendingLeftOffset
                    iCtl.X2 = iCtl.X2 + mPendingLeftOffset
                Else
                    iCtl.Left = iCtl.Left + mPendingLeftOffset
                End If
            End If
        Next
        mPendingLeftOffset = 0
    End If

End Sub

Private Function ControlIsInTab(nCtlName As String, ByVal nTab As Long) As Boolean
    Dim c As Long
    
    For c = 1 To mTabData(nTab).Controls.Count
        If mTabData(nTab).Controls(c) = nCtlName Then
            ControlIsInTab = True
            Exit Function
        End If
    Next c
End Function
    
Private Sub UserControl_Terminate()
    DoTerminate
End Sub

Private Sub DoTerminate()
    If mUserControlTerminated Then Exit Sub
    mUserControlTerminated = True
    
    If mTDIMode = ntTDIModeForms Then
        If mAmbientUserMode Then
            UninstallCBTHook
            mTDIModeFormsUnhooked = False
        End If
    End If
    
    tmrShowTabTTT.Enabled = False
    Set mToolTipEx = Nothing
    
    tmrTabMouseLeave.Enabled = False
    tmrDraw.Enabled = False
    tmrCancelDoubleClick.Enabled = False
    tmrCheckContainedControlsAdditionDesignTime.Enabled = False
    tmrHighlightEffect.Enabled = False
    
    Set mParentControlsTabStop = Nothing
    Set mParentControlsUseMnemonic = Nothing
    Set mContainedControlsThatAreContainers = Nothing
    
    Unsubclass
    
    If Not mTabIconFontsEventsHandler Is Nothing Then
        mTabIconFontsEventsHandler.Release
        Set mTabIconFontsEventsHandler = Nothing
    End If
    If mHandIconHandle <> 0 Then
        DestroyCursor mHandIconHandle
        mHandIconHandle = 0
    End If
End Sub

Private Sub Unsubclass()
    Dim c As Long
    Dim iHwnd As Long
    
    If mSubclassed Then
        If (mFormHwnd <> 0) And mAmbientUserMode Then
            On Error Resume Next
            DetachMessage Me, mFormHwnd, WM_SYSCOLORCHANGE
            DetachMessage Me, mFormHwnd, WM_THEMECHANGED
            DetachMessage Me, mFormHwnd, WM_NCACTIVATE
            DetachMessage Me, mFormHwnd, WM_GETDPISCALEDSIZE
            If mTDIMode = ntTDIModeForms Then
                DetachMessage Me, mFormHwnd, WM_ACTIVATE
                DetachMessage Me, mFormHwnd, WM_ACTIVATEAPP
            End If
            On Error GoTo 0
        End If
        If mSubclassed Then
            If mAmbientUserMode Then
                On Error Resume Next
                DetachMessage Me, mUserControlHwnd, WM_MOUSEACTIVATE
                DetachMessage Me, mUserControlHwnd, WM_SETFOCUS
                DetachMessage Me, mUserControlHwnd, WM_DRAW
                DetachMessage Me, mUserControlHwnd, WM_INIT
                DetachMessage Me, mUserControlHwnd, WM_SETCURSOR
                On Error GoTo 0
                mCanPostDrawMessage = False
            Else
                On Error Resume Next
                DetachMessage Me, mUserControlHwnd, WM_LBUTTONDOWN
                DetachMessage Me, mUserControlHwnd, WM_LBUTTONUP
                DetachMessage Me, mUserControlHwnd, WM_LBUTTONDBLCLK
                DetachMessage Me, mUserControlHwnd, WM_MOUSEMOVE
                On Error GoTo 0
            End If
        End If
        mSubclassed = False
    End If

    If Not mSubclassedControlsForPaintingHwnds Is Nothing Then
        For c = 1 To mSubclassedControlsForPaintingHwnds.Count
            iHwnd = mSubclassedControlsForPaintingHwnds(c)
            On Error Resume Next
            DetachMessage Me, iHwnd, WM_PAINT
            DetachMessage Me, iHwnd, WM_MOVE
            On Error GoTo 0
        Next c
        Set mSubclassedControlsForPaintingHwnds = Nothing
    End If
    
    If Not mSubclassedFramesHwnds Is Nothing Then
        For c = 1 To mSubclassedFramesHwnds.Count
            iHwnd = mSubclassedFramesHwnds(c)
            On Error Resume Next
            DetachMessage Me, iHwnd, WM_PRINTCLIENT
            DetachMessage Me, iHwnd, WM_MOUSELEAVE
            On Error GoTo 0
        Next c
        Set mSubclassedFramesHwnds = Nothing
    End If
    
    If Not mSubclassedControlsForMoveHwnds Is Nothing Then
        For c = 1 To mSubclassedControlsForMoveHwnds.Count
            iHwnd = mSubclassedControlsForMoveHwnds(c)
            On Error Resume Next
            DetachMessage Me, iHwnd, WM_WINDOWPOSCHANGING
            On Error GoTo 0
        Next c
        Set mSubclassedControlsForMoveHwnds = Nothing
    End If
    
    If Not mTDISubclassedFormsHwnds Is Nothing Then
        For c = 1 To mTDISubclassedFormsHwnds.Count
            iHwnd = mTDISubclassedFormsHwnds(c)
            On Error Resume Next
            DetachMessage Me, iHwnd, WM_STYLECHANGING
            DetachMessage Me, iHwnd, WM_DESTROY
            DetachMessage Me, iHwnd, WM_SHOWWINDOW
            On Error GoTo 0
        Next
        Set mTDISubclassedFormsHwnds = Nothing
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Dim c As Long
    Dim c2 As Long
    Dim iTheme As NewTabTheme
    Dim iBytes() As Byte
    
    StoreVisibleControlsInSelectedTab
    
    PropBag.WriteProperty "ControlJustAdded", mControlJustAdded, True
    
    PropBag.WriteProperty "TabsEndFreeSpace", mTabsEndFreeSpace, cPropDef_TabsEndFreeSpace
    PropBag.WriteProperty "SubclassingMethod", mSubclassingMethod, cPropDef_SubclassingMethod
    PropBag.WriteProperty "Tabs", mTabs, 3
    PropBag.WriteProperty "BackColor", mBackColor, Ambient.BackColor
    PropBag.WriteProperty "ForeColor", mForeColor, Ambient.ForeColor
    PropBag.WriteProperty "ForeColorSelectedTab", mForeColorSelectedTab, mForeColor
    PropBag.WriteProperty "ForeColorHighlighted", mForeColorHighlighted, mForeColor
    PropBag.WriteProperty "FlatTabBorderColorHighlight", mFlatTabBorderColorHighlight, mForeColor
    PropBag.WriteProperty "FlatTabBorderColorSelectedTab", mFlatTabBorderColorSelectedTab, mForeColor
    PropBag.WriteProperty "Font", mFont, Ambient.Font
    PropBag.WriteProperty "Enabled", mEnabled, True
    PropBag.WriteProperty "TabsPerRow", mTabsPerRow, cPropDef_TabsPerRow
    PropBag.WriteProperty "Tab", mTab, 0
    PropBag.WriteProperty "TabOrientation", mTabOrientation, ssTabOrientationTop
    PropBag.WriteProperty "ShowFocusRect", mShowFocusRect, cPropDef_ShowFocusRect
    PropBag.WriteProperty "WordWrap", mWordWrap, cPropDef_WordWrap
    PropBag.WriteProperty "Style", mStyle, cPropDef_Style
    PropBag.WriteProperty "TabHeight", Round(mTabHeight), Round(mDefaultTabHeight)  ' in Himetric, for compatibility with the original SSTab
    PropBag.WriteProperty "TabMaxWidth", Round(mTabMaxWidth), cPropDef_TabMaxWidth  ' in Himetric, for compatibility with the original SSTab
    PropBag.WriteProperty "TabMinWidth", Round(mTabMinWidth), cPropDef_TabMinWidth ' in Himetric
    PropBag.WriteProperty "MousePointer", mMousePointer, ssDefault
    PropBag.WriteProperty "MouseIcon", mMouseIcon, Nothing
    PropBag.WriteProperty "OLEDropMode", mOLEDropMode, ssOLEDropNone
    PropBag.WriteProperty "MaskColor", mMaskColor, cPropDef_MaskColor
    PropBag.WriteProperty "UseMaskColor", mUseMaskColor, cPropDef_UseMaskColor
    PropBag.WriteProperty "HighlightTabExtraHeight", Round(mHighlightTabExtraHeight), Round(ScaleY(cPropDef_HighlightTabExtraHeight, vbTwips, vbHimetric))
    PropBag.WriteProperty "HighlightEffect", mHighlightEffect, cPropDef_HighlightEffect
'    PropBag.WriteProperty "Themed", False
    PropBag.WriteProperty "BackColorTabs", mBackColorTabs, Ambient.BackColor
    If mBackColorSelectedTab_IsAutomatic Then
        PropBag.WriteProperty "BackColorSelectedTab", 0, 0 ' delete any value already saved
    Else
        PropBag.WriteProperty "BackColorSelectedTab", mBackColorSelectedTab, -1
    End If
    If mFlatBarColorHighlight_IsAutomatic Then
        PropBag.WriteProperty "FlatBarColorHighlight", 0, 0 ' delete any value already saved
    Else
        PropBag.WriteProperty "FlatBarColorHighlight", mFlatBarColorHighlight, -1
    End If
    If mFlatBarColorInactive_IsAutomatic Then
        PropBag.WriteProperty "FlatBarColorInactive", 0, 0 ' delete any value already saved
    Else
        PropBag.WriteProperty "FlatBarColorInactive", mFlatBarColorInactive, -1
    End If
    If mFlatTabsSeparationLineColor_IsAutomatic Then
        PropBag.WriteProperty "FlatTabsSeparationLineColor", 0, 0 ' delete any value already saved
    Else
        PropBag.WriteProperty "FlatTabsSeparationLineColor", mFlatTabsSeparationLineColor, -1
    End If
    If mFlatBodySeparationLineColor_IsAutomatic Then
        PropBag.WriteProperty "FlatBodySeparationLineColor", 0, 0 ' delete any value already saved
    Else
        PropBag.WriteProperty "FlatBodySeparationLineColor", mFlatBodySeparationLineColor, -1
    End If
    If mFlatBorderColor_IsAutomatic Then
        PropBag.WriteProperty "FlatBorderColor", 0, 0 ' delete any value already saved
    Else
        PropBag.WriteProperty "FlatBorderColor", mFlatBorderColor, -1
    End If
    If mHighlightColor_IsAutomatic Then
        PropBag.WriteProperty "HighlightColor", 0, 0 ' delete any value already saved
    Else
        PropBag.WriteProperty "HighlightColor", mHighlightColor, -1
    End If
    If mHighlightColorSelectedTab_IsAutomatic Then
        PropBag.WriteProperty "HighlightColorSelectedTab", 0, 0 ' delete any value already saved
    Else
        PropBag.WriteProperty "HighlightColorSelectedTab", mHighlightColorSelectedTab, -1
    End If
    PropBag.WriteProperty "ShowDisabledState", mShowDisabledState, cPropDef_ShowDisabledState
    PropBag.WriteProperty "ChangeControlsBackColor", mChangeControlsBackColor, cPropDef_ChangeControlsBackColor
    PropBag.WriteProperty "ChangeControlsForeColor", mChangeControlsForeColor, cPropDef_ChangeControlsForeColor
    PropBag.WriteProperty "TabWidthStyle", mTabWidthStyle, cPropDef_TabWidthStyle
    PropBag.WriteProperty "ShowRowsInPerspective", mShowRowsInPerspective, cPropDef_ShowRowsInPerspective
    PropBag.WriteProperty "TabSeparation", mTabSeparation, cPropDef_TabSeparation
    PropBag.WriteProperty "TabAppearance", mTabAppearance, cPropDef_TabAppearance
    PropBag.WriteProperty "IconAlignment", mIconAlignment, cPropDef_IconAlignment
    PropBag.WriteProperty "AutoRelocateControls", mAutoRelocateControls, cPropDef_AutoRelocateControls
    PropBag.WriteProperty "SoftEdges", mSoftEdges, cPropDef_SoftEdges
    PropBag.WriteProperty "RightToLeft", mRightToLeft, Ambient.RightToLeft
    PropBag.WriteProperty "HandleHighContrastTheme", mHandleHighContrastTheme, cPropDef_HandleHighContrastThem
    PropBag.WriteProperty "LeftOffsetToHideWhenSaved", mLeftOffsetToHide + mPendingLeftOffset, 75000
    PropBag.WriteProperty "LeftThresholdHidedWhenSaved", mLeftThresholdHided, 15000
    PropBag.WriteProperty "BackStyle", mBackStyle, cPropDef_BackStyle
    PropBag.WriteProperty "AutoTabHeight", mAutoTabHeight, False  ' Defaults to False for backward compatibility with SSTab
    PropBag.WriteProperty "OLEDropOnOtherTabs", mOLEDropOnOtherTabs, cPropDef_OLEDropOnOtherTabs
    PropBag.WriteProperty "FlatBarColorSelectedTab", mFlatBarColorSelectedTab, cPropDef_FlatBarColorSelectedTab
    PropBag.WriteProperty "TabTransition", mTabTransition, cPropDef_TabTransition
    PropBag.WriteProperty "FlatRoundnessTop", mFlatRoundnessTop, cPropDef_FlatRoundnessTop
    PropBag.WriteProperty "FlatRoundnessBottom", mFlatRoundnessBottom, cPropDef_FlatRoundnessBottom
    PropBag.WriteProperty "FlatRoundnessTabs", mFlatRoundnessTabs, cPropDef_FlatRoundnessTabs
    PropBag.WriteProperty "TabMousePointerHand", mTabMousePointerHand, cPropDef_TabMousePointerHand
    PropBag.WriteProperty "IconColor", mIconColor, mForeColor
    PropBag.WriteProperty "IconColorSelectedTab", mIconColorSelectedTab, mIconColor
    PropBag.WriteProperty "IconColorMouseHover", IIf(mTDIMode <> ntTDIModeNone, mTDIIconColorMouseHover, mIconColorMouseHover), mIconColor
    PropBag.WriteProperty "IconColorMouseHoverSelectedTab", mIconColorMouseHoverSelectedTab, mIconColor
    PropBag.WriteProperty "IconColorTabHighlighted", mIconColorTabHighlighted, mIconColor
    PropBag.WriteProperty "HighlightMode", mHighlightMode, cPropDef_HighlightMode
    PropBag.WriteProperty "HighlightModeSelectedTab", mHighlightModeSelectedTab, cPropDef_HighlightModeSelectedTab
    PropBag.WriteProperty "CaptionAlignment", mCaptionAlignment, cPropDef_CaptionAlignment
    PropBag.WriteProperty "FlatBorderMode", mFlatBorderMode, cPropDef_FlatBorderMode
    PropBag.WriteProperty "FlatBarHeight", mFlatBarHeight, cPropDef_FlatBarHeight
    PropBag.WriteProperty "FlatBarGripHeight", mFlatBarGripHeight, cPropDef_FlatBarGripHeight
    PropBag.WriteProperty "FlatBarPosition", mFlatBarPosition, cPropDef_FlatBarPosition
    PropBag.WriteProperty "CanReorderTabs", mCanReorderTabs, cPropDef_CanReorderTabs
    PropBag.WriteProperty "TDIMode", mTDIMode, cPropDef_TDIMode
    PropBag.WriteProperty "FlatBodySeparationLineHeight", mFlatBodySeparationLineHeight, cPropDef_FlatBodySeparationLineHeight
    PropBag.WriteProperty "UserControlSizeCorrectionsCounter_ScaleWidthNotToResize", mUserControlSizeCorrectionsCounter_ScaleWidthNotToResize, 0
    PropBag.WriteProperty "UserControlSizeCorrectionsCounter_ScaleHeightNotToResize", mUserControlSizeCorrectionsCounter_ScaleHeightNotToResize, 0
    PropBag.WriteProperty "ControlVersion", App.Major, 0
    
    For c = 0 To mTabs - 1
        PropBag.WriteProperty "TabPicture(" & CStr(c) & ")", mTabData(c).Picture, Nothing
        PropBag.WriteProperty "TabPic16(" & CStr(c) & ")", mTabData(c).Pic16, Nothing
        PropBag.WriteProperty "TabPic20(" & CStr(c) & ")", mTabData(c).Pic20, Nothing
        PropBag.WriteProperty "TabPic24(" & CStr(c) & ")", mTabData(c).Pic24, Nothing
        PropBag.WriteProperty "TabIconChar(" & CStr(c) & ")", mTabData(c).IconChar, 0
        PropBag.WriteProperty "TabIconLeftOffset(" & CStr(c) & ")", mTabData(c).IconLeftOffset, 0
        PropBag.WriteProperty "TabIconTopOffset(" & CStr(c) & ")", mTabData(c).IconTopOffset, 0
        PropBag.WriteProperty "TabCaption(" & CStr(c) & ")", mTabData(c).Caption, "Tab " & CStr(c)
        PropBag.WriteProperty "TabToolTipText(" & CStr(c) & ")", mTabData(c).ToolTipText, ""
        PropBag.WriteProperty "TabEnabled(" & CStr(c) & ")", mTabData(c).Enabled, True
        PropBag.WriteProperty "TabVisible(" & CStr(c) & ")", mTabData(c).Visible, True
        PropBag.WriteProperty "Tab(" & c & ").ControlCount", mTabData(c).Controls.Count, 0
        For c2 = 1 To mTabData(c).Controls.Count
            PropBag.WriteProperty "Tab(" & c & ").Control(" & c2 - 1 & ")", mTabData(c).Controls(c2), ""
        Next
    
        If FontsAreEqual(mTabData(c).IconFont, mDefaultIconFont) Then
            PropBag.WriteProperty "IconFont(" & CStr(c) & ")", Nothing, Nothing
            PropBag.WriteProperty "IconFontName(" & CStr(c) & ")", "", ""
        Else
            PropBag.WriteProperty "IconFont(" & CStr(c) & ")", mTabData(c).IconFont, Nothing
            PropBag.WriteProperty "IconFontName(" & CStr(c) & ")", mTabData(c).IconFontName, ""
        End If
        PropBag.WriteProperty "TabFixedWidth(" & CStr(c) & ")", mTabData(c).FixedWidth, 0
        PropBag.WriteProperty "TabTag(" & CStr(c) & ")", mTabData(c).Tag, ""
        PropBag.WriteProperty "TabData(" & CStr(c) & ")", mTabData(c).Data, 0
        iBytes = mTabData(c).CustomColors.Serialize
        If UBound(iBytes) = -1 Then
            PropBag.WriteProperty "TabCustomColors(" & CStr(c) & ")", "", ""
        Else
            PropBag.WriteProperty "TabCustomColors(" & CStr(c) & ")", iBytes
        End If
        PropBag.WriteProperty "TabKey(" & CStr(c) & ")", mTabData(c).Key, ""
    Next c
    
    If Not mThemesCollection Is Nothing Then
        If mThemesCollection.ThereAreCustomThemes Then
            c2 = 0
            For c = 1 To mThemesCollection.Count
                Set iTheme = mThemesCollection(c)
                If iTheme.Custom Then
                    c2 = c2 + 1
                    PropBag.WriteProperty "Theme(" & CStr(c2) & ")", iTheme.Serialize
                End If
            Next
        End If
    End If
End Sub

Private Sub Draw()
    Dim iTabWidth As Single
    Dim iTabData As T_TabData
    Dim iTabExtraHeight As Long
    Dim iLng As Long
    Dim t As Long
    Dim ctv As Long
    Dim iPosH As Long
    Dim iRow As Long ' this variable is reused and not always means the same thing
    Dim iRowPerspectiveSpace As Long
    Dim iAllRowsPerspectiveSpace As Long
    Dim iTabHeight As Long
    Dim iTmpRect As RECT
    Dim iLastVisibleTab As Long
    Dim iLastVisibleTab_Prev As Long
    Dim iScaleWidth As Long
    Dim iScaleHeight As Long
    Dim iTabMaxWidth As Long
    Dim iTabMinWidth As Long
    Dim iTabLeft As Long
    Dim iRowTabCount As Long
    Dim iAccumulatedTabWith As Long
    Dim iTotalTabWidth As Long
    Dim iTabStretchRatio As Single
    Dim iARPSTmp As Long
    Dim iAvailableSpaceForTabs As Long
    Dim iRowsStretchRatio() As Single
    Dim iRowsStretchRatio_StartingRow As Long
    Dim iRowsStretchRatio_AccumulatedTabWidth As Long
    Dim r As Long
    Dim iAccumulatedAdditionalFixedTabSpace As Long
    Dim iRowsStretchRatio_AccumulatedAdditionalFixedTabWidth As Long
    Dim iSng As Single
    Dim iDecreaseStretchRatio As Boolean
    Dim iIncreaseStretchRatio As Boolean
    Dim iDoNotDecreaseStretchRatio As Boolean
    Dim iStyle2 As NTStyleConstants
    Dim iMessage As T_MSG
    Dim iAlreadyNeedToBePainted As Boolean
    Dim iDoNotDecreaseStretchRatio2 As Boolean
    Dim iTotalRowWidth As Long
    Dim iAvailableSpaceForTabsPrev As Long
    Dim iRowIsFilled() As Boolean
    Dim iRowStretchRatio As Single
    Dim iScaleWidthForTabs As Long
    Dim iMeasureCaptions As Boolean
    Dim iColumnWidthForHorizontalCaptions() As Long
    Dim iColumnLeftForHorizontalCaptions() As Long
    Dim iMaxTabsPerColumn As Long
    Dim iNumberOfColumnsForHorizontalCaptions As Long
    Dim iTabsPerColumn As Long
    Dim iTabHeightForHorizontalCaptions As Long
    Dim iColumnTabCount As Long
    Dim iTabExtraHeightIsNeeded As Boolean
    
    If mUserControlTerminated Or mTDIAddingNewTabForForm Then Exit Sub
    
    'If mResizeEventPending Then RaiseEvent_Resize
    If Not mRedraw Then
        mNeedToDraw = True
        If Not mEnsuringDrawn Then
            Exit Sub
        End If
    End If
    If Not mPropertiesReady Then
        PostDrawMessage
        Exit Sub
    End If
    
    If mSetAutoTabHeightPending Then SetAutoTabHeight
    tmrDraw.Enabled = False
    PeekMessage iMessage, mUserControlHwnd, WM_DRAW, WM_DRAW, PM_REMOVE ' remove posted message, if any
    mDrawMessagePosted = False
    If Not mAccessKeysSet Then
        If mAmbientUserMode Then SetAccessKeys
    End If
    UserControl.ScaleMode = vbPixels
    mScaleWidth = UserControl.ScaleWidth
    mScaleHeight = UserControl.ScaleHeight
    If Not ((mScaleWidth > 0) And (mScaleHeight > 0)) Then
        SetScaleMode
        Exit Sub
    End If
    If Not mFirstDraw Then mFirstDraw = True
    
    mDrawing = True
    For t = 0 To mTabs - 1
        mTabData(t).Selected = (t = mTab)
    Next
    mControlIsThemed = mVisualStyles And (mBackStyle <> ntTransparent)
    If mControlIsThemed Then
        If mTheme <> 0 Then
            CloseThemeData mTheme
            mTheme = 0
        End If
        mTheme = OpenThemeData(mUserControlHwnd, StrPtr("Tab"))
        If mTheme = 0 Then
            mControlIsThemed = False
        End If
        If mControlIsThemed Then
            SetThemeExtraData
        End If
    End If
    If mControlIsThemed Then
        iStyle2 = ssStylePropertyPage
    ElseIf mStyle = ntStyleTabStrip Then
        iStyle2 = ssStylePropertyPage
    Else
        iStyle2 = mStyle
    End If
    
    iLng = mTabAppearance2
    If mTabAppearance = ntTAAuto Then
        If (iStyle2 = ssStylePropertyPage) Or (iStyle2 = ntStyleTabStrip) Then
            mTabAppearance2 = ntTAPropertyPage
        ElseIf (iStyle2 = ntStyleFlat) Then
            mTabAppearance2 = ntTAFlat
        Else
            mTabAppearance2 = ntTATabbedDialog
        End If
    Else
        mTabAppearance2 = mTabAppearance
    End If
    mAppearanceIsPP = (mTabAppearance2 = ntTAPropertyPage) Or (mTabAppearance2 = ntTAPropertyPageRounded) Or mControlIsThemed
    mAppearanceIsFlat = (Not mAppearanceIsPP) And (mTabAppearance2 = ntTAFlat)
    If mTabAppearance2 <> iLng Then ResetCachedThemeImages
    
    iTabHeight = pScaleY(mTabHeight, vbHimetric, vbPixels)
    mFlatRoundnessTop2 = mFlatRoundnessTopDPIScaled
    If mFlatRoundnessTop2 > iTabHeight Then
        mFlatRoundnessTop2 = iTabHeight
    End If
    mFlatRoundnessTabs2 = mFlatRoundnessTabsDPIScaled
    If mFlatRoundnessTabs2 > iTabHeight Then
        mFlatRoundnessTabs2 = iTabHeight
    End If
    If mHighlightAddExtraHeight Or mHighlightAddExtraHeightSelectedTab Then iTabExtraHeight = pScaleY(mHighlightTabExtraHeight, vbHimetric, vbPixels)
    iTabMaxWidth = pScaleX(mTabMaxWidth, vbHimetric, vbPixels)
    iTabMinWidth = pScaleX(mTabMinWidth, vbHimetric, vbPixels)
    If mTabWidthStyle = ntTWAuto Then
        If mTDIMode <> ntTDIModeNone Then
            mTabWidthStyle2 = ntTWTabCaptionWidthFillRows
        ElseIf mStyle = ntStyleTabStrip Then
            mTabWidthStyle2 = ntTWTabStripEmulation
        ElseIf mStyle = ssStylePropertyPage Then
            mTabWidthStyle2 = ntTWTabCaptionWidth
        ElseIf (mStyle = ntStyleFlat) Or (mStyle = ntStyleWindows) Then
            mTabWidthStyle2 = ntTWStretchToFill
        Else
            mTabWidthStyle2 = ntTWFixed
        End If
    Else
        mTabWidthStyle2 = mTabWidthStyle
    End If
    
    If mShowRowsInPerspective = ntYNAuto Then
        If (mStyle = ntStyleTabStrip) Or (mStyle = ntStyleFlat) Or (mStyle = ntStyleWindows) Then
            mShowsRowsPerspective2 = False
        Else
            mShowsRowsPerspective2 = (mTabWidthStyle2 <> ntTWTabStripEmulation) And (mTabWidthStyle2 <> ntTWTabCaptionWidthFillRows)
        End If
    Else
        mShowsRowsPerspective2 = CBool(mShowRowsInPerspective)
    End If
    If mShowsRowsPerspective2 Then
        iRowPerspectiveSpace = pScaleX(cRowPerspectiveSpace, vbTwips, vbPixels)
    End If
    
    mTabSeparation2 = mTabSeparationDPIScaled
    If mControlIsThemed Then
        mTabSeparation2 = mTabSeparation2 - 2
        If mTabSeparation2 < 0 Then mTabSeparation2 = 0
    End If
    
    If (mTabOrientation = ssTabOrientationTop) Then
        m3DShadowH = m3DShadow
        m3DShadowV = m3DShadow
        m3DHighlightH = m3DHighlight
        m3DHighlightV = m3DHighlight
        m3DShadowH_Sel = m3DShadow_Sel
        m3DShadowV_Sel = m3DShadow_Sel
        m3DHighlightH_Sel = m3DHighlight_Sel
        m3DHighlightV_Sel = m3DHighlight_Sel
    ElseIf mTabOrientation = ssTabOrientationBottom Then
        m3DShadowH = m3DHighlight
        m3DShadowV = m3DShadow
        m3DHighlightH = m3DShadow
        m3DHighlightV = m3DHighlight
        m3DShadowH_Sel = m3DHighlight_Sel
        m3DShadowV_Sel = m3DShadow_Sel
        m3DHighlightH_Sel = m3DShadow_Sel
        m3DHighlightV_Sel = m3DHighlight_Sel
    ElseIf (mTabOrientation = ssTabOrientationLeft) Or (mTabOrientation = ntTabOrientationLeftHorizontal) Then
        m3DShadowH = m3DShadow
        m3DShadowV = m3DHighlight
        m3DHighlightH = m3DHighlight
        m3DHighlightV = m3DShadow
        m3DShadowH_Sel = m3DShadow_Sel
        m3DShadowV_Sel = m3DHighlight_Sel
        m3DHighlightH_Sel = m3DHighlight_Sel
        m3DHighlightV_Sel = m3DShadow_Sel
    ElseIf (mTabOrientation = ssTabOrientationRight) Or (mTabOrientation = ntTabOrientationRightHorizontal) Then
        m3DShadowH = m3DHighlight
        m3DShadowV = m3DShadow
        m3DHighlightH = m3DShadow
        m3DHighlightV = m3DHighlight
        m3DShadowH_Sel = m3DHighlight_Sel
        m3DShadowV_Sel = m3DShadow_Sel
        m3DHighlightH_Sel = m3DShadow_Sel
        m3DHighlightV_Sel = m3DHighlight_Sel
    End If
    mTabsAreRotatedButCaptionsAreHorizontal = (mTabOrientation = ntTabOrientationLeftHorizontal) Or (mTabOrientation = ntTabOrientationRightHorizontal)
    
    If mBackStyle = ntOpaque Then
        If mEnabled Or (Not mAmbientUserMode) Or (Not mShowDisabledState) Then
            mBackColorTabs2 = TranslatedColor(mBackColorTabs)
            If IsEmpty(mTabData(mTab).CustomColors.BackColorSelectedTab) Or (mHandleHighContrastTheme And mHighContrastThemeOn) Then
                mBackColorSelectedTab2 = TranslatedColor(mBackColorSelectedTab)
            Else
                mBackColorSelectedTab2 = TranslatedColor(mTabData(mTab).CustomColors.BackColorSelectedTab)
            End If
        Else
            mBackColorTabs2 = TranslatedColor(mBackColorTabsDisabled)
            mBackColorSelectedTab2 = TranslatedColor(mBackColorSelectedTabDisabled)
        End If
    Else
        mBackColorTabs2 = TranslatedColor(mBackColorTabs)
        TranslateColor mBackColorSelectedTab, 0, iLng
        If mBackColorTabs2 = iLng Then
            mBackColorTabs2 = mBackColorTabs2 Xor &H1
        End If
        mBackColorSelectedTab2 = TranslatedColor(mBackColorSelectedTab)
        UserControl.MaskColor = mBackColorSelectedTab
    End If
    UserControl.BackStyle = IIf(mBackStyle = ntOpaque, ntOpaque, ntTransparent)
    
    If (mTabOrientation = ssTabOrientationTop) Or (mTabOrientation = ssTabOrientationBottom) Then
        iScaleWidth = mScaleWidth
        iScaleHeight = mScaleHeight
    Else
        iScaleWidth = mScaleHeight
        iScaleHeight = mScaleWidth
    End If
    
    iScaleWidthForTabs = iScaleWidth - ScaleX(mTabsEndFreeSpace, vbTwips, vbPixels)
    If iScaleWidthForTabs < 1 Then iScaleWidthForTabs = 1
    ' measure tab captions and pic width
    ctv = -1
    If (mTabWidthStyle2 = ntTWTabStripEmulation) Or (mTabWidthStyle2 = ntTWStretchToFill) Then
        iTotalTabWidth = 0
        mRows = 1
    End If
    mVisibleTabs = 0
    iMeasureCaptions = (mTabWidthStyle2 = ntTWTabCaptionWidth) Or (mTabWidthStyle2 = ntTWTabCaptionWidthFillRows) Or (mTabWidthStyle2 = ntTWTabStripEmulation) Or (mTabWidthStyle2 = ntTWStretchToFill) Or mTabsAreRotatedButCaptionsAreHorizontal
    For t = 0 To mTabs - 1
        If mTabData(t).Visible Then
            mVisibleTabs = mVisibleTabs + 1
            ctv = ctv + 1
            If iMeasureCaptions Then
                iLng = MeasureTabIconAndCaption(t)
                If iTabMinWidth > 0 Then
                    If (iLng + 10) < iTabMinWidth Then
                        iLng = iTabMinWidth - 10
                    End If
                End If
                If iTabMaxWidth > 0 Then
                    If (iLng + 10) > iTabMaxWidth Then
                        iLng = iTabMaxWidth - 10
                    End If
                End If
                mTabData(t).IconAndCaptionWidth = iLng
            End If
        End If
    Next t
    
    If (mVisibleTabs = 0) Or (mTab = -1) Then
        Set UserControl.Picture = Nothing
        mBodyReset = True
        GoTo TheExit:
    End If
    
    ' set data about tabs placement on rows
    iLastVisibleTab = 0
    If mTabsAreRotatedButCaptionsAreHorizontal Then
        iAvailableSpaceForTabs = (iScaleWidthForTabs - IIf(mAppearanceIsPP, 4, 0))
        Do
            iMaxTabsPerColumn = Int((iAvailableSpaceForTabs + mTabSeparation2) / (iTabHeight + mTabSeparation2))
            If iMaxTabsPerColumn = 0 Then
                iMaxTabsPerColumn = 1
                Exit Do
            End If
            iNumberOfColumnsForHorizontalCaptions = Int(mTabs / iMaxTabsPerColumn) + IIf(Int(mTabs / iMaxTabsPerColumn) < (mTabs / iMaxTabsPerColumn), 1, 0)
            
            iAvailableSpaceForTabsPrev = iAvailableSpaceForTabs
            iAvailableSpaceForTabs = (iScaleWidthForTabs - IIf(mAppearanceIsPP, 4, 0)) - iRowPerspectiveSpace * (iNumberOfColumnsForHorizontalCaptions - 1) ' - mTabSeparation2 * iMaxTabsPerColumn
            If iAvailableSpaceForTabs >= iAvailableSpaceForTabsPrev Then
                iAvailableSpaceForTabs = iAvailableSpaceForTabsPrev
                Exit Do
            End If
        Loop
        If iNumberOfColumnsForHorizontalCaptions = 0 Then iNumberOfColumnsForHorizontalCaptions = 1
        iTabsPerColumn = Int(mTabs / iNumberOfColumnsForHorizontalCaptions) + IIf(mTabs / iNumberOfColumnsForHorizontalCaptions > Int(mTabs / iNumberOfColumnsForHorizontalCaptions), 1, 0)
        mRows = iNumberOfColumnsForHorizontalCaptions
        
        iRow = 0
        iPosH = 0
        ctv = 0
        For t = 0 To mTabs - 1
            If mTabData(t).Visible Then
                ctv = ctv + 1
                iLastVisibleTab = t
                mTabData(t).LeftTab = False
                mTabData(t).RightTab = False
                iPosH = iPosH + 1
                If iPosH > iTabsPerColumn Then
                    iPosH = 1
                    iRow = iRow + 1
                    mTabData(t - 1).RightTab = True
                End If
                mTabData(t).PosH = iPosH
                mTabData(t).Row = iRow
                If iPosH = 1 Then
                    mTabData(t).LeftTab = True
                End If
                If (ctv = mVisibleTabs) Then
                    mTabData(t).RightTab = True
                End If
            Else
                mTabData(t).Row = -1
            End If
        Next
    ElseIf (mTabWidthStyle2 <> ntTWTabStripEmulation) And (mTabWidthStyle2 <> ntTWStretchToFill) Then
        If (mTabWidthStyle2 = ntTWTabCaptionWidthFillRows) Then
            iTotalTabWidth = 0
            For t = 0 To mTabs - 1
                If mTabData(t).Visible Then
                    iTotalTabWidth = iTotalTabWidth + mTabData(t).IconAndCaptionWidth + 10 + mTabSeparation2
                End If
            Next
            iAvailableSpaceForTabs = (iScaleWidthForTabs - IIf(mAppearanceIsPP, 4, 0))
            ReDim iRowIsFilled(0)
            Do
                iRow = 0
                iPosH = 0
                ctv = 0
                iAccumulatedTabWith = 0
                iAccumulatedAdditionalFixedTabSpace = 0
                For t = 0 To mTabs - 1
                    If mTabData(t).Visible Then
                        'iRowTotalTabWidth(0) = iRowTotalTabWidth(0) + mTabData(t).IconAndCaptionWidth
                        mTabData(t).TopTab = False
                        ctv = ctv + 1
                        iLastVisibleTab = t
                        mTabData(t).LeftTab = False
                        mTabData(t).RightTab = False
                        iPosH = iPosH + 1
                        If Not iPosH = 1 Then ' if not the first tab already exceeds the available width
                            If (iAccumulatedTabWith + iAccumulatedAdditionalFixedTabSpace + mTabData(t).IconAndCaptionWidth + 10) > iAvailableSpaceForTabs Then
                                iPosH = 1
                                If iRow > UBound(iRowIsFilled) Then
                                    ReDim Preserve iRowIsFilled(iRow)
                                End If
                                iRowIsFilled(iRow) = True
                                iRow = iRow + 1
                                iAccumulatedTabWith = 0
                                iAccumulatedAdditionalFixedTabSpace = 0
                                mTabData(t - 1).RightTab = True
                            End If
                        End If
                        iAccumulatedTabWith = iAccumulatedTabWith + mTabData(t).IconAndCaptionWidth
                        iAccumulatedAdditionalFixedTabSpace = iAccumulatedAdditionalFixedTabSpace + 10 + mTabSeparation2
                        mTabData(t).PosH = iPosH
                        If iPosH = 1 Then
                            mTabData(t).LeftTab = True
                        End If
                        If (ctv = mVisibleTabs) Then
                            mTabData(t).RightTab = True
                        End If
                        mTabData(t).Row = iRow
                    Else
                        mTabData(t).Row = -1
                    End If
                Next t
                mRows = iRow + 1
                
                iAvailableSpaceForTabsPrev = iAvailableSpaceForTabs
                iAvailableSpaceForTabs = (iScaleWidthForTabs - IIf(mAppearanceIsPP, 4, 0)) - iRowPerspectiveSpace * (mRows - 1)
                If iAvailableSpaceForTabs = iAvailableSpaceForTabsPrev Then Exit Do
                ReDim iRowIsFilled(mRows - 1)
            Loop
            If UBound(iRowIsFilled) <> mRows - 1 Then
                ReDim Preserve iRowIsFilled(mRows - 1)
            End If
            
            For r = 0 To mRows - 1
                If iRowIsFilled(r) Then
                    iTotalRowWidth = 0
                    For t = 0 To mTabs - 1
                        If mTabData(t).Row = r Then
                            iTotalRowWidth = iTotalRowWidth + mTabData(t).IconAndCaptionWidth '+ 10 + mTabSeparation2
                        End If
                    Next
                    iTotalRowWidth = iTotalRowWidth + 10 + mTabSeparation2
                    For t = 0 To mTabs - 1
                        iRowStretchRatio = iAvailableSpaceForTabs / iTotalRowWidth
                        If mTabData(t).Row = r Then
                            'mTabData(t).Width = mTabData(t).IconAndCaptionWidth * iRowStretchRatio   '- 10 - mTabSeparation2
                            mTabData(t).IconAndCaptionWidth = mTabData(t).IconAndCaptionWidth * iRowStretchRatio - 1 '+ 10 + mTabSeparation2
                        End If
                    Next
                End If
            Next
        Else
            iRow = 0
            iPosH = 0
            ctv = 0
            For t = 0 To mTabs - 1
                If mTabData(t).Visible Then
                    mTabData(t).TopTab = False
                    ctv = ctv + 1
                    iLastVisibleTab = t
                    mTabData(t).LeftTab = False
                    mTabData(t).RightTab = False
                    iPosH = iPosH + 1
                    If iPosH > mTabsPerRow Then
                        iPosH = 1
                        iRow = iRow + 1
                    End If
                    mTabData(t).PosH = iPosH
                    If iPosH = 1 Then
                        mTabData(t).LeftTab = True
                    End If
                    If (iPosH = mTabsPerRow) Or (ctv = mVisibleTabs) Then
                        mTabData(t).RightTab = True
                    End If
                    mTabData(t).Row = iRow
                Else
                    mTabData(t).Row = -1
                End If
            Next t
            mRows = iRow + 1
        End If
    Else
        ' define what tabs to place on each row when tabs are justified
        ' step 1: calculate the number of rows that will be needed and the iRowsStretchRatio for each row (that will be needed in the step 2)
        iARPSTmp = 0
        Do
            iAllRowsPerspectiveSpace = iARPSTmp
            iAvailableSpaceForTabs = (iScaleWidthForTabs - iAllRowsPerspectiveSpace - IIf(mAppearanceIsPP, 4, 0))
            iAccumulatedTabWith = 0
            iAccumulatedAdditionalFixedTabSpace = 0
            iRow = 0
            ReDim iRowsStretchRatio(0)
            iRowsStretchRatio_StartingRow = 0
            iRowsStretchRatio_AccumulatedTabWidth = 0
            iRowsStretchRatio_AccumulatedAdditionalFixedTabWidth = 0
            iRowTabCount = 0
            For t = 0 To mTabs - 1
                If mTabData(t).Visible Then
                    If (iAccumulatedTabWith + iAccumulatedAdditionalFixedTabSpace + mTabData(t).IconAndCaptionWidth + 10) > iAvailableSpaceForTabs Then
                        If iRowTabCount = 0 Then ' this only tab alone passes the available space in the row (and it is the first one or all the previous tabs also entered here)
                            If t < (mTabs - 1) Then
                                iRowsStretchRatio(iRow) = 1
                                iRow = iRow + 1
                                iRowsStretchRatio_StartingRow = iRow
                                iRowsStretchRatio_AccumulatedTabWidth = 0
                                iRowsStretchRatio_AccumulatedAdditionalFixedTabWidth = 0
                                ReDim Preserve iRowsStretchRatio(iRow)
                            End If
                            iRowTabCount = 0
                            iAccumulatedTabWith = 0
                            iAccumulatedAdditionalFixedTabSpace = 0
                        ElseIf iRowTabCount = 1 Then ' this only tab alone passes the available space in the row (and it comes from a previus attempt to put it in the previous row)
                            iSng = ((iRow - iRowsStretchRatio_StartingRow) * iAvailableSpaceForTabs - iRowsStretchRatio_AccumulatedAdditionalFixedTabWidth) / iRowsStretchRatio_AccumulatedTabWidth
                            If iSng < 1 Then
                                iDoNotDecreaseStretchRatio = True
                                iSng = 1
                            End If
                            For r = iRowsStretchRatio_StartingRow To iRow - 1
                                iRowsStretchRatio(r) = iSng
                            Next r
                            iRowsStretchRatio(iRow) = 1
                            iRow = iRow + 1
                            iRowsStretchRatio_StartingRow = iRow
                            ReDim Preserve iRowsStretchRatio(iRow)
                            iRowTabCount = 1
                            iAccumulatedTabWith = mTabData(t).IconAndCaptionWidth
                            iAccumulatedAdditionalFixedTabSpace = 10 + mTabSeparation2
                            iRowsStretchRatio_AccumulatedTabWidth = iAccumulatedTabWith
                            iRowsStretchRatio_AccumulatedAdditionalFixedTabWidth = iAccumulatedAdditionalFixedTabSpace
                        Else
                            iRow = iRow + 1
                            ReDim Preserve iRowsStretchRatio(iRow)
                            iRowTabCount = 1
                            iAccumulatedTabWith = mTabData(t).IconAndCaptionWidth
                            iAccumulatedAdditionalFixedTabSpace = 10 + mTabSeparation2
                            iRowsStretchRatio_AccumulatedTabWidth = iRowsStretchRatio_AccumulatedTabWidth + iAccumulatedTabWith
                            iRowsStretchRatio_AccumulatedAdditionalFixedTabWidth = iRowsStretchRatio_AccumulatedAdditionalFixedTabWidth + iAccumulatedAdditionalFixedTabSpace
                        End If
                    Else
                        iAccumulatedTabWith = iAccumulatedTabWith + mTabData(t).IconAndCaptionWidth
                        iAccumulatedAdditionalFixedTabSpace = iAccumulatedAdditionalFixedTabSpace + 10 + mTabSeparation2
                        iRowsStretchRatio_AccumulatedTabWidth = iRowsStretchRatio_AccumulatedTabWidth + mTabData(t).IconAndCaptionWidth
                        iRowsStretchRatio_AccumulatedAdditionalFixedTabWidth = iRowsStretchRatio_AccumulatedAdditionalFixedTabWidth + 10 + mTabSeparation2
                        iRowTabCount = iRowTabCount + 1
                    End If
                End If
            Next t
            If iRowsStretchRatio_AccumulatedTabWidth > 0 Then
                iSng = ((iRow - iRowsStretchRatio_StartingRow + 1) * iAvailableSpaceForTabs - iRowsStretchRatio_AccumulatedAdditionalFixedTabWidth) / iRowsStretchRatio_AccumulatedTabWidth
                If iSng < 1 Then
                    iDoNotDecreaseStretchRatio = True
                    iSng = 1
                End If
                For r = iRowsStretchRatio_StartingRow To iRow
                    iRowsStretchRatio(r) = iSng
                Next r
            End If
            mRows = iRow + 1
            iARPSTmp = (mRows - 1) * iRowPerspectiveSpace
            If Not mShowsRowsPerspective2 Then
                iAllRowsPerspectiveSpace = iARPSTmp
                Exit Do
            End If
        Loop Until iARPSTmp = iAllRowsPerspectiveSpace ' until it did not add another row
        
        ' step 2: set in what row goes each tab
        iDecreaseStretchRatio = False
        iIncreaseStretchRatio = False
        iDoNotDecreaseStretchRatio2 = False
        Do
            iRowTabCount = 0
            iAccumulatedTabWith = 0
            iRow = 0
            ctv = 0
            If iDecreaseStretchRatio Then
                For r = 0 To mRows - 1
                    iRowsStretchRatio(r) = iRowsStretchRatio(r) * 0.95
                    If iRowsStretchRatio(r) < 1 Then
                        iRowsStretchRatio(r) = 1
                        iDoNotDecreaseStretchRatio2 = True
                    End If
                Next r
                iDecreaseStretchRatio = False
            ElseIf iIncreaseStretchRatio Then
                For r = 0 To mRows - 1
                    iRowsStretchRatio(r) = iRowsStretchRatio(r) * 1.05
                Next r
                iIncreaseStretchRatio = False
            End If
            iLastVisibleTab_Prev = -1
            For t = 0 To mTabs - 1
                If mTabData(t).Visible Then
                    mTabData(t).TopTab = False
                    ctv = ctv + 1
                    iLastVisibleTab_Prev = iLastVisibleTab
                    iLastVisibleTab = t
                    mTabData(t).LeftTab = False
                    mTabData(t).RightTab = False
                    If ctv = mVisibleTabs Then
                        mTabData(t).RightTab = True
                    End If
                    iLng = mTabData(t).IconAndCaptionWidth * iRowsStretchRatio(iRow) + 10
                    If iAccumulatedTabWith + iLng > (iAvailableSpaceForTabs + mTabData(t).IconAndCaptionWidth * 0.38) Then ' 0.38 is an add-hoc value, the right thing to do would be to make another step and recalculate everything several times changing the stretch ratio until an equilibrium point is found (or something like that). But with a couple of examples it seems too work acceptable with this value of 0.38. If there are too many tabs or too few tabs in the top row, here is the problem (probably).
                        If iRowTabCount = 0 Then ' this only tab alone passes the available space in the row (and it is the first one or all the previous tabs also entered here)
                            mTabData(t).Row = iRow
                            mTabData(t).PosH = 1
                            mTabData(t).LeftTab = True
                            mTabData(t).RightTab = True
                            If (iRow + 1) < mRows Then
                                iRow = iRow + 1
                            End If
                            iRowTabCount = 0
                            iAccumulatedTabWith = 0
                        Else
                            If (iRow + 1) < mRows Then
                                If iLastVisibleTab_Prev <> t Then
                                    mTabData(iLastVisibleTab_Prev).RightTab = True
                                End If
                                iRow = iRow + 1
                                iRowTabCount = 1
                                iAccumulatedTabWith = iLng + mTabSeparation2
                            Else
                                iRowTabCount = iRowTabCount + 1
                            End If
                            mTabData(t).PosH = iRowTabCount
                            If iRowTabCount = 1 Then
                                mTabData(t).LeftTab = True
                            End If
                            mTabData(t).Row = iRow
                        End If
                    Else
                        iAccumulatedTabWith = iAccumulatedTabWith + iLng + mTabSeparation2
                        iRowTabCount = iRowTabCount + 1
                        mTabData(t).PosH = iRowTabCount
                        If iRowTabCount = 1 Then
                            mTabData(t).LeftTab = True
                        End If
                        mTabData(t).Row = iRow
                    End If
                Else
                    mTabData(t).Row = -1
                End If
            Next t
            mTabData(iLastVisibleTab).PosH = iRowTabCount
            If iRowTabCount = 1 Then
                mTabData(iLastVisibleTab).LeftTab = True
            End If
            mTabData(iLastVisibleTab).RightTab = True
            
            If (mRows = 1) And (mTabWidthStyle2 <> ntTWStretchToFill) Then
                mTabWidthStyle2 = ntTWTabCaptionWidth
            End If
            If (mTabWidthStyle2 = ntTWTabStripEmulation) Or (mTabWidthStyle2 = ntTWStretchToFill) Then
                ' step 3: set the widths of the tabs for each row
                For iRow = 0 To mRows - 1
                    iAccumulatedTabWith = 0
                    iAccumulatedAdditionalFixedTabSpace = 0
                    iRowTabCount = 0
                    For t = 0 To mTabs - 1
                        If mTabData(t).Row = iRow Then
                            iAccumulatedTabWith = iAccumulatedTabWith + mTabData(t).IconAndCaptionWidth
                            iAccumulatedAdditionalFixedTabSpace = iAccumulatedAdditionalFixedTabSpace + 10 + mTabSeparation2
                            iRowTabCount = iRowTabCount + 1
                        End If
                    Next t
                    If iRowTabCount > 1 Then
                        iAccumulatedAdditionalFixedTabSpace = iAccumulatedAdditionalFixedTabSpace - mTabSeparation2
                        iSng = (iAvailableSpaceForTabs - iAccumulatedAdditionalFixedTabSpace) / iAccumulatedTabWith
                        If iSng < 1 Then
                            If Not (iDoNotDecreaseStretchRatio Or iDoNotDecreaseStretchRatio2) Then
                                iDecreaseStretchRatio = True
                                Exit For
                            End If
                        End If
                    Else
                        If iAccumulatedTabWith = 0 Then
                            iSng = 1
                        Else
                            iSng = (iAvailableSpaceForTabs - iAccumulatedAdditionalFixedTabSpace) / iAccumulatedTabWith
                        End If
                    End If
                    For t = 0 To mTabs - 1
                        If mTabData(t).Row = iRow Then
                            mTabData(t).Width = mTabData(t).IconAndCaptionWidth * iSng
                        End If
                    Next t
                Next iRow
            End If
            
            For iRow = mRows - 1 To 1 Step -1
                iLng = 0
                For t = 0 To mTabs - 1
                    If mTabData(t).Row = iRow Then
                        iLng = iLng + 1
                        Exit For
                    End If
                Next t
                If iLng = 0 Then
                    iIncreaseStretchRatio = True
                End If
            Next iRow
        Loop While (iDecreaseStretchRatio Or iIncreaseStretchRatio) And ((mTabWidthStyle2 = ntTWTabStripEmulation) Or (mTabWidthStyle2 = ntTWStretchToFill))
    End If
    
    iTabExtraHeightIsNeeded = mHighlightAddExtraHeight Or mHighlightAddExtraHeightSelectedTab And (mRows = 1)

    ' Rows positions
    For t = 0 To mTabs - 1
        mTabData(t).RowPos = (mRows - mTabData(t).Row - 1) + mTabData(mTab).Row
        If mTabData(t).RowPos > (mRows - 1) Then mTabData(t).RowPos = mTabData(t).RowPos - mRows
    Next t
    
    If mTabsAreRotatedButCaptionsAreHorizontal Then
        ReDim iColumnWidthForHorizontalCaptions(mRows - 1)
        ReDim iColumnLeftForHorizontalCaptions(mRows - 1)
        For iRow = 0 To mRows - 1
            iColumnWidthForHorizontalCaptions(iRow) = 0
            If iRow = 0 Then
                iColumnLeftForHorizontalCaptions(iRow) = 0
            Else
                iColumnLeftForHorizontalCaptions(iRow) = iColumnLeftForHorizontalCaptions(iRow - 1) + iColumnWidthForHorizontalCaptions(iRow - 1)
            End If
            iColumnTabCount = 0
            For t = 0 To mTabs - 1
                If mTabData(t).RowPos = iRow Then
                    iColumnTabCount = iColumnTabCount + 1
                    If mAppearanceIsFlat Then
                        iLng = mTabData(t).IconAndCaptionWidth + IIf(mFlatBarGripHeightDPIScaled > 0, mFlatBarGripHeightDPIScaled, 0) + mFlatBarHeightDPIScaled
                    Else
                        iLng = mTabData(t).IconAndCaptionWidth
                    End If
                    iLng = iLng + 10 ' some more padding?
                    If iLng > iColumnWidthForHorizontalCaptions(iRow) Then
                        iColumnWidthForHorizontalCaptions(iRow) = iLng
                    End If
                End If
                mTabData(t).TopTab = (iRow = (mRows - 1))
            Next
            If iColumnWidthForHorizontalCaptions(iRow) = 0 Then iColumnWidthForHorizontalCaptions(iRow) = 1600
            If iTabMaxWidth > 0 Then
                If iColumnWidthForHorizontalCaptions(iRow) > iTabMaxWidth Then
                    iColumnWidthForHorizontalCaptions(iRow) = iTabMaxWidth
                End If
            End If
            If iTabMinWidth > 0 Then
                If iColumnWidthForHorizontalCaptions(iRow) < iTabMinWidth Then
                    iColumnWidthForHorizontalCaptions(iRow) = iTabMinWidth
                End If
            End If
            iTabHeightForHorizontalCaptions = (iAvailableSpaceForTabs - (iColumnTabCount - 1) * mTabSeparation2) / iColumnTabCount - 0.5
        
            For t = 0 To mTabs - 1
                If mTabData(t).Row = iRow Then
                    mTabData(t).Width = iTabHeightForHorizontalCaptions
                End If
            Next
        Next
    End If
    
    ' Body position
    If mTabsAreRotatedButCaptionsAreHorizontal Then
        mClientStart = 2
        For iRow = 0 To mRows - 1
            mClientStart = mClientStart + iColumnWidthForHorizontalCaptions(iRow)
        Next
        If iTabExtraHeightIsNeeded Then
            mClientStart = mClientStart + iTabExtraHeight
        End If
    ElseIf iTabExtraHeightIsNeeded Then
        mClientStart = mRows * iTabHeight + iTabExtraHeight + 2
    Else
        mClientStart = mRows * iTabHeight + 2
    End If
    If mHighlightFlatBarWithGrip Or mHighlightFlatBarWithGripSelectedTab Then
        If mTabOrientation = ssTabOrientationBottom Then
            If (mFlatBarPosition = ntBarPositionBottom) And (mFlatBarGripHeightDPIScaled > 0) Then
                mClientStart = mClientStart + mFlatBarGripHeightDPIScaled
            End If
        Else
            If (mFlatBarPosition = ntBarPositionTop) And (mFlatBarGripHeightDPIScaled > 0) Then
                mClientStart = mClientStart + mFlatBarGripHeightDPIScaled
            End If
        End If
    End If
    mClientHeight = iScaleHeight - mClientStart + 2 '+ 1
    
    If mRows > 1 Then
        iAllRowsPerspectiveSpace = iRowPerspectiveSpace * (mRows - 1)
    End If
    mClientWidth = iScaleWidth - iAllRowsPerspectiveSpace '- 1
    If mControlIsThemed Then
        mClientWidth = mClientWidth + mThemedBodyRightShadowPixels - 2
    End If
    
    Select Case mTabOrientation
        Case ssTabOrientationTop
            mClientRect.Top = mClientStart
            mClientRect.Left = 2
            mClientRect.Bottom = mScaleHeight - 4
            mClientRect.Right = mClientWidth - 4
        Case ssTabOrientationBottom
            mClientRect.Top = 0 '2
            mClientRect.Left = 2
            mClientRect.Bottom = mClientHeight - 6 '4
            mClientRect.Right = mClientWidth - 4
        Case ssTabOrientationLeft, ntTabOrientationLeftHorizontal
            mClientRect.Top = mScaleHeight - mClientWidth '+ 2
            mClientRect.Left = mClientStart + 2
            mClientRect.Bottom = mScaleHeight - 4
            mClientRect.Right = mScaleWidth - 4
        Case Else ' ssTabOrientationRight, ntTabOrientationRightHorizontal
            mClientRect.Top = 0 '2
            mClientRect.Left = 2
            mClientRect.Bottom = mClientWidth - 4
            mClientRect.Right = mClientHeight - 6 '4
    End Select
    If mAppearanceIsFlat Then
        mClientRect.Left = mClientRect.Left - 1
        mClientRect.Top = mClientRect.Top + 1
        mClientRect.Right = mClientRect.Right + 3
        mClientRect.Bottom = mClientRect.Bottom + 3
    End If
    
    If (mTabWidthStyle2 = ntTWTabCaptionWidth) Or (mTabWidthStyle2 = ntTWTabCaptionWidthFillRows) Then
        iAvailableSpaceForTabs = (iScaleWidthForTabs - iAllRowsPerspectiveSpace - IIf(mAppearanceIsPP, 4, 0))
        For iRow = 0 To mRows - 1
            iAccumulatedTabWith = 0
            iAccumulatedAdditionalFixedTabSpace = 0
            For t = 0 To mTabs - 1
                If mTabData(t).Row = iRow Then
                    iAccumulatedTabWith = iAccumulatedTabWith + mTabData(t).IconAndCaptionWidth
                    iAccumulatedAdditionalFixedTabSpace = iAccumulatedAdditionalFixedTabSpace + 12
                    If Not mTabData(t).RightTab Then
                        iAccumulatedAdditionalFixedTabSpace = iAccumulatedAdditionalFixedTabSpace + mTabSeparation2
                    End If
                End If
            Next t
            'If mAmbientUserMode Then
            mMinSizeNeeded = (iScaleWidth - iAvailableSpaceForTabs) + iAccumulatedTabWith + iAccumulatedAdditionalFixedTabSpace
            If iAccumulatedTabWith + iAccumulatedAdditionalFixedTabSpace > iAvailableSpaceForTabs Then
                iSng = (iAvailableSpaceForTabs - iAccumulatedAdditionalFixedTabSpace) / iAccumulatedTabWith
                For t = 0 To mTabs - 1
                    If mTabData(t).Row = iRow Then
                        mTabData(t).IconAndCaptionWidth = mTabData(t).IconAndCaptionWidth * iSng
                    End If
                Next t
            End If
            'End If
        Next iRow
    End If
    
    ' minimun size
    If (mUserControlSizeCorrectionsCounter < 3) And (Not mAmbientUserMode) Then
        If (mUserControlSizeCorrectionsCounter_ScaleWidthNotToResize <> mScaleWidth) Or (mUserControlSizeCorrectionsCounter_ScaleHeightNotToResize <> mScaleHeight) Then
            If (mTabOrientation = ssTabOrientationTop) Or (mTabOrientation = ssTabOrientationBottom) Then
                If mClientHeight < 3 Then
                    UserControl.Height = UserControl.Height + pScaleY(3 - mClientHeight, vbPixels, vbTwips)
                    mUserControlSizeCorrectionsCounter = mUserControlSizeCorrectionsCounter + 1
                    GoTo TheExit:
                End If
                If (mTabWidthStyle2 = ntTWFixed) Or (mTabWidthStyle2 = ntTWTabCaptionWidth) Then
                    If UserControl.Width < CLng(mTabsPerRow) * 500 + pScaleX(iAllRowsPerspectiveSpace, vbPixels, vbTwips) Then
                        UserControl.Width = CLng(mTabsPerRow) * 500 + pScaleX(iAllRowsPerspectiveSpace, vbPixels, vbTwips) + Screen_TwipsPerPixelX
                        mUserControlSizeCorrectionsCounter = mUserControlSizeCorrectionsCounter + 1
                        GoTo TheExit:
                    End If
                End If
            Else
                If mClientHeight < 3 Then
                    iLng = UserControl.Width + pScaleX(3 - mClientHeight, vbPixels, vbTwips)
                    UserControl.Width = iLng
                    mUserControlSizeCorrectionsCounter = mUserControlSizeCorrectionsCounter + 1
                    GoTo TheExit:
                End If
                If (mTabWidthStyle2 = ntTWFixed) Or (mTabWidthStyle2 = ntTWTabCaptionWidth) Then
                    If UserControl.Height < mTabsPerRow * 500 + pScaleX(iAllRowsPerspectiveSpace, vbPixels, vbTwips) Then ' we are drawing horizontally, so ScaleX
                        UserControl.Height = mTabsPerRow * 500 + pScaleX(iAllRowsPerspectiveSpace, vbPixels, vbTwips) + Screen_TwipsPerPixely
                        mUserControlSizeCorrectionsCounter = mUserControlSizeCorrectionsCounter + 1
                        GoTo TheExit:
                    End If
                End If
            End If
        End If
    ElseIf Not mAmbientUserMode Then
        mUserControlSizeCorrectionsCounter_ScaleWidthNotToResize = mScaleWidth
        mUserControlSizeCorrectionsCounter_ScaleHeightNotToResize = mScaleHeight
    End If
    If (iTabMaxWidth > 0) And (mTabWidthStyle2 = ntTWFixed) Then
        If (mUserControlSizeCorrectionsCounter < 3) And (Not mAmbientUserMode) Then
            If (mUserControlSizeCorrectionsCounter_ScaleWidthNotToResize <> mScaleWidth) Or (mUserControlSizeCorrectionsCounter_ScaleHeightNotToResize <> mScaleHeight) Then
                iLng = iTabMaxWidth * mTabsPerRow
                If (mTabOrientation = ssTabOrientationTop) Or (mTabOrientation = ssTabOrientationBottom) Then
                    If pScaleX(iLng, vbPixels, vbTwips) > UserControl.Width Then
                        UserControl.Width = pScaleX(iLng, vbPixels, vbTwips)
                        mUserControlSizeCorrectionsCounter = mUserControlSizeCorrectionsCounter + 1
                        GoTo TheExit:
                    End If
                Else
                    If pScaleY(iLng, vbPixels, vbTwips) > UserControl.Height Then
                        UserControl.Height = pScaleY(iLng, vbPixels, vbTwips)
                        mUserControlSizeCorrectionsCounter = mUserControlSizeCorrectionsCounter + 1
                        GoTo TheExit:
                    End If
                End If
            End If
        ElseIf Not mAmbientUserMode Then
            mUserControlSizeCorrectionsCounter_ScaleWidthNotToResize = mScaleWidth
            mUserControlSizeCorrectionsCounter_ScaleHeightNotToResize = mScaleHeight
        End If
        If mAppearanceIsPP Then
            iTabWidth = (iScaleWidthForTabs - 5 - iAllRowsPerspectiveSpace - 1 - IIf(mControlIsThemed, 2 - mThemedBodyRightShadowPixels, 0) - mTabSeparation2 * (mTabsPerRow - 1)) / mTabsPerRow
        Else
            iTabWidth = (iScaleWidthForTabs - 1 - iAllRowsPerspectiveSpace - 1 - IIf(mControlIsThemed, 2 - mThemedBodyRightShadowPixels, 0) - mTabSeparation2 * (mTabsPerRow - 1)) / mTabsPerRow
        End If
        If iTabWidth > iTabMaxWidth Then
            iTabWidth = iTabMaxWidth
        End If
    Else
        iTabWidth = (iScaleWidthForTabs - iAllRowsPerspectiveSpace - 1 - mTabSeparation2 * (mTabsPerRow - 1)) / mTabsPerRow
    End If
    
    mUserControlSizeCorrectionsCounter = 0
    
    If (mClientWidth_Prev <> mClientWidth) And (mClientWidth_Prev <> 0) Or (mClientHeight_Prev <> mClientHeight) And (mClientHeight_Prev <> 0) Then
        ResetCachedThemeImages
    End If
    mClientWidth_Prev = mClientWidth
    mClientHeight_Prev = mClientHeight
    
    If (mTabWidthStyle2 <> ntTWTabStripEmulation) And (mTabWidthStyle2 <> ntTWStretchToFill) Then
        iTabStretchRatio = 1
    End If
    
    ReDim mRightMostTabsRightPos(mRows - 1)
    
    ' set the tab rects
    If mTabsAreRotatedButCaptionsAreHorizontal Then
        For iRow = 0 To mRows - 1
            For t = 0 To mTabs - 1
                If mTabData(t).Visible Then
                    If mTabData(t).RowPos = iRow Then
                        iTabData = mTabData(t)
                        With iTabData.TabRect
                            If t = mTab Then
                                If (iTabExtraHeight > 0) And mHighlightAddExtraHeightSelectedTab Then
                                    .Top = iColumnLeftForHorizontalCaptions(iRow) - iTabExtraHeight
                                Else
                                    .Top = iColumnLeftForHorizontalCaptions(iRow)
                                End If
                                If iTabExtraHeightIsNeeded Then .Top = .Top + iTabExtraHeight
                                .Bottom = .Top + iColumnWidthForHorizontalCaptions(iRow) + IIf(mHighlightAddExtraHeightSelectedTab, iTabExtraHeight, 0)
                                If mHighlightFlatBarWithGrip Or mHighlightFlatBarWithGripSelectedTab Then
                                    If mTabOrientation = ssTabOrientationBottom Then
                                        If (mFlatBarPosition = ntBarPositionBottom) And (mFlatBarGripHeightDPIScaled > 0) Then
                                            .Top = .Top + mFlatBarGripHeightDPIScaled
                                            .Bottom = .Bottom + mFlatBarGripHeightDPIScaled
                                        End If
                                    Else
                                        If (mFlatBarPosition = ntBarPositionTop) And (mFlatBarGripHeightDPIScaled > 0) Then
                                            .Top = .Top + mFlatBarGripHeightDPIScaled
                                            .Bottom = .Bottom + mFlatBarGripHeightDPIScaled
                                        End If
                                    End If
                                End If
                            Else
                                If iTabData.Hovered And mHighlightAddExtraHeight And mAmbientUserMode Then
                                    .Top = iColumnLeftForHorizontalCaptions(iRow) - iTabExtraHeight
                                    If iTabExtraHeightIsNeeded Then .Top = .Top + iTabExtraHeight
                                    .Bottom = .Top + iColumnWidthForHorizontalCaptions(iRow) + IIf(mHighlightAddExtraHeight, iTabExtraHeight, 0)
                                Else
                                    .Top = iColumnLeftForHorizontalCaptions(iRow)
                                    If iTabExtraHeightIsNeeded Then .Top = .Top + iTabExtraHeight
                                    .Bottom = .Top + iColumnWidthForHorizontalCaptions(iRow)
                                End If
                                If mHighlightFlatBarWithGrip Or mHighlightFlatBarWithGripSelectedTab Then
                                    If (mFlatBarPosition = ntBarPositionTop) And (mFlatBarGripHeightDPIScaled > 0) Then
                                        .Top = .Top + mFlatBarGripHeightDPIScaled
                                        .Bottom = .Bottom + mFlatBarGripHeightDPIScaled
                                    End If
                                End If
                            End If
                            
                            .Left = (iTabData.PosH - 1) * IIf(mControlIsThemed, iTabData.Width, Round(iTabData.Width)) + iRowPerspectiveSpace * (mRows - mTabData(t).RowPos - 1) + 1 + (iTabData.PosH - 1) * mTabSeparation2 + 1
                            .Right = .Left + iTabData.Width - 1 '- mTabSeparation2 ' no volver a sacar el -1!!
                            
                            If iTabData.RightTab Then
                                iLng = iScaleWidthForTabs - iRowPerspectiveSpace * mTabData(t).RowPos - 1
                                If mAppearanceIsPP Then
                                    iLng = iLng - 2
                                    If mControlIsThemed Then
                                        If ((mTabWidthStyle2 <> ntTWTabStripEmulation) And (mTabWidthStyle2 <> ntTWStretchToFill)) Or iTabData.Selected Then
                                            iLng = iLng - 1
                                        End If
                                    End If
                                End If
                                If t = mTab Then
                                    If mControlIsThemed Then
                                        iLng = iLng + 1
                                    End If
                                End If
                                If Abs(.Right - iLng) < (6 + iColumnTabCount / 2) Then
                                    .Right = iLng - IIf(mControlIsThemed, mThemedBodyRightShadowPixels - 2, 0)
                                End If
                            End If
                            If iTabData.FixedWidth <> 0 Then
                                .Right = .Left + iTabData.FixedWidth
                            End If
                        End With
                        mTabData(t) = iTabData
                    End If
                End If
            Next t
        Next iRow
    Else
        For iRow = 0 To mRows - 1
            For t = 0 To mTabs - 1
                If mTabData(t).Visible Then
                    If mTabData(t).RowPos = iRow Then
                        iTabData = mTabData(t)
                        With iTabData.TabRect
                            If t = mTab Then
                                If (iTabExtraHeight > 0) And mHighlightAddExtraHeightSelectedTab Then
                                    .Top = (mRows - 1) * iTabHeight - iTabExtraHeight
                                Else
                                    .Top = (mRows - 1) * iTabHeight '+ IIf(iTabData.RowPos = 0, iTabExtraHeight, 0)
                                End If
                                If iTabExtraHeightIsNeeded Then .Top = .Top + iTabExtraHeight
                                .Bottom = .Top + iTabHeight + IIf(mHighlightAddExtraHeightSelectedTab, iTabExtraHeight, 0)
                                If mHighlightFlatBarWithGrip Or mHighlightFlatBarWithGripSelectedTab Then
                                    If mTabOrientation = ssTabOrientationBottom Then
                                        If (mFlatBarPosition = ntBarPositionBottom) And (mFlatBarGripHeightDPIScaled > 0) Then
                                            .Top = .Top + mFlatBarGripHeightDPIScaled
                                            .Bottom = .Bottom + mFlatBarGripHeightDPIScaled
                                        End If
                                    Else
                                        If (mFlatBarPosition = ntBarPositionTop) And (mFlatBarGripHeightDPIScaled > 0) Then
                                            .Top = .Top + mFlatBarGripHeightDPIScaled
                                            .Bottom = .Bottom + mFlatBarGripHeightDPIScaled
                                        End If
                                    End If
                                End If
                            Else
                                If iTabData.Hovered And mHighlightAddExtraHeight And mAmbientUserMode Then
                                    .Top = mTabData(t).RowPos * iTabHeight - iTabExtraHeight
                                    If iTabExtraHeightIsNeeded Then .Top = .Top + iTabExtraHeight
                                    .Bottom = .Top + iTabHeight + IIf(mHighlightAddExtraHeight, iTabExtraHeight, 0)
                                Else
                                    .Top = mTabData(t).RowPos * iTabHeight
                                    If iTabExtraHeightIsNeeded Then .Top = .Top + iTabExtraHeight
                                    .Bottom = .Top + iTabHeight
                                End If
                                If mTabOrientation = ssTabOrientationBottom Then
                                    If mHighlightFlatBarWithGrip Or mHighlightFlatBarWithGripSelectedTab Then
                                        If (mFlatBarPosition = ntBarPositionBottom) And (mFlatBarGripHeightDPIScaled > 0) Then
                                            .Top = .Top + mFlatBarGripHeightDPIScaled
                                            .Bottom = .Bottom + mFlatBarGripHeightDPIScaled
                                        End If
                                    End If
                                Else
                                    If mHighlightFlatBarWithGrip Or mHighlightFlatBarWithGripSelectedTab Then
                                        If (mFlatBarPosition = ntBarPositionTop) And (mFlatBarGripHeightDPIScaled > 0) Then
                                            .Top = .Top + mFlatBarGripHeightDPIScaled
                                            .Bottom = .Bottom + mFlatBarGripHeightDPIScaled
                                        End If
                                    End If
                                End If
                            End If
                            If (mTabWidthStyle2 = ntTWFixed) Then
                                .Left = (iTabData.PosH - 1) * IIf(mControlIsThemed, iTabWidth, Round(iTabWidth)) + iRowPerspectiveSpace * (mRows - mTabData(t).RowPos - 1) + 1 + (iTabData.PosH - 1) * mTabSeparation2 + 1
                                .Right = .Left + iTabWidth - 1 '- mTabSeparation2 ' no volver a sacar el -1!!
                            Else
                                If iTabData.LeftTab Then
                                    iTabLeft = IIf(((mTabWidthStyle2 = ntTWStretchToFill) Or (mTabWidthStyle2 = ntTWTabCaptionWidthFillRows)) And (Not (mAppearanceIsFlat Or (mStyle = ntStyleWindows))), 0, 1) + iRowPerspectiveSpace * (mRows - mTabData(t).RowPos - 1) + 1
                                Else
                                    iTabLeft = iTabLeft + mTabSeparation2
                                End If
                                .Left = iTabLeft
                                If (mTabWidthStyle2 = ntTWTabStripEmulation) Or (mTabWidthStyle2 = ntTWStretchToFill) Then
                                    .Right = .Left + iTabData.Width + 9
                                Else
                                    .Right = .Left + iTabData.IconAndCaptionWidth + 9
                                End If
                                iTabLeft = .Right + 1
                            End If
                            If iTabData.RightTab Then
                                iLng = iScaleWidthForTabs - iRowPerspectiveSpace * mTabData(t).RowPos - 1
                                If mAppearanceIsPP Then
                                    iLng = iLng - 2
                                    If mControlIsThemed Then
                                        If ((mTabWidthStyle2 <> ntTWTabStripEmulation) And (mTabWidthStyle2 <> ntTWStretchToFill)) Or iTabData.Selected Then
                                            iLng = iLng - 1
                                        End If
                                    End If
                                End If
                                If t = mTab Then
                                    If mControlIsThemed Then
                                        iLng = iLng + 1
                                    End If
                                End If
                                If Abs(.Right - iLng) < 6 Then
                                    .Right = iLng - IIf(mControlIsThemed, mThemedBodyRightShadowPixels - 2, 0)
                                End If
                            End If
                            If iTabData.FixedWidth <> 0 Then
                                .Right = .Left + iTabData.FixedWidth
                                iTabLeft = .Right + 1
                            End If
                        End With
                        mTabData(t) = iTabData
                    End If
                End If
            Next t
        Next iRow
    End If
    
    For t = 0 To mTabs - 1
        If mTabData(t).Visible Then
            If mTabData(t).PosH > 1 Then
                If mTabData(t).TabRect.Left <= mTabData(t - 1).TabRect.Right Then
                    iLng = t - 1
                    Do Until mTabData(iLng).Visible = True
                        iLng = iLng - 1
                        If iLng < 0 Then Exit Do
                    Loop
                    If iLng >= 0 Then
                        mTabData(t).TabRect.Left = mTabData(iLng).TabRect.Right + 1
                    End If
                End If
            End If
            If mTabData(t).RightTab Then
                If mTabData(t).RowPos > -1 Then
                    If mTabData(t).TabRect.Right > mRightMostTabsRightPos(mTabData(t).RowPos) Then
                        mRightMostTabsRightPos(mTabData(t).RowPos) = mTabData(t).TabRect.Right
                    End If
                End If
            End If
        End If
    Next t
    
    iLng = 0
    For iRow = 0 To mRows - 1
        For t = 0 To mTabs - 1
            If mTabData(t).RowPos = iRow Then
                If mTabData(t).TabRect.Left > (iLng - 2) Then
                    mTabData(t).TopTab = True
                End If
                If mTabData(t).RightTab Then
                    iLng = mTabData(t).TabRect.Right
                End If
            End If
        Next t
    Next iRow
    
    ' gap between tabs correction
    iLng = 0
    For iRow = 0 To mRows - 1
        For t = 0 To mTabs - 1
            If mTabData(t).Visible Then
                If mTabData(t).RowPos = iRow Then
                    If mTabData(t).TabRect.Left > (iLng + 1 + mTabSeparationDPIScaled) Then
                        If Not mTabData(t).LeftTab Then
                            mTabData(t).TabRect.Left = iLng + 1 + mTabSeparationDPIScaled
                        End If
                    End If
                    If mTabData(t).RightTab Then
                        iLng = 0
                    Else
                        iLng = mTabData(t).TabRect.Right
                    End If
                End If
            End If
        Next t
    Next iRow
    
    If Not mRedraw Then Exit Sub
    mNeedToDraw = False

    ' Do the draw
    
    ' How the "light" need to come according to TabOrientation (because the image later will be rotated). Note: in Windows the llight comes from top-left, and shadows are in bottom right.
    ' Top: from top-left
    ' Left: from top-right
    ' Right: from bottom-left
    ' Bottom: from bottom-left

    ' Do the draw
    picDraw.Width = iScaleWidth
    picDraw.Height = iScaleHeight
    
    If picDraw.BackColor <> mBackColorTabs Then
        picDraw.BackColor = mBackColorTabs ' the pic backcolor determines the focusrect color
    End If
    picDraw.Cls
    
    ' BackColor
    picDraw.Line (0, 0)-(iScaleWidth, iScaleHeight), IIf(mBackStyle = ntOpaque, mBackColor, mBackColorSelectedTab2), BF
    
    ' shadow is at the bottom and all need to be shifted
    If ((mTabOrientation = ssTabOrientationLeft) Or (mTabOrientation = ntTabOrientationLeftHorizontal)) And mControlIsThemed Then
        For t = 0 To mTabs - 1
            mTabData(t).TabRect.Left = mTabData(t).TabRect.Left + mThemedBodyRightShadowPixels
            mTabData(t).TabRect.Right = mTabData(t).TabRect.Right + mThemedBodyRightShadowPixels
        Next t
    End If
    
    ' draw inactive tabs
    For iRow = 0 To mRows - 1
        For t = 0 To mTabs - 1
            If mTabData(t).Visible Then
                If mTabData(t).RowPos = iRow Then
                    If t <> mTab Then
                        If mTabData(t).RightTab And Not (mTabData(t).RowPos = mRows - 1) Then
                            iLng = 4
                            If mAppearanceIsPP Then
                                iLng = iLng + 2 + IIf(mControlIsThemed, mThemedBodyRightShadowPixels - 2, 0)
                            End If
                            If ((mTabWidthStyle2 <> ntTWTabStripEmulation) And (mTabWidthStyle2 <> ntTWStretchToFill)) Or mShowsRowsPerspective2 Then
                                DrawInactiveBodyPart iRowPerspectiveSpace * (mRows - mTabData(t).RowPos - 1) + 3, mTabData(t).TabRect.Bottom + 5, mClientWidth - iLng, CLng(mClientHeight), iLng, mTabData(t).RowPos, t
                            End If
                        End If
                        If mAppearanceIsPP Then
                            mTabData(t).TabRect.Top = mTabData(t).TabRect.Top + 2
                        End If
                        DrawTab t
                        DrawTabPicureAndCaption t
                    End If
                End If
            End If
        Next t
    Next iRow
    
    ' Draw body (body: client area)
    DrawBody iScaleHeight
    
    ' Draw active tab
    If mAppearanceIsPP Then
        mTabData(mTab).TabRect.Left = mTabData(mTab).TabRect.Left - 2
        mTabData(mTab).TabRect.Right = mTabData(mTab).TabRect.Right + 2
    ElseIf (mTabAppearance2 = ntTATabbedDialog) Or (mTabAppearance2 = ntTATabbedDialogRounded) Then
        mTabData(mTab).TabRect.Left = mTabData(mTab).TabRect.Left - 1
    End If
    DrawTab CLng(mTab)
    DrawTabPicureAndCaption CLng(mTab)
    
    If mRows > 0 Then
        If mHighlightFlatBar And (mFlatBarPosition = ntBarPositionBottom) And (mFlatBarHeightDPIScaled > 0) Then
            For t = 0 To mTabs - 1
                If mTabData(t).Visible Then
                    If t <> mTab Then
                        If mTabData(t).RowPos < (mRows - 1) Then
                            If mTabData(t).Hovered Then
                                DrawTab t
                                DrawTabPicureAndCaption t
                                Exit For
                            End If
                        End If
                    End If
                End If
            Next t
        End If
    End If
    
    If mTabOrientation = ssTabOrientationBottom Then
        ' grip over the body
        If mHighlightFlatBar And (mFlatBarGripHeightDPIScaled > 0) Then
            For t = 0 To mTabs - 1
                If mTabData(t).Hovered And mAmbientUserMode Then
                    Const cEpsilon As Single = 0.499
                    Dim iTriangle(2) As POINTAPI
                    
                    With mTabData(t).TabRect
                        ' top point
                        iTriangle(0).X = (.Left + .Right) / 2 + cEpsilon
                        iTriangle(0).Y = .Bottom + 2 + mFlatBarGripHeightDPIScaled
                        ' left point
                        iTriangle(1).X = (.Left + .Right) / 2 - mFlatBarGripHeightDPIScaled + cEpsilon
                        iTriangle(1).Y = .Bottom + 2
                        ' right point
                        iTriangle(2).X = (.Left + .Right) / 2 + mFlatBarGripHeightDPIScaled + cEpsilon
                        iTriangle(2).Y = .Bottom + 2
                        If (Not IsEmpty(mTabData(t).CustomColors.FlatBarColorInactive)) Or (Not IsEmpty(mTabData(t).CustomColors.FlatBarColorHighlight)) Then
                            DrawTriangle iTriangle, mTabData(t).CustomColors.FlatBarGlowColor
                        Else
                            DrawTriangle iTriangle, mFlatBarGlowColor
                        End If
                    End With
                    Exit For
                End If
            Next
        End If
    End If
    
    mEndOfTabs = 0
    For t = 0 To mTabs - 1
        If mTabData(t).Visible Then
            If mTabData(t).TabRect.Right > mEndOfTabs Then
                mEndOfTabs = mTabData(t).TabRect.Right
            End If
        End If
    Next t
    mEndOfTabs = mEndOfTabs + 1
    
    Select Case mTabOrientation
        Case ssTabOrientationTop
            'BitBlt UserControl.hDC, 0, 0, iScaleWidth, iScaleHeight, picDraw.hDC, 0, 0, vbSrcCopy
            Set UserControl.Picture = picDraw.Image
        Case ssTabOrientationBottom
            UserControl.PaintPicture picDraw.Image, 0, iScaleHeight - 1, iScaleWidth, -iScaleHeight
            Set UserControl.Picture = UserControl.Image
            UserControl.Cls
        Case ssTabOrientationLeft, ntTabOrientationLeftHorizontal
            RotatePic picDraw, picRotate, nt90DegreesCounterClockWise
            'BitBlt UserControl.hDC, 0, 0, mScaleWidth, mScaleHeight, picRotate.hDC, 0, 0, vbSrcCopy
            Set UserControl.Picture = picRotate.Image
        Case Else ' ssTabOrientationRight, ntTabOrientationRightHorizontal
            RotatePic picDraw, picRotate, nt90DegreesClockWise
            'BitBlt UserControl.hDC, 0, 0, mScaleWidth, mScaleHeight, picRotate.hDC, 0, 0, vbSrcCopy
            Set UserControl.Picture = picRotate.Image
    End Select
    iAlreadyNeedToBePainted = GetUpdateRect(mUserControlHwnd, iTmpRect, 0&) <> 0&
    picDraw.Cls
    
    ' to avoid flickering on windowless contained controls, if not changed, validate the tab body area
    If (Not mBodyReset) Then
        If Not iAlreadyNeedToBePainted Then
            GetClientRect mUserControlHwnd, iTmpRect
            If mTabOrientation = ssTabOrientationTop Then
                iTmpRect.Top = mClientStart + 3
            ElseIf mTabOrientation = ssTabOrientationBottom Then
                iTmpRect.Bottom = iTmpRect.Bottom - mClientStart - 3
            ElseIf (mTabOrientation = ssTabOrientationLeft) Or (mTabOrientation = ntTabOrientationLeftHorizontal) Then
                iTmpRect.Left = mClientStart + 3
            ElseIf (mTabOrientation = ssTabOrientationRight) Or (mTabOrientation = ntTabOrientationRightHorizontal) Then
                iTmpRect.Right = iTmpRect.Right - mClientStart - 3
            End If
            ValidateRect mUserControlHwnd, iTmpRect
        End If
    End If
    mBodyReset = False
    
    ' rotate caption RECTs according to TabOrientation
    If mTabOrientation = ssTabOrientationBottom Then
        For t = 0 To mTabs - 1
            iTabData = mTabData(t)
            If iTabData.Visible Then
                With iTabData.TabRect
                    iLng = .Top - 2
                    .Top = iScaleHeight - 3 - .Bottom
                    .Bottom = iScaleHeight - 3 - iLng
                End With
                With iTabData.IconRect
                    iLng = .Top - 2
                    .Top = iScaleHeight - 3 - .Bottom
                    .Bottom = iScaleHeight - 3 - iLng
                End With
            End If
            mTabData(t) = iTabData
        Next t
    ElseIf (mTabOrientation = ssTabOrientationLeft) Or (mTabOrientation = ntTabOrientationLeftHorizontal) Then
        For t = 0 To mTabs - 1
            iTabData = mTabData(t)
            If iTabData.Visible Then
                With iTabData.TabRect
                    iTmpRect.Top = .Top
                    iTmpRect.Bottom = .Bottom
                    iTmpRect.Left = .Left
                    iTmpRect.Right = .Right
                    .Top = iScaleWidth - iTmpRect.Right
                    .Bottom = .Top + iTmpRect.Right - iTmpRect.Left
                    .Left = iTmpRect.Top
                    .Right = .Left + iTmpRect.Bottom - iTmpRect.Top
                End With
                With iTabData.IconRect
                    iTmpRect.Top = .Top
                    iTmpRect.Bottom = .Bottom
                    iTmpRect.Left = .Left
                    iTmpRect.Right = .Right
                    .Top = iScaleWidth - iTmpRect.Right
                    .Bottom = .Top + iTmpRect.Right - iTmpRect.Left
                    .Left = iTmpRect.Top
                    .Right = .Left + iTmpRect.Bottom - iTmpRect.Top
                End With
            End If
            mTabData(t) = iTabData
        Next t
    ElseIf (mTabOrientation = ssTabOrientationRight) Or (mTabOrientation = ntTabOrientationRightHorizontal) Then
        For t = 0 To mTabs - 1
            iTabData = mTabData(t)
            If iTabData.Visible Then
                With iTabData.TabRect
                    iTmpRect.Top = .Top
                    iTmpRect.Bottom = .Bottom
                    iTmpRect.Left = .Left
                    iTmpRect.Right = .Right
                    .Top = iTmpRect.Left
                    .Bottom = .Top + iTmpRect.Right - iTmpRect.Left
                    .Left = iScaleHeight - iTmpRect.Bottom
                    .Right = .Left + iTmpRect.Bottom - iTmpRect.Top
                End With
                With iTabData.IconRect
                    iTmpRect.Top = .Top
                    iTmpRect.Bottom = .Bottom
                    iTmpRect.Left = .Left
                    iTmpRect.Right = .Right
                    .Top = iTmpRect.Left
                    .Bottom = .Top + iTmpRect.Right - iTmpRect.Left
                    .Left = iScaleHeight - iTmpRect.Bottom
                    .Right = .Left + iTmpRect.Bottom - iTmpRect.Top
                End With
            End If
            mTabData(t) = iTabData
        Next t
    End If
    If mRows <> mRows_Prev Then
        RaiseEvent RowsChange
    End If
    mRows_Prev = mRows
    If ((mClientStart <> mClientStart_Prev) And (mAutoRelocateControls = ntRelocateAlways) Or (mTabOrientation <> mTabOrientation_Prev) And (mAutoRelocateControls > 0)) And (mTabOrientation_Prev <> -1) Then
        RearrangeContainedControlsPositions
    End If
    mClientStart_Prev = mClientStart
    mTabOrientation_Prev = mTabOrientation
    
    If mBackStyle = ntOpaque Then
        Set UserControl.MaskPicture = Nothing
        tmrCheckContainedControlsAdditionDesignTime.Enabled = False
        tmrCheckContainedControlsAdditionDesignTime.Interval = 1
    Else
        tmrCheckContainedControlsAdditionDesignTime.Interval = 50
        If Not mAmbientUserMode Then tmrCheckContainedControlsAdditionDesignTime.Enabled = True
        picAux.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
        Set picAux.Picture = UserControl.Picture
        
        Dim iCtl As Object
        Dim iLeft As Long
        Dim iWidth As Long
        
        On Error Resume Next
        For Each iCtl In UserControlContainedControls
            iLeft = -mLeftOffsetToHide
            iLeft = iCtl.Left
            If iLeft > -mLeftOffsetToHide Then
                iWidth = -1
                iWidth = iCtl.Width
                If iWidth <> -1 Then
                    picAux.Line (ScaleX(iLeft, vbTwips, vbPixels), ScaleY(iCtl.Top, vbTwips, vbPixels))-(ScaleX(iLeft + iWidth, vbTwips, vbPixels), ScaleY(iCtl.Top + iCtl.Height, vbTwips, vbPixels)), mBackColorSelectedTab2 Xor &H1, BF
                End If
            End If
        Next
        On Error GoTo 0
        If Not mAmbientUserMode Then mLastContainedControlsPositionsStr = GetContainedControlsPositionsStr
        Set UserControl.MaskPicture = picAux.Image
        Set picAux.Picture = Nothing
        picAux.Cls
    End If
    
    RaiseEvent_Resize
    
    If mSubclassControlsPaintingPending Then SubclassControlsPainting
    
    If lblTDILabel.Visible Then
        lblTDILabel.Move ScaleX(mClientRect.Left, vbPixels, UserControl.ScaleMode), ScaleY(mClientRect.Top, vbPixels, UserControl.ScaleMode), ScaleX(mClientRect.Right - mClientRect.Left, vbPixels, UserControl.ScaleMode), ScaleY(mClientRect.Bottom - mClientRect.Top, vbPixels, UserControl.ScaleMode)
    End If

    
TheExit:
    SetScaleMode
    If mTheme <> 0 Then
        CloseThemeData mTheme
        mTheme = 0
    End If
    mDrawing = False
End Sub

Private Sub SetScaleMode()
    Dim iSW As Long
    
    UserControl.ScaleMode = vbTwips
    If mRightToLeft Then
        iSW = UserControl.ScaleWidth
        UserControl.Scale (iSW, 0)-(0, UserControl.ScaleHeight)
    End If
End Sub

Private Sub RaiseEvent_Resize()
    Static sInside As Boolean
    
    If UserControl.EventsFrozen Then
        mResizeEventPending = True
    Else
        If mDrawMessagePosted Or tmrDraw.Enabled Then
            Draw
        ElseIf Not sInside Then
            sInside = True
            mResizeEventPending = False
            If mTDIMode = ntTDIModeForms Then TDIResizeFormContainers
            If (mClientRect_Prev.Left <> mClientRect.Left) Or (mClientRect_Prev.Top <> mClientRect.Top) Or (mClientRect_Prev.Right <> mClientRect.Right) Or (mClientRect_Prev.Bottom <> mClientRect.Bottom) Then
                RaiseEvent Resize
                mClientRect_Prev.Left = mClientRect.Left
                mClientRect_Prev.Top = mClientRect.Top
                mClientRect_Prev.Right = mClientRect.Right
                mClientRect_Prev.Bottom = mClientRect.Bottom
            End If
            sInside = False
        End If
    End If
End Sub

Private Sub TDIResizeFormContainers()
    Dim c As Long
    Dim iSM As Long
    
    iSM = UserControl.ScaleMode
    UserControl.ScaleMode = vbTwips
    For c = 1 To mTabs - 1
        If mTabData(c).Visible Then
            If (mTabData(c).Data >= picTDIFormContainer.LBound) And (mTabData(c).Data <= picTDIFormContainer.UBound) Then
                If IsIconic(mFormHwnd) = 0 Then
                    On Error Resume Next
                    picTDIFormContainer(mTabData(c).Data).Move ClientLeft, ClientTop, ClientWidth, ClientHeight
                    MoveWindow mTDIModeFormsFormData_FormHwnd(mTabData(c).Data), 0, 0, mClientRect.Right - mClientRect.Left + 2, mClientRect.Bottom - mClientRect.Top + 3, 1
                    On Error GoTo 0
                End If
            End If
        End If
    Next
    If iSM = vbUser Then
        SetScaleMode
    Else
        UserControl.ScaleMode = iSM
    End If
End Sub

Private Sub DrawTab(ByVal nTab As Long)
    Dim iCurv As Long
    Dim iLeftOffset As Long
    Dim iRightOffset As Long
    Dim iTopOffset As Long
    Dim iBottomOffset As Long
    Dim iHighlighted As Boolean
    Dim iTabData As T_TabData
    Dim iExtI As Long
    Dim iActive As Boolean
    Dim iState As Long
    Dim iTRect As RECT
    Dim iTRect2 As RECT
    Dim iPartId As Long
    Dim iLeft As Long
    Dim iTop As Long
    Dim iRoundedTabs As Boolean
    Dim iBackColorTabs2 As Long
    Dim iBackColorTabs3 As Long
    
    Dim i3DShadow As Long
    Dim i3DDKShadow As Long
    Dim i3DHighlight As Long
    Dim i3DHighlightH As Long
    Dim i3DHighlightV As Long
    Dim i3DShadowV As Long
    Dim iHighlightColor As Long
    Dim iHighlightGradient As NTHighlightGradientConstants
    Dim iLng As Long
    Dim iColor As Long
    
    Dim iTabColor_BackColorTab As Long
    Dim iTabColor_BackColorSelectedTab As Long
    Dim iTabColor_FlatBarColorInactive As Long
    Dim iTabColor_FlatBarGlowColor As Long
    Dim iTabColor_FlatBarColorSelectedTab As Long
    
    iTabData = mTabData(nTab)
    iActive = iTabData.Selected
    iRoundedTabs = (mTabAppearance2 = ntTAPropertyPageRounded) Or (mTabAppearance2 = ntTATabbedDialogRounded) Or ((mTabAppearance2 = ntTAFlat) And ((mFlatRoundnessTopDPIScaled > 0) Or (mFlatRoundnessTabsDPIScaled > 0)))
    
    If IsEmpty(iTabData.CustomColors.BackColorTab) Or (mHandleHighContrastTheme And mHighContrastThemeOn) Then
        iTabColor_BackColorTab = mBackColorTabs2
    Else
        iTabColor_BackColorTab = iTabData.CustomColors.BackColorTab
    End If
    iTabColor_BackColorSelectedTab = mBackColorSelectedTab2
    If IsEmpty(iTabData.CustomColors.FlatBarColorInactive) Then
        iTabColor_FlatBarColorInactive = mFlatBarColorInactive
    Else
        iTabColor_FlatBarColorInactive = iTabData.CustomColors.FlatBarColorInactive
    End If
    If (IsEmpty(mTabData(nTab).CustomColors.FlatBarColorInactive)) And (IsEmpty(mTabData(nTab).CustomColors.FlatBarColorHighlight)) Then
        iTabColor_FlatBarGlowColor = mFlatBarGlowColor
    Else
        iTabColor_FlatBarGlowColor = iTabData.CustomColors.FlatBarGlowColor
    End If
    If IsEmpty(iTabData.CustomColors.FlatBarColorSelectedTab) Then
        iTabColor_FlatBarColorSelectedTab = mFlatBarColorSelectedTab
    Else
        iTabColor_FlatBarColorSelectedTab = iTabData.CustomColors.FlatBarColorSelectedTab
    End If
    
    If iActive Then
        iHighlighted = ((mHighlightGradientSelectedTab <> ntGradientNone) Or mAppearanceIsFlat Or mControlIsThemed) And iTabData.Enabled
        iBackColorTabs2 = iTabColor_BackColorSelectedTab
        i3DDKShadow = m3DDKShadow_Sel
        i3DHighlightH = m3DHighlightH_Sel
        i3DHighlightV = m3DHighlightV_Sel
        i3DShadowV = m3DShadowV_Sel
        i3DShadow = m3DShadow_Sel
        i3DHighlight = m3DHighlight_Sel
        iHighlightColor = mGlowColor_Sel
        If mBackStyle <> ntOpaque Then iHighlightColor = iHighlightColor Xor 65538
        iHighlightGradient = mHighlightGradientSelectedTab
        If DraggingATab Then
            If mTabOrientation = ssTabOrientationBottom Then
                iTabData.TabRect.Left = iTabData.TabRect.Left + mMouseX2 - mMouseX
                iTabData.TabRect.Right = iTabData.TabRect.Right + mMouseX2 - mMouseX
                iTabData.TabRect.Top = iTabData.TabRect.Top - mMouseY2 + mMouseY
                iTabData.TabRect.Bottom = iTabData.TabRect.Bottom - mMouseY2 + mMouseY
            ElseIf (mTabOrientation = ssTabOrientationLeft) Or (mTabOrientation = ntTabOrientationLeftHorizontal) Then
                iTabData.TabRect.Left = iTabData.TabRect.Left - mMouseY2 + mMouseY
                iTabData.TabRect.Right = iTabData.TabRect.Right - mMouseY2 + mMouseY
                iTabData.TabRect.Top = iTabData.TabRect.Top + mMouseX2 - mMouseX
                iTabData.TabRect.Bottom = iTabData.TabRect.Bottom + mMouseX2 - mMouseX
            ElseIf (mTabOrientation = ssTabOrientationRight) Or (mTabOrientation = ntTabOrientationRightHorizontal) Then
                iTabData.TabRect.Left = iTabData.TabRect.Left + mMouseY2 - mMouseY
                iTabData.TabRect.Right = iTabData.TabRect.Right + mMouseY2 - mMouseY
                iTabData.TabRect.Top = iTabData.TabRect.Top - mMouseX2 + mMouseX
                iTabData.TabRect.Bottom = iTabData.TabRect.Bottom - mMouseX2 + mMouseX
            Else
                iTabData.TabRect.Left = iTabData.TabRect.Left + mMouseX2 - mMouseX
                iTabData.TabRect.Right = iTabData.TabRect.Right + mMouseX2 - mMouseX
                iTabData.TabRect.Top = iTabData.TabRect.Top + mMouseY2 - mMouseY
                iTabData.TabRect.Bottom = iTabData.TabRect.Bottom + mMouseY2 - mMouseY
            End If
        End If
    Else
        iHighlighted = mAmbientUserMode And ((mHighlightGradient <> ntGradientNone) Or mAppearanceIsFlat Or mControlIsThemed) And iTabData.Hovered And (mEnabled Or (Not mAmbientUserMode)) And iTabData.Enabled
        If DraggingATab Then iHighlighted = False
        iBackColorTabs2 = iTabColor_BackColorTab
        i3DDKShadow = m3DDKShadow
        i3DHighlightH = m3DHighlightH
        i3DHighlightV = m3DHighlightV
        i3DShadowV = m3DShadowV
        i3DShadow = m3DShadow
        i3DHighlight = m3DHighlight
        If Not IsEmpty(iTabData.CustomColors.HighlightColor) Then
            iHighlightColor = iTabData.CustomColors.HighlightColor
        Else
            iHighlightColor = mGlowColor
        End If
        iHighlightGradient = mHighlightGradient
    End If
    iBackColorTabs3 = iTabColor_BackColorSelectedTab
    
    If mAppearanceIsFlat Then
        Dim iFlatBarTopColor As Long
        Dim iLineColor As Long
        Dim iFlatTabsSeparationLineColor As Long
        Dim iFlatBorderColor As Long
        Dim iFlatLeftRoundness As Long
        Dim iFlatRightRoundness As Long
        Dim iFlatLeftLineColor As Long
        Dim iFlatRightLineColor As Long
        Dim iFlatBarTopHeight As Long
        Dim iShowFlatBarBottom As Boolean
        Const cEpsilon As Single = 0.499
        Dim iFlatBarTopSet As Boolean
        Dim iFlatBarPosition As NTFlatBarPositionConstants
        Dim iHighlightFlatDrawBorder As Boolean
        Dim iHighlightFlatDrawBorder_Color As Long
        
        iFlatTabsSeparationLineColor = TranslatedColor(mFlatTabsSeparationLineColor)
        iFlatBorderColor = TranslatedColor(mFlatBorderColor)
        
        iFlatBarPosition = mFlatBarPosition
        If mTabOrientation = ssTabOrientationBottom Then
            If iFlatBarPosition = ntBarPositionTop Then
                iFlatBarPosition = ntBarPositionBottom
            Else
                iFlatBarPosition = ntBarPositionTop
            End If
        End If
        
        If iHighlighted And (Not iActive) Then
            If mHighlightFlatBar Or mHighlightFlatBarSelectedTab Then
                If mHighlightFlatBar Or ((iHighlightGradient <> ntGradientNone) And (iTabColor_FlatBarGlowColor = iTabColor_BackColorTab)) Then
                    iFlatBarTopColor = iTabColor_FlatBarGlowColor
                Else
                    iFlatBarTopColor = iTabColor_FlatBarColorInactive
                End If
            Else
                iFlatBarTopHeight = 1
                If (iHighlightGradient <> ntGradientNone) Then
                    If iHighlightGradient = ntGradientDouble Then
                        iFlatBarTopColor = iBackColorTabs2
                    Else
                        iFlatBarTopColor = iHighlightColor
                    End If
                Else
                    If mFlatBorderMode = ntBorderTabs Then
                        iFlatBarTopColor = iFlatBorderColor
                    Else
                        iFlatBarTopColor = iFlatTabsSeparationLineColor
                    End If
                End If
                iFlatBarTopSet = True
            End If
        Else
            If iActive Then
                If mHighlightFlatBarSelectedTab Then
                    iFlatBarTopColor = iTabColor_FlatBarColorSelectedTab
                Else
                    If mHighlightGradient <> ntGradientNone Then
                        iFlatBarTopHeight = 1
                        iFlatBarTopColor = iFlatBorderColor '  iHighlightColor
                        iFlatBarTopSet = True
                    Else
                        If mHighlightFlatBar Then
                            iFlatBarTopColor = iTabColor_FlatBarColorInactive
                        Else
                            iFlatBarTopHeight = 1
                            iFlatBarTopColor = iFlatBorderColor
                            iFlatBarTopSet = True
                        End If
                    End If
                End If
            Else
                If mHighlightFlatBar Or mHighlightFlatBarSelectedTab Then
                    iFlatBarTopColor = iTabColor_FlatBarColorInactive
                Else
                    iFlatBarTopHeight = 1
                    If mFlatBorderMode = ntBorderTabs Then
                        iFlatBarTopColor = iFlatBorderColor ' iFlatTabsSeparationLineColor
                    Else
                        iFlatBarTopColor = iFlatTabsSeparationLineColor
                    End If
                    iFlatBarTopSet = True
                End If
            End If
        End If
        If Not iFlatBarTopSet Then
            If iFlatBarPosition = ntBarPositionTop Then
                iFlatBarTopHeight = mFlatBarHeightDPIScaled
                If (iFlatBarTopHeight = 0) Then
                    iFlatBarTopHeight = 1
                    If (mFlatBorderMode = ntBorderTabs) Or iActive Then
                        iFlatBarTopColor = iFlatBorderColor
                    Else
                        iFlatBarTopColor = iFlatTabsSeparationLineColor
                    End If
                End If
            Else
                iFlatBarTopHeight = 1
                iFlatBarTopColor = iFlatBorderColor '  iHighlightColor
            End If
        End If
        
        If iActive Then
            If mHighlightFlatDrawBorderSelectedTab Then
                iHighlightFlatDrawBorder = True
                If IsEmpty(iTabData.CustomColors.FlatTabBorderColorSelectedTab) Then
                    iHighlightFlatDrawBorder_Color = TranslatedColor(mFlatTabBorderColorSelectedTab)
                Else
                    iHighlightFlatDrawBorder_Color = TranslatedColor(iTabData.CustomColors.FlatTabBorderColorSelectedTab)
                End If
                iFlatBarTopColor = iHighlightFlatDrawBorder_Color
            End If
        Else
            If mHighlightFlatDrawBorder And iHighlighted Then
                iHighlightFlatDrawBorder = True
                If IsEmpty(iTabData.CustomColors.FlatTabBorderColorHighlight) Then
                    iHighlightFlatDrawBorder_Color = TranslatedColor(mFlatTabBorderColorHighlight)
                Else
                    iHighlightFlatDrawBorder_Color = TranslatedColor(iTabData.CustomColors.FlatTabBorderColorHighlight)
                End If
                iFlatBarTopColor = iHighlightFlatDrawBorder_Color
            End If
        End If
        If iTabData.LeftTab And (Not iHighlightFlatDrawBorder) Then
            iFlatLeftRoundness = mFlatRoundnessTop2
            If mFlatRoundnessTabs2 > iFlatLeftRoundness Then
                iFlatLeftRoundness = mFlatRoundnessTabs2
            End If
        Else
            iFlatLeftRoundness = mFlatRoundnessTabs2
        End If
        If iTabData.RightTab And (Not iHighlightFlatDrawBorder) Then
            iFlatRightRoundness = mFlatRoundnessTop2
            If mFlatRoundnessTabs2 > iFlatRightRoundness Then
                iFlatRightRoundness = mFlatRoundnessTabs2
            End If
        Else
            iFlatRightRoundness = mFlatRoundnessTabs2
        End If
        If mAppearanceIsFlat Then
            If iRoundedTabs Then
                If (iFlatLeftRoundness > 0) Then
                    If Not ((Not (iTabData.LeftTab And ((mRows > 1) And (iTabData.RowPos <> 0)))) Or mShowsRowsPerspective2 Or (mFlatRoundnessTabsDPIScaled > 0)) Then
                        iFlatLeftRoundness = 0
                    End If
                End If
                If (iFlatRightRoundness > 0) Then
                    If Not ((Not (iTabData.RightTab And ((mRows > 1) And (iTabData.RowPos <> 0)))) Or mShowsRowsPerspective2 Or (mFlatRoundnessTabsDPIScaled > 0)) Then
                        iFlatRightRoundness = 0
                    End If
                End If
            End If
        End If
        If iTabData.LeftTab Then
            iFlatLeftLineColor = IIf((mFlatBorderMode = ntBorderTabs) Or iActive, iFlatBorderColor, iFlatTabsSeparationLineColor) ' IIf(mFlatBorderMode = ntBorderTabs, iFlatBorderColor, iFlatTabsSeparationLineColor)
        Else
            iFlatLeftLineColor = IIf((mFlatBorderMode = ntBorderSelectedTab) And iActive Or (mFlatBorderMode = ntBorderTabs) And (mTabSeparationDPIScaled > 0), iFlatBorderColor, iFlatTabsSeparationLineColor)
        End If
        If iTabData.RightTab Then
            If (iTabData.RowPos = mRows - 1) Then
                iFlatRightLineColor = IIf((mFlatBorderMode = ntBorderTabs) Or iActive Or (mFlatBorderMode = ntBorderTabs) And (mTabSeparationDPIScaled > 0), iFlatBorderColor, iFlatTabsSeparationLineColor)
            Else
                iFlatRightLineColor = IIf(mFlatBorderMode = ntBorderTabs, iFlatBorderColor, iFlatTabsSeparationLineColor)
            End If
        Else
            iFlatRightLineColor = IIf(iActive And (mFlatBorderMode = ntBorderSelectedTab) Or (mFlatBorderMode = ntBorderTabs) And (mTabSeparationDPIScaled > 0), iFlatBorderColor, iFlatTabsSeparationLineColor)
        End If
        If IIf(nTab = mTab, mHighlightFlatBarSelectedTab, mHighlightFlatBar) And (iFlatBarPosition = ntBarPositionBottom) Then
            iShowFlatBarBottom = (mFlatBarHeightDPIScaled > 0)
        End If
    End If
    
    With iTabData.TabRect
        If mControlIsThemed Then
            If Not iTabData.Enabled Then
                iState = TIS_DISABLED
            ElseIf ((iActive And ControlHasFocus) And (Not mShowFocusRect) And mAmbientUserMode) Or iActive And ((mTabOrientation = ssTabOrientationBottom) Or (mTabOrientation = ssTabOrientationRight) Or (mTabOrientation = ntTabOrientationRightHorizontal)) Then
                iState = TIS_SELECTED ' I had to put TIS_SELECTED instead of TIS_FOCUSED before
            ElseIf iActive Then
                iState = TIS_SELECTED
            ElseIf iHighlighted Then
                iState = TIS_HOT
            Else
                iState = TIS_NORMAL
            End If
            
            If mTabSeparation2 = 0 Then
                iPartId = IIf(iTabData.RightTab, TABP_TABITEMRightEDGE, IIf(iTabData.LeftTab, TABP_TABITEMLEFTEDGE, TABP_TABITEM))
            Else
                iPartId = TABP_TABITEMLEFTEDGE
            End If
            If (mBackColor = vbButtonFace) And (Not (iTabData.RightTab Or (iState = TIS_FOCUSED)) Or (mTabSeparation2 > 0)) Then
                iTRect.Top = .Top
                iTRect.Left = .Left
                iTRect.Right = .Right + 1
                iTRect.Bottom = .Bottom + 1
                If Not iActive Then
                    If (mTabSeparation2 > 0) Then
                        If iTabData.RightTab Then
                            iTRect.Bottom = iTRect.Bottom + 1
                        End If
                    End If
                End If
                If mTabData(nTab).RowPos <> mRows - 1 Then
                    iTRect.Bottom = iTRect.Bottom + 4
                End If
                iTRect2 = iTRect
                iTRect2.Bottom = iTRect.Bottom + 1
                DrawThemeBackground mTheme, picDraw.hDC, iPartId, iState, iTRect2, iTRect
            Else
                iTRect.Left = 0
                iTRect.Top = 0
                iTRect.Bottom = .Bottom - .Top
                iTRect.Bottom = iTRect.Bottom + 1
                If (mTabOrientation = ssTabOrientationBottom) Or (mTabOrientation = ssTabOrientationRight) Or (mTabOrientation = ntTabOrientationRightHorizontal) Then
                    iTRect.Bottom = iTRect.Bottom + 1
                End If
                If Not iActive Then
                    If iTabData.RightTab Then
                        iTRect.Bottom = iTRect.Bottom + 1
                    End If
                End If
                If mTabData(nTab).RowPos <> mRows - 1 Then
                    iTRect.Bottom = iTRect.Bottom + 4
                End If
                iTRect.Right = .Right - .Left + 1
                iLeft = .Left
                iTop = .Top
                On Error Resume Next
                picAux.Width = iTRect.Right
                picAux.Height = iTRect.Bottom
                picAux.Cls
                
                iTRect2 = iTRect
                iTRect2.Bottom = iTRect.Bottom + 1
                
                DrawThemeBackground mTheme, picAux.hDC, iPartId, iState, iTRect2, iTRect
                SetThemedTabTransparentPixels iTabData.LeftTab, (iState = TIS_FOCUSED) Or iTabData.RightTab, (iTabData.TopTab Or (mTabSeparation2 > 0)) And Not (iState = TIS_SELECTED)
                Call TransparentBlt(picDraw.hDC, iLeft, iTop, iTRect.Right, iTRect.Bottom, picAux.hDC, 0, 0, iTRect.Right, iTRect.Bottom, cAuxTransparentColor)
                picAux.Cls
                On Error GoTo 0
            End If
        Else
            If iActive Then
                ' active tab background
                If mAppearanceIsPP Then
                    iExtI = 2
                    iLeftOffset = 1
                    iRightOffset = 1
                    iTopOffset = 0
                    iBottomOffset = 1
                    If iRoundedTabs Then
                        iCurv = 3
                    Else
                        iCurv = 2
                    End If
                ElseIf mAppearanceIsFlat Then
                    iBottomOffset = 3
                    If iTabData.LeftTab Then
                        iLeftOffset = -2
                    End If
                Else
                    iExtI = 1
                    If iRoundedTabs Then
                        iLeftOffset = -1
                    Else
                        iLeftOffset = 0
                    End If
                    iRightOffset = 0
                    iTopOffset = 0
                    iBottomOffset = 2
                    iCurv = 4
                End If
            Else
                ' inactive tab background
                iExtI = 6
                If mAppearanceIsPP Then
                    iLeftOffset = 0
                    iRightOffset = 0
                    iTopOffset = 0 '2
                    iBottomOffset = 5
                    If iRoundedTabs Then
                        iCurv = 3
                    Else
                        iCurv = 2
                    End If
                ElseIf mAppearanceIsFlat Then
                    iBottomOffset = 6
                    If iTabData.LeftTab Then
                        iLeftOffset = -2
                    End If
                    If iTabData.RowPos = (mRows - 1) Then
                        If iFlatLeftRoundness > iBottomOffset Then iBottomOffset = iFlatLeftRoundness
                        If iFlatRightRoundness > iBottomOffset Then iBottomOffset = iFlatRightRoundness
                    End If
                    If iTabData.RightTab Then
                        If Abs(mClientRect.Right - iTabData.TabRect.Right) > 5 Then
                            iBottomOffset = 11
                            If (iTabData.RowPos <> (mRows - 1)) And (iShowFlatBarBottom) Then
                                iBottomOffset = iBottomOffset - 2
                                iExtI = iExtI - 3
                            Else
                                iExtI = 0 ' -5
                            End If
                        End If
                    End If
                    If (iTabData.RowPos <> (mRows - 1)) And (iShowFlatBarBottom) Then
                        iBottomOffset = iBottomOffset - mFlatBarHeightDPIScaled - 2
                        iExtI = iExtI - 3
                    End If
                Else
                    If iRoundedTabs Then
                        iLeftOffset = -1
                    Else
                        iLeftOffset = 0
                    End If
                    iTopOffset = 0
                    iRightOffset = -1
                    iBottomOffset = 6
                    iCurv = 3
                End If
            End If
            If mBackStyle = ntOpaqueSelectedTab Then iBackColorTabs2 = iBackColorTabs2 Xor 257
            
            If mAppearanceIsFlat Then
                If iHighlighted Then
                    If iHighlightGradient = ntGradientSimple Then
                        Call FillCurvedGradient2(.Left + iLeftOffset + 1, .Top + iTopOffset, .Right + iRightOffset + IIf(iTabData.RightTab, 1, 0), (.Bottom), iHighlightColor, iBackColorTabs2, iFlatLeftRoundness, iFlatRightRoundness)
                        If iHighlightFlatDrawBorder Then
                            Call FillCurvedGradient2(.Left + iLeftOffset + 1, .Bottom, .Right + iRightOffset + IIf(iTabData.RightTab, 1, 0), (.Bottom + 1), iBackColorTabs2, iBackColorTabs2, 0, 0, iFlatLeftRoundness, iFlatRightRoundness)
                        Else
                            Call FillCurvedGradient2(.Left + iLeftOffset + 1, .Bottom, .Right + iRightOffset + IIf(iTabData.RightTab, 1, 0), (.Bottom + iBottomOffset), iBackColorTabs2, iBackColorTabs2, 0, 0)
                        End If
                    ElseIf iHighlightGradient = ntGradientDouble Then
                        Call FillCurvedGradient2(.Left + iLeftOffset + 1, .Top + iTopOffset, .Right + iRightOffset, (.Bottom + .Top + iTopOffset) / 2 + 2, iBackColorTabs2, iHighlightColor, iFlatLeftRoundness, iFlatRightRoundness)
                        Call FillCurvedGradient2(.Left + iLeftOffset + 1, (.Bottom + .Top + iTopOffset) / 2, .Right + iRightOffset, .Bottom, iHighlightColor, iBackColorTabs2, 0, 0)
                        If iHighlightFlatDrawBorder Then
                            Call FillCurvedGradient2(.Left + iLeftOffset + 1, .Bottom, .Right + iRightOffset, .Bottom + 1, iBackColorTabs2, iBackColorTabs2, 0, 0, iFlatLeftRoundness, iFlatRightRoundness)
                        Else
                            Call FillCurvedGradient2(.Left + iLeftOffset + 1, .Bottom, .Right + iRightOffset, .Bottom + iBottomOffset, iBackColorTabs2, iBackColorTabs2, 0, 0)
                        End If
                    ElseIf iHighlightGradient = ntGradientPlain Then
                        If iHighlightFlatDrawBorder Then
                            Call FillCurvedGradient2(.Left + iLeftOffset + 1, .Top + iTopOffset, .Right + iRightOffset + IIf(iTabData.RightTab, 1, 0), (.Bottom + 1), iHighlightColor, iHighlightColor, iFlatLeftRoundness, iFlatRightRoundness, iFlatLeftRoundness, iFlatRightRoundness)
                        Else
                            Call FillCurvedGradient2(.Left + iLeftOffset + 1, .Top + iTopOffset, .Right + iRightOffset + IIf(iTabData.RightTab, 1, 0), (.Bottom + iBottomOffset), iHighlightColor, iHighlightColor, iFlatLeftRoundness, iFlatRightRoundness)
                        End If
                    ElseIf iHighlightGradient = ntGradientNone Then
                        Call FillCurvedGradient2(.Left + iLeftOffset + 1, .Top + iTopOffset, .Right + iRightOffset + IIf(iTabData.RightTab, 1, 0), (.Bottom + iBottomOffset), iBackColorTabs2, iBackColorTabs2, iFlatLeftRoundness, iFlatRightRoundness)
                    End If
                Else
                    If iHighlightFlatDrawBorder Then
                        Call FillCurvedGradient2(.Left + iLeftOffset + 1, .Top + iTopOffset, .Right + iRightOffset, .Bottom + 1, iBackColorTabs2, iBackColorTabs2, iFlatLeftRoundness, iFlatRightRoundness, iFlatLeftRoundness, iFlatRightRoundness)
                    Else
                        Call FillCurvedGradient2(.Left + iLeftOffset + 1, .Top + iTopOffset, .Right + iRightOffset, .Bottom + iBottomOffset, iBackColorTabs2, iBackColorTabs2, iFlatLeftRoundness, iFlatRightRoundness)
                    End If
                End If
            Else
                If iHighlighted Then
                    If iHighlightGradient = ntGradientSimple Then
                        Call FillCurvedGradient(.Left + iLeftOffset + 1, .Top + iTopOffset, .Right + iRightOffset + IIf(iTabData.RightTab, 1, 0), (.Bottom), iHighlightColor, iBackColorTabs2, IIf(mAppearanceIsFlat, -1, iCurv), True, True)
                        Call FillCurvedGradient(.Left + iLeftOffset + 1, .Bottom, .Right + iRightOffset + IIf(iTabData.RightTab, 1, 0), (.Bottom + iBottomOffset), iBackColorTabs2, iBackColorTabs2, IIf(mAppearanceIsFlat, -1, iCurv), True, True)
                    ElseIf iHighlightGradient = ntGradientDouble Then
                        Call FillCurvedGradient(.Left + iLeftOffset + 1, .Top + iTopOffset, .Right + iRightOffset, (.Bottom + .Top + iTopOffset) / 2 + 2, iBackColorTabs2, iHighlightColor, IIf(mAppearanceIsFlat, -1, iCurv), True, True)
                        Call FillCurvedGradient(.Left + iLeftOffset + 1, (.Bottom + .Top + iTopOffset) / 2, .Right + iRightOffset, .Bottom, iHighlightColor, iBackColorTabs2, IIf(mAppearanceIsFlat, -1, iCurv), True, True)
                        Call FillCurvedGradient(.Left + iLeftOffset + 1, .Bottom, .Right + iRightOffset, .Bottom + iBottomOffset, iBackColorTabs2, iBackColorTabs2, IIf(mAppearanceIsFlat, -1, iCurv), True, True)
                    ElseIf iHighlightGradient = ntGradientPlain Then
                        Call FillCurvedGradient(.Left + iLeftOffset + 1, .Top + iTopOffset, .Right + iRightOffset + IIf(iTabData.RightTab, 1, 0), (.Bottom + iBottomOffset), iHighlightColor, iHighlightColor, iCurv, True, True)
                    ElseIf iHighlightGradient = ntGradientNone Then
                        Call FillCurvedGradient(.Left + iLeftOffset + 1, .Top + iTopOffset, .Right + iRightOffset + IIf(iTabData.RightTab, 1, 0), (.Bottom + iBottomOffset), iBackColorTabs2, iBackColorTabs2, iCurv, True, True)
                    End If
                Else
                    Call FillCurvedGradient(.Left + iLeftOffset + 1, .Top + iTopOffset, .Right + iRightOffset, .Bottom + iBottomOffset, iBackColorTabs2, iBackColorTabs2, iCurv, True, True)
                End If
            End If
            
            'top line
            If mAppearanceIsPP Then
                If iRoundedTabs Then
                    picDraw.Line (.Left + iLeftOffset + 2, .Top)-(.Right - 2, .Top), i3DHighlightH
                Else
                    picDraw.Line (.Left + iLeftOffset + 2, .Top)-(.Right - 1, .Top), i3DHighlightH
                End If
                If (mTabOrientation = ssTabOrientationBottom) Or (mTabOrientation = ssTabOrientationRight) Or (mTabOrientation = ntTabOrientationRightHorizontal) Then
                    If iRoundedTabs Then
                        picDraw.Line (.Left + iLeftOffset + 4, .Top - 1)-(.Right - 3, .Top - 1), i3DHighlightH
                    Else
                        picDraw.Line (.Left + iLeftOffset + 3, .Top - 1)-(.Right - 2, .Top - 1), i3DHighlightH
                    End If
                End If
            ElseIf mAppearanceIsFlat Then
                If (iFlatBarTopHeight > 0) Then
                    FillCurvedGradient2 .Left + iLeftOffset, .Top + iTopOffset, .Right + iRightOffset + IIf(iTabData.RightTab, 2, 0), .Top + iTopOffset + iFlatBarTopHeight, iFlatBarTopColor, iFlatBarTopColor, iFlatLeftRoundness, iFlatRightRoundness
                    If iHighlighted And (iFlatBarTopHeight > 1) Then
                        If IIf(iActive, mHighlightFlatBarWithGripSelectedTab, mHighlightFlatBarWithGrip) Then
                            Dim iTriangle(2) As POINTAPI
                            
                            If mFlatBarGripHeightDPIScaled > 0 Then
                                ' top point
                                iTriangle(0).X = (.Left + .Right) / 2 + cEpsilon
                                iTriangle(0).Y = .Top - mFlatBarGripHeightDPIScaled
                                ' left point
                                iTriangle(1).X = (.Left + .Right) / 2 - mFlatBarGripHeightDPIScaled + cEpsilon
                                iTriangle(1).Y = .Top
                                ' right point
                                iTriangle(2).X = (.Left + .Right) / 2 + mFlatBarGripHeightDPIScaled + cEpsilon
                                iTriangle(2).Y = .Top
                                DrawTriangle iTriangle, iFlatBarTopColor
                            Else
                                iLng = Abs(mFlatBarGripHeightDPIScaled)
                                ' top point
                                If mFlatBarHeightDPIScaled - Abs(mFlatBarGripHeightDPIScaled) < (mFlatBarHeightDPIScaled * 0.33) Or (mTabOrientation <> ssTabOrientationBottom) Then
                                    iTriangle(0).X = (.Left + .Right) / 2 + cEpsilon
                                    iTriangle(0).Y = .Top + (iLng + iFlatBarTopHeight)
                                    ' left point
                                    iTriangle(1).X = (.Left + .Right) / 2 - (iLng + iFlatBarTopHeight) + cEpsilon
                                    iTriangle(1).Y = .Top
                                    ' right point
                                    iTriangle(2).X = (.Left + .Right) / 2 + (iLng + iFlatBarTopHeight) + cEpsilon
                                    iTriangle(2).Y = .Top
                                    DrawTriangle iTriangle, iFlatBarTopColor
                                End If
                                iTriangle(0).X = (.Left + .Right) / 2 + cEpsilon
                                iTriangle(0).Y = .Top + iLng
                                ' left point
                                iTriangle(1).X = (.Left + .Right) / 2 - iLng + cEpsilon
                                iTriangle(1).Y = .Top
                                ' right point
                                iTriangle(2).X = (.Left + .Right) / 2 + iLng + cEpsilon ' + 1
                                iTriangle(2).Y = .Top
                                DrawTriangle iTriangle, IIf(iTabData.RowPos = 0, TranslatedColor(mBackColor), iTabColor_BackColorTab)
                            End If
                        End If
                    End If
                End If
                If iShowFlatBarBottom Then
                    If iActive Then
                        iColor = iTabColor_FlatBarColorSelectedTab
                    ElseIf iTabData.Hovered And mAmbientUserMode Then
                        iColor = iTabColor_FlatBarGlowColor
                    Else
                        iColor = iFlatBarTopColor
                    End If
                    iColor = TranslatedColor(iColor)
                    
                    FillCurvedGradient2 .Left + iLeftOffset, .Bottom - mFlatBarHeightDPIScaled + 3, .Right + iRightOffset, .Bottom + 3, iColor, iColor, 0, 0
                    If IIf(iActive, mHighlightFlatBarWithGripSelectedTab, mHighlightFlatBarWithGrip) Then
                        If (mFlatBarGripHeightDPIScaled > 0) And iHighlighted Then
                            ' top point
                            iTriangle(0).X = (.Left + .Right) / 2 + cEpsilon
                            iTriangle(0).Y = .Bottom + 2 + mFlatBarGripHeightDPIScaled
                            ' left point
                            iTriangle(1).X = (.Left + .Right) / 2 - mFlatBarGripHeightDPIScaled + cEpsilon
                            iTriangle(1).Y = .Bottom + 2
                            ' right point
                            iTriangle(2).X = (.Left + .Right) / 2 + mFlatBarGripHeightDPIScaled + cEpsilon
                            iTriangle(2).Y = .Bottom + 2
                            DrawTriangle iTriangle, iColor
                        ElseIf iHighlighted Then
                            If mFlatBarHeightDPIScaled - Abs(mFlatBarGripHeightDPIScaled) < (mFlatBarHeightDPIScaled * 0.33) Or (mTabOrientation = ssTabOrientationBottom) Then
                                iLng = Abs(mFlatBarGripHeightDPIScaled) + mFlatBarHeightDPIScaled
                                ' top point
                                iTriangle(0).X = (.Left + .Right) / 2 + cEpsilon
                                iTriangle(0).Y = .Bottom + 2 - iLng
                                ' left point
                                iTriangle(1).X = (.Left + .Right) / 2 - iLng + cEpsilon
                                iTriangle(1).Y = .Bottom + 2
                                ' right point
                                iTriangle(2).X = (.Left + .Right) / 2 + iLng + cEpsilon
                                iTriangle(2).Y = .Bottom + 2
                                DrawTriangle iTriangle, iColor
                            End If
                            
                            If iHighlightGradient <> ntGradientNone Then
                                If iActive Then
                                    If mFlatBodySeparationLineHeightDPIScaled > 0 Then
                                        iColor = TranslatedColor(mFlatBodySeparationLineColor)
                                    Else
                                        If IsEmpty(iTabData.CustomColors.HighlightColorSelectedTab) Then
                                            iColor = mHighlightColorSelectedTab
                                        Else
                                            iColor = iTabData.CustomColors.HighlightColorSelectedTab
                                        End If
                                    End If
                                Else
                                    iColor = iTabColor_BackColorSelectedTab
                                End If
                            Else
                                iColor = iTabColor_BackColorSelectedTab
                            End If
                            iColor = TranslatedColor(iColor)
                            ' top point
                            iTriangle(0).X = (.Left + .Right) / 2 + cEpsilon
                            iTriangle(0).Y = .Bottom + 2 + mFlatBarGripHeightDPIScaled
                            ' left point
                            iTriangle(1).X = (.Left + .Right) / 2 + mFlatBarGripHeightDPIScaled + cEpsilon
                            iTriangle(1).Y = .Bottom + 2
                            ' right point
                            iTriangle(2).X = (.Left + .Right) / 2 - mFlatBarGripHeightDPIScaled + cEpsilon
                            iTriangle(2).Y = .Bottom + 2
                            DrawTriangle iTriangle, iColor
                        End If
                    End If
                End If
            Else
                If iRoundedTabs Then
                    picDraw.Line (.Left + iLeftOffset + 4, .Top)-(.Right - 3, .Top), i3DDKShadow
                    picDraw.Line (.Left + iLeftOffset + 4, .Top + 1)-(.Right - 4, .Top + 1), i3DHighlightH
                    If iActive Then
                        picDraw.Line (.Left + iLeftOffset + 4, .Top + 2)-(.Right - 4, .Top + 2), i3DHighlightH
                    End If
                Else
                    picDraw.Line (.Left + iLeftOffset + 3, .Top)-(.Right - 3, .Top), i3DDKShadow
                    picDraw.Line (.Left + iLeftOffset + 3, .Top + 1)-(.Right - 3, .Top + 1), i3DHighlightH
                    If iActive Then
                        picDraw.Line (.Left + iLeftOffset + 3, .Top + 2)-(.Right - 3, .Top + 2), i3DHighlightH
                    End If
                End If
            End If
            
            'right line
            If mAppearanceIsPP Then
                If mTabOrientation = ssTabOrientationTop Then
                    picDraw.Line (.Right, .Top + 3)-(.Right, .Bottom + iExtI), i3DDKShadow
                    picDraw.Line (.Right - 1, .Top + 3)-(.Right - 1, .Bottom + iExtI), i3DShadowV
                ElseIf (mTabOrientation = ssTabOrientationLeft) Or (mTabOrientation = ntTabOrientationLeftHorizontal) Then
                    picDraw.Line (.Right, .Top + 3)-(.Right, .Bottom + iExtI), i3DHighlightH
                    picDraw.Line (.Right - 1, .Top + 3)-(.Right - 1, .Bottom + iExtI), iBackColorTabs2
                Else
                    picDraw.Line (.Right, .Top + 3)-(.Right, .Bottom + iExtI - 1), i3DDKShadow
                    picDraw.Line (.Right, .Bottom + iExtI - 1)-(.Right + 1, .Bottom + iExtI - 1), i3DShadowV
                    picDraw.Line (.Right - 1, .Top + 3)-(.Right - 1, .Bottom + iExtI), i3DShadowV
                End If
            ElseIf mAppearanceIsFlat Then
                If iHighlightFlatDrawBorder Then
                    picDraw.Line (.Right, .Top + iTopOffset + iFlatRightRoundness)-(.Right, .Bottom - iFlatRightRoundness + 1), iHighlightFlatDrawBorder_Color
                ElseIf Not ((iFlatRightLineColor = iTabColor_BackColorTab) And iTabData.RightTab And (iFlatRightRoundness > 0)) Then
                    If ((iHighlightGradient = ntGradientPlain) Or (iHighlightGradient = ntGradientSimple)) And iHighlighted And (Not iActive) And (mHighlightFlatBar Or mHighlightFlatBarSelectedTab) Then
                        picDraw.Line (.Right, .Top + iTopOffset + iFlatRightRoundness)-(.Right, .Bottom + iBottomOffset + iExtI), iHighlightColor
                    Else
                        picDraw.Line (.Right, .Top + iTopOffset + iFlatRightRoundness)-(.Right, .Bottom + iBottomOffset + iExtI), iFlatRightLineColor
                    End If
                End If
            Else
                picDraw.Line (.Right, .Top + 4)-(.Right, .Bottom + iExtI), i3DDKShadow
                picDraw.Line (.Right - 1, .Top + 4)-(.Right - 1, .Bottom + iExtI), i3DShadowV
                If iActive Then
                    picDraw.Line (.Right - 2, .Top + 4)-(.Right - 2, .Bottom + 1 + iExtI), i3DShadowV
                    If iTabData.RightTab Then
                        picDraw.Line (.Right - 1, .Bottom + iExtI)-(.Right - 1, .Bottom + iExtI + 2), i3DShadowV ' points of top line of body
                        picDraw.Line (.Right - 2, .Bottom + iExtI + 1)-(.Right - 2, .Bottom + iExtI + 2), i3DShadowV ' point of top line of body
                    Else
                        picDraw.Line (.Right - 1, .Bottom + iExtI)-(.Right - 1, .Bottom + iExtI + 2), i3DHighlightH  ' points of top line of body
                        picDraw.Line (.Right - 2, .Bottom + iExtI + 1)-(.Right - 2, .Bottom + iExtI + 2), i3DHighlightH  ' point of top line of body
                    End If
                End If
            End If
            
            'left line
            If mAppearanceIsPP Then
                If (mTabOrientation <> ssTabOrientationLeft) And (mTabOrientation <> ntTabOrientationLeftHorizontal) Then
                    If iRoundedTabs Then
                        picDraw.Line (.Left, .Top + 3)-(.Left, .Bottom + iExtI), i3DHighlightV
                    Else
                        picDraw.Line (.Left, .Top + 2)-(.Left, .Bottom + iExtI), i3DHighlightV
                    End If
                Else
                    picDraw.Line (.Left, .Top + 2)-(.Left, .Bottom), i3DDKShadow
                    picDraw.Line (.Left + 1, .Top + 2)-(.Left + 1, .Bottom + iExtI), i3DShadow
                    If iRoundedTabs Then
                        picDraw.Line (.Left, .Top + 3)-(.Left, .Bottom + iExtI), i3DDKShadow
                        picDraw.Line (.Left + 1, .Top + 3)-(.Left + 1, .Bottom + iExtI), i3DShadow
                    Else
                        picDraw.Line (.Left, .Top + 2)-(.Left, .Bottom + iExtI), i3DDKShadow
                        picDraw.Line (.Left + 1, .Top + 2)-(.Left + 1, .Bottom + iExtI), i3DShadow
                    End If
                End If
            ElseIf mAppearanceIsFlat Then
                If iHighlightFlatDrawBorder Then
                    picDraw.Line (.Left + iLeftOffset, .Top + iTopOffset + iFlatLeftRoundness)-(.Left + iLeftOffset, .Bottom - iFlatLeftRoundness + 1), iHighlightFlatDrawBorder_Color
                ElseIf iFlatLeftLineColor <> iTabColor_BackColorTab Then
                    If ((iHighlightGradient = ntGradientPlain) Or (iHighlightGradient = ntGradientSimple)) And iHighlighted And (Not iActive) And (mHighlightFlatBar Or mHighlightFlatBarSelectedTab) Then
                        picDraw.Line (.Left + iLeftOffset, .Top + iTopOffset + iFlatLeftRoundness)-(.Left + iLeftOffset, .Bottom + iBottomOffset + iExtI), iHighlightColor
                    Else
                        picDraw.Line (.Left + iLeftOffset, .Top + iTopOffset + iFlatLeftRoundness)-(.Left + iLeftOffset, .Bottom + iBottomOffset + iExtI), iFlatLeftLineColor
                    End If
                End If
            Else
                If iRoundedTabs Then
                    If iTabData.LeftTab Then
                        picDraw.Line (.Left, .Top + 5)-(.Left, .Bottom + iExtI), i3DDKShadow
                    End If
                    If (mTabOrientation = ssTabOrientationLeft) Or (mTabOrientation = ntTabOrientationLeftHorizontal) Then
                        If iActive Then
                            picDraw.Line (.Left, .Top + 5)-(.Left, .Bottom + iExtI + 1), i3DHighlightV
                            picDraw.Line (.Left + 1, .Top + 5)-(.Left + 1, .Bottom + 2 + iExtI), i3DHighlightV
                        Else
                            picDraw.Line (.Left, .Top + 5)-(.Left, .Bottom + iExtI), i3DHighlightV
                            picDraw.Line (.Left + 1, .Top + 5)-(.Left + 1, .Bottom + iExtI), iBackColorTabs2
                        End If
                    Else
                        If iActive Then
                            picDraw.Line (.Left, .Top + 5)-(.Left, .Bottom + iExtI + 1), i3DHighlightV
                            picDraw.Line (.Left + 1, .Top + 5)-(.Left + 1, .Bottom + 2 + iExtI), i3DHighlightV
                        Else
                            picDraw.Line (.Left, .Top + 5)-(.Left, .Bottom + iExtI), i3DHighlightV ' iBackColorTabs2
                        End If
                        
                    End If
                Else
                    picDraw.Line (.Left, .Top + 4)-(.Left, .Bottom + iExtI), i3DHighlightV
                    If iActive Then
                        picDraw.Line (.Left + 1, .Top + 4)-(.Left + 1, .Bottom + 1 + iExtI), i3DHighlightV
                        picDraw.Line (.Left, .Bottom + iExtI)-(.Left, .Bottom + iExtI + 2), i3DHighlightV   ' points of top line of body
                        picDraw.Line (.Left + 1, .Bottom + iExtI + 1)-(.Left + 1, .Bottom + iExtI + 2), i3DHighlightV ' point of top line of body
                    End If
                End If
                picDraw.Line (.Left - 1, .Top + 5)-(.Left - 1, .Bottom + iExtI), i3DDKShadow
            End If
            
            'top-right corner
            If mAppearanceIsPP Then
                If (mTabOrientation <> ssTabOrientationLeft) And (mTabOrientation <> ntTabOrientationLeftHorizontal) Then
                    If iRoundedTabs Then
                        picDraw.Line (.Right - 2, .Top + 1)-(.Right - 2, .Top + 2), i3DShadowV
                        picDraw.Line (.Right - 1, .Top + 1)-(.Right - 1, .Top + 2), i3DShadowV
                        picDraw.Line (.Right - 1, .Top + 2)-(.Right - 1, .Top + 3), i3DDKShadow
                    Else
                        picDraw.Line (.Right - 1, .Top + 1)-(.Right - 1, .Top + 2), i3DDKShadow
                        picDraw.Line (.Right - 1, .Top + 2)-(.Right - 1, .Top + 3), i3DShadowV
                        picDraw.Line (.Right, .Top + 2)-(.Right, .Top + 3), i3DDKShadow
                    End If
                Else
                    If iRoundedTabs Then
                        picDraw.Line (.Right - 2, .Top + 1)-(.Right - 2, .Top + 2), i3DHighlight
                        picDraw.Line (.Right - 1, .Top + 1)-(.Right - 1, .Top + 2), i3DHighlight
                        picDraw.Line (.Right - 1, .Top + 2)-(.Right - 1, .Top + 3), i3DHighlight
                    Else
                        picDraw.Line (.Right - 1, .Top + 1)-(.Right - 1, .Top + 2), i3DHighlight
                        picDraw.Line (.Right - 1, .Top + 2)-(.Right - 1, .Top + 3), i3DHighlight
                        picDraw.Line (.Right, .Top + 2)-(.Right, .Top + 3), i3DHighlight
                    End If
                End If
            ElseIf mAppearanceIsFlat Then
                If iFlatRightRoundness > 0 Then
                    ' draw rounded top-right corner
                    If iHighlightFlatDrawBorder Then
                        iLineColor = iHighlightFlatDrawBorder_Color
                    Else
                        If ((iHighlightGradient = ntGradientPlain) Or (iHighlightGradient = ntGradientSimple)) And iHighlighted And (Not iActive) And (mHighlightFlatBar Or mHighlightFlatBarSelectedTab) Then
                            iLineColor = IIf((mFlatBorderMode = ntBorderTabs) Or iActive, iFlatBorderColor, iHighlightColor)
                        Else
                            iLineColor = IIf((mFlatBorderMode = ntBorderTabs) Or iActive, iFlatBorderColor, iFlatTabsSeparationLineColor)
                        End If
                    End If
                    If iLineColor <> iTabColor_BackColorTab Then
                        DrawRoundedCorner ntCornerTopRight, .Right + iRightOffset, .Top + iTopOffset, iFlatRightRoundness, iLineColor, iFlatBarTopHeight
                    End If
                End If
            Else
                If iRoundedTabs Then
                    picDraw.Line (.Right - 1, .Top + 4)-(.Right - 1, .Top + 1), i3DDKShadow
                    picDraw.Line (.Right - 2, .Top + 1)-(.Right - 4, .Top + 1), i3DDKShadow
                    picDraw.Line (.Right - 2, .Top + 2)-(.Right - 3, .Top + 2), i3DShadowV
                    picDraw.Line (.Right - 4, .Top + 1)-(.Right - 3, .Top + 1), i3DShadowV
                    picDraw.Line (.Right - 2, .Top + 3)-(.Right - 2, .Top + 4), i3DShadowV
                    If iActive Then
                        picDraw.Line (.Right - 3, .Top + 2)-(.Right, .Top + 5), i3DShadowV
                        picDraw.Line (.Right - 4, .Top + 2)-(.Right - 1, .Top + 5), i3DShadowV
                        picDraw.Line (.Right - 3, .Top + 4)-(.Right - 1, .Top + 6), i3DShadowV
                    End If
                Else
                    picDraw.Line (.Right - 4, .Top)-(.Right, .Top + 4), i3DDKShadow
                    If iActive Then
                        picDraw.Line (.Right - 3, .Top + 2)-(.Right, .Top + 5), i3DShadowV
                        picDraw.Line (.Right - 4, .Top + 2)-(.Right - 1, .Top + 5), i3DShadowV
                        picDraw.Line (.Right - 4, .Top + 3)-(.Right - 1, .Top + 6), i3DShadowV
                    Else
                        picDraw.Line (.Right - 4, .Top + 1)-(.Right - 1, .Top + 4), i3DShadowV
                    End If
                End If
            End If
            
            'top-left corner
            If mAppearanceIsPP Then
                If (mTabOrientation <> ssTabOrientationLeft) And (mTabOrientation <> ntTabOrientationLeftHorizontal) Then
                    If iRoundedTabs Then
                        picDraw.Line (.Left + 1, .Top + 2)-(.Left + 1, .Top + 3), i3DHighlightH
                        picDraw.Line (.Left, .Top + 3)-(.Left, .Top + 4), i3DHighlightH
                        picDraw.Line (.Left + 1, .Top + 1)-(.Left + 3, .Top + 1), i3DHighlightH
                    Else
                        picDraw.Line (.Left, .Top + 2)-(.Left + 3, .Top - 1), i3DHighlightH
                    End If
                Else
                    If iRoundedTabs Then
                        picDraw.Line (.Left + 1, .Top + 2)-(.Left + 1, .Top + 3), i3DHighlightV
                        picDraw.Line (.Left + 1, .Top + 1)-(.Left + 3, .Top + 1), i3DHighlightV
                    Else
                        picDraw.Line (.Left, .Top + 2)-(.Left + 3, .Top - 1), i3DHighlightV
                    End If
                End If
            ElseIf mAppearanceIsFlat Then
                If iFlatLeftRoundness > 0 Then
                    If iHighlightFlatDrawBorder Then
                        iLineColor = iHighlightFlatDrawBorder_Color
                    Else
                        If ((iHighlightGradient = ntGradientPlain) Or (iHighlightGradient = ntGradientSimple)) And iHighlighted And (Not iActive) And (mHighlightFlatBar Or mHighlightFlatBarSelectedTab) Then
                            iLineColor = IIf((mFlatBorderMode = ntBorderTabs) Or iActive, iFlatBorderColor, iHighlightColor)
                        Else
                            iLineColor = IIf((mFlatBorderMode = ntBorderTabs) Or (iActive And (mFlatBorderMode = ntBorderSelectedTab)), iFlatBorderColor, iFlatTabsSeparationLineColor)
                        End If
                    End If
                    If iLineColor <> iTabColor_BackColorTab Then
                        DrawRoundedCorner ntCornerTopleft, .Left + iLeftOffset, .Top + iTopOffset, iFlatLeftRoundness, iLineColor, iFlatBarTopHeight
                    End If
                End If
            Else
                If iRoundedTabs Then
                    picDraw.Line (.Left + iLeftOffset + 1, .Top + 4)-(.Left + iLeftOffset + 1, .Top + 1), i3DDKShadow
                    picDraw.Line (.Left + iLeftOffset + 2, .Top + 1)-(.Left + iLeftOffset + 4, .Top + 1), i3DDKShadow
                    picDraw.Line (.Left + iLeftOffset + 2, .Top + 3)-(.Left + iLeftOffset + 2, .Top + 2), i3DHighlightH
                    picDraw.Line (.Left + iLeftOffset + 2, .Top + 2)-(.Left + iLeftOffset + 4, .Top + 2), i3DHighlightH
                    picDraw.Line (.Left + iLeftOffset, .Top + 4)-(.Left + iLeftOffset, .Top + 3), i3DDKShadow
                    picDraw.Line (.Left + iLeftOffset + 1, .Top + 4)-(.Left + iLeftOffset + 1, .Top + 5), i3DHighlightH
                    If iActive Then
                        picDraw.Line (.Left + iLeftOffset + 2, .Top + 3)-(.Left + iLeftOffset + 4, .Top + 1), i3DHighlightH
                        picDraw.Line (.Left + iLeftOffset + 2, .Top + 4)-(.Left + iLeftOffset + 5, .Top + 1), i3DHighlightH
                        picDraw.Line (.Left + iLeftOffset + 2, .Top + 5)-(.Left + iLeftOffset + 5, .Top + 2), i3DHighlightH
                    End If
                Else
                    picDraw.Line (.Left + iLeftOffset - 1, .Top + 4)-(.Left + iLeftOffset + 3, .Top), i3DDKShadow
                    If iActive Then
                        picDraw.Line (.Left + iLeftOffset + 1, .Top + 3)-(.Left + iLeftOffset + 3, .Top + 1), i3DHighlightH
                        picDraw.Line (.Left + iLeftOffset + 1, .Top + 4)-(.Left + iLeftOffset + 4, .Top + 1), i3DHighlightH
                        picDraw.Line (.Left + iLeftOffset + 2, .Top + 4)-(.Left + iLeftOffset + 4, .Top + 2), i3DHighlightH
                    Else
                        picDraw.Line (.Left + iLeftOffset, .Top + 4)-(.Left + iLeftOffset + 3, .Top + 1), i3DHighlightH
                    End If
                End If
            End If
            
            ' Bottom line
            If mAppearanceIsFlat Then
                If iHighlightFlatDrawBorder Then
                    picDraw.Line (.Left + iLeftOffset + iFlatLeftRoundness, .Bottom + 1)-(.Right + iRightOffset - iFlatRightRoundness, .Bottom + 1), iHighlightFlatDrawBorder_Color
                    If (iFlatRightRoundness > 0) Then
                        DrawRoundedCorner ntCornerBottomRight, .Right + iRightOffset, .Bottom + 1, iFlatRightRoundness, iHighlightFlatDrawBorder_Color
                    End If
                    If (iFlatLeftRoundness > 0) Then
                        DrawRoundedCorner ntCornerBottomLeft, .Left + iLeftOffset, .Bottom + 1, iFlatLeftRoundness, iHighlightFlatDrawBorder_Color
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Sub DrawInactiveBodyPart(ByVal nLeft As Long, ByVal nTop As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal nXShift As Long, ByVal nRowPos As Long, nTab As Long)
    Dim iDoRightLine As Boolean
    Dim iDoBottomLine As Boolean
    Dim iBackColorTabs As Long
    Dim iLineColor As Long
    Dim iFlatBorderColor As Long
    
    If (nWidth < 1) Or (nHeight < 1) Or (nXShift > mClientWidth) Then Exit Sub
    
    If IsEmpty(mTabData(nTab).CustomColors.BackColorTab) Or (mHandleHighContrastTheme And mHighContrastThemeOn) Then
        iBackColorTabs = mBackColorTabs2
    Else
        iBackColorTabs = mTabData(nTab).CustomColors.BackColorTab
    End If
    
    If mControlIsThemed Then
        If (nWidth > mThemedBodyRightShadowPixels) Then
            EnsureInactiveBodyThemedReady
            BitBlt picDraw.hDC, nLeft, nTop, nWidth, nHeight, picInactiveBodyThemed.hDC, nXShift, 0, vbSrcCopy
        End If
    Else
        iDoRightLine = mClientWidth - (nWidth + nXShift) <= 0
        iDoBottomLine = mClientHeight - nHeight <= 0
        
        If mAppearanceIsFlat Then
            If Not (mHighlightFlatBar And (mFlatBarPosition = ntBarPositionBottom) And (mFlatBarHeightDPIScaled > 0)) Then
                FillCurvedGradient2 nLeft, nTop, nLeft + nWidth, nTop + nHeight, iBackColorTabs, iBackColorTabs, mFlatRoundnessTopDPIScaled, mFlatRoundnessTopDPIScaled, mFlatRoundnessBottomDPIScaled, mFlatRoundnessBottomDPIScaled
            End If
        Else
            picDraw.Line (nLeft, nTop)-(nLeft + nWidth, nTop + nHeight), iBackColorTabs, BF
        End If
        
        'top line
        If mAppearanceIsPP Then
            If (mTabOrientation = ssTabOrientationTop) Or (mTabOrientation = ssTabOrientationLeft) Or (mTabOrientation = ntTabOrientationLeftHorizontal) Then
                picDraw.Line (nLeft - 1, nTop)-(nLeft + nWidth, nTop), m3DHighlight
            Else
                picDraw.Line (nLeft - 1, nTop)-(nLeft + nWidth, nTop), m3DDKShadow
                picDraw.Line (nLeft - 1, nTop + 1)-(nLeft + nWidth - 1, nTop + 1), m3DShadow
            End If
        ElseIf mAppearanceIsFlat Then
            If mFlatBorderMode = ntBorderTabs Then
                iFlatBorderColor = TranslatedColor(mFlatBorderColor)
            Else
                iFlatBorderColor = TranslatedColor(mFlatTabsSeparationLineColor)
            End If
            picDraw.Line (nLeft - 1, nTop)-(nLeft + nWidth - mFlatRoundnessTopDPIScaled, nTop), iFlatBorderColor 'm3DDKShadow
        Else
            picDraw.Line (nLeft - 1, nTop)-(nLeft + nWidth, nTop), m3DDKShadow
            picDraw.Line (nLeft - 1, nTop + 1)-(nLeft + 1 + nWidth, nTop + 1), m3DHighlightH
        End If
        
        'right line
        If iDoRightLine Then
            If mAppearanceIsFlat Then
                If ((nLeft + nWidth) - mRightMostTabsRightPos(nRowPos)) > mFlatRoundnessTopDPIScaled Then
                    picDraw.Line (nLeft + nWidth, nTop + mFlatRoundnessTopDPIScaled)-(nLeft + nWidth, nTop + nHeight - mFlatRoundnessBottomDPIScaled), iFlatBorderColor   'm3DDKShadow
                Else
                    picDraw.Line (nLeft + nWidth, nTop)-(nLeft + nWidth, nTop + nHeight - mFlatRoundnessBottomDPIScaled), iFlatBorderColor   'm3DDKShadow
                End If
                ' top-right corner
                If (mFlatRoundnessTopDPIScaled > 0) Then
                    If ((nLeft + nWidth) - mRightMostTabsRightPos(nRowPos)) > mFlatRoundnessTopDPIScaled Then
                        iLineColor = iFlatBorderColor
                        If iLineColor <> iBackColorTabs Then
                            DrawRoundedCorner ntCornerTopRight, nLeft + nWidth, nTop, mFlatRoundnessTopDPIScaled, iLineColor
                        End If
                    End If
                End If
            Else
                If ((mTabOrientation <> ssTabOrientationLeft) And (mTabOrientation <> ntTabOrientationLeftHorizontal)) Or (Not mAppearanceIsPP) Then
                    picDraw.Line (nLeft + nWidth, nTop)-(nLeft + nWidth, nTop + nHeight), m3DDKShadow
                    picDraw.Line (nLeft + nWidth - 1, nTop + 1)-(nLeft + nWidth - 1, nTop + nHeight), m3DShadowV
                Else
                    picDraw.Line (nLeft + nWidth, nTop)-(nLeft + nWidth, nTop + nHeight), m3DHighlightH
                End If
            End If
        End If
        
        'bottom line
        If iDoBottomLine Then
            If mAppearanceIsFlat Then
                picDraw.Line (nLeft - 1, nTop + nHeight)-(nLeft + nWidth + 1 - mFlatRoundnessBottomDPIScaled, nTop + nHeight), iFlatBorderColor
            Else
                If (mTabOrientation = ssTabOrientationTop) Or (mTabOrientation = ssTabOrientationLeft) Or (mTabOrientation = ntTabOrientationLeftHorizontal) Then
                    picDraw.Line (nLeft - 1, nTop - 1 + nHeight)-(nLeft + nWidth, nTop - 1 + nHeight), m3DShadow
                    picDraw.Line (nLeft - 1, nTop + nHeight)-(nLeft + nWidth + 1, nTop + nHeight), m3DDKShadow
                Else
                    picDraw.Line (nLeft - 1, nTop - 1 + nHeight)-(nLeft + nWidth, nTop - 1 + nHeight), m3DHighlight
                End If
            End If
        End If
        
        If mAppearanceIsFlat Then
            If mFlatRoundnessTopDPIScaled > 0 Then
                iLineColor = iFlatBorderColor ' m3DDKShadow
                ' botom-right corner
                DrawRoundedCorner ntCornerBottomRight, nLeft + nWidth, nTop + nHeight, mFlatRoundnessTop2, iLineColor
            End If
        End If
    End If
End Sub


Private Sub DrawBody(ByVal nScaleHeight As Long)
    Dim iLng As Long
    Dim iLineColor As Long
    Dim iFlatBorderColor As Long
    Dim iFlatBodySeparationLineColor As Long
    Dim iColor As Long
    Dim iActiveTabIsLeft As Boolean
    Dim iTopLeftCornerIsRounded As Boolean
    Dim iTopRightCornerIsRounded As Boolean
    
    If mControlIsThemed Then
        EnsureBodyThemedReady
        BitBlt picDraw.hDC, 0, mClientStart - 2, picBodyThemed.ScaleWidth, picBodyThemed.ScaleHeight, picBodyThemed.hDC, 0, 0, vbSrcCopy
    Else
        ' background
        If mAppearanceIsPP Then
            iLng = -1
        Else
            iLng = 1
        End If
        
        iColor = mBackColorSelectedTab2
        
        If mBackStyle = ntOpaqueSelectedTab Then
            iColor = iColor Xor 1
        End If
        If mAppearanceIsFlat Then
            If mTab > -1 Then
                iActiveTabIsLeft = mTabData(mTab).LeftTab
            End If
            iTopLeftCornerIsRounded = (mFlatBorderMode = ntBorderSelectedTab) And (mHighlightFlatDrawBorderSelectedTab Or (Not iActiveTabIsLeft))
            FillCurvedGradient2 0, mClientStart + iLng, mClientWidth - 1, nScaleHeight - 1, iColor, iColor, IIf(iTopLeftCornerIsRounded And (mFlatRoundnessTopDPIScaled > 0) And (mFlatBodySeparationLineHeight = 1), mFlatRoundnessTopDPIScaled, 0), IIf(((mClientWidth - mTabData(mTab).TabRect.Right) > 3) And ((mFlatBorderMode = ntBorderSelectedTab) Or ((mClientWidth - mRightMostTabsRightPos(mRows - 1)) > 3)), mFlatRoundnessTopDPIScaled, 0), mFlatRoundnessBottomDPIScaled, mFlatRoundnessBottomDPIScaled
        Else
            picDraw.Line (0, mClientStart + iLng)-(mClientWidth - 1, nScaleHeight - 1), iColor, BF
        End If
        
        If mAppearanceIsPP Then
            ' top line
            If (mTabOrientation = ssTabOrientationTop) Or (mTabOrientation = ssTabOrientationLeft) Or (mTabOrientation = ntTabOrientationLeftHorizontal) Then
                picDraw.Line (1, mClientStart - 2)-(mClientWidth - 1, mClientStart - 2), m3DHighlightH_Sel
            Else
                picDraw.Line (0, mClientStart - 2)-(mClientWidth - 1, mClientStart - 2), m3DDKShadow_Sel
                picDraw.Line (1, mClientStart - 1)-(mClientWidth - 1, mClientStart - 1), m3DShadow_Sel
            End If
            
            If (mTabOrientation = ssTabOrientationTop) Then
                'left line
                picDraw.Line (0, mClientStart - 1)-(0, nScaleHeight - 1), m3DHighlightV_Sel
                
                'right line
                picDraw.Line (mClientWidth - 1, mClientStart - 2)-(mClientWidth - 1, nScaleHeight - 1), m3DDKShadow_Sel
                picDraw.Line (mClientWidth - 2, mClientStart - 1)-(mClientWidth - 2, nScaleHeight - 2), m3DShadowV_Sel
                
                'bottom line
                picDraw.Line (0, nScaleHeight - 1)-(mClientWidth, nScaleHeight - 1), m3DDKShadow_Sel
                If mClientHeight > 3 Then
                    picDraw.Line (1, nScaleHeight - 2)-(mClientWidth - 1, nScaleHeight - 2), m3DShadowH_Sel
                End If
            ElseIf (mTabOrientation = ssTabOrientationLeft) Or (mTabOrientation = ntTabOrientationLeftHorizontal) Then
                'left line
                picDraw.Line (0, mClientStart - 1)-(0, nScaleHeight - 1), m3DDKShadow_Sel
                picDraw.Line (1, mClientStart - 1)-(1, nScaleHeight - 1), m3DShadow_Sel
            
                'right line
                picDraw.Line (mClientWidth - 1, mClientStart - 2)-(mClientWidth - 1, nScaleHeight - 1), m3DHighlight_Sel
            
                'bottom line
                picDraw.Line (0, nScaleHeight - 1)-(mClientWidth, nScaleHeight - 1), m3DDKShadow_Sel
                If mClientHeight > 3 Then
                    picDraw.Line (1, nScaleHeight - 2)-(mClientWidth - 1, nScaleHeight - 2), m3DShadowH_Sel
                End If
            Else 'Bottom or Right
                'left line
                picDraw.Line (0, mClientStart - 1)-(0, nScaleHeight - 1), m3DHighlightV_Sel
                
                'right line
                picDraw.Line (mClientWidth - 1, mClientStart - 2)-(mClientWidth - 1, nScaleHeight), m3DDKShadow_Sel
                picDraw.Line (mClientWidth - 2, mClientStart - 1)-(mClientWidth - 2, nScaleHeight - 1), m3DShadowV_Sel
                
                ' bottom line
                picDraw.Line (0, nScaleHeight - 1)-(mClientWidth - 1, nScaleHeight - 1), m3DShadowH_Sel
            End If
        ElseIf mAppearanceIsFlat Then
            iFlatBorderColor = TranslatedColor(mFlatBorderColor)
            If (mFlatBorderMode = ntBorderTabs) Or (mFlatBodySeparationLineHeight > 1) Then
                iFlatBodySeparationLineColor = TranslatedColor(mFlatBodySeparationLineColor)
            Else
                iFlatBodySeparationLineColor = iFlatBorderColor
            End If
            
            ' top line
            If (iFlatBodySeparationLineColor <> mBackColorTabs2) Then
                If Not ((mTabOrientation = ssTabOrientationBottom) And (mFlatBarHeightDPIScaled > 0) And mHighlightFlatBar) Then
                    If mFlatBodySeparationLineHeight = 1 Then
                        picDraw.Line (IIf(iTopLeftCornerIsRounded, mFlatRoundnessTopDPIScaled, 0), mClientStart)-(mClientWidth - 1 - IIf(mFlatBorderMode = ntBorderSelectedTab, mFlatRoundnessTopDPIScaled, 0), mClientStart), iFlatBodySeparationLineColor
                    ElseIf mFlatBodySeparationLineHeight > 1 Then
                        picDraw.Line (0, mClientStart + 1)-(mClientWidth - 2, mClientStart + mFlatBodySeparationLineHeightDPIScaled), iFlatBodySeparationLineColor, BF
                    End If
                End If
            End If
            
            ' left line
            picDraw.Line (0, mClientStart + IIf(iTopLeftCornerIsRounded And (mFlatBodySeparationLineHeight = 1), mFlatRoundnessTopDPIScaled, 0))-(0, nScaleHeight - 1 - mFlatRoundnessBottomDPIScaled), iFlatBorderColor
            ' top-right corner round
            If ((mClientWidth - mTabData(mTab).TabRect.Right) > 3) And ((mFlatBorderMode = ntBorderSelectedTab) Or ((mClientWidth - mRightMostTabsRightPos(mRows - 1)) > 3)) Then
                If (mFlatRoundnessTopDPIScaled > 0) And (mFlatBodySeparationLineHeight = 1) Then
                    If ((mClientWidth - mTabData(mTab).TabRect.Right) > 3) And mFlatBorderMode = ntBorderSelectedTab Then
                        iLineColor = iFlatBodySeparationLineColor
                    Else
                        iLineColor = iFlatBorderColor
                    End If
                    DrawRoundedCorner ntCornerTopRight, mClientWidth - 1, mClientStart, mFlatRoundnessTopDPIScaled, iLineColor
                    iTopRightCornerIsRounded = True
                End If
            End If
            ' right line
            picDraw.Line (mClientWidth - 1, mClientStart + IIf(iTopRightCornerIsRounded, mFlatRoundnessTopDPIScaled, 0))-(mClientWidth - 1, nScaleHeight - 1 - mFlatRoundnessBottomDPIScaled), iFlatBorderColor
            ' bottom line
            picDraw.Line (mFlatRoundnessBottomDPIScaled, nScaleHeight - 1)-(mClientWidth - mFlatRoundnessBottomDPIScaled, nScaleHeight - 1), iFlatBorderColor
            
            If iTopLeftCornerIsRounded Then
                ' top-left corner
                If (mFlatRoundnessTopDPIScaled > 0) And (mFlatBodySeparationLineHeight = 1) Then
                    DrawRoundedCorner ntCornerTopleft, 0, mClientStart, mFlatRoundnessTopDPIScaled, iFlatBorderColor
                End If
            End If
            
            If (mClientWidth - mRightMostTabsRightPos(mRows - 1)) > 3 Then
                picDraw.Line (mRightMostTabsRightPos(mRows - 1), mClientStart)-(mClientWidth - 1 - mFlatRoundnessTopDPIScaled, mClientStart), iFlatBorderColor  ' iFlatBodySeparationLineColor
            End If
            
            If mFlatRoundnessBottomDPIScaled > 0 Then
                iLineColor = iFlatBorderColor
                ' botom-left corner
                DrawRoundedCorner ntCornerBottomLeft, 0, nScaleHeight - 1, mFlatRoundnessBottomDPIScaled, iLineColor
                
                ' botom-right corner
                DrawRoundedCorner ntCornerBottomRight, mClientWidth - 1, nScaleHeight - 1, mFlatRoundnessBottomDPIScaled, iLineColor
            End If
        Else
            ' top line
            picDraw.Line (0, mClientStart - 2)-(mClientWidth - 1, mClientStart - 2), m3DDKShadow_Sel
            picDraw.Line (2, mClientStart - 1)-(mClientWidth - 1, mClientStart - 1), m3DHighlightH_Sel
            picDraw.Line (3, mClientStart)-(mClientWidth - 2, mClientStart), m3DHighlightH_Sel
            
            ' left line
            picDraw.Line (0, mClientStart - 1)-(0, nScaleHeight - 1), m3DDKShadow_Sel
            picDraw.Line (1, mClientStart - 1)-(1, nScaleHeight - 2), m3DHighlightV_Sel
            picDraw.Line (2, mClientStart + 1)-(2, nScaleHeight - 3), m3DHighlightV_Sel

            ' right line
            picDraw.Line (mClientWidth - 1, mClientStart - 2)-(mClientWidth - 1, nScaleHeight - 1), m3DDKShadow_Sel
            picDraw.Line (mClientWidth - 2, mClientStart - 1)-(mClientWidth - 2, nScaleHeight - 2), m3DShadowV_Sel
            picDraw.Line (mClientWidth - 3, mClientStart)-(mClientWidth - 3, nScaleHeight - 3), m3DShadowV_Sel
            
            ' bottom line
            picDraw.Line (0, nScaleHeight - 1)-(mClientWidth, nScaleHeight - 1), m3DDKShadow_Sel
            If mClientHeight > 3 Then
                picDraw.Line (1, nScaleHeight - 2)-(mClientWidth - 1, nScaleHeight - 2), m3DShadowH_Sel
            End If
            If mClientHeight > 4 Then
                picDraw.Line (2, nScaleHeight - 3)-(mClientWidth - 2, nScaleHeight - 3), m3DShadowH_Sel
            End If
        End If
    End If
End Sub

Private Sub DrawTabPicureAndCaption(ByVal nTab As Long)
    Dim iTabData As T_TabData
    Dim iTabRect As RECT
    Dim iTabSpaceRect As RECT
    Dim iCaptionRect As RECT
    Dim iMeasureRect As RECT
    Dim iFocusRect As RECT
    Dim iAuxPicture As StdPicture
    Dim iPicWidth As Long
    Dim iPicHeight As Long
    Dim iCaption As String
    Dim iFontBoldPrev As Boolean
    Dim iFlags As Long
    Dim iPicLeft As Long
    Dim iPicTop As Long
    Dim iLng As Long
    Dim iPicSourceShiftX As Long
    Dim iPicSourceShiftY As Long
    Dim iTabSpaceWidth As Long
    Dim iTabSpaceHeight As Long
    Dim iMeasureWidth As Long
    Dim iMeasureHeight As Long
    Dim iPicWidthToShow As Long
    Dim iPicHeightToShow As Long
    Dim iIconAlignment As NTIconAlignmentConstants
    Dim iBackColorTabs2 As Long
    Dim iForeColor As Long
    Dim iGrayText As Long
    Dim iForeColor2 As Long
    Dim iDrawIcon As Boolean
    Dim iIconCharacter As String
    Dim iIconCharRect As RECT
    Dim iFontPrev As StdFont
    Dim iIconColor As Long
    Dim iForeColorPrev As Long
    Dim iIconFont As StdFont
    Dim iFlatBarHeightTop As Long
    Dim iFlatBarHeightBottom As Long
    Dim iAuxRect As RECT
    Dim iGMPrev As Long
    Dim iTx1 As XFORM
    Dim iGMPrev2 As Long
    Dim iTx2 As XFORM
    Dim iTx2Prev As XFORM
    Dim iFlatBarPosition As NTFlatBarPositionConstants
    Dim iTabCenterX As Long
    Dim iTabCenterY As Long
    Dim iActive As Boolean
    Dim iPropIconColor As Long
    Dim iIconColorSelectedTab As Long
    Dim iIconColorMouseHover As Long
    Dim iIconColorMouseHoverSelectedTab As Long
    Dim iIconColorTabHighlighted As Long
    Dim iTDIFormIcon_Left As Long
    Dim iTDIFormIcon_Top As Long
    Dim iTDIFormIcon_IconCharRect As RECT
    Dim iTDIFormsNoIcons As Boolean
    Dim iDo As Boolean
    Dim iTDIFormsFontPrev As StdFont
    Dim iIsTDIFormsFirstTab As Boolean
    
    If Not mTabData(nTab).Visible Then Exit Sub
    If Not mTabData(nTab).PicToUseSet Then SetPicToUse nTab
    
    iTabData = mTabData(nTab)
    iTabRect = iTabData.TabRect
    
    If mTabsAreRotatedButCaptionsAreHorizontal Then
        iTabCenterX = (iTabRect.Left + iTabRect.Right) / 2
        iTabCenterY = (iTabRect.Top + iTabRect.Bottom) / 2
        iTabRect = RotateRect90Degrees(iTabRect, True, iTabCenterX, iTabCenterY)
    End If
    
    If mTabsAreRotatedButCaptionsAreHorizontal Then
        iGMPrev2 = SetGraphicsMode(picDraw.hDC, GM_ADVANCED)
        GetWorldTransform picDraw.hDC, iTx2Prev
        If (mTabOrientation = ntTabOrientationLeftHorizontal) Then
            iTx2.eM11 = 0
            iTx2.eM12 = 1
            iTx2.eM21 = -1
            iTx2.eM22 = 0
            iTx2.eDx = iTabCenterX
            iTx2.eDy = iTabCenterY
        Else
            iTx2.eM11 = 0
            iTx2.eM12 = -1
            iTx2.eM21 = 1
            iTx2.eM22 = 0
            iTx2.eDx = iTabCenterX
            iTx2.eDy = iTabCenterY
        End If
        SetWorldTransform picDraw.hDC, iTx2
    End If
    
    iActive = (nTab = mTab)
    If CanReorderTabsEffective Then
        If iActive Then
            If DraggingATab Then
                If mTabOrientation = ssTabOrientationBottom Then
                    iTabRect.Left = iTabRect.Left + mMouseX2 - mMouseX
                    iTabRect.Right = iTabRect.Right + mMouseX2 - mMouseX
                    iTabRect.Top = iTabRect.Top - mMouseY2 + mMouseY
                    iTabRect.Bottom = iTabRect.Bottom - mMouseY2 + mMouseY
                ElseIf (mTabOrientation = ssTabOrientationLeft) Then
                    iTabRect.Left = iTabRect.Left - mMouseY2 + mMouseY
                    iTabRect.Right = iTabRect.Right - mMouseY2 + mMouseY
                    iTabRect.Top = iTabRect.Top + mMouseX2 - mMouseX
                    iTabRect.Bottom = iTabRect.Bottom + mMouseX2 - mMouseX
                ElseIf (mTabOrientation = ssTabOrientationRight) Then
                    iTabRect.Left = iTabRect.Left + mMouseY2 - mMouseY
                    iTabRect.Right = iTabRect.Right + mMouseY2 - mMouseY
                    iTabRect.Top = iTabRect.Top - mMouseX2 + mMouseX
                    iTabRect.Bottom = iTabRect.Bottom - mMouseX2 + mMouseX
                Else
                    iTabRect.Left = iTabRect.Left + mMouseX2 - mMouseX
                    iTabRect.Right = iTabRect.Right + mMouseX2 - mMouseX
                    iTabRect.Top = iTabRect.Top + mMouseY2 - mMouseY
                    iTabRect.Bottom = iTabRect.Bottom + mMouseY2 - mMouseY
                End If
            End If
        End If
    End If
    
    iFlatBarPosition = mFlatBarPosition
    
    If mAppearanceIsFlat Then
        If mHighlightFlatBar Or mHighlightFlatBarSelectedTab Then
            If iFlatBarPosition = ntBarPositionTop Then
                If mTabOrientation = ssTabOrientationBottom Then
                    iFlatBarHeightBottom = -mFlatBarHeightDPIScaled
                Else
                    iFlatBarHeightTop = mFlatBarHeightDPIScaled
                End If
                If (mHighlightFlatBarWithGrip And (Not iActive) And iTabData.Hovered And mAmbientUserMode) Or (mHighlightFlatBarWithGripSelectedTab And iActive) Then
                    If mFlatBarGripHeightDPIScaled < 0 Then
                        If mTabOrientation = ssTabOrientationBottom Then
                            iFlatBarHeightBottom = iFlatBarHeightBottom - Abs(mFlatBarGripHeightDPIScaled) - 1
                        Else
                            iFlatBarHeightTop = iFlatBarHeightTop + Abs(mFlatBarGripHeightDPIScaled) + 1
                        End If
                    Else
                        If mTabOrientation = ssTabOrientationBottom Then
                            iFlatBarHeightBottom = iFlatBarHeightBottom - Abs(mFlatBarGripHeightDPIScaled) - 1
                        End If
                    End If
                End If
            Else
                iFlatBarHeightBottom = mFlatBarHeightDPIScaled
                If mHighlightFlatBarWithGrip Or mHighlightFlatBarWithGripSelectedTab Then
                    If mFlatBarGripHeightDPIScaled < 0 Then
                        iFlatBarHeightBottom = iFlatBarHeightBottom + Abs(mFlatBarGripHeightDPIScaled) + 1
                    End If
                End If
            End If
        End If
    End If
    
    If iActive Then
        iBackColorTabs2 = mBackColorSelectedTab2
        If IsEmpty(iTabData.CustomColors.ForeColorSelectedTab) Then
            iForeColor = mForeColorSelectedTab
        Else
            iForeColor = iTabData.CustomColors.ForeColorSelectedTab
        End If
        If IsEmpty(iTabData.CustomColors.IconColorSelectedTab) Then
            iIconColorSelectedTab = mIconColorSelectedTab
        Else
            iIconColorSelectedTab = iTabData.CustomColors.IconColorSelectedTab
        End If
        If IsEmpty(iTabData.CustomColors.IconColorMouseHoverSelectedTab) Then
            iIconColorMouseHoverSelectedTab = mIconColorMouseHoverSelectedTab
        Else
            iIconColorMouseHoverSelectedTab = iTabData.CustomColors.IconColorMouseHoverSelectedTab
        End If
        If iIconColorMouseHoverSelectedTab <> iIconColorSelectedTab Then
            If (mMouseIsOverIcon_Tab = CInt(nTab) Or (tmrHighlightIcon.Enabled And (Val(tmrHighlightIcon.Tag) = nTab))) And (Not tmrPreHighlightIcon.Enabled) Then
                iIconColor = iIconColorMouseHoverSelectedTab
            Else
                iIconColor = iIconColorSelectedTab
            End If
        Else
            iIconColor = iIconColorSelectedTab
        End If
        iGrayText = mGrayText_Sel
    Else
        If IsEmpty(iTabData.CustomColors.BackColorTab) Or (mHandleHighContrastTheme And mHighContrastThemeOn) Then
            iBackColorTabs2 = mBackColorTabs2
        Else
            iBackColorTabs2 = iTabData.CustomColors.BackColorTab
        End If
        If iTabData.Hovered And mAmbientUserMode And Not DraggingATab Then
            If IsEmpty(iTabData.CustomColors.ForeColorHighlighted) Then
                iForeColor = mForeColorHighlighted
            Else
                iForeColor = iTabData.CustomColors.ForeColorHighlighted
            End If
        Else
            If IsEmpty(iTabData.CustomColors.ForeColor) Then
                iForeColor = mForeColor
            Else
                iForeColor = iTabData.CustomColors.ForeColor
            End If
        End If
        If IsEmpty(iTabData.CustomColors.IconColor) Then
            iPropIconColor = mIconColor
        Else
            iPropIconColor = iTabData.CustomColors.IconColor
        End If
        If IsEmpty(iTabData.CustomColors.IconColorMouseHover) Then
            iIconColorMouseHover = mIconColorMouseHover
        Else
            iIconColorMouseHover = iTabData.CustomColors.IconColorMouseHover
        End If
        If IsEmpty(iTabData.CustomColors.IconColorTabHighlighted) Then
            iIconColorTabHighlighted = mIconColorTabHighlighted
        Else
            iIconColorTabHighlighted = iTabData.CustomColors.IconColorTabHighlighted
        End If
        If iIconColorMouseHover <> iPropIconColor Then
            If mAmbientUserMode And (mMouseIsOverIcon_Tab = CInt(nTab) Or (tmrHighlightIcon.Enabled And (Val(tmrHighlightIcon.Tag) = nTab))) And (Not tmrPreHighlightIcon.Enabled) Then
            If (mMouseIsOverIcon_Tab = CInt(nTab)) Then
                    iIconColor = iIconColorMouseHover
                Else
                    If iTabData.Hovered And mAmbientUserMode Then
                        iIconColor = iIconColorTabHighlighted
                    Else
                        iIconColor = iPropIconColor
                    End If
                End If
            Else
                If mAmbientUserMode And iTabData.Hovered Then
                    iIconColor = iIconColorTabHighlighted
                Else
                    iIconColor = iPropIconColor
                End If
            End If
        Else
            If iTabData.Hovered And mAmbientUserMode Then
                iIconColor = iIconColorTabHighlighted
            Else
                iIconColor = iPropIconColor
            End If
        End If
        iGrayText = mGrayText
    End If
    If Not (iTabData.Enabled And mEnabled) Then
        iIconColor = iGrayText
    End If
    If iTabData.IconFont Is Nothing Then
        Set iIconFont = mDefaultIconFont
    Else
        Set iIconFont = iTabData.IconFont
    End If
    iForeColor = TranslatedColor(iForeColor)
    iIconColor = TranslatedColor(iIconColor)
    
    If iTabData.Enabled And mEnabled Then
        picDraw.ForeColor = iForeColor
    Else
        picDraw.ForeColor = iGrayText
    End If
    iForeColor2 = picDraw.ForeColor
    
    iFontBoldPrev = picDraw.FontBold
    If iActive Then
        picDraw.FontBold = mFont.Bold Or mHighlightCaptionBoldSelectedTab
        picDraw.FontUnderLine = mFont.Underline Or mHighlightCaptionUnderlinedSelectedTab
    Else
        If iTabData.Hovered And mAmbientUserMode And mHighlightCaptionBold Then
            picDraw.FontBold = True
        Else
            picDraw.FontBold = mFont.Bold
        End If
        If iTabData.Hovered And mAmbientUserMode And mHighlightCaptionUnderlined Then
            picDraw.FontUnderLine = True
        Else
            picDraw.FontUnderLine = mFont.Underline
        End If
    End If
    
    iTabSpaceRect.Left = iTabRect.Left + 2
    If mAppearanceIsFlat Then iTabSpaceRect.Left = iTabSpaceRect.Left + 1
    iTabSpaceRect.Top = iTabRect.Top
    If mTabOrientation = ssTabOrientationBottom Then
        If iActive And Not mAppearanceIsFlat Then
            iTabSpaceRect.Bottom = iTabRect.Bottom + 2
        Else
            iTabSpaceRect.Bottom = iTabRect.Bottom + 4
        End If
    Else
        If mAppearanceIsFlat Then
            iTabSpaceRect.Bottom = iTabRect.Bottom
        Else
            iTabSpaceRect.Bottom = iTabRect.Bottom - 2
        End If
    End If
    iTabSpaceRect.Right = iTabRect.Right - 2
    
    If mAppearanceIsPP And iTabData.Selected Then
        iTabSpaceRect.Top = iTabSpaceRect.Top - 1
    End If
    
    If (Not iTabData.DoNotUseIconFont) And (iTabData.IconChar <> 0) Then
        iIconCharacter = ChrU(iTabData.IconChar)
        iIconCharRect.Left = 0
        iIconCharRect.Top = 0
        iIconCharRect.Right = 0
        iIconCharRect.Bottom = 0 ' iTabSpaceRect.Bottom
        iFlags = DT_CALCRECT Or DT_SINGLELINE Or DT_CENTER
        Set picAuxIconFont.Font = iIconFont
        iDo = True
        If mAmbientUserMode And (mTDIMode = ntTDIModeForms) Then iDo = Not mTDIFormWithoutCloseButton(iTabData.Data)
        If iDo Then
            DrawTextW picAuxIconFont.hDC, StrPtr(iIconCharacter), -1, iIconCharRect, iFlags Or IIf(mRightToLeft, DT_RTLREADING, 0)
            iPicWidth = (iIconCharRect.Right - iIconCharRect.Left)
            iPicHeight = (iIconCharRect.Bottom - iIconCharRect.Top)
            iDrawIcon = True
        End If
    ElseIf Not iTabData.PicToUse Is Nothing Then
        iDrawIcon = True
        If iTabData.Enabled Then
            Set iAuxPicture = iTabData.PicToUse
        Else
            If iTabData.PicToUse.Type = vbPicTypeBitmap Then
                If Not iTabData.PicDisabledSet Then
                    Set mTabData(nTab).PicDisabled = PictureToGrayScale(iTabData.PicToUse)
                End If
                Set iAuxPicture = mTabData(nTab).PicDisabled
            Else
                Set iAuxPicture = iTabData.PicToUse
            End If
        End If
        
        iPicWidth = pScaleX(iAuxPicture.Width, vbHimetric, vbPixels)
        iPicHeight = pScaleY(iAuxPicture.Height, vbHimetric, vbPixels)
        If (mTabOrientation = ssTabOrientationLeft) Then
            picAux.Width = iPicWidth
            picAux.Height = iPicHeight
            picAux.Cls
            picAux.BackColor = mBackColorTabs
            picRotate.Cls
            picAux.PaintPicture iAuxPicture, 0, 0
            RotatePic picAux, picRotate, nt90DegreesClockWise
            Set iAuxPicture = picRotate.Image
            picRotate.Cls
            picAux.Cls
            iPicWidth = pScaleX(iAuxPicture.Width, vbHimetric, vbPixels)
            iPicHeight = pScaleY(iAuxPicture.Height, vbHimetric, vbPixels)
        ElseIf (mTabOrientation = ssTabOrientationRight) Then
            picAux.Width = iPicWidth
            picAux.Height = iPicHeight
            picAux.Cls
            picAux.BackColor = mBackColorTabs
            picRotate.Cls
            picAux.PaintPicture iAuxPicture, 0, 0
            RotatePic picAux, picRotate, nt90DegreesCounterClockWise
            Set iAuxPicture = picRotate.Image
            picRotate.Cls
            picAux.Cls
            iPicWidth = pScaleX(iAuxPicture.Width, vbHimetric, vbPixels)
            iPicHeight = pScaleY(iAuxPicture.Height, vbHimetric, vbPixels)
        End If
    End If
    If iDrawIcon Then
        iIconAlignment = mIconAlignment
        If (mTDIMode = ntTDIModeForms) Or mIsTDIFormsSampleInPropertyPage Then
            If nTab = FindTab(0, ntFindOriginalIndex) Then
                iIconAlignment = ntIconAlignStart
                iIsTDIFormsFirstTab = True
            End If
        End If
        If (mTabOrientation = ssTabOrientationLeft) Then
            If iIconAlignment = ntIconAlignAfterCaption Then
                iIconAlignment = ntIconAlignBeforeCaption
            ElseIf iIconAlignment = ntIconAlignBeforeCaption Then
                iIconAlignment = ntIconAlignAfterCaption
            ElseIf iIconAlignment = ntIconAlignCenteredAfterCaption Then
                iIconAlignment = ntIconAlignCenteredBeforeCaption
            ElseIf iIconAlignment = ntIconAlignCenteredBeforeCaption Then
                iIconAlignment = ntIconAlignCenteredAfterCaption
            ElseIf iIconAlignment = ntIconAlignStart Then
                iIconAlignment = ntIconAlignEnd
            ElseIf iIconAlignment = ntIconAlignEnd Then
                iIconAlignment = ntIconAlignStart
            End If
        End If
    End If
    
    iTabSpaceWidth = (iTabSpaceRect.Right - iTabSpaceRect.Left) + 1
    iTabSpaceHeight = (iTabSpaceRect.Bottom - iTabSpaceRect.Top) + 1
    
    ' Calculate iMeasureRect for one liner and without elipsis for both cases, WordWrap or not
    iMeasureRect = iTabSpaceRect
    If mTDIMode = ntTDIModeForms Then
        If iTabData.Data > 0 Then
            If (mTDIFormIconCustomData(iTabData.Data).Width > 0) Then
                iTabSpaceRect.Left = iTabSpaceRect.Left + mTDIFormIconCustomData(iTabData.Data).Width * 1.06 - 4 '+ 6 - 4
            End If
        End If
    End If
    iMeasureRect.Bottom = iMeasureRect.Top + 5
    
    iFlags = DT_CALCRECT Or DT_SINGLELINE Or mAlignmentDTFlag
    iCaption = iTabData.Caption & IIf(picDraw.Font.Italic, "  ", "")
    DrawTextW picDraw.hDC, StrPtr(iCaption), -1, iMeasureRect, iFlags Or IIf(mRightToLeft, DT_RTLREADING, 0)
    iMeasureWidth = (iMeasureRect.Right - iMeasureRect.Left)
    
    If iDrawIcon Then
        If iPicWidth + iMeasureWidth + mTabIconDistanceToCaptionDPIScaled > iTabSpaceWidth Then
            If iPicWidth < iTabSpaceWidth / 2 Then
                iPicWidthToShow = iPicWidth
            Else
                If mWordWrap Then
                    If iPicWidth > iTabSpaceWidth * 0.67 Then
                        iPicWidthToShow = iTabSpaceWidth * 0.67
                    Else
                        iPicWidthToShow = iPicWidth
                    End If
                Else
                    If iPicWidth > iTabSpaceWidth * 0.5 Then
                        iPicWidthToShow = iTabSpaceWidth * 0.5
                    Else
                        iPicWidthToShow = iPicWidth
                    End If
                End If
            End If
            If iPicWidthToShow + iMeasureWidth + mTabIconDistanceToCaptionDPIScaled < iTabSpaceWidth Then
                iPicWidthToShow = iTabSpaceWidth - iMeasureWidth - mTabIconDistanceToCaptionDPIScaled
            End If
            If iPicWidthToShow > iPicWidth Then
                iPicWidthToShow = iPicWidth
            End If
        Else
            iPicWidthToShow = iPicWidth
        End If
    End If
    
    If iPicHeight > iTabSpaceHeight Then
        iPicHeightToShow = iTabSpaceHeight
    Else
        iPicHeightToShow = iPicHeight
    End If
    
    iMeasureRect.Right = iTabSpaceRect.Right
    If iDrawIcon And ((iIconAlignment <> ntIconAlignCenteredOnTab) And (iIconAlignment <> ntIconAlignAtTop) And (iIconAlignment <> ntIconAlignAtBottom)) Then
        iMeasureRect.Left = iTabSpaceRect.Left + iPicWidthToShow + IIf(iIsTDIFormsFirstTab, 0, mTabIconDistanceToCaptionDPIScaled)
    Else
        iMeasureRect.Left = iTabSpaceRect.Left
    End If
    
    iMeasureRect.Top = 0
    iMeasureRect.Bottom = iMeasureRect.Top + 5
    
    iCaptionRect.Left = iMeasureRect.Left
    iCaptionRect.Right = iMeasureRect.Right
    ' Calculate iMeasureRect again, without elipsis for WordWrap and with elipsis for single line, and without both text centering
    If mWordWrap Then
        iFlags = DT_CALCRECT Or DT_WORDBREAK
    Else
        iFlags = DT_CALCRECT Or DT_SINGLELINE Or DT_END_ELLIPSIS Or DT_MODIFYSTRING
    End If
    iCaption = iTabData.Caption & IIf(picDraw.Font.Italic, "  ", "")

    DrawTextW picDraw.hDC, StrPtr(iCaption & IIf(iFlags And DT_MODIFYSTRING, "    ", "")), IIf(iFlags And DT_MODIFYSTRING, Len(iCaption), -1), iMeasureRect, iFlags Or IIf(mRightToLeft, DT_RTLREADING, 0)
'    If nTab = 0 Then Stop
    iMeasureWidth = (iMeasureRect.Right - iMeasureRect.Left)
    iMeasureHeight = (iMeasureRect.Bottom - iMeasureRect.Top)
    
    If iDrawIcon Then
        If (iIconAlignment = ntIconAlignAfterCaption) Or (iIconAlignment = ntIconAlignCenteredAfterCaption) Or (iIconAlignment = ntIconAlignEnd) Then
            iLng = iTabSpaceRect.Right - iPicWidthToShow - mTabIconDistanceToCaptionDPIScaled
            iCaptionRect.Left = iCaptionRect.Left - iCaptionRect.Right + iLng
            iCaptionRect.Right = iLng
        End If
    End If
    
    If (iIconAlignment = ntIconAlignAtBottom) Then
        iCaptionRect.Top = iTabSpaceRect.Top + iFlatBarHeightTop
    ElseIf (iIconAlignment = ntIconAlignAtTop) Then
        iCaptionRect.Bottom = iTabSpaceRect.Bottom - iFlatBarHeightBottom
        iCaptionRect.Top = iCaptionRect.Bottom - iMeasureHeight
    Else
        iCaptionRect.Top = iTabSpaceRect.Top + (iTabSpaceHeight - iFlatBarHeightTop - iFlatBarHeightBottom) / 2 - iMeasureHeight / 2 + iFlatBarHeightTop
    End If
    iCaptionRect.Bottom = iCaptionRect.Top + iMeasureHeight + 1
    
    If mTabOrientation = ssTabOrientationBottom Then
        iGMPrev = SetGraphicsMode(picDraw.hDC, GM_ADVANCED)
        iTx1.eM11 = 1: iTx1.eM22 = -1: iTx1.eDx = 0: iTx1.eDy = iTabSpaceRect.Top + iFlatBarHeightTop / 2 + iTabSpaceRect.Bottom
        SetWorldTransform picDraw.hDC, iTx1
    End If
    
    If iDrawIcon Then
        If iIconAlignment = ntIconAlignAtBottom Then
            iPicTop = (iCaptionRect.Bottom - iFlatBarHeightBottom + iTabSpaceRect.Bottom - iPicHeight) / 2
        ElseIf iIconAlignment = ntIconAlignAtTop Then
            iPicTop = (iTabSpaceRect.Top + iCaptionRect.Top - iPicHeight) / 2
        Else
            iPicTop = iTabSpaceRect.Top + (iTabSpaceHeight - iPicHeightToShow - iFlatBarHeightTop - iFlatBarHeightBottom) / 2 + 0.49 + iFlatBarHeightTop
        End If
        ' Position of pic
        If iIconAlignment = ntIconAlignStart Then
            iPicLeft = iTabSpaceRect.Left + 4.5 * mDPIScale
        ElseIf iIconAlignment = ntIconAlignEnd Then
            iPicLeft = iTabSpaceRect.Right - iPicWidthToShow - 4.5 * mDPIScale
        ElseIf (iIconAlignment = ntIconAlignCenteredOnTab) Or (iIconAlignment = ntIconAlignAtTop) Or (iIconAlignment = ntIconAlignAtBottom) Then
            iPicLeft = (iTabSpaceRect.Right + iTabSpaceRect.Left - iPicWidthToShow) / 2
        ElseIf iTabData.Caption <> "" Then
            If iIconAlignment = ntIconAlignBeforeCaption Then
                iPicLeft = (iCaptionRect.Right + iCaptionRect.Left) / 2 - iMeasureWidth / 2 - mTabIconDistanceToCaptionDPIScaled - iPicWidthToShow
            ElseIf iIconAlignment = ntIconAlignAfterCaption Then
                iPicLeft = (iCaptionRect.Right + iCaptionRect.Left) / 2 + iMeasureWidth / 2 + mTabIconDistanceToCaptionDPIScaled
            ElseIf iIconAlignment = ntIconAlignCenteredBeforeCaption Then
                iPicLeft = iTabSpaceRect.Left + (((iCaptionRect.Right + iCaptionRect.Left) / 2 - iMeasureWidth / 2) - iTabSpaceRect.Left) / 2 - iPicWidthToShow / 2
            ElseIf iIconAlignment = ntIconAlignCenteredAfterCaption Then
                iPicLeft = iTabSpaceRect.Right - (iTabSpaceRect.Right - ((iCaptionRect.Right + iCaptionRect.Left) / 2 + iMeasureWidth / 2)) / 2 - iPicWidthToShow / 2
            End If
        Else
            iPicLeft = (iTabSpaceRect.Right + iTabSpaceRect.Left) / 2 - iPicWidthToShow / 2
        End If
        If iPicLeft < iTabSpaceRect.Left Then
            iPicLeft = iTabSpaceRect.Left
        End If
        If (iPicLeft + iPicWidthToShow) > iTabSpaceRect.Right Then
            iPicLeft = iTabSpaceRect.Right - iPicWidthToShow
        End If
        
        If iPicHeightToShow >= iPicHeight Then
            iPicSourceShiftY = 0
        Else
            iPicSourceShiftY = (iPicHeight - iPicHeightToShow) / 2
        End If
        If iPicWidthToShow >= iPicWidth Then
            iPicSourceShiftX = 0
        Else
            iPicSourceShiftX = (iPicWidth - iPicWidthToShow) / 2
        End If
        
        If iPicWidth < 1 Then iPicWidth = 1
        If iPicHeight < 1 Then iPicHeight = 1
        
        ' draw the icon
        If iIconCharacter <> "" Then
            iFlags = DT_SINGLELINE Or DT_CENTER
            Set iFontPrev = picDraw.Font
            Set picDraw.Font = iIconFont
            iForeColorPrev = picDraw.ForeColor
            picDraw.ForeColor = iIconColor
            iLng = ((iTabRect.Bottom - iTabRect.Top) - (iIconCharRect.Bottom - iIconCharRect.Top)) / 2
            iIconCharRect.Left = iPicLeft + iTabData.IconLeftOffset * mDPIScale
            iIconCharRect.Right = iPicLeft + iTabData.IconLeftOffset * mDPIScale + iPicWidth
            iIconCharRect.Top = iPicTop + iTabData.IconTopOffset * mDPIScale
            iIconCharRect.Bottom = iPicTop + iTabData.IconTopOffset * mDPIScale + iPicHeight + 1
            If (mIconAlignment = ntIconAlignAfterCaption) Or (mIconAlignment = ntIconAlignCenteredAfterCaption) Or (mIconAlignment = ntIconAlignEnd) Then
                If iIconCharRect.Right > (iTabRect.Right - 4) Then
                    iLng = iIconCharRect.Right - (iTabRect.Right - 4)
                    iIconCharRect.Left = iIconCharRect.Left - iLng
                    iIconCharRect.Right = iIconCharRect.Right - iLng
                End If
            ElseIf (mIconAlignment = ntIconAlignBeforeCaption) Or (mIconAlignment = ntIconAlignCenteredBeforeCaption) Or (mIconAlignment = ntIconAlignStart) Then
                If iIconCharRect.Left < (iTabRect.Left + 5) Then
                    iLng = (iTabRect.Left + 5) - iIconCharRect.Left
                    iIconCharRect.Left = iIconCharRect.Left + iLng
                    iIconCharRect.Right = iIconCharRect.Right + iLng
                End If
            End If
            If (mTabOrientation = ssTabOrientationLeft) Or (mTabOrientation = ssTabOrientationRight) Then
                iGMPrev2 = SetGraphicsMode(picDraw.hDC, GM_ADVANCED)
                GetWorldTransform picDraw.hDC, iTx2Prev
                If (mTabOrientation = ssTabOrientationLeft) Then
                    iTx2.eM11 = 0
                    iTx2.eM12 = 1
                    iTx2.eM21 = -1
                    iTx2.eM22 = 0
                    iTx2.eDx = (iIconCharRect.Left + iIconCharRect.Right) / 2 + iPicWidth / 2
                    iTx2.eDy = (iIconCharRect.Top + iIconCharRect.Bottom) / 2 - iPicHeight / 2
                    iAuxRect = iIconCharRect
                    iIconCharRect.Left = 0
                    iIconCharRect.Top = 0
                    iIconCharRect.Right = iPicHeight
                    iIconCharRect.Bottom = iPicWidth + 1
'                ElseIf mTabOrientation = ssTabOrientationBottom Then
'                    iTx2.eM11 = 1
'                    iTx2.eM12 = 0
'                    iTx2.eM21 = 0
'                    iTx2.eM22 = -1
'                    iTx2.eDx = (iIconCharRect.Left + iIconCharRect.Right) / 2 - iPicWidth / 2
'                    iTx2.eDy = (iIconCharRect.Top + iIconCharRect.Bottom) / 2 + iPicHeight / 2
'                    iLng = iIconCharRect.Left
'                    iIconCharRect.Left = 0
'                    iIconCharRect.Top = 0
'                    iIconCharRect.Right = iPicHeight
'                    iIconCharRect.Bottom = iPicWidth + 1
                Else
                    iTx2.eM11 = 0
                    iTx2.eM12 = -1
                    iTx2.eM21 = 1
                    iTx2.eM22 = 0
                    iTx2.eDx = (iIconCharRect.Left + iIconCharRect.Right) / 2 - iPicWidth / 2
                    iTx2.eDy = (iIconCharRect.Top + iIconCharRect.Bottom) / 2 + iPicHeight / 2
                    iAuxRect = iIconCharRect
                    iIconCharRect.Left = 0
                    iIconCharRect.Top = 0
                    iIconCharRect.Right = iPicHeight
                    iIconCharRect.Bottom = iPicWidth + 1
                End If
                SetWorldTransform picDraw.hDC, iTx2
            End If
            If mTabsAreRotatedButCaptionsAreHorizontal Then
                iAuxRect = iIconCharRect
                iIconCharRect.Right = iIconCharRect.Right - iTabCenterX
                iIconCharRect.Left = iIconCharRect.Left - iTabCenterX
                iIconCharRect.Bottom = iIconCharRect.Bottom - iTabCenterY
                iIconCharRect.Top = iIconCharRect.Top - iTabCenterY
            End If
            DrawTextW picDraw.hDC, StrPtr(iIconCharacter), -1, iIconCharRect, iFlags Or IIf(mRightToLeft, DT_RTLREADING, 0)
            If (mTabOrientation = ssTabOrientationLeft) Or (mTabOrientation = ssTabOrientationRight) Then
                SetWorldTransform picDraw.hDC, iTx2Prev
                SetGraphicsMode picDraw.hDC, iGMPrev2
                iIconCharRect = iAuxRect
            End If
            If mTabsAreRotatedButCaptionsAreHorizontal Then
                iIconCharRect = RotateRect90Degrees(iAuxRect, False, iTabCenterX, iTabCenterY)
                If mTabOrientation = ntTabOrientationRightHorizontal Then
                    iIconCharRect.Top = (iTabRect.Right - iTabRect.Left) - iIconCharRect.Top
                    iIconCharRect.Bottom = (iTabRect.Right - iTabRect.Left) - iIconCharRect.Bottom
                Else
                    iLng = iIconCharRect.Top: iIconCharRect.Top = iIconCharRect.Bottom: iIconCharRect.Bottom = iLng
                End If
                iLng = iIconCharRect.Left: iIconCharRect.Left = iIconCharRect.Right: iIconCharRect.Right = iLng
            End If
            mTabData(nTab).IconRect = iIconCharRect
            Set picDraw.Font = iFontPrev
            picDraw.ForeColor = iForeColorPrev
        Else
            If mTabsAreRotatedButCaptionsAreHorizontal Then
                iPicLeft = iPicLeft - iTabCenterX
                iPicTop = iPicTop - iTabCenterY
            End If
            If mRightToLeft Then
                SetLayout GetDC(picDraw.hWnd), LAYOUT_RTL Or LAYOUT_BITMAPORIENTATIONPRESERVED
            End If
            If iAuxPicture.Type = vbPicTypeBitmap And mUseMaskColor Then
                Call DrawImage(picDraw.hDC, iAuxPicture.Handle, TranslatedColor(mMaskColor), iPicLeft, iPicTop, iPicWidthToShow, iPicHeightToShow, iPicSourceShiftX, iPicSourceShiftY)
            Else
                picDraw.PaintPicture iAuxPicture, iPicLeft, iPicTop, iPicWidthToShow, iPicHeightToShow, iPicSourceShiftX, iPicSourceShiftY, iPicWidthToShow, iPicHeightToShow
            End If
            If mRightToLeft Then
                SetLayout GetDC(picDraw.hWnd), LAYOUT_RTL
            End If
        End If
    End If
    'Now draw the text
    If mWordWrap Then
        iFlags = DT_WORDBREAK Or DT_END_ELLIPSIS Or DT_MODIFYSTRING Or mAlignmentDTFlag
    Else
        iFlags = DT_SINGLELINE Or DT_END_ELLIPSIS Or DT_MODIFYSTRING Or mAlignmentDTFlag Or DT_VCENTER
    End If
    
    iCaption = iTabData.Caption
    If iCaptionRect.Bottom > iTabRect.Bottom Then
        iCaptionRect.Bottom = iTabRect.Bottom
    End If
    picDraw.ForeColor = iForeColor2
    If iCaptionRect.Right - iCaptionRect.Left < 8 Then iCaptionRect.Right = iCaptionRect.Left + 8 ' protection, because this API crashes the program in this case
    
    If mTabsAreRotatedButCaptionsAreHorizontal Then
        iCaptionRect.Right = iCaptionRect.Right - iTabCenterX
        iCaptionRect.Left = iCaptionRect.Left - iTabCenterX
        iCaptionRect.Bottom = iCaptionRect.Bottom - iTabCenterY
        iCaptionRect.Top = iCaptionRect.Top - iTabCenterY
    End If
    
    If mTDIMode = ntTDIModeForms Then
        If (iTabData.Data <> 0) Then
            If (mTDIFormIconCustomData(iTabData.Data).Width > 0) Then
                iTDIFormsNoIcons = True
                iTDIFormIcon_Left = iTabRect.Left + 6 + mTDIFormIconCustomData(iTabData.Data).Width * 0.03
                iTDIFormIcon_Top = iCaptionRect.Top + (iCaptionRect.Bottom - iCaptionRect.Top) / 2 - mTDIFormIconCustomData(iTabData.Data).Height / 2 + 0.49
                If mTDIFormIconCustomData(iTabData.Data).UseIconFont Then
                    iTDIFormIcon_Top = iTDIFormIcon_Top + 1.49
                End If
                iTDIFormIcon_Left = iTDIFormIcon_Left + mTDIFormIconCustomData(iTabData.Data).LeftOffset
                iTDIFormIcon_Top = iTDIFormIcon_Top + mTDIFormIconCustomData(iTabData.Data).TopOffset
                If mTDIFormIconCustomData(iTabData.Data).UseIconFont Then
                    iTDIFormIcon_IconCharRect.Left = iTDIFormIcon_Left
                    iTDIFormIcon_IconCharRect.Top = iTDIFormIcon_Top
                    iTDIFormIcon_IconCharRect.Right = iTDIFormIcon_Left + mTDIFormIconCustomData(iTabData.Data).Width
                    iTDIFormIcon_IconCharRect.Bottom = iTabRect.Bottom ' iTDIFormIcon_Top + mTDIFormIconCustomData(iTabData.Data).Height
                    
                    iForeColorPrev = picDraw.ForeColor
                    If mAmbientUserMode And iTabData.Hovered And (Not iActive) And iTabData.Enabled Then
                        picDraw.ForeColor = mTDIFormIconCustomData(iTabData.Data).IconColorMouseHover
                    ElseIf mAmbientUserMode And iActive And iTabData.Enabled Then
                        picDraw.ForeColor = mTDIFormIconCustomData(iTabData.Data).IconColorSelected
                    Else
                        picDraw.ForeColor = mTDIFormIconCustomData(iTabData.Data).IconColor
                    End If
                    Set iTDIFormsFontPrev = picDraw.Font
                    Set picDraw.Font = mTDIFormIconCustomData(iTabData.Data).IconFont
                    DrawTextW picDraw.hDC, StrPtr(mTDIFormIconCustomData(iTabData.Data).IconCharStr), -1, iTDIFormIcon_IconCharRect, DT_SINGLELINE Or DT_CENTER Or IIf(mRightToLeft, DT_RTLREADING, 0)
                    Set picDraw.Font = iTDIFormsFontPrev
                    picDraw.ForeColor = iForeColorPrev
                    
                    iTDIFormsNoIcons = False
                ElseIf (Not mTDIFormIconCustomData(iTabData.Data).Icon Is Nothing) Then
                    If Not (mTDIFormIconCustomData(iTabData.Data).Icon.Handle = 0) Then
                        If mTDIFormIconCustomData(iTabData.Data).Icon.Type = vbPicTypeBitmap And mUseMaskColor Then
                            Call DrawImage(picDraw.hDC, mTDIFormIconCustomData(iTabData.Data).Icon.Handle, TranslatedColor(mMaskColor), iTDIFormIcon_Left, iTDIFormIcon_Top, mTDIFormIconCustomData(iTabData.Data).Width, mTDIFormIconCustomData(iTabData.Data).Height, iPicSourceShiftX, iPicSourceShiftY)
                        Else
                            picDraw.PaintPicture mTDIFormIconCustomData(iTabData.Data).Icon, iTDIFormIcon_Left, iTDIFormIcon_Top, mTDIFormIconCustomData(iTabData.Data).Width, mTDIFormIconCustomData(iTabData.Data).Height, 1, 1, UserControl.ScaleX(mTDIFormIconCustomData(iTabData.Data).Icon.Width, vbHimetric, vbPixels), UserControl.ScaleY(mTDIFormIconCustomData(iTabData.Data).Icon.Height, vbHimetric, vbPixels)
                        End If
                        iTDIFormsNoIcons = False
                    End If
                End If
                If iTDIFormsNoIcons And (Not mTDIModeFormsFormData_FormIcon(iTabData.Data) Is Nothing) Then
                    iCaptionRect.Left = iCaptionRect.Left - mTDIFormIconCustomData(iTabData.Data).Width * 1.06 + 4
                End If
            End If
        End If
    End If
    
    DrawTextW picDraw.hDC, StrPtr(iCaption & IIf(iFlags And DT_MODIFYSTRING, "    ", "")), IIf(iFlags And DT_MODIFYSTRING, Len(iCaption), -1), iCaptionRect, iFlags Or IIf(mRightToLeft, DT_RTLREADING, 0)
    
    ' Draw the focus rect
    If mAmbientUserMode Then    'only at run time
        If (iActive) And ControlHasFocus And mShowFocusRect Then
            If mTabsAreRotatedButCaptionsAreHorizontal Then
                If mAppearanceIsPP Then
                    iFocusRect = iTabRect
                    
                    iLng = iFocusRect.Top
                    iFocusRect.Top = iFocusRect.Left
                    iFocusRect.Left = iLng
                    iLng = iFocusRect.Bottom
                    iFocusRect.Bottom = iFocusRect.Right
                    iFocusRect.Right = iLng
                    
                    iFocusRect.Left = iFocusRect.Left + 3
                    iFocusRect.Top = iFocusRect.Top + 4
                    iFocusRect.Right = iFocusRect.Right - 2

                    iFocusRect.Right = (iFocusRect.Right - iFocusRect.Left) / 2
                    iFocusRect.Left = -iFocusRect.Right
                    iFocusRect.Bottom = (iFocusRect.Bottom - iFocusRect.Top) / 2
                    iFocusRect.Top = -iFocusRect.Bottom
                Else
                    iFocusRect.Left = -iMeasureWidth / 2 - 2
                    iFocusRect.Right = iFocusRect.Left + iMeasureWidth + 3
                    iFocusRect.Top = iCaptionRect.Top - 1
                    iFocusRect.Bottom = iFocusRect.Top + iMeasureHeight + 2
                End If
                picDraw.ForeColor = iForeColor
                If mAppearanceIsPP Then
                    iFocusRect.Top = iFocusRect.Top - 1
                    iFocusRect.Bottom = iFocusRect.Bottom - 1
                End If
                
                If iFocusRect.Right > ((iTabSpaceRect.Bottom - iTabSpaceRect.Top) / 2) Then
                    iFocusRect.Right = ((iTabSpaceRect.Bottom - iTabSpaceRect.Top) / 2)
                End If
                If iFocusRect.Left < ((-(iTabSpaceRect.Bottom - iTabSpaceRect.Top) / 2) + 1) Then
                    iFocusRect.Left = (-(iTabSpaceRect.Bottom - iTabSpaceRect.Top) / 2) + 1
                End If
                If iFocusRect.Bottom > ((iTabSpaceRect.Right - iTabSpaceRect.Left) / 2) Then
                    iFocusRect.Bottom = ((iTabSpaceRect.Right - iTabSpaceRect.Left) / 2)
                End If
                If iFocusRect.Top < ((-(iTabSpaceRect.Right - iTabSpaceRect.Left) / 2) + 1) Then
                    iFocusRect.Top = (-(iTabSpaceRect.Right - iTabSpaceRect.Left) / 2) + 1
                End If
            Else
                If mAppearanceIsPP Then
                    iFocusRect = iTabRect
                    iFocusRect.Left = iFocusRect.Left + 3
                    iFocusRect.Top = iFocusRect.Top + 4
                    iFocusRect.Right = iFocusRect.Right - 2
                    If (mTabOrientation = ssTabOrientationLeft) Then
                        iFocusRect.Left = iFocusRect.Left + 1
                        iFocusRect.Right = iFocusRect.Right + 1
                    End If
                Else
                    iFocusRect.Left = (iCaptionRect.Left + iCaptionRect.Right) / 2 - iMeasureWidth / 2 - 2
                    iFocusRect.Right = iFocusRect.Left + iMeasureWidth + 3
                    iFocusRect.Top = iCaptionRect.Top - 1
                    iFocusRect.Bottom = iFocusRect.Top + iMeasureHeight + 2
                End If
                picDraw.ForeColor = iForeColor
                If mAppearanceIsPP Then
                    iFocusRect.Top = iFocusRect.Top - 1
                    iFocusRect.Bottom = iFocusRect.Bottom - 1
                End If
                
                If iFocusRect.Right > (iTabSpaceRect.Right) Then
                    iFocusRect.Right = iTabSpaceRect.Right
                End If
                If iFocusRect.Left < (iTabSpaceRect.Left + 1) Then
                    iFocusRect.Left = iTabSpaceRect.Left + 1
                End If
                If iFocusRect.Bottom > (iTabSpaceRect.Bottom) Then
                    iFocusRect.Bottom = iTabSpaceRect.Bottom
                End If
                If iFocusRect.Top < (iTabSpaceRect.Top + 1) Then
                    iFocusRect.Top = iTabSpaceRect.Top + 1
                End If
            End If
            
            Call DrawFocusRect(picDraw.hDC, iFocusRect)
        End If
    End If

    If mTabsAreRotatedButCaptionsAreHorizontal Then
        SetWorldTransform picDraw.hDC, iTx2Prev
        SetGraphicsMode picDraw.hDC, iGMPrev2
    End If

    If mTabOrientation = ssTabOrientationBottom Then
        iTx1.eM11 = 1: iTx1.eM22 = 1: iTx1.eDx = 0: iTx1.eDy = 0
        SetWorldTransform picDraw.hDC, iTx1
        SetGraphicsMode picDraw.hDC, iGMPrev
    End If

    If picDraw.FontBold <> iFontBoldPrev Then
        picDraw.FontBold = iFontBoldPrev
    End If
End Sub

Private Function RotateRect90Degrees(nRect As RECT, Optional nCCW As Boolean, Optional nCenterX As Variant, Optional nCenterY As Variant) As RECT
    Dim iCenterX As Single
    Dim iCenterY As Single
    Const pi = 3.141592654
    Dim iAngleInRadians As Single
    Dim iCos As Single
    Dim iSin As Single
    
    If nCCW Then
        iAngleInRadians = -90 * pi / 180
    Else
        iAngleInRadians = 90 * pi / 180
    End If
    iCos = Cos(iAngleInRadians)
    iSin = sIn(iAngleInRadians)
    
    If IsMissing(nCenterX) Then nCenterX = (nRect.Left + nRect.Right) / 2
    If IsMissing(nCenterY) Then nCenterY = (nRect.Top + nRect.Bottom) / 2
    iCenterX = CSng(nCenterX)
    iCenterY = CSng(nCenterY)
    
    RotateRect90Degrees.Left = (nRect.Left - iCenterX) * iCos - (nRect.Top - iCenterY) * iSin + iCenterX
    RotateRect90Degrees.Bottom = (nRect.Left - iCenterX) * iSin + (nRect.Top - iCenterY) * iCos + iCenterY
    
    RotateRect90Degrees.Right = (nRect.Right - iCenterX) * iCos - (nRect.Bottom - iCenterY) * iSin + iCenterX
    RotateRect90Degrees.Top = (nRect.Right - iCenterX) * iSin + (nRect.Bottom - iCenterY) * iCos + iCenterY
End Function

' The following procedure was taken from http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=56462&lngWId=1
' Kinda over-riden function for pFillCurvedGradientR, performs same job,
' but takes integers instead of Rect as parameter
Private Sub FillCurvedGradient(ByVal lLeft As Long, ByVal lTop As Long, ByVal lRight As Long, ByVal lBottom As Long, ByVal lStartColor As Long, ByVal lEndColor As Long, Optional ByVal iCurveValue As Long = -1, Optional bCurveLeft As Boolean = False, Optional bCurveRight As Boolean = False)
    Dim utRect As RECT
    
    utRect.Left = lLeft
    utRect.Top = lTop
    utRect.Right = lRight
    utRect.Bottom = lBottom
    
    If utRect.Bottom - utRect.Top > 0 Then Call FillCurvedGradientR(utRect, lStartColor, lEndColor, iCurveValue, bCurveLeft, bCurveRight)
End Sub

' The following procedure was taken from http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=56462&lngWId=1
' function used to Fill a rectangular area by Gradient
' This function can draw using the curve value to generate a rounded rect kinda effect
Private Sub FillCurvedGradientR(utRect As RECT, ByVal lStartColor As Long, ByVal lEndColor As Long, Optional ByVal iCurveValue As Long = -1, Optional bCurveLeft As Boolean = False, Optional bCurveRight As Boolean = False)

    Dim sngRedInc As Single, sngGreenInc As Single, sngBlueInc As Single
    Dim sngRed As Single, sngGreen As Single, sngBlue As Single
    
    lStartColor = TranslatedColor(lStartColor)
    lEndColor = TranslatedColor(lEndColor)

    Dim intCnt As Integer
    
    sngRed = GetRedValue(lStartColor)
    sngGreen = GetGreenValue(lStartColor)
    sngBlue = GetBlueValue(lStartColor)
    
    sngRedInc = (GetRedValue(lEndColor) - sngRed) / (utRect.Bottom - utRect.Top)
    sngGreenInc = (GetGreenValue(lEndColor) - sngGreen) / (utRect.Bottom - utRect.Top)
    sngBlueInc = (GetBlueValue(lEndColor) - sngBlue) / (utRect.Bottom - utRect.Top)

    If sngRed > 255 Then sngRed = 255
    If sngGreen > 255 Then sngGreen = 255
    If sngBlue > 255 Then sngBlue = 255
    If sngRed < 0 Then sngRed = 0
    If sngGreen < 0 Then sngGreen = 0
    If sngBlue < 0 Then sngBlue = 0

    If iCurveValue < 1 Then
        For intCnt = utRect.Top To utRect.Bottom
            picDraw.Line (utRect.Left, intCnt)-(utRect.Right, intCnt), RGB(sngRed, sngGreen, sngBlue)
            sngRed = sngRed + sngRedInc
            sngGreen = sngGreen + sngGreenInc
            sngBlue = sngBlue + sngBlueInc
            
            If sngRed > 255 Then sngRed = 255
            If sngGreen > 255 Then sngGreen = 255
            If sngBlue > 255 Then sngBlue = 255
            If sngRed < 0 Then sngRed = 0
            If sngGreen < 0 Then sngGreen = 0
            If sngBlue < 0 Then sngBlue = 0
        Next
    Else
        If bCurveLeft And bCurveRight Then
            For intCnt = utRect.Top To utRect.Bottom
                picDraw.Line (utRect.Left + iCurveValue + 1, intCnt)-(utRect.Right - iCurveValue, intCnt), RGB(sngRed, sngGreen, sngBlue)
                sngRed = sngRed + sngRedInc
                sngGreen = sngGreen + sngGreenInc
                sngBlue = sngBlue + sngBlueInc

                If sngRed > 255 Then sngRed = 255
                If sngGreen > 255 Then sngGreen = 255
                If sngBlue > 255 Then sngBlue = 255
                If sngRed < 0 Then sngRed = 0
                If sngGreen < 0 Then sngGreen = 0
                If sngBlue < 0 Then sngBlue = 0

                If iCurveValue > 0 Then
                    iCurveValue = iCurveValue - 1
                End If
            Next
        ElseIf bCurveLeft Then
            For intCnt = utRect.Top To utRect.Bottom
                picDraw.Line (utRect.Left + iCurveValue + 1, intCnt)-(utRect.Right, intCnt), RGB(sngRed, sngGreen, sngBlue)

                sngRed = sngRed + sngRedInc
                sngGreen = sngGreen + sngGreenInc
                sngBlue = sngBlue + sngBlueInc

                If sngRed > 255 Then sngRed = 255
                If sngGreen > 255 Then sngGreen = 255
                If sngBlue > 255 Then sngBlue = 255
                If sngRed < 0 Then sngRed = 0
                If sngGreen < 0 Then sngGreen = 0
                If sngBlue < 0 Then sngBlue = 0

                If iCurveValue > 0 Then
                    iCurveValue = iCurveValue - 1
                End If
            Next
        Else    'curve right
            For intCnt = utRect.Top To utRect.Bottom
                picDraw.Line (utRect.Left, intCnt)-(utRect.Right - iCurveValue, intCnt), RGB(sngRed, sngGreen, sngBlue)

                sngRed = sngRed + sngRedInc
                sngGreen = sngGreen + sngGreenInc
                sngBlue = sngBlue + sngBlueInc
                
                If sngRed > 255 Then sngRed = 255
                If sngGreen > 255 Then sngGreen = 255
                If sngBlue > 255 Then sngBlue = 255
                If sngRed < 0 Then sngRed = 0
                If sngGreen < 0 Then sngGreen = 0
                If sngBlue < 0 Then sngBlue = 0
                
                If iCurveValue > 0 Then
                    iCurveValue = iCurveValue - 1
                End If
            Next
        End If
    End If
End Sub

Private Sub FillCurvedGradient2(ByVal nLeft As Long, ByVal nTop As Long, ByVal nRight As Long, ByVal nBottom As Long, ByVal nStartColor As Long, ByVal nEndColor As Long, ByVal nCurveTopLeft As Long, ByVal nCurveTopRight As Long, Optional ByVal nCurveBottomLeft As Long, Optional ByVal nCurveBottomRight As Long)
    Dim sngRedInc As Single, sngGreenInc As Single, sngBlueInc As Single
    Dim sngRed As Single, sngGreen As Single, sngBlue As Single
    Dim iCurvPixelsTopY As Long
    Dim iCurvPixelsBottomY As Long
    Dim iCurvPixelsXLeft As Single
    Dim iCurvPixelsXRight As Single
    Dim iCurvPixelsXLeftInt As Long
    Dim iCurvPixelsXRightInt As Long
    Dim iX As Long
    Dim iY As Long
    Dim iDistance As Single
    Dim iCol As Long
    
    If (nBottom - nTop) <= 0 Then Exit Sub
    If (nRight - nLeft) <= 0 Then Exit Sub
    
    nStartColor = TranslatedColor(nStartColor)
    nEndColor = TranslatedColor(nEndColor)
    
    sngRed = GetRedValue(nStartColor)
    sngGreen = GetGreenValue(nStartColor)
    sngBlue = GetBlueValue(nStartColor)
    
    sngRedInc = (GetRedValue(nEndColor) - sngRed) / (nBottom - nTop)
    sngGreenInc = (GetGreenValue(nEndColor) - sngGreen) / (nBottom - nTop)
    sngBlueInc = (GetBlueValue(nEndColor) - sngBlue) / (nBottom - nTop)

    If sngRed > 255 Then sngRed = 255
    If sngGreen > 255 Then sngGreen = 255
    If sngBlue > 255 Then sngBlue = 255
    If sngRed < 0 Then sngRed = 0
    If sngGreen < 0 Then sngGreen = 0
    If sngBlue < 0 Then sngBlue = 0
    
    If (nCurveTopLeft < 1) And (nCurveTopRight < 1) And (nCurveBottomLeft < 1) And (nCurveBottomRight < 1) Then
        For iY = nTop To nBottom - 1
            picDraw.Line (nLeft, iY)-(nRight, iY), RGB(sngRed, sngGreen, sngBlue)
            sngRed = sngRed + sngRedInc
            sngGreen = sngGreen + sngGreenInc
            sngBlue = sngBlue + sngBlueInc
            
            If sngRed > 255 Then sngRed = 255
            If sngGreen > 255 Then sngGreen = 255
            If sngBlue > 255 Then sngBlue = 255
            If sngRed < 0 Then sngRed = 0
            If sngGreen < 0 Then sngGreen = 0
            If sngBlue < 0 Then sngBlue = 0
        Next
    Else
        If nCurveTopLeft > nCurveTopRight Then
            iCurvPixelsTopY = nCurveTopLeft
        Else
            iCurvPixelsTopY = nCurveTopRight
        End If
        If iCurvPixelsTopY > (nBottom - nTop) Then
            iCurvPixelsTopY = (nBottom - nTop)
        End If
        For iY = nTop To nTop + iCurvPixelsTopY - 1
            If (iY - nTop) <= nCurveTopLeft Then
                iCurvPixelsXLeft = nCurveTopLeft - Sqr(nCurveTopLeft ^ 2 - (nCurveTopLeft - iY + nTop) ^ 2)
                iCurvPixelsXLeftInt = Round(iCurvPixelsXLeft) + 1
            Else
                iCurvPixelsXLeft = 0
                iCurvPixelsXLeftInt = 0
            End If
            If (iY - nTop) <= nCurveTopRight Then
                iCurvPixelsXRight = nCurveTopRight - Sqr(nCurveTopRight ^ 2 - (nCurveTopRight - iY + nTop) ^ 2)
                iCurvPixelsXRightInt = Round(iCurvPixelsXRight) + 1
            Else
                iCurvPixelsXRight = 0
                iCurvPixelsXRightInt = 0
            End If
            
            iCol = RGB(sngRed, sngGreen, sngBlue)
            If (iCurvPixelsXLeftInt > 0) Or (iCurvPixelsXRightInt > 0) Then
                picDraw.Line (nLeft + iCurvPixelsXLeftInt, iY)-(nRight - iCurvPixelsXRightInt, iY), iCol
                If iCurvPixelsXLeftInt > 0 Then
                    For iX = iCurvPixelsXLeftInt - 1 To iCurvPixelsXLeftInt \ 2 - 1 Step -1
                        If iX < 0 Then Exit For
                        iDistance = nCurveTopLeft - (Sqr((nCurveTopLeft - iX) ^ 2 + (nCurveTopLeft - iY + nTop) ^ 2))
                        If iDistance > 0 Then iDistance = 0
                        If iDistance > -1 Then
                            iDistance = (1 - Abs(iDistance)) * 100
                            SetPixelV picDraw.hDC, nLeft + iX, iY, MixColors(iCol, GetPixel(picDraw.hDC, nLeft + iX, iY), iDistance)
                        End If
                    Next
                End If
                If iCurvPixelsXRightInt > 0 Then
                    For iX = iCurvPixelsXRightInt - 1 To iCurvPixelsXRightInt \ 2 - 1 Step -1
                        If iX < 0 Then Exit For
                        iDistance = nCurveTopRight - (Sqr((nCurveTopRight - iX) ^ 2 + (nCurveTopRight - iY + nTop) ^ 2))
                        If iDistance > 0 Then iDistance = 0
                        If iDistance > -1 Then
                            iDistance = (1 - Abs(iDistance)) * 100
                            SetPixelV picDraw.hDC, nRight - iX - 1, iY, MixColors(iCol, GetPixel(picDraw.hDC, nRight - iX - 1, iY), iDistance)
                        End If
                    Next
                End If
            Else
                picDraw.Line (nLeft + iCurvPixelsXLeftInt, iY)-(nRight - iCurvPixelsXRightInt, iY), iCol
            End If
            sngRed = sngRed + sngRedInc
            sngGreen = sngGreen + sngGreenInc
            sngBlue = sngBlue + sngBlueInc
            
            If sngRed > 255 Then sngRed = 255
            If sngGreen > 255 Then sngGreen = 255
            If sngBlue > 255 Then sngBlue = 255
            If sngRed < 0 Then sngRed = 0
            If sngGreen < 0 Then sngGreen = 0
            If sngBlue < 0 Then sngBlue = 0
        Next
        
        If nCurveBottomLeft > nCurveBottomRight Then
            iCurvPixelsBottomY = nCurveBottomLeft
        Else
            iCurvPixelsBottomY = nCurveBottomRight
        End If
        If iCurvPixelsBottomY > (nBottom - (nTop + iCurvPixelsTopY)) Then
            iCurvPixelsBottomY = (nBottom - (nTop + iCurvPixelsTopY))
        End If
        
        For iY = nTop + iCurvPixelsTopY To nBottom - 1 - iCurvPixelsBottomY
            picDraw.Line (nLeft, iY)-(nRight, iY), RGB(sngRed, sngGreen, sngBlue)
            sngRed = sngRed + sngRedInc
            sngGreen = sngGreen + sngGreenInc
            sngBlue = sngBlue + sngBlueInc
            
            If sngRed > 255 Then sngRed = 255
            If sngGreen > 255 Then sngGreen = 255
            If sngBlue > 255 Then sngBlue = 255
            If sngRed < 0 Then sngRed = 0
            If sngGreen < 0 Then sngGreen = 0
            If sngBlue < 0 Then sngBlue = 0
        Next
        
        For iY = nBottom - iCurvPixelsBottomY To nBottom - 1
            If (nBottom - iY) <= nCurveBottomLeft Then
                iCurvPixelsXLeft = nCurveBottomLeft - Sqr(nCurveBottomLeft ^ 2 - (nCurveBottomLeft + iY - nBottom) ^ 2)
                iCurvPixelsXLeftInt = Round(iCurvPixelsXLeft) + 1
            Else
                iCurvPixelsXLeft = 0
                iCurvPixelsXLeftInt = 0
            End If
            If (nBottom - iY) <= nCurveBottomRight Then
                iCurvPixelsXRight = nCurveBottomRight - Sqr(nCurveBottomRight ^ 2 - (nCurveBottomRight + iY - nBottom) ^ 2)
                iCurvPixelsXRightInt = Round(iCurvPixelsXRight) + 1
            Else
                iCurvPixelsXRight = 0
                iCurvPixelsXRightInt = 0
            End If
            
            iCol = RGB(sngRed, sngGreen, sngBlue)
            If (iCurvPixelsXLeftInt > 0) Or (iCurvPixelsXRightInt > 0) Then
                picDraw.Line (nLeft + iCurvPixelsXLeftInt, iY)-(nRight - iCurvPixelsXRightInt, iY), iCol
                If iCurvPixelsXLeftInt > 0 Then
                    For iX = iCurvPixelsXLeftInt - 1 To iCurvPixelsXLeftInt \ 2 - 1 Step -1
                        If iX < 0 Then Exit For
                        iDistance = nCurveBottomLeft - (Sqr((nCurveBottomLeft - iX) ^ 2 + (nCurveBottomLeft + iY - nBottom) ^ 2))
                        If iDistance > 0 Then iDistance = 0
                        If iDistance > -1 Then
                            iDistance = (1 - Abs(iDistance)) * 100
                            SetPixelV picDraw.hDC, nLeft + iX, iY, MixColors(iCol, GetPixel(picDraw.hDC, nLeft + iX, iY), iDistance)
                        End If
                    Next
                End If
                If iCurvPixelsXRightInt > 0 Then
                    For iX = iCurvPixelsXRightInt - 1 To iCurvPixelsXRightInt \ 2 - 1 Step -1
                        If iX < 0 Then Exit For
                        iDistance = nCurveBottomRight - (Sqr((nCurveBottomRight - iX) ^ 2 + (nCurveBottomRight + iY - nBottom) ^ 2))
                        If iDistance > 0 Then iDistance = 0
                        If iDistance > -1 Then
                            iDistance = (1 - Abs(iDistance)) * 100
                            SetPixelV picDraw.hDC, nRight - iX - 1, iY, MixColors(iCol, GetPixel(picDraw.hDC, nRight - iX - 1, iY), iDistance)
                        End If
                    Next
                End If
            Else
                picDraw.Line (nLeft + iCurvPixelsXLeftInt, iY)-(nRight - iCurvPixelsXRightInt, iY), iCol
            End If
            sngRed = sngRed + sngRedInc
            sngGreen = sngGreen + sngGreenInc
            sngBlue = sngBlue + sngBlueInc
            
            If sngRed > 255 Then sngRed = 255
            If sngGreen > 255 Then sngGreen = 255
            If sngBlue > 255 Then sngBlue = 255
            If sngRed < 0 Then sngRed = 0
            If sngGreen < 0 Then sngGreen = 0
            If sngBlue < 0 Then sngBlue = 0
        Next
    End If
End Sub

Private Sub DrawRoundedCorner(ByVal nCorner As NTCornerPositionConstants, ByVal nX As Long, ByVal nY As Long, ByVal nRoundness As Long, ByVal nColor As Long, Optional ByVal nSkipAtTop As Long)
    Dim iX As Long
    Dim iY As Long
    Dim iDistance As Single
    
    If nCorner = ntCornerTopRight Then
        For iX = nX To nX - nRoundness Step -1
            For iY = nY + nSkipAtTop To nY + nRoundness
                iDistance = Sqr((nRoundness - (iX - nX + nRoundness * 2)) ^ 2 + (nRoundness - (iY - (nY))) ^ 2) - nRoundness
                If (Abs(iDistance) < 1) Then
                    SetPixelV picDraw.hDC, iX, iY, MixColors(nColor, GetPixel(picDraw.hDC, iX, iY), (1 - Abs(iDistance)) * 100)
                End If
            Next
        Next
    ElseIf nCorner = ntCornerTopleft Then
        For iX = nX To nX + nRoundness
            For iY = nY + nSkipAtTop To nY + nRoundness
                iDistance = Sqr((nRoundness - ((nX - iX) + nRoundness * 2)) ^ 2 + (nRoundness - (iY - (nY))) ^ 2) - nRoundness
                If (Abs(iDistance) < 1) Then
                    SetPixelV picDraw.hDC, iX, iY, MixColors(nColor, GetPixel(picDraw.hDC, iX, iY), (1 - Abs(iDistance)) * 100)
                End If
            Next
        Next
    ElseIf nCorner = ntCornerBottomRight Then
        For iX = nX To nX - nRoundness Step -1
            For iY = nY To nY - nRoundness Step -1
                iDistance = Sqr((nX - iX - nRoundness) ^ 2 + (nRoundness - (nY - iY)) ^ 2) - nRoundness
                If (Abs(iDistance) < 1) Then
                    SetPixelV picDraw.hDC, iX, iY, MixColors(nColor, GetPixel(picDraw.hDC, iX, iY), (1 - Abs(iDistance)) * 100)
                End If
            Next
        Next
    ElseIf nCorner = ntCornerBottomLeft Then
        For iX = nX To nX + nRoundness
            For iY = nY To nY - nRoundness Step -1
                iDistance = Sqr((nRoundness - (nX - iX) - nRoundness * 2) ^ 2 + (nRoundness - (nY - iY)) ^ 2) - nRoundness
                If (Abs(iDistance) < 1) Then
                    SetPixelV picDraw.hDC, iX, iY, MixColors(nColor, GetPixel(picDraw.hDC, iX, iY), (1 - Abs(iDistance)) * 100)
                End If
            Next
        Next
    End If
End Sub

Private Sub DrawTriangle(nTriangle() As POINTAPI, ByVal iColor As Long)
    Dim iBrush As Long
    Dim iPrevObj As Long
    
    iBrush = CreateSolidBrush(iColor)
    picDraw.ForeColor = iColor
    iPrevObj = SelectObject(picDraw.hDC, iBrush)
    Polygon picDraw.hDC, nTriangle(0), 3
    iPrevObj = SelectObject(picDraw.hDC, iPrevObj)
    DeleteObject iBrush
End Sub

Private Sub DrawImage(ByVal lDestHDC As Long, ByVal lhBmp As Long, ByVal lTransColor As Long, ByVal nLeft As Long, ByVal nTop As Long, ByVal nWidth As Long, ByVal nHeight As Long, Optional ByVal nXOrigin As Long, Optional ByVal nYOrigin As Long)
    Dim lHDC As Long
    Dim lhBmpOld As Long
    Dim utBmp As BITMAP
    
    lHDC = CreateCompatibleDC(lDestHDC)
    lhBmpOld = SelectObject(lHDC, lhBmp)
    Call GetObjectA(lhBmp, Len(utBmp), utBmp)
    Call TransparentBlt(lDestHDC, nLeft, nTop, nWidth, nHeight, lHDC, nXOrigin, nYOrigin, nWidth, nHeight, lTransColor)
    Call SelectObject(lHDC, lhBmpOld)
    DeleteDC (lHDC)
End Sub

Private Function TranslatedColor(ByVal lOleColor As Long) As Long
    Dim lRGBColor As Long
    
    Call TranslateColor(lOleColor, 0, lRGBColor)
    TranslatedColor = lRGBColor
End Function

'extract Red component from a color
Private Function GetRedValue(ByVal RGBValue As Long) As Integer
    GetRedValue = RGBValue And &HFF
End Function

'extract Green component from a color
Private Function GetGreenValue(ByVal RGBValue As Long) As Integer
    GetGreenValue = ((RGBValue And &HFF00) / &H100) And &HFF
End Function

'extract Blue component from a color
Private Function GetBlueValue(ByVal RGBValue As Long) As Integer
    GetBlueValue = ((RGBValue And &HFF0000) / &H10000) And &HFF
End Function

Private Sub SetColors()
    Dim iBCol As Long
    Dim iCol_L As Integer
    Dim iCol_S As Integer
    Dim iCol_H As Integer
    Dim iBackColorTabs_H As Integer
    Dim iBackColorTabs_L As Integer
    Dim iBackColorTabs_S As Integer
    Dim iBackColorSelectedTab_H As Integer
    Dim iBackColorSelectedTab_L As Integer
    Dim iBackColorSelectedTab_S As Integer
    Dim iAmbientBackColor_H As Integer
    Dim iAmbientBackColor_L As Integer
    Dim iAmbientBackColor_S As Integer
    Dim c As Long
    Dim iCol As Long
    Dim iFlatBarColorInactiveTab As Long
    Dim iFlatSeparationColorTabs As Long
    Dim iFlatSeparationColorBody As Long
    Dim iBackColorSelectedTab As Long
    
    ResetAllPicsDisabled
    mBodyReset = True
    
    If mHandleHighContrastTheme And mHighContrastThemeOn Then
        m3DDKShadow = vb3DDKShadow
        m3DHighlight = vb3DHighlight
        m3DShadow = vb3DShadow
        mGrayText = vbGrayText
        iFlatBarColorInactiveTab = vb3DHighlight
        
        iBCol = TranslatedColor(mBackColorTabs)
        ColorRGBToHLS iBCol, iBackColorTabs_H, iBackColorTabs_L, iBackColorTabs_S
        mBackColorTabsDisabled = ColorHLSToRGB(iBackColorTabs_H, iBackColorTabs_L * 0.98, iBackColorTabs_S * 0.6)
    Else
        iBCol = TranslatedColor(mBackColorTabs)
        ColorRGBToHLS iBCol, iBackColorTabs_H, iBackColorTabs_L, iBackColorTabs_S
        
        iBCol = TranslatedColor(Ambient.BackColor)
        ColorRGBToHLS iBCol, iAmbientBackColor_H, iAmbientBackColor_L, iAmbientBackColor_S
        If mSoftEdges Then
            If (iBackColorTabs_L > 150) Then
                m3DDKShadow = ColorHLSToRGB(iBackColorTabs_H, iBackColorTabs_L * 0.65, iBackColorTabs_S * 0.5)
                m3DShadow = ColorHLSToRGB(iBackColorTabs_H, iBackColorTabs_L * 0.818, iBackColorTabs_S * 0.6)
            Else
                iCol_L = iBackColorTabs_L * 3
                If iCol_L > 240 Then iCol_L = 240
                m3DDKShadow = ColorHLSToRGB(iBackColorTabs_H, iCol_L * 0.65, iBackColorTabs_S * 0.5)
                iCol_L = iBackColorTabs_L * 2
                If iCol_L > 240 Then iCol_L = 240
                m3DShadow = ColorHLSToRGB(iBackColorTabs_H, iCol_L * 0.818, iBackColorTabs_S * 0.6)
            End If
        Else
            If (iBackColorTabs_L > 150) Then
                m3DDKShadow = ColorHLSToRGB(iBackColorTabs_H, iBackColorTabs_L * 0.47, iBackColorTabs_S * 0.18)
                m3DShadow = ColorHLSToRGB(iBackColorTabs_H, iBackColorTabs_L * 0.718, iBackColorTabs_S * 0.3)
            Else
                iCol_L = iBackColorTabs_L * 3
                If iCol_L > 240 Then iCol_L = 240
                m3DDKShadow = ColorHLSToRGB(iBackColorTabs_H, iCol_L, iBackColorTabs_S * 0.18)
                iCol_L = iBackColorTabs_L * 2
                If iCol_L > 240 Then iCol_L = 240
                m3DShadow = ColorHLSToRGB(iBackColorTabs_H, iCol_L * 0.718, iBackColorTabs_S * 0.3)
            End If
        End If
        mBackColorTabsDisabled = ColorHLSToRGB(iBackColorTabs_H, iBackColorTabs_L * 0.98, iBackColorTabs_S * 0.6)
        mGrayText = vbGrayText
        
        If iBackColorTabs_L > 150 Then
            If (iBackColorTabs_L > 200) Then
                iCol_L = iBackColorTabs_L * 1.2
                If iCol_L > 240 Then iCol_L = 240
                m3DHighlight = ColorHLSToRGB(iBackColorTabs_H, iCol_L, iBackColorTabs_S * 0.3)
                iFlatBarColorInactiveTab = ColorHLSToRGB(iBackColorTabs_H, iCol_L * 0.85, iBackColorTabs_S * 0.3)
            Else
                iCol_L = iBackColorTabs_L * 1.1
                If iCol_L > 240 Then iCol_L = 240
                m3DHighlight = ColorHLSToRGB(iBackColorTabs_H, iCol_L, iBackColorTabs_S * 0.2)
                iFlatBarColorInactiveTab = ColorHLSToRGB(iBackColorTabs_H, iCol_L, iBackColorTabs_S * 0.2)
            End If
        Else
            iCol_L = iBackColorTabs_L + (240 - iBackColorTabs_L) * 0.7
            If iCol_L > 240 Then iCol_L = 240
            m3DHighlight = ColorHLSToRGB(iBackColorTabs_H, iCol_L, iBackColorTabs_S * 0.9)
            iCol_L = iBackColorTabs_L + (240 - iBackColorTabs_L) * 0.3
            iFlatBarColorInactiveTab = ColorHLSToRGB(iBackColorTabs_H, iCol_L, iBackColorTabs_S * 0.6)
        End If
    End If
    mBlendDisablePicWithBackColorTabs_NotThemed = (iBackColorTabs_L < 200)
    If mBlendDisablePicWithBackColorTabs_NotThemed Then
        mBackColorTabs_R = iBCol And 255
        mBackColorTabs_G = (iBCol \ 256) And 255
        mBackColorTabs_B = (iBCol \ 65536) And 255
    End If
    
    mGlowColor = GetGlowColor(iBackColorTabs_H, iBackColorTabs_L, iBackColorTabs_S)
    
    mHighlightColor_ColorAutomatic = mGlowColor
    If mHighlightColor_IsAutomatic Then
        mHighlightColor = mGlowColor
    Else
        mGlowColor = mHighlightColor
    End If
    
    mFlatBarColorInactive_ColorAutomatic = iFlatBarColorInactiveTab
    If mFlatBarColorInactive_IsAutomatic Then
        mFlatBarColorInactive = mFlatBarColorInactive_ColorAutomatic
    End If
    
    If Not IsEmpty(mTabData(mTab).CustomColors.FlatBarColorSelectedTab) Then
        iCol = MixColors(mTabData(mTab).CustomColors.FlatBarColorSelectedTab, mFlatBarColorInactive, 60)
    Else
        iCol = MixColors(mFlatBarColorSelectedTab, mFlatBarColorInactive, 60)
    End If
    ColorRGBToHLS iCol, iCol_H, iCol_L, iCol_S
    mFlatBarColorHighlight_ColorAutomatic = ColorHLSToRGB(iCol_H, iCol_L, iCol_S * 0.75)
    If mFlatBarColorHighlight_IsAutomatic Then
        mFlatBarColorHighlight = mFlatBarColorHighlight_ColorAutomatic
    End If
    
    For c = 1 To 10
        mHighlightEffectColors_Strong(c) = MixColors(mGlowColor, mBackColorTabs, 10 * c)
        mHighlightEffectColors_Light(c) = MixColors(mGlowColor, mBackColorTabs, 6 * c)
        mFlatBarHighlightEffectColors(c) = MixColors(mFlatBarColorHighlight, mFlatBarColorInactive, 10 * c)
    Next c
    
    ' Active Tab (TabSel) colors
    If mHandleHighContrastTheme And mHighContrastThemeOn Then
        m3DDKShadow_Sel = vb3DDKShadow
        m3DHighlight_Sel = vb3DHighlight
        m3DShadow_Sel = vb3DShadow
        mGrayText_Sel = vbGrayText
        
        iBCol = TranslatedColor(mBackColorTabs)
        ColorRGBToHLS iBCol, iBackColorSelectedTab_H, iBackColorSelectedTab_L, iBackColorSelectedTab_S
        mBackColorSelectedTabDisabled = ColorHLSToRGB(iBackColorSelectedTab_H, iBackColorSelectedTab_L * 0.98, iBackColorSelectedTab_S * 0.6)
    Else
        If IsEmpty(mTabData(mTab).CustomColors.BackColorSelectedTab) Or (mHandleHighContrastTheme And mHighContrastThemeOn) Then
            iBackColorSelectedTab = mBackColorSelectedTab
        Else
            iBackColorSelectedTab = mTabData(mTab).CustomColors.BackColorSelectedTab
        End If
        iBCol = TranslatedColor(iBackColorSelectedTab)
        ColorRGBToHLS iBCol, iBackColorSelectedTab_H, iBackColorSelectedTab_L, iBackColorSelectedTab_S
        
        iBCol = TranslatedColor(Ambient.BackColor)
        ColorRGBToHLS iBCol, iAmbientBackColor_H, iAmbientBackColor_L, iAmbientBackColor_S
        If mSoftEdges Then
            If (iBackColorSelectedTab_L > 150) Then
                m3DDKShadow_Sel = ColorHLSToRGB(iBackColorSelectedTab_H, iBackColorSelectedTab_L * 0.65, iBackColorSelectedTab_S * 0.5)
                m3DShadow_Sel = ColorHLSToRGB(iBackColorSelectedTab_H, iBackColorSelectedTab_L * 0.818, iBackColorSelectedTab_S * 0.6)
            Else
                iCol_L = iBackColorSelectedTab_L * 3
                If iCol_L > 240 Then iCol_L = 240
                m3DDKShadow_Sel = ColorHLSToRGB(iBackColorSelectedTab_H, iCol_L * 0.65, iBackColorSelectedTab_S * 0.5)
                iCol_L = iBackColorSelectedTab_L * 2
                If iCol_L > 240 Then iCol_L = 240
                m3DShadow_Sel = ColorHLSToRGB(iBackColorSelectedTab_H, iCol_L * 0.818, iBackColorSelectedTab_S * 0.6)
            End If
        Else
            If (iBackColorSelectedTab_L > 150) Then
                m3DDKShadow_Sel = ColorHLSToRGB(iBackColorSelectedTab_H, iBackColorSelectedTab_L * 0.47, iBackColorSelectedTab_S * 0.18)
                m3DShadow_Sel = ColorHLSToRGB(iBackColorSelectedTab_H, iBackColorSelectedTab_L * 0.718, iBackColorSelectedTab_S * 0.3)
            Else
                iCol_L = iBackColorSelectedTab_L * 3
                If iCol_L > 240 Then iCol_L = 240
                m3DDKShadow_Sel = ColorHLSToRGB(iBackColorSelectedTab_H, iCol_L, iBackColorSelectedTab_S * 0.18)
                iCol_L = iBackColorSelectedTab_L * 2
                If iCol_L > 240 Then iCol_L = 240
                m3DShadow_Sel = ColorHLSToRGB(iBackColorSelectedTab_H, iCol_L * 0.718, iBackColorSelectedTab_S * 0.3)
            End If
        End If
        mGrayText_Sel = vbGrayText
        
        If iBackColorSelectedTab_L > 150 Then
            If (iBackColorSelectedTab_L > 200) And (iBackColorSelectedTab_S < 150) Then
                iCol_L = iBackColorSelectedTab_L * 1.2
                If iCol_L > 240 Then iCol_L = 240
                m3DHighlight_Sel = ColorHLSToRGB(iBackColorSelectedTab_H, iCol_L, iBackColorSelectedTab_S * 0.3)
            Else
                iCol_L = iBackColorSelectedTab_L * 1.1
                If iCol_L > 240 Then iCol_L = 240
                m3DHighlight_Sel = ColorHLSToRGB(iBackColorSelectedTab_H, iCol_L, iBackColorSelectedTab_S * 0.2)
            End If
        Else
            iCol_L = iBackColorSelectedTab_L + (240 - iBackColorSelectedTab_L) * 0.7
            If iCol_L > 240 Then iCol_L = 240
            m3DHighlight_Sel = ColorHLSToRGB(iBackColorSelectedTab_H, iCol_L, iBackColorSelectedTab_S * 0.9)
        End If
    End If
    If mBlendDisablePicWithBackColorTabs_NotThemed Then
        mBackColorSelectedTab_R = iBCol And 255
        mBackColorSelectedTab_G = (iBCol \ 256) And 255
        mBackColorSelectedTab_B = (iBCol \ 65536) And 255
    End If
    
    If iBackColorSelectedTab_L > 150 Then
        If (iBackColorSelectedTab_L > 233) Then
            If iBackColorSelectedTab_S = 0 Then
                iCol_L = iBackColorSelectedTab_L * 0.95
                iCol_S = 0
                iCol_H = iBackColorSelectedTab_H
            Else
                iCol_L = iBackColorSelectedTab_L * 0.9
                iCol_S = 80
                iCol_H = iBackColorSelectedTab_H
            End If
            mGlowColor_Sel = ColorHLSToRGB(iCol_H, iCol_L, iCol_S)
        ElseIf (iBackColorSelectedTab_L > 200) And (iBackColorSelectedTab_S < 150) Then
            iCol_L = iBackColorSelectedTab_L + (240 - iBackColorSelectedTab_L) * 0.1 + iBackColorSelectedTab_L * 0.05 + 5
            If iCol_L > 240 Then iCol_L = 240
            mGlowColor_Sel = ColorHLSToRGB(iBackColorSelectedTab_H, iCol_L, iBackColorSelectedTab_S)
        Else
            iCol_S = iBackColorSelectedTab_S
            If iBackColorSelectedTab_L > 160 Then
                iCol_L = iBackColorSelectedTab_L * 1.15
            Else
                iCol_L = iBackColorSelectedTab_L + (240 - iBackColorSelectedTab_L) * 0.2 + iBackColorSelectedTab_L * 0.06 + 5
            End If
            If iCol_L > 240 Then iCol_L = 240
            If iCol_L > 200 Then
                If iBackColorSelectedTab_L > 210 Then
                    iCol_S = 1
                Else
                    If iCol_S > 100 Then
                        If ((iBackColorSelectedTab_H > 35) And (iBackColorSelectedTab_H < 45)) Then
                            iCol_S = iCol_S - 100
                            If iCol_S < 1 Then iCol_S = 1
                            iCol_L = iCol_L + 20
                            If iCol_L > 240 Then iCol_L = 240
                        End If
                    End If
                End If
            End If
            mGlowColor_Sel = ColorHLSToRGB(iBackColorSelectedTab_H, iCol_L, iCol_S)
        End If
    Else
        If iBackColorSelectedTab_S > 60 Then
            Select Case iBackColorSelectedTab_H
                Case 0 To 30, 220 To 240 ' reds
                    If iBackColorSelectedTab_L < 100 Then
                        iCol_L = iBackColorSelectedTab_L + (240 - iBackColorSelectedTab_L) * 0.07
                    Else
                        iCol_L = iBackColorSelectedTab_L
                    End If
                Case 200 To 219 ' violet
                    iCol_L = iBackColorSelectedTab_L + (240 - iBackColorSelectedTab_L) * 0.3
                Case 31 To 120 ' yellows, greenes, cyanes
                    iCol_L = iBackColorSelectedTab_L + (240 - iBackColorSelectedTab_L) * 0.2
                Case Else ' blues
                    If iBackColorSelectedTab_L < 100 Then
                        iCol_L = iBackColorSelectedTab_L + (240 - iBackColorSelectedTab_L) * 0.15
                    Else
                        iCol_L = iBackColorSelectedTab_L '+ (240 - iBackColorSelectedTab_L) * 0.07
                    End If
            End Select
        Else ' gray
            iCol_L = iBackColorSelectedTab_L + (240 - iBackColorSelectedTab_L) * 0.2
        End If
        iCol_L = iCol_L + 15
        If iCol_L > 240 Then iCol_L = 240
        iCol_S = iBackColorSelectedTab_S
        If iCol_S > 200 Then
            iCol_S = iCol_S * 0.65
            If iCol_S < 1 Then iCol_S = 1
            iCol_L = iCol_L * 1.4
            If iCol_L > 240 Then iCol_L = 240
        Else
            iCol_S = iCol_S * 1.1
        End If
        If iCol_S > 240 Then iCol_S = 240

        mGlowColor_Sel = ColorHLSToRGB(iBackColorSelectedTab_H, iCol_L, iCol_S)
    End If
    
    mHighlightColorSelectedTab_ColorAutomatic = mGlowColor_Sel
    If mHighlightColorSelectedTab_IsAutomatic Then
        mHighlightColorSelectedTab = mGlowColor_Sel
    Else
        mGlowColor_Sel = mHighlightColorSelectedTab
    End If
    If Not IsEmpty(mTabData(mTab).CustomColors.HighlightColorSelectedTab) Then
        mGlowColor_Sel = mTabData(mTab).CustomColors.HighlightColorSelectedTab
    End If
    mGlowColor_Sel_Bk = mGlowColor_Sel
    mGlowColor_Sel_Light = MixColors(mGlowColor_Sel, iBackColorSelectedTab, 60)
    
    If iBackColorTabs_L > 150 Then
        iFlatSeparationColorTabs = MixColors(m3DDKShadow, TranslatedColor(mBackColorTabs), 12)
        iFlatSeparationColorBody = MixColors(m3DDKShadow, TranslatedColor(mBackColorTabs), 17)
    Else
        iFlatSeparationColorTabs = MixColors(m3DDKShadow, TranslatedColor(mBackColorTabs), 15)
        iFlatSeparationColorBody = MixColors(m3DDKShadow, TranslatedColor(mBackColorTabs), 15)
    End If
    SetHighlightMode
    
    mFlatTabsSeparationLineColor_ColorAutomatic = iFlatSeparationColorTabs
    If mFlatTabsSeparationLineColor_IsAutomatic Then
        mFlatTabsSeparationLineColor = mFlatTabsSeparationLineColor_ColorAutomatic
    End If
    mFlatBodySeparationLineColor_ColorAutomatic = iFlatSeparationColorBody
    If mFlatBodySeparationLineColor_IsAutomatic Then
        mFlatBodySeparationLineColor = mFlatBodySeparationLineColor_ColorAutomatic
    End If
    mFlatBorderColor_ColorAutomatic = m3DShadow_Sel
    If mFlatBorderColor_IsAutomatic Then
        mFlatBorderColor = mFlatBorderColor_ColorAutomatic
    End If
    
End Sub

Private Function GetGlowColor(nBackColorTabs_H As Integer, nBackColorTabs_L As Integer, nBackColorTabs_S As Integer) As Long
    Dim iBCol As Long
    Dim iCol_L As Integer
    Dim iCol_S As Integer
    Dim iCol_H As Integer
    
    If nBackColorTabs_L > 150 Then
        If (nBackColorTabs_L > 233) Then
            If nBackColorTabs_S = 0 Then
                iCol_L = nBackColorTabs_L * 0.95
                iCol_S = 0
                iCol_H = nBackColorTabs_H
            Else
                iCol_L = nBackColorTabs_L * 0.9
                iCol_S = 80
                iCol_H = nBackColorTabs_H
            End If
            GetGlowColor = ColorHLSToRGB(iCol_H, iCol_L, iCol_S)
        ElseIf (nBackColorTabs_L > 200) And (nBackColorTabs_S < 150) Then
            iCol_L = nBackColorTabs_L + (240 - nBackColorTabs_L) * 0.1 + nBackColorTabs_L * 0.05 + 5
            If iCol_L > 240 Then iCol_L = 240
            GetGlowColor = ColorHLSToRGB(nBackColorTabs_H, iCol_L, nBackColorTabs_S)
        Else
            iCol_S = nBackColorTabs_S
            If nBackColorTabs_L > 160 Then
                iCol_L = nBackColorTabs_L * 1.15
            Else
                iCol_L = nBackColorTabs_L + (240 - nBackColorTabs_L) * 0.2 + nBackColorTabs_L * 0.06 + 5
            End If
            If iCol_L > 240 Then iCol_L = 240
            If iCol_L > 200 Then
                If nBackColorTabs_L > 210 Then
                    iCol_S = 1
                Else
                    If iCol_S > 100 Then
                        If ((nBackColorTabs_H > 35) And (nBackColorTabs_H < 45)) Then
                            iCol_S = iCol_S - 100
                            If iCol_S < 1 Then iCol_S = 1
                            iCol_L = iCol_L + 20
                            If iCol_L > 240 Then iCol_L = 240
                        End If
                    End If
                End If
            End If
            GetGlowColor = ColorHLSToRGB(nBackColorTabs_H, iCol_L, iCol_S)
        End If
    Else
        If nBackColorTabs_S > 60 Then
            Select Case nBackColorTabs_H
                Case 0 To 30, 220 To 240 ' reds
                    If nBackColorTabs_L < 100 Then
                        iCol_L = nBackColorTabs_L + (240 - nBackColorTabs_L) * 0.07
                    Else
                        iCol_L = nBackColorTabs_L
                    End If
                Case 200 To 219 ' violet
                    iCol_L = nBackColorTabs_L + (240 - nBackColorTabs_L) * 0.3
                Case 31 To 120 ' yellows, greenes, cyanes
                    iCol_L = nBackColorTabs_L + (240 - nBackColorTabs_L) * 0.2
                Case Else ' blues
                    If nBackColorTabs_L < 100 Then
                        iCol_L = nBackColorTabs_L + (240 - nBackColorTabs_L) * 0.15
                    Else
                        iCol_L = nBackColorTabs_L '+ (240 - nBackColorTabs_L) * 0.07
                    End If
            End Select
        Else ' gray
            iCol_L = nBackColorTabs_L + (240 - nBackColorTabs_L) * 0.2
        End If
        iCol_L = iCol_L + 15
        If iCol_L > 240 Then iCol_L = 240
        iCol_S = nBackColorTabs_S
        If iCol_S > 200 Then
            iCol_S = iCol_S * 0.65
            If iCol_S < 1 Then iCol_S = 1
            iCol_L = iCol_L * 1.4
            If iCol_L > 240 Then iCol_L = 240
        Else
            iCol_S = iCol_S * 1.1
        End If
        If iCol_S > 240 Then iCol_S = 240
        
        GetGlowColor = ColorHLSToRGB(nBackColorTabs_H, iCol_L, iCol_S)
    End If
End Function

Private Function MixColors(ByVal nColor1 As Long, ByVal nColor2 As Long, ByVal nPercentageColor1 As Long) As Long
    Dim iColor1 As Long
    Dim iColor2 As Long
    Dim iColor1_R  As Byte
    Dim iColor1_G   As Byte
    Dim iColor1_B   As Byte
    Dim iColor2_R  As Byte
    Dim iColor2_G   As Byte
    Dim iColor2_B   As Byte
    
    iColor1 = TranslatedColor(nColor1)
    iColor2 = TranslatedColor(nColor2)
    
    iColor1_R = iColor1 And 255
    iColor1_G = (iColor1 \ 256) And 255
    iColor1_B = (iColor1 \ 65536) And 255
    iColor2_R = iColor2 And 255
    iColor2_G = (iColor2 \ 256) And 255
    iColor2_B = (iColor2 \ 65536) And 255
    
    If nPercentageColor1 > 100 Then nPercentageColor1 = 100
    If nPercentageColor1 < 0 Then nPercentageColor1 = 0
    
    MixColors = RGB((iColor1_R * nPercentageColor1 + iColor2_R * (100 - nPercentageColor1)) / 100, (iColor1_G * nPercentageColor1 + iColor2_G * (100 - nPercentageColor1)) / 100, (iColor1_B * nPercentageColor1 + iColor2_B * (100 - nPercentageColor1)) / 100)
    
End Function


Public Property Get ClientLeft() As Single
Attribute ClientLeft.VB_Description = "Returns the left of the tab body."
Attribute ClientLeft.VB_ProcData.VB_Invoke_Property = ";Position"
    EnsureDrawn
    ClientLeft = FixRoundingError(UserControl.ScaleX(mClientRect.Left, vbPixels, vbTwips))
End Property

Public Property Get ClientTop() As Single
Attribute ClientTop.VB_Description = "Returns the space occupied by tabs, or where the tab body begins."
Attribute ClientTop.VB_ProcData.VB_Invoke_Property = ";Position"
    EnsureDrawn
    ClientTop = FixRoundingError(UserControl.ScaleY(mClientRect.Top, vbPixels, vbTwips))
End Property

Public Property Get ClientWidth() As Single
Attribute ClientWidth.VB_Description = "Returns the width of the tab body."
Attribute ClientWidth.VB_ProcData.VB_Invoke_Property = ";Position"
    EnsureDrawn
    ClientWidth = FixRoundingError(UserControl.ScaleX(mClientRect.Right - mClientRect.Left, vbPixels, vbTwips))
End Property

Public Property Get ClientHeight() As Single
Attribute ClientHeight.VB_Description = "Returns the height of the tab body."
Attribute ClientHeight.VB_ProcData.VB_Invoke_Property = ";Position"
    EnsureDrawn
    ClientHeight = FixRoundingError(UserControl.ScaleY(mClientRect.Bottom - mClientRect.Top, vbPixels, vbTwips))
End Property

Private Sub EnsureDrawn()
    Dim c  As Long
    
    If (Not mFirstDraw) Or tmrDraw.Enabled Or mDrawMessagePosted Then
        mEnsuringDrawn = True
        Draw
        Do Until Not (mDrawMessagePosted Or tmrDraw.Enabled)
            Draw
            c = c + 1
            If c > 5 Then Exit Do
        Loop
        mEnsuringDrawn = False
    End If
End Sub

Private Sub RotatePic(picSrc As PictureBox, picDest As PictureBox, ByVal nDirection As NTRotatePicDirectionConstants)
    Dim pt(0 To 2) As POINTAPI
    Dim iHeight As Long
    Dim iWidth As Long
    
    iHeight = picSrc.Height
    iWidth = picSrc.Width
    
    'Set the coordinates of the parallelogram
    If nDirection = nt90DegreesClockWise Then ' 90 degrees
        pt(0).X = iHeight
        pt(0).Y = 0
        pt(1).X = iHeight
        pt(1).Y = iWidth
        pt(2).X = 0
        pt(2).Y = 0
    ElseIf nDirection = nt90DegreesCounterClockWise Then ' 270 degrees
        pt(0).X = 0
        pt(0).Y = iWidth
        pt(1).X = 0
        pt(1).Y = 0
        pt(2).X = iHeight
        pt(2).Y = iWidth
    ElseIf nDirection = ntFlipVertical Then ' vertical
        pt(0).X = 0
        pt(0).Y = iHeight
        pt(1).X = iWidth
        pt(1).Y = iHeight
        pt(2).X = 0
        pt(2).Y = 0
    ElseIf nDirection = ntFlipHorizontal Then ' horizontal
        pt(0).X = iWidth
        pt(0).Y = 0
        pt(1).X = 0
        pt(1).Y = 0
        pt(2).X = iWidth
        pt(2).Y = iHeight
    End If
    
    picDest.Width = picSrc.Height
    picDest.Height = picSrc.Width
    picDest.Cls
    
    picDest.Cls
    PlgBlt picDest.hDC, pt(0), picSrc.hDC, 0, 0, iWidth, iHeight, ByVal 0&, ByVal 0&, ByVal 0&
End Sub

Private Function ContainerScaleMode() As ScaleModeConstants
    ContainerScaleMode = vbTwips
    On Error Resume Next
    ContainerScaleMode = UserControl.Extender.Container.ScaleMode
    Err.Clear
End Function

Friend Function FromContainerSizeY(nValue As Variant, Optional nToScale As ScaleModeConstants = vbTwips) As Single
    FromContainerSizeY = pScaleY(nValue, ContainerScaleMode, nToScale)
End Function

Private Function ToContainerSizeY(nValue As Variant, Optional nFromScale As ScaleModeConstants = vbTwips) As Single
    ToContainerSizeY = pScaleY(nValue, nFromScale, ContainerScaleMode)
End Function


Friend Function FromContainerSizeX(nValue As Variant, Optional nToScale As ScaleModeConstants = vbTwips) As Single
    FromContainerSizeX = pScaleX(nValue, ContainerScaleMode, nToScale)
End Function

Private Function ToContainerSizeX(nValue As Variant, Optional nFromScale As ScaleModeConstants = vbTwips) As Single
    ToContainerSizeX = pScaleX(nValue, nFromScale, ContainerScaleMode)
End Function

Private Function FixRoundingError(nNumber As Single, Optional ByVal nDecimals As Long = 3) As Single
    Dim iNum As Single
    
    iNum = Round(nNumber * 10 ^ nDecimals) / 10 ^ nDecimals
    If iNum = Int(iNum) Then
        FixRoundingError = iNum
    Else
        If (ContainerScaleMode = vbTwips) Or (ContainerScaleMode = vbPixels) Then
            FixRoundingError = Round(nNumber)
        Else
            FixRoundingError = nNumber
        End If
    End If
End Function
    
Private Sub SetControlsBackColor(ByVal nColor As Long, Optional ByVal nPrevColor As Long = -1)
    Dim iCtl As Object
    Dim iLng As Long
    Dim iCancel As Boolean
    Dim iControls As Object
    Dim iContainer As Object
    Dim iContainer_Prev As Object
    Dim iStr As String
    Dim iCtlIsNewTab As Boolean
    Dim iLngT As Long
    Dim iLngTS As Long
    
    On Error Resume Next
    Set iControls = UserControl.Parent.Controls
    If nPrevColor = -1 Then nPrevColor = mControlsBackColor_PrevColor
    
    If iControls Is Nothing Then ' at least let's change the backcolor of the contained controls in the usercontrol
        For Each iCtl In UserControlContainedControls
            iLng = -1
            iLng = iCtl.BackColor
            If iLng <> -1 Then
                If (iLng = vbButtonFace) And (nColor <> vbButtonFace) Or (iLng = nPrevColor) Then
                    iCancel = False
                    iStr = iCtl.Name
                    RaiseEvent ChangeControlBackColor(iStr, TypeName(iCtl), iCancel)
                    If Not iCancel Then
                        iCtl.BackColor = nColor
                    End If
                End If
            End If
        Next
    Else 'let's change the backcolor of all the controls inside the tabs
        For Each iCtl In iControls
            Set iContainer = Nothing
            Set iContainer = iCtl.Container
            Do Until iContainer Is Nothing
                If iContainer Is UserControl.Extender Then
                    iLng = -1
                    iLng = iCtl.BackColor
                    If TypeName(iCtl) = TypeName(Me) Then
                        iLngT = iCtl.BackColorTabs
                        iLngTS = iCtl.BackColorSelectedTab
                        iCtlIsNewTab = True
                    End If
                    If iLng <> -1 Then
                        If (iLng = vbButtonFace) And (nColor <> vbButtonFace) Or (iLng = nPrevColor) Then
                            iCancel = False
                            If Not iContainer_Prev Is Nothing Then
                                If iContainer_Prev.Container Is UserControl.Extender Then
                                    iStr = iContainer_Prev.Name
                                    RaiseEvent ChangeControlBackColor(iStr, TypeName(iContainer_Prev), iCancel)
                                End If
                            End If
                            If Not iCancel Then
                                iCancel = False
                                iStr = iCtl.Name
                                RaiseEvent ChangeControlBackColor(iStr, TypeName(iCtl), iCancel)
                                If Not iCancel Then
                                    iCtl.BackColor = nColor
                                End If
                            End If
                        End If
                    End If
                    If iCtlIsNewTab Then
                        If (iLngT = vbButtonFace) And (nColor <> vbButtonFace) Or (iLngT = nPrevColor) Then
                            iCancel = False
                            If Not iContainer_Prev Is Nothing Then
                                If iContainer_Prev.Container Is UserControl.Extender Then
                                    iStr = iContainer_Prev.Name
                                    RaiseEvent ChangeControlBackColor(iStr, TypeName(iContainer_Prev), iCancel)
                                End If
                            End If
                            If Not iCancel Then
                                iCancel = False
                                iStr = iCtl.Name
                                RaiseEvent ChangeControlBackColor(iStr, TypeName(iCtl), iCancel)
                                If Not iCancel Then
                                    iCtl.BackColorTabs = nColor
                                End If
                            End If
                        End If
                        If iLngTS <> iLngT Then
                            If (iLngTS = vbButtonFace) And (nColor <> vbButtonFace) Or (iLngTS = nPrevColor) Then
                                iCancel = False
                                If Not iContainer_Prev Is Nothing Then
                                    If iContainer_Prev.Container Is UserControl.Extender Then
                                        iStr = iContainer_Prev.Name
                                        RaiseEvent ChangeControlBackColor(iStr, TypeName(iContainer_Prev), iCancel)
                                    End If
                                End If
                                If Not iCancel Then
                                    iCancel = False
                                    iStr = iCtl.Name
                                    RaiseEvent ChangeControlBackColor(iStr, TypeName(iCtl), iCancel)
                                    If Not iCancel Then
                                        iCtl.BackColorSelectedTab = nColor
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                Set iContainer_Prev = iContainer
                Set iContainer = Nothing
                Set iContainer = iContainer_Prev.Container
            Loop
        Next
    End If
    Err.Clear
    mControlsBackColor_PrevColor = nColor
End Sub

Private Sub SetControlsForeColor(ByVal nColor As Long, Optional ByVal nPrevColor As Long = -1)
    If nColor <> mControlsForeColor_PrevColor Then
        SetControlsForeColor2 nColor, nPrevColor
        mControlsForeColor_PrevColor = nColor
    End If
End Sub

Private Sub SetControlsForeColor2(ByVal nColor As Long, Optional ByVal nPrevColor As Long = -1)
    Dim iCtl As Object
    Dim iLng As Long
    Dim iCancel As Boolean
    Dim iControls As Object
    Dim iContainer As Object
    Dim iContainer_Prev As Object
    Dim iStr As String
    Dim iCtlIsNewTab As Boolean
    
    On Error Resume Next
    Set iControls = UserControl.Parent.Controls
    If nPrevColor = -1 Then nPrevColor = mControlsForeColor_PrevColor
    
    If iControls Is Nothing Then ' at least let's change the Forecolor of the contained controls in the usercontrol
        For Each iCtl In UserControlContainedControls
            iLng = -1
            iLng = iCtl.ForeColor
            If iLng <> -1 Then
                If (iLng = vbButtonText) And (nColor <> vbButtonText) Or (iLng = nPrevColor) Then
                    iCancel = False
                    iStr = iCtl.Name
                    RaiseEvent ChangeControlForeColor(iStr, TypeName(iCtl), iCancel)
                    If Not iCancel Then
                        iCtl.ForeColor = nColor
                    End If
                End If
            End If
        Next
    Else 'let's change the Forecolor of all the controls inside the tabs
        For Each iCtl In iControls
            Set iContainer = Nothing
            Set iContainer = iCtl.Container
            Do Until iContainer Is Nothing
                If iContainer Is UserControl.Extender Then
                    iLng = -1
                    If TypeName(iCtl) = TypeName(Me) Then
                        iLng = iCtl.ForeColorSelectedTab
                        iCtlIsNewTab = True
                    Else
                        iLng = iCtl.ForeColor
                    End If
                    If iLng <> -1 Then
                        If (iLng = vbButtonText) And (nColor <> vbButtonText) Or (iLng = nPrevColor) Then
                            iCancel = False
                            If Not iContainer_Prev Is Nothing Then
                                If iContainer_Prev.Container Is UserControl.Extender Then
                                    iStr = iContainer_Prev.Name
                                    RaiseEvent ChangeControlForeColor(iStr, TypeName(iContainer_Prev), iCancel)
                                End If
                            End If
                            If Not iCancel Then
                                iCancel = False
                                iStr = iCtl.Name
                                RaiseEvent ChangeControlForeColor(iStr, TypeName(iCtl), iCancel)
                                If Not iCancel Then
                                    If iCtlIsNewTab Then
                                        iCtl.ForeColorSelectedTab = nColor
                                    Else
                                        iCtl.ForeColor = nColor
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                Set iContainer_Prev = iContainer
                Set iContainer = Nothing
                Set iContainer = iContainer_Prev.Container
            Loop
        Next
    End If
    Err.Clear
End Sub

Public Sub Refresh()
Attribute Refresh.VB_Description = "Redraws the control."
    Dim iWv As Boolean
    
    iWv = IsWindowVisible(mUserControlHwnd) <> 0
    If iWv Then SendMessage mUserControlHwnd, WM_SETREDRAW, False, 0&
    mBodyReset = True
    If mChangeControlsBackColor Then
        SetControlsBackColor mBackColorSelectedTab
    End If
    If mChangeControlsForeColor Then
        SetControlsProperForeColor
    End If
    StoreControlsTabStop
    mRedraw = True
    mSubclassControlsPaintingPending = True
    mRepaintSubclassedControls = True
    Draw
    If iWv Then SendMessage mUserControlHwnd, WM_SETREDRAW, True, 0&
    If iWv Then RedrawWindow mUserControlHwnd, ByVal 0&, 0&, RDW_INVALIDATE Or RDW_ALLCHILDREN
    If iWv Then UserControl.Refresh
End Sub

Private Sub RaiseEvent_TabMouseEnter(ByVal nTab As Long)
    If DraggingATab Then Exit Sub
    mTabData(nTab).Hovered = True
    RaiseEvent TabMouseEnter(nTab)
    mCurrentMousePointerIsHand = False
    If mTabMousePointerHand Then
        If nTab <> mTab Then
            mCurrentMousePointerIsHand = True
            If GetCursor <> IDC_HAND Then
                SetCursor mHandIconHandle
            End If
        End If
    End If
    If mHighlightEffect And (Not mControlIsThemed) Then
        tmrHighlightEffect.Enabled = False
        tmrHighlightEffect.Enabled = True
        mHighlightEffect_Step = 1
        If mHighlightIntensity = ntHighlightIntensityStrong Then
            mGlowColor = mHighlightEffectColors_Strong(mHighlightEffect_Step)
        Else
            mGlowColor = mHighlightEffectColors_Light(mHighlightEffect_Step)
        End If
        If mTabUnderMouse > -1 Then
            If Not IsEmpty(mTabData(mTabUnderMouse).CustomColors.HighlightColor) Then
                If mHighlightIntensity = ntHighlightIntensityStrong Then
                    mTabData(mTabUnderMouse).CustomColors.GlowColor = mTabData(mTabUnderMouse).CustomColors.HighlightEffectColors_Strong(mHighlightEffect_Step)
                Else
                    mTabData(mTabUnderMouse).CustomColors.GlowColor = mTabData(mTabUnderMouse).CustomColors.HighlightEffectColors_Light(mHighlightEffect_Step)
                End If
            End If
        End If
        mFlatBarGlowColor = mFlatBarHighlightEffectColors(mHighlightEffect_Step)
        If (Not IsEmpty(mTabData(mTabUnderMouse).CustomColors.FlatBarColorInactive)) Or (Not IsEmpty(mTabData(mTabUnderMouse).CustomColors.FlatBarColorHighlight)) Then
            mTabData(mTabUnderMouse).CustomColors.FlatBarGlowColor = mTabData(mTabUnderMouse).CustomColors.FlatBarHighlightEffectColors(mHighlightEffect_Step)
        End If
    ElseIf (Not mControlIsThemed) Then
        mGlowColor = mHighlightEffectColors_Strong(10)
        If mTabUnderMouse > -1 Then
            If Not IsEmpty(mTabData(mTabUnderMouse).CustomColors.HighlightColor) Then
                mTabData(mTabUnderMouse).CustomColors.GlowColor = mTabData(mTabUnderMouse).CustomColors.HighlightEffectColors_Strong(10)
            End If
        End If
        mFlatBarGlowColor = mFlatBarHighlightEffectColors(10)
        If (Not IsEmpty(mTabData(mTabUnderMouse).CustomColors.FlatBarColorInactive)) Or (Not IsEmpty(mTabData(mTabUnderMouse).CustomColors.FlatBarColorHighlight)) Then
            mTabData(mTabUnderMouse).CustomColors.FlatBarGlowColor = mTabData(mTabUnderMouse).CustomColors.FlatBarHighlightEffectColors(10)
        End If
        Draw
    End If
    If (mHighlightGradient <> ntGradientNone) Or mAppearanceIsFlat Or mControlIsThemed Then PostDrawMessage
    
    If mThereAreTabsToolTipTexts Then
        If mTabData(nTab).ToolTipText <> "" Then
            ShowTabTTT nTab
        End If
    End If
End Sub

Private Sub RaiseEvent_TabMouseLeave(ByVal nTab As Long)
    mCurrentMousePointerIsHand = False
    If tmrHighlightEffect.Enabled Then
        tmrHighlightEffect.Enabled = False
        If mHighlightIntensity = ntHighlightIntensityStrong Then
            mGlowColor = mHighlightEffectColors_Strong(10)
        Else
            mGlowColor = mHighlightEffectColors_Light(10)
        End If
    End If
    If mTabUnderMouse > -1 Then
        If Not IsEmpty(mTabData(mTabUnderMouse).CustomColors.HighlightColor) Then
            If mHighlightIntensity = ntHighlightIntensityStrong Then
                mTabData(mTabUnderMouse).CustomColors.GlowColor = mTabData(mTabUnderMouse).CustomColors.HighlightEffectColors_Strong(10)
            Else
                mTabData(mTabUnderMouse).CustomColors.GlowColor = mTabData(mTabUnderMouse).CustomColors.HighlightEffectColors_Light(10)
            End If
        End If
    End If
    mTabData(nTab).Hovered = False
    RaiseEvent TabMouseLeave(nTab)
    If nTab <> mTab Then
        If (mHighlightGradient <> ntGradientNone) Or mAppearanceIsFlat Or mControlIsThemed Then PostDrawMessage
    End If
    mMouseIsOverIcon = False
    mMouseIsOverIcon_Tab = -1
    DrawDelayed
    tmrShowTabTTT.Enabled = False
    Set mToolTipEx = Nothing
End Sub

Private Sub ShowTabTTT(ByVal nTab As Long)
    tmrShowTabTTT.Enabled = False
    tmrShowTabTTT.Enabled = True
    tmrShowTabTTT.Tag = nTab
End Sub

Private Sub CheckIfThereAreTabsToolTipTexts()
    Dim c As Long
    
    'If Not mAmbientUserMode Then Exit Sub
    mThereAreTabsToolTipTexts = False
    For c = 0 To mTabs - 1
        If mTabData(c).ToolTipText <> "" Then
            mThereAreTabsToolTipTexts = True
            Exit Sub
        End If
    Next c
End Sub

Private Sub SetButtonFaceColor()
    Dim iCol As Long
    
    iCol = TranslatedColor(vbButtonFace)
    ColorRGBToHLS iCol, mButtonFace_H, mButtonFace_L, mButtonFace_S
    
End Sub

Private Sub SetThemedTabTransparentPixels(nIsLeftTab As Boolean, nIsRightTab As Boolean, nIsTopTab As Boolean)
    Dim X As Long
    Dim X2 As Long
    Dim iYLenght As Long
    
    If nIsLeftTab Or nIsTopTab Then
        For X = 0 To 5
            iYLenght = mTABITEM_TopLeftCornerTransparencyMask(X)
            If iYLenght < 0 Then
                iYLenght = picAux.ScaleHeight - iYLenght
            End If
            If iYLenght > 0 Then
                picAux.Line (X, 0)-(X, iYLenght), cAuxTransparentColor
            End If
        Next X
    End If
    If nIsRightTab Then
        For X = 0 To 5
            X2 = picAux.ScaleWidth - 1 - X
            iYLenght = mTABITEMRIGHTEDGE_RightSideTransparencyMask(X)
            If iYLenght < 0 Then
                iYLenght = picAux.ScaleHeight - iYLenght
            End If
            If iYLenght > 0 Then
                picAux.Line (X2, 0)-(X2, iYLenght), cAuxTransparentColor
            End If
        Next X
    ElseIf nIsTopTab Then
        For X = 0 To 5
            X2 = picAux.ScaleWidth - 1 - X
            iYLenght = mTABITEM_TopRightCornerTransparencyMask(X)
            If iYLenght < 0 Then
                iYLenght = picAux.ScaleHeight - iYLenght
            End If
            If iYLenght > 0 Then
                picAux.Line (X2, 0)-(X2, iYLenght), cAuxTransparentColor
            End If
        Next X
    End If
    
End Sub

Private Sub EnsureBodyThemedReady()
    If Not mBodyThemedReady Then
        Dim iRect As RECT
        
        iRect.Left = 0
        iRect.Top = 0
        iRect.Right = mClientWidth  '+ 1 '- 1
        iRect.Bottom = mClientHeight '- 1 '+ 1 '- 1
        picBodyThemed.Width = iRect.Right
        If iRect.Bottom > 0 Then
            picBodyThemed.Height = iRect.Bottom
        End If
        picBodyThemed.BackColor = mBackColor
        picBodyThemed.Cls
        If (mTabOrientation = ssTabOrientationTop) Then
            DrawThemeBackground mTheme, picBodyThemed.hDC, TABP_PANE, 0&, iRect, iRect
        ElseIf (mTabOrientation = ssTabOrientationLeft) Or (mTabOrientation = ntTabOrientationLeftHorizontal) Then
            ' shadow must be at the bottom, and since the image will be rotated it must be at the left here.
            picAux.Cls
            picAux.Width = picBodyThemed.Width
            picAux.Height = picBodyThemed.Height
            DrawThemeBackground mTheme, picAux.hDC, TABP_PANE, 0&, iRect, iRect
            picBodyThemed.PaintPicture picAux.Image, picAux.ScaleWidth - 1, 0, -picAux.ScaleWidth, picAux.ScaleHeight
        Else ' (mTabOrientation = ssTabOrientationBottom) Or (mTabOrientation = ssTabOrientationRight)  Or (mTabOrientation = ntTabOrientationRightHorizontal)
            picAux.Cls
            picAux.Width = picBodyThemed.Width
            picAux.Height = picBodyThemed.Height
            iRect.Bottom = iRect.Bottom + mThemedBodyBottomShadowPixels
            DrawThemeBackground mTheme, picAux.hDC, TABP_PANE, 0&, iRect, iRect
            picBodyThemed.PaintPicture picAux.Image, 0, picAux.ScaleHeight - 1, picAux.ScaleWidth, -picAux.ScaleHeight
        End If
        mThemedBodyReferenceTopBackColor = GetPixel(picBodyThemed.hDC, picBodyThemed.ScaleWidth / 2, picBodyThemed.ScaleHeight * 0.1)
        mBodyThemedReady = True
    End If
End Sub

Private Sub EnsureInactiveBodyThemedReady()
    If Not mInactiveBodyThemedReady Then
        Dim iCA As COLORADJUSTMENT
        
        EnsureBodyThemedReady
        picInactiveBodyThemed.Width = picBodyThemed.Width
        picInactiveBodyThemed.Height = picBodyThemed.Height
        iCA = GetInactiveBodyColorAdjustment
        picAux2.Cls
        picAux2.Width = picBodyThemed.Width
        picAux2.Height = picBodyThemed.Height
        
        SetStretchBltMode picAux2.hDC, HALFTONE
        SetColorAdjustment picAux2.hDC, iCA
        
        StretchBlt picAux2.hDC, 0, 0, picBodyThemed.Width, picBodyThemed.Height, picBodyThemed.hDC, 0, 0, picBodyThemed.Width, picBodyThemed.Height, vbSrcCopy
        picInactiveBodyThemed.Cls
        BitBlt picInactiveBodyThemed.hDC, 0, 0, picAux2.ScaleWidth, picAux2.ScaleHeight, picAux2.hDC, 0, 0, vbSrcCopy
        mInactiveBodyThemedReady = True
    End If
End Sub

Private Sub SetThemeExtraData()
    Dim iRect As RECT
    Dim X As Long
    Dim X2 As Long
    Dim Y As Long
    Dim iCol As Long
    Dim iCol_H As Integer
    Dim iCol_L As Integer
    Dim iCol_S As Integer
    Dim iToChange As Boolean
    Const cHTolerance As Integer = 3
    Const cLTolerance As Integer = 5
    Const cSTolerance As Integer = 14
    Dim iColB As Long
    Dim iColB_H As Integer
    Dim iColB_L As Integer
    Dim iColB_S As Integer
    Dim iThreshold As Integer
    
    If mThemeExtraDataAlreadySet Then Exit Sub
    mThemeExtraDataAlreadySet = True
    
    iRect.Left = 0
    iRect.Top = 0
    iRect.Right = 30
    iRect.Bottom = 30
    picAux.Width = 30
    picAux.Height = 30
    
    DrawThemeBackground mTheme, picAux.hDC, TABP_TABITEM, TIS_NORMAL, iRect, iRect
    mThemedInactiveReferenceBackColorTabs = GetPixel(picAux.hDC, 15, 27)
    ColorRGBToHLS mThemedInactiveReferenceBackColorTabs, mThemedInactiveReferenceBackColorTabs_H, mThemedInactiveReferenceBackColorTabs_L, mThemedInactiveReferenceBackColorTabs_S
    
    ' transparency mask for top left corner of TABITEM and TABITEMRIGHTEDGE
    For X = 0 To 5
        mTABITEM_TopLeftCornerTransparencyMask(X) = 0
    Next X
    For X = 0 To 5
        For Y = 0 To picAux.ScaleHeight - 1
            iToChange = False
            iCol = GetPixel(picAux.hDC, X, Y)
            ColorRGBToHLS iCol, iCol_H, iCol_L, iCol_S
            If Abs(iCol_H - mButtonFace_H) <= cHTolerance Then
                If Abs(iCol_L - mButtonFace_L) <= cLTolerance Then
                    If Abs(iCol_S - mButtonFace_S) <= cSTolerance Then
                        iToChange = True
                    End If
                End If
            End If
            If Not iToChange Then
                If Y < (6) Then
                    mTABITEM_TopLeftCornerTransparencyMask(X) = Y
                Else
                    mTABITEM_TopLeftCornerTransparencyMask(X) = Y - picAux.ScaleHeight - 1 ' negative values point to pixels left to reach the bottom
                End If
                Exit For
            End If
        Next Y
        If Y = picAux.ScaleHeight Then
            mTABITEM_TopLeftCornerTransparencyMask(X) = -1
        End If
        If mTABITEM_TopLeftCornerTransparencyMask(X) = 0 Then Exit For
    Next X
    
    ' transparency mask for top right corner of TABITEM
    For X = 0 To 5
        mTABITEM_TopRightCornerTransparencyMask(X) = 0
    Next X
    For X = 0 To 5
        X2 = picAux.ScaleWidth - 1 - X
        For Y = 0 To picAux.ScaleHeight - 1
            iToChange = False
            iCol = GetPixel(picAux.hDC, X2, Y)
            ColorRGBToHLS iCol, iCol_H, iCol_L, iCol_S
            If Abs(iCol_H - mButtonFace_H) <= cHTolerance Then
                If Abs(iCol_L - mButtonFace_L) <= cLTolerance Then
                    If Abs(iCol_S - mButtonFace_S) <= cSTolerance Then
                        iToChange = True
                    End If
                End If
            End If
            If Not iToChange Then
                If Y < (6) Then
                    mTABITEM_TopRightCornerTransparencyMask(X) = Y
                Else
                    mTABITEM_TopRightCornerTransparencyMask(X) = Y - picAux.ScaleHeight - 1 ' negative values point to pixels left to reach the bottom
                End If
                Exit For
            End If
        Next Y
        If Y = picAux.ScaleHeight Then
            mTABITEM_TopRightCornerTransparencyMask(X) = -1
        End If
        If mTABITEM_TopRightCornerTransparencyMask(X) = 0 Then Exit For
    Next X
    
    ' transparency mask for right side of TABITEMRIGHTEDGE
    picAux.Cls
    DrawThemeBackground mTheme, picAux.hDC, TABP_TABITEMRightEDGE, TIS_NORMAL, iRect, iRect
    For X = 0 To 5
        mTABITEMRIGHTEDGE_RightSideTransparencyMask(X) = 0
    Next X
    For X = 0 To 5
        X2 = picAux.ScaleWidth - 1 - X
        For Y = 0 To picAux.ScaleHeight - 1
            iToChange = False
            iCol = GetPixel(picAux.hDC, X2, Y)
            ColorRGBToHLS iCol, iCol_H, iCol_L, iCol_S
            If Abs(iCol_H - mButtonFace_H) <= cHTolerance Then
                If Abs(iCol_L - mButtonFace_L) <= cLTolerance Then
                    If Abs(iCol_S - mButtonFace_S) <= cSTolerance Then
                        iToChange = True
                    End If
                End If
            End If
            If Not iToChange Then
                If Y < (6) Then
                    mTABITEMRIGHTEDGE_RightSideTransparencyMask(X) = Y
                Else
                    mTABITEMRIGHTEDGE_RightSideTransparencyMask(X) = Y - picAux.ScaleHeight - 1 ' negative values point to pixels left to reach the bottom
                End If
                Exit For
            End If
        Next Y
        If Y = picAux.ScaleHeight Then
            mTABITEMRIGHTEDGE_RightSideTransparencyMask(X) = -1 ' all the column of pixels
        End If
        If mTABITEMRIGHTEDGE_RightSideTransparencyMask(X) = 0 Then Exit For
    Next X
    
    DrawThemeBackground mTheme, picAux.hDC, TABP_PANE, 0&, iRect, iRect
    iColB = GetPixel(picAux.hDC, 15, 10)
    ColorRGBToHLS iColB, iColB_H, iColB_L, iColB_S
    
    mBlendDisablePicWithBackColorTabs_Themed = (iColB_L <= 200)
    If mBlendDisablePicWithBackColorTabs_Themed Then
        mThemedBodyBackColor_R = iColB And 255
        mThemedBodyBackColor_G = (iColB \ 256) And 255
        mThemedBodyBackColor_B = (iColB \ 65536) And 255
    End If
    
    iThreshold = 120
    mThemedBodyBottomShadowPixels = 0
    Do
        For Y = picAux.ScaleHeight - 9 To picAux.ScaleHeight - 1
            iCol = GetPixel(picAux.hDC, 15, Y)
            ColorRGBToHLS iCol, iCol_H, iCol_L, iCol_S
            If Abs(iCol_L - iColB_L) > iThreshold Then
                mThemedBodyBottomShadowPixels = picAux.ScaleHeight - Y - 1
                Exit For
            End If
        Next Y
        If mThemedBodyBottomShadowPixels = 0 Then
            iThreshold = iThreshold - 10
            If iThreshold < 1 Then
                iThreshold = 20
                Exit Do
            End If
        End If
    Loop While mThemedBodyBottomShadowPixels = 0
    
    mThemedBodyRightShadowPixels = 0
    For X = picAux.ScaleWidth - 9 To picAux.ScaleWidth - 1
        iCol = GetPixel(picAux.hDC, X, 15)
        ColorRGBToHLS iCol, iCol_H, iCol_L, iCol_S
        If Abs(iCol_L - iColB_L) > iThreshold Then
            mThemedBodyRightShadowPixels = picAux.ScaleWidth - X - 1
            Exit For
        End If
    Next X
    
    picAux.Cls
End Sub

Private Function GetInactiveBodyColorAdjustment() As COLORADJUSTMENT
    Dim iCA As COLORADJUSTMENT
    Dim iCol As Long
    Dim iCol_H As Integer
    Dim iCol_L As Integer
    Dim iCol_S As Integer
    Dim c As Long
    Dim iLng As Long
    
    picAux2.Width = 1
    picAux2.Height = 1
    picAux2.Cls
    SetStretchBltMode picAux2.hDC, HALFTONE
    GetColorAdjustment picAux2.hDC, iCA
    
    picAux.Width = 1
    picAux.Height = 1
    SetPixelV picAux.hDC, 0, 0, mThemedBodyReferenceTopBackColor
    
    ' luminance
    c = 0
    Do
        c = c + 1
        StretchBlt picAux2.hDC, 0, 0, 1, 1, picAux.hDC, 0, 0, 1, 1, vbSrcCopy
        iCol = GetPixel(picAux2.hDC, 0, 0)
        ColorRGBToHLS iCol, iCol_H, iCol_L, iCol_S
        If Abs(mThemedInactiveReferenceBackColorTabs_L - iCol_L) < 3 Then
            Exit Do
        ElseIf c > 5 Then
            Exit Do
        End If
        iLng = mThemedInactiveReferenceBackColorTabs_L - iCol_L
        If iLng > 50 Then iLng = 50
        If iLng < -50 Then iLng = -50
        iCA.caBrightness = iLng
        SetColorAdjustment picAux2.hDC, iCA
    Loop
    
    GetInactiveBodyColorAdjustment = iCA
End Function

Private Sub ResetCachedThemeImages()
    mBodyThemedReady = False
    mInactiveBodyThemedReady = False
    mSubclassControlsPaintingPending = True
    mRepaintSubclassedControls = True
    mBodyReset = True
End Sub

Private Function MeasureTabIconAndCaption(ByVal nTab As Long) As Long
    Dim iPicWidth As Long
    Dim iCaptionWidth As Long
    Dim iCaptionRect As RECT
    Dim iTabMaxWidth As Long
    Dim iFlags As Long
    Dim iFontBoldPrev As Boolean
    Dim iCaption As String
    Dim iTabHeight As Long
    
    ' pic
    If (Not mTabData(nTab).DoNotUseIconFont) And (mTabData(nTab).IconChar <> 0) Then
        Dim iIconCharacter As String
        Dim iIconCharRect As RECT
        Dim iIconFont As StdFont
        
        If mTabData(nTab).IconFont Is Nothing Then
            Set iIconFont = mDefaultIconFont
        Else
            Set iIconFont = mTabData(nTab).IconFont
        End If
        iIconCharacter = ChrU(mTabData(nTab).IconChar)
        iIconCharRect.Left = 0
        iIconCharRect.Top = 0
        iIconCharRect.Right = 0
        iIconCharRect.Bottom = 0
        iFlags = DT_CALCRECT Or DT_SINGLELINE Or DT_CENTER
        Set picAuxIconFont.Font = iIconFont
        DrawTextW picAuxIconFont.hDC, StrPtr(iIconCharacter), -1, iIconCharRect, iFlags Or IIf(mRightToLeft, DT_RTLREADING, 0)
        If (mTabOrientation = ssTabOrientationTop) Or (mTabOrientation = ssTabOrientationBottom) Then
            iPicWidth = (iIconCharRect.Right - iIconCharRect.Left)
        Else
            iPicWidth = (iIconCharRect.Bottom - iIconCharRect.Top)
        End If
        If (mIconAlignment = ntIconAlignAfterCaption) Or (mIconAlignment = ntIconAlignCenteredAfterCaption) Or (mIconAlignment = ntIconAlignEnd) Then
            If mTabData(nTab).IconLeftOffset > 0 Then
                iPicWidth = iPicWidth + mTabData(nTab).IconLeftOffset * mDPIScale
            End If
        ElseIf (mIconAlignment = ntIconAlignBeforeCaption) Or (mIconAlignment = ntIconAlignCenteredBeforeCaption) Or (mIconAlignment = ntIconAlignStart) Then
            If mTabData(nTab).IconLeftOffset < 0 Then
                iPicWidth = iPicWidth - mTabData(nTab).IconLeftOffset * mDPIScale
            End If
        End If
    Else
        iPicWidth = 0
        If Not mTabData(nTab).PicToUseSet Then SetPicToUse nTab
        If Not mTabData(nTab).PicToUse Is Nothing Then
            If (mTabOrientation = ssTabOrientationTop) Or (mTabOrientation = ssTabOrientationBottom) Then
                iPicWidth = pScaleX(mTabData(nTab).PicToUse.Width, vbHimetric, vbPixels)
            Else
                iPicWidth = pScaleX(mTabData(nTab).PicToUse.Height, vbHimetric, vbPixels)
            End If
        End If
    End If
    
    ' caption
    iFontBoldPrev = picAux.FontBold
    
    picAux.FontBold = mFont.Bold Or mHighlightCaptionBoldSelectedTab Or mHighlightCaptionBold
    picAux.FontUnderLine = mFont.Underline Or mHighlightCaptionUnderlinedSelectedTab Or mHighlightCaptionUnderlined
    
    With mTabData(nTab).TabRect
        iCaptionRect.Left = 0
        iCaptionRect.Top = 0
        iCaptionRect.Bottom = .Bottom - .Top - 4
        iCaptionRect.Right = mScaleWidth
    End With
    
    If mTabMaxWidth > 0 Then
        iTabMaxWidth = pScaleX(mTabMaxWidth, vbHimetric, vbPixels)
        iCaptionRect.Right = iTabMaxWidth
    End If
    
    iFlags = DT_CALCRECT Or DT_SINGLELINE Or mAlignmentDTFlag Or DT_VCENTER
    iCaption = mTabData(nTab).Caption & IIf(picAux.Font.Italic, " ", "") & IIf(mVisualStyles Or mAppearanceIsFlat, " ", "")
    DrawTextW picAux.hDC, StrPtr(iCaption), -1, iCaptionRect, iFlags
    iCaptionWidth = iCaptionRect.Right
    
    If picAux.FontBold <> iFontBoldPrev Then
        picAux.FontBold = iFontBoldPrev
    End If
    
    If mIconAlignment = ntIconAlignCenteredOnTab Then
        If iPicWidth > iCaptionWidth Then
            MeasureTabIconAndCaption = iPicWidth
        Else
            MeasureTabIconAndCaption = iCaptionWidth
        End If
    ElseIf (mIconAlignment = ntIconAlignAtTop) Or (mIconAlignment = ntIconAlignAtBottom) Then
        If iPicWidth > iCaptionWidth Then
            MeasureTabIconAndCaption = iPicWidth
        Else
            MeasureTabIconAndCaption = iCaptionWidth
        End If
        MeasureTabIconAndCaption = MeasureTabIconAndCaption + 4
    Else
        MeasureTabIconAndCaption = iPicWidth + mTabIconDistanceToCaptionDPIScaled + iCaptionWidth
    End If
    
    ' TDI mode forms
    If mTDIMode = ntTDIModeForms Then
        If mTabData(nTab).Data > 0 Then
            Dim iTabData As T_TabData
            Dim iTDIFormsNoIcons As Boolean
            Dim iTDIFormIcon_Icon As StdPicture
            Dim iTDIFormIcon_Width As Long
            Dim iTDIFormIcon_Height As Long
            Dim iTDIFormIcon_UseIconFont As Boolean
            Dim iTDIFormIcon_IconFont As StdFont
            Dim iTDIFormIcon_IconCharRect As RECT
            Dim iTDIFormIcon_IconChar As Variant
            Dim iTDIFormIcon_IconCharStr As String
            Dim iTDIFormsFontPrev As StdFont
            Dim iTDIFormIcon_IconColor As Long
            Dim iTDIFormIcon_IconColorMouseHover As Long
            Dim iTDIFormIcon_IconColorSelected As Long
            Dim iTDIFormIcon_LeftOffset As Long
            Dim iTDIFormIcon_TopOffset As Long
            Dim iTabPosString As String
            Dim iNoIcon As Boolean
            
            iTabData = mTabData(nTab)
            
            iTDIFormsNoIcons = True
            iTabPosString = CStr(iTabData.PosH) & "-" & CStr(iTabData.RowPos)
            If (mTDIFormIconCustomData(iTabData.Data).hWnd <> mTDIModeFormsFormData_FormHwnd(iTabData.Data)) Then
                If (mTDIFormIconCustomData(iTabData.Data).hWnd = mTDIModeFormsFormData_FormHwnd(iTabData.Data)) And (Not mTDIFormIconCustomData(iTabData.Data).Icon Is Nothing) Then
                    Set iTDIFormIcon_Icon = mTDIFormIconCustomData(iTabData.Data).Icon
                Else
                    Set iTDIFormIcon_Icon = mTDIModeFormsFormData_FormIcon(iTabData.Data)
                End If
                
                If Not iTDIFormIcon_Icon Is Nothing Then
                    iTDIFormIcon_Width = UserControl.ScaleX(iTDIFormIcon_Icon.Width, vbHimetric, vbPixels)
                    iTDIFormIcon_Height = UserControl.ScaleY(iTDIFormIcon_Icon.Height, vbHimetric, vbPixels)
                Else
                    iTDIFormIcon_Width = 32
                    iTDIFormIcon_Height = 32
                End If
                
                iTDIFormIcon_UseIconFont = False
                Set iTDIFormIcon_IconFont = New StdFont
                If (Not mTabData(nTab).IconFont Is Nothing) Then
                    If FontExists(iTabData.IconFont.Name) Then
                        iTDIFormIcon_IconFont.Name = iTabData.IconFont.Name
                        iTDIFormIcon_IconFont.Size = 41
                        iTDIFormIcon_IconFont.Bold = True
                        iTDIFormIcon_IconChar = "&HE160&"
                    Else
                        iTDIFormIcon_IconFont.Name = "Arial"
                        iTDIFormIcon_IconFont.Size = 41
                        iTDIFormIcon_IconFont.Bold = True
                        iTDIFormIcon_IconChar = "&H2666&"
                    End If
                End If
                
                If mTabData(0).IconChar <> 0 Then
                    iTDIFormIcon_IconFont.Size = mTabData(0).IconFont.Size
                Else
                    Set iTDIFormsFontPrev = picDraw.Font
                    Set picDraw.Font = iTDIFormIcon_IconFont
                    iTabHeight = (iTabData.TabRect.Bottom - iTabData.TabRect.Top)
                    If iTabHeight = 0 Then iTabHeight = ScaleY(mTabHeight, vbHimetric, vbPixels)
                    Do
                        iTDIFormIcon_IconFont.Size = iTDIFormIcon_IconFont.Size - 1
                        iTDIFormIcon_IconCharRect.Left = 0
                        iTDIFormIcon_IconCharRect.Right = 0
                        iTDIFormIcon_IconCharRect.Top = 0
                        iTDIFormIcon_IconCharRect.Bottom = 0
                        DrawTextW picDraw.hDC, StrPtr(iTDIFormIcon_IconChar), -1, iTDIFormIcon_IconCharRect, DT_CALCRECT Or DT_SINGLELINE Or DT_CENTER Or IIf(mRightToLeft, DT_RTLREADING, 0)
                        If iTDIFormIcon_IconFont.Size < 5 Then Exit Do
                    Loop While ((iTDIFormIcon_IconCharRect.Bottom - iTDIFormIcon_IconCharRect.Top) + 8) > iTabHeight
                    Set picDraw.Font = iTDIFormsFontPrev
                End If
                
                iTDIFormIcon_IconColor = mIconColor
                iTDIFormIcon_IconColorMouseHover = mIconColorTabHighlighted
                iTDIFormIcon_IconColorSelected = mIconColorSelectedTab
                
                iTDIFormIcon_LeftOffset = 0
                iTDIFormIcon_TopOffset = 0
                
                RaiseEvent TDIFormsShowTabIcon(nTab, GetWindowCaption(mTDIModeFormsFormData_FormHwnd(iTabData.Data)), mTDIModeFormsFormData_FormHwnd(iTabData.Data), iTDIFormIcon_Icon, iTDIFormIcon_LeftOffset, iTDIFormIcon_TopOffset, iTDIFormIcon_Width, iTDIFormIcon_Height, iTDIFormIcon_UseIconFont, iTDIFormIcon_IconFont, iTDIFormIcon_IconChar, iTDIFormIcon_IconColor, iTDIFormIcon_IconColorMouseHover, iTDIFormIcon_IconColorSelected, iNoIcon)
                
                If iNoIcon Then iTDIFormIcon_Width = 0
                If iTDIFormIcon_UseIconFont Then
                    mTDIFormIconCustomData(iTabData.Data).UseIconFont = True
                    Set mTDIFormIconCustomData(iTabData.Data).IconFont = iTDIFormIcon_IconFont
                    If (VarType(iTDIFormIcon_IconChar) = vbString) And (Not IsNumeric(iTDIFormIcon_IconChar)) Then
                        If Left$(iTDIFormIcon_IconChar, 2) = "&H" Then
                            mTDIFormIconCustomData(iTabData.Data).IconCharStr = ChrU(Val(iTDIFormIcon_IconChar))
                        Else
                            mTDIFormIconCustomData(iTabData.Data).IconCharStr = iTDIFormIcon_IconChar
                        End If
                    ElseIf IsNumeric(iTDIFormIcon_IconChar) Then
                        mTDIFormIconCustomData(iTabData.Data).IconCharStr = ChrU(iTDIFormIcon_IconChar)
                    Else
                        RaiseError 1385, TypeName(Me), "IconChar format not recognized."
                    End If
                    iTDIFormIcon_IconCharStr = mTDIFormIconCustomData(iTabData.Data).IconCharStr
                    
                    If IsValidOLE_COLOR(iTDIFormIcon_IconColor) Then
                        mTDIFormIconCustomData(iTabData.Data).IconColor = iTDIFormIcon_IconColor
                    Else
                        mTDIFormIconCustomData(iTabData.Data).IconColor = mIconColor
                    End If
                    If IsValidOLE_COLOR(iTDIFormIcon_IconColorMouseHover) Then
                        mTDIFormIconCustomData(iTabData.Data).IconColorMouseHover = iTDIFormIcon_IconColorMouseHover
                    Else
                        mTDIFormIconCustomData(iTabData.Data).IconColorMouseHover = mIconColorTabHighlighted
                    End If
                    If IsValidOLE_COLOR(iTDIFormIcon_IconColorSelected) Then
                        mTDIFormIconCustomData(iTabData.Data).IconColorSelected = iTDIFormIcon_IconColorSelected
                    Else
                        mTDIFormIconCustomData(iTabData.Data).IconColorSelected = mIconColorSelectedTab
                    End If
                    
                    iTDIFormIcon_IconCharRect.Left = 0
                    iTDIFormIcon_IconCharRect.Right = 0
                    iTDIFormIcon_IconCharRect.Top = 0
                    iTDIFormIcon_IconCharRect.Bottom = 0
                    
                    Set iTDIFormsFontPrev = picDraw.Font
                    Set picDraw.Font = iTDIFormIcon_IconFont
                    
                    DrawTextW picDraw.hDC, StrPtr(iTDIFormIcon_IconCharStr), -1, iTDIFormIcon_IconCharRect, DT_CALCRECT Or DT_SINGLELINE Or DT_CENTER Or IIf(mRightToLeft, DT_RTLREADING, 0)
                    
                    Set picDraw.Font = iTDIFormsFontPrev
                    
                    iTDIFormIcon_Width = iTDIFormIcon_IconCharRect.Right - iTDIFormIcon_IconCharRect.Left
                    iTDIFormIcon_Height = iTDIFormIcon_IconCharRect.Bottom - iTDIFormIcon_IconCharRect.Top
                Else
                    If iTDIFormIcon_Height <= 0 Then iTDIFormIcon_Width = 0
                    Set mTDIFormIconCustomData(iTabData.Data).Icon = iTDIFormIcon_Icon
                    If iTDIFormIcon_Icon Is Nothing Then
                        iTDIFormIcon_Width = 0
                    ElseIf iTDIFormIcon_Icon.Handle = 0 Then
                        iTDIFormIcon_Width = 0
                    End If
                End If
                mTDIFormIconCustomData(iTabData.Data).hWnd = mTDIModeFormsFormData_FormHwnd(iTabData.Data)
                mTDIFormIconCustomData(iTabData.Data).TabPosString = iTabPosString
                mTDIFormIconCustomData(iTabData.Data).Width = iTDIFormIcon_Width
                mTDIFormIconCustomData(iTabData.Data).Height = iTDIFormIcon_Height
                mTDIFormIconCustomData(iTabData.Data).LeftOffset = iTDIFormIcon_LeftOffset
                mTDIFormIconCustomData(iTabData.Data).TopOffset = iTDIFormIcon_TopOffset
            End If
            
            If Not mTDIModeFormsFormData_FormIcon(mTabData(nTab).Data) Is Nothing Then
                MeasureTabIconAndCaption = MeasureTabIconAndCaption + mTDIFormIconCustomData(iTabData.Data).Width + 4
            Else
                MeasureTabIconAndCaption = MeasureTabIconAndCaption + 4
            End If
        End If
    End If
End Function

Public Function IsThemed() As Boolean
Attribute IsThemed.VB_Description = "Returns a boolean value indicating whether the Windows visual styles are applied to the control or not."
    Dim iTheme As Long
    
    IsThemed = mVisualStyles And (mBackStyle <> ntTransparent)
    If IsThemed Then
        iTheme = OpenThemeData(mUserControlHwnd, StrPtr("Tab"))
        If iTheme = 0 Then
            IsThemed = False
        Else
            CloseThemeData iTheme
        End If
    End If
End Function

' Obsolete, hidden , left just for binary compatibility
Public Property Get ForceVisualStyles() As Boolean
Attribute ForceVisualStyles.VB_Description = "Hidden property intended for testing purposes. Allows the control to show visual styles on an un-themed IDE."
Attribute ForceVisualStyles.VB_MemberFlags = "40"
    '
End Property

Public Property Let ForceVisualStyles(ByVal nValue As Boolean)
    '
End Property

Private Sub SetVisibleControls(ByVal iPreviousTab As Long)
    Dim iCtl As Object
    Dim iCtlName As Variant
    Dim iContainedControlsString As String
    Dim iHwnd As Long
    Dim c As Long
    Dim iLeft As Long
    Dim iIsLine As Boolean
    
    If mUserControlTerminated Then Exit Sub
    If Not mAmbientUserMode Then CheckIfContainedControlChangedToArray
    
    If (Not mAmbientUserMode) And mChangeControlsBackColor And (mBackColorSelectedTab <> vbButtonFace) Then
        iContainedControlsString = GetContainedControlsString
        If iContainedControlsString <> mLastContainedControlsString Then
            SetControlsBackColor mBackColorSelectedTab
        End If
    End If
    If (Not mAmbientUserMode) And mChangeControlsForeColor And (mForeColorSelectedTab <> vbButtonText) Then
        iContainedControlsString = GetContainedControlsString
        If iContainedControlsString <> mLastContainedControlsString Then
            SetControlsProperForeColor
        End If
    End If
    
    If mSubclassedControlsForMoveHwnds.Count > 0 Then
        For c = 1 To mSubclassedControlsForMoveHwnds.Count
            iHwnd = mSubclassedControlsForMoveHwnds(c)
            DetachMessage Me, iHwnd, WM_WINDOWPOSCHANGING
        Next c
        Set mSubclassedControlsForMoveHwnds = New Collection
    End If
    
    If mPendingLeftOffset <> 0 Then
        DoPendingLeftOffset
    End If
    
    ' hide controls in previous tab
    If mAmbientUserMode Then StoreControlsTabStop
    If (iPreviousTab >= 0) And (iPreviousTab <= UBound(mTabData)) Then
        Set mTabData(iPreviousTab).Controls = New Collection
    End If
    For Each iCtl In UserControlContainedControls
        iIsLine = TypeName(iCtl) = "Line"
        iLeft = -15001
        On Error Resume Next
        If iIsLine Then
            iLeft = iCtl.X1
        Else
            iLeft = iCtl.Left
        End If
        On Error GoTo 0
        If iLeft > -mLeftThresholdHided Then
            iCtlName = ControlName(iCtl)
            If (iPreviousTab >= 0) And (iPreviousTab <= UBound(mTabData)) Then
                If Not IsControlInOtherTab(iCtlName, iPreviousTab) Then
                    mTabData(iPreviousTab).Controls.Add iCtlName, iCtlName
                End If
            End If
            If iIsLine Then
                iCtl.X1 = iCtl.X1 - mLeftOffsetToHide
                iCtl.X2 = iCtl.X2 - mLeftOffsetToHide
            Else
                iCtl.Left = iCtl.Left - mLeftOffsetToHide
            End If
        End If
    Next
    
    ' show controls in selected tab
    If (mTab > -1) And (mTab < mTabs) Then
        For Each iCtlName In mTabData(mTab).Controls
            Set iCtl = GetContainedControlByName(iCtlName)
            If Not iCtl Is Nothing Then
                On Error Resume Next
                iIsLine = TypeName(iCtl) = "Line"
                If iIsLine Then
                    iCtl.X1 = iCtl.X1 + mLeftOffsetToHide
                    iCtl.X2 = iCtl.X2 + mLeftOffsetToHide
                Else
                    iCtl.Left = iCtl.Left + mLeftOffsetToHide
                End If
                On Error GoTo 0
                If mAmbientUserMode Then
                    On Error Resume Next
                    iCtl.TabStop = mParentControlsTabStop(iCtlName)
                    iCtl.UseMnemonic = mParentControlsUseMnemonic(iCtlName)
                    If TypeName(iCtl) = "ComboBox" Then
                        ' ComboBox fix
                        If iCtl.Style = vbComboDropdown Then
                            iCtl.SelLength = 0
                        End If
                    End If
                    On Error GoTo 0
                    If ControlIsContainer(iCtlName) Then
                        SetTabStopsToParentControlsContainedInControl iCtl
                    End If
                End If
            End If
        Next
    End If
    
    If (Not mAmbientUserMode) And (mChangeControlsBackColor Or mChangeControlsForeColor) Then
        mLastContainedControlsString = iContainedControlsString
    End If
    
    If mSubclassed And (Not mOnlySubclassUserControl) Then
        On Error Resume Next
        For Each iCtl In UserControlContainedControls
            If iCtl.Left < -mLeftThresholdHided Then
                iHwnd = 0
                iHwnd = GetControlHwnd(iCtl)
                If iHwnd <> 0 Then
                    mSubclassedControlsForMoveHwnds.Add iHwnd
                    AttachMessage Me, iHwnd, WM_WINDOWPOSCHANGING
                End If
            End If
        Next
        Err.Clear
    End If
    
    If mTDIMode = ntTDIModeForms Then
        If iPreviousTab > -1 Then
            If (mTabData(iPreviousTab).Data >= picTDIFormContainer.LBound) And (mTabData(iPreviousTab).Data <= picTDIFormContainer.UBound) Then
               picTDIFormContainer(mTabData(iPreviousTab).Data).Visible = False
            End If
        End If
        If mTab > -1 Then
            If mTabData(mTab).Data > 0 Then
                If (mTabData(mTab).Data >= picTDIFormContainer.LBound) And (mTabData(mTab).Data <= picTDIFormContainer.UBound) Then
                    picTDIFormContainer(mTabData(mTab).Data).Visible = True
                End If
            End If
        End If
    End If
End Sub

Private Function GetControlHwnd(nControl As Object) As Long
    On Error Resume Next
    GetControlHwnd = nControl.hWndUserControl
    If GetControlHwnd = 0 Then
        GetControlHwnd = nControl.hWnd
    End If
End Function

Private Function GetControlHwnd2(nControl As Object) As Long
    On Error Resume Next
    GetControlHwnd2 = nControl.hWnd
End Function

Private Function IsControlInOtherTab(nCtlName As Variant, ByVal nTab As Long) As Boolean
    Dim t As Long
    Dim iStr As String
    
    On Error Resume Next
    For t = 0 To mTabs - 1
        If t <> nTab Then
            iStr = ""
            iStr = mTabData(t).Controls(nCtlName)
            If iStr <> "" Then
                IsControlInOtherTab = True
                Exit Function
            End If
        End If
    Next t
End Function

Private Function GetContainedControlsString() As String
    Dim iCtl As Object
    
    For Each iCtl In UserControlContainedControls
        GetContainedControlsString = GetContainedControlsString & iCtl.Name
    Next
End Function

Private Sub StoreControlsTabStop(Optional nInitialize As Boolean)
    Dim iControls As Object
    Dim iCtl As Object
    Dim iContainer As Object
    Dim iContainer_Prev As Object
    Dim iStr As String
    Dim iParent As Object
    Dim iVisible As Boolean
    
    On Error Resume Next
    Set iParent = UserControl.Parent
    Set iControls = iParent.Controls
    If iControls Is Nothing Then ' this parent doesn't have a controls collection
        Set iControls = UserControlContainedControls ' let's do it just with the contained controls then
    End If
    For Each iCtl In iControls
        Set iContainer_Prev = Nothing
        Set iContainer = Nothing
        Set iContainer = iCtl.Container
        Do Until iContainer Is Nothing
            If iContainer Is UserControl.Extender Then
                iVisible = False
                If Not (iContainer_Prev Is Nothing Or iContainer_Prev Is iParent) Then ' the control is contained in another control that is contained in the usercontrol
                    iVisible = iContainer_Prev.Left > -mLeftThresholdHided
                    If iVisible Or nInitialize Then
                        iStr = ControlName(iCtl)
                        mParentControlsTabStop.Add iCtl.TabStop, iStr
                        mParentControlsUseMnemonic.Add iCtl.UseMnemonic, iStr
                        iStr = ControlName(iContainer_Prev)
                        mContainedControlsThatAreContainers.Add iStr, iStr
                        If nInitialize Then
                            If Not iVisible Then
                                iCtl.TabStop = False
                                iCtl.UseMnemonic = False
                            End If
                        Else
                            iCtl.TabStop = False
                            iCtl.UseMnemonic = False
                        End If
                    End If
                Else ' the control is directly contained in the usercontrol
                    iVisible = iCtl.Left > -mLeftThresholdHided
                    If iVisible Or nInitialize Then
                        iStr = ControlName(iCtl)
                        mParentControlsTabStop.Add iCtl.TabStop, iStr
                        mParentControlsUseMnemonic.Add iCtl.UseMnemonic, iStr
                        If nInitialize Then
                            If Not iVisible Then
                                iCtl.TabStop = False
                                iCtl.UseMnemonic = False
                            End If
                        Else
                            iCtl.TabStop = False
                            iCtl.UseMnemonic = False
                        End If
                    End If
                End If
                Exit Do
            End If
            Set iContainer_Prev = iContainer
            Set iContainer = Nothing
            Set iContainer = iContainer_Prev.Container
        Loop
    Next
    mTabStopsInitialized = True
    Err.Clear
End Sub

Private Sub SubclassControlsPainting()
    Dim iSubclassTheControls As Boolean
    Dim iHwnd As Long
    Dim c As Long
    Dim iBKColor As Long
    Dim iControls As Object
    Dim iCtl As Object
    Dim iContainer As Object
    Dim iContainer_Prev As Object
    Dim iParent As Object
    Dim iVisible As Boolean
    Dim iBackColorTabs As Long
    Dim iCancel As Boolean
    Dim iClantNotHandled As Boolean
    Dim iCtlTypeName As String
    Dim iStr As String
    
  '  If Not mAmbientUserMode Then Exit Sub
    If (Not mSubclassed) Or mOnlySubclassUserControl Then Exit Sub
    If Not mUserControlShown Then
        If Val(tmrSubclassControls.Tag) < 200 Then
            tmrSubclassControls.Enabled = True
        End If
        tmrSubclassControls.Tag = Val(tmrSubclassControls.Tag) + 1
        Exit Sub
    End If
    tmrSubclassControls.Tag = ""
    mSubclassControlsPaintingPending = False
    
    iSubclassTheControls = mVisualStyles And mChangeControlsBackColor
    If mSubclassedControlsForPaintingHwnds.Count > 0 Then
        For c = 1 To mSubclassedControlsForPaintingHwnds.Count
            iHwnd = mSubclassedControlsForPaintingHwnds(c)
            DetachMessage Me, iHwnd, WM_PAINT
            DetachMessage Me, iHwnd, WM_MOVE
            If Not iSubclassTheControls And mRepaintSubclassedControls Then
                ' redraw the control
                RedrawWindow iHwnd, ByVal 0&, 0&, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_INTERNALPAINT Or RDW_ALLCHILDREN
            End If
        Next c
        Set mSubclassedControlsForPaintingHwnds = New Collection
    End If
    
    If mSubclassedFramesHwnds.Count > 0 Then
        For c = 1 To mSubclassedFramesHwnds.Count
            iHwnd = mSubclassedFramesHwnds(c)
            DetachMessage Me, iHwnd, WM_PRINTCLIENT
            DetachMessage Me, iHwnd, WM_MOUSELEAVE
        Next c
        Set mSubclassedFramesHwnds = New Collection
    End If
    
    If Not iSubclassTheControls Then
        mRepaintSubclassedControls = False
'        Exit Sub
    End If
    
    If mChangeControlsBackColor Then
        If Not mChangedControlsBackColor Then
            SetControlsBackColor mBackColorSelectedTab
            mChangedControlsBackColor = True
        End If
    End If
    If mChangeControlsForeColor Then
        If Not mChangedControlsForeColor Then
            SetControlsProperForeColor
            mChangedControlsForeColor = True
        End If
    End If
    
    If mShowDisabledState And (Not mEnabled) Then
        iBackColorTabs = mBackColorSelectedTabDisabled
    Else
        iBackColorTabs = mBackColorSelectedTab
    End If
    
    On Error Resume Next
    Set iParent = UserControl.Parent
    Set iControls = iParent.Controls
    If iControls Is Nothing Then ' this parent doesn't have a controls collection
        Set iControls = UserControlContainedControls ' let's do it just with the contained controls then
    End If
    For Each iCtl In iControls
        iCtlTypeName = TypeName(iCtl)
        iClantNotHandled = (iCtlTypeName = "ButtonEx") Or (iCtlTypeName = "ButtonExNoFocus")
        If Not iClantNotHandled Then
            Set iContainer_Prev = Nothing
            Set iContainer = Nothing
            Set iContainer = iCtl.Container
            If iContainer Is Nothing Then
                iHwnd = 0
                iHwnd = GetControlHwnd2(iCtl)
                If iHwnd <> 0 Then
                    Set iContainer = GetContainerByHwnd(iHwnd)
                End If
            End If
            Do Until iContainer Is Nothing
                If iContainer Is UserControl.Extender Then
                    iVisible = False
                    If Not (iContainer_Prev Is Nothing Or iContainer_Prev Is iParent) Then ' the control is contained in another control that is contained in the usercontrol
                        iVisible = iContainer_Prev.Left > -mLeftThresholdHided
                        If iVisible Then
                            iHwnd = 0
                            iHwnd = GetControlHwnd2(iCtl)
                            If iHwnd <> 0 Then
                                If iSubclassTheControls Then
                                    iBKColor = -1
                                    iBKColor = iCtl.BackColor
                                    If (iBKColor = iBackColorTabs) Then
                                        iCancel = False
                                        If iContainer_Prev.Container Is UserControl.Extender Then
                                            iStr = iContainer_Prev.Name
                                            RaiseEvent ChangeControlBackColor(iStr, TypeName(iContainer_Prev), iCancel)
                                        End If
                                        If Not iCancel Then
                                            iCancel = False
                                            iStr = iCtl.Name
                                            RaiseEvent ChangeControlBackColor(iStr, TypeName(iCtl), iCancel)
                                            If Not iCancel Then
                                                mSubclassedControlsForPaintingHwnds.Add iHwnd, CStr(iHwnd)
                                            End If
                                        End If
                                    End If
                                End If
                                If iCtlTypeName = "Frame" Then
                                    mSubclassedFramesHwnds.Add iHwnd, CStr(iHwnd)
                                End If
                            ElseIf (iCtlTypeName = "Label") Or (iCtlTypeName = "LabelW") Then
                                If iCtl.BackStyle = 1 Then ' solid
                                    If iSubclassTheControls Then
                                        iBKColor = -1
                                        iBKColor = iCtl.BackColor
                                        If (iBKColor = iBackColorTabs) Then
                                            iCancel = False
                                            If iContainer_Prev.Container Is UserControl.Extender Then
                                                iStr = iContainer_Prev.Name
                                                RaiseEvent ChangeControlBackColor(iStr, TypeName(iCtl), iCancel)
                                            End If
                                            If Not iCancel Then
                                                iCancel = False
                                                iStr = iCtl.Name
                                                RaiseEvent ChangeControlBackColor(iStr, TypeName(iCtl), iCancel)
                                                If Not iCancel Then
                                                    iCtl.BackStyle = 0 ' transparent
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else ' the control is directly contained in the usercontrol
                        iVisible = iCtl.Left > -mLeftThresholdHided
                        If iVisible Then
                            iHwnd = 0
                            iHwnd = GetControlHwnd2(iCtl)
                            If iHwnd <> 0 Then
                                If iSubclassTheControls Then
                                    iBKColor = -1
                                    iBKColor = iCtl.BackColor
                                    If (iBKColor = iBackColorTabs) Then
                                        iCancel = False
                                        iStr = iCtl.Name
                                        RaiseEvent ChangeControlBackColor(iStr, TypeName(iCtl), iCancel)
                                        If Not iCancel Then
                                            mSubclassedControlsForPaintingHwnds.Add iHwnd, CStr(iHwnd)
                                        End If
                                    End If
                                End If
                                If iCtlTypeName = "Frame" Then
                                    mSubclassedFramesHwnds.Add iHwnd, CStr(iHwnd)
                                End If
                            ElseIf (iCtlTypeName = "Label") Or (iCtlTypeName = "LabelW") Then
                                If iCtl.BackStyle = 1 Then ' solid
                                    If iSubclassTheControls Then
                                        iBKColor = -1
                                        iBKColor = iCtl.BackColor
                                        If (iBKColor = iBackColorTabs) Then
                                            iCancel = False
                                            If iContainer_Prev.Container Is UserControl.Extender Then
                                                iStr = iContainer_Prev.Name
                                                RaiseEvent ChangeControlBackColor(iStr, TypeName(iContainer_Prev), iCancel)
                                            End If
                                            If Not iCancel Then
                                                iCancel = False
                                                iStr = iCtl.Name
                                                RaiseEvent ChangeControlBackColor(iStr, TypeName(iCtl), iCancel)
                                                If Not iCancel Then
                                                    iCtl.BackStyle = 0 ' transparent
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                    Exit Do
                End If
                Set iContainer_Prev = iContainer
                Set iContainer = Nothing
                Set iContainer = iContainer_Prev.Container
            Loop
        End If
    Next
    On Error GoTo 0
    
    
    For c = 1 To mSubclassedFramesHwnds.Count
        iHwnd = mSubclassedFramesHwnds(c)
        AttachMessage Me, iHwnd, WM_PRINTCLIENT
        AttachMessage Me, iHwnd, WM_MOUSELEAVE
    Next
    For c = 1 To mSubclassedControlsForPaintingHwnds.Count
        iHwnd = mSubclassedControlsForPaintingHwnds(c)
        AttachMessage Me, iHwnd, WM_PAINT
        AttachMessage Me, iHwnd, WM_MOVE
        If mRepaintSubclassedControls Then
            ' redraw the control
            RedrawWindow iHwnd, ByVal 0&, 0&, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_INTERNALPAINT Or RDW_ALLCHILDREN
        End If
    Next c
    mRepaintSubclassedControls = False
    
End Sub

Private Function GetContainerByHwnd(ByVal nHwnd As Long) As Object
    Dim iParent As Object
    Dim iControls As Object
    Dim iCtl As Object
    Dim iHwndParent As Long
    Dim iHwnd As Long
    
    On Error Resume Next
    Set iParent = UserControl.Extender.Parent
    If iParent Is Nothing Then GoTo Exit_Function
    Set iControls = iParent.Controls
    If iControls Is Nothing Then GoTo Exit_Function
    
    iHwndParent = GetParent(nHwnd)
    
    For Each iCtl In iControls
        iHwnd = 0
        iHwnd = GetControlHwnd(iCtl)
        If iHwnd = iHwndParent Then
            Set GetContainerByHwnd = iCtl
            GoTo Exit_Function
        End If
    Next
    
Exit_Function:
    Err.Clear
End Function

Private Function ControlIsContainer(nControlName As Variant) As Boolean
    Dim iStr As String
    
    On Error Resume Next
    iStr = mContainedControlsThatAreContainers(nControlName)
    ControlIsContainer = Err.Number = 0
    Err.Clear
End Function

Private Sub SetTabStopsToParentControlsContainedInControl(nContainer As Object)
    Dim iControls As Object
    Dim iCtl As Object
    Dim iContainer As Object
    Dim iStr As String
    Dim iObj As Object
    
    If nContainer Is Nothing Then Exit Sub
    On Error Resume Next
    Set iControls = GetContainedControlsInControlContainer(nContainer)
    If Not iControls Is Nothing Then
        For Each iCtl In iControls
            Set iContainer = Nothing
            Set iContainer = iCtl.Container
            Do Until iContainer Is Nothing
                If iContainer Is nContainer Then
                    iStr = ControlName(iCtl)
                    iCtl.TabStop = mParentControlsTabStop(iStr)
                    iCtl.UseMnemonic = mParentControlsUseMnemonic(iStr)
                    If TypeName(iCtl) = "ComboBox" Then
                        ' ComboBox fix
                        If iCtl.Style = vbComboDropdown Then
                            iCtl.SelLength = 0
                        End If
                    End If
                End If
                Set iObj = iContainer
                Set iContainer = Nothing
                Set iContainer = iObj.Container
            Loop
        Next
    End If
    Err.Clear
End Sub

Private Function GetContainedControlsInControlContainer(nContainer As Object) As Object
    Dim iControls As Object
    Dim iCtl As Object
    Dim iContainer As Object
    Dim iContainer_Prev As Object
    
    Set GetContainedControlsInControlContainer = New Collection
    
    If nContainer Is Nothing Then Exit Function
    On Error Resume Next
    Set iControls = UserControl.Parent.Controls
    If iControls Is Nothing Then GoTo Exit_Function
    
    For Each iCtl In iControls
        Set iContainer_Prev = Nothing
        Set iContainer = Nothing
        Set iContainer = iCtl.Container
        Do Until iContainer Is Nothing
            If iContainer Is nContainer Then
                GetContainedControlsInControlContainer.Add iCtl
            End If
            Set iContainer_Prev = iContainer
            Set iContainer = Nothing
            Set iContainer = iContainer_Prev.Container
        Loop
    Next
    
Exit_Function:
    Err.Clear
End Function

Private Function ControlName(nCtl As Object) As String
    Dim iIndex As Integer
    
    On Error GoTo NoIndex:
    ControlName = nCtl.Name
    iIndex = -1
    iIndex = nCtl.Index
    If iIndex >= 0 Then
        ControlName = ControlName & "(" & iIndex & ")"
    End If

NoIndex:
End Function

Private Function GetContainedControlByName(ByVal nControlName As String) As Object
    Dim iCtl As Object

    On Error GoTo ErrorExit
    For Each iCtl In UserControlContainedControls
        If StrComp(nControlName, ControlName(iCtl), vbTextCompare) = 0 Then
            Set GetContainedControlByName = iCtl
            Exit For
        End If
    Next
    
ErrorExit:
End Function

Private Sub SetAccessKeys()
    Dim c As Long
    Dim iPos As Long
    Dim iChr As String
    Dim iAsc As Long
    Dim iAK As String
    
    mAccessKeys = ""
    iAK = ""
    
    For c = 0 To mTabs - 1
        iChr = ""
        If mTabData(c).Enabled And mTabData(c).Visible Then
            iPos = InStr(mTabData(c).Caption, "&")
            If iPos > 0 Then
                iChr = LCase$(Mid$(mTabData(c).Caption, iPos + 1, 1))
                If (iChr <> "") Then
                    iAsc = Asc(iChr)
                    If Not (((iAsc > 47) And (iAsc < 58)) Or ((iAsc > 96) And (iAsc < 123))) Then
                        iChr = ""
                    End If
                End If
            End If
        End If
        iAK = iAK & iChr
        If iChr = "" Then iChr = Chr$(0)
        mAccessKeys = mAccessKeys & iChr
    Next c
    UserControl.AccessKeys = iAK
    mAccessKeysSet = True
End Sub

Private Sub SetPicToUse(ByVal nTab As Long)
    Dim iTx As Single
    
    If mTabData(nTab).PicToUseSet Then Exit Sub
    
    iTx = Screen_TwipsPerPixelX
    If Not mTabData(nTab).Pic16 Is Nothing Then
        If iTx >= 15 Then ' 96 DPI
            Set mTabData(nTab).PicToUse = mTabData(nTab).Pic16
        ElseIf iTx >= 12 Then ' 120 DPI
            If Not mTabData(nTab).Pic20 Is Nothing Then
                Set mTabData(nTab).PicToUse = mTabData(nTab).Pic20
            Else
                Set mTabData(nTab).PicToUse = mTabData(nTab).Pic16
            End If
        ElseIf iTx >= 10 Then ' 144 DPI
            If Not mTabData(nTab).Pic24 Is Nothing Then
                Set mTabData(nTab).PicToUse = mTabData(nTab).Pic24
            ElseIf Not mTabData(nTab).Pic20 Is Nothing Then
                Set mTabData(nTab).PicToUse = mTabData(nTab).Pic20
            Else
                Set mTabData(nTab).PicToUse = mTabData(nTab).Pic16
            End If
        ElseIf iTx >= 7 Then ' 192 DPI
            Set mTabData(nTab).PicToUse = StretchPicNN(mTabData(nTab).Pic16, 2)
        ElseIf iTx >= 6 Then
            If Not mTabData(nTab).Pic20 Is Nothing Then
                Set mTabData(nTab).PicToUse = StretchPicNN(mTabData(nTab).Pic20, 2)
            Else
                Set mTabData(nTab).PicToUse = StretchPicNN(mTabData(nTab).Pic16, 2)
            End If
        ElseIf iTx >= 5 Then
            If Not mTabData(nTab).Pic24 Is Nothing Then
                Set mTabData(nTab).PicToUse = StretchPicNN(mTabData(nTab).Pic24, 2)
            ElseIf Not mTabData(nTab).Pic20 Is Nothing Then
                Set mTabData(nTab).PicToUse = StretchPicNN(mTabData(nTab).Pic20, 2)
            Else
                Set mTabData(nTab).PicToUse = StretchPicNN(mTabData(nTab).Pic16, 3)
            End If
        ElseIf iTx >= 4 Then  ' 289 to 360 DPI
            If Not mTabData(nTab).Pic20 Is Nothing Then
                Set mTabData(nTab).PicToUse = StretchPicNN(mTabData(nTab).Pic20, 3)
            Else
                Set mTabData(nTab).PicToUse = StretchPicNN(mTabData(nTab).Pic16, 4)
            End If
        ElseIf iTx >= 3 Then   ' 361 to 480 DPI
            If Not mTabData(nTab).Pic24 Is Nothing Then
                Set mTabData(nTab).PicToUse = StretchPicNN(mTabData(nTab).Pic24, 3)
            ElseIf Not mTabData(nTab).Pic20 Is Nothing Then
                Set mTabData(nTab).PicToUse = StretchPicNN(mTabData(nTab).Pic20, 4)
            Else
                Set mTabData(nTab).PicToUse = StretchPicNN(mTabData(nTab).Pic16, 6)
            End If
        ElseIf iTx >= 2 Then   ' 481 to 720 DPI
            If Not mTabData(nTab).Pic24 Is Nothing Then
                Set mTabData(nTab).PicToUse = StretchPicNN(mTabData(nTab).Pic24, 5)
            Else
                Set mTabData(nTab).PicToUse = StretchPicNN(mTabData(nTab).Pic16, 8)
            End If
        Else ' greater than 720 DPI
            If Not mTabData(nTab).Pic24 Is Nothing Then
                Set mTabData(nTab).PicToUse = StretchPicNN(mTabData(nTab).Pic24, 10)
            Else
                Set mTabData(nTab).PicToUse = StretchPicNN(mTabData(nTab).Pic16, 16)
            End If
        End If
    Else
        If Not mTabData(nTab).Picture Is Nothing Then
            Set mTabData(nTab).PicToUse = mTabData(nTab).Picture
        Else
            Set mTabData(nTab).PicToUse = Nothing
        End If
    End If
    mTabData(nTab).PicToUseSet = True
End Sub

Private Function StretchPicNN(nPic As StdPicture, ByVal nFactor As Long) As StdPicture
    Dim iWidth As Long
    Dim iHeight As Long
    
    iWidth = pScaleX(nPic.Width, vbHimetric, vbPixels)
    iHeight = pScaleX(nPic.Height, vbHimetric, vbPixels)
    picAux.Width = iWidth * nFactor
    picAux.Height = iHeight * nFactor
    picAux.Cls
    
    picAux.PaintPicture nPic, 0, 0, picAux.Width, picAux.Height, 0, 0, iWidth, iHeight
    Set StretchPicNN = picAux.Image
    picAux.Cls
End Function

Private Function PictureToGrayScale(nPic As StdPicture) As StdPicture
    Dim iWidth As Long
    Dim iHeight As Long
    Dim X As Long
    Dim Y As Long
    Dim iColor As Long

    If nPic Is Nothing Then Exit Function
    
    iWidth = pScaleX(nPic.Width, vbHimetric, vbPixels)
    iHeight = pScaleX(nPic.Height, vbHimetric, vbPixels)
    picAux.Width = iWidth
    picAux.Height = iHeight
    picAux.Cls
    picAux2.Width = picAux.Width
    picAux2.Height = picAux.Height
    picAux2.Cls
    
    Set picAux.Picture = nPic

    For X = 0 To picAux.ScaleWidth - 1
        For Y = 0 To picAux.ScaleHeight - 1
            iColor = GetPixel(picAux.hDC, X, Y)
            If iColor <> mMaskColor Then
                iColor = ToGray(iColor)
            End If
            SetPixelV picAux2.hDC, X, Y, iColor
        Next Y
    Next X

    Set PictureToGrayScale = picAux2.Image
    picAux.Cls
    picAux2.Cls
End Function

Private Function ToGray(ByVal nColor As Long) As Long
    Dim iR As Long
    Dim iG As Long
    Dim iB As Long
    Dim iC As Long
    Dim iBlendDisablePicWithBackColorTabs As Boolean
    
    iR = nColor And 255
    iG = (nColor \ 256) And 255
    iB = (nColor \ 65536) And 255
    iC = (0.2125 * iR + 0.7154 * iG + 0.0721 * iB)
    
    If mControlIsThemed Then
        iBlendDisablePicWithBackColorTabs = mBlendDisablePicWithBackColorTabs_Themed
    Else
        iBlendDisablePicWithBackColorTabs = mBlendDisablePicWithBackColorTabs_NotThemed
    End If
        
    If iBlendDisablePicWithBackColorTabs Then
        If mControlIsThemed Then
            ToGray = RGB(iC / 255 * mThemedBodyBackColor_R * 0.7 + 88, iC / 255 * mThemedBodyBackColor_G * 0.7 + 88, iC / 255 * mThemedBodyBackColor_B * 0.7 + 88)
        Else
            ToGray = RGB(iC / 255 * mBackColorTabs_R * 0.7 + 88, iC / 255 * mBackColorTabs_G * 0.7 + 88, iC / 255 * mBackColorTabs_B * 0.7 + 88)
        End If
    Else
        ToGray = RGB(iC * 0.6 + 90, iC * 0.6 + 90, iC * 0.6 + 90)
    End If

End Function

Private Sub ResetAllPicsDisabled()
    Dim t As Long
    
    For t = 0 To mTabs - 1
        mTabData(t).PicDisabledSet = False
    Next t
End Sub

Private Function MouseIsOverAContainedControl() As Boolean
    Dim iPt As POINTAPI
    Dim iCtl As Object
    Dim iWidth As Long
    
    GetCursorPos iPt
    ScreenToClient mUserControlHwnd, iPt
    iPt.X = iPt.X * Screen_TwipsPerPixelX
    iPt.Y = iPt.Y * Screen_TwipsPerPixely
    
    On Error Resume Next
    For Each iCtl In UserControlContainedControls
        iWidth = -1
        iWidth = iCtl.Width
        If iWidth <> -1 Then
            If iCtl.Left <= iPt.X Then
                If iCtl.Left + iCtl.Width >= iPt.X Then
                    If iCtl.Top <= iPt.Y Then
                        If iCtl.Top + iCtl.Height >= iPt.Y Then
                            MouseIsOverAContainedControl = True
                            Err.Clear
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Next
    Err.Clear
End Function


Private Sub DrawDelayed()
    If mAmbientUserMode And mSubclassed Then
        PostDrawMessage
    Else
        Draw
    End If
End Sub

Private Sub PostDrawMessage()
    If mCanPostDrawMessage Then
        If Not mDrawMessagePosted Then
            PostMessage mUserControlHwnd, WM_DRAW, 0&, 0&
            mDrawMessagePosted = True
        End If
    Else
        tmrDraw.Enabled = True
    End If
End Sub

Friend Property Get TabControlsNames(ByVal Index As Long) As Object
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 381, TypeName(Me) ' invalid property array index
        Exit Property
    End If
    Set TabControlsNames = mTabData(Index).Controls
End Property

Friend Property Set TabControlsNames(ByVal Index As Long, nValue As Object)
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 381, TypeName(Me) ' invalid property array index
        Exit Property
    End If
    Set mTabData(Index).Controls = nValue
End Property

Friend Sub HideAllContainedControls()
    Dim iCtl As Object
    Dim c As Long
    Dim iHwnd As Long
    Dim iIsLine As Boolean
    
    If mUserControlTerminated Then Exit Sub
    
    If mSubclassedControlsForMoveHwnds.Count > 0 Then
        For c = 1 To mSubclassedControlsForMoveHwnds.Count
            iHwnd = mSubclassedControlsForMoveHwnds(c)
            DetachMessage Me, iHwnd, WM_WINDOWPOSCHANGING
        Next c
        Set mSubclassedControlsForMoveHwnds = New Collection
    End If
    
    On Error Resume Next
    For Each iCtl In UserControlContainedControls
        iIsLine = TypeName(iCtl) = "Line"
        If iIsLine Then
            If iCtl.X1 > -mLeftThresholdHided Then
                iCtl.X1 = iCtl.X1 - mLeftOffsetToHide
                iCtl.X2 = iCtl.X2 - mLeftOffsetToHide
            End If
        Else
            If iCtl.Left > -mLeftThresholdHided Then
                iCtl.Left = iCtl.Left - mLeftOffsetToHide
            End If
        End If
    Next
    Err.Clear
End Sub

Friend Sub MakeContainedControlsInSelTabVisible()
    Dim iCtl As Object
    Dim iCtlName As Variant
    Dim iHwnd As Long
    Dim c As Long
    Dim iIsLine As Boolean
    
    If mUserControlTerminated Then Exit Sub
    
    If mSubclassedControlsForMoveHwnds.Count > 0 Then
        For c = 1 To mSubclassedControlsForMoveHwnds.Count
            iHwnd = mSubclassedControlsForMoveHwnds(c)
            DetachMessage Me, iHwnd, WM_WINDOWPOSCHANGING
        Next c
        Set mSubclassedControlsForMoveHwnds = New Collection
    End If
    
    On Error Resume Next
    For Each iCtlName In mTabData(mTab).Controls
        Set iCtl = GetContainedControlByName(iCtlName)
        If Not iCtl Is Nothing Then
            iIsLine = TypeName(iCtl) = "Line"
            If iIsLine Then
                If iCtl.X1 < -mLeftThresholdHided Then
                    iCtl.X1 = iCtl.X1 + mLeftOffsetToHide
                    iCtl.X2 = iCtl.X2 + mLeftOffsetToHide
                End If
            Else
                If iCtl.Left < -mLeftThresholdHided Then
                    iCtl.Left = iCtl.Left + mLeftOffsetToHide
                End If
            End If
            If mAmbientUserMode And mSubclassed And (Not mOnlySubclassUserControl) Then
                iHwnd = 0
                iHwnd = GetControlHwnd(iCtl)
                If iHwnd <> 0 Then
                    mSubclassedControlsForMoveHwnds.Add iHwnd
                    AttachMessage Me, iHwnd, WM_WINDOWPOSCHANGING
                End If
            End If
        End If
    Next
    Err.Clear
End Sub

Private Sub CheckContainedControlsConsistency(Optional nCheckControlsThatChangedToArray As Boolean)
    Dim t As Long
    Dim iCCList As Collection
    Dim iAllCtInTabs As Collection
    Dim c As Long
    Dim iStr As String
    Dim iCtl As Object
    Dim iCtlName As Variant
    Dim iCtlName2 As Variant
    Dim iCtlsInTabsToRemove As Collection
    Dim iShowedNewControls As Boolean
    Dim iThereAreMissingControls As Boolean
    Dim iAuxFound As Boolean
    Dim iCtlsTypesAndRects As Collection
    Dim iAuxTypeAndRect_CtrlInTab As String
    Dim iAuxTypeAndRect_CC As String
    Dim iFound As Boolean
    Dim t2 As Long
    Dim c2 As Long
    Dim iListCtlsNowArrayToUpdateInfo As Collection
    Dim iIsLine As Boolean
    
    Set iCCList = New Collection
    For Each iCtl In UserControlContainedControls
        iStr = ControlName(iCtl)
        iCCList.Add iStr, iStr
    Next
    
    On Error Resume Next
    Set iAllCtInTabs = New Collection
    For t = 0 To mTabs - 1
        For c = 1 To mTabData(t).Controls.Count
            iStr = mTabData(t).Controls(c)
            iAllCtInTabs.Add iStr, iStr
        Next c
    Next t
    On Error GoTo 0
    
    iThereAreMissingControls = False
    For Each iCtlName In iAllCtInTabs
        iAuxFound = False
        For Each iCtlName2 In iCCList
            If iCtlName2 = iCtlName Then
                iAuxFound = True
                Exit For
            End If
        Next
        If Not iAuxFound Then
            iThereAreMissingControls = True
            If nCheckControlsThatChangedToArray Then
                iAuxFound = False
                For Each iCtlName2 In iCCList
                    If iCtlName2 = iCtlName & "(0)" Then
                        iAuxFound = True
                        Exit For
                    End If
                Next
                If iAuxFound Then
                    If iListCtlsNowArrayToUpdateInfo Is Nothing Then Set iListCtlsNowArrayToUpdateInfo = New Collection
                    iListCtlsNowArrayToUpdateInfo.Add iCtlName, iCtlName
                End If
            Else
                Exit For
            End If
        End If
    Next
    
    If iThereAreMissingControls Then
        If nCheckControlsThatChangedToArray Then
            If Not iListCtlsNowArrayToUpdateInfo Is Nothing Then
                For t = 0 To mTabs - 1
                    For c = 1 To mTabData(t).Controls.Count
                        iStr = mTabData(t).Controls(c)
                        iFound = False
                        For Each iCtlName In iListCtlsNowArrayToUpdateInfo
                            If iCtlName = iStr Then
                                iFound = True
                            End If
                        Next
                        If iFound Then
                            iStr = iStr & "(0)"
                            mTabData(t).Controls.Add iStr, iStr, c
                            mTabData(t).Controls.Remove (c + 1)
                        End If
                    Next c
                Next t
            End If
        Else
            ' This fixes SSTab paste bug, read http://www.vbforums.com/showthread.php?871285&p=5359379&viewfull=1#post5359379
            Set iCtlsTypesAndRects = New Collection
            For Each iCtl In UserControlContainedControls
                iStr = ControlName(iCtl)
                iCtlsTypesAndRects.Add GetControlTypeAndRect(iStr), iStr
            Next
            
            For t = 0 To mTabs - 1 ' enumerate tabs
                For c = 1 To mTabData(t).Controls.Count ' enumerate controls that are in that tab
                    iStr = mTabData(t).Controls(c) ' in iStr: get the name of one control in the "current" tab
                    iAuxTypeAndRect_CtrlInTab = GetControlTypeAndRect(iStr)
                    If iAuxTypeAndRect_CtrlInTab = "-" Then ' if the control is not found it may have been en converted to an array
                        iAuxTypeAndRect_CtrlInTab = GetControlTypeAndRect(iStr & "(0)")
                    End If
                    For Each iCtlName In iCCList ' iCCList has al the Contained Controls that are in the UserControl (inside the NewTab)
                        iAuxTypeAndRect_CC = GetControlTypeAndRect(CStr(iCtlName))
                        If iAuxTypeAndRect_CC = iAuxTypeAndRect_CtrlInTab Then
                            iFound = False
                            For t2 = 0 To mTabs - 1
                                For c2 = 1 To mTabData(t).Controls.Count
                                    If mTabData(t).Controls(c2) = iCtlName Then
                                        iFound = True
                                    End If
                                Next c2
                            Next t2
                            If Not iFound Then
                                mTabData(t).Controls.Add iCtlName, iCtlName, c
                                mTabData(t).Controls.Remove (c + 1)
                            End If
                        End If
                    Next
                Next c
            Next t
            
            On Error Resume Next
            Set iAllCtInTabs = New Collection
            For t = 0 To mTabs - 1
                For c = 1 To mTabData(t).Controls.Count
                    iStr = mTabData(t).Controls(c)
                    iAllCtInTabs.Add iStr, iStr
                Next c
            Next t
            On Error GoTo 0
        End If
    End If
    
    If nCheckControlsThatChangedToArray Then Exit Sub
    
    ' check if contained control is on any tab
    iShowedNewControls = False
    On Error Resume Next
    For Each iCtlName In iCCList
        iStr = ""
        iStr = iAllCtInTabs(iCtlName)
        If iStr = "" Then ' the control is not placed on any tab
            ' place it in the visible tab
            mTabData(mTab).Controls.Add iCtlName, iCtlName
            Set iCtl = GetContainedControlByName(iCtlName)
            iIsLine = TypeName(iCtl) = "Line"
            If iIsLine Then
                If iCtl.X1 <= -mLeftThresholdHided Then
                    iCtl.X1 = iCtl.X1 + mLeftOffsetToHide
                    iCtl.X2 = iCtl.X2 + mLeftOffsetToHide
                    iShowedNewControls = True
                End If
            Else
                If iCtl.Left <= -mLeftThresholdHided Then
                    iCtl.Left = iCtl.Left + mLeftOffsetToHide
                    iShowedNewControls = True
                End If
            End If
        End If
    Next
    
    If iShowedNewControls Then
        mSubclassControlsPaintingPending = True
        mRepaintSubclassedControls = True
        SubclassControlsPainting
    End If
    
    ' now check the inverse: if there are controls in tabs but they don't exists
    Set iCtlsInTabsToRemove = New Collection
    For Each iCtlName In iAllCtInTabs
        iStr = ""
        iStr = iCCList(iCtlName)
        If iStr = "" Then ' the control doesn't exist
            iCtlsInTabsToRemove.Add iStr, iStr
        End If
    Next
    
    ' remove the controls that don't exists, if any
    If iCtlsInTabsToRemove.Count > 0 Then
        For t = 0 To mTabs - 1
            For Each iCtlName In mTabData(t).Controls
                iStr = ""
                iStr = iCtlsInTabsToRemove(iCtlName)
                If iStr <> "" Then ' the control name is in the list of controls to remove
                    mTabData(t).Controls.Remove iCtlName
                End If
            Next
        Next t
    End If
    Err.Clear
End Sub

Private Sub CheckIfContainedControlChangedToArray()
    CheckContainedControlsConsistency True
End Sub

Private Function GetControlTypeAndRect(iCtlName As String) As String
    Dim iCtl As Object
    Dim iSng As Long
    
    Set iCtl = GetParentControlByName(iCtlName)
    If Not iCtl Is Nothing Then
        On Error Resume Next
        GetControlTypeAndRect = TypeName(iCtl) & "."
        iSng = 0
        iSng = iCtl.Left
        GetControlTypeAndRect = GetControlTypeAndRect & CStr(iSng) & "."
        iSng = 0
        iSng = iCtl.Top
        GetControlTypeAndRect = GetControlTypeAndRect & CStr(iSng) & "."
        iSng = 0
        iSng = iCtl.Width
        GetControlTypeAndRect = GetControlTypeAndRect & CStr(iSng) & "."
        iSng = 0
        iSng = iCtl.Height
        GetControlTypeAndRect = GetControlTypeAndRect & CStr(iSng)
    Else
        GetControlTypeAndRect = "-"
    End If
End Function

Private Function GetParentControlByName(ByVal nControlName As String) As Object
    Dim iCtl As Object
    
    On Error GoTo ErrorExit
    For Each iCtl In UserControl.Parent.Controls
        If StrComp(nControlName, ControlName(iCtl), vbTextCompare) = 0 Then
            Set GetParentControlByName = iCtl
            Exit For
        End If
    Next
    
ErrorExit:
End Function

Public Property Get Controls() As VBRUN.ContainedControls
Attribute Controls.VB_Description = "Returns a collection of the controls that were added to the control."
    Set Controls = UserControlContainedControls
End Property

Private Sub RaiseError(ByVal Number As Long, Optional ByVal Source As Variant, Optional ByVal Description As Variant, Optional ByVal HelpFile As Variant, Optional ByVal HelpContext As Variant)
    If mInIDE Then
        On Error Resume Next
        Err.Raise Number, Source, Description, HelpFile, HelpContext
        MsgBox "Error " & Err.Number & ". " & Err.Description, vbCritical
    Else
        Err.Raise Number, Source, Description, HelpFile, HelpContext
    End If
End Sub

Private Function MakeTrue(Value As Boolean) As Boolean
    MakeTrue = True
    Value = True
End Function

Private Function ControlHasFocus() As Boolean
    ControlHasFocus = mHasFocus And mFormIsActive
End Function

Private Sub RearrangeContainedControlsPositions()
    Dim iCtl As Object
    Dim iBodyStart As Single
    Dim iBodyStart_Prev As Single
    Dim iIsLine As Boolean
    
    If (mTabOrientation = ssTabOrientationTop) Or (mTabOrientation = ssTabOrientationBottom) Then
        iBodyStart = pScaleY(mClientStart - 5, vbPixels, vbTwips)
    Else
        iBodyStart = pScaleX(mClientStart - 5, vbPixels, vbTwips)
    End If
    If (mTabOrientation_Prev = ssTabOrientationTop) Or (mTabOrientation_Prev = ssTabOrientationBottom) Then
        iBodyStart_Prev = pScaleY(mClientStart_Prev - 5, vbPixels, vbTwips)
    Else
        iBodyStart_Prev = pScaleX(mClientStart_Prev - 5, vbPixels, vbTwips)
    End If
    
    On Error Resume Next
    If mTabOrientation = mTabOrientation_Prev Then
        For Each iCtl In UserControlContainedControls
            iIsLine = TypeName(iCtl) = "Line"
            If mTabOrientation = ssTabOrientationTop Then
                If iIsLine Then
                    iCtl.Y1 = iCtl.Y1 - iBodyStart_Prev + iBodyStart
                    iCtl.Y2 = iCtl.Y2 - iBodyStart_Prev + iBodyStart
                Else
                    iCtl.Top = iCtl.Top - iBodyStart_Prev + iBodyStart
                End If
            ElseIf mTabOrientation = ssTabOrientationBottom Then
                If iIsLine Then
                    iCtl.Y1 = iCtl.Y1 + iBodyStart_Prev - iBodyStart
                    iCtl.Y2 = iCtl.Y2 + iBodyStart_Prev - iBodyStart
                Else
                    iCtl.Top = iCtl.Top + iBodyStart_Prev - iBodyStart
                End If
            ElseIf (mTabOrientation = ssTabOrientationLeft) Or (mTabOrientation = ntTabOrientationLeftHorizontal) Then
                If iIsLine Then
                    iCtl.X1 = iCtl.X1 - iBodyStart_Prev + iBodyStart
                    iCtl.X2 = iCtl.X2 - iBodyStart_Prev + iBodyStart
                Else
                    iCtl.Left = iCtl.Left - iBodyStart_Prev + iBodyStart
                End If
            ElseIf (mTabOrientation = ssTabOrientationRight) Or (mTabOrientation = ntTabOrientationRightHorizontal) Then
                If iIsLine Then
                    iCtl.X1 = iCtl.X1 + iBodyStart_Prev - iBodyStart
                    iCtl.X2 = iCtl.X2 + iBodyStart_Prev - iBodyStart
                Else
                    iCtl.Left = iCtl.Left + iBodyStart_Prev - iBodyStart
                End If
            End If
        Next
    Else
        For Each iCtl In UserControlContainedControls
            iIsLine = TypeName(iCtl) = "Line"
            If mTabOrientation_Prev = ssTabOrientationTop Then
                If iIsLine Then
                    iCtl.Y1 = iCtl.Y1 - iBodyStart_Prev
                    iCtl.Y2 = iCtl.Y2 - iBodyStart_Prev
                Else
                    iCtl.Top = iCtl.Top - iBodyStart_Prev
                End If
            ElseIf mTabOrientation_Prev = ssTabOrientationBottom Then
                '
            ElseIf (mTabOrientation_Prev = ssTabOrientationLeft) Or (mTabOrientation_Prev = ntTabOrientationLeftHorizontal) Then
                If iIsLine Then
                    iCtl.X1 = iCtl.X1 - iBodyStart_Prev
                    iCtl.X2 = iCtl.X2 - iBodyStart_Prev
                Else
                    iCtl.Left = iCtl.Left - iBodyStart_Prev
                End If
            ElseIf (mTabOrientation_Prev = ssTabOrientationRight) Or (mTabOrientation_Prev = ntTabOrientationRightHorizontal) Then
                '
            End If
        
            If mTabOrientation = ssTabOrientationTop Then
                If iIsLine Then
                    iCtl.Y1 = iCtl.Y1 + iBodyStart
                    iCtl.Y2 = iCtl.Y2 + iBodyStart
                Else
                    iCtl.Top = iCtl.Top + iBodyStart
                End If
            ElseIf mTabOrientation = ssTabOrientationBottom Then
                '
            ElseIf (mTabOrientation = ssTabOrientationLeft) Or (mTabOrientation = ntTabOrientationLeftHorizontal) Then
                If iIsLine Then
                    iCtl.X1 = iCtl.X1 + iBodyStart
                    iCtl.X2 = iCtl.X2 + iBodyStart
                Else
                    iCtl.Left = iCtl.Left + iBodyStart
                End If
            ElseIf (mTabOrientation = ssTabOrientationRight) Or (mTabOrientation = ntTabOrientationRightHorizontal) Then
                '
            End If
        Next
    End If
    Err.Clear
End Sub

Public Property Get TabControls(ByVal nTab As Long, Optional GetChilds As Boolean = True) As Collection
Attribute TabControls.VB_Description = "Returns a collection of the controls that are inside a tab."
    Dim iCtlName As Variant
    Dim iCtl As Object
    Dim iCtl2 As Object
    Dim iObj As Object
    
    If (nTab < 0) Or (nTab > (mTabs - 1)) Then
        RaiseError 5, TypeName(Me) ' Invalid procedure call or argument
        Exit Property
    End If
    
    Set TabControls = New Collection
    
    If GetChilds Then
        If Not mTabStopsInitialized Then
            StoreControlsTabStop True
            mTabStopsInitialized = True
        End If
    End If
    
    For Each iCtlName In mTabData(nTab).Controls
        Set iCtl = GetContainedControlByName(iCtlName)
        If Not iCtl Is Nothing Then
            Set iObj = iCtl
            TabControls.Add iObj, iCtlName
            If GetChilds Then
                If ControlIsContainer(iCtlName) Then
                    For Each iCtl2 In GetContainedControlsInControlContainer(iCtl)
                        Set iObj = iCtl2
                        TabControls.Add iObj, iCtl2.Name
                    Next
                End If
            End If
        End If
    Next
    
End Property

Public Property Get EndOfTabs() As Single
Attribute EndOfTabs.VB_Description = "Returns a value that indicates where the last tab ends."
    EnsureDrawn
    If (mTabOrientation = ssTabOrientationTop) Or (mTabOrientation = ssTabOrientationBottom) Then
        EndOfTabs = FixRoundingError(ToContainerSizeX(mEndOfTabs, vbPixels))
    Else
        EndOfTabs = FixRoundingError(ToContainerSizeY(mEndOfTabs, vbPixels))
    End If
End Property

Public Property Get MinSizeNeeded() As Single
Attribute MinSizeNeeded.VB_Description = "Returns the minimun Width (or Height, depending on the TabOpientation setting) of the control needed to show all the tabs in one row (without adding new rows)."
Attribute MinSizeNeeded.VB_ProcData.VB_Invoke_Property = ";Data"
    EnsureDrawn
    If (mTabOrientation = ssTabOrientationTop) Or (mTabOrientation = ssTabOrientationBottom) Then
        MinSizeNeeded = FixRoundingError(ToContainerSizeX(mMinSizeNeeded, vbPixels))
    Else
        MinSizeNeeded = FixRoundingError(ToContainerSizeY(mMinSizeNeeded, vbPixels))
    End If
End Property

Public Property Get HandleHighContrastTheme() As Boolean
Attribute HandleHighContrastTheme.VB_Description = "When True (default setting), handles the system changes to High contrast theme automatically."
Attribute HandleHighContrastTheme.VB_ProcData.VB_Invoke_Property = ";Behavior"
    HandleHighContrastTheme = mHandleHighContrastTheme
End Property

Public Property Let HandleHighContrastTheme(ByVal nValue As Boolean)
    If nValue <> mHandleHighContrastTheme Then
        mHandleHighContrastTheme = nValue
        SetPropertyChanged "HandleHighContrastTheme"
        If mHandleHighContrastTheme Then
            CheckHighContrastTheme
        End If
        SetColors
        DrawDelayed
    End If
End Property


Friend Function pScaleX(Width As Variant, Optional ByVal FromScale As Variant, Optional ByVal ToScale As Variant) As Single
    Select Case True
        Case ToScale = vbPixels
            Select Case FromScale
                Case vbCentimeters
                    pScaleX = Width * mDPIX / 2.54
                Case vbCharacters
                    pScaleX = Width / 1440 * mDPIX * 120
                Case vbHimetric
                    pScaleX = Width * mDPIX / 2540
                Case vbInches
                    pScaleX = Width * mDPIX
                Case vbMillimeters
                    pScaleX = Width * mDPIX / 25.4
                Case vbPixels
                    pScaleX = Width
                Case vbPoints
                    pScaleX = Width / 1440 * mDPIX * 20
                Case vbTwips
                    pScaleX = Width / 1440 * mDPIX
                Case Else
                    pScaleX = UserControl.ScaleX(Width, FromScale, ToScale)
            End Select
        Case FromScale = vbPixels
            Select Case ToScale
                Case vbCentimeters
                    pScaleX = Width / mDPIX * 2.54
                Case vbCharacters
                    pScaleX = Width * 1440 / mDPIX / 120
                Case vbHimetric
                    pScaleX = Width / mDPIX * 2540
                Case vbInches
                    pScaleX = Width / mDPIX
                Case vbMillimeters
                    pScaleX = Width / mDPIX * 25.4
                Case vbPixels
                    pScaleX = Width
                Case vbPoints
                    pScaleX = Width * 1440 / mDPIX / 20
                Case vbTwips
                    pScaleX = Width * 1440 / mDPIX
                Case vbUser
                    pScaleX = UserControl.ScaleX(Width, FromScale, ToScale)
                Case Else
                    pScaleX = UserControl.ScaleX(Width, FromScale, ToScale)
            End Select
        Case Else
            pScaleX = UserControl.ScaleX(Width, FromScale, ToScale)
    End Select
End Function

Friend Function pScaleY(Height As Variant, Optional ByVal FromScale As Variant, Optional ByVal ToScale As Variant) As Single
    Select Case True
        Case ToScale = vbPixels
            Select Case FromScale
                Case vbCentimeters
                    pScaleY = Height * mDPIY / 2.54
                Case vbCharacters
                    pScaleY = Height / 1440 * mDPIY * 120
                Case vbHimetric
                    pScaleY = Height * mDPIY / 2540
                Case vbInches
                    pScaleY = Height * mDPIY
                Case vbMillimeters
                    pScaleY = Height * mDPIY / 25.4
                Case vbPixels
                    pScaleY = Height
                Case vbPoints
                    pScaleY = Height / 1440 * mDPIY * 20
                Case vbTwips
                    pScaleY = Height / 1440 * mDPIY
                Case Else
                    pScaleY = UserControl.ScaleY(Height, FromScale, ToScale)
            End Select
        Case FromScale = vbPixels
            Select Case ToScale
                Case vbCentimeters
                    pScaleY = Height / mDPIY * 2.54
                Case vbCharacters
                    pScaleY = Height * 1440 / mDPIY / 120
                Case vbHimetric
                    pScaleY = Height / mDPIY * 2540
                Case vbInches
                    pScaleY = Height / mDPIY
                Case vbMillimeters
                    pScaleY = Height / mDPIY * 25.4
                Case vbPixels
                    pScaleY = Height
                Case vbPoints
                    pScaleY = Height * 1440 / mDPIY / 20
                Case vbTwips
                    pScaleY = Height * 1440 / mDPIY
                Case vbUser
                    pScaleY = UserControl.ScaleY(Height, FromScale, ToScale)
                Case Else
                    pScaleY = UserControl.ScaleY(Height, FromScale, ToScale)
            End Select
        Case Else
            pScaleY = UserControl.ScaleY(Height, FromScale, ToScale)
    End Select
End Function

Private Sub SetDPI()
    Dim iDC As Long
    Dim iTx As Single
    Dim iTY As Single
    
    iDC = GetDC(0)
    mDPIX = GetDeviceCaps(iDC, LOGPIXELSX)
    mDPIY = GetDeviceCaps(iDC, LOGPIXELSY)
    ReleaseDC 0, iDC
    
    iTx = 1440 / mDPIX
    iTY = 1440 / mDPIY
    
    mXCorrection = iTx / Screen.TwipsPerPixelX
    mYCorrection = iTY / Screen.TwipsPerPixelY
    
    SetLeftOffsetToHide Screen.TwipsPerPixelX
    mDPIScale = 1 / 96 * mDPIX
End Sub

Private Sub SetLeftOffsetToHide(ByVal nTwipsPerPixel As Long)
    If nTwipsPerPixel >= 6 Then
        mLeftOffsetToHide = 75000 ' compatible with original SSTab up to 250% DPI
        mLeftThresholdHided = 15000
    Else
        mLeftOffsetToHide = nTwipsPerPixel * 16384 * 0.6 ' Windows has a limit on controls positions out of screen (in pixels), need to handle that for very hight DPI setting (> 250%) https://www.vbforums.com/showthread.php?888201
        If mLeftOffsetToHide > 30000 Then
            mLeftThresholdHided = 15000
        Else
            mLeftThresholdHided = mLeftOffsetToHide / 2
        End If
    End If
End Sub

Private Function Screen_TwipsPerPixelX() As Single
    Screen_TwipsPerPixelX = Screen.TwipsPerPixelX * mXCorrection
End Function

Private Function Screen_TwipsPerPixely() As Single
    Screen_TwipsPerPixely = Screen.TwipsPerPixelY * mYCorrection
End Function


Public Property Get Object() As Object
Attribute Object.VB_Description = "Returns the control instance without the extender."
    Set Object = Me
End Property

Private Function IsMsgBoxShown() As Boolean
    Dim iHwnd As Long
     
    Do Until IsWindowLocal(iHwnd)
        iHwnd = FindWindowEx(0&, iHwnd, "#32770", vbNullString)
        If iHwnd = 0 Then Exit Function
    Loop
    IsMsgBoxShown = True
End Function

Private Function IsWindowLocal(ByVal nHwnd As Long) As Boolean
    Dim iIdProcess As Long
    
    Call GetWindowThreadProcessId(nHwnd, iIdProcess)
    IsWindowLocal = (iIdProcess = GetCurrentProcessId())
End Function

Private Function IsHighContrastTheme() As Boolean
    Dim iHC As tagHIGHCONTRAST
    
    iHC.cbSize = Len(iHC)
    SystemParametersInfo SPI_GETHIGHCONTRAST, Len(iHC), iHC, 0
    IsHighContrastTheme = (iHC.dwFlags And HCF_HIGHCONTRASTON) = HCF_HIGHCONTRASTON
End Function

Private Sub CheckHighContrastTheme()
    Dim iAuxBool As Boolean
    
'    If Not mAmbientUserMode Then Exit Sub
    If mHighContrastThemeOn <> IsHighContrastTheme Then
        iAuxBool = Not mHighContrastThemeOn
        If iAuxBool Then
            mHandleHighContrastTheme_OrigForeColor = ForeColor
            mHandleHighContrastTheme_OrigBackColorTabs = BackColorTabs
            mHandleHighContrastTheme_OrigForeColorSelectedTab = ForeColorSelectedTab
            mHandleHighContrastTheme_OrigForeColorHighlighted = ForeColorHighlighted
            mHandleHighContrastTheme_OrigFlatTabBorderColorHighlight = FlatTabBorderColorHighlight
            mHandleHighContrastTheme_OrigFlatTabBorderColorSelectedTab = FlatTabBorderColorSelectedTab
            mHandleHighContrastTheme_OrigBackColorSelectedTab = BackColorSelectedTab
            mHandleHighContrastTheme_OrigIconColor = IconColor
            mHandleHighContrastTheme_OrigIconColorSelectedTab = IconColorSelectedTab
            mHandleHighContrastTheme_OrigIconColorMouseHover = IconColorMouseHover
            mHandleHighContrastTheme_OrigIconColorMouseHoverSelectedTab = IconColorMouseHoverSelectedTab
            mHandleHighContrastTheme_OrigIconColorTabHighlighted = IconColorTabHighlighted
            ForeColor = vbButtonText
            BackColorTabs = vbButtonFace
            ForeColorSelectedTab = vbButtonText
            ForeColorHighlighted = vbButtonText
            FlatTabBorderColorHighlight = vbButtonText
            FlatTabBorderColorSelectedTab = vbButtonText
            mChangingHighContrastTheme = True
            BackColorSelectedTab = vbButtonFace
            mChangingHighContrastTheme = False
            IconColorSelectedTab = vbButtonText
            IconColorMouseHover = vbButtonText
            IconColorMouseHoverSelectedTab = vbButtonText
            IconColorTabHighlighted = vbButtonText
            mHighContrastThemeOn = True
        Else
            mHighContrastThemeOn = False
            ForeColor = mHandleHighContrastTheme_OrigForeColor
            BackColorTabs = mHandleHighContrastTheme_OrigBackColorTabs
            ForeColorSelectedTab = mHandleHighContrastTheme_OrigForeColorSelectedTab
            ForeColorHighlighted = mHandleHighContrastTheme_OrigForeColorHighlighted
            FlatTabBorderColorHighlight = mHandleHighContrastTheme_OrigFlatTabBorderColorHighlight
            FlatTabBorderColorSelectedTab = mHandleHighContrastTheme_OrigFlatTabBorderColorSelectedTab
            If mBackColorSelectedTab_IsAutomatic Then
                BackColorSelectedTab = BackColorTabs
            Else
                BackColorSelectedTab = mHandleHighContrastTheme_OrigBackColorSelectedTab
            End If
            IconColor = mHandleHighContrastTheme_OrigIconColor
            IconColorSelectedTab = mHandleHighContrastTheme_OrigIconColorSelectedTab
            IconColorMouseHover = mHandleHighContrastTheme_OrigIconColorMouseHover
            IconColorMouseHoverSelectedTab = mHandleHighContrastTheme_OrigIconColorMouseHoverSelectedTab
            IconColorTabHighlighted = mHandleHighContrastTheme_OrigIconColorTabHighlighted
        End If
    End If
End Sub

Public Property Get LeftOffsetToHide() As Long
Attribute LeftOffsetToHide.VB_Description = "Returns the shift to the left in twips that it is using to hide the controls in not active tabs."
Attribute LeftOffsetToHide.VB_ProcData.VB_Invoke_Property = ";Data"
    LeftOffsetToHide = mLeftOffsetToHide
End Property


Public Property Get ControlLeft(ByVal ControlName As String) As Single
Attribute ControlLeft.VB_Description = "Returns/sets the left of the contained control whose name was provided by the ControlName parameter."
Attribute ControlLeft.VB_ProcData.VB_Invoke_Property = ";Position"
    Dim iCtl As Object
    Dim iFound As Boolean
    Dim iWithIndex As Boolean
    Dim iName As String
    Dim iIndex As Long
    Dim iLeft As Single
    Dim iIsLine As Boolean
    
    ControlName = LCase$(ControlName)
    iWithIndex = InStr(ControlName, "(") > 0
    For Each iCtl In UserControlContainedControls
        iName = LCase$(iCtl.Name)
        If iWithIndex Then
            iIndex = -1
            On Error Resume Next
            iIndex = iCtl.Index
            On Error GoTo 0
            If iIndex <> -1 Then
                iName = iName & "(" & iIndex & ")"
            End If
        End If
        If iName = ControlName Then
            iFound = True
            Exit For
        End If
    Next
    If Not iFound Then
        RaiseError 1501, , "Control not found."
    Else
        iIsLine = TypeName(iCtl) = "Line"
        iLeft = 0
        On Error Resume Next
        If iIsLine Then
            iLeft = iCtl.X1
        Else
            iLeft = iCtl.Left
        End If
        On Error GoTo 0
        If iLeft < -mLeftThresholdHided Then
           ControlLeft = iLeft + mLeftOffsetToHide + mPendingLeftOffset
        Else
            ControlLeft = iLeft
        End If
    End If
End Property

Public Property Let ControlLeft(ByVal ControlName As String, ByVal Left As Single)
    Dim iCtl As Object
    Dim iFound As Boolean
    Dim iWithIndex As Boolean
    Dim iName As String
    Dim iIndex As Long
    
    Left = Left - mPendingLeftOffset
    
    ControlName = LCase$(ControlName)
    iWithIndex = InStr(ControlName, "(") > 0
    For Each iCtl In UserControlContainedControls
        iName = LCase$(iCtl.Name)
        If iWithIndex Then
            iIndex = -1
            On Error Resume Next
            iIndex = iCtl.Index
            On Error GoTo 0
            If iIndex <> -1 Then
                iName = iName & "(" & iIndex & ")"
            End If
        End If
        If iName = ControlName Then
            iFound = True
            Exit For
        End If
    Next
    If Not iFound Then
        RaiseError 1501, , "Control not found."
    Else
        If iCtl.Left < -mLeftThresholdHided Then
            iCtl.Left = Left - mLeftOffsetToHide
        Else
            iCtl.Left = Left
        End If
    End If
End Property

Public Sub ControlMove(ByVal nControlName As String, ByVal Left As Single, ByVal Top As Single, Optional ByVal Width As Variant, Optional ByVal Height As Variant, Optional ByVal IndexOfOtherTabToMoveTheControl As Long = -1)
Attribute ControlMove.VB_Description = "Replaces the ControlName.Move method. The difference is that it takes into account the Left offset of controls on inactive tabs."
    Dim iCtl As Object
    Dim iFound As Boolean
    Dim iWithIndex As Boolean
    Dim iName As String
    Dim iIndex As Long
    Dim t As Long
    Dim iCtlName As String
    Dim c As Long
    Dim iIsLine As Boolean
    Dim iAuxLeft As Single
    
    Left = Left - mPendingLeftOffset
    
    nControlName = LCase$(nControlName)
    iWithIndex = InStr(nControlName, "(") > 0
    For Each iCtl In UserControlContainedControls
        iName = LCase$(iCtl.Name)
        If iWithIndex Then
            iIndex = -1
            On Error Resume Next
            iIndex = iCtl.Index
            On Error GoTo 0
            If iIndex <> -1 Then
                iName = iName & "(" & iIndex & ")"
            End If
        End If
        If iName = nControlName Then
            iFound = True
            Exit For
        End If
    Next
    If Not iFound Then
        RaiseError 1501, , "Control not found."
    Else
        iAuxLeft = 0
        iIsLine = False
        If TypeName(iCtl) = "Line" Then
            iAuxLeft = iCtl.X1
            iIsLine = True
        Else
            iAuxLeft = iCtl.Left
        End If
        If iIsLine Then
            If IsMissing(Width) Then
                Width = Abs(iCtl.X2 - iCtl.X1)
            End If
            If IsMissing(Height) Then
                Height = Abs(iCtl.Y2 - iCtl.Y1)
            End If
            If iAuxLeft < -mLeftThresholdHided Then
                iCtl.X1 = Left - mLeftOffsetToHide
            Else
                iCtl.X1 = Left
            End If
            iCtl.X2 = iCtl.X1 + Width
            iCtl.Y1 = Top
            iCtl.Y2 = iCtl.Y1 + Height
            iAuxLeft = iCtl.X1
        Else
            If IsMissing(Width) Then
                Width = iCtl.Width
            End If
            If IsMissing(Height) Then
                Height = iCtl.Height
            End If
            If iAuxLeft < -mLeftThresholdHided Then
                iCtl.Move Left - mLeftOffsetToHide, Top, Width, Height
            Else
                iCtl.Move Left, Top, Width, Height
            End If
            iAuxLeft = iCtl.Left
        End If
        If IndexOfOtherTabToMoveTheControl > -1 Then
            iCtlName = ControlName(iCtl)
            iFound = False
            For t = 0 To mTabs - 1
                For c = 1 To mTabData(t).Controls.Count
                    If mTabData(t).Controls(c) = iCtlName Then
                        If t <> IndexOfOtherTabToMoveTheControl Then
                            mTabData(t).Controls.Remove iCtlName
                        Else
                            iFound = True
                            Exit For
                        End If
                    End If
                Next
                If iFound Then Exit For
            Next
            mTabData(IndexOfOtherTabToMoveTheControl).Controls.Add iCtlName, iCtlName
            If (iAuxLeft < -mLeftThresholdHided) And (IndexOfOtherTabToMoveTheControl = mTab) Then
                If iIsLine Then
                    iCtl.X1 = iCtl.X1 + mLeftOffsetToHide
                    iCtl.X2 = iCtl.X2 + mLeftOffsetToHide
                Else
                    iCtl.Left = iCtl.Left + mLeftOffsetToHide
                End If
            ElseIf (iAuxLeft >= -mLeftThresholdHided) And (IndexOfOtherTabToMoveTheControl <> mTab) Then
                If iIsLine Then
                    iCtl.X1 = iCtl.X1 - mLeftOffsetToHide
                    iCtl.X2 = iCtl.X2 - mLeftOffsetToHide
                Else
                    iCtl.Left = iCtl.Left - mLeftOffsetToHide
                End If
            End If
        End If
    End If
End Sub

Public Sub ControlSetTab(ByVal nControlName As String, ByVal nTab As Long)
Attribute ControlSetTab.VB_Description = "Sets or change the tab where a contained control is."
    Dim iControlName As String
    Dim iCtl As Object
    Dim iWithIndex As Boolean
    Dim iName As String
    Dim iIndex As Long
    Dim iFound As Boolean
    
    iControlName = LCase$(nControlName)
    iWithIndex = InStr(iControlName, "(") > 0
    For Each iCtl In UserControlContainedControls
        iName = LCase$(iCtl.Name)
        If iWithIndex Then
            iIndex = -1
            On Error Resume Next
            iIndex = iCtl.Index
            On Error GoTo 0
            If iIndex <> -1 Then
                iName = iName & "(" & iIndex & ")"
            End If
        End If
        If iName = iControlName Then
            iFound = True
            Exit For
        End If
    Next
    If Not iFound Then
        RaiseError 1501, , "Control not found."
    Else
        If TypeName(iCtl) = "Line" Then
            ControlMove nControlName, ControlLeft(nControlName), iCtl.Y1, iCtl.X2 - iCtl.X1, iCtl.Y2 - iCtl.Y1, nTab
        Else
            ControlMove nControlName, ControlLeft(nControlName), iCtl.Top, iCtl.Width, iCtl.Height, nTab
            MakeContainedControlsInSelTabVisible
        End If
    End If
End Sub

Private Sub SetAutoTabHeight()
    Dim iHeight As Single
    Dim t As Long
    Dim iPicHeight As Long
    Dim iOrigHeight As Long
    Dim iVerticalSpaceFromIconToCaption As Long
    
    If Not mAutoTabHeight Then Exit Sub
    
    If Not picAux2.Font Is mFont Then
        Set picAux2.Font = mFont
    End If
    
    iHeight = picAux2.ScaleY(picAux2.TextHeight("Atjq_"), picAux2.ScaleMode, vbHimetric)
    mTabHeight = iHeight * 1.02 + pScaleY(8 * mDPIScale, vbPixels, vbHimetric)
    iOrigHeight = mTabHeight
    
    iVerticalSpaceFromIconToCaption = mTabHeight * 0.05
    For t = 0 To mTabs - 1
        If (Not mTabData(t).DoNotUseIconFont) And (mTabData(t).IconChar <> 0) Then
            Dim iIconCharacter As String
            Dim iIconCharRect As RECT
            Dim iFlags As Long
            Dim iIconFont As StdFont
            Dim iIconPadding As Long
            
            If mTabData(t).IconFont Is Nothing Then
                Set iIconFont = mDefaultIconFont
            Else
                Set iIconFont = mTabData(t).IconFont
            End If
            iIconPadding = pScaleY(3 * mDPIScale + iIconFont.Size * 0.22, vbPixels, vbHimetric)
            iIconCharacter = ChrU(mTabData(t).IconChar)
            iIconCharRect.Left = 0
            iIconCharRect.Top = 0
            iIconCharRect.Right = 0
            iIconCharRect.Bottom = 0
            iFlags = DT_CALCRECT Or DT_SINGLELINE Or DT_CENTER
            Set picAuxIconFont.Font = iIconFont
            DrawTextW picAuxIconFont.hDC, StrPtr(iIconCharacter), -1, iIconCharRect, iFlags Or IIf(mRightToLeft, DT_RTLREADING, 0)
            If (mTabOrientation = ssTabOrientationTop) Or (mTabOrientation = ssTabOrientationBottom) Then
                iPicHeight = (iIconCharRect.Bottom - iIconCharRect.Top)
            Else
                iPicHeight = (iIconCharRect.Right - iIconCharRect.Left)
            End If
            iPicHeight = ScaleY(iPicHeight, vbPixels, vbHimetric)
            If (mIconAlignment = ntIconAlignAtTop) Or (mIconAlignment = ntIconAlignAtBottom) Then
                If (iOrigHeight + iPicHeight + iVerticalSpaceFromIconToCaption) > mTabHeight Then
                    mTabHeight = iOrigHeight + iPicHeight + iVerticalSpaceFromIconToCaption
                End If
            ElseIf iPicHeight + iIconPadding > mTabHeight Then
                mTabHeight = iPicHeight + iIconPadding
            End If
        Else
            If Not mTabData(t).PicToUseSet Then SetPicToUse t
    
            iPicHeight = 0
            If Not mTabData(t).PicToUse Is Nothing Then
                If (mTabOrientation = ssTabOrientationTop) Or (mTabOrientation = ssTabOrientationBottom) Then
                    iPicHeight = mTabData(t).PicToUse.Height
                Else
                    iPicHeight = mTabData(t).PicToUse.Width
                End If
            End If
            iPicHeight = iPicHeight + pScaleY(6 * mDPIScale, vbPixels, vbHimetric)
            If iPicHeight > mTabHeight Then
                mTabHeight = iPicHeight
            End If
        End If
    Next
    iVerticalSpaceFromIconToCaption = pScaleY(iVerticalSpaceFromIconToCaption * 0.15, vbHimetric, vbPixels)
    
    If mAppearanceIsFlat Then
        If mHighlightFlatBar Or mHighlightFlatBarSelectedTab Then
            If (mFlatBarPosition = ntBarPositionTop) And (mTabOrientation <> ssTabOrientationBottom) Or (mFlatBarPosition = ntBarPositionBottom) And (mTabOrientation = ssTabOrientationBottom) Then
                mTabHeight = mTabHeight + ScaleY(mFlatBarHeightDPIScaled, vbPixels, vbHimetric)
                If mHighlightFlatBarWithGrip Or mHighlightFlatBarWithGripSelectedTab Then
                    If mFlatBarGripHeightDPIScaled < 0 Then
                        mTabHeight = mTabHeight - ScaleY(mFlatBarGripHeightDPIScaled, vbPixels, vbHimetric)
                    End If
                End If
            Else
                mTabHeight = mTabHeight + ScaleY(mFlatBarHeightDPIScaled, vbPixels, vbHimetric)
                If mHighlightFlatBarWithGrip Or mHighlightFlatBarWithGripSelectedTab Then
                    If mFlatBarGripHeightDPIScaled < 0 Then
                        'If mFlatBarHeightDPIScaled - Abs(mFlatBarGripHeightDPIScaled) < (mFlatBarHeightDPIScaled * 0.33) Then
                        mTabHeight = mTabHeight - ScaleY(mFlatBarGripHeightDPIScaled, vbPixels, vbHimetric)
                        'End If
                    End If
                End If
            End If
        End If
    End If
    If mTDIMode = ntTDIModeForms Then
        If mTabHeight < UserControl.ScaleY(37, vbPixels, vbHimetric) Then
            mTabHeight = UserControl.ScaleY(37, vbPixels, vbHimetric)
        End If
    End If
    PropertyChanged "TabHeight"
    mSetAutoTabHeightPending = False
End Sub

Private Property Get UserControlContainedControlsCount() As Long
    On Error Resume Next
    UserControlContainedControlsCount = UserControl.ContainedControls.Count
End Property


Private Property Get UserControlContainedControls() As Object
    On Error Resume Next
    Set UserControlContainedControls = UserControl.ContainedControls
    If UserControlContainedControls Is Nothing Then
        Set UserControlContainedControls = New Collection
    End If
End Property

Friend Property Get UserControlWidth() As Single
    UserControlWidth = UserControl.Width
End Property

Friend Property Get UserControlHeight() As Single
    UserControlHeight = UserControl.Height
End Property


Public Property Get OLEDropOnOtherTabs() As Boolean
Attribute OLEDropOnOtherTabs.VB_Description = "Returns/sets a value that determines if the user in a OLE drag operation will be able to drop over inactive tabs or just over the selected tab."
Attribute OLEDropOnOtherTabs.VB_ProcData.VB_Invoke_Property = ";Behavior"
    OLEDropOnOtherTabs = mOLEDropOnOtherTabs
End Property

Public Property Let OLEDropOnOtherTabs(ByVal nValue As Boolean)
    If nValue <> mOLEDropOnOtherTabs Then
        mOLEDropOnOtherTabs = nValue
        SetPropertyChanged "OLEDropOnOtherTabs"
    End If
End Property


Public Property Get TabMousePointerHand() As Boolean
Attribute TabMousePointerHand.VB_Description = "Returns/sets a value that determines if the mouse pointer over tabs will be the hand."
Attribute TabMousePointerHand.VB_ProcData.VB_Invoke_Property = ";Behavior"
    TabMousePointerHand = mTabMousePointerHand
End Property

Public Property Let TabMousePointerHand(ByVal nValue As Boolean)
    If nValue <> mTabMousePointerHand Then
        mTabMousePointerHand = nValue
        SetPropertyChanged "TabMousePointerHand"
    End If
End Property


Public Property Get CanReorderTabs() As Boolean
Attribute CanReorderTabs.VB_Description = "Returns/sets a value that determines whether the user will be able to change tab positions by dragging them."
Attribute CanReorderTabs.VB_ProcData.VB_Invoke_Property = ";Behavior"
    CanReorderTabs = mCanReorderTabs
End Property

Public Property Let CanReorderTabs(ByVal nValue As Boolean)
    If nValue <> mCanReorderTabs Then
        mCanReorderTabs = nValue
        DraggingATab = False
        SetPropertyChanged "CanReorderTabs"
    End If
End Property


Public Property Get TDIMode() As NTTDIModeConstants
Attribute TDIMode.VB_Description = "Returns/sets a value that determines if the control will be used for TDI (tabbed document interface) and which mode (controls or forms). The control behavior changes significantly in this mode because some things are automated. "
Attribute TDIMode.VB_ProcData.VB_Invoke_Property = ";Behavior"
    TDIMode = mTDIMode
End Property

Public Property Let TDIMode(ByVal nValue As NTTDIModeConstants)
    If nValue <> mTDIMode Then
        If mAmbientUserMode Then
            MsgBox "This property cannot be changed at run time!", vbExclamation
            Exit Property
        End If
        If Not mControlJustAdded Then
            MsgBox "This property needs to be set immediately after adding the " & TypeName(Me) & " control (without changing other properties first)", vbExclamation
            Exit Property
        End If
        If (mTDIMode = ntTDIModeNone) And (mTabs <> cPropDef_TabsPerRow) And (mTabs <> 2) Then
            MsgBox "This property needs to be set immediately after adding the " & TypeName(Me) & " control (without changing other properties first)", vbExclamation
            Exit Property
        End If
        If ContainedControls.Count > 0 Then
            MsgBox "This property needs to be set when there is still no contained control inside the " & TypeName(Me) & " control.", vbExclamation
            Exit Property
        End If
        If mTDIMode = ntTDIModeForms Then
            If mAmbientUserMode Then UninstallCBTHook
        End If
        mTDIMode = nValue
        TDIConfigureTDIModeOnce
        If mTDIMode <> ntTDIModeNone Then TDISetTDIMode
        If Not Ambient.UserMode Then lblTDILabel.Visible = (mTDIMode <> ntTDIModeNone)
        If mTDIMode = ntTDIModeControls Then
            lblTDILabel.Caption = "Tabbed Document Interface mode 'Controls'. Use Tab 0 as a template. Add all controls here, all them must be control arrays with Index = 0."
        ElseIf mTDIMode = ntTDIModeForms Then
            lblTDILabel.Caption = "Tabbed Document Interface mode 'Forms'. It is like MDI (Multimple Document Interface) but with tabs. This form must be the main/startup form, and this tab will be as a menu. Put here the controls that go in the main form."
        End If
        SetPropertyChanged "TDIMode"
        mControlJustAdded = True
    End If
End Property


Public Property Get Theme() As String
Attribute Theme.VB_Description = "Returns/sets a value in string format that determines a set of property settings with a name, called ""theme""."
Attribute Theme.VB_ProcData.VB_Invoke_Property = "pagNewTabThemes;Appearance"
Attribute Theme.VB_MemberFlags = "200"
    If mCurrentTheme Is Nothing Then SetCurrentTheme
    Theme = mCurrentThemeName
End Property

Private Sub SetCurrentTheme()
    Dim iTheme As NewTabTheme
    
    Set mCurrentTheme = New NewTabTheme
    mCurrentTheme.ThemeString = GetThemeStringFromControl(Me, Ambient.BackColor, Ambient.ForeColor)
    If mThemesCollection Is Nothing Then Set mThemesCollection = New NewTabThemes
    For Each iTheme In mThemesCollection
        If iTheme.Hash = mCurrentTheme.Hash Then
            mCurrentTheme.Name = iTheme.Name
            Exit For
        End If
    Next
    mCurrentThemeName = mCurrentTheme.Name
    If mCurrentThemeName = "" Then mCurrentThemeName = "Custom"
End Sub

Public Property Let Theme(ByVal nValue As String)
    nValue = LCase$(Trim$(nValue))
    If nValue = "custom" Then Exit Property
    If mCurrentTheme Is Nothing Then SetCurrentTheme
    If nValue = mCurrentThemeName Then Exit Property
    If mThemesCollection Is Nothing Then Set mThemesCollection = New NewTabThemes
    If Not mThemesCollection.Exists(nValue) Then
        RaiseError 380, TypeName(Me)
        Exit Property
    End If
    mCurrentThemeName = nValue
    ApplyThemeToControl mThemesCollection(mCurrentThemeName).Data, Me, Ambient.BackColor, Ambient.ForeColor
End Property


Public Property Get Themes() As NewTabThemes
Attribute Themes.VB_Description = "Returns a collection of NewTabTheme objects. They basically provide the names of the themes that are available (and you can set in the Theme property)."
Attribute Themes.VB_ProcData.VB_Invoke_Property = ";Appearance"
    If mThemesCollection Is Nothing Then Set mThemesCollection = New NewTabThemes
    Set Themes = mThemesCollection
End Property

Friend Property Set Themes(ByVal nThemes As NewTabThemes)
    Dim iTheme As NewTabTheme
    
    If Not nThemes Is Nothing Then
        For Each iTheme In mThemesCollection
            If iTheme.Custom Then
                mThemesCollection.Remove iTheme.Name
            End If
        Next
        For Each iTheme In nThemes
            If iTheme.Custom Then
                mThemesCollection.Add iTheme
            End If
        Next
    End If
    SetPropertyChanged "Themes"
End Property


Public Property Get TabData(ByVal Index As Long) As Long
Attribute TabData.VB_Description = "Used to store any data in Long format, similar to ListBox's ItemData. If the tabs are reordered, it will keep this data for this tab."
Attribute TabData.VB_ProcData.VB_Invoke_Property = ";Data"
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 381, TypeName(Me) ' invalid property array index
        Exit Property
    End If
    TabData = mTabData(Index).Data
End Property

Public Property Let TabData(ByVal Index As Long, ByVal nValue As Long)
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 381, TypeName(Me) ' invalid property array index
        Exit Property
    End If
    If nValue <> mTabData(Index).Data Then
        mTabData(Index).Data = nValue
        PropertyChanged "TabData"
    End If
End Property


Public Property Get TabTag(ByVal Index As Long) As String
Attribute TabTag.VB_Description = "Returns/sets a custom String data for each tab."
Attribute TabTag.VB_ProcData.VB_Invoke_Property = ";Data"
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 381, TypeName(Me) ' invalid property array index
        Exit Property
    End If
    TabTag = mTabData(Index).Tag
End Property

Public Property Let TabTag(ByVal Index As Long, ByVal nValue As String)
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 381, TypeName(Me) ' invalid property array index
        Exit Property
    End If
    If nValue <> mTabData(Index).Tag Then
        mTabData(Index).Tag = nValue
        PropertyChanged "TabTag"
    End If
End Property


Public Property Get TabKey(ByVal Index As Long) As String
Attribute TabKey.VB_Description = "A key that can be used to set the current/active/selected tab with the 'Tab' property.  Keys are case insensitive."
Attribute TabKey.VB_ProcData.VB_Invoke_Property = ";Data"
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 381, TypeName(Me) ' invalid property array index
        Exit Property
    End If
    TabKey = mTabData(Index).Key
End Property

Public Property Let TabKey(ByVal Index As Long, ByVal nValue As String)
    Dim c As Long
    Dim iLCKey As String
    
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 381, TypeName(Me) ' invalid property array index
        Exit Property
    End If
    If nValue <> "" Then
        If IsNumeric(Left$(nValue, 1)) Then
            RaiseError 380, TypeName(Me), "The TabKey must not start with a number.", vbExclamation
            Exit Property
        End If
    End If
    If nValue <> mTabData(Index).Key Then
        If nValue <> "" Then
            For c = 0 To mTabs - 1
                If c <> Index Then
                    iLCKey = LCase$(nValue)
                    If mTabData(c).Key <> "" Then
                        If LCase$(mTabData(c).Key) = iLCKey Then
                            RaiseError 457, TypeName(Me)
                            Exit Property
                        End If
                    End If
                End If
            Next
        End If
        mTabData(Index).Key = nValue
        PropertyChanged "TabKey"
    End If
End Property


Public Property Get SubclassingMethod() As NTSubclassingMethodConstants
Attribute SubclassingMethod.VB_Description = "Returns/sets the sublassing method, what to subclass (only the UserControl or the Form too) and if the subclassing will be enabled at all. If subclassing is disabled or partially disabled some features may not work."
Attribute SubclassingMethod.VB_ProcData.VB_Invoke_Property = ";Advanced"
    SubclassingMethod = mSubclassingMethod
End Property

Public Property Let SubclassingMethod(ByVal nValue As NTSubclassingMethodConstants)
    Dim iPrev As NTSubclassingMethodConstants
    
    If (nValue < ntSMSetWindowSubclass) Or (nValue > ntSM_SWLOnlyUserControl) Then
        RaiseError 380, TypeName(Me) ' invalid property value
        Exit Property
    End If
    If nValue <> SubclassingMethod Then
        iPrev = mSubclassingMethod
        mSubclassingMethod = nValue
        SetPropertyChanged "SubclassingMethod"
        If iPrev <> ntSMDisabled Then
            Unsubclass
        End If
        gSubclassWithSetWindowLong = (mSubclassingMethod = ntSMSetWindowLong) Or (mSubclassingMethod = ntSM_SWLOnlyUserControl)
        mOnlySubclassUserControl = (mSubclassingMethod = ntSM_SWSOnlyUserControl) Or (mSubclassingMethod = ntSM_SWLOnlyUserControl)
        If mSubclassingMethod <> ntSMDisabled Then
            mSubclassed = True
            #If NOSUBCLASSINIDE Then
                If mInIDE Then
                    mSubclassed = False
                End If
            #End If
            Set mSubclassedControlsForPaintingHwnds = New Collection
            Set mSubclassedFramesHwnds = New Collection
            Set mSubclassedControlsForMoveHwnds = New Collection
            If mSubclassed Then
                SubclassUserControl
                SubclassForm
            End If
        End If
    End If
End Property


Private Sub ShowPicCover()
    Dim iRect As RECT
    Const LWA_ALPHA = &H2&
    Const WS_EX_LAYERED = &H80000
    Const GWL_EXSTYLE = (-20)
    Const WS_EX_TOOLWINDOW = &H80
    Const WS_EX_TRANSPARENT As Long = &H20
    Dim iWindowStyle As Long
    Dim iDC As Long
    Static sShowing As Boolean
    Dim iFormRect As RECT
    Dim iPt As POINTAPI
    Dim iFormHwnd As Long
    
    If sShowing Or mSettingTDIMode Then Exit Sub
    sShowing = True
    If picCover.Visible Then
        tmrTabTransition.Enabled = False
        HidePicCover
    End If
    
    GetWindowRect mUserControlHwnd, iRect
    If mFormHwnd = 0 Then
        iFormHwnd = GetAncestor(UserControl.ContainerHwnd, GA_ROOT)
    Else
        iFormHwnd = mFormHwnd
    End If
    GetClientRect iFormHwnd, iFormRect
    ClientToScreen iFormHwnd, iPt
    
    iFormRect.Left = iFormRect.Left + iPt.X
    iFormRect.Right = iFormRect.Right + iPt.X
    iFormRect.Top = iFormRect.Top + iPt.Y
    iFormRect.Bottom = iFormRect.Bottom + iPt.Y
    
    iRect.Right = iRect.Left + mClientRect.Right
    iRect.Bottom = iRect.Top + mClientRect.Bottom
    iRect.Left = iRect.Left + mClientRect.Left
    iRect.Top = iRect.Top + mClientRect.Top
    
    If iFormHwnd <> 0 Then
        If iRect.Right > (iFormRect.Right) Then iRect.Right = iFormRect.Right
        If iRect.Bottom > (iFormRect.Bottom) Then iRect.Bottom = iFormRect.Bottom
    End If
    
    picCover.Visible = False
    SetParent picCover.hWnd, 0
    MoveWindow picCover.hWnd, -100, -100, 1, 1, 0
    
    iWindowStyle = GetWindowLong(picCover.hWnd, GWL_EXSTYLE)
    iWindowStyle = iWindowStyle Or WS_EX_TOOLWINDOW Or WS_EX_LAYERED Or WS_EX_TRANSPARENT
    SetWindowLong picCover.hWnd, GWL_EXSTYLE, iWindowStyle
    SetLayeredWindowAttributes picCover.hWnd, 0, 0, LWA_ALPHA
    picCover.Visible = True ' for some reason this is necessary to avoid a flicker the first time
    SetLayeredWindowAttributes picCover.hWnd, 0, 220, LWA_ALPHA
    picCover.Visible = False
    
    picCover.Visible = True
    
    mTabTransition_Step = 5
    tmrTabTransition.Interval = 10 ' 500
    tmrTabTransition.Enabled = True
    MoveWindow picCover.hWnd, iRect.Left, iRect.Top, iRect.Right - iRect.Left, iRect.Bottom - iRect.Top, 0
    iDC = GetDC(mUserControlHwnd)
    BitBlt picCover.hDC, 0, 0, iRect.Right - iRect.Left, iRect.Bottom - iRect.Top, iDC, mClientRect.Left, mClientRect.Top, vbSrcCopy
    ReleaseDC mUserControlHwnd, iDC
    picCover.Refresh
    
    Sleep 30

    sShowing = False
End Sub

Private Sub HidePicCover()
    Const WS_EX_LAYERED = &H80000
    Const GWL_EXSTYLE = (-20)
    Const WS_EX_TOOLWINDOW = &H80
    Dim iWindowStyle As Long
    
    picCover.Visible = False
    picCover.Cls
    SetParent picCover.hWnd, mUserControlHwnd
    iWindowStyle = GetWindowLong(picCover.hWnd, GWL_EXSTYLE)
    iWindowStyle = iWindowStyle And Not WS_EX_TOOLWINDOW And Not WS_EX_LAYERED
    SetWindowLong picCover.hWnd, GWL_EXSTYLE, iWindowStyle
    mTabTransition_Step = 0
End Sub

Friend Property Get AmbientFont() As StdFont
    On Error Resume Next
    Set AmbientFont = Ambient.Font
End Property

Friend Property Get AmbientBackColor() As Long
    On Error Resume Next
    AmbientBackColor = Ambient.BackColor
End Property

Friend Property Get AmbientForeColor() As Long
    On Error Resume Next
    AmbientForeColor = Ambient.ForeColor
End Property

Private Function GetAutomaticBackColorSelectedTab() As Long
    Dim iBackColorTabs_H As Integer
    Dim iBackColorTabs_L As Integer
    Dim iBackColorTabs_S As Integer
    Dim iBCol As Long
    Dim iCol_L As Integer
    
    If mStyle = ntStyleFlat Then

        If mHandleHighContrastTheme And (mHighContrastThemeOn Or mChangingHighContrastTheme) And (mBackColorTabs = vbButtonFace) Then
            GetAutomaticBackColorSelectedTab = vbButtonFace
        Else
            iBCol = TranslatedColor(mBackColorTabs)
            ColorRGBToHLS iBCol, iBackColorTabs_H, iBackColorTabs_L, iBackColorTabs_S
            If iBackColorTabs_L > 150 Then
                If (iBackColorTabs_L > 200) And (iBackColorTabs_S < 150) Then
                    iCol_L = iBackColorTabs_L * 1.08
                    If iCol_L > 240 Then iCol_L = 240
                    GetAutomaticBackColorSelectedTab = ColorHLSToRGB(iBackColorTabs_H, iCol_L, iBackColorTabs_S * 0.8)
                Else
                    iCol_L = iBackColorTabs_L * 1.08
                    If iCol_L > 240 Then iCol_L = 240
                    GetAutomaticBackColorSelectedTab = ColorHLSToRGB(iBackColorTabs_H, iCol_L, iBackColorTabs_S * 0.5)
                End If
            Else
                iCol_L = iBackColorTabs_L * 1.35
                If iCol_L > 240 Then iCol_L = 240
                GetAutomaticBackColorSelectedTab = ColorHLSToRGB(iBackColorTabs_H, iCol_L, iBackColorTabs_S)
            End If
        End If
    Else
        GetAutomaticBackColorSelectedTab = mBackColorTabs
    End If
End Function

Private Function CloneFont(nOrigFont As iFont) As StdFont
    If nOrigFont Is Nothing Then Exit Function
    nOrigFont.Clone CloneFont
End Function

Private Function FontsAreEqual(nFont1 As StdFont, nFont2 As StdFont) As Boolean
    If nFont1 Is Nothing Or nFont2 Is Nothing Then Exit Function
    
    If (nFont1 Is Nothing) And (nFont2 Is Nothing) Then
        FontsAreEqual = True
        Exit Function
    End If
    If (nFont1 Is Nothing) Then Exit Function
    If (nFont2 Is Nothing) Then Exit Function
    
    If nFont1.Name = nFont2.Name Then
        If nFont1.Size = nFont2.Size Then
            If nFont1.Bold = nFont2.Bold Then
                If nFont1.Italic = nFont2.Italic Then
                    If nFont1.Strikethrough = nFont2.Strikethrough Then
                        If nFont1.Underline = nFont2.Underline Then
                            If nFont1.Weight = nFont2.Weight Then
                                If nFont1.Charset = nFont2.Charset Then
                                    FontsAreEqual = True
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function

Private Function ChrU(ByVal nCharCodeU As Long) As String
    Const cPOW10 As Long = 2 ^ 10
    
'    If (nCharCodeU >= 0) And (nCharCodeU <= &HFF&) Then
'        ChrU = Chr$(nCharCodeU)
'    ElseIf nCharCodeU <= &HFFFF& Then
    If nCharCodeU <= &HFFFF& Then
        ChrU = ChrW$(nCharCodeU)
    Else
        ChrU = ChrW$(&HD800& + (nCharCodeU And &HFFFF&) \ cPOW10) & ChrW$(&HDC00& + (nCharCodeU And (cPOW10 - 1)))
    End If
End Function

Private Sub SetHighlightMode()
    Dim iHighlightMode As NTHighlightModeFlagsConstants
    Dim iHighlightModeSelectedTab As NTHighlightModeFlagsConstants
    
    iHighlightMode = mHighlightMode
    If iHighlightMode = ntHLAuto Then
        If mTDIMode <> ntTDIModeNone Then
            iHighlightMode = (ntHLBackgroundPlain Or ntHLCaptionBold)
        ElseIf mStyle = ntStyleFlat Then
            iHighlightMode = (ntHLBackgroundGradient Or ntHLBackgroundLight Or ntHLFlatBar)
        ElseIf mStyle = ntStyleWindows Then
            iHighlightMode = ntHLNone
        ElseIf mStyle = ssStyleTabbedDialog Then
            iHighlightMode = ntHLNone
        Else
            iHighlightMode = ntHLNone ' ntHLBackgroundDoubleGradient
        End If
    End If
    If (iHighlightMode And ntHLBackgroundTypeFilter) = ntHLBackgroundDoubleGradient Then
        mHighlightGradient = ntGradientDouble
    ElseIf (iHighlightMode And ntHLBackgroundTypeFilter) = ntHLBackgroundGradient Then
        mHighlightGradient = ntGradientSimple
    ElseIf (iHighlightMode And ntHLBackgroundTypeFilter) = ntHLBackgroundPlain Then
        mHighlightGradient = ntGradientPlain
    Else
        mHighlightGradient = ntGradientNone
    End If
    
    mHighlightCaptionBold = (iHighlightMode <> ntHLNone) And ((iHighlightMode And ntHLCaptionBold) = ntHLCaptionBold)
    mHighlightCaptionUnderlined = (iHighlightMode <> ntHLNone) And ((iHighlightMode And ntHLCaptionUnderlined) = ntHLCaptionUnderlined)
    If (iHighlightMode <> ntHLNone) And ((iHighlightMode And ntHLBackgroundLight) = ntHLBackgroundLight) Then
        mHighlightIntensity = ntHighlightIntensityLight
    Else
        mHighlightIntensity = ntHighlightIntensityStrong
    End If
    mHighlightFlatBar = (iHighlightMode <> ntHLNone) And ((iHighlightMode And ntHLFlatBar) = ntHLFlatBar)
    mHighlightFlatBarWithGrip = mHighlightFlatBar And ((iHighlightMode And ntHLFlatBarGrip) = ntHLFlatBarGrip)
    mHighlightAddExtraHeight = (iHighlightMode <> ntHLNone) And ((iHighlightMode And ntHLExtraHeight) = ntHLExtraHeight)
    mHighlightFlatDrawBorder = (iHighlightMode <> ntHLNone) And ((iHighlightMode And ntHLFlatDrawBorder) = ntHLFlatDrawBorder)
    
    iHighlightModeSelectedTab = mHighlightModeSelectedTab
    If iHighlightModeSelectedTab = ntHLAuto Then
        If mTDIMode <> ntTDIModeNone Then
            iHighlightModeSelectedTab = (ntHLBackgroundPlain Or ntHLBackgroundLight)
        ElseIf mStyle = ntStyleFlat Then
            iHighlightModeSelectedTab = (ntHLBackgroundGradient Or ntHLBackgroundLight Or ntHLFlatBar)
        ElseIf mStyle = ntStyleWindows Then
            iHighlightModeSelectedTab = ntHLNone
        ElseIf mStyle = ssStyleTabbedDialog Then
            iHighlightModeSelectedTab = ntHLCaptionBold
        Else
            iHighlightModeSelectedTab = ntHLNone ' ntHLBackgroundDoubleGradient
        End If
    End If
    
    If (iHighlightModeSelectedTab And ntHLBackgroundTypeFilter) = ntHLBackgroundDoubleGradient Then
        mHighlightGradientSelectedTab = ntGradientDouble
    ElseIf (iHighlightModeSelectedTab And ntHLBackgroundTypeFilter) = ntHLBackgroundGradient Then
        mHighlightGradientSelectedTab = ntGradientSimple
    ElseIf (iHighlightModeSelectedTab And ntHLBackgroundTypeFilter) = ntHLBackgroundPlain Then
        mHighlightGradientSelectedTab = ntGradientPlain
    Else
        mHighlightGradientSelectedTab = ntGradientNone
    End If
    
    mHighlightCaptionBoldSelectedTab = (iHighlightModeSelectedTab <> ntHLNone) And ((iHighlightModeSelectedTab And ntHLCaptionBold) = ntHLCaptionBold)
    mHighlightCaptionUnderlinedSelectedTab = (iHighlightModeSelectedTab <> ntHLNone) And ((iHighlightModeSelectedTab And ntHLCaptionUnderlined) = ntHLCaptionUnderlined)
    If (iHighlightModeSelectedTab <> ntHLNone) And ((iHighlightModeSelectedTab And ntHLBackgroundLight) = ntHLBackgroundLight) Then
        mHighlightIntensitySelectedTab = ntHighlightIntensityLight
    Else
        mHighlightIntensitySelectedTab = ntHighlightIntensityStrong
    End If
    mHighlightFlatBarSelectedTab = (iHighlightModeSelectedTab <> ntHLNone) And ((iHighlightModeSelectedTab And ntHLFlatBar) = ntHLFlatBar)
    mHighlightFlatBarWithGripSelectedTab = mHighlightFlatBarSelectedTab And ((iHighlightModeSelectedTab And ntHLFlatBarGrip) = ntHLFlatBarGrip)
    mHighlightAddExtraHeightSelectedTab = (iHighlightModeSelectedTab <> ntHLNone) And ((iHighlightModeSelectedTab And ntHLExtraHeight) = ntHLExtraHeight)
    mHighlightFlatDrawBorderSelectedTab = (iHighlightModeSelectedTab <> ntHLNone) And ((iHighlightModeSelectedTab And ntHLFlatDrawBorder) = ntHLFlatDrawBorder)
    
    If mHighlightIntensitySelectedTab = ntHighlightIntensityStrong Then
        mGlowColor_Sel = mGlowColor_Sel_Bk
    Else
        mGlowColor_Sel = mGlowColor_Sel_Light
    End If
End Sub

Public Function GetTabLeft(ByVal Index As Long) As Single
Attribute GetTabLeft.VB_Description = "Returns the left position of a tab."
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 5, TypeName(Me) ' Invalid procedure call or argument
        Exit Function
    End If
    EnsureDrawn
    GetTabLeft = FixRoundingError(UserControl.ScaleX(mTabData(Index).TabRect.Left, vbPixels, vbTwips))
End Function

Public Function GetTabTop(ByVal Index As Long) As Single
Attribute GetTabTop.VB_Description = "Returns the top position of a tab."
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 5, TypeName(Me) ' Invalid procedure call or argument
        Exit Function
    End If
    EnsureDrawn
    GetTabTop = FixRoundingError(UserControl.ScaleY(mTabData(Index).TabRect.Top, vbPixels, vbTwips))
End Function

Public Function GetTabWidth(ByVal Index As Long) As Single
Attribute GetTabWidth.VB_Description = "Returns the width of a tab."
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 5, TypeName(Me) ' Invalid procedure call or argument
        Exit Function
    End If
    EnsureDrawn
    GetTabWidth = FixRoundingError(UserControl.ScaleX(mTabData(Index).TabRect.Right - mTabData(Index).TabRect.Left, vbPixels, vbTwips))
End Function

Public Function GetTabHeight(ByVal Index As Long) As Single
Attribute GetTabHeight.VB_Description = "Returns the height of a tab."
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 5, TypeName(Me) ' Invalid procedure call or argument
        Exit Function
    End If
    EnsureDrawn
    GetTabHeight = FixRoundingError(UserControl.ScaleY(mTabData(Index).TabRect.Bottom - mTabData(Index).TabRect.Top, vbPixels, vbTwips))
End Function

Private Sub SetPropertyChanged(Optional nPropertyName As String)
    If mPropertiesReady Then
        If Not mSettingTDIMode Then
            PropertyChanged nPropertyName
            mControlJustAdded = False
            Set mCurrentTheme = Nothing
        End If
    End If
End Sub

Public Sub MoveTab(ByVal CurrentIndex As Long, ByVal NewIndex As Long)
Attribute MoveTab.VB_Description = "Moves a tab to another position."
    Dim iCanceled As Boolean
    Dim iTempTabData As T_TabData
    Dim c As Long
    Dim iCurTab As Boolean
    Dim iRedraw As Boolean
    
    If NewIndex = CurrentIndex Then Exit Sub
    If (CurrentIndex < 0) Or (CurrentIndex > (mTabs - 1)) Or (NewIndex < 0) Or (NewIndex > (mTabs - 1)) Then
        RaiseError 5, TypeName(Me)
        Exit Sub
    End If
    
    RaiseEvent BeforeTabReorder(CurrentIndex, NewIndex, iCanceled)
    If NewIndex = CurrentIndex Then Exit Sub
    If iCanceled Then Exit Sub
    
    mMovingATab = True
    iRedraw = Redraw
    Redraw = False
    iCurTab = (CurrentIndex = mTab)
    
    iTempTabData = mTabData(CurrentIndex)
    If NewIndex > CurrentIndex Then
        For c = CurrentIndex + 1 To NewIndex
            mTabData(c - 1) = mTabData(c)
        Next
        mTabData(NewIndex) = iTempTabData
    Else
        For c = CurrentIndex - 1 To NewIndex Step -1
            mTabData(c + 1) = mTabData(c)
        Next
        mTabData(NewIndex) = iTempTabData
    End If
    If iCurTab Then
        mTab = -1
        If TabVisible(NewIndex) Then
            TabSel = NewIndex
            Draw
        End If
    Else
        For c = 0 To mTabs - 1
            If mTabData(c).Selected Then
                If TabVisible(c) Then
                    TabSel = c
                End If
                Exit For
            End If
        Next
        Draw
    End If
    RecreateTabIconFontsEventHandler
    Redraw = iRedraw
    RaiseEvent TabReordered(NewIndex, CurrentIndex)
    mMovingATab = False
End Sub

Private Sub RecreateTabIconFontsEventHandler()
    Dim t As Long
    
    mTabIconFontsEventsHandler.Release
    Set mTabIconFontsEventsHandler = New cFontEventHandlers
    
    For t = 0 To mTabs - 1
        If Not mTabData(t).IconFont Is Nothing Then
            mTabIconFontsEventsHandler.AddFont mTabData(t).IconFont, t
        End If
    Next
End Sub

Public Property Get VisibleTabs() As Long
Attribute VisibleTabs.VB_Description = "Returns the numbers of tabs that are visible. [TabVisible(Index) = True]."
Attribute VisibleTabs.VB_ProcData.VB_Invoke_Property = ";Data"
    EnsureDrawn
    VisibleTabs = mVisibleTabs
End Property


Private Property Get DraggingATab() As Boolean
    DraggingATab = mDraggingATab And ((mMouseX <> 0 And mMouseX2 <> 0) Or (mMouseY <> 0 And mMouseY2 <> 0))
End Property

Private Property Let DraggingATab(nValue As Boolean)
    Dim iRc As RECT
    Dim iPt As POINTAPI
    Dim iPt2 As POINTAPI
    Dim iPt3 As POINTAPI
    
    If mMovingATab Then Exit Property
    If nValue = mDraggingATab Then Exit Property
    
    mDraggingATab = nValue And (mTab > -1)
    If mDraggingATab Then
        mPreviousTabBeforeDragging = mTab
        ClientToScreen mUserControlHwnd, iPt
        
        If mTabOrientation = ssTabOrientationBottom Then
            iRc.Left = iPt.X
            iRc.Right = UserControl.ScaleX(UserControl.ScaleWidth, UserControl.ScaleMode, vbPixels) + iPt.X
            iPt2.Y = mClientRect.Bottom
            iPt3.Y = mScaleHeight
            ClientToScreen mUserControlHwnd, iPt2
            ClientToScreen mUserControlHwnd, iPt3
            iRc.Top = iPt2.Y + mMouseY - mTabData(mTab).TabRect.Top
            iRc.Bottom = iPt3.Y + mMouseY - mTabData(mTab).TabRect.Bottom
        ElseIf (mTabOrientation = ssTabOrientationLeft) Or (mTabOrientation = ntTabOrientationLeftHorizontal) Then
            iRc.Left = iPt.X + mMouseX - mTabData(mTab).TabRect.Left
            iRc.Right = iPt.X + mClientRect.Left + mMouseX - mTabData(mTab).TabRect.Right
            iRc.Top = iPt.Y
            iPt3.Y = mScaleHeight
            ClientToScreen mUserControlHwnd, iPt3
            iRc.Bottom = iPt3.Y
        ElseIf (mTabOrientation = ssTabOrientationRight) Or (mTabOrientation = ntTabOrientationRightHorizontal) Then
            iRc.Left = iPt.X + mClientRect.Left + mClientRect.Right + mMouseX - mTabData(mTab).TabRect.Left
            iRc.Right = UserControl.ScaleX(UserControl.ScaleWidth, UserControl.ScaleMode, vbPixels) + iPt.X + mClientRect.Left + mMouseX - mTabData(mTab).TabRect.Right
            iRc.Top = iPt.Y
            iPt3.Y = mScaleHeight
            ClientToScreen mUserControlHwnd, iPt3
            iRc.Bottom = iPt3.Y
        Else
            iRc.Left = iPt.X
            iRc.Right = UserControl.ScaleX(UserControl.ScaleWidth, UserControl.ScaleMode, vbPixels) + iPt.X
            iRc.Top = iPt.Y + mMouseY - mTabData(mTab).TabRect.Top
            iRc.Bottom = mClientRect.Top + iPt.Y + mMouseY - mTabData(mTab).TabRect.Bottom
        End If
        
        tmrTabDragging.Enabled = True
        If mTDIMode = ntTDIModeForms Then
            If mRows = 1 Then
                iRc.Left = iRc.Left + mTabData(0).TabRect.Right - mTabData(0).TabRect.Left
            End If
        End If
        ClipCursor iRc
    Else
        mMouseX = 0
        mMouseY = 0
        tmrTabDragging.Enabled = False
        ClipCursor ByVal 0
    End If
End Property


Friend Property Get BackColorSelectedTab_IsAutomatic() As Boolean
    BackColorSelectedTab_IsAutomatic = mBackColorSelectedTab_IsAutomatic
End Property

Friend Property Get FlatBarColorHighlight_IsAutomatic() As Boolean
    FlatBarColorHighlight_IsAutomatic = mFlatBarColorHighlight_IsAutomatic
End Property

Friend Property Get HighlightColor_IsAutomatic() As Boolean
    HighlightColor_IsAutomatic = mHighlightColor_IsAutomatic
End Property

Friend Property Get HighlightColorSelectedTab_IsAutomatic() As Boolean
    HighlightColorSelectedTab_IsAutomatic = mHighlightColorSelectedTab_IsAutomatic
End Property

Friend Property Get FlatBarColorInactive_IsAutomatic() As Boolean
    FlatBarColorInactive_IsAutomatic = mFlatBarColorInactive_IsAutomatic
End Property

Friend Property Get FlatTabsSeparationLineColor_IsAutomatic() As Boolean
    FlatTabsSeparationLineColor_IsAutomatic = mFlatTabsSeparationLineColor_IsAutomatic
End Property

Friend Property Get FlatBodySeparationLineColor_IsAutomatic() As Boolean
    FlatBodySeparationLineColor_IsAutomatic = mFlatBodySeparationLineColor_IsAutomatic
End Property

Friend Property Get FlatBorderColor_IsAutomatic() As Boolean
    FlatBorderColor_IsAutomatic = mFlatBorderColor_IsAutomatic
End Property


' Extender properties and methods
Public Property Get Name() As String
Attribute Name.VB_Description = "Returns the name used in code to identify the control."
Attribute Name.VB_ProcData.VB_Invoke_Property = ";Misc"
    Name = Ambient.DisplayName
End Property

Public Property Get Tag() As String
Attribute Tag.VB_Description = "Returns/sets an expression that stores any extra data needed for your program. "
Attribute Tag.VB_ProcData.VB_Invoke_Property = ";Data"
    Tag = Extender.Tag
End Property

Public Property Let Tag(ByVal Value As String)
    Extender.Tag = Value
End Property

Public Property Get Parent() As Object
Attribute Parent.VB_Description = "Returns the object in which the control is located."
    Set Parent = UserControl.Parent
End Property

Public Property Get Container() As Object
Attribute Container.VB_Description = "Returns the control's container."
    Set Container = Extender.Container
End Property

Public Property Set Container(ByVal Value As Object)
    Set Extender.Container = Value
End Property

Public Property Get Left() As Single
Attribute Left.VB_Description = "Returns/sets the distance between the internal left edge of the control and the left edge of its container."
Attribute Left.VB_ProcData.VB_Invoke_Property = ";Position"
    Left = Extender.Left
End Property

Public Property Let Left(ByVal Value As Single)
    Extender.Left = Value
End Property

Public Property Get Top() As Single
Attribute Top.VB_Description = "Returns/sets the distance between the internal top edge of the control and the top edge of its container."
Attribute Top.VB_ProcData.VB_Invoke_Property = ";Position"
    Top = Extender.Top
End Property

Public Property Let Top(ByVal Value As Single)
    Extender.Top = Value
End Property

Public Property Get Width() As Single
Attribute Width.VB_Description = "Returns/sets the width of the control."
Attribute Width.VB_ProcData.VB_Invoke_Property = ";Position"
    Width = Extender.Width
End Property

Public Property Let Width(ByVal Value As Single)
    Extender.Width = Value
End Property

Public Property Get Height() As Single
Attribute Height.VB_Description = "Returns/sets the height of the control."
Attribute Height.VB_ProcData.VB_Invoke_Property = ";Position"
    Height = Extender.Height
End Property

Public Property Let Height(ByVal Value As Single)
    Extender.Height = Value
End Property

Public Property Get Visible() As Boolean
Attribute Visible.VB_Description = "Returns/sets a value indicating whether the control is visible or hidden."
Attribute Visible.VB_ProcData.VB_Invoke_Property = ";Behavior"
    Visible = Extender.Visible
End Property

Public Property Let Visible(ByVal Value As Boolean)
    Extender.Visible = Value
End Property

Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/sets a the tool tip text to be displayed when the mouse is over the control."
Attribute ToolTipText.VB_ProcData.VB_Invoke_Property = ";Misc"
    ToolTipText = Extender.ToolTipText
End Property

Public Property Let ToolTipText(ByVal Value As String)
    Extender.ToolTipText = Value
End Property

Public Property Get HelpContextID() As Long
Attribute HelpContextID.VB_Description = "Returns/sets a string expression containing the context ID for a topic in a Help file."
Attribute HelpContextID.VB_ProcData.VB_Invoke_Property = ";Misc"
    HelpContextID = Extender.HelpContextID
End Property

Public Property Let HelpContextID(ByVal Value As Long)
    Extender.HelpContextID = Value
End Property

Public Property Get WhatsThisHelpID() As Long
Attribute WhatsThisHelpID.VB_Description = "Returns/sets an associated context number for the control. "
Attribute WhatsThisHelpID.VB_ProcData.VB_Invoke_Property = ";Misc"
    WhatsThisHelpID = Extender.WhatsThisHelpID
End Property

Public Property Let WhatsThisHelpID(ByVal Value As Long)
    Extender.WhatsThisHelpID = Value
End Property

Public Property Get DragIcon() As IPictureDisp
Attribute DragIcon.VB_Description = "Returns/sets the icon to be used as mouse pointer in a drag-and-drop operation."
Attribute DragIcon.VB_ProcData.VB_Invoke_Property = ";Behavior"
    Set DragIcon = Extender.DragIcon
End Property

Public Property Let DragIcon(ByVal Value As IPictureDisp)
    Extender.DragIcon = Value
End Property

Public Property Set DragIcon(ByVal Value As IPictureDisp)
    Set Extender.DragIcon = Value
End Property

Public Property Get DragMode() As Integer
Attribute DragMode.VB_Description = "Returns/sets a value that determines whether manual or automatic drag mode is used for a drag-and-drop operation."
Attribute DragMode.VB_ProcData.VB_Invoke_Property = ";Behavior"
    DragMode = Extender.DragMode
End Property

Public Property Let DragMode(ByVal Value As Integer)
    Extender.DragMode = Value
End Property

Public Sub Drag(Optional ByRef Action As Variant)
Attribute Drag.VB_Description = "Begins, ends, or cancels a drag operation."
    If IsMissing(Action) Then Extender.Drag Else Extender.Drag Action
End Sub

Public Sub SetFocus()
Attribute SetFocus.VB_Description = "Moves the focus to this control."
    Extender.SetFocus
End Sub

Public Sub ZOrder(Optional ByRef Position As Variant)
Attribute ZOrder.VB_Description = "Places the control at the front or back of the z-order within its graphical level."
    If IsMissing(Position) Then Extender.ZOrder Else Extender.ZOrder Position
End Sub

Private Sub TDIConfigureTDIModeOnce()
    Dim iFont As StdFont
    
    CanReorderTabs = True
    IconColorMouseHover = vbRed
    IconColorMouseHoverSelectedTab = vbRed
    mTDIIconColorMouseHover = IconColorMouseHover
    
    mTDIChangingTabCount = True
    If mTDIMode = ntTDIModeControls Then
        Tabs = 2
    ElseIf mTDIMode = ntTDIModeForms Then
        Tabs = 1
        TabIconLeftOffset(0) = 0
        TabIconTopOffset(0) = 0
        TabIconCharHex(0) = ""
    Else
        Tabs = 1
        TabCaption(0) = "Tab 0"
        Tabs = cPropDef_TabsPerRow
        MoveTab 0, 2
        Tabs = 1
        TabCaption(0) = "Tab 0"
        Tabs = cPropDef_TabsPerRow
        mCanReorderTabs = cPropDef_CanReorderTabs
        mIconColorMouseHover = Ambient.ForeColor
        mIconColorMouseHoverSelectedTab = Ambient.ForeColor
    End If
    mTDIChangingTabCount = False
End Sub
    
Private Function FontExists(nFontName As String) As Boolean
    Dim iFont As New StdFont
    
    'If mInIDE Then If nFontName = "Segoe MDL2 Assets" Then Exit Function
    
    If nFontName = "[Auto]" Then Exit Function
    If Trim$(nFontName) = "" Then Exit Function
    
    iFont.Name = nFontName
    FontExists = StrComp(nFontName, iFont.Name, vbTextCompare) = 0
End Function
    
    
Private Sub TDISetTDIMode()
    Dim iTabCaption As String
    Dim iLoadTabControls As Boolean
    Dim iFont As StdFont
    Dim iShowTabCloseButton As Boolean
    
    Redraw = False
    
    Set iFont = New StdFont
    If FontExists("Segoe MDL2 Assets") Then
        If mTDIMode = ntTDIModeControls Then
            iFont.Name = "Segoe MDL2 Assets"
            iFont.Bold = True
            iFont.Size = 10 ' size of the X to close the tab
            Set TabIconFont(0) = iFont
        End If
        If mTDIMode = ntTDIModeControls Then
            Set TabIconFont(1) = CloneFont(iFont)
            TabIconFont(1).Size = 10 ' size of the + to add a new tab
            TabIconFont(1).Bold = True
            TabToolTipText(1) = "Add a new tab"
            TabIconLeftOffset(1) = -2 * mDPIScale
            TabIconTopOffset(1) = 1 * mDPIScale
            TabIconCharHex(1) = "&HF8AA&"
            TabIconLeftOffset(0) = -3 * mDPIScale
            TabIconTopOffset(0) = 1 * mDPIScale
            TabIconCharHex(0) = "&HE106&"
        End If
    ElseIf FontExists("Marlett") Then
        iFont.Name = "Marlett"
        'iFont.Bold = True
        iFont.Size = 15 ' size of the X to close the tab
        Set TabIconFont(0) = iFont
        If mTDIMode = ntTDIModeControls Then
            Set iFont = New StdFont
            iFont.Name = "Arial"
            iFont.Size = 15
            iFont.Bold = True
            Set TabIconFont(1) = CloneFont(iFont)
            TabToolTipText(1) = "Add a new tab"
            TabIconLeftOffset(1) = -2 * mDPIScale
            TabIconTopOffset(1) = 2 * mDPIScale
            TabIconCharHex(1) = "&H2B&"
            TabIconLeftOffset(0) = -3 * mDPIScale
            TabIconTopOffset(0) = 1 * mDPIScale
            TabIconCharHex(0) = "&H72&"
        End If
    Else
        iFont.Name = "Arial"
        iFont.Size = 15
        iFont.Bold = True
        Set TabIconFont(0) = iFont
        If mTDIMode = ntTDIModeControls Then
            Set TabIconFont(1) = CloneFont(iFont)
            TabToolTipText(1) = "Add a new tab"
            TabIconLeftOffset(1) = -2 * mDPIScale
            TabIconTopOffset(1) = 2 * mDPIScale
            TabIconCharHex(1) = "&H2B&"
            TabIconLeftOffset(0) = 0
            TabIconTopOffset(0) = 0
            TabIconCharHex(0) = "&H78&"
        End If
    End If
    
    mSettingTDIMode = True
    mTDIIconColorMouseHover = IconColorMouseHover
    mTDIChangingTabCount = True
    If mTDIMode = ntTDIModeControls Then
        Tabs = 2
        TabVisible(1) = True
    ElseIf mTDIMode = ntTDIModeForms Then
        Tabs = 1
    End If
    mTDIChangingTabCount = False
    If mTDIMode = ntTDIModeControls Then
        TabCaption(1) = ""
        mTabData(1).Data = -1
    End If
    mIconAlignment = ntIconAlignEnd
    mBackColor = Ambient.BackColor
    
    If Not mAmbientUserMode Then
        TabSel = 0
        If mTDIMode = ntTDIModeControls Then
            TabCaption(0) = "New tab template   "
        ElseIf (TabCaption(0) = "New tab template   ") Or (TabCaption(0) = "Tab 0") Then
            TabCaption(0) = "Home"
        End If
        lblTDILabel.Visible = mTDIMode <> ntTDIModeNone
        lblTDILabel.ZOrder
        TabVisible(0) = True
        lblTDILabel.ForeColor = mForeColorSelectedTab
        DrawDelayed
    Else
        If Not FontExists(TabIconFont(0).Name) Then
            Set iFont = New StdFont
            If FontExists("Segoe MDL2 Assets") Then
                iFont.Name = "Segoe MDL2 Assets"
                iFont.Size = 10
                iFont.Bold = True
                Set TabIconFont(0) = iFont
                TabIconLeftOffset(0) = -3 * mDPIScale
                TabIconTopOffset(0) = 1 * mDPIScale
                If TabIconCharHex(0) <> "" Then TabIconCharHex(0) = "&HE106&"
            ElseIf FontExists("Marlett") Then
                iFont.Name = "Marlett"
                iFont.Size = 15
                iFont.Bold = True
                Set TabIconFont(0) = iFont
                TabIconLeftOffset(0) = 0
                TabIconTopOffset(0) = 0
                If TabIconCharHex(0) <> "" Then TabIconCharHex(0) = "&H72&"
            Else
                iFont.Name = "Arial"
                iFont.Size = 15
                iFont.Bold = True
                Set TabIconFont(0) = iFont
                TabIconLeftOffset(0) = 0
                TabIconTopOffset(0) = 0
                If TabIconCharHex(0) <> "" Then TabIconCharHex(0) = "&H78&"
            End If
        End If
        If mTDIMode = ntTDIModeControls Then
            If Not FontExists(TabIconFont(1).Name) Then
                Set iFont = New StdFont
                If FontExists("Segoe MDL2 Assets") Then
                    iFont.Name = "Segoe MDL2 Assets"
                    iFont.Size = 10
                    iFont.Bold = True
                    Set TabIconFont(1) = iFont
                    TabToolTipText(1) = "Add a new tab"
                    TabIconLeftOffset(1) = -2 * mDPIScale
                    TabIconTopOffset(1) = 1 * mDPIScale
                    If TabIconCharHex(1) <> "" Then TabIconCharHex(1) = "&HF8AA&"
                Else
                    iFont.Name = "Arial"
                    iFont.Size = 15
                    iFont.Bold = True
                    Set TabIconFont(1) = iFont
                    TabToolTipText(1) = "Add a new tab"
                    TabIconLeftOffset(1) = -2 * mDPIScale
                    TabIconTopOffset(1) = 2 * mDPIScale
                    If TabIconCharHex(1) <> "" Then TabIconCharHex(1) = "&H2B&"
                End If
            End If
        End If
        mTDIIconColorMouseHover = mIconColorMouseHover
        mIconColorMouseHover = mIconColor
        mIconColorMouseHoverSelectedTab = mIconColor
        If mTDIMode = ntTDIModeControls Then
            TDIStoreTab0ControlInfo
            mTDILastTabNumber = mTDILastTabNumber + 1
            If mTDIMode = ntTDIModeControls Then
                iTabCaption = "Default tab"
            End If
            iLoadTabControls = True
            iShowTabCloseButton = True
            RaiseEvent TDIBeforeNewTab(ntDefaultTab, mTDILastTabNumber, iTabCaption, iLoadTabControls, iShowTabCloseButton, False)
            TDIPrepareNewTab iTabCaption, iLoadTabControls, , , iShowTabCloseButton
            TabVisible(0) = False
        End If
    End If
    mSettingTDIMode = False
    If mTDIMode = ntTDIModeForms Then
        If mAmbientUserMode Then
            ReDim mTDIModeFormsFormData_FormHwnd(100)
            ReDim mTDIModeFormsFormData_OldParentHwnd(100)
            ReDim mTDIModeFormsFormData_FormIcon(100)
            ReDim mTDIFormIconCustomData(100)
            ReDim mTDIFormWithoutCloseButton(100)
            InstallCBTHook Me
        End If
    End If
    Redraw = True
    
End Sub

Private Sub TDIStoreTab0ControlInfo()
    Dim c As Long
    Dim ub As Long
    Dim i As Long
    Dim iCtl As Object
    
    c = -1
    ub = 100
    ReDim mTDIControlNames(ub)
    For Each iCtl In TabControls(0)
        i = -1
        On Error Resume Next
        i = iCtl.Index
        On Error GoTo 0
        If i = 0 Then ' only controls with Index = 0
            c = c + 1
            If c > ub Then
                ub = ub + 100
                ReDim Preserve mTDIControlNames(ub)
            End If
            mTDIControlNames(c) = iCtl.Name
        End If
    Next
    mTDIControlNames_Count = c + 1
    If mTDIControlNames_Count > 0 Then
        ReDim Preserve mTDIControlNames(c)
    Else
        ReDim mTDIControlNames(c To c)
    End If
End Sub

Private Function TDIAddNewTab(Optional Position As Variant, Optional Focused As Boolean = True, Optional ByVal nTabType As NTTDINewTabTypeConstants = ntNewTabByClickingIcon, Optional nTabCaption As String = "") As Boolean
    Dim iCancel As Boolean
    Dim iLoadTabControls As Boolean
    Dim iShowTabCloseButton As Boolean
    
    If Not mTDIAddingNewTabForForm Then
        If mTDIMode <> ntTDIModeControls Then
            RaiseError 1381, TypeName(Me), "TDIAddNewTab only available for TDIMode = ntTDIModeControls"
            Exit Function
        End If
        If Not IsMissing(Position) Then
            If Position < 0 Then
                RaiseError 5, TypeName(Me) ' Invalid procedure call or argument
                Exit Function
            ElseIf Position > mTabs Then
                RaiseError 5, TypeName(Me) ' Invalid procedure call or argument
                Exit Function
            End If
        End If
    End If
    mTDILastTabNumber = mTDILastTabNumber + 1
    
    If (nTabType <> ntTDIForm) And (nTabCaption = "") Then
        nTabCaption = "New tab"
    End If
    iLoadTabControls = True
    iShowTabCloseButton = True
    RaiseEvent TDIBeforeNewTab(nTabType, mTDILastTabNumber, nTabCaption, iLoadTabControls, iShowTabCloseButton, iCancel)
    If Not iCancel Then
        TDIPrepareNewTab nTabCaption, iLoadTabControls, IIf(IsMissing(Position), -1, CLng(Position)), Focused, iShowTabCloseButton
        TDIAddNewTab = True
    End If
End Function

Private Sub TDIPrepareNewTab(nTabCaption As String, nLoadTabControls As Boolean, Optional ByVal nPosition As Long = -1, Optional nFocused As Boolean = True, Optional nShowTabCloseButton As Boolean = True)
    Dim iRedraw As Boolean
    
    iRedraw = mRedraw
    mRedraw = False
    mTDIChangingTabCount = True
    Tabs = mTabs + 1
    mTDIChangingTabCount = False
    
    mTDIAddingNewTab = True
    MoveTab mTabs - 2, mTabs - 1
    mTabData(mTabs - 2).TDITabNumber = mTDILastTabNumber
    
    If mTDIMode = ntTDIModeForms Then
        Set TabIconFont(mTabs - 2) = CloneFont(TabIconFont(mTabs - 1))
    Else
        Set TabIconFont(mTabs - 2) = TabIconFont(0)
    End If
    
    If nShowTabCloseButton Then
        If TabIconFont(mTabs - 2).Name = "Arial" Then
            TabIconFont(mTabs - 2).Size = 15
            TabIconLeftOffset(mTabs - 2) = -3 * mDPIScale
            TabIconTopOffset(mTabs - 2) = -1 * mDPIScale
            TabIconCharHex(mTabs - 2) = "&H78&"
        ElseIf TabIconFont(mTabs - 2).Name = "Marlett" Then
            TabIconFont(mTabs - 2).Size = 15
            TabIconLeftOffset(mTabs - 2) = -3 * mDPIScale
            TabIconTopOffset(mTabs - 2) = -1 * mDPIScale
            TabIconCharHex(mTabs - 2) = "&H72&"
        Else
            TabIconFont(mTabs - 2).Size = 10
            TabIconLeftOffset(mTabs - 2) = -3 * mDPIScale
            TabIconTopOffset(mTabs - 2) = 1 * mDPIScale
            TabIconCharHex(mTabs - 2) = "&HE106&"
        End If
    End If
    TabCaption(mTabs - 2) = nTabCaption & IIf(nShowTabCloseButton, "   ", "")
    If mAmbientUserMode Then
        mIconColorMouseHover = mIconColor
        mIconColorMouseHoverSelectedTab = mIconColor
    End If
    tmrTDIIconColor.Enabled = False
    tmrTDIIconColor.Enabled = True
    
    If nPosition = -1 Then
        nPosition = mTabs - 2
    End If
    MoveTab mTabs - 2, CInt(nPosition)
    If nLoadTabControls Then
        TDILoadNewTabControls nPosition
    End If
    
    If nFocused Then
        TabSel = nPosition
    End If
    mTDIAddingNewTab = False
    RaiseEvent TDINewTabAdded(mTDILastTabNumber)
    mTabData(mTabs - 1).Hovered = False
    mRedraw = iRedraw
    DrawDelayed
End Sub

Private Sub TDILoadNewTabControls(ByVal nTabPosition As Long)
    Dim c As Long
    Dim iCtl As Object
    Dim iCtl0 As Object
    Dim iAuxLeft As Long
    Dim iIsLine  As Boolean
    Dim iContainer As Object
    Dim iContainer0 As Object
    
    ' load controls and set position
    For c = 0 To mTDIControlNames_Count - 1
        Set iCtl = UserControl.Parent.Controls(mTDIControlNames(c), mTDILastTabNumber)
        Set iCtl0 = UserControl.Parent.Controls(mTDIControlNames(c), 0) ' same control on first tab
        On Error Resume Next
        Load iCtl
        If Err.Number Then
            MsgBox "Control arrays must have only Index 0! Check " & iCtl.Name & "(" & iCtl.Index & "), vbExclamation"
            Exit Sub
        End If
        Err.Clear
        iAuxLeft = -1.01
        iIsLine = False
        If TypeName(iCtl) = "Line" Then
            iAuxLeft = iCtl.X1
            iIsLine = True
        Else
            iAuxLeft = iCtl.Left
        End If
        On Error GoTo 0
        If iAuxLeft <> -1.01 Then
            On Error Resume Next
            Set iContainer = Nothing
            Set iContainer = iCtl.Container
            On Error GoTo 0
            If iContainer Is UserControl.Extender Then
                If iIsLine Then
                    ControlMove iCtl.Name & "(" & iCtl.Index & ")", ControlLeft(mTDIControlNames(c) & "(0)"), iCtl0.Y1, iCtl0.X2 - iCtl0.X1, iCtl0.Y2 - iCtl0.Y1, CInt(nTabPosition)
                Else
                    ControlMove iCtl.Name & "(" & iCtl.Index & ")", ControlLeft(mTDIControlNames(c) & "(0)"), iCtl0.Top, iCtl0.Width, iCtl0.Height, CInt(nTabPosition)
                End If
            End If
        End If
    Next
    ' set containers
    For c = 0 To mTDIControlNames_Count - 1
        Set iCtl = UserControl.Parent.Controls(mTDIControlNames(c), mTDILastTabNumber)
        Set iCtl0 = UserControl.Parent.Controls(mTDIControlNames(c), 0) ' same control on first tab
        Set iContainer0 = Nothing
        On Error Resume Next
        Set iContainer0 = iCtl0.Container
        On Error GoTo 0
        Set iContainer = Nothing
        If Not iContainer0 Is Nothing Then
            If Not iContainer0 Is UserControl.Extender Then
                Set iContainer = UserControl.Parent.Controls(iContainer0.Name, mTDILastTabNumber)
                Set iCtl.Container = iContainer
            End If
        End If
    Next
    ' set visible
    For c = 0 To mTDIControlNames_Count - 1
        Set iCtl = UserControl.Parent.Controls(mTDIControlNames(c), mTDILastTabNumber)
        Set iCtl0 = UserControl.Parent.Controls(mTDIControlNames(c), 0)
        On Error Resume Next
        iCtl.Visible = iCtl0.Visible
        On Error GoTo 0
    Next
    PropertyChanged False
End Sub

Public Function TDIGetTabIndexByTabNumber(ByVal nTabNumber As Long) As Long
Attribute TDIGetTabIndexByTabNumber.VB_Description = "When in TDI mode, it returns the Index of a tab given its number."
    Dim c As Long
    
    For c = 0 To mTabs - 1
        If mTabData(mTabUnderMouse).TDITabNumber = nTabNumber Then
            TDIGetTabIndexByTabNumber = c
            Exit Function
        End If
    Next
    TDIGetTabIndexByTabNumber -1
End Function

Public Function TDIGetTabNumberByTabIndex(ByVal Index As Long) As Long
Attribute TDIGetTabNumberByTabIndex.VB_Description = "When in TDI mode, it returns the number of a tab given its Index."
    If (Index < 0) Or (Index >= mTabs) Then
        TDIGetTabNumberByTabIndex = -1
        Exit Function
    End If
    TDIGetTabNumberByTabIndex = mTabData(Index).TDITabNumber
End Function

Private Sub TDIUnloadTabControls(ByVal nTabNumber As Long)
    Dim iCtl As Object
    Dim c As Long
    
    For c = 0 To mTDIControlNames_Count - 1
        Set iCtl = UserControl.Parent.Controls(mTDIControlNames(c), nTabNumber)
        On Error Resume Next
        Set iCtl.Container = UserControl.Extender
        On Error GoTo 0
    Next
    For c = 0 To mTDIControlNames_Count - 1
        Set iCtl = UserControl.Parent.Controls(mTDIControlNames(c), nTabNumber)
        Unload iCtl
    Next
End Sub

Public Property Let TabsEndFreeSpace(ByVal nValue As Long)
Attribute TabsEndFreeSpace.VB_Description = "Returns/sets the size of an optional free space after the last tab (for top-otiented tabs that is the rightmost tab)."
Attribute TabsEndFreeSpace.VB_ProcData.VB_Invoke_PropertyPut = ";Tabs"
    If nValue <> mTabsEndFreeSpace Then
        mTabsEndFreeSpace = nValue
        PropertyChanged "TabsEndFreeSpace"
        DrawDelayed
    End If
End Property

Public Property Get TabsEndFreeSpace() As Long
    TabsEndFreeSpace = mTabsEndFreeSpace
End Property

Public Property Get MouseTab() As Long
Attribute MouseTab.VB_Description = "Returns the index of the tab under the mouse. If there is no tab under the mouse it returns -1."
    MouseTab = mTabUnderMouse
End Property

Friend Function IsParentEnabled() As Boolean
    On Error Resume Next
    IsParentEnabled = UserControl.Parent.Enabled
End Function

Friend Sub TDIPutFormIntoTab(ByVal nHwndForm As Long)
    Const WS_CAPTION As Long = &HC00000
    Const WS_THICKFRAME As Long = &H40000
    Const GWL_STYLE As Long = (-16)
    Const WM_GETICON As Long = &H7F
    Const ICON_BIG As Long = 1
    Const WS_SYSMENU = &H80000
    Dim iIconHandle As Long
    Dim iFormCaption As String
    
    mTDIAddingNewTabForForm = True
    iFormCaption = GetWindowCaption(nHwndForm)
    If TDIAddNewTab(mTabs, , ntTDIForm, iFormCaption) Then
        If UBound(mTDIModeFormsFormData_FormHwnd) < (mTabs - 1) Then
            ReDim Preserve mTDIModeFormsFormData_FormHwnd(mTabs + 100)
            ReDim Preserve mTDIModeFormsFormData_OldParentHwnd(mTabs + 100)
            ReDim Preserve mTDIModeFormsFormData_FormIcon(mTabs + 100)
            ReDim Preserve mTDIFormIconCustomData(mTabs + 100)
            ReDim Preserve mTDIFormWithoutCloseButton(mTabs + 100)
        End If
        mTDIModeFormsFormData_FormHwnd(mTabs - 1) = nHwndForm
        mTabData(mTabs - 1).Data = mTabs - 1
        
        Load picTDIFormContainer(mTabs - 1)
        UserControl.ScaleMode = vbTwips
        picTDIFormContainer(mTabs - 1).Move ClientLeft, ClientTop, ClientWidth, ClientHeight
        SetScaleMode
        SetWindowLong nHwndForm, GWL_STYLE, GetWindowLong(nHwndForm, GWL_STYLE) And Not (WS_CAPTION Or WS_THICKFRAME)
        MoveWindow nHwndForm, 0, 0, mClientRect.Right - mClientRect.Left + 2, mClientRect.Bottom - mClientRect.Top + 3, 1
        mTDIFormWithoutCloseButton(mTabs - 1) = (GetWindowLong(nHwndForm, GWL_STYLE) And WS_SYSMENU) = 0
        
        mTDISubclassedFormsHwnds.Add nHwndForm, CStr(nHwndForm)
        AttachMessage Me, nHwndForm, WM_STYLECHANGING
        AttachMessage Me, nHwndForm, WM_DESTROY
        AttachMessage Me, nHwndForm, WM_SHOWWINDOW
        
        mTDIModeFormsFormData_OldParentHwnd(mTabs - 1) = SetParent(nHwndForm, picTDIFormContainer(mTabs - 1).hWnd)
        picTDIFormContainer(mTabs - 1).Visible = True
        
        iIconHandle = SendMessage(nHwndForm, WM_GETICON, ICON_BIG, 0)
        If iIconHandle <> 0 Then
            Set mTDIModeFormsFormData_FormIcon(mTabs - 1) = CreatePicture(iIconHandle, vbPicTypeIcon)
        End If
        
        mTabData(mTabs - 1).Caption = GetUniqueCaption(iFormCaption, mTabs - 1)
        mTDIAddingNewTabForForm = False
        DrawDelayed
    End If
    mTDIAddingNewTabForForm = False
End Sub

Private Function CreatePicture(ByVal hImage As Long, ByVal nType As PictureTypeConstants) As IPictureDisp
    'This function creates a picture object from a handle to a bitmap or a icon
    'hImage is the handle to the bitmap or icon
    'Type is the type of the image (can be either vbPicTypeBitmap or vbPicTypeIcon
    Dim PicInfo As PicBmp
    Dim TmpPic As IPictureDisp
    Dim IID_IDispatch As GUID
 
    'Setup the Guid for the function
    With IID_IDispatch
        
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    
    End With
    
    'Setup the pic structure
    With PicInfo
        .Size = Len(PicInfo)
        .Type = nType
        .hBmp = hImage
    End With
 
    'create the picture
    OleCreatePictureIndirect PicInfo, IID_IDispatch, 1, TmpPic
 
    Set CreatePicture = TmpPic
 
End Function

Private Function GetUniqueCaption(nCaption As String, nTabToIgnore As Long) As String
    Dim c As Long
    Dim iFound As Boolean
    Dim iCaption As String
    Dim n As Long
    Dim n2 As Long
    Dim iStr As String
    Dim iLenStr As String
    Dim iLng As Long
    
    n = 1
    Do
        iFound = False
        If n > 1 Then
            n2 = n
            iStr = nCaption & "("
            iLenStr = Len(iStr)
            For c = 1 To mTabs - 1
                If Left$(mTabData(c).Caption, iLenStr) = iStr Then
                    iLng = Val(Mid$(mTabData(c).Caption, iLenStr + 1)) + 1
                    If iLng > n2 Then
                        n2 = iLng
                    End If
                End If
            Next
            If mRightToLeft Then
                iCaption = "(" & n2 & ") " & nCaption
            Else
                iCaption = nCaption & " (" & n2 & ")"
            End If
        Else
            iCaption = nCaption
        End If
        For c = 1 To mTabs - 1
            If mTabData(c).Visible Then
                If c <> nTabToIgnore Then
                    If mTabData(c).Caption = iCaption Then
                        iFound = True
                        Exit For
                    End If
                End If
            End If
        Next
        If Not iFound Then
            GetUniqueCaption = iCaption
            Exit Function
        End If
        n = n + 1
    Loop
End Function

Private Function GetWindowCaption(ByVal hWnd As Long) As String
    Dim Buffer As String
    Const WM_GETTEXT = &HD
    
    Buffer = String$(GetWindowTextLength(hWnd) + 1, vbNullChar)
    GetWindowCaption = Left$(Buffer, DefWindowProcW(hWnd, WM_GETTEXT, Len(Buffer), StrPtr(Buffer)))
End Function

Friend Sub TDIFormClosing(ByVal nHwndForm As Long)
    Dim c As Long
    Dim i As Long
    Dim iDone As Boolean
    Const WA_ACTIVE As Long = 1
    Dim iHwnd As Long
    
    If mShowingModalForm Then
        PostMessage mFormHwnd, WM_NCACTIVATE, WA_ACTIVE, 0&
        mShowingModalForm = False
    End If
    
    iHwnd = 0
    On Error Resume Next
    iHwnd = mTDISubclassedFormsHwnds(CStr(nHwndForm))
    On Error GoTo 0
    If iHwnd <> 0 Then
        DetachMessage Me, nHwndForm, WM_STYLECHANGING
        DetachMessage Me, nHwndForm, WM_DESTROY
        DetachMessage Me, nHwndForm, WM_SHOWWINDOW
        mTDISubclassedFormsHwnds.Remove CStr(nHwndForm)
    End If
    
    i = TDIGetTabByFormHwnd(nHwndForm)
    If i > -1 Then
        If (mTabData(i).Data >= picTDIFormContainer.LBound) And (mTabData(i).Data <= picTDIFormContainer.UBound) Then
           picTDIFormContainer(mTabData(i).Data).Visible = False
        End If
        mTabData(i).Visible = False
        DrawDelayed
        For c = i - 1 To 1 Step -1
            If mTabData(c).Visible Then
                TabSel = c
                iDone = True
                Exit For
            End If
        Next
        If Not iDone Then
            For c = i + 1 To mTabs - 1
                If mTabData(c).Visible Then
                    TabSel = c
                    iDone = True
                    Exit For
                End If
            Next
        End If
        If Not iDone Then
            If TabVisible(0) = True Then
                TabSel = 0
            End If
        End If
    End If
End Sub

Friend Sub TDIFocusForm(ByVal nHwndForm As Long)
    Dim c As Long
    Const WA_ACTIVE As Long = 1
    
    If mAppDeactivated Then
        PostMessage mFormHwnd, WM_NCACTIVATE, WA_ACTIVE, 0&
        mAppDeactivated = False
    End If
    
    For c = mTabs - 1 To 1 Step -1
        If mTDIModeFormsFormData_FormHwnd(mTabData(c).Data) = nHwndForm Then
            If mTabData(c).Visible Then
                TabSel = c
            End If
            Exit For
        End If
    Next
End Sub

Friend Sub ShowsModalForm()
    Const WA_INACTIVE As Long = 0
    
    PostMessage mFormHwnd, WM_NCACTIVATE, WA_INACTIVE, 0&
    mShowingModalForm = True
End Sub
    
Public Function SelectTab(ByVal Find As Variant, Optional ByVal Method As NTFindTabMethodConstants = ntFindCaption, Optional StartingTab As Variant) As Boolean
Attribute SelectTab.VB_Description = "Sets the selected tab by means some not available in the Tab property."
    Dim t As Long
    
    t = FindTab(Find, Method, StartingTab)
    If t > -1 Then
        TabSel = t
        SelectTab = True
    End If
End Function

Public Function FindTab(ByVal Find As Variant, Optional ByVal Method As NTFindTabMethodConstants = ntFindCaption, Optional StartingTab As Variant) As Long
Attribute FindTab.VB_Description = "Finds a tab based on the values provided in the parameters and returns its index. If the tab is not found, it returns -1."
    Dim c As Long
    Dim iStartingTab As Long
    
    If Not IsMissing(StartingTab) Then
        iStartingTab = Val(StartingTab)
    End If
    FindTab = -1
    If Method = ntFindCaption Then
        Find = LCase$(Find)
        For c = iStartingTab To mTabs - 1
            If LCase$(mTabData(c).Caption) = Find Then
                FindTab = c
                Exit Function
            End If
        Next
    ElseIf Method = ntFindOriginalIndex Then
        For c = iStartingTab To mTabs - 1
            If mTabData(c).OriginalIndex = Find Then
                FindTab = c
                Exit Function
            End If
        Next
    ElseIf Method = ntFindData Then
        For c = iStartingTab To mTabs - 1
            If mTabData(c).Data = Find Then
                FindTab = c
                Exit Function
            End If
        Next
    ElseIf Method = ntFindTag Then
        For c = iStartingTab To mTabs - 1
            If mTabData(c).Tag = Find Then
                FindTab = c
                Exit Function
            End If
        Next
    ElseIf Method = ntFindPartialCaption Then
        For c = iStartingTab To mTabs - 1
            If InStr(1, mTabData(c).Caption, Find, vbTextCompare) > 0 Then
                FindTab = c
                Exit Function
            End If
        Next
    ElseIf Method = ntFindPartialCaptionReverse Then
        If IsMissing(StartingTab) Then
            iStartingTab = mTabs - 1
        End If
        For c = iStartingTab To 0 Step -1
            If InStr(1, mTabData(c).Caption, Find, vbTextCompare) > 0 Then
                FindTab = c
                Exit Function
            End If
        Next
    ElseIf Method = ntFindKey Then
        For c = iStartingTab To mTabs - 1
            Dim iLCFind As String
            
            iLCFind = LCase(Find)
            If LCase$(mTabData(c).Key) = iLCFind Then
                FindTab = c
                Exit Function
            End If
        Next
    End If
End Function


Public Function GetTabsOrder() As String
Attribute GetTabsOrder.VB_Description = "Returns a string with data containing the tabs order. CanReorderTabs must be set to True."
    Dim c As Long
    
    If CanReorderTabsEffective = False Then
        RaiseError 1381, TypeName(Me), "GetTabsOrder/SetTabsOrder need CanReorderTabs property to True."
        Exit Function
    End If
    
    For c = 0 To mTabs - 1
        If c > 0 Then GetTabsOrder = GetTabsOrder & ","
        GetTabsOrder = GetTabsOrder & mTabData(c).OriginalIndex
    Next
End Function


Public Sub SetTabsOrder(ByVal PosData As String)
Attribute SetTabsOrder.VB_Description = "Sets the tabs order using data previously obtained with GetTabsOrder. CanReorderTabs must be set to True."
    Dim iData() As String
    Dim c As Long
    Dim iAuxTabData() As T_TabData
    Dim iSet() As Boolean
    Dim oi As Long
    
    If CanReorderTabsEffective = False Then
        RaiseError 1381, TypeName(Me), "GetTabsOrder/SetTabsOrder need CanReorderTabs property to True."
        Exit Sub
    End If
    
    ReDim iAuxTabData(mTabs - 1)
    ReDim iSet(mTabs - 1)
    
    iData = Split(PosData, ",")
    
    For c = 0 To UBound(iData)
        If c < mTabs Then
            oi = FindTab(Val(iData(c)), ntFindOriginalIndex)
            If oi > -1 Then
                iAuxTabData(c) = mTabData(oi)
                iSet(c) = True
            End If
        End If
    Next
    
    For c = 0 To mTabs - 1
        If iSet(c) Then
            mTabData(c) = iAuxTabData(c)
        End If
    Next
    
    DrawDelayed
End Sub

Public Function GetThemeData() As String
Attribute GetThemeData.VB_Description = "Returns a string with the current theme data. It can be later applied with SetThemeData."
    GetThemeData = GetThemeStringFromControl(Me, Ambient.BackColor, Ambient.ForeColor)
End Function


Public Sub SetThemeData(ByVal nThemeData As String)
Attribute SetThemeData.VB_Description = "Applies a theme from data contained in a String. That string can be obtained with GetThemeData."
    Dim iTheme As NewTabTheme
    
    Set iTheme = New NewTabTheme
    iTheme.ThemeString = nThemeData
    ApplyThemeToControl iTheme.Data, Me, Ambient.BackColor, Ambient.ForeColor
End Sub


Public Property Get TabFixedWidth(ByVal Index As Long) As Long
Attribute TabFixedWidth.VB_Description = "Returns/sets a fixed width for a tab. To set it for its automatic width (according to the TabWidthStyle setting), set it to 0."
Attribute TabFixedWidth.VB_ProcData.VB_Invoke_Property = ";Tabs"
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 381, TypeName(Me) ' invalid property array index
        Exit Property
    End If
    TabFixedWidth = UserControl.ScaleX(mTabData(Index).FixedWidth, vbPixels, vbTwips)
End Property

Public Property Let TabFixedWidth(ByVal Index As Long, ByVal nValue As Long)
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 381, TypeName(Me) ' invalid property array index
        Exit Property
    End If
    If nValue <> mTabData(Index).FixedWidth Then
        If (nValue < 0) Then
            RaiseError 380, TypeName(Me) ' invalid property value
            Exit Property
        End If
        mTabData(Index).FixedWidth = UserControl.ScaleX(nValue, vbTwips, vbPixels)
        PropertyChanged "TabFixedWidth"
        DrawDelayed
    End If
End Property

Private Function ColorsHaveEnoughContrast(ByVal nColor1 As Long, ByVal nColor2 As Long, Optional ByVal Threshold As Long = 120) As Boolean
    Dim h1 As Integer
    Dim l1 As Integer
    Dim s1 As Integer
    Dim h2 As Integer
    Dim l2 As Integer
    Dim s2 As Integer
    Dim c1 As Long
    Dim c2 As Long
    
    TranslateColor nColor1, 0&, c1
    TranslateColor nColor2, 0&, c2
    
    ColorRGBToHLS c1, h1, l1, s1
    ColorRGBToHLS c2, h2, l2, s2
    ColorsHaveEnoughContrast = Abs(l1 - l2) > Threshold
End Function


Public Property Get TabCustomColor(ByVal Index As Long, Optional ByVal ColorID As NTTabCustomColorIDConstants = ntCCBackColorTab) As OLE_COLOR
Attribute TabCustomColor.VB_Description = "Returns/sets a custom color for a tab. It may be a color when it is inactive or active."
Attribute TabCustomColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 381, TypeName(Me) ' invalid property array index
        Exit Property
    End If
    If (ColorID < 0) Or (ColorID > ntCCIconColorTabHighlighted) Then
        RaiseError 1382, TypeName(Me), "invalid ColorID."
        Exit Property
    End If
    
    Select Case ColorID
        Case ntCCBackColorTab
            If IsEmpty(mTabData(Index).CustomColors.BackColorTab) Then
                TabCustomColor = mBackColorTabs
            Else
                TabCustomColor = mTabData(Index).CustomColors.BackColorTab
            End If
        Case ntCCBackColorSelectedTab
            If IsEmpty(mTabData(Index).CustomColors.BackColorSelectedTab) Then
                TabCustomColor = mBackColorSelectedTab
            Else
                TabCustomColor = mTabData(Index).CustomColors.BackColorSelectedTab
            End If
        Case ntCCHighlightColor
            If IsEmpty(mTabData(Index).CustomColors.HighlightColor) Then
                TabCustomColor = mHighlightColor
            Else
                TabCustomColor = mTabData(Index).CustomColors.HighlightColor
            End If
        Case ntCCHighlightColorSelectedTab
            If IsEmpty(mTabData(Index).CustomColors.HighlightColorSelectedTab) Then
                TabCustomColor = mHighlightColorSelectedTab
            Else
                TabCustomColor = mTabData(Index).CustomColors.HighlightColorSelectedTab
            End If
        Case ntCCFlatBarColorInactive
            If IsEmpty(mTabData(Index).CustomColors.FlatBarColorInactive) Then
                TabCustomColor = mFlatBarColorInactive
            Else
                TabCustomColor = mTabData(Index).CustomColors.FlatBarColorInactive
            End If
        Case ntCCFlatBarColorHighlight
            If IsEmpty(mTabData(Index).CustomColors.FlatBarColorHighlight) Then
                TabCustomColor = mFlatBarColorHighlight
            Else
                TabCustomColor = mTabData(Index).CustomColors.FlatBarColorHighlight
            End If
        Case ntCCFlatBarColorSelectedTab
            If IsEmpty(mTabData(Index).CustomColors.FlatBarColorSelectedTab) Then
                TabCustomColor = mFlatBarColorSelectedTab
            Else
                TabCustomColor = mTabData(Index).CustomColors.FlatBarColorSelectedTab
            End If
        Case ntCCFlatTabBorderColorHighlight
            If IsEmpty(mTabData(Index).CustomColors.FlatTabBorderColorHighlight) Then
                TabCustomColor = mFlatTabBorderColorHighlight
            Else
                TabCustomColor = mTabData(Index).CustomColors.FlatTabBorderColorHighlight
            End If
        Case ntCCFlatTabBorderColorSelectedTab
            If IsEmpty(mTabData(Index).CustomColors.FlatTabBorderColorSelectedTab) Then
                TabCustomColor = mFlatTabBorderColorSelectedTab
            Else
                TabCustomColor = mTabData(Index).CustomColors.FlatTabBorderColorSelectedTab
            End If
        Case ntCCForeColor
            If IsEmpty(mTabData(Index).CustomColors.ForeColor) Then
                TabCustomColor = mForeColor
            Else
                TabCustomColor = mTabData(Index).CustomColors.ForeColor
            End If
        Case ntCCForeColorHighlighted
            If IsEmpty(mTabData(Index).CustomColors.ForeColorHighlighted) Then
                TabCustomColor = mForeColorHighlighted
            Else
                TabCustomColor = mTabData(Index).CustomColors.ForeColorHighlighted
            End If
        Case ntCCForeColorSelectedTab
            If IsEmpty(mTabData(Index).CustomColors.ForeColorSelectedTab) Then
                TabCustomColor = mForeColorSelectedTab
            Else
                TabCustomColor = mTabData(Index).CustomColors.ForeColorSelectedTab
            End If
        Case ntCCIconColor
            If IsEmpty(mTabData(Index).CustomColors.IconColor) Then
                TabCustomColor = mIconColor
            Else
                TabCustomColor = mTabData(Index).CustomColors.IconColor
            End If
        Case ntCCIconColorSelectedTab
            If IsEmpty(mTabData(Index).CustomColors.IconColorSelectedTab) Then
                TabCustomColor = mIconColorSelectedTab
            Else
                TabCustomColor = mTabData(Index).CustomColors.IconColorSelectedTab
            End If
        Case ntCCIconColorMouseHover
            If IsEmpty(mTabData(Index).CustomColors.IconColorMouseHover) Then
                TabCustomColor = mIconColorMouseHover
            Else
                TabCustomColor = mTabData(Index).CustomColors.IconColorMouseHover
            End If
        Case ntCCIconColorMouseHoverSelectedTab
            If IsEmpty(mTabData(Index).CustomColors.IconColorMouseHoverSelectedTab) Then
                TabCustomColor = mIconColorMouseHoverSelectedTab
            Else
                TabCustomColor = mTabData(Index).CustomColors.IconColorMouseHoverSelectedTab
            End If
        Case ntCCIconColorTabHighlighted
            If IsEmpty(mTabData(Index).CustomColors.IconColorTabHighlighted) Then
                TabCustomColor = mIconColorTabHighlighted
            Else
                TabCustomColor = mTabData(Index).CustomColors.IconColorTabHighlighted
            End If
    End Select
End Property

Public Property Let TabCustomColor(ByVal Index As Long, Optional ByVal ColorID As NTTabCustomColorIDConstants = ntCCBackColorTab, ByVal nValue As OLE_COLOR)
    Dim c As Long
    Dim iCol As Long
    
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 381, TypeName(Me) ' invalid property array index
        Exit Property
    End If
    If (ColorID < 0) Or (ColorID > ntCCIconColorTabHighlighted) Then
        RaiseError 1382, TypeName(Me), "invalid ColorID."
        Exit Property
    End If
    If nValue <> -1 Then
        If Not IsValidOLE_COLOR(nValue) Then RaiseError 380, TypeName(Me): Exit Property
    End If
    
    Select Case ColorID
        Case ntCCBackColorTab
            If (nValue = mBackColorTabs) Or (nValue = -1) Then
                mTabData(Index).CustomColors.BackColorTab = Empty
            Else
                mTabData(Index).CustomColors.BackColorTab = nValue
            
                Dim iBCol As Long
                Dim iBackColorTabs_H As Integer
                Dim iBackColorTabs_L As Integer
                Dim iBackColorTabs_S As Integer
                Dim iGlowColor As Long
                
                iBCol = TranslatedColor(nValue)
                ColorRGBToHLS iBCol, iBackColorTabs_H, iBackColorTabs_L, iBackColorTabs_S
                iGlowColor = GetGlowColor(iBackColorTabs_H, iBackColorTabs_L, iBackColorTabs_S)
                For c = 1 To 10
                    mTabData(Index).CustomColors.HighlightEffectColors_Strong(c) = MixColors(iGlowColor, nValue, 10 * c)
                    mTabData(Index).CustomColors.HighlightEffectColors_Light(c) = MixColors(iGlowColor, nValue, 6 * c)
                Next c
            End If
        Case ntCCBackColorSelectedTab
            If (nValue = mBackColorSelectedTab) Or (nValue = -1) Then
                mTabData(Index).CustomColors.BackColorSelectedTab = Empty
                iCol = mBackColorSelectedTab
            Else
                mTabData(Index).CustomColors.BackColorSelectedTab = nValue
                iCol = nValue
            End If
            SetColors
            If Index = mTab Then SetControlsBackColor iCol
        Case ntCCHighlightColor
            If (nValue = mHighlightColor) Or (nValue = -1) Then
                mTabData(Index).CustomColors.HighlightColor = Empty
            Else
                mTabData(Index).CustomColors.HighlightColor = nValue
            End If
        Case ntCCHighlightColorSelectedTab
            If (nValue = mHighlightColorSelectedTab) Or (nValue = -1) Then
                mTabData(Index).CustomColors.HighlightColorSelectedTab = Empty
            Else
                mTabData(Index).CustomColors.HighlightColorSelectedTab = nValue
            End If
            SetColors
        Case ntCCFlatBarColorInactive
            If (nValue = mFlatBarColorInactive) Or (nValue = -1) Then
                mTabData(Index).CustomColors.FlatBarColorInactive = Empty
            Else
                mTabData(Index).CustomColors.FlatBarColorInactive = nValue
                Dim iFlatBarColorHighlight As Long
                
                If IsEmpty(mTabData(Index).CustomColors.FlatBarColorHighlight) Then
                    iFlatBarColorHighlight = mFlatBarColorHighlight
                Else
                    iFlatBarColorHighlight = mTabData(Index).CustomColors.FlatBarColorHighlight
                End If
                For c = 1 To 10
                    mTabData(Index).CustomColors.FlatBarHighlightEffectColors(c) = MixColors(iFlatBarColorHighlight, mTabData(Index).CustomColors.FlatBarColorInactive, 10 * c)
                Next c
                mTabData(Index).CustomColors.FlatBarGlowColor = mTabData(Index).CustomColors.FlatBarHighlightEffectColors(10)
            End If
        Case ntCCFlatBarColorHighlight
            If (nValue = mFlatBarColorHighlight) Or (nValue = -1) Then
                mTabData(Index).CustomColors.FlatBarColorHighlight = Empty
            Else
                mTabData(Index).CustomColors.FlatBarColorHighlight = nValue
                
                Dim iFlatBarColorInactive As Long
                
                If IsEmpty(mTabData(Index).CustomColors.FlatBarColorInactive) Then
                    iFlatBarColorInactive = mFlatBarColorInactive
                Else
                    iFlatBarColorInactive = mTabData(Index).CustomColors.FlatBarColorInactive
                End If
                For c = 1 To 10
                    mTabData(Index).CustomColors.FlatBarHighlightEffectColors(c) = MixColors(mTabData(Index).CustomColors.FlatBarColorHighlight, iFlatBarColorInactive, 10 * c)
                Next c
                mTabData(Index).CustomColors.FlatBarGlowColor = mTabData(Index).CustomColors.FlatBarHighlightEffectColors(10)
            End If
        Case ntCCFlatBarColorSelectedTab
            If (nValue = mFlatBarColorSelectedTab) Or (nValue = -1) Then
                mTabData(Index).CustomColors.FlatBarColorSelectedTab = Empty
            Else
                mTabData(Index).CustomColors.FlatBarColorSelectedTab = nValue
            End If
            SetColors
        Case ntCCFlatTabBorderColorHighlight
            If (nValue = mFlatTabBorderColorHighlight) Or (nValue = -1) Then
                mTabData(Index).CustomColors.FlatTabBorderColorHighlight = Empty
            Else
                mTabData(Index).CustomColors.FlatTabBorderColorHighlight = nValue
            End If
        Case ntCCFlatTabBorderColorSelectedTab
            If (nValue = mFlatTabBorderColorSelectedTab) Or (nValue = -1) Then
                mTabData(Index).CustomColors.FlatTabBorderColorSelectedTab = Empty
            Else
                mTabData(Index).CustomColors.FlatTabBorderColorSelectedTab = nValue
            End If
        Case ntCCForeColor
            If (nValue = mForeColor) Or (nValue = -1) Then
                mTabData(Index).CustomColors.ForeColor = Empty
            Else
                mTabData(Index).CustomColors.ForeColor = nValue
            End If
        Case ntCCForeColorHighlighted
            If (nValue = mForeColorHighlighted) Or (nValue = -1) Then
                mTabData(Index).CustomColors.ForeColorHighlighted = Empty
            Else
                mTabData(Index).CustomColors.ForeColorHighlighted = nValue
            End If
        Case ntCCForeColorSelectedTab
            If (nValue = mForeColorSelectedTab) Or (nValue = -1) Then
                mTabData(Index).CustomColors.ForeColorSelectedTab = Empty
            Else
                mTabData(Index).CustomColors.ForeColorSelectedTab = nValue
            End If
            SetControlsProperForeColor
        Case ntCCIconColor
            If (nValue = mIconColor) Or (nValue = -1) Then
                mTabData(Index).CustomColors.IconColor = Empty
            Else
                mTabData(Index).CustomColors.IconColor = nValue
            End If
        Case ntCCIconColorSelectedTab
            If (nValue = mIconColorSelectedTab) Or (nValue = -1) Then
                mTabData(Index).CustomColors.IconColorSelectedTab = Empty
            Else
                mTabData(Index).CustomColors.IconColorSelectedTab = nValue
            End If
        Case ntCCIconColorMouseHover
            If (nValue = mIconColorMouseHover) Or (nValue = -1) Then
                mTabData(Index).CustomColors.IconColorMouseHover = Empty
            Else
                mTabData(Index).CustomColors.IconColorMouseHover = nValue
            End If
        Case ntCCIconColorMouseHoverSelectedTab
            If (nValue = mIconColorMouseHoverSelectedTab) Or (nValue = -1) Then
                mTabData(Index).CustomColors.IconColorMouseHoverSelectedTab = Empty
            Else
                mTabData(Index).CustomColors.IconColorMouseHoverSelectedTab = nValue
            End If
        Case ntCCIconColorTabHighlighted
            If (nValue = mIconColorTabHighlighted) Or (nValue = -1) Then
                mTabData(Index).CustomColors.IconColorTabHighlighted = Empty
            Else
                mTabData(Index).CustomColors.IconColorTabHighlighted = nValue
            End If
    End Select
    PropertyChanged "TabCustomColor"
    DrawDelayed
End Property

Private Property Get CanReorderTabsEffective() As Boolean
    CanReorderTabsEffective = mCanReorderTabs And ((mTabOrientation = ssTabOrientationTop) Or (mTabOrientation = ssTabOrientationBottom))
End Property


Public Property Get CaptionAlignment() As NTAlignmentConstants
Attribute CaptionAlignment.VB_Description = "Returns/sets the alignment of the tab captions."
Attribute CaptionAlignment.VB_ProcData.VB_Invoke_Property = ";Appearance"
    CaptionAlignment = mCaptionAlignment
End Property

Public Property Let CaptionAlignment(ByVal nValue As NTAlignmentConstants)
    If nValue <> mCaptionAlignment Then
        If (nValue < ntAlignmentLeft) Or (nValue > ntAlignmentCenter) Then
            RaiseError 380, TypeName(Me) ' invalid property value
            Exit Property
        End If
        mCaptionAlignment = nValue
        If mCaptionAlignment = ntAlignmentLeft Then
            mAlignmentDTFlag = DT_LEFT
        ElseIf mCaptionAlignment = ntAlignmentRight Then
            mAlignmentDTFlag = DT_RIGHT
        Else
            mAlignmentDTFlag = DT_CENTER
        End If
        DrawDelayed
        mSetAutoTabHeightPending = True
        SetPropertyChanged "CaptionAlignment"
    End If
End Property

Public Function TDIGetFormHwndByTab(ByVal Index As Long) As Long
Attribute TDIGetFormHwndByTab.VB_Description = "When in TDI mode forms, it returns the hWnd of the form that is in that tab index."
    On Error Resume Next
    TDIGetFormHwndByTab = mTDIModeFormsFormData_FormHwnd(mTabData(Index).Data)
End Function

Public Function TDIGetTabByFormHwnd(ByVal FormhWnd As Long) As Long
Attribute TDIGetTabByFormHwnd.VB_Description = "When in TDI mode forms, it returns the tab index by supplying the hWnd of the form."
    Dim c As Long
    
    TDIGetTabByFormHwnd = -1
    For c = 0 To mTabs - 1
        If mTDIModeFormsFormData_FormHwnd(mTabData(c).Data) = FormhWnd Then
            TDIGetTabByFormHwnd = c
            Exit For
        End If
    Next
End Function

Friend Property Let IsTDIFormsSampleInPropertyPage(ByVal nValue As Boolean)
    mIsTDIFormsSampleInPropertyPage = nValue
End Property

Public Function GetTabOriginalIndex(ByVal Index As Long) As Long
Attribute GetTabOriginalIndex.VB_Description = "It returns the original index of the tab by supplying the current index. If CanReorderTabs is set to True, the tab index can change."
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 381, TypeName(Me) ' invalid property array index
        Exit Function
    End If
    GetTabOriginalIndex = mTabData(Index).OriginalIndex
End Function

Private Function IsNumericInteger(nValue As Variant) As Boolean
    Dim iVal As Variant
    
    IsNumericInteger = IsNumeric(nValue)
    If IsNumericInteger Then
        iVal = Val(nValue)
        IsNumericInteger = (iVal = Int(iVal))
    End If
End Function

Public Sub RefreshContainedControls()
Attribute RefreshContainedControls.VB_Description = "Forces the redraw of some controls that can be problematic because they don't draw correctly after changing tab. Invoke it in the Click event of the control (after changing tab)."
    Dim iCtl As Object
    
    On Error Resume Next
    For Each iCtl In TabControls(mTab)
        iCtl.Width = iCtl.Width + Screen.TwipsPerPixelX
        iCtl.Width = iCtl.Width - Screen.TwipsPerPixelX
        iCtl.Refresh
    Next
End Sub

Friend Property Get Subclassed() As Boolean
    Subclassed = mSubclassed
End Property

'Tab is a reserved keyword in VB6, but you can remove that restriction.
'To be able to compile with Tab property, you need to replace VBA6.DLL with this version: https://github.com/EduardoVB/NewTab/raw/main/control-source/lib/VBA6.DLL
'VBA6.DLL is in VS6's installation folder, usually:
'C:\Program Files (x86)\Microsoft Visual Studio\VB98\

#Const COMPILE_WITH_TAB_PROPERTY = 0
#If COMPILE_WITH_TAB_PROPERTY Then
Public Property Get Tab() As Variant
Attribute Tab.VB_Description = "Returns or sets the selected tab by its index, TabKey or TabCaption."
Attribute Tab.VB_ProcData.VB_Invoke_Property = ";Tabs"

'Attribute Tab.VB_Description = "Returns or sets the index of the current (""selected"" or ""active"") tab."
    Tab = mTab
End Property

Public Property Let Tab(ByVal nValue As Variant)
    Dim iTab As Long
    
    If IsNumericInteger(nValue) Then
        ' OK, it is an index
    ElseIf VarType(nValue) = vbString Then
        iTab = FindTab(nValue, ntFindKey)
        If iTab = -1 Then
            iTab = FindTab(nValue, ntFindCaption)
            If iTab = -1 Then
                RaiseError 5, TypeName(Me), "Invalid procedure call or argument. String not found as TabKey or Caption. Or nValue must be numeric integer if it is a tab index."   ' Invalid procedure call or argument
                Exit Property
            Else
                nValue = iTab
            End If
        Else
            nValue = iTab
        End If
    Else
        RaiseError 380, TypeName(Me), "Invalid property value. nValue must be a tab index or a TabKey."
        Exit Property
    End If
    TabSel = nValue
End Property
#Else
Public Property Get SelectedTab() As Variant
Attribute SelectedTab.VB_Description = "Returns or sets the selected tab by its index, TabKey or TabCaption."
Attribute SelectedTab.VB_ProcData.VB_Invoke_Property = ";Tabs"

'Attribute Tab.VB_Description = "Returns or sets the index of the current (""selected"" or ""active"") tab."
    SelectedTab = mTab
End Property

Public Property Let SelectedTab(ByVal nValue As Variant)
    Dim iTab As Long
    
    If IsNumericInteger(nValue) Then
        ' OK, it is an index
    ElseIf VarType(nValue) = vbString Then
        iTab = FindTab(nValue, ntFindKey)
        If iTab = -1 Then
            iTab = FindTab(nValue, ntFindCaption)
            If iTab = -1 Then
                RaiseError 5, TypeName(Me), "Invalid procedure call or argument. String not found as TabKey or Caption. Or nValue must be numeric integer if it is a tab index."   ' Invalid procedure call or argument
                Exit Property
            Else
                nValue = iTab
            End If
        Else
            nValue = iTab
        End If
    Else
        RaiseError 380, TypeName(Me), "Invalid property value. nValue must be a tab index or a TabKey."
        Exit Property
    End If
    TabSel = nValue
End Property
#End If

