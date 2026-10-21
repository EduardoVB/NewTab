VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cColorSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' The color choice the property pages offer: automatic, a system color or any other one.
' A popup menu and not a dropdown window: TPM_RETURNCMD hands the choice back in the same call,
' and a property page can neither own a VB menu nor keep a window of its own above the dialog

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type TPMPARAMS
    cbSize As Long
    rcExclude As RECT
End Type

Private Type MENUITEMINFOW
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As LongPtr
    hbmpChecked As LongPtr
    hbmpUnchecked As LongPtr
    dwItemData As LongPtr
    dwTypeData As LongPtr
    cch As Long
    hbmpItem As LongPtr
End Type

Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Declare Function CreatePopupMenu Lib "user32" () As LongPtr
Private Declare Function AppendMenuW Lib "user32" (ByVal hMenu As LongPtr, ByVal uFlags As Long, ByVal uIDNewItem As LongPtr, ByVal lpNewItem As LongPtr) As Long
Private Declare Function TrackPopupMenuEx Lib "user32" (ByVal hMenu As LongPtr, ByVal uFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal hWnd As LongPtr, ByVal lptpm As LongPtr) As Long
Private Declare Function DestroyMenu Lib "user32" (ByVal hMenu As LongPtr) As Long
Private Declare Function SetMenuItemInfoW Lib "user32" (ByVal hMenu As LongPtr, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As MENUITEMINFOW) As Long
Private Declare Function SetMenuDefaultItem Lib "user32" (ByVal hMenu As LongPtr, ByVal uItem As Long, ByVal fByPos As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As RECT) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hDC As LongPtr) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As LongPtr, pbmi As BITMAPINFOHEADER, ByVal iUsage As Long, ppvBits As LongPtr, ByVal hSection As LongPtr, ByVal dwOffset As Long) As LongPtr
Private Declare Function SetDIBits Lib "gdi32" (ByVal hDC As LongPtr, ByVal hBitmap As LongPtr, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFOHEADER, ByVal wUsage As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
Private Declare Function TranslateColor Lib "oleaut32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As LongPtr, Col As Long) As Long

Private Const MF_STRING = &H0&
Private Const MF_POPUP = &H10&
Private Const MF_SEPARATOR = &H800&
Private Const MIIM_BITMAP = &H80&
Private Const TPM_LEFTALIGN = &H0&
Private Const TPM_TOPALIGN = &H0&
Private Const TPM_VERTICAL = &H40&
Private Const TPM_NONOTIFY = &H80&
Private Const TPM_RETURNCMD = &H100&
Private Const SM_CXMENUCHECK = 71&
Private Const SM_CYMENUCHECK = 72&
Private Const DIB_RGB_COLORS = 0&
Private Const BI_RGB = 0&

Private Const cCmdAutomatic As Long = 1
Private Const cCmdOther As Long = 2
Private Const cCmdSystemBase As Long = 100
Private Const cSystemColorCount As Long = 25

' nCurrentColor and nPickedColor use -1 for automatic, as the control's properties do.
' False when the menu was dismissed or the color dialog canceled. nAllowAutomatic is False for the
' properties that have no automatic state
Public Function PickColor(ByVal nOwnerHwnd As LongPtr, ByVal nAnchorHwnd As LongPtr, ByVal nCurrentColor As Long, ByVal nAutomaticColor As Long, ByRef nPickedColor As Long, Optional ByVal nAllowAutomatic As Boolean = True) As Boolean
    Dim iMenu As LongPtr
    Dim iSubMenu As LongPtr
    Dim iBitmaps(0 To cSystemColorCount) As LongPtr
    Dim iSwatchWidth As Long
    Dim iSwatchHeight As Long
    Dim iCmd As Long
    Dim iIdx As Long
    Dim utParams As TPMPARAMS
    Dim iDlg As cCommonDialog
    
    iMenu = CreatePopupMenu
    If iMenu = 0 Then Exit Function
    iSubMenu = CreatePopupMenu
    If iSubMenu = 0 Then
        DestroyMenu iMenu
        Exit Function
    End If
    iSwatchWidth = GetSystemMetrics(SM_CXMENUCHECK)
    iSwatchHeight = GetSystemMetrics(SM_CYMENUCHECK)
    If iSwatchWidth < 8 Then iSwatchWidth = 16
    If iSwatchHeight < 8 Then iSwatchHeight = 16
    
    For iIdx = 0 To cSystemColorCount - 1
        AppendMenuW iSubMenu, MF_STRING, cCmdSystemBase + iIdx, StrPtr(SystemColorName(iIdx))
        iBitmaps(iIdx) = SwatchBitmap(&H80000000 Or iIdx, iSwatchWidth, iSwatchHeight)
        SetItemBitmap iSubMenu, cCmdSystemBase + iIdx, iBitmaps(iIdx)
    Next iIdx
    
    If nAllowAutomatic Then
        AppendMenuW iMenu, MF_STRING, cCmdAutomatic, StrPtr("Automatic")
        iBitmaps(cSystemColorCount) = SwatchBitmap(nAutomaticColor, iSwatchWidth, iSwatchHeight)
        SetItemBitmap iMenu, cCmdAutomatic, iBitmaps(cSystemColorCount)
        AppendMenuW iMenu, MF_SEPARATOR, 0, 0
    End If
    AppendMenuW iMenu, MF_POPUP, iSubMenu, StrPtr("System colors")
    AppendMenuW iMenu, MF_STRING, cCmdOther, StrPtr("Other color...")
    
    ' With a bitmap in hbmpItem the system draws no check mark, so the current choice is the
    ' default item, drawn in bold
    If nCurrentColor = -1 Then
        SetMenuDefaultItem iMenu, cCmdAutomatic, 0
    ElseIf (nCurrentColor And &HFFFFFF00) = &H80000000 Then
        iIdx = nCurrentColor And &HFF&
        If iIdx < cSystemColorCount Then
            SetMenuDefaultItem iSubMenu, cCmdSystemBase + iIdx, 0
        End If
    Else
        SetMenuDefaultItem iMenu, cCmdOther, 0
    End If
    
    ' Below the anchor, or above it when there is no room: the exclusion rect is what lets the
    ' system choose
    GetWindowRect nAnchorHwnd, utParams.rcExclude
    utParams.cbSize = LenB(utParams)
    iCmd = TrackPopupMenuEx(iMenu, TPM_RETURNCMD Or TPM_NONOTIFY Or TPM_LEFTALIGN Or TPM_TOPALIGN Or TPM_VERTICAL, utParams.rcExclude.Left, utParams.rcExclude.Bottom, nOwnerHwnd, VarPtr(utParams))
    
    ' the submenu goes with its parent
    DestroyMenu iMenu
    For iIdx = 0 To cSystemColorCount
        If iBitmaps(iIdx) <> 0 Then DeleteObject iBitmaps(iIdx)
    Next iIdx
    
    Select Case iCmd
        Case cCmdAutomatic
            nPickedColor = -1
            PickColor = True
        Case cCmdOther
            Set iDlg = New cCommonDialog
            ' The dialog opens on the colour the swatch shows
            If nCurrentColor = -1 Then
                iDlg.Color = nAutomaticColor
            Else
                iDlg.Color = nCurrentColor
            End If
            If (nCurrentColor And &HFF000000) = 0 Then
                iDlg.ShowColor cdeCCFullOpen
            Else
                iDlg.ShowColor
            End If
            If Not iDlg.Canceled Then
                nPickedColor = iDlg.Color
                PickColor = True
            End If
        Case cCmdSystemBase To cCmdSystemBase + cSystemColorCount - 1
            nPickedColor = &H80000000 Or (iCmd - cCmdSystemBase)
            PickColor = True
    End Select
End Function

' The name the page shows for a color: "Automatic", the system color's name, or its hex value
Public Function ColorDisplayName(ByVal nColor As Long) As String
    Dim iIdx As Long
    
    If nColor = -1 Then
        ColorDisplayName = "Automatic"
    ElseIf (nColor And &HFFFFFF00) = &H80000000 Then
        iIdx = nColor And &HFF&
        If iIdx < cSystemColorCount Then
            ColorDisplayName = SystemColorName(iIdx)
        Else
            ColorDisplayName = "&H" & Hex$(nColor) & "&"
        End If
    Else
        ColorDisplayName = "&H" & Right$("00000000" & Hex$(nColor), 8) & "&"
    End If
End Function

' The names VB's own color list gives them
Private Function SystemColorName(ByVal nIndex As Long) As String
    Dim iNames() As String
    
    iNames = Split("Scroll Bars|Desktop|Active Title Bar|Inactive Title Bar|Menu Bar|Window Background|Window Frame|Menu Text|Window Text|Active Title Bar Text|Active Border|Inactive Border|Application Workspace|Highlight|Highlight Text|Button Face|Button Shadow|Disabled Text|Button Text|Inactive Title Bar Text|Button Highlight|Button Dark Shadow|Button Light Shadow|ToolTip Text|ToolTip", "|")
    SystemColorName = iNames(nIndex)
End Function

Private Sub SetItemBitmap(ByVal nMenu As LongPtr, ByVal nCmd As Long, ByVal nBitmap As LongPtr)
    Dim utItem As MENUITEMINFOW
    
    If nBitmap = 0 Then Exit Sub
    utItem.cbSize = LenB(utItem)
    utItem.fMask = MIIM_BITMAP
    utItem.hbmpItem = nBitmap
    SetMenuItemInfoW nMenu, nCmd, 0, utItem
End Sub

' 32 bpp with the alpha byte set: the themed menu reads a zero alpha as transparent
Private Function SwatchBitmap(ByVal nColor As Long, ByVal nWidth As Long, ByVal nHeight As Long) As LongPtr
    Dim utHeader As BITMAPINFOHEADER
    Dim iPixels() As Long
    Dim iRGB As Long
    Dim iFill As Long
    Dim iFrame As Long
    Dim iX As Long
    Dim iY As Long
    Dim iDC As LongPtr
    Dim iBits As LongPtr
    Dim iBmp As LongPtr
    
    If TranslateColor(nColor, 0, iRGB) <> 0 Then Exit Function
    iFill = &HFF000000 Or ((iRGB And &HFF&) * &H10000) Or (iRGB And &HFF00&) Or ((iRGB \ &H10000) And &HFF&)
    iFrame = &HFF808080
    ReDim iPixels(0 To nWidth * nHeight - 1)
    For iY = 0 To nHeight - 1
        For iX = 0 To nWidth - 1
            If (iX = 0) Or (iY = 0) Or (iX = nWidth - 1) Or (iY = nHeight - 1) Then
                iPixels(iY * nWidth + iX) = iFrame
            Else
                iPixels(iY * nWidth + iX) = iFill
            End If
        Next iX
    Next iY
    utHeader.biSize = LenB(utHeader)
    utHeader.biWidth = nWidth
    utHeader.biHeight = -nHeight
    utHeader.biPlanes = 1
    utHeader.biBitCount = 32
    utHeader.biCompression = BI_RGB
    iDC = GetDC(0)
    iBmp = CreateDIBSection(iDC, utHeader, DIB_RGB_COLORS, iBits, 0, 0)
    If iBmp <> 0 Then
        SetDIBits iDC, iBmp, 0, nHeight, iPixels(0), utHeader, DIB_RGB_COLORS
    End If
    ReleaseDC 0, iDC
    SwatchBitmap = iBmp
End Function
