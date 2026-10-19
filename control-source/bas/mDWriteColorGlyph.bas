Attribute VB_Name = "mDWriteColorGlyph"
Option Explicit

' Color glyphs (emoji, color icon fonts) through DirectWrite and Direct2D. GDI draws any glyph in
' a single color whatever color tables the font carries. Color fonts need Windows 8.1 or later:
' before that ColorGlyphsSupported answers False and the callers keep drawing with GDI.
' The COM interfaces are called through their vtables with DispCallFunc, so the cVt constants
' must match the order of the methods in d2d1.h and dwrite.h exactly

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type D2D1_COLOR_F
    Red As Single
    Green As Single
    Blue As Single
    Alpha As Single
End Type

Private Type D2D1_RECT_F
    Left As Single
    Top As Single
    Right As Single
    Bottom As Single
End Type

Private Type D2D1_RENDER_TARGET_PROPERTIES
    RenderTargetType As Long
    PixelFormat As Long
    AlphaMode As Long
    DpiX As Single
    DpiY As Single
    Usage As Long
    MinLevel As Long
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

Private Type XFORM
    eM11 As Single
    eM12 As Single
    eM21 As Single
    eM22 As Single
    eDx As Single
    eDy As Single
End Type

Private Type TEXTMETRICW
    tmHeight As Long
    tmAscent As Long
    tmDescent As Long
    tmInternalLeading As Long
    tmExternalLeading As Long
    tmAveCharWidth As Long
    tmMaxCharWidth As Long
    tmWeight As Long
    tmOverhang As Long
    tmDigitizedAspectX As Long
    tmDigitizedAspectY As Long
    tmFirstChar As Integer
    tmLastChar As Integer
    tmDefaultChar As Integer
    tmBreakChar As Integer
    tmItalic As Byte
    tmUnderlined As Byte
    tmStruckOut As Byte
    tmPitchAndFamily As Byte
    tmCharSet As Byte
End Type

Private Declare Function DWriteCreateFactory Lib "dwrite" (ByVal factoryType As Long, ByRef iid As GUID, ByRef factory As LongPtr) As Long
Private Declare Function D2D1CreateFactory Lib "d2d1" (ByVal factoryType As Long, ByRef riid As GUID, ByVal pFactoryOptions As LongPtr, ByRef ppIFactory As LongPtr) As Long
Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As LongPtr, ByVal oVft As LongPtr, ByVal cc As Long, ByVal vtReturn As Integer, ByVal cActuals As Long, ByRef prgvt As Integer, ByRef prgpvarg As LongPtr, ByRef pvargResult As Variant) As Long
Private Declare Function IIDFromString Lib "ole32" (ByVal lpsz As LongPtr, ByRef lpiid As GUID) As Long
Private Declare Function OleTranslateColor Lib "oleaut32" (ByVal clr As Long, ByVal hpal As LongPtr, ByRef lpcolorref As Long) As Long
Private Declare Function GetTextFaceW Lib "gdi32" (ByVal hDC As LongPtr, ByVal nCount As Long, ByVal lpFaceName As LongPtr) As Long
Private Declare Function GetTextMetricsW Lib "gdi32" (ByVal hDC As LongPtr, ByRef lptm As TEXTMETRICW) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As LongPtr) As LongPtr
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As LongPtr) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As LongPtr, ByRef pbmi As BITMAPINFOHEADER, ByVal iUsage As Long, ByRef ppvBits As LongPtr, ByVal hSection As LongPtr, ByVal dwOffset As Long) As LongPtr
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As LongPtr, ByVal hObject As LongPtr) As LongPtr
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
Private Declare Function GetCurrentObject Lib "gdi32" (ByVal hDC As LongPtr, ByVal uObjectType As Long) As LongPtr
Private Declare Function GdiFlush Lib "gdi32" () As Long
Private Declare Function GetLayout Lib "gdi32" (ByVal hDC As LongPtr) As Long
Private Declare Function SetLayout Lib "gdi32" (ByVal hDC As LongPtr, ByVal dwLayout As Long) As Long
Private Declare Function GetGraphicsMode Lib "gdi32" (ByVal hDC As LongPtr) As Long
Private Declare Function GetWorldTransform Lib "gdi32" (ByVal hDC As LongPtr, ByRef lpXform As XFORM) As Long
Private Declare Function SetWorldTransform Lib "gdi32" (ByVal hDC As LongPtr, ByRef lpXform As XFORM) As Long
Private Declare Function ModifyWorldTransform Lib "gdi32" (ByVal hDC As LongPtr, ByRef lpXform As Any, ByVal iMode As Long) As Long
Private Declare Function AlphaBlend Lib "msimg32" (ByVal hdcDest As LongPtr, ByVal xDest As Long, ByVal yDest As Long, ByVal wDest As Long, ByVal hDest As Long, ByVal hdcSrc As LongPtr, ByVal xSrc As Long, ByVal ySrc As Long, ByVal wSrc As Long, ByVal hSrc As Long, ByVal blendFunction As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

#If Win64 Then
Private Const cPtrSize As Long = 8
#Else
Private Const cPtrSize As Long = 4
#End If

Private Const CC_STDCALL As Long = 4
Private Const OBJ_FONT As Long = 6
Private Const DIB_RGB_COLORS As Long = 0
Private Const LAYOUT_RTL As Long = &H1
Private Const LAYOUT_BITMAPORIENTATIONPRESERVED As Long = &H8
Private Const GM_ADVANCED As Long = 2
Private Const MWT_IDENTITY As Long = 1
' BLENDFUNCTION: AC_SRC_OVER, no flags, constant alpha 255, AC_SRC_ALPHA
Private Const cBlendPerPixelAlpha As Long = &H1FF0000
Private Const E_POINTER As Long = &H80004003
Private Const D2DERR_RECREATE_TARGET As Long = &H8899000C

Private Const DXGI_FORMAT_B8G8R8A8_UNORM As Long = 87
Private Const D2D1_ALPHA_MODE_PREMULTIPLIED As Long = 1
Private Const D2D1_TEXT_ANTIALIAS_MODE_GRAYSCALE As Long = 2
Private Const D2D1_DRAW_TEXT_OPTIONS_ENABLE_COLOR_FONT As Long = 4
Private Const DWRITE_FONT_STYLE_ITALIC As Long = 2
Private Const DWRITE_FONT_STRETCH_NORMAL As Long = 5
Private Const DWRITE_TEXT_ALIGNMENT_CENTER As Long = 2
Private Const DWRITE_PARAGRAPH_ALIGNMENT_NEAR As Long = 0
Private Const DWRITE_LINE_SPACING_METHOD_UNIFORM As Long = 1
Private Const DWRITE_WORD_WRAPPING_NO_WRAP As Long = 1

' vtable indexes
Private Const cVtQueryInterface As Long = 0
Private Const cVtRelease As Long = 2
Private Const cVtD2DFactory_CreateDCRenderTarget As Long = 16
Private Const cVtRenderTarget_CreateSolidColorBrush As Long = 8
Private Const cVtRenderTarget_DrawText As Long = 27
Private Const cVtRenderTarget_SetTextAntialiasMode As Long = 34
Private Const cVtRenderTarget_Clear As Long = 47
Private Const cVtRenderTarget_BeginDraw As Long = 48
Private Const cVtRenderTarget_EndDraw As Long = 49
Private Const cVtDCRenderTarget_BindDC As Long = 57
Private Const cVtSolidColorBrush_SetColor As Long = 8
Private Const cVtDWFactory_CreateTextFormat As Long = 15
Private Const cVtTextFormat_SetTextAlignment As Long = 3
Private Const cVtTextFormat_SetParagraphAlignment As Long = 4
Private Const cVtTextFormat_SetWordWrapping As Long = 5
Private Const cVtTextFormat_SetLineSpacing As Long = 10

Private mSupportChecked As Boolean
Private mSupported As Boolean
Private mDWriteFactory As LongPtr
Private mD2DFactory As LongPtr
Private mRenderTarget As LongPtr
Private mBrush As LongPtr
Private mTextFormat As LongPtr
Private mTextFormatKey As String
Private mTextHeight As Long
Private mUsers As Long

' Every NewTab registers on Initialize and leaves on Terminate. Registering creates nothing: the
' objects are still made on the first color glyph, and released when the last control goes
Public Sub ColorGlyphsAddUser()
    mUsers = mUsers + 1
End Sub

Public Sub ColorGlyphsRemoveUser()
    If mUsers > 0 Then mUsers = mUsers - 1
    If mUsers = 0 Then ColorGlyphsReleaseAll
End Sub

Private Sub ColorGlyphsReleaseAll()
    ReleaseObject mTextFormat
    mTextFormatKey = ""
    ReleaseRenderTarget
    ReleaseObject mD2DFactory
    ReleaseObject mDWriteFactory
    mSupportChecked = False
    mSupported = False
End Sub

Public Function ColorGlyphsSupported() As Boolean
    If Not mSupportChecked Then
        mSupportChecked = True
        mSupported = CreateFactories
    End If
    ColorGlyphsSupported = mSupported
End Function

' Draws nCharCode centered in the rectangle, with the font currently selected into nHdc and at its
' pixel size. nForeColor paints the parts of the glyph that have no color of their own.
' nBackColor = -1 leaves the background transparent, in premultiplied alpha, which is what a
' 32-bit DIB section needs. Returns False when nothing was drawn, and the caller falls back to GDI
Public Function DrawColorGlyph(ByVal nHdc As LongPtr, ByVal nCharCode As Long, ByVal nForeColor As Long, ByVal nBackColor As Long, ByVal nLeft As Long, ByVal nTop As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Boolean
    Dim iBindRect As RECT
    Dim iForeColor As D2D1_COLOR_F
    Dim iBackColor As D2D1_COLOR_F
    Dim iLayoutRect As D2D1_RECT_F
    Dim iText As String
    Dim iNull As LongPtr
    Dim iHR As Long

    If Not ColorGlyphsSupported Then Exit Function
    If (nWidth <= 0) Or (nHeight <= 0) Then Exit Function
    If Not EnsureTextFormat(nHdc) Then Exit Function
    If Not EnsureRenderTarget Then Exit Function

    iBindRect.Left = nLeft
    iBindRect.Top = nTop
    iBindRect.Right = nLeft + nWidth
    iBindRect.Bottom = nTop + nHeight
    If CallVtbl(mRenderTarget, cVtDCRenderTarget_BindDC, nHdc, VarPtr(iBindRect)) <> 0 Then Exit Function

    iForeColor = ColorFFromOleColor(nForeColor)
    CallVtbl mBrush, cVtSolidColorBrush_SetColor, VarPtr(iForeColor)
    If nBackColor <> -1 Then iBackColor = ColorFFromOleColor(nBackColor)
    ' The line box is GDI's own, placed where DT_CALCRECT centring would put it
    iLayoutRect.Top = (nHeight - mTextHeight) \ 2
    iLayoutRect.Right = nWidth
    iLayoutRect.Bottom = iLayoutRect.Top + mTextHeight
    iText = ChrU(nCharCode)

    CallVtbl mRenderTarget, cVtRenderTarget_BeginDraw
    CallVtbl mRenderTarget, cVtRenderTarget_Clear, VarPtr(iBackColor)
    CallVtbl mRenderTarget, cVtRenderTarget_DrawText, StrPtr(iText), CLng(Len(iText)), mTextFormat, VarPtr(iLayoutRect), mBrush, D2D1_DRAW_TEXT_OPTIONS_ENABLE_COLOR_FONT, CLng(0)
    iHR = CallVtbl(mRenderTarget, cVtRenderTarget_EndDraw, iNull, iNull)
    If iHR = D2DERR_RECREATE_TARGET Then ReleaseRenderTarget
    DrawColorGlyph = (iHR = 0)
End Function

' The glyph is drawn upright in a transparent bitmap of nGlyphWidth x nGlyphHeight, with the font
' selected into nFontDC, then turned if nRotation asks for it (1 = 90 degrees clockwise, -1 =
' counterclockwise, so the destination is nGlyphHeight x nGlyphWidth) and alpha blended onto
' nDestDC at nDestLeft, nDestTop. A world transform on nDestDC is set aside for the blend,
' because the rotation is already in the pixels, and a mirrored layout keeps the bitmap unflipped
Public Function DrawColorGlyphOnDC(ByVal nDestDC As LongPtr, ByVal nFontDC As LongPtr, ByVal nCharCode As Long, ByVal nForeColor As Long, ByVal nDestLeft As Long, ByVal nDestTop As Long, ByVal nGlyphWidth As Long, ByVal nGlyphHeight As Long, ByVal nRotation As Long) As Boolean
    Dim iGlyphDC As LongPtr
    Dim iGlyphBitmap As LongPtr
    Dim iGlyphPrevBitmap As LongPtr
    Dim iGlyphBits As LongPtr
    Dim iPrevFont As LongPtr
    Dim iOutDC As LongPtr
    Dim iOutBitmap As LongPtr
    Dim iOutPrevBitmap As LongPtr
    Dim iOutBits As LongPtr
    Dim iPixels() As Long
    Dim iTurned() As Long
    Dim iX As Long
    Dim iY As Long
    
    If Not ColorGlyphsSupported Then Exit Function
    If (nGlyphWidth <= 0) Or (nGlyphHeight <= 0) Then Exit Function
    iGlyphDC = CreateDIBDC(nGlyphWidth, nGlyphHeight, iGlyphBitmap, iGlyphPrevBitmap, iGlyphBits)
    If iGlyphDC = 0 Then Exit Function
    iPrevFont = SelectObject(iGlyphDC, GetCurrentObject(nFontDC, OBJ_FONT))
    If DrawColorGlyph(iGlyphDC, nCharCode, nForeColor, -1, 0, 0, nGlyphWidth, nGlyphHeight) Then
        GdiFlush
        If nRotation = 0 Then
            BlendOnto nDestDC, nDestLeft, nDestTop, nGlyphWidth, nGlyphHeight, iGlyphDC
            DrawColorGlyphOnDC = True
        Else
            iOutDC = CreateDIBDC(nGlyphHeight, nGlyphWidth, iOutBitmap, iOutPrevBitmap, iOutBits)
            If iOutDC <> 0 Then
                ReDim iPixels(nGlyphWidth * nGlyphHeight - 1)
                ReDim iTurned(nGlyphWidth * nGlyphHeight - 1)
                CopyMemory iPixels(0), ByVal iGlyphBits, nGlyphWidth * nGlyphHeight * 4
                ' both bitmaps are top-down; the turned one is nGlyphHeight pixels wide
                For iY = 0 To nGlyphHeight - 1
                    For iX = 0 To nGlyphWidth - 1
                        If nRotation > 0 Then
                            iTurned(iX * nGlyphHeight + (nGlyphHeight - 1 - iY)) = iPixels(iY * nGlyphWidth + iX)
                        Else
                            iTurned((nGlyphWidth - 1 - iX) * nGlyphHeight + iY) = iPixels(iY * nGlyphWidth + iX)
                        End If
                    Next
                Next
                CopyMemory ByVal iOutBits, iTurned(0), nGlyphWidth * nGlyphHeight * 4
                BlendOnto nDestDC, nDestLeft, nDestTop, nGlyphHeight, nGlyphWidth, iOutDC
                FreeDIBDC iOutDC, iOutBitmap, iOutPrevBitmap
                DrawColorGlyphOnDC = True
            End If
        End If
    End If
    SelectObject iGlyphDC, iPrevFont
    FreeDIBDC iGlyphDC, iGlyphBitmap, iGlyphPrevBitmap
End Function

Private Sub BlendOnto(ByVal nDestDC As LongPtr, ByVal nLeft As Long, ByVal nTop As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal nSourceDC As LongPtr)
    Dim iLayoutPrev As Long
    Dim iTransformPrev As XFORM
    Dim iTransformSet As Boolean
    
    iLayoutPrev = GetLayout(nDestDC)
    If (iLayoutPrev And LAYOUT_RTL) <> 0 Then SetLayout nDestDC, iLayoutPrev Or LAYOUT_BITMAPORIENTATIONPRESERVED
    If GetGraphicsMode(nDestDC) = GM_ADVANCED Then
        GetWorldTransform nDestDC, iTransformPrev
        ModifyWorldTransform nDestDC, ByVal 0&, MWT_IDENTITY
        iTransformSet = True
    End If
    AlphaBlend nDestDC, nLeft, nTop, nWidth, nHeight, nSourceDC, 0, 0, nWidth, nHeight, cBlendPerPixelAlpha
    If iTransformSet Then SetWorldTransform nDestDC, iTransformPrev
    If (iLayoutPrev And LAYOUT_RTL) <> 0 Then SetLayout nDestDC, iLayoutPrev
End Sub

' A 32-bit top-down DIB section selected into a memory DC, zero filled
Private Function CreateDIBDC(ByVal nWidth As Long, ByVal nHeight As Long, ByRef nBitmap As LongPtr, ByRef nPrevBitmap As LongPtr, ByRef nBits As LongPtr) As LongPtr
    Dim iBIH As BITMAPINFOHEADER
    Dim iDC As LongPtr
    Dim iNull As LongPtr
    
    iDC = CreateCompatibleDC(iNull)
    If iDC = 0 Then Exit Function
    iBIH.biSize = LenB(iBIH)
    iBIH.biWidth = nWidth
    iBIH.biHeight = -nHeight
    iBIH.biPlanes = 1
    iBIH.biBitCount = 32
    nBitmap = CreateDIBSection(iDC, iBIH, DIB_RGB_COLORS, nBits, iNull, 0)
    If (nBitmap = 0) Or (nBits = 0) Then
        If nBitmap <> 0 Then DeleteObject nBitmap
        DeleteDC iDC
        Exit Function
    End If
    nPrevBitmap = SelectObject(iDC, nBitmap)
    CreateDIBDC = iDC
End Function

Private Sub FreeDIBDC(ByVal nDC As LongPtr, ByVal nBitmap As LongPtr, ByVal nPrevBitmap As LongPtr)
    SelectObject nDC, nPrevBitmap
    DeleteObject nBitmap
    DeleteDC nDC
End Sub

Private Function CreateFactories() As Boolean
    Dim iIID As GUID
    Dim iDWriteFactory2 As LongPtr
    Dim iNull As LongPtr

    ' dwrite.dll and d2d1.dll do not exist before Vista SP2
    On Error GoTo ErrHandler
    IIDFromString StrPtr("{b859ee5a-d838-4b5b-a2e8-1adc7d93db48}"), iIID
    If DWriteCreateFactory(0, iIID, mDWriteFactory) <> 0 Then GoTo ErrHandler
    ' IDWriteFactory2 came with Windows 8.1, together with color font support
    IIDFromString StrPtr("{0439fc60-ca44-4994-8dee-3a9af7b732ec}"), iIID
    If CallVtbl(mDWriteFactory, cVtQueryInterface, VarPtr(iIID), VarPtr(iDWriteFactory2)) <> 0 Then GoTo ErrHandler
    ReleaseObject iDWriteFactory2
    IIDFromString StrPtr("{06152247-6f50-465a-9245-118bfd3b6007}"), iIID
    If D2D1CreateFactory(0, iIID, iNull, mD2DFactory) <> 0 Then GoTo ErrHandler
    CreateFactories = True
    Exit Function

ErrHandler:
    ReleaseObject mD2DFactory
    ReleaseObject mDWriteFactory
End Function

Private Function EnsureRenderTarget() As Boolean
    Dim iProps As D2D1_RENDER_TARGET_PROPERTIES
    Dim iColor As D2D1_COLOR_F
    Dim iNull As LongPtr

    If mRenderTarget <> 0 Then
        EnsureRenderTarget = True
        Exit Function
    End If
    iProps.PixelFormat = DXGI_FORMAT_B8G8R8A8_UNORM
    iProps.AlphaMode = D2D1_ALPHA_MODE_PREMULTIPLIED
    ' Not an assumption about the screen: at 96 a DIP is a pixel, so every size handed to
    ' Direct2D is in pixels of the DC
    iProps.DpiX = 96
    iProps.DpiY = 96
    If CallVtbl(mD2DFactory, cVtD2DFactory_CreateDCRenderTarget, VarPtr(iProps), VarPtr(mRenderTarget)) <> 0 Then
        mRenderTarget = 0
        Exit Function
    End If
    ' ClearType needs an opaque target
    CallVtbl mRenderTarget, cVtRenderTarget_SetTextAntialiasMode, D2D1_TEXT_ANTIALIAS_MODE_GRAYSCALE
    iColor.Alpha = 1
    If CallVtbl(mRenderTarget, cVtRenderTarget_CreateSolidColorBrush, VarPtr(iColor), iNull, VarPtr(mBrush)) <> 0 Then
        mBrush = 0
        ReleaseObject mRenderTarget
        Exit Function
    End If
    EnsureRenderTarget = True
End Function

Private Sub ReleaseRenderTarget()
    ReleaseObject mBrush
    ReleaseObject mRenderTarget
End Sub

' The text format is rebuilt only when the font selected into the DC changes
Private Function EnsureTextFormat(ByVal nHdc As LongPtr) As Boolean
    Dim iTM As TEXTMETRICW
    Dim iFaceName As String
    Dim iLen As Long
    Dim iEmHeight As Long
    Dim iWeight As Long
    Dim iStyle As Long
    Dim iKey As String
    Dim iLocale As String
    Dim iNull As LongPtr

    If GetTextMetricsW(nHdc, iTM) = 0 Then Exit Function
    iLen = GetTextFaceW(nHdc, 0, iNull)
    If iLen <= 1 Then Exit Function
    iFaceName = String$(iLen, vbNullChar)
    GetTextFaceW nHdc, iLen, StrPtr(iFaceName)
    iFaceName = Left$(iFaceName, InStr(iFaceName & vbNullChar, vbNullChar) - 1)
    ' The em height, which is what a GDI point size turns into and what DirectWrite calls the size
    iEmHeight = iTM.tmHeight - iTM.tmInternalLeading
    If iEmHeight <= 0 Then Exit Function
    iWeight = iTM.tmWeight
    If iWeight = 0 Then iWeight = 400
    If iTM.tmItalic <> 0 Then iStyle = DWRITE_FONT_STYLE_ITALIC

    iKey = iFaceName & "|" & iEmHeight & "|" & iWeight & "|" & iStyle & "|" & iTM.tmHeight & "|" & iTM.tmAscent
    If (iKey = mTextFormatKey) And (mTextFormat <> 0) Then
        EnsureTextFormat = True
        Exit Function
    End If
    ReleaseObject mTextFormat
    mTextFormatKey = ""
    iLocale = "en-us"
    If CallVtbl(mDWriteFactory, cVtDWFactory_CreateTextFormat, StrPtr(iFaceName), iNull, iWeight, iStyle, DWRITE_FONT_STRETCH_NORMAL, CSng(iEmHeight), StrPtr(iLocale), VarPtr(mTextFormat)) <> 0 Then
        mTextFormat = 0
        Exit Function
    End If
    CallVtbl mTextFormat, cVtTextFormat_SetTextAlignment, DWRITE_TEXT_ALIGNMENT_CENTER
    CallVtbl mTextFormat, cVtTextFormat_SetParagraphAlignment, DWRITE_PARAGRAPH_ALIGNMENT_NEAR
    CallVtbl mTextFormat, cVtTextFormat_SetWordWrapping, DWRITE_WORD_WRAPPING_NO_WRAP
    ' GDI's line height and baseline instead of the ones DirectWrite takes from the font, which
    ' can differ and would sit the glyph a pixel or two off the GDI drawn ones
    CallVtbl mTextFormat, cVtTextFormat_SetLineSpacing, DWRITE_LINE_SPACING_METHOD_UNIFORM, CSng(iTM.tmHeight), CSng(iTM.tmAscent)
    mTextHeight = iTM.tmHeight
    mTextFormatKey = iKey
    EnsureTextFormat = True
End Function

Private Function ColorFFromOleColor(ByVal nColor As Long) As D2D1_COLOR_F
    Dim iRGB As Long

    If OleTranslateColor(nColor, 0, iRGB) <> 0 Then iRGB = 0
    ColorFFromOleColor.Red = (iRGB And &HFF&) / 255
    ColorFFromOleColor.Green = ((iRGB \ &H100&) And &HFF&) / 255
    ColorFFromOleColor.Blue = ((iRGB \ &H10000) And &HFF&) / 255
    ColorFFromOleColor.Alpha = 1
End Function

' Every argument has to arrive with its exact type: a bare literal would be an Integer, two bytes
' where the method expects four, so the callers pass typed variables, constants As Long, or CLng
Private Function CallVtbl(ByVal nObject As LongPtr, ByVal nIndex As Long, ParamArray nArgs() As Variant) As Long
    Dim iTypes() As Integer
    Dim iPtrs() As LongPtr
    Dim iArgs() As Variant
    Dim iResult As Variant
    Dim iCount As Long
    Dim iArg As Long
    Dim iHR As Long

    If nObject = 0 Then
        CallVtbl = E_POINTER
        Exit Function
    End If
    iCount = UBound(nArgs) + 1
    ' one spare slot so that element 0 exists even with no arguments
    ReDim iTypes(iCount)
    ReDim iPtrs(iCount)
    ReDim iArgs(iCount)
    For iArg = 0 To iCount - 1
        iArgs(iArg) = nArgs(iArg)
        iTypes(iArg) = VarType(iArgs(iArg))
        iPtrs(iArg) = VarPtr(iArgs(iArg))
    Next
    iHR = DispCallFunc(nObject, nIndex * cPtrSize, CC_STDCALL, vbLong, iCount, iTypes(0), iPtrs(0), iResult)
    If iHR <> 0 Then
        CallVtbl = iHR
    ElseIf VarType(iResult) = vbLong Then
        CallVtbl = iResult
    End If
End Function

Private Sub ReleaseObject(ByRef nObject As LongPtr)
    If nObject <> 0 Then
        CallVtbl nObject, cVtRelease
        nObject = 0
    End If
End Sub

Private Function ChrU(ByVal nCharCode As Long) As String
    If nCharCode <= &HFFFF& Then
        ChrU = ChrW$(nCharCode)
    Else
        ChrU = ChrW$(&HD800& + (nCharCode - &H10000) \ &H400&) & ChrW$(&HDC00& + (nCharCode And &H3FF&))
    End If
End Function
