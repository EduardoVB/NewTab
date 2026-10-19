Attribute VB_Name = "mNewTabPicScale"
'=========================================================================
'
' NewTab - DPI scaling of tab pictures
'
' Four parts, kept apart on purpose:
'
'   - ScalePictureToIcon and the pv* helpers around it are NewTab code: they
'     drive the pipeline and turn the result back into an icon.
'
'   - pvGetDIBits is adapted from a bas module by wqweto
'     (https://github.com/wqweto/Xbrz), MIT-0. The icon branch was reworked
'     here - see the comment on it.
'
'   - pvMagnify2x is a VB6 port of MMPX, Copyright 2020 Morgan McGuire and
'     Mara Gagiu, available under the MIT license.
'     https://casual-effects.com/research/McGuire2021PixelArt/index.html
'     The reference implementation's letter names for the neighbourhood
'     (A..I, P, Q, R, S) and the outputs (J, K, L, M) are kept as iPixA,
'     iOutJ and so on, so the rules can still be read against the paper.
'     Do not "fix" those names.
'
'   - pvUpscale and its helpers (pvBuildSampleTable, pvColorDist, pvIsSimilar)
'     are a scaler written for this file. It borrows two published ideas, not
'     code:
'       * from xBRZ (Zenju): deciding which diagonal of every 2x2 block is the
'         edge by comparing colour gradients over a 4x4 neighbourhood, and the
'         45 degree / 2:1 / 1:2 slope classification. That is what makes
'         curves and diagonals come out smooth.
'       * from MMPX: looking two pixels further along the edge before cutting
'         a corner, so right angles and one pixel steps on a straight edge
'         stay square. That is what keeps boxes and frames crisp.
'
' Why pixel-art scalers and not a plain resampler: they reconstruct curves,
' diagonals and sharp corners from the rules of the art instead of interpolating
' across them. That is what keeps a 16 pixel icon readable when it grows.
'
' The caller picks what to favour (NTPicScaleModeConstants):
'
'   ntPicScaleBalanced  pvUpscale with its corner guards on. Smooth curves and
'                       diagonals, square corners on solid shapes, and edges kept
'                       sharp at small factors (see SAMPLE_MAX_FOOTPRINT).
'   ntPicScaleSmooth    pvUpscale with the guards off and full antialiasing at
'                       any factor: every corner found is cut or rounded (closest
'                       to what xBRZ produces) and thin lines keep an even weight
'                       at small factors, at the cost of soft grey edges.
'   ntPicScaleCrisp     MMPX, doubled until it reaches the target and then
'                       reduced by area. Never invents a colour inside a shape
'                       and keeps the hardest corners.
'
' pvUpscale stores the edges it finds as shapes inside each source pixel, not as
' pixels of a fixed 2x or 3x grid, so it renders straight to the size actually
' wanted. MMPX only magnifies by two, so that path overshoots to the next power
' of two and then reduces - and reducing a supersampled image does not blur.
'
' No external dependency: VB6 plus gdi32 and user32.
'
'=========================================================================
Option Explicit

' What ScalePictureToIcon favours when it has to choose - see the notes above
Public Enum NTPicScaleModeConstants
    ntPicScaleBalanced = 0
    ntPicScaleSmooth = 1
    ntPicScaleCrisp = 2
End Enum

'=========================================================================
' API
'=========================================================================

'--- for CreateDIBSection / GetDIBits
Private Const DIB_RGB_COLORS            As Long = 0

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As LongPtr) As LongPtr
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As LongPtr) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As LongPtr, ByVal hObject As LongPtr) As LongPtr
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As LongPtr, lpBitsInfo As BITMAPINFOHEADER, ByVal wUsage As Long, lpBits As LongPtr, ByVal Handle As LongPtr, ByVal dw As Long) As LongPtr
Private Declare Function GetDIBits Lib "gdi32" (ByVal hDC As LongPtr, ByVal hBitmap As LongPtr, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As Any, ByVal wUsage As Long) As Long
Private Declare Function GetObjectW Lib "gdi32" (ByVal hObject As LongPtr, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetIconInfo Lib "user32" (ByVal hIcon As LongPtr, pIconInfo As ICONINFO) As Long
Private Declare Function CreateIconIndirect Lib "user32" (pIconInfo As ICONINFO) As LongPtr

Private Type BITMAPINFOHEADER
    biSize              As Long
    biWidth             As Long
    biHeight            As Long
    biPlanes            As Integer
    biBitCount          As Integer
    biCompression       As Long
    biSizeImage         As Long
    biXPelsPerMeter     As Long
    biYPelsPerMeter     As Long
    biClrUsed           As Long
    biClrImportant      As Long
End Type

Private Type BITMAPAPI
    bmType              As Long
    bmWidth             As Long
    bmHeight            As Long
    bmWidthBytes        As Long
    bmPlanes            As Integer
    bmBitsPixel         As Integer
    bmBits              As LongPtr
End Type

Private Type ICONINFO
    fIcon               As Long
    xHotspot            As Long
    yHotspot            As Long
    hbmMask             As LongPtr
    hbmColor            As LongPtr
End Type

'=========================================================================
' Constants
'=========================================================================

'--- MMPX doubles, so four passes reach 16x, which is as far as the old
'--- nearest neighbour ladder ever went
Private Const SCALE_MAX_PASSES          As Long = 4

' How wide a window each destination pixel averages over when reducing, as a fraction of its own
' footprint. 1 is a plain box filter, which mixes a new colour into every pixel that straddles an
' edge - and after a pixel art magnifier almost every pixel does. A narrower window centred on the
' same place leaves interiors sampling a single source pixel
Private Const AREA_REDUCE_WINDOW        As Double = 0.35

' ---- Scaler tunables ---------------------------------------------------------------
' Two colors closer than this are "the same" for the guards (0 = exact match only).
Private Const EQUAL_TOLERANCE As Single = 30!
' Weight of the center diagonal in the gradient test (xBRZ uses 4).
Private Const CENTER_BIAS As Single = 4!
' One diagonal must beat the other by this factor to be a "dominant" edge, which is always
' drawn as a slope. Higher = more corners go through the guards and stay square.
Private Const DOMINANT_FACTOR_BALANCED As Single = 8!
Private Const DOMINANT_FACTOR_SMOOTH As Single = 3.6
' Ratio needed to turn a 45 degree cut into a 2:1 / 1:2 slope.
Private Const STEEP_FACTOR As Single = 2.2
' Balanced mode only (smooth always rounds):
' True: free-standing 1px corners (line tips, dots) get a small rounding.
' False: they stay perfectly square.
Private Const ROUND_THIN_CORNERS As Boolean = False

' ---- Corner shapes -----------------------------------------------------------------
Private Const SHAPE_NONE As Byte = 0
Private Const SHAPE_ROUND As Byte = 1
Private Const SHAPE_DIAGONAL As Byte = 2
Private Const SHAPE_SHALLOW As Byte = 3
Private Const SHAPE_STEEP As Byte = 4
Private Const SHAPE_BOTH As Byte = 5

' ---- Fixed point (sub-pixel position inside a source pixel, 0..4095) ---------------
Private Const FP_ONE As Long = 4096
Private Const FP_HALF As Long = 2048
' Threshold of the 45 degree cut (6144 = exactly through the edge midpoints; higher = smaller cut)
Private Const FP_DIAGONAL_CUT_BALANCED As Long = 6200
Private Const FP_DIAGONAL_CUT_SMOOTH As Long = 6144
Private Const FP_TWO As Long = 8192

' Sub-samples per output pixel side (4 -> 16 samples).
Private Const SUB_SAMPLES As Long = 4
Private Const SAMPLE_COUNT As Long = 16
' How much of the source an output pixel may average over, in source pixels. At small factors
' (1.25x, 1.5x) an output pixel is almost as big as a source pixel, and averaging over all of it
' turns every edge it straddles into a grey fringe. Capping the area keeps interiors and most
' edges on a single source pixel; from about 1 / this value upwards (3.3x) the cap no longer
' applies and edges get the full antialiasing. Lower = crisper and more uneven line widths,
' higher = smoother and greyer. 1 turns the cap off. Balanced mode only: smooth always uses 1.
Private Const SAMPLE_MAX_FOOTPRINT As Double = 0.3

'==============================================================================
' NewTab entry point
'==============================================================================

' Scales nPic to nTargetWidth x nTargetHeight and returns an icon handle, or 0
' if it could not be done - in which case the caller keeps the original picture.
'
' The handle comes back raw instead of wrapped in a StdPicture so the control
' can build it with its own CreatePicture, which already owns the IID and the
' fPictureOwnsHandle convention. Ownership passes to the caller: whoever gets a
' non zero handle either hands it to CreatePicture (which destroys it with the
' picture) or destroys it.
'
' nMaskColor is a Windows COLORREF - &H00BBGGRR, what OleTranslateColor gives
' every VB6 caller - already resolved to a literal colour, or -1 for none. It is compared once against the untouched source and turned into alpha
' before any scaling runs. That matters because the scaler finds edges by
' comparing pixels: a mask left as a colour would be treated as just another
' colour, and it would be blended into the antialiased edge of the picture.
'
' nSrcWidth / nSrcHeight are only a fallback for pictures with no GDI object to
' measure (metafiles). For bitmaps and icons the real bitmap is measured here,
' because a size that came back through himetric can be off by a pixel and a one
' pixel error in the stride corrupts the whole buffer.
'
' nMode is optional and defaults to ntPicScaleBalanced.
Public Function ScalePictureToIcon(nPic As StdPicture, ByVal nMaskColor As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal nTargetWidth As Long, ByVal nTargetHeight As Long, Optional ByVal nMode As NTPicScaleModeConstants = ntPicScaleBalanced) As LongPtr
    Dim iSrcBits() As Long
    Dim iDstBits() As Long
    Dim iWidth As Long
    Dim iHeight As Long
    Dim iPasses As Long
    Dim iMemDC As LongPtr
    Dim iHeader As BITMAPINFOHEADER
    Dim iDib As LongPtr
    Dim iBitsPtr As LongPtr
    Dim iInfo As ICONINFO

    If nPic Is Nothing Then Exit Function
    If nPic.Handle = 0 Then Exit Function
    If nTargetWidth <= 0 Or nTargetHeight <= 0 Then Exit Function

    iWidth = nSrcWidth
    iHeight = nSrcHeight
    If Not pvGetDIBits(nPic, nMaskColor, iSrcBits, iWidth, iHeight) Then Exit Function
    If iWidth <= 0 Or iHeight <= 0 Then Exit Function

    If nMode = ntPicScaleCrisp Then
        ' Grow first, shrink second, and never the other way round. Every pass
        ' doubles, so the loop stops at or above the target and what follows is
        ' always a reduction
        Do While iWidth < nTargetWidth Or iHeight < nTargetHeight
            If iPasses >= SCALE_MAX_PASSES Then Exit Do
            ReDim iDstBits(0 To 4 * iWidth * iHeight - 1) As Long
            pvMagnify2x iSrcBits, iWidth, iHeight, iDstBits
            iSrcBits = iDstBits
            iWidth = iWidth * 2
            iHeight = iHeight * 2
            iPasses = iPasses + 1
        Loop

        If iWidth <> nTargetWidth Or iHeight <> nTargetHeight Then
            ReDim iDstBits(0 To nTargetWidth * nTargetHeight - 1) As Long
            pvAreaReduce iSrcBits, iWidth, iHeight, iDstBits, nTargetWidth, nTargetHeight
            iSrcBits = iDstBits
        End If
    Else
        ' One pass, straight to the target size. It also covers a target smaller
        ' than the source, where it degrades to an antialiased reduction
        pvUpscale iSrcBits, iWidth, iHeight, iDstBits, nTargetWidth, nTargetHeight, (nMode = ntPicScaleSmooth)
        iSrcBits = iDstBits
    End If

    ' Everything above works in straight alpha, because that is what the scalers
    ' need to compare and average pixels correctly. A 32bpp icon is drawn by
    ' Windows through AlphaBlend, which wants premultiplied, so the conversion
    ' happens once here at the end
    pvPremultiply iSrcBits

    iMemDC = CreateCompatibleDC(0)
    If iMemDC = 0 Then Exit Function
    With iHeader
        .biSize = LenB(iHeader)
        .biPlanes = 1
        .biBitCount = 32
        .biWidth = nTargetWidth
        ' negative height asks for top-down rows, matching the buffer order
        .biHeight = -nTargetHeight
        .biSizeImage = 4 * nTargetWidth * nTargetHeight
    End With
    iDib = CreateDIBSection(iMemDC, iHeader, DIB_RGB_COLORS, iBitsPtr, 0, 0)
    If iDib <> 0 Then
        Call CopyMemory(ByVal iBitsPtr, iSrcBits(0), iHeader.biSizeImage)
        With iInfo
            .fIcon = 1
            .hbmColor = iDib
            ' the same 32bpp DIB as the mask: Windows derives the AND mask from
            ' it and builds an alpha icon. A real 1bpp mask here would make it a
            ' classic icon and throw the alpha away
            .hbmMask = iDib
        End With
        ScalePictureToIcon = CreateIconIndirect(iInfo)
        ' CreateIconIndirect copies both bitmaps
        Call DeleteObject(iDib)
    End If
    Call DeleteDC(iMemDC)
End Function

' Reads nPic into a top-down 32bpp ARGB buffer with straight (not premultiplied) alpha.
' On success every pixel has a meaningful alpha: a source that carries none - a plain bitmap with
' no mask colour - is made fully opaque here, not in a consumer.
' Adapted from mxBRZ.bas by wqweto (MIT-0)
Private Function pvGetDIBits(nPic As StdPicture, ByVal nMaskColor As Long, nBits() As Long, nWidth As Long, nHeight As Long) As Boolean
    Dim iMemDC As LongPtr
    Dim iInfo As ICONINFO
    Dim iHeader As BITMAPINFOHEADER
    Dim iBmp As BITMAPAPI
    Dim iDib As LongPtr
    Dim iBitsPtr As LongPtr
    Dim iPrevDib As LongPtr
    Dim iMaskBits() As Long
    Dim iPos As Long
    Dim iMaskRGB As Long
    Dim iPic As IPicture
    Dim iSourceIsIcon As Boolean
    Dim iHasAlpha As Boolean

    iSourceIsIcon = (nPic.Type = vbPicTypeIcon)
    If iSourceIsIcon Then
        If GetIconInfo(nPic.Handle, iInfo) = 0 Then Exit Function
        ' A monochrome icon has no colour bitmap and there is nothing to scale
        If iInfo.hbmColor = 0 Then GoTo QH
        If GetObjectW(iInfo.hbmColor, LenB(iBmp), iBmp) = 0 Then GoTo QH
        nWidth = iBmp.bmWidth
        nHeight = iBmp.bmHeight
    ElseIf nPic.Type = vbPicTypeBitmap Then
        If GetObjectW(nPic.Handle, LenB(iBmp), iBmp) <> 0 Then
            nWidth = iBmp.bmWidth
            nHeight = iBmp.bmHeight
        End If
    End If
    If nWidth <= 0 Or nHeight <= 0 Then GoTo QH

    iMemDC = CreateCompatibleDC(0)
    If iMemDC = 0 Then GoTo QH
    With iHeader
        .biSize = LenB(iHeader)
        .biPlanes = 1
        .biBitCount = 32
        .biWidth = nWidth
        .biHeight = -nHeight
        .biSizeImage = 4 * nWidth * nHeight
    End With

    If iSourceIsIcon Then
        ReDim nBits(0 To nWidth * nHeight - 1) As Long
        If GetDIBits(iMemDC, iInfo.hbmColor, 0, nHeight, nBits(0), iHeader, DIB_RGB_COLORS) = 0 Then GoTo QH
        ' An ordinary 32bpp DDB has an unused fourth byte that reads back as
        ' zero, so alpha only means something when at least one pixel has a non
        ' zero one - the same test IconToGrayScale in the control has to make
        For iPos = 0 To UBound(nBits)
            If (nBits(iPos) And &HFF000000) <> 0 Then
                iHasAlpha = True
                Exit For
            End If
        Next
        If iHasAlpha Then
            ' A 32bpp icon stores premultiplied colour: undone, because the scalers work
            ' in straight alpha. The AND mask is ignored here: rebuilding alpha from it would
            ' flatten the soft edge of an alpha icon
            pvUnpremultiply nBits
        Else
            ReDim iMaskBits(0 To nWidth * nHeight - 1) As Long
            If GetDIBits(iMemDC, iInfo.hbmMask, 0, nHeight, iMaskBits(0), iHeader, DIB_RGB_COLORS) = 0 Then GoTo QH
            ' Classic icon: the AND mask is the transparency. A set bit means
            ' the destination shows through
            For iPos = 0 To UBound(iMaskBits)
                If iMaskBits(iPos) = 0 Then
                    nBits(iPos) = nBits(iPos) Or &HFF000000
                Else
                    nBits(iPos) = 0
                End If
            Next
        End If
    Else
        iDib = CreateDIBSection(iMemDC, iHeader, DIB_RGB_COLORS, iBitsPtr, 0, 0)
        If iDib = 0 Then GoTo QH
        iPrevDib = SelectObject(iMemDC, iDib)
        Set iPic = nPic
        ' Render and not PaintPicture: the source rectangle goes in himetric and
        ' is computed here instead of being converted by VB6 with its rounded
        ' TwipsPerPixel, which comes out short where 1440 / DPI is not a whole
        ' number. The Y origin is the picture height and the height negative,
        ' because himetric measures Y upwards while the DIB rows go down
        iPic.Render iMemDC, 0, 0, nWidth, nHeight, 0, iPic.Height, iPic.Width, -iPic.Height, ByVal 0
        ReDim nBits(0 To nWidth * nHeight - 1) As Long
        Call CopyMemory(nBits(0), ByVal iBitsPtr, iHeader.biSizeImage)
        ' No alpha detection on this branch: a bitmap has no alpha channel for GDI and what Render
        ' leaves in the fourth byte is undefined. The unscaled draw path masks by colour too
        If nMaskColor <> -1 Then
            ' The mask is resolved once, against the untouched source. nMaskColor is a COLORREF,
            ' &H00BBGGRR, and the DIB reads back as &H00RRGGBB, so it is byte reversed first
            iMaskRGB = ((nMaskColor And &HFF) * &H10000) Or (nMaskColor And &HFF00&) Or ((nMaskColor And &HFF0000) \ &H10000)
            For iPos = 0 To UBound(nBits)
                If (nBits(iPos) And &HFFFFFF) = iMaskRGB Then
                    nBits(iPos) = 0
                Else
                    nBits(iPos) = nBits(iPos) Or &HFF000000
                End If
            Next
        Else
            ' No mask colour: opaque everywhere
            For iPos = 0 To UBound(nBits)
                nBits(iPos) = nBits(iPos) Or &HFF000000
            Next
        End If
    End If

    '--- success
    pvGetDIBits = True
QH:
    If iPrevDib <> 0 Then Call SelectObject(iMemDC, iPrevDib)
    If iDib <> 0 Then Call DeleteObject(iDib)
    If iInfo.hbmColor <> 0 Then Call DeleteObject(iInfo.hbmColor)
    If iInfo.hbmMask <> 0 Then Call DeleteObject(iInfo.hbmMask)
    If iMemDC <> 0 Then Call DeleteDC(iMemDC)
End Function

' Weighted average with fractional coverage: every destination pixel is the
' average of the source area under a window centred on it, weighted by how much
' of each source pixel falls inside. The window is AREA_REDUCE_WINDOW of the
' pixel's own footprint - see the constant. Only ever called to shrink.
'
' The averaging is done on colour premultiplied by alpha and undone at the end.
' Averaging straight alpha instead would pull the colour of fully transparent
' pixels - which is black, since that is what a masked out pixel becomes - into
' every partly transparent edge pixel, and the icon would come out with a dark
' fringe all around it
Private Sub pvAreaReduce(nSrc() As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, nDst() As Long, ByVal nDstWidth As Long, ByVal nDstHeight As Long)
    Dim iDstX As Long
    Dim iDstY As Long
    Dim iKernelX As Long
    Dim iKernelY As Long
    Dim iSpanX As Long
    Dim iSpanY As Long
    Dim iStartX() As Long
    Dim iStartY() As Long
    Dim iCountX() As Long
    Dim iCountY() As Long
    Dim iWeightsX() As Double
    Dim iWeightsY() As Double
    Dim iBaseX As Long
    Dim iBaseY As Long
    Dim iRow As Long
    Dim iPix As Long
    Dim iAlpha As Long
    Dim iStepX As Double
    Dim iStepY As Double
    Dim iWeightY As Double
    Dim iWeight As Double
    Dim iWeighted As Double
    Dim iSumWeight As Double
    Dim iSumAlpha As Double
    Dim iSumRed As Double
    Dim iSumGreen As Double
    Dim iSumBlue As Double

    iStepX = CDbl(nSrcWidth) / nDstWidth
    iStepY = CDbl(nSrcHeight) / nDstHeight
    ' widest footprint a destination pixel can have: a span of nStep can straddle
    ' Int(nStep) whole source pixels plus a partial one at each end
    iSpanX = Int(iStepX) + 2
    iSpanY = Int(iStepY) + 2
    ' The coverage of a source column is the same for every row, so it is worked out once per
    ' destination column
    pvBuildWeights nDstWidth, nSrcWidth, iStepX, iSpanX, iStartX, iCountX, iWeightsX
    pvBuildWeights nDstHeight, nSrcHeight, iStepY, iSpanY, iStartY, iCountY, iWeightsY

    For iDstY = 0 To nDstHeight - 1
        iBaseY = iDstY * iSpanY
        For iDstX = 0 To nDstWidth - 1
            iBaseX = iDstX * iSpanX
            iSumWeight = 0
            iSumAlpha = 0
            iSumRed = 0
            iSumGreen = 0
            iSumBlue = 0
            For iKernelY = 0 To iCountY(iDstY) - 1
                iWeightY = iWeightsY(iBaseY + iKernelY)
                If iWeightY > 0 Then
                    iRow = (iStartY(iDstY) + iKernelY) * nSrcWidth + iStartX(iDstX)
                    For iKernelX = 0 To iCountX(iDstX) - 1
                        iWeight = iWeightsX(iBaseX + iKernelX) * iWeightY
                        If iWeight > 0 Then
                            iPix = nSrc(iRow + iKernelX)
                            ' same arithmetic as pvGetAlpha, inlined to avoid the call overhead
                            If iPix < 0 Then
                                iAlpha = 128 Or ((iPix And &H7F000000) \ &H1000000)
                            Else
                                iAlpha = iPix \ &H1000000
                            End If
                            iWeighted = iWeight * iAlpha
                            iSumWeight = iSumWeight + iWeight
                            iSumAlpha = iSumAlpha + iWeighted
                            iSumRed = iSumRed + ((iPix And &HFF0000) \ &H10000) * iWeighted
                            iSumGreen = iSumGreen + ((iPix And &HFF00&) \ &H100&) * iWeighted
                            iSumBlue = iSumBlue + (iPix And &HFF) * iWeighted
                        End If
                    Next iKernelX
                End If
            Next iKernelY
            If iSumAlpha > 0 Then
                nDst(iDstY * nDstWidth + iDstX) = pvMakePixel( _
                    pvClamp(Int(iSumAlpha / iSumWeight + 0.5), 0, 255), _
                    pvClamp(Int(iSumRed / iSumAlpha + 0.5), 0, 255), _
                    pvClamp(Int(iSumGreen / iSumAlpha + 0.5), 0, 255), _
                    pvClamp(Int(iSumBlue / iSumAlpha + 0.5), 0, 255))
            Else
                nDst(iDstY * nDstWidth + iDstX) = 0
            End If
        Next iDstX
    Next iDstY
End Sub

' One axis of the coverage table: for every destination index, the first source
' index it touches, how many it touches, and the weight of each
Private Sub pvBuildWeights(ByVal nDstCount As Long, ByVal nSrcCount As Long, ByVal nStep As Double, ByVal nSpan As Long, nStart() As Long, nCount() As Long, nWeights() As Double)
    Dim iDst As Long
    Dim iFirst As Long
    Dim iLast As Long
    Dim iIndex As Long
    Dim iLow As Double
    Dim iHigh As Double
    Dim iCentre As Double

    ReDim nStart(0 To nDstCount - 1) As Long
    ReDim nCount(0 To nDstCount - 1) As Long
    ReDim nWeights(0 To nDstCount * nSpan - 1) As Double
    For iDst = 0 To nDstCount - 1
        iLow = iDst * nStep
        iHigh = iLow + nStep
        ' Left alone at 1 rather than recentred on itself, so the full width case
        ' keeps the exact same doubles it had before the window existed
        If AREA_REDUCE_WINDOW < 1 Then
            iCentre = (iLow + iHigh) / 2
            iLow = iCentre - nStep * AREA_REDUCE_WINDOW / 2
            iHigh = iCentre + nStep * AREA_REDUCE_WINDOW / 2
        End If
        iFirst = Int(iLow)
        iLast = pvLastIndex(iHigh, nSrcCount)
        ' belt and braces against a rounding edge widening the footprint past
        ' the span the table was sized for
        If iLast > iFirst + nSpan - 1 Then iLast = iFirst + nSpan - 1
        nStart(iDst) = iFirst
        nCount(iDst) = iLast - iFirst + 1
        For iIndex = 0 To iLast - iFirst
            nWeights(iDst * nSpan + iIndex) = pvOverlap(iLow, iHigh, iFirst + iIndex)
        Next
    Next
End Sub

' How much of source pixel nIndex - which spans [nIndex, nIndex + 1) - falls
' inside the destination footprint [nFrom, nTo)
Private Function pvOverlap(ByVal nFrom As Double, ByVal nTo As Double, ByVal nIndex As Long) As Double
    Dim iLow As Double
    Dim iHigh As Double

    iLow = nIndex
    If nFrom > iLow Then iLow = nFrom
    iHigh = nIndex + 1
    If nTo < iHigh Then iHigh = nTo
    If iHigh > iLow Then pvOverlap = iHigh - iLow
End Function

' Index of the last source pixel touched by a footprint ending at nTo, clamped
' to the image. A footprint ending exactly on a boundary does not touch the
' pixel that starts there
Private Function pvLastIndex(ByVal nTo As Double, ByVal nCount As Long) As Long
    Dim iIndex As Long

    iIndex = Int(nTo)
    If iIndex = nTo Then iIndex = iIndex - 1
    If iIndex > nCount - 1 Then iIndex = nCount - 1
    pvLastIndex = iIndex
End Function

Private Sub pvPremultiply(nBits() As Long)
    Dim iPos As Long
    Dim iPix As Long
    Dim iAlpha As Long

    For iPos = 0 To UBound(nBits)
        iPix = nBits(iPos)
        iAlpha = pvGetAlpha(iPix)
        If iAlpha = 0 Then
            nBits(iPos) = 0
        ElseIf iAlpha < 255 Then
            nBits(iPos) = pvMakePixel(iAlpha, _
                (pvGetRed(iPix) * iAlpha) \ 255, _
                (pvGetGreen(iPix) * iAlpha) \ 255, _
                (pvGetBlue(iPix) * iAlpha) \ 255)
        End If
    Next
End Sub

Private Sub pvUnpremultiply(nBits() As Long)
    Dim iPos As Long
    Dim iPix As Long
    Dim iAlpha As Long

    For iPos = 0 To UBound(nBits)
        iPix = nBits(iPos)
        iAlpha = pvGetAlpha(iPix)
        If iAlpha = 0 Then
            nBits(iPos) = 0
        ElseIf iAlpha < 255 Then
            nBits(iPos) = pvMakePixel(iAlpha, _
                pvClamp((pvGetRed(iPix) * 255) \ iAlpha, 0, 255), _
                pvClamp((pvGetGreen(iPix) * 255) \ iAlpha, 0, 255), _
                pvClamp((pvGetBlue(iPix) * 255) \ iAlpha, 0, 255))
        End If
    Next
End Sub

'==============================================================================
' Pixel channel access
'
' VB6 has no unsigned Long, so bit 31 is the sign bit: an integer divide by
' &H1000000 truncates towards zero on a negative value instead of shifting it
' down. Alpha therefore has to be taken apart around that bit, and put back the
' same way
'==============================================================================

Private Function pvGetBlue(ByVal nPix As Long) As Long
    pvGetBlue = nPix And &HFF
End Function

Private Function pvGetGreen(ByVal nPix As Long) As Long
    pvGetGreen = (nPix And &HFF00&) \ &H100&
End Function

Private Function pvGetRed(ByVal nPix As Long) As Long
    pvGetRed = (nPix And &HFF0000) \ &H10000
End Function

Private Function pvGetAlpha(ByVal nPix As Long) As Long
    If nPix < 0 Then
        pvGetAlpha = 128 Or ((nPix And &H7F000000) \ &H1000000)
    Else
        pvGetAlpha = nPix \ &H1000000
    End If
End Function

Private Function pvMakePixel(ByVal nAlpha As Long, ByVal nRed As Long, ByVal nGreen As Long, ByVal nBlue As Long) As Long
    Dim iResult As Long

    iResult = (nRed * &H10000) Or (nGreen * &H100&) Or nBlue
    If nAlpha And &H80 Then
        pvMakePixel = iResult Or &H80000000 Or ((nAlpha And &H7F) * &H1000000)
    Else
        pvMakePixel = iResult Or (nAlpha * &H1000000)
    End If
End Function

Private Function pvClamp(ByVal nValue As Long, ByVal nLow As Long, ByVal nHigh As Long) As Long
    If nValue < nLow Then
        pvClamp = nLow
    ElseIf nValue > nHigh Then
        pvClamp = nHigh
    Else
        pvClamp = nValue
    End If
End Function

'==============================================================================
' MMPX 2x magnification
' Copyright 2020 Morgan McGuire and Mara Gagiu - MIT license
' https://casual-effects.com/research/McGuire2021PixelArt/index.html
'
' Ported from the reference C++/C implementation. Upstream letter naming kept.
' nDst must already be dimensioned to 4 * nSrcWidth * nSrcHeight elements.
'
' The rules never mix two colours: every destination pixel is one of the source
' pixels, which is why the palette, the flat interiors and the exact alpha of a
' masked icon all survive the magnification untouched
'==============================================================================

Private Sub pvMagnify2x(nSrc() As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, nDst() As Long)
    Dim iSrcX As Long
    Dim iSrcY As Long
    Dim iDstWidth As Long
    Dim iRowTop As Long
    Dim iRowBottom As Long
    Dim iPixA As Long
    Dim iPixB As Long
    Dim iPixC As Long
    Dim iPixD As Long
    Dim iPixE As Long
    Dim iPixF As Long
    Dim iPixG As Long
    Dim iPixH As Long
    Dim iPixI As Long
    Dim iPixP As Long
    Dim iPixQ As Long
    Dim iPixR As Long
    Dim iPixS As Long
    Dim iOutJ As Long
    Dim iOutK As Long
    Dim iOutL As Long
    Dim iOutM As Long
    Dim iLumaB As Long
    Dim iLumaD As Long
    Dim iLumaE As Long
    Dim iLumaF As Long
    Dim iLumaH As Long

    iDstWidth = nSrcWidth * 2
    For iSrcY = 0 To nSrcHeight - 1
        '--- inputs carried along the row
        iPixA = pvSrcPixel(nSrc, nSrcWidth, nSrcHeight, -1, iSrcY - 1)
        iPixB = pvSrcPixel(nSrc, nSrcWidth, nSrcHeight, 0, iSrcY - 1)
        iPixC = pvSrcPixel(nSrc, nSrcWidth, nSrcHeight, 1, iSrcY - 1)
        iPixD = pvSrcPixel(nSrc, nSrcWidth, nSrcHeight, -1, iSrcY)
        iPixE = pvSrcPixel(nSrc, nSrcWidth, nSrcHeight, 0, iSrcY)
        iPixF = pvSrcPixel(nSrc, nSrcWidth, nSrcHeight, 1, iSrcY)
        iPixG = pvSrcPixel(nSrc, nSrcWidth, nSrcHeight, -1, iSrcY + 1)
        iPixH = pvSrcPixel(nSrc, nSrcWidth, nSrcHeight, 0, iSrcY + 1)
        iPixI = pvSrcPixel(nSrc, nSrcWidth, nSrcHeight, 1, iSrcY + 1)
        iPixQ = pvSrcPixel(nSrc, nSrcWidth, nSrcHeight, -2, iSrcY)
        iPixR = pvSrcPixel(nSrc, nSrcWidth, nSrcHeight, 2, iSrcY)
        For iSrcX = 0 To nSrcWidth - 1
            iOutJ = iPixE
            iOutK = iPixE
            iOutL = iPixE
            iOutM = iPixE
            If iPixA <> iPixE Or iPixB <> iPixE Or iPixC <> iPixE Or iPixD <> iPixE Or iPixF <> iPixE Or iPixG <> iPixE Or iPixH <> iPixE Or iPixI <> iPixE Then
                iPixP = pvSrcPixel(nSrc, nSrcWidth, nSrcHeight, iSrcX, iSrcY - 2)
                iPixS = pvSrcPixel(nSrc, nSrcWidth, nSrcHeight, iSrcX, iSrcY + 2)
                iLumaB = pvLuma(iPixB)
                iLumaD = pvLuma(iPixD)
                iLumaE = pvLuma(iPixE)
                iLumaF = pvLuma(iPixF)
                iLumaH = pvLuma(iPixH)

                '--- 1:1 slope rules
                If (iPixD = iPixB And iPixD <> iPixH And iPixD <> iPixF) And (iLumaE >= iLumaD Or iPixE = iPixA) And (iPixE = iPixA Or iPixE = iPixC Or iPixE = iPixG) And (iLumaE < iLumaD Or iPixA <> iPixD Or iPixE <> iPixP Or iPixE <> iPixQ) Then iOutJ = iPixD
                If (iPixB = iPixF And iPixB <> iPixD And iPixB <> iPixH) And (iLumaE >= iLumaB Or iPixE = iPixC) And (iPixE = iPixA Or iPixE = iPixC Or iPixE = iPixI) And (iLumaE < iLumaB Or iPixC <> iPixB Or iPixE <> iPixP Or iPixE <> iPixR) Then iOutK = iPixB
                If (iPixH = iPixD And iPixH <> iPixF And iPixH <> iPixB) And (iLumaE >= iLumaH Or iPixE = iPixG) And (iPixE = iPixA Or iPixE = iPixG Or iPixE = iPixI) And (iLumaE < iLumaH Or iPixG <> iPixH Or iPixE <> iPixS Or iPixE <> iPixQ) Then iOutL = iPixH
                If (iPixF = iPixH And iPixF <> iPixB And iPixF <> iPixD) And (iLumaE >= iLumaF Or iPixE = iPixI) And (iPixE = iPixC Or iPixE = iPixG Or iPixE = iPixI) And (iLumaE < iLumaF Or iPixI <> iPixH Or iPixE <> iPixR Or iPixE <> iPixS) Then iOutM = iPixF

                '--- intersection rules
                If iPixE <> iPixF And iPixE = iPixC And iPixE = iPixI And iPixE = iPixD And iPixE = iPixQ And iPixF = iPixB And iPixF = iPixH And iPixF <> pvSrcPixel(nSrc, nSrcWidth, nSrcHeight, iSrcX + 3, iSrcY) Then
                    iOutK = iPixF
                    iOutM = iPixF
                End If
                If iPixE <> iPixD And iPixE = iPixA And iPixE = iPixG And iPixE = iPixF And iPixE = iPixR And iPixD = iPixB And iPixD = iPixH And iPixD <> pvSrcPixel(nSrc, nSrcWidth, nSrcHeight, iSrcX - 3, iSrcY) Then
                    iOutJ = iPixD
                    iOutL = iPixD
                End If
                If iPixE <> iPixH And iPixE = iPixG And iPixE = iPixI And iPixE = iPixB And iPixE = iPixP And iPixH = iPixD And iPixH = iPixF And iPixH <> pvSrcPixel(nSrc, nSrcWidth, nSrcHeight, iSrcX, iSrcY + 3) Then
                    iOutL = iPixH
                    iOutM = iPixH
                End If
                If iPixE <> iPixB And iPixE = iPixA And iPixE = iPixC And iPixE = iPixH And iPixE = iPixS And iPixB = iPixD And iPixB = iPixF And iPixB <> pvSrcPixel(nSrc, nSrcWidth, nSrcHeight, iSrcX, iSrcY - 3) Then
                    iOutJ = iPixB
                    iOutK = iPixB
                End If
                If iLumaB < iLumaE And iPixE = iPixG And iPixE = iPixH And iPixE = iPixI And iPixE = iPixS And iPixE <> iPixA And iPixE <> iPixD And iPixE <> iPixC And iPixE <> iPixF Then
                    iOutJ = iPixB
                    iOutK = iPixB
                End If
                If iLumaH < iLumaE And iPixE = iPixA And iPixE = iPixB And iPixE = iPixC And iPixE = iPixP And iPixE <> iPixD And iPixE <> iPixG And iPixE <> iPixI And iPixE <> iPixF Then
                    iOutL = iPixH
                    iOutM = iPixH
                End If
                If iLumaF < iLumaE And iPixE = iPixA And iPixE = iPixD And iPixE = iPixG And iPixE = iPixQ And iPixE <> iPixB And iPixE <> iPixC And iPixE <> iPixI And iPixE <> iPixH Then
                    iOutK = iPixF
                    iOutM = iPixF
                End If
                If iLumaD < iLumaE And iPixE = iPixC And iPixE = iPixF And iPixE = iPixI And iPixE = iPixR And iPixE <> iPixB And iPixE <> iPixA And iPixE <> iPixG And iPixE <> iPixH Then
                    iOutJ = iPixD
                    iOutL = iPixD
                End If

                '--- 2:1 slope rules
                If iPixH <> iPixB Then
                    If iPixH <> iPixA And iPixH <> iPixE And iPixH <> iPixC Then
                        If iPixH = iPixG And iPixH = iPixF And iPixH = iPixR And iPixH <> iPixD And iPixH <> pvSrcPixel(nSrc, nSrcWidth, nSrcHeight, iSrcX + 2, iSrcY - 1) Then iOutL = iOutM
                        If iPixH = iPixI And iPixH = iPixD And iPixH = iPixQ And iPixH <> iPixF And iPixH <> pvSrcPixel(nSrc, nSrcWidth, nSrcHeight, iSrcX - 2, iSrcY - 1) Then iOutM = iOutL
                    End If
                    If iPixB <> iPixI And iPixB <> iPixG And iPixB <> iPixE Then
                        If iPixB = iPixA And iPixB = iPixF And iPixB = iPixR And iPixB <> iPixD And iPixB <> pvSrcPixel(nSrc, nSrcWidth, nSrcHeight, iSrcX + 2, iSrcY + 1) Then iOutJ = iOutK
                        If iPixB = iPixC And iPixB = iPixD And iPixB = iPixQ And iPixB <> iPixF And iPixB <> pvSrcPixel(nSrc, nSrcWidth, nSrcHeight, iSrcX - 2, iSrcY + 1) Then iOutK = iOutJ
                    End If
                End If
                If iPixF <> iPixD Then
                    If iPixD <> iPixI And iPixD <> iPixE And iPixD <> iPixC Then
                        If iPixD = iPixA And iPixD = iPixH And iPixD = iPixS And iPixD <> iPixB And iPixD <> pvSrcPixel(nSrc, nSrcWidth, nSrcHeight, iSrcX + 1, iSrcY + 2) Then iOutJ = iOutL
                        If iPixD = iPixG And iPixD = iPixB And iPixD = iPixP And iPixD <> iPixH And iPixD <> pvSrcPixel(nSrc, nSrcWidth, nSrcHeight, iSrcX + 1, iSrcY - 2) Then iOutL = iOutJ
                    End If
                    If iPixF <> iPixE And iPixF <> iPixA And iPixF <> iPixG Then
                        If iPixF = iPixC And iPixF = iPixH And iPixF = iPixS And iPixF <> iPixB And iPixF <> pvSrcPixel(nSrc, nSrcWidth, nSrcHeight, iSrcX - 1, iSrcY + 2) Then iOutK = iOutM
                        If iPixF = iPixI And iPixF = iPixB And iPixF = iPixP And iPixF <> iPixH And iPixF <> pvSrcPixel(nSrc, nSrcWidth, nSrcHeight, iSrcX - 1, iSrcY - 2) Then iOutM = iOutK
                    End If
                End If
            End If

            iRowTop = (iSrcY * 2) * iDstWidth + iSrcX * 2
            iRowBottom = iRowTop + iDstWidth
            nDst(iRowTop) = iOutJ
            nDst(iRowTop + 1) = iOutK
            nDst(iRowBottom) = iOutL
            nDst(iRowBottom + 1) = iOutM

            '--- slide the window one pixel to the right
            iPixA = iPixB
            iPixB = iPixC
            iPixC = pvSrcPixel(nSrc, nSrcWidth, nSrcHeight, iSrcX + 2, iSrcY - 1)
            iPixQ = iPixD
            iPixD = iPixE
            iPixE = iPixF
            iPixF = iPixR
            iPixR = pvSrcPixel(nSrc, nSrcWidth, nSrcHeight, iSrcX + 3, iSrcY)
            iPixG = iPixH
            iPixH = iPixI
            iPixI = pvSrcPixel(nSrc, nSrcWidth, nSrcHeight, iSrcX + 2, iSrcY + 1)
        Next iSrcX
    Next iSrcY
End Sub

' Clamp to border. The C99 reference port has "min > max" where it means
' "v > max" here, which leaves a coordinate past the right or bottom edge
' unclamped and reads outside the buffer; this is the corrected form
Private Function pvSrcPixel(nSrc() As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal nX As Long, ByVal nY As Long) As Long
    Dim iX As Long
    Dim iY As Long

    iX = pvClamp(nX, 0, nSrcWidth - 1)
    iY = pvClamp(nY, 0, nSrcHeight - 1)
    pvSrcPixel = nSrc(iY * nSrcWidth + iX)
End Function

' Weighs the summed channels by how transparent the pixel is, so the rules can
' tell a dark or opaque pixel - which MMPX treats as foreground - from a light
' or transparent one. Deliberately not a perceptual luminance: it only has to
' make that one coarse distinction
Private Function pvLuma(ByVal nPix As Long) As Long
    pvLuma = (pvGetRed(nPix) + pvGetGreen(nPix) + pvGetBlue(nPix) + 1) * (256 - pvGetAlpha(nPix))
End Function

'==============================================================================
' Shape based scaler (balanced / smooth)
'==============================================================================

'-------------------------------------------------------------------------------------
' pvUpscale - shape based scaler (see the notes at the top of the file)
' Pixels are &HAARRGGBB with straight alpha, row-major, top-down, zero based.
'   nSrc()            source pixels, at least nSrcWidth * nSrcHeight elements
'   nSrcWidth/Height  source size in pixels
'   nDst()            receives the result; it is ReDim'ed to nDstWidth * nDstHeight
'   nDstWidth/Height  output size in pixels (any size, not only integer factors)
'   nSmooth           False = balanced (corner guards on, capped sampling area)
'                     True = smooth (guards off, full antialiasing)
'-------------------------------------------------------------------------------------
Private Sub pvUpscale(nSrc() As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, nDst() As Long, ByVal nDstWidth As Long, ByVal nDstHeight As Long, ByVal nSmooth As Boolean)
    Dim iPad() As Long
    Dim iDominantFactor As Single
    Dim iDiagonalCut As Long
    Dim iMaxFootprint As Double
    Dim iBlend() As Byte
    Dim iShape() As Byte
    Dim iCornerColor() As Long
    Dim iHasShape() As Byte
    Dim iColTable() As Long
    Dim iFracXTable() As Long
    Dim iRowTable() As Long
    Dim iFracYTable() As Long
    Dim iOffRight(0 To 3) As Long
    Dim iOffDown(0 To 3) As Long
    Dim iPadW As Long
    Dim iPadH As Long
    Dim iPosX As Long
    Dim iPosY As Long
    Dim iSrcX As Long
    Dim iSrcY As Long
    Dim iPos As Long
    Dim iSrcIdx As Long
    Dim iRot As Long
    Dim iRight As Long
    Dim iDown As Long
    Dim iPixel As Long
    Dim iPixE As Long
    Dim iPixF As Long
    Dim iPixH As Long
    Dim iPixI As Long
    Dim iPixG As Long
    Dim iPixC As Long
    Dim iPixD As Long
    Dim iPixB As Long
    Dim iPixJ As Long
    Dim iPixK As Long
    Dim iDiagJG As Single
    Dim iDiagFK As Single
    Dim iDistFG As Single
    Dim iDistHC As Single
    Dim iType As Byte
    Dim iShapeId As Byte
    Dim iDoLine As Boolean
    Dim iShallow As Boolean
    Dim iSteep As Boolean
    Dim iTick As Boolean
    Dim iInside As Boolean
    Dim iOutX As Long
    Dim iOutY As Long
    Dim iSample As Long
    Dim iRowBase As Long
    Dim iColBase As Long
    Dim iFracU As Long
    Dim iFracV As Long
    Dim iCornerX As Long
    Dim iCornerY As Long
    Dim iDeltaX As Long
    Dim iDeltaY As Long
    Dim iAlpha As Long
    Dim iSumA As Long
    Dim iSumR As Long
    Dim iSumG As Long
    Dim iSumB As Long
    Dim iOutA As Long
    Dim iOutColor As Long

    If nSrcWidth < 1 Or nSrcHeight < 1 Or nDstWidth < 1 Or nDstHeight < 1 Then Exit Sub
    ReDim nDst(0 To nDstWidth * nDstHeight - 1)

    If nSmooth Then
        iDominantFactor = DOMINANT_FACTOR_SMOOTH
        iDiagonalCut = FP_DIAGONAL_CUT_SMOOTH
    Else
        iDominantFactor = DOMINANT_FACTOR_BALANCED
        iDiagonalCut = FP_DIAGONAL_CUT_BALANCED
    End If

    '=================================================================================
    ' STEP 0 - Padded copy (2px border, edge pixels replicated) so the kernels never
    '          need bounds checks. Fully transparent pixels are normalized to 0 so
    '          hidden RGB garbage cannot create fake edges.
    '=================================================================================
    iPadW = nSrcWidth + 4
    iPadH = nSrcHeight + 4
    ReDim iPad(0 To iPadW * iPadH - 1)
    For iPosY = 0 To iPadH - 1
        iSrcY = iPosY - 2
        If iSrcY < 0 Then iSrcY = 0
        If iSrcY > nSrcHeight - 1 Then iSrcY = nSrcHeight - 1
        For iPosX = 0 To iPadW - 1
            iSrcX = iPosX - 2
            If iSrcX < 0 Then iSrcX = 0
            If iSrcX > nSrcWidth - 1 Then iSrcX = nSrcWidth - 1
            iPixel = nSrc(iSrcY * nSrcWidth + iSrcX)
            If (iPixel And &HFF000000) = 0 Then iPixel = 0
            iPad(iPosY * iPadW + iPosX) = iPixel
        Next iPosX
    Next iPosY

    '=================================================================================
    ' STEP 1 - DETECT (xBRZ). For every 2x2 block
    '
    '              F G
    '              J K
    '
    '          compare the "cost" of both diagonals using their 4x4 surroundings.
    '          The diagonal with the LOWER cost is the one whose pixels belong
    '          together, so the edge crosses the other two pixels: their corner that
    '          touches the block center is flagged (1 = weak, 2 = dominant).
    '
    '          iBlend(pos * 4 + n): n = 0 bottom-right, 1 top-right, 2 top-left,
    '          3 bottom-left corner of the pixel at padded position pos.
    '=================================================================================
    ReDim iBlend(0 To iPadW * iPadH * 4 - 1)
    For iPosY = 1 To nSrcHeight + 1
        For iPosX = 1 To nSrcWidth + 1
            iPos = iPosY * iPadW + iPosX
            iPixF = iPad(iPos)
            iPixG = iPad(iPos + 1)
            iPixJ = iPad(iPos + iPadW)
            iPixK = iPad(iPos + iPadW + 1)
            ' Flat block or plain horizontal/vertical edge: nothing to do (fast path,
            ' this is what most of an icon looks like).
            If Not ((iPixF = iPixG And iPixJ = iPixK) Or (iPixF = iPixJ And iPixG = iPixK)) Then
                iDiagJG = pvColorDist(iPad(iPos + iPadW - 1), iPixF) + pvColorDist(iPixF, iPad(iPos - iPadW + 1))
                iDiagJG = iDiagJG + pvColorDist(iPad(iPos + 2 * iPadW), iPixK) + pvColorDist(iPixK, iPad(iPos + 2))
                iDiagJG = iDiagJG + CENTER_BIAS * pvColorDist(iPixJ, iPixG)
                iDiagFK = pvColorDist(iPad(iPos - 1), iPixJ) + pvColorDist(iPixJ, iPad(iPos + 2 * iPadW + 1))
                iDiagFK = iDiagFK + pvColorDist(iPad(iPos - iPadW), iPixG) + pvColorDist(iPixG, iPad(iPos + iPadW + 2))
                iDiagFK = iDiagFK + CENTER_BIAS * pvColorDist(iPixF, iPixK)
                If iDiagJG < iDiagFK Then
                    ' J and G belong together -> F and K get their inner corner cut.
                    iType = 1
                    If iDominantFactor * iDiagJG < iDiagFK Then iType = 2
                    If iPixF <> iPixG And iPixF <> iPixJ Then iBlend(iPos * 4) = iType
                    If iPixK <> iPixJ And iPixK <> iPixG Then iBlend((iPos + iPadW + 1) * 4 + 2) = iType
                ElseIf iDiagFK < iDiagJG Then
                    ' F and K belong together -> J and G get their inner corner cut.
                    iType = 1
                    If iDominantFactor * iDiagFK < iDiagJG Then iType = 2
                    If iPixJ <> iPixF And iPixJ <> iPixK Then iBlend((iPos + iPadW) * 4 + 1) = iType
                    If iPixG <> iPixF And iPixG <> iPixK Then iBlend((iPos + 1) * 4 + 3) = iType
                End If
            End If
        Next iPosX
    Next iPosY

    '=================================================================================
    ' STEP 2 - CLASSIFY. For each flagged corner decide its shape. The same code
    '          handles the 4 corners by rotating the neighbourhood: "right" and "down"
    '          are offsets that turn 90 degrees per rotation, so the corner being
    '          examined is always the bottom-right one of this local frame:
    '
    '                  B  C
    '               D  E  F  F2         E = this pixel (corner under test: E/F/H/I)
    '            G2 G  H  I             F2 = two steps right, H2 = two steps down
    '                  H2               G2 = two steps left of G, C2 = two up of C
    '
    '   CURVE vs CORNER (the "selector"):
    '     - dominant edge (type 2)            -> always a slope (curve / diagonal)
    '     - weak edge (type 1) is a slope unless one of the MMPX-style guards says
    '       this is really a right angle:
    '         a) an adjacent corner of E is also flagged and E is not continued
    '            diagonally there (E is an L/T junction or a line tip),
    '         b) F, H, I, G, C form a solid "L" around E (E is a concave notch).
    '       In those cases the corner stays square (or is only slightly rounded
    '       when E is a thin feature, see ROUND_THIN_CORNERS).
    '
    '   SLOPE TYPE (xBRZ): 2:1 shallow if F looks like G, 1:2 steep if H looks like C,
    '       otherwise 45 degrees.
    '
    '   STRAIGHT-EDGE-WITH-NOISE guard (MMPX): a 2:1 slope is only real if the
    '       staircase goes on. If two pixels further along the row the edge is back
    '       at the same level (E == F2 and G == G2), the step is a 1px bump/notch on
    '       a straight edge: do not smear it into a long wedge, keep it square.
    '=================================================================================
    iOffRight(0) = 1
    iOffDown(0) = iPadW
    iOffRight(1) = -iPadW
    iOffDown(1) = 1
    iOffRight(2) = -1
    iOffDown(2) = -iPadW
    iOffRight(3) = iPadW
    iOffDown(3) = -1

    ReDim iShape(0 To nSrcWidth * nSrcHeight * 4 - 1)
    ReDim iCornerColor(0 To nSrcWidth * nSrcHeight * 4 - 1)
    ReDim iHasShape(0 To nSrcWidth * nSrcHeight - 1)

    For iPosY = 0 To nSrcHeight - 1
        For iPosX = 0 To nSrcWidth - 1
            iPos = (iPosY + 2) * iPadW + iPosX + 2
            iSrcIdx = iPosY * nSrcWidth + iPosX
            iPixE = iPad(iPos)
            For iRot = 0 To 3
                iType = iBlend(iPos * 4 + iRot)
                If iType <> 0 Then
                    iRight = iOffRight(iRot)
                    iDown = iOffDown(iRot)
                    iPixF = iPad(iPos + iRight)
                    iPixH = iPad(iPos + iDown)
                    iPixI = iPad(iPos + iRight + iDown)
                    iPixG = iPad(iPos - iRight + iDown)
                    iPixC = iPad(iPos + iRight - iDown)
                    iPixD = iPad(iPos - iRight)
                    iPixB = iPad(iPos - iDown)

                    ' --- Curve or right angle? ---
                    iDoLine = True
                    If iType < 2 Then
                        If iBlend(iPos * 4 + ((iRot + 1) And 3)) <> 0 Then
                            If Not pvIsSimilar(iPixE, iPixG) Then iDoLine = False
                        End If
                        If iDoLine Then
                            If iBlend(iPos * 4 + ((iRot + 3) And 3)) <> 0 Then
                                If Not pvIsSimilar(iPixE, iPixC) Then iDoLine = False
                            End If
                        End If
                        If iDoLine Then
                            If Not pvIsSimilar(iPixE, iPixI) Then
                                If pvIsSimilar(iPixG, iPixH) Then
                                    If pvIsSimilar(iPixH, iPixI) Then
                                        If pvIsSimilar(iPixI, iPixF) Then
                                            If pvIsSimilar(iPixF, iPixC) Then iDoLine = False
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If

                    ' Color that fills the cut: the neighbour closest to E.
                    If pvColorDist(iPixE, iPixF) <= pvColorDist(iPixE, iPixH) Then
                        iPixel = iPixF
                    Else
                        iPixel = iPixH
                    End If

                    iShapeId = SHAPE_NONE
                    If iDoLine Then
                        iDistFG = pvColorDist(iPixF, iPixG)
                        iDistHC = pvColorDist(iPixH, iPixC)
                        iShallow = False
                        iSteep = False
                        iTick = False
                        If STEEP_FACTOR * iDistFG <= iDistHC Then
                            If iPixE <> iPixG And iPixD <> iPixG Then iShallow = True
                        End If
                        If STEEP_FACTOR * iDistHC <= iDistFG Then
                            If iPixE <> iPixC And iPixB <> iPixC Then iSteep = True
                        End If
                        ' --- Straight edge with a 1px bump: cancel the long slope (balanced only) ---
                        If iShallow And Not nSmooth Then
                            If pvIsSimilar(iPad(iPos + 2 * iRight), iPixE) Then
                                If pvIsSimilar(iPad(iPos - 2 * iRight + iDown), iPixG) Then
                                    iShallow = False
                                    iTick = True
                                End If
                            End If
                        End If
                        If iSteep And Not nSmooth Then
                            If pvIsSimilar(iPad(iPos + 2 * iDown), iPixE) Then
                                If pvIsSimilar(iPad(iPos + iRight - 2 * iDown), iPixC) Then
                                    iSteep = False
                                    iTick = True
                                End If
                            End If
                        End If
                        If iShallow And iSteep Then
                            iShapeId = SHAPE_BOTH
                        ElseIf iShallow Then
                            iShapeId = SHAPE_SHALLOW
                        ElseIf iSteep Then
                            iShapeId = SHAPE_STEEP
                        ElseIf Not iTick Then
                            iShapeId = SHAPE_DIAGONAL
                        End If
                    Else
                        ' Right angle. If E continues on both far sides (D and B) this
                        ' is the corner of a solid area: keep it perfectly square.
                        ' Otherwise E is thin here (line tip, dot): optional rounding.
                        ' Smooth skips the question and rounds every one of them.
                        If nSmooth Then
                            iShapeId = SHAPE_ROUND
                        ElseIf pvIsSimilar(iPixE, iPixD) And pvIsSimilar(iPixE, iPixB) Then
                            iShapeId = SHAPE_NONE
                        ElseIf ROUND_THIN_CORNERS Then
                            iShapeId = SHAPE_ROUND
                        End If
                    End If

                    If iShapeId <> SHAPE_NONE Then
                        iShape(iSrcIdx * 4 + iRot) = iShapeId
                        iCornerColor(iSrcIdx * 4 + iRot) = iPixel
                        iHasShape(iSrcIdx) = 1
                    End If
                End If
            Next iRot
        Next iPosX
    Next iPosY

    '=================================================================================
    ' STEP 3 - RENDER. Every output pixel takes 4x4 sub-samples. Each sample maps to
    '          a source pixel plus a position inside it (fixed point 0..4095). If the
    '          source pixel has shaped corners, the sample is tested against the
    '          half-plane (or arc) of each one and takes the corner color when it
    '          falls inside. Samples are averaged with alpha weighting.
    '          All position math is precomputed per output column / row.
    '=================================================================================
    If nSmooth Then
        iMaxFootprint = 1#
    Else
        iMaxFootprint = SAMPLE_MAX_FOOTPRINT
    End If
    pvBuildSampleTable nDstWidth, nSrcWidth, False, iMaxFootprint, iColTable, iFracXTable
    pvBuildSampleTable nDstHeight, nSrcHeight, True, iMaxFootprint, iRowTable, iFracYTable

    For iOutY = 0 To nDstHeight - 1
        iRowBase = iOutY * SAMPLE_COUNT
        For iOutX = 0 To nDstWidth - 1
            iColBase = iOutX * SAMPLE_COUNT
            iSumA = 0
            iSumR = 0
            iSumG = 0
            iSumB = 0
            For iSample = 0 To SAMPLE_COUNT - 1
                iSrcX = iColTable(iColBase + iSample)
                iSrcY = iRowTable(iRowBase + iSample)
                iSrcIdx = iSrcY * nSrcWidth + iSrcX
                iPixel = iPad((iSrcY + 2) * iPadW + iSrcX + 2)
                If iHasShape(iSrcIdx) <> 0 Then
                    iFracU = iFracXTable(iColBase + iSample)
                    iFracV = iFracYTable(iRowBase + iSample)
                    For iRot = 0 To 3
                        iShapeId = iShape(iSrcIdx * 4 + iRot)
                        If iShapeId <> SHAPE_NONE Then
                            ' Rotate the sample into the corner's local frame, where
                            ' the corner under test sits at (FP_ONE, FP_ONE).
                            Select Case iRot
                                Case 0
                                    iCornerX = iFracU
                                    iCornerY = iFracV
                                Case 1
                                    iCornerX = FP_ONE - iFracV
                                    iCornerY = iFracU
                                Case 2
                                    iCornerX = FP_ONE - iFracU
                                    iCornerY = FP_ONE - iFracV
                                Case Else
                                    iCornerX = iFracV
                                    iCornerY = FP_ONE - iFracU
                            End Select
                            iInside = False
                            Select Case iShapeId
                                Case SHAPE_DIAGONAL
                                    ' 45 degree cut through the midpoints of both edges
                                    iInside = (iCornerX + iCornerY > iDiagonalCut)
                                Case SHAPE_SHALLOW
                                    ' 2:1 slope: from the far bottom corner to mid right edge
                                    iInside = (iCornerX + 2 * iCornerY > FP_TWO)
                                Case SHAPE_STEEP
                                    ' 1:2 slope: from the far right corner to mid bottom edge
                                    iInside = (2 * iCornerX + iCornerY > FP_TWO)
                                Case SHAPE_BOTH
                                    iInside = (iCornerX + 2 * iCornerY > FP_TWO) Or (2 * iCornerX + iCornerY > FP_TWO)
                                Case SHAPE_ROUND
                                    ' Outside of a quarter circle of radius 1/2 pixel
                                    If iCornerX > FP_HALF And iCornerY > FP_HALF Then
                                        iDeltaX = iCornerX - FP_HALF
                                        iDeltaY = iCornerY - FP_HALF
                                        iInside = (iDeltaX * iDeltaX + iDeltaY * iDeltaY > FP_HALF * FP_HALF)
                                    End If
                            End Select
                            If iInside Then iPixel = iCornerColor(iSrcIdx * 4 + iRot)
                        End If
                    Next iRot
                End If
                ' Accumulate (alpha-weighted so transparent pixels do not bleed color)
                iAlpha = ((iPixel And &HFF000000) \ &H1000000) And &HFF&
                If iAlpha <> 0 Then
                    iSumA = iSumA + iAlpha
                    iSumR = iSumR + ((iPixel And &HFF0000) \ &H10000) * iAlpha
                    iSumG = iSumG + ((iPixel And &HFF00&) \ &H100&) * iAlpha
                    iSumB = iSumB + (iPixel And &HFF&) * iAlpha
                End If
            Next iSample

            If iSumA = 0 Then
                iOutColor = 0
            Else
                iOutA = (iSumA + SAMPLE_COUNT \ 2) \ SAMPLE_COUNT
                iOutColor = ((iSumR + iSumA \ 2) \ iSumA) * &H10000
                iOutColor = iOutColor Or (((iSumG + iSumA \ 2) \ iSumA) * &H100&)
                iOutColor = iOutColor Or ((iSumB + iSumA \ 2) \ iSumA)
                ' Alpha goes in the sign byte: avoid overflow for values > 127
                If iOutA > 127 Then
                    iOutColor = iOutColor Or ((iOutA - 256) * &H1000000)
                Else
                    iOutColor = iOutColor Or (iOutA * &H1000000)
                End If
            End If
            nDst(iOutY * nDstWidth + iOutX) = iOutColor
        Next iOutX
    Next iOutY
End Sub

'-------------------------------------------------------------------------------------
' Precomputes, for every output coordinate and every sub-sample, which source
' pixel it falls in and the fixed-point position inside that pixel.
' The sub-sample grid is slightly sheared (each row of samples is shifted a bit)
' so perfectly vertical / horizontal edges get 16 coverage levels instead of 4.
' The grid is shrunk around the centre of the output pixel when enlarging by a small
' factor: nMaxFootprint is the most of the source, in source pixels, that one output
' pixel may average over - see SAMPLE_MAX_FOOTPRINT. 1 or more = never shrunk.
'   nIsVertical = False -> table for X (uses the sample column, sheared by its row)
'   nIsVertical = True  -> table for Y (uses the sample row, sheared by its column)
'-------------------------------------------------------------------------------------
Private Sub pvBuildSampleTable(ByVal nDstSize As Long, ByVal nSrcSize As Long, ByVal nIsVertical As Boolean, ByVal nMaxFootprint As Double, nIndexTable() As Long, nFracTable() As Long)
    Dim iOut As Long
    Dim iSample As Long
    Dim iSubX As Long
    Dim iSubY As Long
    Dim iCell As Long
    Dim iCenter As Double
    Dim iOffset As Double
    Dim iSrcPos As Double
    Dim iWindow As Double
    Const JITTER As Double = 0.125

    ReDim nIndexTable(0 To nDstSize * SAMPLE_COUNT - 1)
    ReDim nFracTable(0 To nDstSize * SAMPLE_COUNT - 1)
    iCenter = (SUB_SAMPLES - 1) / 2#
    iWindow = 1#
    If nDstSize > nSrcSize Then
        iWindow = nMaxFootprint * nDstSize / nSrcSize
        If iWindow > 1# Then iWindow = 1#
    End If

    For iOut = 0 To nDstSize - 1
        For iSample = 0 To SAMPLE_COUNT - 1
            iSubY = iSample \ SUB_SAMPLES
            iSubX = iSample Mod SUB_SAMPLES
            If nIsVertical Then
                iOffset = (iSubY + 0.5 - JITTER * (iSubX - iCenter)) / SUB_SAMPLES
            Else
                iOffset = (iSubX + 0.5 + JITTER * (iSubY - iCenter)) / SUB_SAMPLES
            End If
            iSrcPos = (iOut + 0.5 + (iOffset - 0.5) * iWindow) * nSrcSize / nDstSize
            iCell = Int(iSrcPos)
            If iCell < 0 Then iCell = 0
            If iCell > nSrcSize - 1 Then iCell = nSrcSize - 1
            nIndexTable(iOut * SAMPLE_COUNT + iSample) = iCell
            nFracTable(iOut * SAMPLE_COUNT + iSample) = Int((iSrcPos - iCell) * FP_ONE)
        Next iSample
    Next iOut
End Sub

'-------------------------------------------------------------------------------------
' Perceptual color distance (YCbCr, as in xBRZ), alpha aware: the color difference
' is scaled by the smaller alpha and the alpha difference is added on top.
' Returns 0 for identical colors, about 255 for black vs white.
'-------------------------------------------------------------------------------------
Private Function pvColorDist(ByVal nColorA As Long, ByVal nColorB As Long) As Single
    Dim iAlphaA As Long
    Dim iAlphaB As Long
    Dim iDiffR As Long
    Dim iDiffG As Long
    Dim iDiffB As Long
    Dim iLuma As Single
    Dim iCb As Single
    Dim iCr As Single
    Dim iDist As Single

    If nColorA = nColorB Then Exit Function

    iAlphaA = ((nColorA And &HFF000000) \ &H1000000) And &HFF&
    iAlphaB = ((nColorB And &HFF000000) \ &H1000000) And &HFF&
    iDiffR = ((nColorA And &HFF0000) \ &H10000) - ((nColorB And &HFF0000) \ &H10000)
    iDiffG = ((nColorA And &HFF00&) \ &H100&) - ((nColorB And &HFF00&) \ &H100&)
    iDiffB = (nColorA And &HFF&) - (nColorB And &HFF&)

    iLuma = 0.2627 * iDiffR + 0.678 * iDiffG + 0.0593 * iDiffB
    iCb = 0.531519 * (iDiffB - iLuma)
    iCr = 0.67815 * (iDiffR - iLuma)
    iDist = Sqr(iLuma * iLuma + iCb * iCb + iCr * iCr)

    If iAlphaA < iAlphaB Then
        pvColorDist = iAlphaA * iDist / 255! + (iAlphaB - iAlphaA)
    Else
        pvColorDist = iAlphaB * iDist / 255! + (iAlphaA - iAlphaB)
    End If
End Function

Private Function pvIsSimilar(ByVal nColorA As Long, ByVal nColorB As Long) As Boolean
    If nColorA = nColorB Then
        pvIsSimilar = True
    Else
        pvIsSimilar = (pvColorDist(nColorA, nColorB) < EQUAL_TOLERANCE)
    End If
End Function
