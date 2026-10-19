Attribute VB_Name = "mNewTabDefaults"
Option Explicit

' The version of the persisted data format, written into the property bag as "ControlVersion".
' Up to 15 it was App.Major, so the numbering continues that line; from 16 on it moves only when
' the format changes
Public Const cDataFormatVersion As Long = 19

Public Const cPropDef_Style As Long = ntStyleWindows ' ssStyleTabbedDialog
Public Const cPropDef_TabHeight As Long = 238  ' in twips
Public Const cPropDef_WordWrap As Boolean = True
Public Const cPropDef_MaskColor As Long = &HFF00FF
Public Const cPropDef_UseMaskColor As Boolean = True
Public Const cPropDef_ScalePicturesForDPI As Long = ntScalePicturesBalanced
Public Const cPropDef_TDIFormsIconSize As Long = 24 ' ntTDIFormsIcon24
Public Const cPropDef_ShowFocusRect As Boolean = False
Public Const cPropDef_TabsPerRow As Long = 3
Public Const cPropDef_Tabs As Long = 3 ' the tab count of a new control, the same as SSTab's
Public Const cPropDef_ShowDisabledState As Boolean = False
Public Const cPropDef_ChangeControlsBackColor As Boolean = True
Public Const cPropDef_ChangeControlsForeColor As Boolean = True
Public Const cPropDef_HighlightEffect As Long = ntAnimationFast
Public Const cPropDef_HighlightEffectOut As Long = ntAnimationMedium
Public Const cPropDef_TabWidthStyle As Long = ntTWAuto
Public Const cPropDef_TabHeightStyleForHorizontal As Long = ntTHMaxCaptionHeight
Public Const cPropDef_ShowRowsInPerspective As Long = ntYNAuto
Public Const cPropDef_NewTabButton As Long = ntYNAuto ' the button shows only in the TDI Controls mode, which is what the control did before the property existed
Public Const cPropDef_TabSeparation As Integer = 0
Public Const cPropDef_TabPaddingWidth As Long = 5
Public Const cPropDef_CloseButtonPadding As Long = 6
Public Const cPropDef_TabPaddingHeight As Long = 4
Public Const cPropDef_TabAppearance As Long = ntTAAuto
Public Const cPropDef_IconAlignment As Long = ntIconAlignCenteredBeforeCaption ' ntIconAlignBeforeCaption
Public Const cPropDef_AutoRelocateControls As Long = ntRelocateAlways
Public Const cPropDef_SoftEdges As Boolean = True
Public Const cPropDef_HandleHighContrastThem As Boolean = True
Public Const cPropDef_BackStyle As Long = ntOpaque
Public Const cPropDef_AutoTabHeight As Boolean = True
Public Const cPropDef_OLEDropOnOtherTabs As Boolean = True
Public Const cPropDef_TabTransition As Long = ntAnimationNone
Public Const cPropDef_FlatBarColorSelectedTab As Long = &HFC7A43   ' &HF1956D
Public Const cPropDef_FlatRoundnessTop As Long = 8
Public Const cPropDef_FlatRoundnessBottom As Long = 8
Public Const cPropDef_FlatRoundnessTabs As Long = 0
Public Const cPropDef_TabMousePointerHand As Boolean = False
Public Const cPropDef_IconFontName As String = "Segoe MDL2 Assets" ' "Wingdings"
Public Const cPropDef_IconFontSize As Single = 11
Public Const cPropDef_HighlightMode As Long = ntHLAuto
Public Const cPropDef_HighlightModeSelectedTab As Long = ntHLAuto
Public Const cPropDef_FlatBorderMode  As Long = ntBorderTabs
Public Const cPropDef_FlatBarHeight As Long = 3
Public Const cPropDef_FlatBarGripHeight As Long = -5
Public Const cPropDef_FlatBarPosition As Long = ntBarPositionTop
Public Const cPropDef_FlatTabBodyGap As Long = 0
Public Const cPropDef_CanReorderTabs As Boolean = False
Public Const cPropDef_TDIMode As Long = ntTDIModeNone
Public Const cPropDef_TDIDefaultTabCaption As String = "Default tab"
Public Const cPropDef_TDINewTabCaption As String = "New tab"
' In twips and DPI scaled, like the property. Read only through the control's
' DefaultHighlightTabExtraHeightHimetric; the selected tab's widening uses the same 30 twips
Public Const cPropDef_HighlightTabExtraHeight As Long = 30 'twips
' 15 twips, 1 pixel at 96 DPI. Not zero: the value does nothing until ntHLShiftCaption asks for it
Public Const cPropDef_HighlightTabCaptionShift As Long = 15 'twips ' the pixel that separates the TabStrip's selected caption from SSTab property page's, at 96 DPI
Public Const cPropDef_FlatBodySeparationLineHeight As Long = 1
Public Const cPropDef_TabMaxWidth As Long = 0
Public Const cPropDef_TabMinWidth As Long = 0
Public Const cPropDef_TabGeneralFixedWidth As Long = 3175 ' Himetric: 1.25 inch, 120 pixels at 96 DPI
Public Const cPropDef_SubclassingMethod As Long = ntSMSetWindowSubclass
Public Const cPropDef_StripEndFreeSpace As Long = 0
Public Const cPropDef_CaptionAlignment As Long = ntAlignmentCenter
Public Const cPropDef_CloseButtonColorMouseHover As Long = vbRed


' HighlightEffect shipped as a Boolean and is now a speed. A stored True (-1) resolves to the
' default speed and not to a pinned member. Public because UserControl_ReadProperties and the
' theme string reader both need it
Public Function AnimationSpeedFromStoredValue(ByVal nStored As Variant) As Long
    Dim lStored As Long

    If VarType(nStored) = vbBoolean Then
        If nStored Then
            AnimationSpeedFromStoredValue = cPropDef_HighlightEffect
        Else
            AnimationSpeedFromStoredValue = ntAnimationNone
        End If
        Exit Function
    End If
    lStored = CLng(nStored)
    If lStored < ntAnimationNone Then
        AnimationSpeedFromStoredValue = cPropDef_HighlightEffect
    ElseIf lStored > ntAnimationSlower Then
        AnimationSpeedFromStoredValue = ntAnimationSlower
    Else
        AnimationSpeedFromStoredValue = lStored
    End If
End Function
