VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTabColors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mFlatBarHighlightEffectColors(10) As Long
Private mHighlightEffectColors_Strong(10) As Long
Private mHighlightEffectColors_Light(10) As Long

Public BackColorTab As Variant
Public BackColorTabSel As Variant
Public HighlightColor As Variant
Public HighlightColorTabSel As Variant

Public FlatBarColorInactive As Variant
Public FlatBarColorHighlight As Variant
Public FlatBarColorTabSel As Variant

Public FlatTabBorderColorHighlight As Variant
Public FlatTabBorderColorTabSel As Variant

Public ForeColor As Variant
Public ForeColorHighlighted As Variant
Public ForeColorTabSel As Variant

Public IconColor As Variant
Public IconColorTabSel As Variant
Public IconColorMouseHover As Variant
Public IconColorMouseHoverTabSel As Variant
Public IconColorTabHighlighted As Variant


Public GlowColor As Long
Public FlatBarGlowColor As Long

Public Property Get FlatBarHighlightEffectColors(ByVal nStep As Long) As Long
    FlatBarHighlightEffectColors = mFlatBarHighlightEffectColors(nStep)
End Property

Public Property Let FlatBarHighlightEffectColors(ByVal nStep As Long, ByVal nColor As Long)
    mFlatBarHighlightEffectColors(nStep) = nColor
End Property


Public Property Get HighlightEffectColors_Light(ByVal nStep As Long) As Long
    HighlightEffectColors_Light = mHighlightEffectColors_Light(nStep)
End Property

Public Property Let HighlightEffectColors_Light(ByVal nStep As Long, ByVal nColor As Long)
    mHighlightEffectColors_Light(nStep) = nColor
End Property


Public Property Get HighlightEffectColors_Strong(ByVal nStep As Long) As Long
    HighlightEffectColors_Strong = mHighlightEffectColors_Strong(nStep)
End Property

Public Property Let HighlightEffectColors_Strong(ByVal nStep As Long, ByVal nColor As Long)
    mHighlightEffectColors_Strong(nStep) = nColor
End Property


