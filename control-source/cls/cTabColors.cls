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

Friend Function Serialize() As Byte()
    Dim iPb As PropertyBag
    
    Set iPb = New PropertyBag
    
    iPb.WriteProperty "BackColorTab", BackColorTab, Empty
    iPb.WriteProperty "BackColorTabSel", BackColorTabSel, Empty
    iPb.WriteProperty "HighlightColor", HighlightColor, Empty
    iPb.WriteProperty "HighlightColorTabSel", HighlightColorTabSel, Empty
    iPb.WriteProperty "FlatBarColorInactive", FlatBarColorInactive, Empty
    iPb.WriteProperty "FlatBarColorHighlight", FlatBarColorHighlight, Empty
    iPb.WriteProperty "FlatBarColorTabSel", FlatBarColorTabSel, Empty
    iPb.WriteProperty "FlatTabBorderColorHighlight", FlatTabBorderColorHighlight, Empty
    iPb.WriteProperty "FlatTabBorderColorTabSel", FlatTabBorderColorTabSel, Empty
    iPb.WriteProperty "ForeColor", ForeColor, Empty
    iPb.WriteProperty "ForeColorHighlighted", ForeColorHighlighted, Empty
    iPb.WriteProperty "ForeColorTabSel", ForeColorTabSel, Empty
    iPb.WriteProperty "IconColor", IconColor, Empty
    iPb.WriteProperty "IconColorTabSel", IconColorTabSel, Empty
    iPb.WriteProperty "IconColorMouseHover", IconColorMouseHover, Empty
    iPb.WriteProperty "IconColorMouseHoverTabSel", IconColorMouseHoverTabSel, Empty
    iPb.WriteProperty "IconColorTabHighlighted", IconColorTabHighlighted, Empty
    
    Serialize = iPb.Contents
End Function

Friend Sub Deserialize(nBytes() As Byte)
    Dim iPb As PropertyBag
    
    Set iPb = New PropertyBag
    iPb.Contents = nBytes
    
    BackColorTab = iPb.ReadProperty("BackColorTab", Empty)
    BackColorTabSel = iPb.ReadProperty("BackColorTabSel", Empty)
    HighlightColor = iPb.ReadProperty("HighlightColor", Empty)
    HighlightColorTabSel = iPb.ReadProperty("HighlightColorTabSel", Empty)
    FlatBarColorInactive = iPb.ReadProperty("FlatBarColorInactive", Empty)
    FlatBarColorHighlight = iPb.ReadProperty("FlatBarColorHighlight", Empty)
    FlatBarColorTabSel = iPb.ReadProperty("FlatBarColorTabSel", Empty)
    FlatTabBorderColorHighlight = iPb.ReadProperty("FlatTabBorderColorHighlight", Empty)
    FlatTabBorderColorTabSel = iPb.ReadProperty("FlatTabBorderColorTabSel", Empty)
    ForeColor = iPb.ReadProperty("ForeColor", Empty)
    ForeColorHighlighted = iPb.ReadProperty("ForeColorHighlighted", Empty)
    ForeColorTabSel = iPb.ReadProperty("ForeColorTabSel", Empty)
    IconColor = iPb.ReadProperty("IconColor", Empty)
    IconColorTabSel = iPb.ReadProperty("IconColorTabSel", Empty)
    IconColorMouseHover = iPb.ReadProperty("IconColorMouseHover", Empty)
    IconColorMouseHoverTabSel = iPb.ReadProperty("IconColorMouseHoverTabSel", Empty)
    IconColorTabHighlighted = iPb.ReadProperty("IconColorTabHighlighted", Empty)
End Sub
