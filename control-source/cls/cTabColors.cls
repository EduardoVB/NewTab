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
Public BackColorSelectedTab As Variant
Public HighlightColor As Variant
Public HighlightColorSelectedTab As Variant

Public FlatBarColorInactive As Variant
Public FlatBarColorHighlight As Variant
Public FlatBarColorSelectedTab As Variant

Public FlatTabBorderColorHighlight As Variant
Public FlatTabBorderColorSelectedTab As Variant

Public ForeColor As Variant
Public ForeColorHighlighted As Variant
Public ForeColorSelectedTab As Variant

Public IconColor As Variant
Public IconColorSelectedTab As Variant
Public IconColorMouseHover As Variant
Public IconColorMouseHoverSelectedTab As Variant
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
    iPb.WriteProperty "BackColorSelectedTab", BackColorSelectedTab, Empty
    iPb.WriteProperty "HighlightColor", HighlightColor, Empty
    iPb.WriteProperty "HighlightColorSelectedTab", HighlightColorSelectedTab, Empty
    iPb.WriteProperty "FlatBarColorInactive", FlatBarColorInactive, Empty
    iPb.WriteProperty "FlatBarColorHighlight", FlatBarColorHighlight, Empty
    iPb.WriteProperty "FlatBarColorSelectedTab", FlatBarColorSelectedTab, Empty
    iPb.WriteProperty "FlatTabBorderColorHighlight", FlatTabBorderColorHighlight, Empty
    iPb.WriteProperty "FlatTabBorderColorSelectedTab", FlatTabBorderColorSelectedTab, Empty
    iPb.WriteProperty "ForeColor", ForeColor, Empty
    iPb.WriteProperty "ForeColorHighlighted", ForeColorHighlighted, Empty
    iPb.WriteProperty "ForeColorSelectedTab", ForeColorSelectedTab, Empty
    iPb.WriteProperty "IconColor", IconColor, Empty
    iPb.WriteProperty "IconColorSelectedTab", IconColorSelectedTab, Empty
    iPb.WriteProperty "IconColorMouseHover", IconColorMouseHover, Empty
    iPb.WriteProperty "IconColorMouseHoverSelectedTab", IconColorMouseHoverSelectedTab, Empty
    iPb.WriteProperty "IconColorTabHighlighted", IconColorTabHighlighted, Empty
    
    Serialize = iPb.Contents
End Function

Friend Sub Deserialize(nBytes() As Byte)
    Dim iPb As PropertyBag
    
    Set iPb = New PropertyBag
    iPb.Contents = nBytes
    
    BackColorTab = iPb.ReadProperty("BackColorTab", Empty)
    BackColorSelectedTab = iPb.ReadProperty("BackColorSelectedTab", Empty)
    HighlightColor = iPb.ReadProperty("HighlightColor", Empty)
    HighlightColorSelectedTab = iPb.ReadProperty("HighlightColorSelectedTab", Empty)
    FlatBarColorInactive = iPb.ReadProperty("FlatBarColorInactive", Empty)
    FlatBarColorHighlight = iPb.ReadProperty("FlatBarColorHighlight", Empty)
    FlatBarColorSelectedTab = iPb.ReadProperty("FlatBarColorSelectedTab", Empty)
    FlatTabBorderColorHighlight = iPb.ReadProperty("FlatTabBorderColorHighlight", Empty)
    FlatTabBorderColorSelectedTab = iPb.ReadProperty("FlatTabBorderColorSelectedTab", Empty)
    ForeColor = iPb.ReadProperty("ForeColor", Empty)
    ForeColorHighlighted = iPb.ReadProperty("ForeColorHighlighted", Empty)
    ForeColorSelectedTab = iPb.ReadProperty("ForeColorSelectedTab", Empty)
    IconColor = iPb.ReadProperty("IconColor", Empty)
    IconColorSelectedTab = iPb.ReadProperty("IconColorSelectedTab", Empty)
    IconColorMouseHover = iPb.ReadProperty("IconColorMouseHover", Empty)
    IconColorMouseHoverSelectedTab = iPb.ReadProperty("IconColorMouseHoverSelectedTab", Empty)
    IconColorTabHighlighted = iPb.ReadProperty("IconColorTabHighlighted", Empty)
End Sub
