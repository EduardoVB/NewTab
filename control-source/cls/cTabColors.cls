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


' The result of the inheritance pass for this tab, indexed by NTChainColorConstants of ctlNewTab
' (private there, hence the literal bound). Never persisted and never written back into the public
' slots below, where Empty says nobody set the colour. Slots 9, 14, 15 and 16 have no per tab meaning
Private mResolvedChainColors(17) As Long

Public BackColorTab As Variant
Public BackColorSelectedTab As Variant
Public HighlightColor As Variant
Public HighlightColorSelectedTab As Variant
Public HighlightGradientColor As Variant
Public HighlightGradientColorSelectedTab As Variant

Public FlatBarColorInactive As Variant
Public FlatBarColorHighlighted As Variant
Public FlatBarColorSelectedTab As Variant

Public FlatTabBorderColorHighlighted As Variant
Public FlatTabBorderColorSelectedTab As Variant
Public FlatTabBorderColorInactive As Variant

Public ForeColor As Variant
Public ForeColorHighlighted As Variant
Public ForeColorSelectedTab As Variant

Public IconColor As Variant
Public IconColorSelectedTab As Variant
Public IconColorMouseHover As Variant
Public IconColorMouseHoverSelectedTab As Variant
Public IconColorTabHighlighted As Variant

Public CloseButtonColor As Variant
Public CloseButtonColorSelectedTab As Variant
Public CloseButtonColorMouseHover As Variant
Public CloseButtonColorMouseHoverSelectedTab As Variant
Public CloseButtonColorTabHighlighted As Variant


Public FlatBarGlowColor As Long


Public Property Get ResolvedChainColor(ByVal nIndex As Long) As Long
    ResolvedChainColor = mResolvedChainColors(nIndex)
End Property

Public Property Let ResolvedChainColor(ByVal nIndex As Long, ByVal nColor As Long)
    mResolvedChainColors(nIndex) = nColor
End Property

' No DefaultValue is passed to WriteProperty: PropertyBag drops a value equal to the default, and
' the Long 0 compares equal to Empty, so black would not be persisted. Absence in the bag is the
' "nobody set this" mark
Private Sub WriteIfSet(ByVal nPb As PropertyBag, ByVal nName As String, ByRef nValue As Variant)
    If Not IsEmpty(nValue) Then nPb.WriteProperty nName, nValue
End Sub

Friend Function Serialize() As Byte()
    Dim iPb As PropertyBag
    
    Set iPb = New PropertyBag
    
    WriteIfSet iPb, "BackColorTab", BackColorTab
    WriteIfSet iPb, "BackColorSelectedTab", BackColorSelectedTab
    WriteIfSet iPb, "HighlightColor", HighlightColor
    WriteIfSet iPb, "HighlightColorSelectedTab", HighlightColorSelectedTab
    WriteIfSet iPb, "HighlightGradientColor", HighlightGradientColor
    WriteIfSet iPb, "HighlightGradientColorSelectedTab", HighlightGradientColorSelectedTab
    WriteIfSet iPb, "FlatBarColorInactive", FlatBarColorInactive
    WriteIfSet iPb, "FlatBarColorHighlighted", FlatBarColorHighlighted
    WriteIfSet iPb, "FlatBarColorSelectedTab", FlatBarColorSelectedTab
    WriteIfSet iPb, "FlatTabBorderColorHighlighted", FlatTabBorderColorHighlighted
    WriteIfSet iPb, "FlatTabBorderColorSelectedTab", FlatTabBorderColorSelectedTab
    WriteIfSet iPb, "FlatTabBorderColorInactive", FlatTabBorderColorInactive
    WriteIfSet iPb, "ForeColor", ForeColor
    WriteIfSet iPb, "ForeColorHighlighted", ForeColorHighlighted
    WriteIfSet iPb, "ForeColorSelectedTab", ForeColorSelectedTab
    WriteIfSet iPb, "IconColor", IconColor
    WriteIfSet iPb, "IconColorSelectedTab", IconColorSelectedTab
    WriteIfSet iPb, "IconColorMouseHover", IconColorMouseHover
    WriteIfSet iPb, "IconColorMouseHoverSelectedTab", IconColorMouseHoverSelectedTab
    WriteIfSet iPb, "IconColorTabHighlighted", IconColorTabHighlighted
    WriteIfSet iPb, "CloseButtonColor", CloseButtonColor
    WriteIfSet iPb, "CloseButtonColorSelectedTab", CloseButtonColorSelectedTab
    WriteIfSet iPb, "CloseButtonColorMouseHover", CloseButtonColorMouseHover
    WriteIfSet iPb, "CloseButtonColorMouseHoverSelectedTab", CloseButtonColorMouseHoverSelectedTab
    WriteIfSet iPb, "CloseButtonColorTabHighlighted", CloseButtonColorTabHighlighted
    
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
    HighlightGradientColor = iPb.ReadProperty("HighlightGradientColor", Empty)
    HighlightGradientColorSelectedTab = iPb.ReadProperty("HighlightGradientColorSelectedTab", Empty)
    FlatBarColorInactive = iPb.ReadProperty("FlatBarColorInactive", Empty)
    FlatBarColorHighlighted = iPb.ReadProperty("FlatBarColorHighlighted", Empty)
    If IsEmpty(FlatBarColorHighlighted) Then FlatBarColorHighlighted = iPb.ReadProperty("FlatBarColorHighlight", Empty)
    FlatBarColorSelectedTab = iPb.ReadProperty("FlatBarColorSelectedTab", Empty)
    FlatTabBorderColorHighlighted = iPb.ReadProperty("FlatTabBorderColorHighlighted", Empty)
    If IsEmpty(FlatTabBorderColorHighlighted) Then FlatTabBorderColorHighlighted = iPb.ReadProperty("FlatTabBorderColorHighlight", Empty)
    FlatTabBorderColorSelectedTab = iPb.ReadProperty("FlatTabBorderColorSelectedTab", Empty)
    FlatTabBorderColorInactive = iPb.ReadProperty("FlatTabBorderColorInactive", Empty)
    ForeColor = iPb.ReadProperty("ForeColor", Empty)
    ForeColorHighlighted = iPb.ReadProperty("ForeColorHighlighted", Empty)
    ForeColorSelectedTab = iPb.ReadProperty("ForeColorSelectedTab", Empty)
    IconColor = iPb.ReadProperty("IconColor", Empty)
    IconColorSelectedTab = iPb.ReadProperty("IconColorSelectedTab", Empty)
    IconColorMouseHover = iPb.ReadProperty("IconColorMouseHover", Empty)
    IconColorMouseHoverSelectedTab = iPb.ReadProperty("IconColorMouseHoverSelectedTab", Empty)
    IconColorTabHighlighted = iPb.ReadProperty("IconColorTabHighlighted", Empty)
    CloseButtonColor = iPb.ReadProperty("CloseButtonColor", Empty)
    CloseButtonColorSelectedTab = iPb.ReadProperty("CloseButtonColorSelectedTab", Empty)
    CloseButtonColorMouseHover = iPb.ReadProperty("CloseButtonColorMouseHover", Empty)
    CloseButtonColorMouseHoverSelectedTab = iPb.ReadProperty("CloseButtonColorMouseHoverSelectedTab", Empty)
    CloseButtonColorTabHighlighted = iPb.ReadProperty("CloseButtonColorTabHighlighted", Empty)
End Sub
