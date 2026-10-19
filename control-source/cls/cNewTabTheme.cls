VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NewTabTheme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "A Theme is a set of property settings stored together."
Option Explicit

Private Declare Function HashData Lib "shlwapi" (ByVal pbData As LongPtr, ByVal cbData As Long, ByRef pbHash As Any, ByVal cbHash As Long) As Long

Private mName As String
Private mHash As String
Private mData As Collection
Private mCustom As Boolean
Private mTag As String
Private mThemeString As String

Friend Property Get ThemeString() As String
    If mThemeString = "" Then
        MakeThemeString
    End If
    ThemeString = mThemeString
End Property

Private Sub MakeThemeString()
    Dim iProp As cPropertyData
    Dim c As Long
    
    For c = 1 To mData.Count
        Set iProp = mData(c)
        mThemeString = mThemeString & IIf(mThemeString = "", "", "|") & iProp.Name & "=" & Trim$(CStr(iProp.Value))
    Next
    mHash = GetThemeStringHash(mThemeString)
End Sub

Friend Property Let ThemeString(nThemeString As String)
    Dim s3() As String
    Dim s4() As String
    Dim c3 As Long
    Dim iProp As cPropertyData
    
    Set mData = New Collection
    mThemeString = ""
    mHash = ""
    s3 = Split(nThemeString, "|") ' tuples of property name and property value
    For c3 = 0 To UBound(s3)
        s4 = Split(s3(c3), "=") ' property name = property value
        If UBound(s4) = 1 Then
            Set iProp = New cPropertyData
            iProp.Name = s4(0)
            If IsNumeric(s4(1)) Then
                iProp.Value = Val(s4(1))
                If iProp.Value = Int(iProp.Value) Then
                    iProp.Value = CLng(iProp.Value)
                End If
            Else
                iProp.Value = s4(1)
            End If
            mData.Add iProp, iProp.Name
        End If
    Next
    TranslateLegacyPropertyNames
End Property

Friend Property Get Data() As Collection
    Set Data = mData
End Property

Friend Property Get Hash() As String
    If mHash = "" Then
        MakeThemeString
    End If
    Hash = mHash
End Property

Private Function GetThemeStringHash(nThemeString As String) As String
    GetThemeStringHash = SimpleHash(nThemeString)
End Function

Private Function SimpleHash(ByVal nData As Variant, Optional nNumberOfHashCharacters_MustBeEvenAndLessThan512 As Long = 16) As String
    Dim iHashBytes() As Byte
    Dim c As Long
    Dim n As Long
    Dim iStr As String
    Dim iVarType As Long
    Dim iDataBytes() As Byte


    n = (nNumberOfHashCharacters_MustBeEvenAndLessThan512 / 2)
    ReDim iHashBytes(n - 1)
    iVarType = VarType(nData)
    If iVarType = vbString Then
        iStr = nData
        HashData StrPtr(iStr), 2 * Len(iStr), iHashBytes(0), n
    Else
        If iVarType <> vbArray + vbByte Then
            Err.Raise 2345, , "Invalid data type"
            Exit Function
        Else
            iDataBytes = nData
            HashData VarPtr(iDataBytes(0)), UBound(iDataBytes) + 1, iHashBytes(0), n
        End If
    End If
    For c = 0 To UBound(iHashBytes)
        iStr = Hex$(iHashBytes(c))
        If Len(iStr) = 1 Then
            iStr = "0" & iStr
        End If
        SimpleHash = SimpleHash & iStr
    Next c
End Function

Friend Function Serialize() As Byte()
    Dim iPb As PropertyBag
    Dim iProp As cPropertyData
    Dim c As Long
    
    Set iPb = New PropertyBag
    iPb.WriteProperty "Name", Name
    iPb.WriteProperty "Custom", Custom, False
    iPb.WriteProperty "Hash", Hash
    iPb.WriteProperty "PropertyCount", mData.Count
    For c = 1 To mData.Count
        Set iProp = mData(c)
        iPb.WriteProperty "PN" & c, iProp.Name
        iPb.WriteProperty "PV" & c, iProp.Value
    Next
    Serialize = iPb.Contents
End Function

Friend Sub Deserialize(nBytes() As Byte)
    Dim iPb As PropertyBag
    Dim iProp As cPropertyData
    Dim c As Long
    Dim iPropertyCount As Long
    
    Set mData = New Collection
    mThemeString = ""
    Set iPb = New PropertyBag
    iPb.Contents = nBytes
    Name = iPb.ReadProperty("Name", "")
    Custom = iPb.ReadProperty("Custom", False)
    mHash = iPb.ReadProperty("Hash", "")
    iPropertyCount = iPb.ReadProperty("PropertyCount", 0)
    For c = 1 To iPropertyCount
        Set iProp = New cPropertyData
        iProp.Name = iPb.ReadProperty("PN" & c, "")
        iProp.Value = iPb.ReadProperty("PV" & c, Empty)
        If Not IsEmpty(iProp.Value) Then
            mData.Add iProp, iProp.Name
        End If
    Next
    TranslateLegacyPropertyNames
End Sub

' The hash is taken over the entries in order, so they are always put in the order
' GetThemeStringFromControl sorts them in and the hash is recomputed. Old property names are
' renamed on the way; when a string carries both names the current one wins
Private Sub TranslateLegacyPropertyNames()
    Dim iProp As cPropertyData
    Dim iNewData As Collection
    Dim iNewName As String
    Dim iSortKey As String
    Dim p As Long
    
    Set iNewData = New Collection
    For Each iProp In mData
        iNewName = CurrentPropertyName(iProp.Name)
        If (iNewName = iProp.Name) Or (Not DataHasProperty(iNewName)) Then
            iProp.Name = iNewName
            iSortKey = PropertySortKey(iProp)
            For p = 1 To iNewData.Count
                If StrComp(iSortKey, PropertySortKey(iNewData(p)), vbBinaryCompare) < 0 Then Exit For
            Next
            If p > iNewData.Count Then
                iNewData.Add iProp, iProp.Name
            Else
                iNewData.Add iProp, iProp.Name, p
            End If
        End If
    Next
    Set mData = iNewData
    mThemeString = ""
    mHash = ""
End Sub

Private Function CurrentPropertyName(ByVal nName As String) As String
    Select Case nName
        Case "BackColorTabSel": CurrentPropertyName = "BackColorSelectedTab"
        Case "ForeColorTabSel": CurrentPropertyName = "ForeColorSelectedTab"
        Case "FlatTabBorderColorHighlight": CurrentPropertyName = "FlatTabBorderColorHighlighted"
        Case "FlatTabBorderColorTabSel": CurrentPropertyName = "FlatTabBorderColorSelectedTab"
        Case "IconColorTabSel": CurrentPropertyName = "IconColorSelectedTab"
        Case "IconColorMouseHoverTabSel": CurrentPropertyName = "IconColorMouseHoverSelectedTab"
        Case "FlatBarColorTabSel": CurrentPropertyName = "FlatBarColorSelectedTab"
        Case "FlatBarColorHighlight": CurrentPropertyName = "FlatBarColorHighlighted"
        Case "HighlightColorTabSel": CurrentPropertyName = "HighlightColorSelectedTab"
        Case "HighlightModeTabSel": CurrentPropertyName = "HighlightModeSelectedTab"
        Case Else: CurrentPropertyName = nName
    End Select
End Function

Private Function DataHasProperty(ByVal nName As String) As Boolean
    Dim iProp As cPropertyData
    
    On Error Resume Next
    Set iProp = mData(nName)
    DataHasProperty = Not iProp Is Nothing
End Function

' The same text QuickSort compares in GetThemeStringFromControl
Private Function PropertySortKey(ByVal nProp As cPropertyData) As String
    PropertySortKey = "|" & nProp.Name & "=" & Trim$(CStr(nProp.Value))
End Function

Friend Function Clone() As NewTabTheme
    Set Clone = New NewTabTheme
    Clone.SetData mThemeString, mHash, mData
    Clone.Name = Me.Name
    Clone.Custom = Me.Custom
End Function

Friend Sub SetData(ByRef nThemeString As String, ByRef nHash As String, ByRef nData As Collection)
    Dim iProp As cPropertyData
    Dim iPropNew As cPropertyData
    
    mThemeString = nThemeString
    mHash = nHash
    Set mData = New Collection
    For Each iProp In nData
        Set iPropNew = New cPropertyData
        iPropNew.Name = iProp.Name
        iPropNew.Value = iProp.Value
        mData.Add iPropNew, iPropNew.Name
    Next
End Sub


Public Property Get Name() As String
Attribute Name.VB_Description = "Returns the name of the theme."
Attribute Name.VB_MemberFlags = "200"
    Name = mName
End Property

Friend Property Let Name(ByVal nName As String)
    mName = Trim$(nName)
End Property


Public Property Get Custom() As Boolean
Attribute Custom.VB_Description = "Returns whether this theme is custom or built-in by default."
    Custom = mCustom
End Property

Friend Property Let Custom(nValue As Boolean)
    mCustom = nValue
End Property


Friend Property Get Tag() As String
    Tag = mTag
End Property

Friend Property Let Tag(nTag As String)
    mTag = nTag
End Property

Private Sub Class_Initialize()
    mCustom = True
End Sub
