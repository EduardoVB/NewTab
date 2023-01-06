VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NewTabThemes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Returns the colection of ""Themes""."
Option Explicit

Event ThemeRenamed()
Attribute ThemeRenamed.VB_Description = "Occurs after a theme has ben renamed."
Event ThemeRemoved()
Attribute ThemeRemoved.VB_Description = "Occurs after a theme has been removed."

Private mThemes As Collection
Private mDefaultThemesCopied As Boolean

Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of themes in the collection."
Attribute Count.VB_UserMemId = -531
    EnsureDefaultThemesCopied
    Count = mThemes.Count
End Property

Private Sub Class_Initialize()
    Set mThemes = New Collection
    mDefaultThemesCopied = False
End Sub

Private Sub Class_Terminate()
    Set mThemes = Nothing
End Sub

Public Property Get Item(ByVal IndexOrKey As Variant) As NewTabTheme
Attribute Item.VB_Description = "Returns a NewTabTheme object."
Attribute Item.VB_UserMemId = 0
    EnsureDefaultThemesCopied
    If VarType(IndexOrKey) = vbString Then
        IndexOrKey = Trim$(IndexOrKey)
    End If
    On Error Resume Next
    Set Item = mThemes(IndexOrKey)
    If Err.Number Then
        Err.Clear
        Err.Raise 5  ' Invalid procedure call or argument
        Dim iStr As String
        iStr = Err.Description
        On Error GoTo 0
        Err.Raise 5, App.Title & ", " & TypeName(Me), iStr & ". Theme doesn't exist."
    End If
End Property

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    EnsureDefaultThemesCopied
    Set NewEnum = mThemes.[_NewEnum]
End Function

Friend Sub Add(ByVal nTheme As NewTabTheme)
    EnsureDefaultThemesCopied
    mThemes.Add nTheme, nTheme.Name
End Sub

Public Function Exists(ByVal nThemeName As String) As Boolean
Attribute Exists.VB_Description = "Returns True if the theme exists in the collection."
    Dim iTheme As NewTabTheme
    
    nThemeName = Trim$(nThemeName)
    EnsureDefaultThemesCopied
    On Error Resume Next
    Set iTheme = mThemes(nThemeName)
    Exists = Not iTheme Is Nothing
End Function

Public Sub Remove(ByVal nThemeName As String)
Attribute Remove.VB_Description = "Removes a theme from the collection."
    EnsureDefaultThemesCopied
    
    nThemeName = Trim$(nThemeName)
    If Not Exists(nThemeName) Then
        Err.Clear
        Err.Raise 5  ' Invalid procedure call or argument
        Dim iStr As String
        iStr = Err.Description
        On Error GoTo 0
        Err.Raise 5, App.Title & ", " & TypeName(Me), iStr & ". Theme doesn't exist."
        Exit Sub
    End If
    mThemes.Remove nThemeName
    RaiseEvent ThemeRemoved
End Sub

Public Sub Rename(ByVal nOldThemeName As String, ByVal nNewThemeName As String)
Attribute Rename.VB_Description = "Renames a theme."
    Dim iTheme As NewTabTheme
    Dim iStr As String
    
    nOldThemeName = Trim$(nOldThemeName)
    nNewThemeName = Trim$(nNewThemeName)
    EnsureDefaultThemesCopied
    If nOldThemeName = nNewThemeName Then Exit Sub
    If Not Exists(nOldThemeName) Then
        Err.Clear
        Err.Raise 5  ' Invalid procedure call or argument
        iStr = Err.Description
        On Error GoTo 0
        Err.Raise 5, App.Title & ", " & TypeName(Me), iStr & ". Theme doesn't exist."
        Exit Sub
    End If
    If LCase$(nOldThemeName) <> LCase$(nNewThemeName) Then
        If Exists(nNewThemeName) Then
            Err.Clear
            Err.Raise 5  ' Invalid procedure call or argument
            iStr = Err.Description
            On Error GoTo 0
            Err.Raise 5, App.Title & ", " & TypeName(Me), iStr & ". New name already exist."
            Exit Sub
        End If
    End If
    
    Set iTheme = mThemes(nOldThemeName)
    Remove nOldThemeName
    iTheme.Name = nNewThemeName
    Add iTheme
    RaiseEvent ThemeRenamed
End Sub

Private Sub EnsureDefaultThemesCopied()
    Dim iDefThemes As NewTabThemes
    Dim iTheme As NewTabTheme
    Dim iTheme2 As NewTabTheme
    
    If Not mDefaultThemesCopied Then
        Set iDefThemes = GetDefaultThemes
        mDefaultThemesCopied = True
        For Each iTheme In iDefThemes
            Set iTheme2 = iTheme.Clone
            Add iTheme2
        Next
    End If
End Sub

Friend Function Clone() As NewTabThemes
    Dim iTheme As NewTabTheme
    
    Set Clone = New NewTabThemes
    For Each iTheme In mThemes
        If iTheme.Custom Then
            Clone.Add iTheme.Clone
        End If
    Next
End Function

Friend Function ThereAreCustomThemes() As Boolean
    If mThemes.Count > 0 Then
        ThereAreCustomThemes = mThemes(mThemes.Count).Custom
    End If
End Function

Friend Property Let DoNotCopyDefaultThemes(nValue As Boolean)
    mDefaultThemesCopied = nValue
End Property
