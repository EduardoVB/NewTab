Attribute VB_Name = "mNewTabThemes"
Option Explicit

Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, Col As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetProp Lib "user32.dll" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long

Public gPagTabsApplyTime As Double
Public gPagTabsLastSelectedTab As Double

Private mDefaultThemes As NewTabThemes
Private mThemesFile As NewTabThemes
Private mThemesRegistry As NewTabThemes
Private mThemesLoaded As Boolean
Private mControlTypeName As String

Private mAuxAmbientBackColor As Long
Private mAuxAmbientForeColor As Long
Private mAuxThemeData As Collection

Public Const cThemeIsCustomSettings As String = "Current custom settings, not saved as a Theme"

' Default Themes strings
Private Const cThemeString_Default                       As String = ""
Private Const cThemeString_SSTab                         As String = "HighlightEffect=0|HighlightMode=1|HighlightModeSelectedTab=16|ShowFocusRect=-1|SoftEdges=0|Style=0"
Private Const cThemeString_SSTabWindows                  As String = "HighlightModeSelectedTab=16|ShowFocusRect=-1"
Private Const cThemeString_SSTabPropertyPage             As String = "HighlightEffect=0|HighlightMode=1|HighlightModeSelectedTab=1|ShowFocusRect=-1|SoftEdges=0|Style=1"
Private Const cThemeString_SSTabPropertyPageWindows      As String = "ShowFocusRect=-1|ShowRowsInPerspective=1|TabWidthStyle=1"
Private Const cThemeString_TabStrip                      As String = "HighlightEffect=0|HighlightMode=1|HighlightModeSelectedTab=1|ShowFocusRect=-1|SoftEdges=0|Style=2"
Private Const cThemeString_TabStripWindows               As String = "ShowFocusRect=-1|TabWidthStyle=0"
Private Const cThemeString_FlatSilver                    As String = "BackColorTabs=15658734|FlatBarColorSelectedTab=14181684|HighlightMode=2|HighlightModeSelectedTab=66|Style=3"
Private Const cThemeString_FlatBronze                    As String = "BackColorSelectedTab=16383485|BackColorTabs=14611960|FlatBarColorHighlight=3431538|FlatBarColorInactive=13559786|FlatBarColorSelectedTab=1148870|FlatBorderColor=1148870|FlatBorderMode=1|ForeColorHighlighted=16777215|HighlightColor=3431538|HighlightEffect=0|HighlightMode=68|HighlightModeSelectedTab=90|IconColorTabHighlighted=16777215|Style=3"
Private Const cThemeString_FlatAppleGreen                As String = "BackColorSelectedTab=16514553|BackColorTabs=15136990|FlatBarColorHighlight=3633716|FlatBarColorInactive=14150350|FlatBarColorSelectedTab=1820177|FlatBorderColor=1820177|FlatBorderMode=1|ForeColorHighlighted=16777215|HighlightColor=3633716|HighlightEffect=0|HighlightMode=68|HighlightModeSelectedTab=90|IconColorTabHighlighted=16777215|Style=3"
Private Const cThemeString_FlatGolden                    As String = "BackColorSelectedTab=16777215|BackColorTabs=15202556|FlatBarColorHighlight=3530228|FlatBarColorInactive=13559786|FlatBarColorSelectedTab=768981|HighlightColor=3530228|HighlightMode=76|HighlightModeSelectedTab=90|IconColorTabHighlighted=12664841|Style=3"
Private Const cThemeString_FlatSeaBlue                   As String = "BackColorSelectedTab=16250871|FlatBarColorSelectedTab=16546371|FlatBarHeight=0|FlatBorderColor=10184001|FlatBorderMode=1|FlatRoundnessTabs=8|FlatTabsSeparationLineColor=-2147483633|ForeColorHighlighted=16777215|ForeColorSelectedTab=10184001|HighlightColor=16477710|HighlightMode=4|HighlightModeSelectedTab=10|IconColorTabHighlighted=16777215|IconColorSelectedTab=10184001|Style=3|TabMousePointerHand=-1"
Private Const cThemeString_FlatEmerald                   As String = "BackColorSelectedTab=2183936|BackColorTabs=1588736|FlatBarColorHighlight=9615225|FlatBarColorInactive=4422175|FlatBarColorSelectedTab=16777215|FlatBodySeparationLineColor=1983492|FlatBorderColor=3960091|FlatTabsSeparationLineColor=1983492|ForeColor=16777215|HighlightColor=5281568|HighlightColorSelectedTab=5942308|HighlightMode=64|HighlightModeSelectedTab=90|Style=3|TabMousePointerHand=-1"
Private Const cThemeString_FlatRedWine                   As String = "BackColorSelectedTab=2424902|BackColorTabs=1835059|FlatBarColorHighlight=8542630|FlatBarColorInactive=4332134|FlatBarColorSelectedTab=16777215|FlatBodySeparationLineColor=2032442|FlatBorderColor=3937884|FlatTabsSeparationLineColor=2032442|ForeColor=16777215|HighlightColor=5184382|HighlightColorSelectedTab=5906065|HighlightMode=64|HighlightModeSelectedTab=90|Style=3|TabMousePointerHand=-1"
Private Const cThemeString_FlatDeepWaters                As String = "BackColorSelectedTab=6699520|BackColorTabs=5057536|FlatBarColorHighlight=16777215|FlatBarColorInactive=9856549|FlatBarColorSelectedTab=16777215|FlatBodySeparationLineColor=5057536|FlatBorderColor=8870689|FlatTabsSeparationLineColor=5386501|ForeColor=16777215|HighlightColor=7554571|HighlightColorSelectedTab=13863980|HighlightMode=76|HighlightModeSelectedTab=90|Style=3|TabMousePointerHand=-1"
Private Const cThemeString_FlatOpenAir                   As String = "BackColorTabs=16250871|FlatBarColorHighlight=16250871|FlatBarColorInactive=16250871|FlatBarColorSelectedTab=16731706|FlatBarHeight=4|FlatBarPosition=1|FlatBodySeparationLineColor=14869218|FlatBorderColor=16250871|FlatRoundnessBottom=0|FlatRoundnessTop=0|FlatTabsSeparationLineColor=16250871|ForeColorHighlighted=16731706|ForeColorSelectedTab=16731706|HighlightColor=16477710|HighlightMode=1|HighlightModeSelectedTab=64|IconColorTabHighlighted=16731706|IconColorSelectedTab=16731706|Style=3|TabMousePointerHand=-1"
Private Const cThemeString_GhostTab                      As String = "BackColorSelectedTab=16250871|FlatBarColorSelectedTab=7699508|FlatBarHeight=0|FlatBodySeparationLineColor=7699508|FlatBodySeparationLineHeight=3|FlatBorderColor=7699508|FlatBorderMode=1|FlatRoundnessTabs=8|FlatTabsSeparationLineColor=11250603|ForeColor=0|ForeColorSelectedTab=16777215|HighlightColor=12766860|HighlightColorSelectedTab=7699508|HighlightMode=4|HighlightModeSelectedTab=20|IconColorSelectedTab=16777215|Style=3|TabMousePointerHand=-1|TabSeparation=8"
Private Const cThemeString_Buttons                       As String = "BackColorSelectedTab=-2147483633|BackColorTabs=12766860|FlatBarColorHighlight=7699508|FlatBarColorInactive=12766860|FlatBarColorSelectedTab=7699508|FlatBarPosition=1|FlatBodySeparationLineColor=11250603|FlatBodySeparationLineHeight=0|FlatBorderColor=-2147483633|FlatRoundnessBottom=0|FlatRoundnessTop=0|FlatTabsSeparationLineColor=11250603|HighlightColor=12766860|HighlightColorSelectedTab=-2147483633|HighlightMode=196|HighlightModeSelectedTab=196|Style=3|TabMousePointerHand=-1|TabSeparation=8"
Private Const cThemeString_Buttons2                      As String = "BackColorSelectedTab=-2147483633|BackColorTabs=12766860|FlatBarColorHighlight=7699508|FlatBarColorInactive=12766860|FlatBarColorSelectedTab=7699508|FlatBarPosition=1|FlatBodySeparationLineColor=11250603|FlatBodySeparationLineHeight=0|FlatBorderColor=-2147483633|FlatRoundnessBottom=0|FlatRoundnessTabs=16|FlatRoundnessTop=0|FlatTabsSeparationLineColor=11250603|HighlightColor=12766860|HighlightColorSelectedTab=-2147483633|HighlightMode=196|HighlightModeSelectedTab=196|Style=3|TabMousePointerHand=-1|TabSeparation=8"
Private Const cThemeString_Buttons3                      As String = "BackColorSelectedTab=-2147483633|BackColorTabs=12766860|FlatBarColorHighlight=7699508|FlatBarColorInactive=12766860|FlatBarColorSelectedTab=7699508|FlatBodySeparationLineColor=7699508|FlatBodySeparationLineHeight=3|FlatBorderColor=-2147483633|FlatRoundnessBottom=0|FlatRoundnessTop=0|FlatTabsSeparationLineColor=11250603|HighlightColor=12766860|HighlightColorSelectedTab=-2147483633|HighlightMode=196|HighlightModeSelectedTab=1|Style=3|TabMousePointerHand=-1|TabSeparation=8"
Private Const cThemeString_Buttons4                      As String = "BackColorSelectedTab=-2147483633|BackColorTabs=12766860|FlatBarColorHighlight=7699508|FlatBarColorInactive=12766860|FlatBarColorSelectedTab=7699508|FlatBarHeight=5|FlatBodySeparationLineColor=7699508|FlatBorderColor=7699508|FlatBorderMode=1|FlatRoundnessBottom=4|FlatRoundnessTabs=4|FlatRoundnessTop=4|FlatTabsSeparationLineColor=-2147483633|HighlightColor=12766860|HighlightColorSelectedTab=-2147483633|HighlightMode=196|HighlightModeSelectedTab=64|Style=3|TabMousePointerHand=-1|TabSeparation=8"
Private Const cThemeString_Buttons5                      As String = "BackColorSelectedTab=-2147483633|BackColorTabs=12766860|FlatBarColorHighlight=7699508|FlatBarColorInactive=12766860|FlatBarColorSelectedTab=7699508|FlatBarGripHeight=0|FlatBarHeight=5|FlatBodySeparationLineColor=7699508|FlatBorderColor=7699508|FlatBorderMode=1|FlatRoundnessBottom=4|FlatRoundnessTabs=4|FlatRoundnessTop=4|FlatTabsSeparationLineColor=-2147483633|HighlightColor=12766860|HighlightColorSelectedTab=-2147483633|HighlightMode=196|HighlightModeSelectedTab=64|Style=3|TabMousePointerHand=-1|TabSeparation=8"
Private Const cThemeString_Buttons6                      As String = "BackColorSelectedTab=16777215|BackColorTabs=12766860|FlatBarColorHighlight=7699508|FlatBarColorInactive=12766860|FlatBarColorSelectedTab=7699508|FlatBarGripHeight=0|FlatBarHeight=5|FlatBodySeparationLineColor=7699508|FlatBorderColor=7699508|FlatBorderMode=1|FlatRoundnessBottom=4|FlatRoundnessTabs=4|FlatRoundnessTop=4|FlatTabsSeparationLineColor=-2147483633|HighlightColor=12766860|HighlightColorSelectedTab=-2147483633|HighlightMode=196|HighlightModeSelectedTab=64|Style=3|TabMousePointerHand=-1|TabSeparation=8"
Private Const cThemeString_Buttons7                      As String = "BackColorSelectedTab=-2147483633|BackColorTabs=10646340|FlatBarColorHighlight=4228799|FlatBarColorInactive=10646340|FlatBarColorSelectedTab=10646340|FlatBarPosition=1|FlatBodySeparationLineColor=7434609|FlatBodySeparationLineHeight=0|FlatBorderColor=-2147483633|FlatRoundnessBottom=0|FlatRoundnessTop=0|FlatTabsSeparationLineColor=7434609|ForeColor=16777215|ForeColorSelectedTab=0|HighlightColor=10646340|HighlightColorSelectedTab=-2147483633|HighlightMode=196|HighlightModeSelectedTab=196|IconColorSelectedTab=0|Style=3|TabMousePointerHand=-1|TabSeparation=8"
Private Const cThemeString_Buttons8                      As String = "BackColorSelectedTab=-2147483633|BackColorTabs=10316098|FlatBarColorHighlight=1729514|FlatBarColorInactive=10316098|FlatBarColorSelectedTab=10316098|FlatBarPosition=1|FlatBodySeparationLineColor=7171437|FlatBodySeparationLineHeight=0|FlatBorderColor=-2147483633|FlatRoundnessBottom=0|FlatRoundnessTabs=8|FlatRoundnessTop=0|FlatTabsSeparationLineColor=7171437|ForeColor=16777215|ForeColorSelectedTab=0|HighlightColor=10316098|HighlightColorSelectedTab=-2147483633|HighlightMode=196|HighlightModeSelectedTab=196|IconColorSelectedTab=0|Style=3|TabMousePointerHand=-1|TabSeparation=8|TabWidthStyle=1"
Private Const cThemeString_Buttons9                      As String = "BackColorSelectedTab=-2147483633|BackColorTabs=8804407|FlatBarColorHighlight=1729514|FlatBarColorInactive=8804407|FlatBarColorSelectedTab=8804407|FlatBarPosition=1|FlatBodySeparationLineColor=6184542|FlatBodySeparationLineHeight=0|FlatBorderColor=-2147483633|FlatRoundnessBottom=0|FlatRoundnessTabs=8|FlatRoundnessTop=0|FlatTabsSeparationLineColor=6184542|ForeColor=16777215|ForeColorSelectedTab=0|HighlightColor=8804407|HighlightColorSelectedTab=-2147483633|HighlightMode=196|HighlightModeSelectedTab=196|IconColorSelectedTab=0|Style=3|TabMousePointerHand=-1|TabSeparation=8|TabWidthStyle=1"
Private Const cThemeString_Buttons10                     As String = "BackColorSelectedTab=16775410|BackColorTabs=7742876|FlatBarColorHighlight=5751007|FlatBarColorInactive=7742876|FlatBarColorSelectedTab=7742876|FlatBarPosition=1|FlatBodySeparationLineColor=7829367|FlatBodySeparationLineHeight=0|FlatBorderColor=-2147483633|FlatRoundnessBottom=0|FlatRoundnessTabs=8|FlatRoundnessTop=0|FlatTabsSeparationLineColor=7829367|ForeColor=16777215|ForeColorSelectedTab=0|HighlightColor=7742876|HighlightColorSelectedTab=-2147483633|HighlightMode=196|HighlightModeSelectedTab=196|IconColorSelectedTab=0|Style=3|TabMousePointerHand=-1|TabSeparation=8|TabWidthStyle=1"
Private Const cThemeString_Buttons11                     As String = "BackColorSelectedTab=16775410|BackColorTabs=6955405|FlatBarColorHighlight=1283056|FlatBarColorInactive=6955405|FlatBarColorSelectedTab=6955405|FlatBarPosition=1|FlatBodySeparationLineColor=7039851|FlatBodySeparationLineHeight=0|FlatBorderColor=-2147483633|FlatRoundnessBottom=0|FlatRoundnessTabs=8|FlatRoundnessTop=0|FlatTabsSeparationLineColor=7039851|ForeColor=16777215|ForeColorSelectedTab=0|HighlightColor=6955405|HighlightColorSelectedTab=-2147483633|HighlightMode=196|HighlightModeSelectedTab=196|IconColorSelectedTab=0|Style=3|TabMousePointerHand=-1|TabSeparation=8|TabWidthStyle=1"
Private Const cThemeString_WebLinks                      As String = "BackColorTabs=16250871|FlatBarColorHighlight=16250871|FlatBarColorInactive=16250871|FlatBarColorSelectedTab=16731706|FlatBarHeight=4|FlatBarPosition=1|FlatBodySeparationLineColor=14869218|FlatBorderColor=16250871|FlatRoundnessBottom=0|FlatRoundnessTop=0|FlatTabsSeparationLineColor=16250871|ForeColorHighlighted=16731706|ForeColorSelectedTab=16777215|HighlightColor=16477710|HighlightColorSelectedTab=16477710|HighlightMode=32|HighlightModeSelectedTab=20|IconColorTabHighlighted=16731706|IconColorSelectedTab=16777215|Style=3|TabMousePointerHand=-1"
Private Const cThemeString_WebLinks2                     As String = "BackColorTabs=16250871|FlatBarColorHighlight=16250871|FlatBarColorInactive=16250871|FlatBarColorSelectedTab=16731706|FlatBarHeight=4|FlatBarPosition=1|FlatBodySeparationLineColor=14869218|FlatBorderColor=16731706|FlatBorderMode=1|FlatRoundnessTabs=8|FlatTabsSeparationLineColor=16250871|ForeColorHighlighted=16731706|ForeColorSelectedTab=16777215|HighlightColor=16477710|HighlightColorSelectedTab=16477710|HighlightMode=16|HighlightModeSelectedTab=20|IconColorTabHighlighted=16731706|IconColorSelectedTab=16777215|Style=3|TabMousePointerHand=-1"
Private Const cThemeString_AnotherButtonLight            As String = "BackColor=16250871|FlatBarColorSelectedTab=16546371|FlatBarHeight=0|FlatBodySeparationLineColor=16777215|FlatBodySeparationLineHeight=0|FlatBorderColor=16777215|FlatBorderMode=1|FlatRoundnessBottom=0|FlatRoundnessTabs=4|FlatRoundnessTop=0|FlatTabBorderColorHighlight=8603431|FlatTabBorderColorSelectedTab=15492185|FlatTabsSeparationLineColor=16777215|ForeColor=0|HighlightColor=15461355|HighlightMode=12|HighlightModeSelectedTab=512|Style=3|TabMousePointerHand=-1"
Private Const cThemeString_AnotherButtonDark             As String = "BackColor=3476744|BackColorSelectedTab=592137|FlatBarColorSelectedTab=12335619|FlatBarHeight=0|FlatBodySeparationLineHeight=0|FlatBorderColor=3476744|FlatBorderMode=1|FlatRoundnessBottom=0|FlatRoundnessTabs=4|FlatRoundnessTop=0|FlatTabBorderColorHighlight=14195837|FlatTabBorderColorSelectedTab=14195837|FlatTabsSeparationLineColor=3476744|ForeColor=16777215|HighlightColor=4210752|HighlightMode=12|HighlightModeSelectedTab=512|Style=3|TabMousePointerHand=-1"

Private Function ClientProjectFolder() As String
    Dim iInIDE As Boolean
    Static sValue As String
    
    If sValue = "" Then
        Debug.Assert MakeTrue(iInIDE)
        If iInIDE Then
            sValue = GetClientProjectFolder
        End If
    End If
    ClientProjectFolder = sValue
End Function

Private Function GetClientProjectFolder() As String
    Dim hwndMain As Long
    Dim hProp As Long
    Dim iObjIDE As Object
    Dim iObjVBE As Object
    
    hwndMain = FindWindow("wndclass_desked_gsk", vbNullString)
    If hwndMain <> 0 Then
        hProp = GetProp(hwndMain, "VBAutomation")
        If hProp <> 0 Then
            CopyMemory iObjIDE, hProp, 4&    '= VBIDE.Window
            On Error Resume Next
            Set iObjVBE = iObjIDE.VBE
            GetClientProjectFolder = iObjVBE.ActiveVBProject.FileName
            If InStr(GetClientProjectFolder, "\") > 2 Then
                GetClientProjectFolder = Left$(GetClientProjectFolder, InStrRev(GetClientProjectFolder, "\"))
            End If
            On Error GoTo 0
            CopyMemory iObjIDE, 0&, 4&
        End If
    End If
End Function
    
Private Function MakeTrue(Value As Boolean) As Boolean
    MakeTrue = True
    Value = True
End Function

Private Function ConfigFilePathInProjectFolder() As String
    Static sValue As String
    
    If sValue = "" Then
        sValue = ClientProjectFolder & ConfigFileName
    End If
    ConfigFilePathInProjectFolder = sValue
End Function

Public Function ConfigFileName() As String
    ConfigFileName = App.Title & "Themes.ntt"
End Function

Private Sub EnsureThemesLoaded()
    If Not mThemesLoaded Then
        If mControlTypeName = "" Then Err.Raise 5558
        LoadThemesFromRegistry
        LoadThemesFromFile
        mThemesLoaded = True
    End If
End Sub

Private Sub LoadThemesFromRegistry()
    Dim iStr As String
    Dim s1() As String
    Dim s2() As String
    Dim c1 As Long
    Dim iTheme As NewTabTheme
    
    Set mThemesRegistry = New NewTabThemes
    mThemesRegistry.DoNotCopyDefaultThemes = True
    iStr = GetSetting(mControlTypeName, "Themes", "Data")
    If iStr <> "" Then
        s1 = Split(iStr, "\") ' get individual Theme data, tuples of Theme name and Theme data
        For c1 = 0 To UBound(s1)
            s2 = Split(s1(c1), ":") ' Theme name : Theme data
            If UBound(s2) = 1 Then
                If (s2(0) <> "") And (s2(1) <> "") Then
                    Set iTheme = New NewTabTheme
                    iTheme.Name = s2(0)
                    iTheme.ThemeString = s2(1)
                    iTheme.Custom = True
                    If Not mThemesRegistry.Exists(iTheme.Name) Then
                        mThemesRegistry.Add iTheme
                    End If
                End If
            End If
        Next
    End If
End Sub

Private Sub LoadThemesFromFile()
    Dim iStr As String
    Dim s1() As String
    Dim s2() As String
    Dim c1 As Long
    Dim iTheme As NewTabTheme
    
    Set mThemesFile = New NewTabThemes
    mThemesFile.DoNotCopyDefaultThemes = True
    iStr = LoadTextFile(ConfigFilePathInProjectFolder)
    If iStr = "" Then
        iStr = GetSetting(mControlTypeName, "Themes", "Data")
    End If
    If iStr <> "" Then
        s1 = Split(iStr, "\") ' get individual Theme data, tuples of Theme name and Theme data
        For c1 = 0 To UBound(s1)
            s2 = Split(s1(c1), ":") ' Theme name : Theme data
            If UBound(s2) = 1 Then
                If (s2(0) <> "") And (s2(1) <> "") Then
                    Set iTheme = New NewTabTheme
                    iTheme.Name = s2(0)
                    iTheme.ThemeString = s2(1)
                    iTheme.Custom = True
                    If Not mThemesFile.Exists(iTheme.Name) Then
                        mThemesFile.Add iTheme
                    End If
                End If
            End If
        Next
    End If
End Sub

Public Function GetThemesRegistry() As Collection
    Dim iTheme As NewTabTheme
    
    EnsureThemesLoaded
    Set GetThemesRegistry = New Collection
    For Each iTheme In mThemesRegistry
        GetThemesRegistry.Add iTheme, iTheme.Name
    Next
End Function

Public Function GetThemesFile() As Collection
    Dim iTheme As NewTabTheme
    
    EnsureThemesLoaded
    Set GetThemesFile = New Collection
    For Each iTheme In mThemesFile
        GetThemesFile.Add iTheme, iTheme.Name
    Next
End Function

Public Sub SaveThemesInRegistry(nThemes As NewTabThemes)
    Dim iStr As String
    Dim iTheme As NewTabTheme
    
    Set mThemesRegistry = nThemes
    For Each iTheme In mThemesRegistry
        iStr = iStr & IIf(iStr = "", "", "\") & iTheme.Name & ":" & iTheme.ThemeString
    Next
    If iStr = "" Then
        DeleteSetting mControlTypeName, "Themes", "Data"
    Else
        SaveSetting mControlTypeName, "Themes", "Data", iStr
    End If
End Sub

Public Sub SaveThemesInFile(nThemes As NewTabThemes)
    Dim iStr As String
    Dim iTheme As NewTabTheme
    
    Set mThemesFile = nThemes
    For Each iTheme In mThemesFile
        iStr = iStr & IIf(iStr = "", "", "\") & iTheme.Name & ":" & iTheme.ThemeString
    Next
    If iStr = "" Then
        If FileExists(ConfigFilePathInProjectFolder) Then
            On Error Resume Next
            Kill ConfigFilePathInProjectFolder
            On Error GoTo 0
        End If
    Else
        If FileExists(ConfigFilePathInProjectFolder) Then
            On Error Resume Next
            Kill ConfigFilePathInProjectFolder
            On Error GoTo 0
        End If
        On Error Resume Next
        SaveTextFile ConfigFilePathInProjectFolder, iStr
        On Error GoTo 0
    End If
End Sub

Private Function FileExists(ByVal strPathName As String) As Boolean
    Dim intFileNum As Integer

    On Error Resume Next
    intFileNum = FreeFile
    Open strPathName For Input As intFileNum
    FileExists = (Err.Number = 0) Or (Err.Number = 70) Or (Err.Number = 55)
    Close intFileNum
    Err.Clear
End Function

Private Sub SaveTextFile(nPath As String, nText As String)
    Dim iFreeFile As Long
    
    If FileExists(nPath) Then
        Err.Raise 867, , "File already exists"
    Else
        On Error Resume Next
        iFreeFile = FreeFile
        Open nPath For Output As #iFreeFile
        Print #iFreeFile, nText
        Close #iFreeFile
    End If
End Sub

Private Function LoadTextFile(nFilePath As String) As String
    Dim iFile As Long
    
    If FileExists(nFilePath) Then
        iFile = FreeFile
        Open nFilePath For Input Access Read As #iFile
        If LOF(iFile) > 0 Then
            LoadTextFile = Input(LOF(iFile), iFile)
        End If
        Close #iFile
    End If
End Function

Public Sub SetControlTypeName(nName As String)
    mControlTypeName = nName
End Sub

Public Sub CopyControlProperties(nCtlSrc As NewTab, nCtlDest As NewTab, Optional nAmbientFont As StdFont)
    Dim iRedraw As Boolean
    
    iRedraw = nCtlDest.Redraw
    nCtlDest.Redraw = False
    
    nCtlDest.SetDefaultPropertyValuesForThemedProperties True
    
    nCtlDest.Style = nCtlSrc.Style
    nCtlDest.BackColor = nCtlSrc.BackColor
    nCtlDest.BackColorTabs = nCtlSrc.BackColorTabs
    If Not nCtlSrc.BackColorSelectedTab_IsAutomatic Then nCtlDest.BackColorSelectedTab = nCtlSrc.BackColorSelectedTab
     
    nCtlDest.ForeColor = nCtlSrc.ForeColor
    
    nCtlDest.ForeColorSelectedTab = nCtlSrc.ForeColorSelectedTab
    nCtlDest.ForeColorHighlighted = nCtlSrc.ForeColorHighlighted
    nCtlDest.FlatTabBorderColorHighlight = nCtlSrc.FlatTabBorderColorHighlight
    nCtlDest.FlatTabBorderColorSelectedTab = nCtlSrc.FlatTabBorderColorSelectedTab
    nCtlDest.IconColor = nCtlSrc.IconColor
    nCtlDest.IconColorSelectedTab = nCtlSrc.IconColorSelectedTab
    nCtlDest.IconColorTabHighlighted = nCtlSrc.IconColorTabHighlighted
     
    nCtlDest.FlatBarColorSelectedTab = nCtlSrc.FlatBarColorSelectedTab
    If Not nCtlSrc.FlatBarColorInactive_IsAutomatic Then nCtlDest.FlatBarColorInactive = nCtlSrc.FlatBarColorInactive
    If Not nCtlSrc.FlatTabsSeparationLineColor_IsAutomatic Then nCtlDest.FlatTabsSeparationLineColor = nCtlSrc.FlatTabsSeparationLineColor
    If Not nCtlSrc.FlatBodySeparationLineColor_IsAutomatic Then nCtlDest.FlatBodySeparationLineColor = nCtlSrc.FlatBodySeparationLineColor
    If Not nCtlSrc.FlatBorderColor_IsAutomatic Then nCtlDest.FlatBorderColor = nCtlSrc.FlatBorderColor
    If Not nCtlSrc.FlatBarColorHighlight_IsAutomatic Then nCtlDest.FlatBarColorHighlight = nCtlSrc.FlatBarColorHighlight
    If Not nCtlSrc.HighlightColor_IsAutomatic Then nCtlDest.HighlightColor = nCtlSrc.HighlightColor
    If Not nCtlSrc.HighlightColorSelectedTab_IsAutomatic Then nCtlDest.HighlightColorSelectedTab = nCtlSrc.HighlightColorSelectedTab
    
    nCtlDest.TabAppearance = nCtlSrc.TabAppearance
    nCtlDest.ShowRowsInPerspective = nCtlSrc.ShowRowsInPerspective
    nCtlDest.BackStyle = nCtlSrc.BackStyle
    nCtlDest.HighlightMode = nCtlSrc.HighlightMode
    nCtlDest.HighlightModeSelectedTab = nCtlSrc.HighlightModeSelectedTab
    nCtlDest.AutoRelocateControls = nCtlSrc.AutoRelocateControls
    nCtlDest.FlatBorderMode = nCtlSrc.FlatBorderMode
    nCtlDest.FlatBarPosition = nCtlSrc.FlatBarPosition
    nCtlDest.FlatBarHeight = nCtlSrc.FlatBarHeight
    nCtlDest.FlatBarGripHeight = nCtlSrc.FlatBarGripHeight
    nCtlDest.FlatBodySeparationLineHeight = nCtlSrc.FlatBodySeparationLineHeight
    nCtlDest.TabSeparation = nCtlSrc.TabSeparation
    nCtlDest.HighlightTabExtraHeight = nCtlSrc.HighlightTabExtraHeight
    nCtlDest.TabMaxWidth = nCtlSrc.TabMaxWidth
    nCtlDest.TabMinWidth = nCtlSrc.TabMinWidth
    nCtlDest.FlatRoundnessTop = nCtlSrc.FlatRoundnessTop
    nCtlDest.FlatRoundnessBottom = nCtlSrc.FlatRoundnessBottom
    nCtlDest.FlatRoundnessTabs = nCtlSrc.FlatRoundnessTabs
    nCtlDest.ShowDisabledState = nCtlSrc.ShowDisabledState
    nCtlDest.ShowFocusRect = nCtlSrc.ShowFocusRect
    nCtlDest.HighlightEffect = nCtlSrc.HighlightEffect
    nCtlDest.WordWrap = nCtlSrc.WordWrap
    nCtlDest.SoftEdges = nCtlSrc.SoftEdges
    nCtlDest.TabMousePointerHand = nCtlSrc.TabMousePointerHand
    nCtlDest.MaskColor = nCtlSrc.MaskColor
    If FontsAreEqual(nCtlSrc.AmbientFont, nCtlSrc.Font) Then
        Set nCtlDest.Font = nCtlDest.AmbientFont
    Else
        Set nCtlDest.Font = CloneFont(nCtlSrc.Font)
    End If
    If (nCtlDest.TDIMode <> ntTDIModeNone) = (nCtlSrc.TDIMode <> ntTDIModeNone) Then
        nCtlDest.IconColorMouseHover = nCtlSrc.IconColorMouseHover
        nCtlDest.IconColorMouseHoverSelectedTab = nCtlSrc.IconColorMouseHoverSelectedTab
        nCtlDest.TabWidthStyle = nCtlSrc.TabWidthStyle
    End If
    nCtlDest.Redraw = iRedraw
End Sub

Private Function CloneFont(nOrigFont As iFont) As StdFont
    If nOrigFont Is Nothing Then Exit Function
    nOrigFont.Clone CloneFont
End Function

Public Function GetDefaultThemes() As NewTabThemes
    If mDefaultThemes Is Nothing Then
        LoadDefaultThemes
    End If
    Set GetDefaultThemes = mDefaultThemes
End Function

Private Function PropertyExists(nPropertyName As String) As Boolean
    Dim iProp As cPropertyData
    
    On Error Resume Next
    Set iProp = mAuxThemeData(nPropertyName)
    On Error GoTo 0
    PropertyExists = Not iProp Is Nothing
End Function
    
Private Function GetPropertyValue(nPropertyName As String) As Variant
    Dim iProp As cPropertyData
    
    Set iProp = mAuxThemeData(nPropertyName)
    If iProp.Value = "A" Then
        GetPropertyValue = -1
    ElseIf iProp.Value = "B" Then
        GetPropertyValue = mAuxAmbientBackColor
    ElseIf iProp.Value = "F" Then
        GetPropertyValue = mAuxAmbientForeColor
    Else
        GetPropertyValue = Val(iProp.Value)
    End If
End Function
    
Public Sub ApplyThemeToControl(ByRef nThemeData As Collection, nCtl As NewTab, nAmbientBackColor As Long, nAmbientForeColor As Long)
    Dim iProp As cPropertyData
    Dim iRedraw As Boolean
    
    mAuxAmbientBackColor = nAmbientBackColor
    mAuxAmbientForeColor = nAmbientForeColor
    Set mAuxThemeData = nThemeData
    
    iRedraw = nCtl.Redraw
    nCtl.Redraw = False
    nCtl.SetDefaultPropertyValuesForThemedProperties True
    
    If PropertyExists("BackColor") Then nCtl.BackColor = GetPropertyValue("BackColor")
    If PropertyExists("ForeColor") Then nCtl.ForeColor = GetPropertyValue("ForeColor")
    
    Set iProp = Nothing
    On Error Resume Next
    Set iProp = nThemeData("IconColorSelectedTab")
    On Error GoTo 0
    If iProp Is Nothing Then
        On Error Resume Next
        Set iProp = nThemeData("IconColorTabSel")
        On Error GoTo 0
    End If
    If iProp Is Nothing Then
        nCtl.IconColorSelectedTab = nCtl.ForeColor
    Else
        If iProp.Value = "F" Then nCtl.IconColorSelectedTab = nAmbientForeColor Else nCtl.IconColorSelectedTab = Val(iProp.Value)
    End If
    
    Set iProp = Nothing
    On Error Resume Next
    Set iProp = nThemeData("IconColorTabHighlighted")
    On Error GoTo 0
    If iProp Is Nothing Then
        nCtl.IconColorTabHighlighted = nCtl.ForeColor
    Else
        If iProp.Value = "F" Then nCtl.IconColorTabHighlighted = nAmbientForeColor Else nCtl.IconColorTabHighlighted = Val(iProp.Value)
    End If
    If nCtl.TDIMode = ntTDIModeNone Then
        If iProp Is Nothing Then
            nCtl.IconColorMouseHover = nCtl.ForeColor
        Else
            If iProp.Value = "F" Then nCtl.IconColorMouseHover = nAmbientForeColor Else nCtl.IconColorMouseHover = Val(iProp.Value)
        End If
        Set iProp = Nothing
        On Error Resume Next
        Set iProp = nThemeData("ForeColorSelectedTab")
        On Error GoTo 0
        If iProp Is Nothing Then
            On Error Resume Next
            Set iProp = nThemeData("ForeColorTabSel")
            On Error GoTo 0
        End If
        If iProp Is Nothing Then
            nCtl.IconColorMouseHoverSelectedTab = nCtl.ForeColor
        Else
            If iProp.Value = "F" Then nCtl.IconColorMouseHoverSelectedTab = nAmbientForeColor Else nCtl.IconColorMouseHoverSelectedTab = Val(iProp.Value)
        End If
    End If
    
    If PropertyExists("BackColorTabs") Then nCtl.BackColorTabs = GetPropertyValue("BackColorTabs") Else nCtl.BackColorTabs = nCtl.BackColor
    
    If PropertyExists("BackColorSelectedTab") Then
        nCtl.BackColorSelectedTab = GetPropertyValue("BackColorSelectedTab")
    ElseIf PropertyExists("BackColorTabSel") Then
        nCtl.BackColorSelectedTab = GetPropertyValue("BackColorTabSel")
    Else
        nCtl.BackColorSelectedTab = nCtl.BackColorTabs
    End If
    
    
    If PropertyExists("ForeColorSelectedTab") Then
        nCtl.ForeColorSelectedTab = GetPropertyValue("ForeColorSelectedTab")
    ElseIf PropertyExists("ForeColorTabSel") Then
        nCtl.ForeColorSelectedTab = GetPropertyValue("ForeColorTabSel")
    Else
        nCtl.ForeColorSelectedTab = nCtl.ForeColor
    End If
    
    If PropertyExists("ForeColorHighlighted") Then nCtl.ForeColorHighlighted = GetPropertyValue("ForeColorHighlighted") Else nCtl.ForeColorHighlighted = nCtl.ForeColor
    If PropertyExists("FlatTabBorderColorHighlight") Then nCtl.FlatTabBorderColorHighlight = GetPropertyValue("FlatTabBorderColorHighlight") Else nCtl.FlatTabBorderColorHighlight = nCtl.ForeColor
    
    If PropertyExists("FlatTabBorderColorSelectedTab") Then
        nCtl.FlatTabBorderColorSelectedTab = GetPropertyValue("FlatTabBorderColorSelectedTab")
    ElseIf PropertyExists("FlatTabBorderColorTabSel") Then
        nCtl.FlatTabBorderColorSelectedTab = GetPropertyValue("FlatTabBorderColorTabSel")
    Else
        nCtl.FlatTabBorderColorSelectedTab = nCtl.ForeColor
    End If
    
    If PropertyExists("IconColor") Then nCtl.IconColor = GetPropertyValue("IconColor") Else nCtl.IconColor = nCtl.ForeColor
    
    If PropertyExists("IconColorSelectedTab") Then
        nCtl.IconColorSelectedTab = GetPropertyValue("IconColorSelectedTab")
    ElseIf PropertyExists("IconColorTabSel") Then
        nCtl.IconColorSelectedTab = GetPropertyValue("IconColorTabSel")
    Else
        nCtl.IconColorSelectedTab = nCtl.ForeColor
    End If
    
    If PropertyExists("IconColorTabHighlighted") Then nCtl.IconColorTabHighlighted = GetPropertyValue("IconColorTabHighlighted") Else nCtl.IconColorTabHighlighted = nCtl.ForeColor
    
    If PropertyExists("FlatBarColorSelectedTab") Then
        nCtl.FlatBarColorSelectedTab = GetPropertyValue("FlatBarColorSelectedTab")
    ElseIf PropertyExists("FlatBarColorTabSel") Then
        nCtl.FlatBarColorSelectedTab = GetPropertyValue("FlatBarColorTabSel")
    End If
    
    If PropertyExists("FlatBarColorInactive") Then nCtl.FlatBarColorInactive = GetPropertyValue("FlatBarColorInactive")
    If PropertyExists("FlatTabsSeparationLineColor") Then nCtl.FlatTabsSeparationLineColor = GetPropertyValue("FlatTabsSeparationLineColor")
    If PropertyExists("FlatBodySeparationLineColor") Then nCtl.FlatBodySeparationLineColor = GetPropertyValue("FlatBodySeparationLineColor")
    If PropertyExists("FlatBorderColor") Then nCtl.FlatBorderColor = GetPropertyValue("FlatBorderColor")
    If PropertyExists("FlatBarColorHighlight") Then nCtl.FlatBarColorHighlight = GetPropertyValue("FlatBarColorHighlight")
    If PropertyExists("HighlightColor") Then nCtl.HighlightColor = GetPropertyValue("HighlightColor")
    
    If PropertyExists("HighlightColorSelectedTab") Then
        nCtl.HighlightColorSelectedTab = GetPropertyValue("HighlightColorSelectedTab")
    ElseIf PropertyExists("HighlightColorTabSel") Then
        nCtl.HighlightColorSelectedTab = GetPropertyValue("HighlightColorTabSel")
    End If
    
    If PropertyExists("Style") Then nCtl.Style = GetPropertyValue("Style")
    If PropertyExists("TabAppearance") Then nCtl.TabAppearance = GetPropertyValue("TabAppearance")
    
    For Each iProp In nThemeData
        Select Case iProp.Name
            Case "TabWidthStyle"
                If nCtl.TDIMode = ntTDIModeNone Then
                    nCtl.TabWidthStyle = Val(iProp.Value)
                End If
            Case "ShowRowsInPerspective"
                nCtl.ShowRowsInPerspective = Val(iProp.Value)
           Case "BackStyle"
                nCtl.BackStyle = Val(iProp.Value)
            Case "HighlightMode"
                nCtl.HighlightMode = Val(iProp.Value)
            Case "HighlightModeSelectedTab", "HighlightModeTabSel"
                nCtl.HighlightModeSelectedTab = Val(iProp.Value)
            Case "AutoRelocateControls"
                nCtl.AutoRelocateControls = Val(iProp.Value)
            Case "FlatBorderMode"
                nCtl.FlatBorderMode = Val(iProp.Value)
            Case "FlatBarPosition"
                nCtl.FlatBarPosition = Val(iProp.Value)
            Case "FlatBarHeight"
                nCtl.FlatBarHeight = Val(iProp.Value)
            Case "FlatBarGripHeight"
                nCtl.FlatBarGripHeight = Val(iProp.Value)
            Case "FlatBodySeparationLineHeight"
                nCtl.FlatBodySeparationLineHeight = Val(iProp.Value)
            Case "TabSeparation"
                nCtl.TabSeparation = Val(iProp.Value)
            Case "HighlightTabExtraHeight"
                nCtl.HighlightTabExtraHeight = Val(iProp.Value)
            Case "TabMaxWidth"
                nCtl.TabMaxWidth = Val(iProp.Value)
            Case "TabMinWidth"
                nCtl.TabMinWidth = Val(iProp.Value)
            Case "FlatRoundnessTop"
                nCtl.FlatRoundnessTop = Val(iProp.Value)
            Case "FlatRoundnessBottom"
                nCtl.FlatRoundnessBottom = Val(iProp.Value)
            Case "FlatRoundnessTabs"
                nCtl.FlatRoundnessTabs = Val(iProp.Value)
            Case "ShowDisabledState"
                nCtl.ShowDisabledState = CBool(Val(iProp.Value))
            Case "ShowFocusRect"
                nCtl.ShowFocusRect = CBool(Val(iProp.Value))
            Case "HighlightEffect"
                nCtl.HighlightEffect = CBool(Val(iProp.Value))
            Case "WordWrap"
                nCtl.WordWrap = CBool(Val(iProp.Value))
            Case "SoftEdges"
                nCtl.SoftEdges = CBool(Val(iProp.Value))
            Case "TabMousePointerHand"
                nCtl.TabMousePointerHand = CBool(Val(iProp.Value))
            Case "MaskColor"
                nCtl.MaskColor = Val(iProp.Value)
        End Select
    Next
    nCtl.Redraw = iRedraw
End Sub

Public Function GetThemeStringFromControl(nCtl As NewTab, nAmbientBackColor As Long, nAmbientForeColor As Long, Optional nHash As String) As String
    Dim iPropsStr() As String
    Dim ub As Long
    Dim iTheme As NewTabTheme
    Dim iTakeAmbientColors As Boolean
    Dim c As Long
    
    c = -1
    ub = 100
    ReDim iPropsStr(ub)
    
    If nCtl.Style <> cPropDef_Style Then AddPropStrToArray iPropsStr, c, "Style", nCtl.Style
    If nCtl.Style <> ntStyleWindows Then
        If nCtl.TabAppearance <> cPropDef_TabAppearance Then AddPropStrToArray iPropsStr, c, "TabAppearance", nCtl.TabAppearance
        If nCtl.HighlightEffect <> cPropDef_HighlightEffect Then AddPropStrToArray iPropsStr, c, "HighlightEffect", Val(Str$(CLng(nCtl.HighlightEffect)))
        If nCtl.ShowDisabledState <> cPropDef_ShowDisabledState Then AddPropStrToArray iPropsStr, c, "ShowDisabledState", Val(Str$(CLng(nCtl.ShowDisabledState)))
    End If
    If nCtl.HighlightMode <> cPropDef_HighlightMode Then AddPropStrToArray iPropsStr, c, "HighlightMode", nCtl.HighlightMode
    If nCtl.HighlightModeSelectedTab <> cPropDef_HighlightModeSelectedTab Then AddPropStrToArray iPropsStr, c, "HighlightModeSelectedTab", nCtl.HighlightModeSelectedTab
    If nCtl.TabWidthStyle <> cPropDef_TabWidthStyle Then AddPropStrToArray iPropsStr, c, "TabWidthStyle", nCtl.TabWidthStyle
    If nCtl.ShowRowsInPerspective <> cPropDef_ShowRowsInPerspective Then AddPropStrToArray iPropsStr, c, "ShowRowsInPerspective", nCtl.ShowRowsInPerspective
    If nCtl.BackStyle <> cPropDef_BackStyle Then AddPropStrToArray iPropsStr, c, "BackStyle", nCtl.BackStyle
    If nCtl.AutoRelocateControls <> cPropDef_AutoRelocateControls Then AddPropStrToArray iPropsStr, c, "AutoRelocateControls", nCtl.AutoRelocateControls
'    If nCtl.TabTransition <> cPropDef_TabTransition Then AddPropStrToArray iPropsStr, c, "TabTransition", nCtl.TabTransition
'    Select Case nCtl.TabWidthStyle
'        Case ntTWFixed, ntTWTabCaptionWidth
'            If nCtl.TabsPerRow <> cPropDef_TabsPerRow Then AddPropStrToArray iPropsStr, c, "TabsPerRow", Val(Str$(nCtl.TabsPerRow))
'        Case ntTWAuto
'            Select Case nCtl.Style
'                Case ssStyleTabbedDialog, ssStylePropertyPage
'                    If nCtl.TabsPerRow <> cPropDef_TabsPerRow Then AddPropStrToArray iPropsStr, c, "TabsPerRow", Val(Str$(nCtl.TabsPerRow))
'            End Select
'    End Select
    If nCtl.TabSeparation <> cPropDef_TabSeparation Then AddPropStrToArray iPropsStr, c, "TabSeparation", Val(Str$(nCtl.TabSeparation))
    If nCtl.HighlightTabExtraHeight <> cPropDef_HighlightTabExtraHeight Then AddPropStrToArray iPropsStr, c, "HighlightTabExtraHeight", Val(Str$(nCtl.HighlightTabExtraHeight))
    If nCtl.TabMaxWidth <> cPropDef_TabMaxWidth Then AddPropStrToArray iPropsStr, c, "TabMaxWidth", Val(Str$(nCtl.TabMaxWidth))
    If nCtl.TabMinWidth <> cPropDef_TabMinWidth Then AddPropStrToArray iPropsStr, c, "TabMinWidth", Val(Str$(nCtl.TabMinWidth))
    If nCtl.ShowFocusRect <> cPropDef_ShowFocusRect Then AddPropStrToArray iPropsStr, c, "ShowFocusRect", Val(Str$(CLng(nCtl.ShowFocusRect)))
'    If nCtl.ChangeControlsBackColor <> cPropDef_ChangeControlsBackColor Then AddPropStrToArray iPropsStr, c, "ChangeControlsBackColor", Val(Str$(CLng(nCtl.ChangeControlsBackColor)))
'    If nCtl.ChangeControlsForeColor <> cPropDef_ChangeControlsForeColor Then AddPropStrToArray iPropsStr, c, "ChangeControlsForeColor", Val(Str$(CLng(nCtl.ChangeControlsForeColor)))
    If nCtl.WordWrap <> cPropDef_WordWrap Then AddPropStrToArray iPropsStr, c, "WordWrap", Val(Str$(CLng(nCtl.WordWrap)))
    If nCtl.TabMousePointerHand <> cPropDef_TabMousePointerHand Then AddPropStrToArray iPropsStr, c, "TabMousePointerHand", Val(Str$(CLng(nCtl.TabMousePointerHand)))
    If nCtl.Style <> ntStyleWindows Then
        Select Case nCtl.TabAppearance
            Case ntTATabbedDialog, ntTATabbedDialogRounded, ntTAPropertyPage, ntTAPropertyPageRounded
                If nCtl.SoftEdges <> cPropDef_SoftEdges Then AddPropStrToArray iPropsStr, c, "SoftEdges", Val(Str$(CLng(nCtl.SoftEdges)))
            Case ntTAAuto
                Select Case nCtl.Style
                    Case ssStyleTabbedDialog, ssStylePropertyPage, ntStyleTabStrip
                        If nCtl.SoftEdges <> cPropDef_SoftEdges Then AddPropStrToArray iPropsStr, c, "SoftEdges", Val(Str$(CLng(nCtl.SoftEdges)))
                End Select
        End Select
    End If
    If nCtl.Style = ntStyleFlat Then
        If nCtl.FlatBorderMode <> cPropDef_FlatBorderMode Then AddPropStrToArray iPropsStr, c, "FlatBorderMode", nCtl.FlatBorderMode
        If nCtl.FlatBarPosition <> cPropDef_FlatBarPosition Then AddPropStrToArray iPropsStr, c, "FlatBarPosition", nCtl.FlatBarPosition
        If nCtl.FlatBarHeight <> cPropDef_FlatBarHeight Then AddPropStrToArray iPropsStr, c, "FlatBarHeight", nCtl.FlatBarHeight
        If nCtl.FlatBarGripHeight <> cPropDef_FlatBarGripHeight Then AddPropStrToArray iPropsStr, c, "FlatBarGripHeight", nCtl.FlatBarGripHeight
        If nCtl.FlatBodySeparationLineHeight <> cPropDef_FlatBodySeparationLineHeight Then AddPropStrToArray iPropsStr, c, "FlatBodySeparationLineHeight", nCtl.FlatBodySeparationLineHeight
        If nCtl.FlatRoundnessTop <> cPropDef_FlatRoundnessTop Then AddPropStrToArray iPropsStr, c, "FlatRoundnessTop", Val(Str$(nCtl.FlatRoundnessTop))
        If nCtl.FlatRoundnessBottom <> cPropDef_FlatRoundnessBottom Then AddPropStrToArray iPropsStr, c, "FlatRoundnessBottom", Val(Str$(nCtl.FlatRoundnessBottom))
        If nCtl.FlatRoundnessTabs <> cPropDef_FlatRoundnessTabs Then AddPropStrToArray iPropsStr, c, "FlatRoundnessTabs", Val(Str$(nCtl.FlatRoundnessTabs))
    End If
    
    ' colors
    If nCtl.BackColor <> nAmbientBackColor Then
        AddPropStrToArray iPropsStr, c, "BackColor", nCtl.BackColor
    End If
    
    iTakeAmbientColors = ControlTakesAmbientColors(nCtl, nAmbientBackColor, nAmbientForeColor)
    
    If nCtl.ForeColor <> nAmbientForeColor Then
        AddPropStrToArray iPropsStr, c, "ForeColor", nCtl.ForeColor
    End If
    If nCtl.MaskColor <> cPropDef_MaskColor Then AddPropStrToArray iPropsStr, c, "MaskColor", nCtl.MaskColor
    If nCtl.BackColorTabs <> nCtl.BackColor Then
        If Not iTakeAmbientColors Then
            AddPropStrToArray iPropsStr, c, "BackColorTabs", nCtl.BackColorTabs
        End If
    End If
    If Not nCtl.BackColorSelectedTab_IsAutomatic Then
        If iTakeAmbientColors Then
            AddPropStrToArray iPropsStr, c, "BackColorSelectedTab", "B"
        Else
            AddPropStrToArray iPropsStr, c, "BackColorSelectedTab", nCtl.BackColorSelectedTab
        End If
    End If
    If nCtl.ForeColorSelectedTab <> nCtl.ForeColor Then
        AddPropStrToArray iPropsStr, c, "ForeColorSelectedTab", nCtl.ForeColorSelectedTab
    End If
    If nCtl.ForeColorHighlighted <> nCtl.ForeColor Then
        If iTakeAmbientColors And ((nCtl.ForeColorHighlighted = nAmbientForeColor) Or (nCtl.ForeColorHighlighted = vbButtonText)) Then
            AddPropStrToArray iPropsStr, c, "ForeColorHighlighted", "F"
        Else
            AddPropStrToArray iPropsStr, c, "ForeColorHighlighted", nCtl.ForeColorHighlighted
        End If
    End If
    If nCtl.FlatTabBorderColorHighlight <> nCtl.ForeColor Then
        If iTakeAmbientColors And ((nCtl.FlatTabBorderColorHighlight = nAmbientForeColor) Or (nCtl.FlatTabBorderColorHighlight = vbButtonText)) Then
            AddPropStrToArray iPropsStr, c, "FlatTabBorderColorHighlight", "F"
        Else
            AddPropStrToArray iPropsStr, c, "FlatTabBorderColorHighlight", nCtl.FlatTabBorderColorHighlight
        End If
    End If
    If nCtl.FlatTabBorderColorSelectedTab <> nCtl.ForeColor Then
        If iTakeAmbientColors And ((nCtl.FlatTabBorderColorSelectedTab = nAmbientForeColor) Or (nCtl.FlatTabBorderColorSelectedTab = vbButtonText)) Then
            AddPropStrToArray iPropsStr, c, "FlatTabBorderColorSelectedTab", "F"
        Else
            AddPropStrToArray iPropsStr, c, "FlatTabBorderColorSelectedTab", nCtl.FlatTabBorderColorSelectedTab
        End If
    End If
    If nCtl.IconColor <> nCtl.ForeColor Then
        If Not iTakeAmbientColors Then
            AddPropStrToArray iPropsStr, c, "IconColor", nCtl.IconColor
        End If
    End If
    If nCtl.IconColorSelectedTab <> nCtl.ForeColor Then
        If Not iTakeAmbientColors Then
            AddPropStrToArray iPropsStr, c, "IconColorSelectedTab", nCtl.IconColorSelectedTab
        End If
    End If
    If nCtl.IconColorTabHighlighted <> nCtl.ForeColor Then
        If iTakeAmbientColors And ((nCtl.IconColorTabHighlighted = nAmbientForeColor) Or (nCtl.IconColorTabHighlighted = vbButtonText)) Then
            AddPropStrToArray iPropsStr, c, "IconColorTabHighlighted", "F"
        Else
            AddPropStrToArray iPropsStr, c, "IconColorTabHighlighted", nCtl.IconColorTabHighlighted
        End If
    End If
    If Not nCtl.HighlightColor_IsAutomatic Then
        If iTakeAmbientColors And ((nCtl.HighlightColor = nAmbientForeColor) Or (nCtl.HighlightColor = vbButtonText)) Then
            AddPropStrToArray iPropsStr, c, "HighlightColor", "F"
        Else
            AddPropStrToArray iPropsStr, c, "HighlightColor", nCtl.HighlightColor
        End If
    End If
    If Not nCtl.HighlightColorSelectedTab_IsAutomatic Then
        If iTakeAmbientColors And ((nCtl.HighlightColorSelectedTab = nAmbientForeColor) Or (nCtl.HighlightColorSelectedTab = vbButtonText)) Then
            AddPropStrToArray iPropsStr, c, "HighlightColorSelectedTab", "F"
        Else
            AddPropStrToArray iPropsStr, c, "HighlightColorSelectedTab", nCtl.HighlightColorSelectedTab
        End If
    End If
    
    If nCtl.Style = ntStyleFlat Then
        If Not nCtl.FlatBarColorInactive_IsAutomatic Then
            If iTakeAmbientColors And ((nCtl.FlatBarColorInactive = nAmbientBackColor) Or (nCtl.FlatBarColorInactive = vbButtonFace)) Then
                AddPropStrToArray iPropsStr, c, "FlatBarColorInactive", "B"
            ElseIf iTakeAmbientColors And ((nCtl.FlatBarColorInactive = nAmbientForeColor) Or (nCtl.FlatBarColorInactive = vbButtonText)) Then
                AddPropStrToArray iPropsStr, c, "FlatBarColorInactive", "F"
            Else
                AddPropStrToArray iPropsStr, c, "FlatBarColorInactive", nCtl.FlatBarColorInactive
            End If
        End If
        If Not nCtl.FlatTabsSeparationLineColor_IsAutomatic Then
            If iTakeAmbientColors And ((nCtl.FlatTabsSeparationLineColor = nAmbientBackColor) Or (nCtl.FlatTabsSeparationLineColor = vbButtonFace)) Then
                AddPropStrToArray iPropsStr, c, "FlatTabsSeparationLineColor", "B"
            ElseIf iTakeAmbientColors And ((nCtl.FlatTabsSeparationLineColor = nAmbientForeColor) Or (nCtl.FlatTabsSeparationLineColor = vbButtonText)) Then
                AddPropStrToArray iPropsStr, c, "FlatTabsSeparationLineColor", "F"
            Else
                AddPropStrToArray iPropsStr, c, "FlatTabsSeparationLineColor", nCtl.FlatTabsSeparationLineColor
            End If
        End If
        If Not nCtl.FlatBodySeparationLineColor_IsAutomatic Then
            If iTakeAmbientColors And ((nCtl.FlatBodySeparationLineColor = nAmbientBackColor) Or (nCtl.FlatBodySeparationLineColor = vbButtonFace)) Then
                AddPropStrToArray iPropsStr, c, "FlatBodySeparationLineColor", "B"
            ElseIf iTakeAmbientColors And ((nCtl.FlatBodySeparationLineColor = nAmbientForeColor) Or (nCtl.FlatBodySeparationLineColor = vbButtonText)) Then
                AddPropStrToArray iPropsStr, c, "FlatBodySeparationLineColor", "F"
            Else
                AddPropStrToArray iPropsStr, c, "FlatBodySeparationLineColor", nCtl.FlatBodySeparationLineColor
            End If
        End If
        If Not nCtl.FlatBorderColor_IsAutomatic Then
            If iTakeAmbientColors And ((nCtl.FlatBorderColor = nAmbientBackColor) Or (nCtl.FlatBorderColor = vbButtonFace)) Then
                AddPropStrToArray iPropsStr, c, "FlatBorderColor", "B"
            ElseIf iTakeAmbientColors And ((nCtl.FlatBorderColor = nAmbientForeColor) Or (nCtl.FlatBorderColor = vbButtonText)) Then
                AddPropStrToArray iPropsStr, c, "FlatBorderColor", "F"
            Else
                AddPropStrToArray iPropsStr, c, "FlatBorderColor", nCtl.FlatBorderColor
            End If
        End If
        If Not nCtl.FlatBarColorHighlight_IsAutomatic Then
            If iTakeAmbientColors And ((nCtl.FlatBarColorHighlight = nAmbientBackColor) Or (nCtl.FlatBarColorHighlight = vbButtonFace)) Then
                AddPropStrToArray iPropsStr, c, "FlatBarColorHighlight", "B"
            ElseIf iTakeAmbientColors And ((nCtl.FlatBarColorHighlight = nAmbientForeColor) Or (nCtl.FlatBarColorHighlight = vbButtonText)) Then
                AddPropStrToArray iPropsStr, c, "FlatBarColorHighlight", "F"
            Else
                AddPropStrToArray iPropsStr, c, "FlatBarColorHighlight", nCtl.FlatBarColorHighlight
            End If
        End If
        If iTakeAmbientColors And ((nCtl.FlatBarColorSelectedTab = nAmbientBackColor) Or (nCtl.FlatBarColorSelectedTab = vbButtonFace)) Then
            AddPropStrToArray iPropsStr, c, "FlatBarColorSelectedTab", "B"
        ElseIf iTakeAmbientColors And ((nCtl.FlatBarColorSelectedTab = nAmbientForeColor) Or (nCtl.FlatBarColorSelectedTab = vbButtonText)) Then
            AddPropStrToArray iPropsStr, c, "FlatBarColorSelectedTab", "F"
        Else
            AddPropStrToArray iPropsStr, c, "FlatBarColorSelectedTab", nCtl.FlatBarColorSelectedTab
        End If
    End If
    
    If c > -1 Then
        ReDim Preserve iPropsStr(c)
    Else
        ReDim iPropsStr(-1 To -1)
    End If
    QuickSort iPropsStr
    Set iTheme = New NewTabTheme
    iTheme.ThemeString = Mid$(Join(iPropsStr), 2)
    GetThemeStringFromControl = iTheme.ThemeString
    nHash = iTheme.Hash
End Function

Private Sub AddPropStrToArray(ByRef nPropStrArray() As String, ByRef nPos As Long, ByRef nPropName As String, ByRef nPropValue As String)
    nPos = nPos + 1
    nPropStrArray(nPos) = "|" & nPropName & "=" & nPropValue
End Sub

Private Function ControlTakesAmbientColors(ByRef nCtl As NewTab, ByVal nAmbientBackColor As Long, ByVal nAmbientForeColor As Long) As Boolean
    ControlTakesAmbientColors = False
    If (nCtl.ForeColor = nAmbientForeColor) Or (nCtl.ForeColor = vbButtonText) Then
        If (nCtl.BackColorTabs = nAmbientBackColor) Or (nCtl.BackColorTabs = vbButtonFace) Then
            If (nCtl.BackColorSelectedTab = nAmbientBackColor) Or (nCtl.BackColorSelectedTab = vbButtonFace) Then
                If (nCtl.ForeColorSelectedTab = nAmbientForeColor) Or (nCtl.ForeColorSelectedTab = vbButtonText) Then
                    If (nCtl.IconColor = nAmbientForeColor) Or (nCtl.IconColor = vbButtonText) Then
                        If (nCtl.IconColorSelectedTab = nAmbientForeColor) Or (nCtl.IconColorSelectedTab = vbButtonText) Then
                            ControlTakesAmbientColors = True
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function

Private Function FontsAreEqual(nFont1 As StdFont, nFont2 As StdFont) As Boolean
    If nFont1 Is Nothing Or nFont2 Is Nothing Then Exit Function
    
    If (nFont1 Is Nothing) And (nFont2 Is Nothing) Then
        FontsAreEqual = True
        Exit Function
    End If
    If (nFont1 Is Nothing) Then Exit Function
    If (nFont2 Is Nothing) Then Exit Function
    
    If nFont1.Name = nFont2.Name Then
        If nFont1.Size = nFont2.Size Then
            If nFont1.Bold = nFont2.Bold Then
                If nFont1.Italic = nFont2.Italic Then
                    If nFont1.Strikethrough = nFont2.Strikethrough Then
                        If nFont1.Underline = nFont2.Underline Then
                            If nFont1.Weight = nFont2.Weight Then
                                If nFont1.Charset = nFont2.Charset Then
                                    FontsAreEqual = True
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function

Private Sub LoadDefaultThemes()
    Set mDefaultThemes = New NewTabThemes
    mDefaultThemes.DoNotCopyDefaultThemes = True
    AddDefaultTheme "Default (Style: Windows)", cThemeString_Default
    AddDefaultTheme "SSTab", cThemeString_SSTab
    AddDefaultTheme "SSTab with Windows styles", cThemeString_SSTabWindows
    AddDefaultTheme "SSTab Style property page", cThemeString_SSTabPropertyPage
    AddDefaultTheme "SSTab Style property page with Windows styles", cThemeString_SSTabPropertyPageWindows
    AddDefaultTheme "TabStrip", cThemeString_TabStrip
    AddDefaultTheme "TabStrip with Windows styles", cThemeString_TabStripWindows
    AddDefaultTheme "Silver", cThemeString_FlatSilver
    AddDefaultTheme "Bronze", cThemeString_FlatBronze
    AddDefaultTheme "Apple Green", cThemeString_FlatAppleGreen
    AddDefaultTheme "Golden", cThemeString_FlatGolden
    AddDefaultTheme "Sea Blue", cThemeString_FlatSeaBlue
    AddDefaultTheme "Emerald", cThemeString_FlatEmerald
    AddDefaultTheme "Red Wine", cThemeString_FlatRedWine
    AddDefaultTheme "Deep Waters", cThemeString_FlatDeepWaters
    AddDefaultTheme "Open Air", cThemeString_FlatOpenAir
    AddDefaultTheme "Ghost Tab", cThemeString_GhostTab
    AddDefaultTheme "Buttons", cThemeString_Buttons
    AddDefaultTheme "Buttons 2", cThemeString_Buttons2
    AddDefaultTheme "Buttons 3", cThemeString_Buttons3
    AddDefaultTheme "Buttons 4", cThemeString_Buttons4
    AddDefaultTheme "Buttons 5", cThemeString_Buttons5
    AddDefaultTheme "Buttons 6", cThemeString_Buttons6
    AddDefaultTheme "Buttons 7", cThemeString_Buttons7
    AddDefaultTheme "Buttons 8", cThemeString_Buttons8
    AddDefaultTheme "Buttons 9", cThemeString_Buttons9
    AddDefaultTheme "Buttons 10", cThemeString_Buttons10
    AddDefaultTheme "Buttons 11", cThemeString_Buttons11
    AddDefaultTheme "Web Links", cThemeString_WebLinks
    AddDefaultTheme "Web Links 2", cThemeString_WebLinks2
    AddDefaultTheme "Another Button Light", cThemeString_AnotherButtonLight
    AddDefaultTheme "Another Button Dark", cThemeString_AnotherButtonDark
End Sub

Private Sub AddDefaultTheme(nName As String, nThemeString As String)
    Dim iTheme As NewTabTheme
    
    Set iTheme = New NewTabTheme
    iTheme.Name = nName
    iTheme.Custom = False
    iTheme.ThemeString = nThemeString
    mDefaultThemes.Add iTheme
End Sub

' Omit plngLeft & plngRight; they are used internally during recursion
Private Sub QuickSort(ByRef pvarArray As Variant, Optional ByVal plngLeft As Long, Optional ByVal plngRight As Long)
    Dim lngFirst As Long
    Dim lngLast As Long
    Dim varMid As Variant
    Dim varSwap As Variant
    
    If plngRight = 0 Then
        plngLeft = LBound(pvarArray)
        plngRight = UBound(pvarArray)
    End If
    lngFirst = plngLeft
    lngLast = plngRight
    varMid = pvarArray((plngLeft + plngRight) \ 2)
    Do
        Do While pvarArray(lngFirst) < varMid And lngFirst < plngRight
            lngFirst = lngFirst + 1
        Loop
        Do While varMid < pvarArray(lngLast) And lngLast > plngLeft
            lngLast = lngLast - 1
        Loop
        If lngFirst <= lngLast Then
            varSwap = pvarArray(lngFirst)
            pvarArray(lngFirst) = pvarArray(lngLast)
            pvarArray(lngLast) = varSwap
            lngFirst = lngFirst + 1
            lngLast = lngLast - 1
        End If
    Loop Until lngFirst > lngLast
    If plngLeft < lngLast Then QuickSort pvarArray, plngLeft, lngLast
    If lngFirst < plngRight Then QuickSort pvarArray, lngFirst, plngRight
End Sub

Public Function IsValidOLE_COLOR(ByVal nColor As Long) As Boolean
    Const S_OK As Long = 0
    IsValidOLE_COLOR = (TranslateColor(nColor, 0, nColor) = S_OK)
End Function


