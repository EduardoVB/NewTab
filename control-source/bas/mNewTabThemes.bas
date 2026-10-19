Attribute VB_Name = "mNewTabThemes"
Option Explicit

Private Declare Function PathGetCharType Lib "shlwapi.dll" Alias "PathGetCharTypeW" (ByVal ch As Long) As Long
Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As LongPtr, Col As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExW" (ByVal hWndParent As LongPtr, ByVal hWndChildAfter As LongPtr, ByVal lpszClass As LongPtr, ByVal lpszWindow As LongPtr) As LongPtr
Private Declare Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hWnd As LongPtr, lpdwProcessId As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32.dll" () As Long
Private Declare Function GetProp Lib "user32.dll" Alias "GetPropW" (ByVal hWnd As LongPtr, ByVal lpString As LongPtr) As LongPtr

Public gPagTabsApplyTime As Double
Public gPagTabsLastSelectedTab As Double

Private mDefaultThemes As NewTabThemes
Private mThemesFile As NewTabThemes
Private mThemesRegistry As NewTabThemes
Private mThemesLoaded As Boolean
Private mThemesFilePath As String
Private mControlTypeName As String

Private mAuxAmbientBackColor As Long
Private mAuxAmbientForeColor As Long
Private mAuxThemeData As Collection

Public Const cThemeIsCustomSettings As String = "Current custom settings, not saved as a Theme"

' Default Themes strings
Private Const cThemeString_Default                       As String = ""
' A theme names only what the defaults and the Auto values do not already give for its Style:
' anything a theme names is frozen and stops following Style. That is why the SSTab and TabStrip
' themes name no highlight entry and no ShowFocusRect, and the Windows ones name the mode they need.
' TabPaddingWidth: 7 matches SSTab property page and the TabStrip with visual styles, 6 the TabStrip
' without them, and only where the width style measures captions. TabPaddingHeight = 3 gives the
' TabStrip its shorter row; with visual styles the default is the closest
Private Const cThemeString_SSTab                         As String = "SoftEdges=0|Style=0"
Private Const cThemeString_SSTabWindows                  As String = "HighlightModeSelectedTab=16|ShowRowsInPerspective=1|TabWidthStyle=2"
Private Const cThemeString_SSTabPropertyPage             As String = "SoftEdges=0|Style=1|TabPaddingWidth=7"
Private Const cThemeString_SSTabPropertyPageWindows      As String = "HighlightModeSelectedTab=2304|ShowRowsInPerspective=1|TabPaddingWidth=7|TabWidthStyle=1"
Private Const cThemeString_TabStrip                      As String = "SoftEdges=0|Style=2|TabPaddingHeight=3|TabPaddingWidth=6"
Private Const cThemeString_TabStripWindows               As String = "TabPaddingWidth=7|TabWidthStyle=0"
Private Const cThemeString_FlatSilver                    As String = "BackColorTabs=15658734|ForeColor=0|Style=3"
Private Const cThemeString_FlatBronze                    As String = "BackColorSelectedTab=16383485|BackColorTabs=14611960|FlatBarColorHighlighted=3431538|FlatBarColorInactive=13559786|FlatBarColorSelectedTab=1148870|FlatBorderColor=1148870|FlatBorderMode=1|ForeColor=0|ForeColorHighlighted=16777215|HighlightColor=3431538|HighlightEffect=0|HighlightEffectOut=0|HighlightMode=68|HighlightModeSelectedTab=90|Style=3"
Private Const cThemeString_FlatAppleGreen                As String = "BackColorSelectedTab=16514553|BackColorTabs=15136990|FlatBarColorHighlighted=3633716|FlatBarColorInactive=14150350|FlatBarColorSelectedTab=1820177|FlatBorderColor=1820177|FlatBorderMode=1|ForeColor=0|ForeColorHighlighted=16777215|HighlightColor=3633716|HighlightEffect=0|HighlightEffectOut=0|HighlightMode=68|HighlightModeSelectedTab=90|Style=3"
Private Const cThemeString_FlatGolden                    As String = "BackColorSelectedTab=16777215|BackColorTabs=15202556|FlatBarColorHighlighted=3530228|FlatBarColorInactive=13559786|FlatBarColorSelectedTab=768981|ForeColor=0|HighlightColor=3530228|HighlightMode=76|HighlightModeSelectedTab=90|IconColorTabHighlighted=12664841|Style=3"
Private Const cThemeString_FlatSeaBlue                   As String = "BackColorSelectedTab=16250871|FlatBarHeight=0|FlatBorderColor=10184001|FlatBorderMode=1|FlatRoundnessTabs=8|FlatTabsSeparationLineColor=-2147483633|ForeColorHighlighted=16777215|ForeColorSelectedTab=10184001|HighlightColor=16477710|HighlightMode=4|HighlightModeSelectedTab=10|Style=3|TabMousePointerHand=-1"
Private Const cThemeString_FlatSeaGreen                  As String = "BackColorSelectedTab=16250871|FlatBarHeight=0|FlatBorderColor=4681259|FlatBorderMode=1|FlatRoundnessTabs=8|FlatTabsSeparationLineColor=-2147483633|ForeColorHighlighted=16777215|ForeColorSelectedTab=4681259|HighlightColor=4356630|HighlightMode=4|HighlightModeSelectedTab=10|Style=3|TabMousePointerHand=-1"
Private Const cThemeString_FlatCherry                    As String = "BackColorSelectedTab=16250871|FlatBarHeight=0|FlatBorderColor=4273563|FlatBorderMode=1|FlatRoundnessTabs=8|FlatTabsSeparationLineColor=-2147483633|ForeColorHighlighted=16777215|ForeColorSelectedTab=4273563|HighlightColor=4075721|HighlightMode=4|HighlightModeSelectedTab=10|Style=3|TabMousePointerHand=-1"
Private Const cThemeString_FlatAmber                     As String = "BackColorSelectedTab=16250871|FlatBarHeight=0|FlatBorderColor=1202826|FlatBorderMode=1|FlatRoundnessTabs=8|FlatTabsSeparationLineColor=-2147483633|ForeColorHighlighted=16777215|ForeColorSelectedTab=1202826|HighlightColor=24480|HighlightMode=4|HighlightModeSelectedTab=10|Style=3|TabMousePointerHand=-1"
Private Const cThemeString_FlatViolet                    As String = "BackColorSelectedTab=16250871|FlatBarHeight=0|FlatBorderColor=10177118|FlatBorderMode=1|FlatRoundnessTabs=8|FlatTabsSeparationLineColor=-2147483633|ForeColorHighlighted=16777215|ForeColorSelectedTab=10177118|HighlightColor=14175087|HighlightMode=4|HighlightModeSelectedTab=10|Style=3|TabMousePointerHand=-1"
Private Const cThemeString_FlatTeal                      As String = "BackColorSelectedTab=16250871|FlatBarHeight=0|FlatBorderColor=7630623|FlatBorderMode=1|FlatRoundnessTabs=8|FlatTabsSeparationLineColor=-2147483633|ForeColorHighlighted=16777215|ForeColorSelectedTab=7630623|HighlightColor=9076495|HighlightMode=4|HighlightModeSelectedTab=10|Style=3|TabMousePointerHand=-1"
Private Const cThemeString_FlatGraphite                  As String = "BackColorSelectedTab=16250871|FlatBarHeight=0|FlatBorderColor=4603964|FlatBorderMode=1|FlatRoundnessTabs=8|FlatTabsSeparationLineColor=-2147483633|ForeColorHighlighted=16777215|ForeColorSelectedTab=4603964|HighlightColor=6840410|HighlightMode=4|HighlightModeSelectedTab=10|Style=3|TabMousePointerHand=-1"
Private Const cThemeString_FlatClay                      As String = "BackColorTabs=14083057|FlatBarColorSelectedTab=4019656|ForeColor=1451067|Style=3"
Private Const cThemeString_FlatLavender                  As String = "BackColorTabs=16114921|FlatBarColorSelectedTab=12736382|ForeColor=4202794|Style=3"
Private Const cThemeString_FlatEmerald                   As String = "BackColorSelectedTab=2183936|BackColorTabs=1588736|FlatBarColorHighlighted=9615225|FlatBarColorInactive=4422175|FlatBarColorSelectedTab=16777215|FlatBodySeparationLineColor=1983492|FlatBorderColor=3960091|FlatTabsSeparationLineColor=1983492|ForeColor=16777215|HighlightColor=5281568|HighlightColorSelectedTab=5942308|HighlightMode=64|HighlightModeSelectedTab=90|Style=3|TabMousePointerHand=-1"
Private Const cThemeString_FlatRedWine                   As String = "BackColorSelectedTab=2424902|BackColorTabs=1835059|FlatBarColorHighlighted=8542630|FlatBarColorInactive=4332134|FlatBarColorSelectedTab=16777215|FlatBodySeparationLineColor=2032442|FlatBorderColor=3937884|FlatTabsSeparationLineColor=2032442|ForeColor=16777215|HighlightColor=5184382|HighlightColorSelectedTab=5906065|HighlightMode=64|HighlightModeSelectedTab=90|Style=3|TabMousePointerHand=-1"
Private Const cThemeString_FlatDeepWaters                As String = "BackColorSelectedTab=6699520|BackColorTabs=5057536|FlatBarColorHighlighted=16777215|FlatBarColorInactive=9856549|FlatBarColorSelectedTab=16777215|FlatBodySeparationLineColor=5057536|FlatBorderColor=8870689|ForeColor=16777215|HighlightColor=7554571|HighlightColorSelectedTab=13863980|HighlightMode=76|HighlightModeSelectedTab=90|Style=3|TabMousePointerHand=-1"
Private Const cThemeString_FlatCandlelight               As String = "BackColorTabs=4210752|FlatBarColorSelectedTab=8454143|FlatRoundnessBottom=0|FlatRoundnessTop=0|ForeColor=12648447|HighlightColor=4871256|HighlightColorSelectedTab=7697781|HighlightMode=4|Style=3"
Private Const cThemeString_FlatMoonlight                 As String = "BackColorTabs=4210752|FlatBarColorSelectedTab=8454143|FlatRoundnessBottom=0|FlatRoundnessTop=0|ForeColor=14942207|HighlightColor=5855577|HighlightColorSelectedTab=7697781|HighlightMode=4|Style=3"
Private Const cThemeString_FlatMidnight                  As String = "BackColorTabs=3351579|FlatBarColorSelectedTab=16762972|FlatRoundnessBottom=0|FlatRoundnessTop=0|ForeColor=16446187|HighlightColorSelectedTab=8413002|HighlightMode=4|Style=3"
Private Const cThemeString_FlatEspresso                  As String = "BackColorTabs=2239034|FlatBarColorSelectedTab=4956400|FlatRoundnessBottom=0|FlatRoundnessTop=0|ForeColor=14545146|HighlightColorSelectedTab=4741754|HighlightMode=4|Style=3"
Private Const cThemeString_FlatTwilight                  As String = "BackColorTabs=3679278|FlatBarColorSelectedTab=15372999|FlatRoundnessBottom=0|FlatRoundnessTop=0|ForeColor=16445428|HighlightColorSelectedTab=8015971|HighlightMode=4|Style=3"
Private Const cThemeString_FlatPitchBlack                As String = "BackColorSelectedTab=0|BackColorTabs=0|FlatBarColorHighlighted=16777215|FlatBarColorInactive=6316128|FlatBarColorSelectedTab=54527|FlatBodySeparationLineColor=16777215|FlatBorderColor=16777215|FlatTabsSeparationLineColor=8421504|ForeColor=16777215|HighlightColor=4210752|HighlightMode=68|HighlightModeSelectedTab=80|Style=3"
Private Const cThemeString_FlatInk                       As String = "BackColorSelectedTab=16777215|BackColorTabs=16777215|FlatBarHeight=0|FlatBodySeparationLineColor=0|FlatBorderColor=0|FlatRoundnessBottom=0|FlatRoundnessTop=0|FlatTabsSeparationLineColor=12632256|ForeColor=0|ForeColorSelectedTab=16777215|HighlightColor=14737632|HighlightColorSelectedTab=0|HighlightMode=4|HighlightModeSelectedTab=20|Style=3"
Private Const cThemeString_FlatOpenAir                   As String = "BackColorTabs=16250871|FlatBarColorHighlighted=16250871|FlatBarColorInactive=16250871|FlatBarColorSelectedTab=16731706|FlatBarHeight=4|FlatBarPosition=1|FlatBodySeparationLineColor=14869218|FlatBorderColor=16250871|FlatRoundnessBottom=0|FlatRoundnessTop=0|FlatTabsSeparationLineColor=16250871|ForeColor=0|ForeColorHighlighted=16731706|ForeColorSelectedTab=16731706|HighlightColor=16477710|HighlightMode=1|HighlightModeSelectedTab=64|Style=3|TabMousePointerHand=-1"
Private Const cThemeString_GhostTab                      As String = "BackColorSelectedTab=16250871|FlatBarColorSelectedTab=7699508|FlatBarHeight=0|FlatBodySeparationLineColor=7699508|FlatBodySeparationLineHeight=3|FlatBorderColor=7699508|FlatBorderMode=1|FlatRoundnessTabs=8|FlatTabsSeparationLineColor=11250603|ForeColorSelectedTab=16777215|HighlightColor=12766860|HighlightColorSelectedTab=7699508|HighlightMode=4|HighlightModeSelectedTab=20|Style=3|TabMousePointerHand=-1|TabSeparation=8"
Private Const cThemeString_Buttons                       As String = "BackColorSelectedTab=-2147483633|BackColorTabs=12766860|FlatBarColorHighlighted=7699508|FlatBarColorInactive=12766860|FlatBarColorSelectedTab=7699508|FlatBarPosition=1|FlatBodySeparationLineColor=11250603|FlatBodySeparationLineHeight=0|FlatBorderColor=-2147483633|FlatRoundnessBottom=0|FlatRoundnessTop=0|FlatTabsSeparationLineColor=11250603|ForeColor=0|ForeColorSelectedTab=-2147483630|HighlightColor=12766860|HighlightColorSelectedTab=-2147483633|HighlightEffectOut=1|HighlightMode=196|HighlightModeSelectedTab=196|Style=3|TabMousePointerHand=-1|TabSeparation=8"
Private Const cThemeString_Buttons2                      As String = "BackColorSelectedTab=-2147483633|BackColorTabs=12766860|FlatBarColorHighlighted=7699508|FlatBarColorInactive=12766860|FlatBarColorSelectedTab=7699508|FlatBarPosition=1|FlatBodySeparationLineColor=11250603|FlatBodySeparationLineHeight=0|FlatBorderColor=-2147483633|FlatRoundnessBottom=0|FlatRoundnessTabs=16|FlatRoundnessTop=0|FlatTabsSeparationLineColor=11250603|ForeColor=0|ForeColorSelectedTab=-2147483630|HighlightColor=12766860|HighlightColorSelectedTab=-2147483633|HighlightEffectOut=1|HighlightMode=196|HighlightModeSelectedTab=196|Style=3|TabMousePointerHand=-1|TabSeparation=8"
Private Const cThemeString_Buttons3                      As String = "BackColorSelectedTab=-2147483633|BackColorTabs=12766860|FlatBarColorHighlighted=7699508|FlatBarColorInactive=12766860|FlatBarColorSelectedTab=7699508|FlatBodySeparationLineColor=7699508|FlatBodySeparationLineHeight=3|FlatBorderColor=-2147483633|FlatRoundnessBottom=0|FlatRoundnessTop=0|FlatTabsSeparationLineColor=11250603|ForeColor=0|ForeColorSelectedTab=-2147483630|HighlightColor=12766860|HighlightColorSelectedTab=-2147483633|HighlightEffectOut=1|HighlightMode=196|HighlightModeSelectedTab=1|Style=3|TabMousePointerHand=-1|TabSeparation=8"
Private Const cThemeString_Buttons4                      As String = "BackColorSelectedTab=-2147483633|BackColorTabs=12766860|FlatBarColorHighlighted=7699508|FlatBarColorInactive=12766860|FlatBarColorSelectedTab=7699508|FlatBarHeight=5|FlatBodySeparationLineColor=7699508|FlatBorderColor=7699508|FlatBorderMode=1|FlatRoundnessBottom=4|FlatRoundnessTabs=4|FlatRoundnessTop=4|FlatTabsSeparationLineColor=-2147483633|ForeColor=0|ForeColorSelectedTab=-2147483630|HighlightColor=12766860|HighlightColorSelectedTab=-2147483633|HighlightEffectOut=1|HighlightMode=196|HighlightModeSelectedTab=64|Style=3|TabMousePointerHand=-1|TabSeparation=8"
Private Const cThemeString_Buttons5                      As String = "BackColorSelectedTab=-2147483633|BackColorTabs=12766860|FlatBarColorHighlighted=7699508|FlatBarColorInactive=12766860|FlatBarColorSelectedTab=7699508|FlatBarGripHeight=0|FlatBarHeight=5|FlatBodySeparationLineColor=7699508|FlatBorderColor=7699508|FlatBorderMode=1|FlatRoundnessBottom=4|FlatRoundnessTabs=4|FlatRoundnessTop=4|FlatTabsSeparationLineColor=-2147483633|ForeColor=0|ForeColorSelectedTab=-2147483630|HighlightColor=12766860|HighlightColorSelectedTab=-2147483633|HighlightEffectOut=1|HighlightMode=196|HighlightModeSelectedTab=64|Style=3|TabMousePointerHand=-1|TabSeparation=8"
Private Const cThemeString_Buttons6                      As String = "BackColorSelectedTab=16777215|BackColorTabs=12766860|FlatBarColorHighlighted=7699508|FlatBarColorInactive=12766860|FlatBarColorSelectedTab=7699508|FlatBarGripHeight=0|FlatBarHeight=5|FlatBodySeparationLineColor=7699508|FlatBorderColor=7699508|FlatBorderMode=1|FlatRoundnessBottom=4|FlatRoundnessTabs=4|FlatRoundnessTop=4|FlatTabsSeparationLineColor=-2147483633|ForeColor=0|HighlightColor=12766860|HighlightColorSelectedTab=-2147483633|HighlightEffectOut=1|HighlightMode=196|HighlightModeSelectedTab=64|Style=3|TabMousePointerHand=-1|TabSeparation=8"
Private Const cThemeString_Buttons7                      As String = "BackColorSelectedTab=-2147483633|BackColorTabs=10646340|FlatBarColorHighlighted=4228799|FlatBarColorInactive=10646340|FlatBarColorSelectedTab=10646340|FlatBarPosition=1|FlatBodySeparationLineColor=7434609|FlatBodySeparationLineHeight=0|FlatBorderColor=-2147483633|FlatRoundnessBottom=0|FlatRoundnessTop=0|FlatTabsSeparationLineColor=7434609|ForeColor=16777215|ForeColorSelectedTab=-2147483630|HighlightColor=10646340|HighlightColorSelectedTab=-2147483633|HighlightEffectOut=1|HighlightMode=196|HighlightModeSelectedTab=196|Style=3|TabMousePointerHand=-1|TabSeparation=8"
Private Const cThemeString_Buttons8                      As String = "BackColorSelectedTab=-2147483633|BackColorTabs=10316098|FlatBarColorHighlighted=1729514|FlatBarColorInactive=10316098|FlatBarColorSelectedTab=10316098|FlatBarPosition=1|FlatBodySeparationLineColor=7171437|FlatBodySeparationLineHeight=0|FlatBorderColor=-2147483633|FlatRoundnessBottom=0|FlatRoundnessTabs=8|FlatRoundnessTop=0|FlatTabsSeparationLineColor=7171437|ForeColor=16777215|ForeColorSelectedTab=-2147483630|HighlightColor=10316098|HighlightColorSelectedTab=-2147483633|HighlightEffectOut=1|HighlightMode=196|HighlightModeSelectedTab=196|Style=3|TabMousePointerHand=-1|TabSeparation=8"
Private Const cThemeString_Buttons9                      As String = "BackColorSelectedTab=-2147483633|BackColorTabs=8804407|FlatBarColorHighlighted=1729514|FlatBarColorInactive=8804407|FlatBarColorSelectedTab=8804407|FlatBarPosition=1|FlatBodySeparationLineColor=6184542|FlatBodySeparationLineHeight=0|FlatBorderColor=-2147483633|FlatRoundnessBottom=0|FlatRoundnessTabs=8|FlatRoundnessTop=0|FlatTabsSeparationLineColor=6184542|ForeColor=16777215|ForeColorSelectedTab=-2147483630|HighlightColor=8804407|HighlightColorSelectedTab=-2147483633|HighlightEffectOut=1|HighlightMode=196|HighlightModeSelectedTab=196|Style=3|TabMousePointerHand=-1|TabSeparation=8"
Private Const cThemeString_Buttons10                     As String = "BackColorSelectedTab=16775410|BackColorTabs=7742876|FlatBarColorHighlighted=5751007|FlatBarColorInactive=7742876|FlatBarColorSelectedTab=7742876|FlatBarPosition=1|FlatBodySeparationLineColor=7829367|FlatBodySeparationLineHeight=0|FlatBorderColor=-2147483633|FlatRoundnessBottom=0|FlatRoundnessTabs=8|FlatRoundnessTop=0|FlatTabsSeparationLineColor=7829367|ForeColor=16777215|ForeColorSelectedTab=0|HighlightColor=7742876|HighlightColorSelectedTab=-2147483633|HighlightEffectOut=1|HighlightMode=196|HighlightModeSelectedTab=196|Style=3|TabMousePointerHand=-1|TabSeparation=8"
Private Const cThemeString_Buttons11                     As String = "BackColorSelectedTab=16775410|BackColorTabs=5577584|FlatBarColorHighlighted=1283056|FlatBarColorInactive=5577584|FlatBarColorSelectedTab=5577584|FlatBarPosition=1|FlatBodySeparationLineColor=5658198|FlatBodySeparationLineHeight=0|FlatBorderColor=-2147483633|FlatRoundnessBottom=0|FlatRoundnessTabs=8|FlatRoundnessTop=0|FlatTabsSeparationLineColor=5658198|ForeColor=16777215|ForeColorSelectedTab=0|HighlightColor=5577584|HighlightColorSelectedTab=-2147483633|HighlightEffectOut=1|HighlightMode=196|HighlightModeSelectedTab=196|Style=3|TabMousePointerHand=-1|TabSeparation=8"
Private Const cThemeString_Buttons12                     As String = "BackColorSelectedTab=-2147483633|BackColorTabs=4021028|FlatBarColorHighlighted=4620782|FlatBarColorInactive=4021028|FlatBarColorSelectedTab=4021028|FlatBarPosition=1|FlatBodySeparationLineColor=7829367|FlatBodySeparationLineHeight=0|FlatBorderColor=-2147483633|FlatRoundnessBottom=0|FlatRoundnessTabs=8|FlatRoundnessTop=0|FlatTabsSeparationLineColor=7829367|ForeColor=16777215|ForeColorSelectedTab=-2147483630|HighlightColor=4021028|HighlightColorSelectedTab=-2147483633|HighlightEffectOut=1|HighlightMode=196|HighlightModeSelectedTab=196|Style=3|TabMousePointerHand=-1|TabSeparation=8"
Private Const cThemeString_WebLinks                      As String = "BackColorTabs=16250871|FlatBarColorHighlighted=16250871|FlatBarColorInactive=16250871|FlatBarColorSelectedTab=16731706|FlatBarHeight=4|FlatBarPosition=1|FlatBodySeparationLineColor=14869218|FlatBorderColor=16250871|FlatRoundnessBottom=0|FlatRoundnessTop=0|FlatTabsSeparationLineColor=16250871|ForeColor=0|ForeColorHighlighted=16731706|ForeColorSelectedTab=16777215|HighlightColor=16477710|HighlightColorSelectedTab=16477710|HighlightMode=32|HighlightModeSelectedTab=20|Style=3|TabMousePointerHand=-1"
Private Const cThemeString_WebLinks2                     As String = "BackColorTabs=16250871|FlatBarColorHighlighted=16250871|FlatBarColorInactive=16250871|FlatBarColorSelectedTab=16731706|FlatBarHeight=4|FlatBarPosition=1|FlatBodySeparationLineColor=14869218|FlatBorderColor=16731706|FlatBorderMode=1|FlatRoundnessTabs=8|FlatTabsSeparationLineColor=16250871|ForeColor=0|ForeColorHighlighted=16731706|ForeColorSelectedTab=16777215|HighlightColor=16477710|HighlightColorSelectedTab=16477710|HighlightMode=16|HighlightModeSelectedTab=20|Style=3|TabMousePointerHand=-1"
Private Const cThemeString_AnotherButtonLight            As String = "BackColorTabs=16250871|FlatBarHeight=0|FlatBodySeparationLineColor=16777215|FlatBodySeparationLineHeight=0|FlatBorderColor=16777215|FlatBorderMode=1|FlatRoundnessBottom=0|FlatRoundnessTabs=4|FlatRoundnessTop=0|FlatTabBorderColorHighlighted=8603431|FlatTabBorderColorSelectedTab=15492185|FlatTabsSeparationLineColor=16777215|ForeColor=0|HighlightColor=15461355|HighlightMode=12|HighlightModeSelectedTab=512|Style=3|TabMousePointerHand=-1"
Private Const cThemeString_AnotherButtonDark             As String = "BackColorSelectedTab=592137|BackColorTabs=3476744|FlatBarColorSelectedTab=12335619|FlatBarHeight=0|FlatBodySeparationLineHeight=0|FlatBorderColor=3476744|FlatBorderMode=1|FlatRoundnessBottom=0|FlatRoundnessTabs=4|FlatRoundnessTop=0|FlatTabBorderColorHighlighted=14195837|FlatTabBorderColorSelectedTab=14195837|FlatTabsSeparationLineColor=3476744|ForeColor=16777215|HighlightColor=4210752|HighlightMode=12|HighlightModeSelectedTab=512|Style=3|TabMousePointerHand=-1"
Private Const cThemeString_Slate                         As String = "BackColorSelectedTab=-2147483633|BackColorTabs=6837311|ForeColor=16777215|ForeColorSelectedTab=-2147483630|Style=1"
Private Const cThemeString_Brick                         As String = "BackColorSelectedTab=-2147483633|BackColorTabs=3029898|ForeColor=16777215|ForeColorSelectedTab=-2147483630|Style=0"
Private Const cThemeString_FlatRaised                    As String = "BackColorTabs=15393757|FlatBarColorSelectedTab=11562799|FlatRoundnessTabs=6|ForeColor=3549726|HighlightModeSelectedTab=330|HighlightTabExtraHeight=60|Style=3"
Private Const cThemeString_FlatMarker                    As String = "BackColorSelectedTab=16777215|BackColorTabs=16777215|FlatBarColorHighlighted=11072511|FlatBarColorInactive=16777215|FlatBarColorSelectedTab=6087935|FlatBarGripHeight=0|FlatBarHeight=9|FlatBarPosition=1|FlatBodySeparationLineColor=13684944|FlatBorderColor=16777215|FlatRoundnessBottom=0|FlatRoundnessTop=0|FlatTabsSeparationLineColor=16777215|ForeColor=2105376|HighlightMode=64|HighlightModeSelectedTab=80|Style=3"
Private Const cThemeString_FlatOutline                   As String = "BackColorSelectedTab=16777215|BackColorTabs=16777215|FlatBarHeight=0|FlatBodySeparationLineHeight=0|FlatBorderColor=16777215|FlatBorderMode=1|FlatRoundnessTabs=6|FlatRoundnessTop=0|FlatTabBodyGap=2|FlatTabBorderColorHighlighted=11642266|FlatTabBorderColorSelectedTab=15429407|FlatTabsSeparationLineColor=16777215|ForeColor=2105376|HighlightMode=4608|HighlightModeSelectedTab=4624|Style=3|TabSeparation=6"
Private Const cThemeString_FlatBlueRibbon                As String = "BackColorTabs=16118254|FlatBarColorInactive=16118254|FlatBarColorSelectedTab=12608789|FlatBarHeight=7|ForeColor=3155227|Style=3"
Private Const cThemeString_FlatRedRibbon                 As String = "BackColorTabs=15658997|FlatBarColorInactive=15658997|FlatBarColorSelectedTab=2631878|FlatBarHeight=7|ForeColor=1776432|Style=3"
Private Const cThemeString_FlatGreenRibbon               As String = "BackColorTabs=15725806|FlatBarColorInactive=15725806|FlatBarColorSelectedTab=3308846|FlatBarHeight=7|ForeColor=1978395|Style=3"
Private Const cThemeString_FlatNeon                      As String = "BackColorSelectedTab=2829099|BackColorTabs=1710618|FlatBarColorInactive=1710618|FlatBarColorSelectedTab=10544384|FlatBarHeight=6|ForeColor=15921906|HighlightColorSelectedTab=3815994|HighlightMode=4|Style=3"
Private Const cThemeString_FlatChips                     As String = "BackColorSelectedTab=16777215|BackColorTabs=15525860|FlatBarHeight=0|FlatBodySeparationLineColor=15426341|FlatBodySeparationLineHeight=3|FlatBorderColor=15426341|FlatBorderMode=1|FlatRoundnessTabs=8|FlatTabsSeparationLineColor=15525860|ForeColor=3352863|ForeColorSelectedTab=16777215|HighlightColor=14471883|HighlightColorSelectedTab=15426341|HighlightMode=4|HighlightModeSelectedTab=20|Style=3|TabMousePointerHand=-1|TabSeparation=6"
Private Const cThemeString_FlatCards                     As String = "BackColorSelectedTab=16777215|BackColorTabs=16777215|FlatBarColorHighlighted=9085168|FlatBarColorInactive=14077125|FlatBarColorSelectedTab=3037156|FlatBarHeight=4|FlatBorderColor=14077125|FlatRoundnessTabs=6|ForeColor=3813415|HighlightColor=16184305|HighlightMode=68|HighlightModeSelectedTab=80|Style=3|TabSeparation=5"
Private Const cThemeString_FlatSegments                  As String = "BackColorSelectedTab=16777215|BackColorTabs=14933207|FlatBarHeight=0|FlatBodySeparationLineColor=3813415|FlatBodySeparationLineHeight=2|FlatBorderColor=3813415|FlatBorderMode=1|FlatRoundnessTop=0|FlatTabsSeparationLineColor=14933207|ForeColor=3813415|ForeColorSelectedTab=16777215|HighlightColor=13748159|HighlightColorSelectedTab=3813415|HighlightMode=4|HighlightModeSelectedTab=20|Style=3|TabSeparation=2"
Private Const cThemeString_FlatCarbon                    As String = "BackColorSelectedTab=4603450|BackColorTabs=3353898|FlatBarColorHighlighted=1733816|FlatBarColorInactive=3353898|FlatBarColorSelectedTab=1875967|FlatBarPosition=1|FlatBodySeparationLineHeight=0|FlatBorderColor=4603450|FlatRoundnessTabs=6|FlatTabsSeparationLineColor=3353898|ForeColor=15592168|HighlightColor=4603450|HighlightEffectOut=1|HighlightMode=196|HighlightModeSelectedTab=208|Style=3|TabMousePointerHand=-1|TabSeparation=6"
Private Const cThemeString_FlatCarbonBlue                As String = "BackColorSelectedTab=4603450|BackColorTabs=3353898|FlatBarColorHighlighted=12089400|FlatBarColorInactive=3353898|FlatBarColorSelectedTab=16754253|FlatBarPosition=1|FlatBodySeparationLineHeight=0|FlatBorderColor=4603450|FlatRoundnessTabs=6|FlatTabsSeparationLineColor=3353898|ForeColor=15592168|HighlightColor=4603450|HighlightEffectOut=1|HighlightMode=196|HighlightModeSelectedTab=208|Style=3|TabMousePointerHand=-1|TabSeparation=6"
Private Const cThemeString_FlatCarbonGreen               As String = "BackColorSelectedTab=4603450|BackColorTabs=3353898|FlatBarColorHighlighted=7246625|FlatBarColorInactive=3353898|FlatBarColorSelectedTab=10079278|FlatBarPosition=1|FlatBodySeparationLineHeight=0|FlatBorderColor=4603450|FlatRoundnessTabs=6|FlatTabsSeparationLineColor=3353898|ForeColor=15592168|HighlightColor=4603450|HighlightEffectOut=1|HighlightMode=196|HighlightModeSelectedTab=208|Style=3|TabMousePointerHand=-1|TabSeparation=6"
Private Const cThemeString_FlatCarbonViolet              As String = "BackColorSelectedTab=4603450|BackColorTabs=3353898|FlatBarColorHighlighted=12081019|FlatBarColorInactive=3353898|FlatBarColorSelectedTab=16742570|FlatBarPosition=1|FlatBodySeparationLineHeight=0|FlatBorderColor=4603450|FlatRoundnessTabs=6|FlatTabsSeparationLineColor=3353898|ForeColor=15592168|HighlightColor=4603450|HighlightEffectOut=1|HighlightMode=196|HighlightModeSelectedTab=208|Style=3|TabMousePointerHand=-1|TabSeparation=6"
Private Const cThemeString_FlatSpacious                  As String = "BackColorSelectedTab=16777215|BackColorTabs=16448250|FlatBarColorHighlighted=16435112|FlatBarColorInactive=16448250|FlatBarColorSelectedTab=13788953|FlatBarPosition=1|FlatBodySeparationLineColor=14736602|FlatBorderColor=16448250|FlatRoundnessBottom=0|FlatRoundnessTop=0|FlatTabsSeparationLineColor=16448250|ForeColor=4407356|ForeColorSelectedTab=13788953|HighlightMode=64|HighlightModeSelectedTab=64|Style=3|TabPaddingHeight=8|TabPaddingWidth=12"
Private Const cThemeString_FlatPills                     As String = "FlatBarHeight=0|FlatBorderMode=1|FlatRoundnessTabs=10|FlatTabBodyGap=2|FlatTabsSeparationLineColor=-2147483633|ForeColorSelectedTab=16777215|HighlightColor=14866384|HighlightColorSelectedTab=15426341|HighlightMode=4100|HighlightModeSelectedTab=4116|Style=3|TabMousePointerHand=-1|TabSeparation=4"
Private Const cThemeString_FlatMintPills                 As String = "FlatBarHeight=0|FlatBorderMode=1|FlatRoundnessTabs=10|FlatTabBodyGap=2|FlatTabsSeparationLineColor=-2147483633|ForeColorSelectedTab=16777215|HighlightColor=14866384|HighlightColorSelectedTab=4030485|HighlightMode=4100|HighlightModeSelectedTab=4116|Style=3|TabMousePointerHand=-1|TabSeparation=4"
Private Const cThemeString_FlatCoralPills                As String = "FlatBarHeight=0|FlatBorderMode=1|FlatRoundnessTabs=10|FlatTabBodyGap=2|FlatTabsSeparationLineColor=-2147483633|ForeColorSelectedTab=16777215|HighlightColor=14866384|HighlightColorSelectedTab=803266|HighlightMode=4100|HighlightModeSelectedTab=4116|Style=3|TabMousePointerHand=-1|TabSeparation=4"
Private Const cThemeString_FlatGraphitePills             As String = "FlatBarHeight=0|FlatBorderMode=1|FlatRoundnessTabs=10|FlatTabBodyGap=2|FlatTabsSeparationLineColor=-2147483633|ForeColorSelectedTab=16777215|HighlightColor=14866384|HighlightColorSelectedTab=5325111|HighlightMode=4100|HighlightModeSelectedTab=4116|Style=3|TabMousePointerHand=-1|TabSeparation=4"
Private Const cThemeString_FlatCapsules                  As String = "FlatBarHeight=0|FlatBorderMode=1|FlatRoundnessTabs=10|FlatTabBodyGap=2|FlatTabBorderColorHighlighted=15547004|FlatTabsSeparationLineColor=-2147483633|ForeColorSelectedTab=16777215|HighlightColorSelectedTab=14231661|HighlightMode=4608|HighlightModeSelectedTab=4116|Style=3|TabMousePointerHand=-1|TabSeparation=4"
Private Const cThemeString_FlatTinted                    As String = "FlatBarHeight=0|FlatBorderMode=1|FlatRoundnessTabs=10|FlatTabBodyGap=2|FlatTabBorderColorSelectedTab=15426341|FlatTabsSeparationLineColor=-2147483633|ForeColorSelectedTab=14175773|HighlightColor=15460325|HighlightColorSelectedTab=16706267|HighlightMode=4100|HighlightModeSelectedTab=4628|Style=3|TabMousePointerHand=-1|TabSeparation=4"
Private Const cThemeString_FlatFloatingMint              As String = "FlatBarHeight=0|FlatBorderColor=-2147483633|FlatBorderMode=1|FlatRoundnessTabs=10|FlatTabBodyGap=5|FlatTabBorderColorHighlighted=10403978|FlatTabBorderColorSelectedTab=-2147483630|FlatTabsSeparationLineColor=-2147483633|ForeColorSelectedTab=16777215|HighlightColor=15266274|HighlightColorSelectedTab=4030485|HighlightMode=4612|HighlightModeSelectedTab=4116|Style=3|TabMousePointerHand=-1|TabSeparation=4"
Private Const cThemeString_FlatFloatingBlue              As String = "FlatBarHeight=0|FlatBorderColor=-2147483633|FlatBorderMode=1|FlatRoundnessTabs=10|FlatTabBodyGap=5|FlatTabBorderColorHighlighted=16101778|FlatTabBorderColorSelectedTab=-2147483630|FlatTabsSeparationLineColor=-2147483633|ForeColorSelectedTab=16777215|HighlightColor=16641253|HighlightColorSelectedTab=15426341|HighlightMode=4612|HighlightModeSelectedTab=4116|Style=3|TabMousePointerHand=-1|TabSeparation=4"
Private Const cThemeString_FlatFloatingCoral             As String = "FlatBarHeight=0|FlatBorderColor=-2147483633|FlatBorderMode=1|FlatRoundnessTabs=10|FlatTabBodyGap=5|FlatTabBorderColorHighlighted=8823008|FlatTabBorderColorSelectedTab=-2147483630|FlatTabsSeparationLineColor=-2147483633|ForeColorSelectedTab=16777215|HighlightColor=14870776|HighlightColorSelectedTab=803266|HighlightMode=4612|HighlightModeSelectedTab=4116|Style=3|TabMousePointerHand=-1|TabSeparation=4"
Private Const cThemeString_FlatFloatingGraphite          As String = "FlatBarHeight=0|FlatBorderColor=-2147483633|FlatBorderMode=1|FlatRoundnessTabs=10|FlatTabBodyGap=5|FlatTabBorderColorHighlighted=11051163|FlatTabBorderColorSelectedTab=-2147483630|FlatTabsSeparationLineColor=-2147483633|ForeColorSelectedTab=16777215|HighlightColor=15395047|HighlightColorSelectedTab=5325111|HighlightMode=4612|HighlightModeSelectedTab=4116|Style=3|TabMousePointerHand=-1|TabSeparation=4"
Private Const cThemeString_FlatTagsMint                  As String = "FlatBarHeight=0|FlatBorderColor=-2147483633|FlatBorderMode=1|FlatRoundnessTabs=10|FlatTabBodyGap=5|FlatTabBorderColorHighlighted=10403978|FlatTabBorderColorInactive=13684944|FlatTabBorderColorSelectedTab=-2147483630|FlatTabsSeparationLineColor=-2147483633|ForeColorSelectedTab=16777215|HighlightColor=15266274|HighlightColorSelectedTab=4030485|HighlightMode=12804|HighlightModeSelectedTab=4116|Style=3|TabMousePointerHand=-1|TabSeparation=4"
Private Const cThemeString_FlatTagsBlue                  As String = "FlatBarHeight=0|FlatBorderColor=-2147483633|FlatBorderMode=1|FlatRoundnessTabs=10|FlatTabBodyGap=5|FlatTabBorderColorHighlighted=16101778|FlatTabBorderColorInactive=13684944|FlatTabBorderColorSelectedTab=-2147483630|FlatTabsSeparationLineColor=-2147483633|ForeColorSelectedTab=16777215|HighlightColor=16641253|HighlightColorSelectedTab=15426341|HighlightMode=12804|HighlightModeSelectedTab=4116|Style=3|TabMousePointerHand=-1|TabSeparation=4"
Private Const cThemeString_FlatTagsCoral                 As String = "FlatBarHeight=0|FlatBorderColor=-2147483633|FlatBorderMode=1|FlatRoundnessTabs=10|FlatTabBodyGap=5|FlatTabBorderColorHighlighted=8823008|FlatTabBorderColorInactive=13684944|FlatTabBorderColorSelectedTab=-2147483630|FlatTabsSeparationLineColor=-2147483633|ForeColorSelectedTab=16777215|HighlightColor=14870776|HighlightColorSelectedTab=803266|HighlightMode=12804|HighlightModeSelectedTab=4116|Style=3|TabMousePointerHand=-1|TabSeparation=4"
Private Const cThemeString_FlatTagsGraphite              As String = "FlatBarHeight=0|FlatBorderColor=-2147483633|FlatBorderMode=1|FlatRoundnessTabs=10|FlatTabBodyGap=5|FlatTabBorderColorHighlighted=11051163|FlatTabBorderColorInactive=13684944|FlatTabBorderColorSelectedTab=-2147483630|FlatTabsSeparationLineColor=-2147483633|ForeColorSelectedTab=16777215|HighlightColor=15395047|HighlightColorSelectedTab=5325111|HighlightMode=12804|HighlightModeSelectedTab=4116|Style=3|TabMousePointerHand=-1|TabSeparation=4"
Private Const cThemeString_FlatBadgesMint                As String = "BackColorTabs=14803425|FlatBarHeight=0|FlatBorderColor=-2147483633|FlatBorderMode=1|FlatRoundnessTabs=10|FlatTabBodyGap=5|FlatTabBorderColorHighlighted=10403978|FlatTabBorderColorInactive=11915174|FlatTabBorderColorSelectedTab=-2147483630|FlatTabsSeparationLineColor=-2147483633|ForeColorSelectedTab=16777215|HighlightColor=15266274|HighlightColorSelectedTab=4030485|HighlightMode=12804|HighlightModeSelectedTab=4116|Style=3|TabMousePointerHand=-1|TabSeparation=4"
Private Const cThemeString_FlatBadgesBlue                As String = "BackColorTabs=14803425|FlatBarHeight=0|FlatBorderColor=-2147483633|FlatBorderMode=1|FlatRoundnessTabs=10|FlatTabBodyGap=5|FlatTabBorderColorHighlighted=16101778|FlatTabBorderColorInactive=16237740|FlatTabBorderColorSelectedTab=-2147483630|FlatTabsSeparationLineColor=-2147483633|ForeColorSelectedTab=16777215|HighlightColor=16641253|HighlightColorSelectedTab=15426341|HighlightMode=12804|HighlightModeSelectedTab=4116|Style=3|TabMousePointerHand=-1|TabSeparation=4"
Private Const cThemeString_FlatBadgesCoral               As String = "BackColorTabs=14803425|FlatBarHeight=0|FlatBorderColor=-2147483633|FlatBorderMode=1|FlatRoundnessTabs=10|FlatTabBodyGap=5|FlatTabBorderColorHighlighted=8823008|FlatTabBorderColorInactive=10729448|FlatTabBorderColorSelectedTab=-2147483630|FlatTabsSeparationLineColor=-2147483633|ForeColorSelectedTab=16777215|HighlightColor=14870776|HighlightColorSelectedTab=803266|HighlightMode=12804|HighlightModeSelectedTab=4116|Style=3|TabMousePointerHand=-1|TabSeparation=4"
Private Const cThemeString_FlatBadgesGraphite            As String = "BackColorTabs=14803425|FlatBarHeight=0|FlatBorderColor=-2147483633|FlatBorderMode=1|FlatRoundnessTabs=10|FlatTabBodyGap=5|FlatTabBorderColorHighlighted=11051163|FlatTabBorderColorInactive=12433331|FlatTabBorderColorSelectedTab=-2147483630|FlatTabsSeparationLineColor=-2147483633|ForeColorSelectedTab=16777215|HighlightColor=15395047|HighlightColorSelectedTab=5325111|HighlightMode=12804|HighlightModeSelectedTab=4116|Style=3|TabMousePointerHand=-1|TabSeparation=4"

' Not cached: the IDE can close a project and open another in the same process
Private Function ClientProjectFolder() As String
    Dim iHwndIDE As LongPtr
    
    iHwndIDE = IDEWindowOfThisProcess
    If iHwndIDE <> 0 Then
        ClientProjectFolder = GetClientProjectFolder(iHwndIDE)
    End If
End Function

' Debug.Assert does not exist in the compiled OCX, so being inside the IDE is asked of the
' windows instead: an IDE main window owned by this process. Several IDEs can be open at once
Private Function IDEWindowOfThisProcess() As LongPtr
    Dim iHwnd As LongPtr
    Dim iProcessId As Long
    Dim iThisProcessId As Long
    
    iThisProcessId = GetCurrentProcessId
    Do
        iHwnd = FindWindowEx(0, iHwnd, StrPtr("wndclass_desked_gsk"), 0)
        If iHwnd = 0 Then Exit Do
        GetWindowThreadProcessId iHwnd, iProcessId
        If iProcessId = iThisProcessId Then Exit Do
    Loop
    IDEWindowOfThisProcess = iHwnd
End Function

Private Function GetClientProjectFolder(ByVal nHwndIDE As LongPtr) As String
    Dim hwndMain As LongPtr
    Dim hProp As LongPtr
    Dim iObjIDE As Object
    Dim iObjVBE As Object
    
    hwndMain = nHwndIDE
    If hwndMain <> 0 Then
        hProp = GetProp(hwndMain, StrPtr("VBAutomation"))
        If hProp <> 0 Then
            CopyMemory iObjIDE, hProp, LenB(hProp)    '= VBIDE.Window
            On Error Resume Next
            Set iObjVBE = iObjIDE.VBE
            GetClientProjectFolder = iObjVBE.ActiveVBProject.FileName
            If InStr(GetClientProjectFolder, "\") > 2 Then
                GetClientProjectFolder = Left$(GetClientProjectFolder, InStrRev(GetClientProjectFolder, "\"))
            End If
            On Error GoTo 0
            hProp = 0
            CopyMemory iObjIDE, hProp, LenB(hProp)
        End If
    End If
End Function
    
' Empty for a project not saved yet: a bare file name would land in the current directory
Private Function ConfigFilePathInProjectFolder() As String
    Dim iFolder As String
    
    iFolder = ClientProjectFolder
    If iFolder <> "" Then ConfigFilePathInProjectFolder = iFolder & ConfigFileName
End Function

Public Function ConfigFileName() As String
    ConfigFileName = App.EXEName & "_Themes.ntt"
End Function

Private Sub EnsureThemesLoaded()
    If Not mThemesLoaded Then
        If mControlTypeName = "" Then Err.Raise 5558
        LoadThemesFromRegistry
        mThemesLoaded = True
    End If
    If mThemesFile Is Nothing Then
        LoadThemesFromFile
    ElseIf ConfigFilePathInProjectFolder <> mThemesFilePath Then
        LoadThemesFromFile
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
    mThemesFilePath = ConfigFilePathInProjectFolder
    If mThemesFilePath <> "" Then iStr = LoadTextFile(mThemesFilePath)
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

' Copies one control onto another. With nCopyOnlyThemedProperties it copies just the set a theme
' writes: pagNewTabThemes.PropertyPage_ApplyChanges writes onto the developer's real control and
' must not stamp the preview's non themed values on it. A line without an
' "If Not nCopyOnlyThemedProperties" guard is copied in both modes, so it declares itself themed;
' the reference is ApplyThemeToControl: what it assigns is themed
Public Sub CopyControlProperties(nCtlSrc As NewTab, nCtlDest As NewTab, Optional ByVal nCopyOnlyThemedProperties As Boolean = False)
    Dim iRedraw As Boolean
    Dim iTabCount As Long
    Dim t As Long
    
    iRedraw = nCtlDest.Redraw
    nCtlDest.Redraw = False
    
    ' Before the roundness: under Bottom its Let swaps top and bottom
    If Not nCopyOnlyThemedProperties Then nCtlDest.TabOrientation = PreviewTabOrientation(nCtlSrc.TabOrientation)
    nCtlDest.SetDefaultPropertyValuesForThemedProperties True
    
    nCtlDest.Style = nCtlSrc.Style
    ' A hooked colour travels as -1 and not as the colour: the Get answers a resolved colour, and
    ' assigning that would unhook the destination from its own container
    If Not nCopyOnlyThemedProperties Then
        If nCtlSrc.BackColor_IsFromAmbient Then nCtlDest.BackColor = -1 Else nCtlDest.BackColor = nCtlSrc.BackColor
    End If
    If nCtlSrc.BackColorTabs_IsFromAmbient Then nCtlDest.BackColorTabs = -1 Else nCtlDest.BackColorTabs = nCtlSrc.BackColorTabs
    If Not nCtlSrc.BackColorSelectedTab_IsAutomatic Then nCtlDest.BackColorSelectedTab = nCtlSrc.BackColorSelectedTab
     
    If nCtlSrc.ForeColor_IsFromAmbient Then nCtlDest.ForeColor = -1 Else nCtlDest.ForeColor = nCtlSrc.ForeColor
    
    ' An inherited colour is skipped: its Get answers a resolved colour, and copying it would make it
    ' explicit on the destination. SetDefaultPropertyValuesForThemedProperties left them automatic.
    ' The two mouse hover slots of the 'x' have no flag, they fall back to a constant
    If Not nCtlSrc.ForeColorSelectedTab_IsAutomatic Then nCtlDest.ForeColorSelectedTab = nCtlSrc.ForeColorSelectedTab
    If Not nCtlSrc.ForeColorHighlighted_IsAutomatic Then nCtlDest.ForeColorHighlighted = nCtlSrc.ForeColorHighlighted
    If Not nCtlSrc.FlatTabBorderColorHighlighted_IsAutomatic Then nCtlDest.FlatTabBorderColorHighlighted = nCtlSrc.FlatTabBorderColorHighlighted
    If Not nCtlSrc.FlatTabBorderColorSelectedTab_IsAutomatic Then nCtlDest.FlatTabBorderColorSelectedTab = nCtlSrc.FlatTabBorderColorSelectedTab
    If Not nCtlSrc.FlatTabBorderColorInactive_IsAutomatic Then nCtlDest.FlatTabBorderColorInactive = nCtlSrc.FlatTabBorderColorInactive
    If Not nCtlSrc.IconColor_IsAutomatic Then nCtlDest.IconColor = nCtlSrc.IconColor
    If Not nCtlSrc.IconColorSelectedTab_IsAutomatic Then nCtlDest.IconColorSelectedTab = nCtlSrc.IconColorSelectedTab
    If Not nCtlSrc.IconColorTabHighlighted_IsAutomatic Then nCtlDest.IconColorTabHighlighted = nCtlSrc.IconColorTabHighlighted
    If Not nCtlSrc.IconColorMouseHover_IsAutomatic Then nCtlDest.IconColorMouseHover = nCtlSrc.IconColorMouseHover
    If Not nCtlSrc.IconColorMouseHoverSelectedTab_IsAutomatic Then nCtlDest.IconColorMouseHoverSelectedTab = nCtlSrc.IconColorMouseHoverSelectedTab
    If Not nCtlSrc.CloseButtonColor_IsAutomatic Then nCtlDest.CloseButtonColor = nCtlSrc.CloseButtonColor
    If Not nCtlSrc.CloseButtonColorSelectedTab_IsAutomatic Then nCtlDest.CloseButtonColorSelectedTab = nCtlSrc.CloseButtonColorSelectedTab
    nCtlDest.CloseButtonColorMouseHover = nCtlSrc.CloseButtonColorMouseHover
    nCtlDest.CloseButtonColorMouseHoverSelectedTab = nCtlSrc.CloseButtonColorMouseHoverSelectedTab
    If Not nCtlSrc.CloseButtonColorTabHighlighted_IsAutomatic Then nCtlDest.CloseButtonColorTabHighlighted = nCtlSrc.CloseButtonColorTabHighlighted
    If Not nCtlSrc.NewTabButtonColor_IsAutomatic Then nCtlDest.NewTabButtonColor = nCtlSrc.NewTabButtonColor
    If Not nCtlSrc.NewTabButtonColorHighlighted_IsAutomatic Then nCtlDest.NewTabButtonColorHighlighted = nCtlSrc.NewTabButtonColorHighlighted
    ' Not themed and not reset by SetDefaultPropertyValuesForThemedProperties, so the automatic state
    ' travels as -1
    If Not nCopyOnlyThemedProperties Then
        If nCtlSrc.StripButtonsColor_IsAutomatic Then nCtlDest.StripButtonsColor = -1 Else nCtlDest.StripButtonsColor = nCtlSrc.StripButtonsColor
        If nCtlSrc.StripButtonsColorHighlighted_IsAutomatic Then nCtlDest.StripButtonsColorHighlighted = -1 Else nCtlDest.StripButtonsColorHighlighted = nCtlSrc.StripButtonsColorHighlighted
    End If
     
    nCtlDest.FlatBarColorSelectedTab = nCtlSrc.FlatBarColorSelectedTab
    If Not nCtlSrc.FlatBarColorInactive_IsAutomatic Then nCtlDest.FlatBarColorInactive = nCtlSrc.FlatBarColorInactive
    If Not nCtlSrc.FlatTabsSeparationLineColor_IsAutomatic Then nCtlDest.FlatTabsSeparationLineColor = nCtlSrc.FlatTabsSeparationLineColor
    If Not nCtlSrc.FlatBodySeparationLineColor_IsAutomatic Then nCtlDest.FlatBodySeparationLineColor = nCtlSrc.FlatBodySeparationLineColor
    If Not nCtlSrc.FlatBorderColor_IsAutomatic Then nCtlDest.FlatBorderColor = nCtlSrc.FlatBorderColor
    If Not nCtlSrc.FlatBarColorHighlighted_IsAutomatic Then nCtlDest.FlatBarColorHighlighted = nCtlSrc.FlatBarColorHighlighted
    If Not nCtlSrc.HighlightColor_IsAutomatic Then nCtlDest.HighlightColor = nCtlSrc.HighlightColor
    If Not nCtlSrc.HighlightColorSelectedTab_IsAutomatic Then nCtlDest.HighlightColorSelectedTab = nCtlSrc.HighlightColorSelectedTab
    If Not nCtlSrc.HighlightGradientColor_IsAutomatic Then nCtlDest.HighlightGradientColor = nCtlSrc.HighlightGradientColor
    If Not nCtlSrc.HighlightGradientColorSelectedTab_IsAutomatic Then nCtlDest.HighlightGradientColorSelectedTab = nCtlSrc.HighlightGradientColorSelectedTab
    
    nCtlDest.TabAppearance = nCtlSrc.TabAppearance
    nCtlDest.ShowRowsInPerspective = nCtlSrc.ShowRowsInPerspective
    If Not nCopyOnlyThemedProperties Then nCtlDest.BackStyle = nCtlSrc.BackStyle
    nCtlDest.HighlightMode = nCtlSrc.HighlightMode
    nCtlDest.HighlightModeSelectedTab = nCtlSrc.HighlightModeSelectedTab
    If Not nCopyOnlyThemedProperties Then nCtlDest.AutoRelocateControls = nCtlSrc.AutoRelocateControls
    nCtlDest.FlatBorderMode = nCtlSrc.FlatBorderMode
    nCtlDest.FlatBarPosition = nCtlSrc.FlatBarPosition
    nCtlDest.FlatBarHeight = nCtlSrc.FlatBarHeight
    nCtlDest.FlatBarGripHeight = nCtlSrc.FlatBarGripHeight
    nCtlDest.FlatTabBodyGap = nCtlSrc.FlatTabBodyGap
    nCtlDest.FlatBodySeparationLineHeight = nCtlSrc.FlatBodySeparationLineHeight
    nCtlDest.TabSeparation = nCtlSrc.TabSeparation
    nCtlDest.TabPaddingWidth = nCtlSrc.TabPaddingWidth
    nCtlDest.TabPaddingHeight = nCtlSrc.TabPaddingHeight
    nCtlDest.HighlightTabExtraHeight = nCtlSrc.HighlightTabExtraHeight
    nCtlDest.HighlightTabCaptionShift = nCtlSrc.HighlightTabCaptionShift
    nCtlDest.TabMaxWidth = nCtlSrc.TabMaxWidth
    nCtlDest.TabMinWidth = nCtlSrc.TabMinWidth
    nCtlDest.TabGeneralFixedWidth = nCtlSrc.TabGeneralFixedWidth
    nCtlDest.TabHeightStyleForHorizontal = nCtlSrc.TabHeightStyleForHorizontal
    nCtlDest.FlatRoundnessTop = nCtlSrc.FlatRoundnessTop
    nCtlDest.FlatRoundnessBottom = nCtlSrc.FlatRoundnessBottom
    nCtlDest.FlatRoundnessTabs = nCtlSrc.FlatRoundnessTabs
    If Not nCopyOnlyThemedProperties Then nCtlDest.ShowDisabledState = nCtlSrc.ShowDisabledState
    nCtlDest.ShowFocusRect = nCtlSrc.ShowFocusRect
    nCtlDest.HighlightEffect = nCtlSrc.HighlightEffect
    nCtlDest.HighlightEffectOut = nCtlSrc.HighlightEffectOut
    If Not nCopyOnlyThemedProperties Then nCtlDest.WordWrap = nCtlSrc.WordWrap
    nCtlDest.SoftEdges = nCtlSrc.SoftEdges
    nCtlDest.TabMousePointerHand = nCtlSrc.TabMousePointerHand
    ' For the previews that show pictures (pagNewTabIcons and pagNewTabTabs)
    If Not nCopyOnlyThemedProperties Then nCtlDest.MaskColor = nCtlSrc.MaskColor
    If Not nCopyOnlyThemedProperties Then nCtlDest.UseMaskColor = nCtlSrc.UseMaskColor
    If Not nCopyOnlyThemedProperties Then nCtlDest.ScalePicturesForDPI = nCtlSrc.ScalePicturesForDPI
    ' The buttons, for the previews. NewTabButton is resolved and not copied raw because a preview
    ' is never in a TDI mode - see NewTabButtonEffective
    If Not nCopyOnlyThemedProperties Then nCtlDest.CloseButton = nCtlSrc.CloseButton
    If Not nCopyOnlyThemedProperties Then nCtlDest.CloseButtonPosition = nCtlSrc.CloseButtonPosition
    If Not nCopyOnlyThemedProperties Then nCtlDest.StripButtonsRow = nCtlSrc.StripButtonsRow
    If Not nCopyOnlyThemedProperties Then nCtlDest.TabsOverflow = nCtlSrc.TabsOverflow
    If Not nCopyOnlyThemedProperties Then nCtlDest.NewTabButtonPosition = nCtlSrc.NewTabButtonPosition
    If Not nCopyOnlyThemedProperties Then nCtlDest.NewTabButton = IIf(nCtlSrc.NewTabButtonEffective, ntYes, ntNo)
    If Not nCopyOnlyThemedProperties Then nCtlDest.CopyGlyphSettingsFrom nCtlSrc
    ' A tab property, the only one here: the pages seed each tab's icon themselves, and without
    ' this their samples draw every color icon in one color
    If Not nCopyOnlyThemedProperties Then
        iTabCount = nCtlSrc.Tabs
        If nCtlDest.Tabs < iTabCount Then iTabCount = nCtlDest.Tabs
        For t = 0 To iTabCount - 1
            nCtlDest.TabIconInColor(t) = nCtlSrc.TabIconInColor(t)
        Next
    End If
    If Not nCopyOnlyThemedProperties Then
        If FontsAreEqual(nCtlSrc.AmbientFont, nCtlSrc.Font) Then
            Set nCtlDest.Font = nCtlDest.AmbientFont
        Else
            Set nCtlDest.Font = CloneFont(nCtlSrc.Font)
        End If
    End If
    ' TabWidthStyle is the only thing this condition is about: a TDI strip sizes its tabs by
    ' itself, so a preview that is not in TDI must not carry its width style across
    If (nCtlDest.TDIMode <> ntTDIModeNone) = (nCtlSrc.TDIMode <> ntTDIModeNone) Then
        nCtlDest.TabWidthStyle = nCtlSrc.TabWidthStyle
    End If
    nCtlDest.Redraw = iRedraw
End Sub

' The orientation of a page sample. The rotated strips would not fit in a sample, so Left and
' Right are drawn as Top, and LeftTop and RightBottom as TopRight, the same layout without the
' rotation. Their FlatBarPosition and roundness already mean the same as in Top: only the Bottom
' pair takes them as screen sides. The rest are shown as they are
Public Function PreviewTabOrientation(ByVal nTabOrientation As NTTabOrientationConstants) As NTTabOrientationConstants
    Select Case nTabOrientation
        Case ssTabOrientationLeft, ssTabOrientationRight
            PreviewTabOrientation = ssTabOrientationTop
        Case ntTabOrientationLeftTop, ntTabOrientationRightBottom
            PreviewTabOrientation = ntTabOrientationTopRight
        Case Else
            PreviewTabOrientation = nTabOrientation
    End Select
End Function

' For a sample that only shows the strip buttons: a *Horizontal column is drawn as a row, at the
' same end of the strip
Public Function PreviewTabOrientationAsRow(ByVal nTabOrientation As NTTabOrientationConstants) As NTTabOrientationConstants
    Select Case nTabOrientation
        Case ntTabOrientationLeftHorizontal, ntTabOrientationRightBottomHorizontal
            PreviewTabOrientationAsRow = ntTabOrientationTopRight
        Case ntTabOrientationRightHorizontal, ntTabOrientationLeftBottomHorizontal
            PreviewTabOrientationAsRow = ssTabOrientationTop
        Case Else
            PreviewTabOrientationAsRow = nTabOrientation
    End Select
End Function

' Top of a sample in its container, with the strip and the selected tab in sight. The selected row
' is the one next to the body: with the strip at the bottom it goes to the upper edge, and a strip
' shorter than the container is shown whole. A *Horizontal strip is a column, moved up only as
' far as the selected tab needs
Public Function PreviewTop(ByVal nPreview As NewTab, ByVal nContainerHeight As Single) As Single
    Dim iTop As Single
    Dim iTabSel As Long
    
    iTabSel = nPreview.TabSel
    Select Case nPreview.TabOrientation
        Case ssTabOrientationBottom, ntTabOrientationBottomRight
            If iTabSel > -1 Then iTop = -nPreview.GetTabTop(iTabSel)
            If (iTop + nPreview.Height) < nContainerHeight Then iTop = nContainerHeight - nPreview.Height
        Case ntTabOrientationLeftHorizontal, ntTabOrientationRightHorizontal, ntTabOrientationLeftBottomHorizontal, ntTabOrientationRightBottomHorizontal
            If iTabSel > -1 Then
                If (nPreview.GetTabTop(iTabSel) + nPreview.GetTabHeight(iTabSel)) > nContainerHeight Then iTop = nContainerHeight - (nPreview.GetTabTop(iTabSel) + nPreview.GetTabHeight(iTabSel))
            End If
        Case Else
            If (nPreview.Rows > 1) And (iTabSel > -1) Then iTop = nContainerHeight - (nPreview.GetTabTop(iTabSel) + nPreview.GetTabHeight(iTabSel))
    End Select
    If iTop > 0 Then iTop = 0
    PreviewTop = iTop
End Function

' Whether the sample has tabs out of sight above or below, which the pages that move it let the user
' drag into view
Public Function PreviewNeedsVerticalDrag(ByVal nPreview As NewTab, ByVal nContainerHeight As Single) As Boolean
    Select Case nPreview.TabOrientation
        Case ntTabOrientationLeftHorizontal, ntTabOrientationRightHorizontal, ntTabOrientationLeftBottomHorizontal, ntTabOrientationRightBottomHorizontal
            PreviewNeedsVerticalDrag = nPreview.Height > nContainerHeight
        Case Else
            PreviewNeedsVerticalDrag = (nPreview.Height - nPreview.ClientHeight) > 800
    End Select
End Function

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
    
' "B" and "F" are read but no longer written, so that a theme saved with an older build keeps
' applying. "A" means -1, not set
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
    
    If PropertyExists("ForeColor") Then nCtl.ForeColor = GetPropertyValue("ForeColor")
    
    ' -1 and not the ambient value: a theme that says nothing about the tabs colour leaves it
    ' hooked to the container
    If PropertyExists("BackColorTabs") Then nCtl.BackColorTabs = GetPropertyValue("BackColorTabs") Else nCtl.BackColorTabs = -1
    
    If PropertyExists("BackColorSelectedTab") Then
        nCtl.BackColorSelectedTab = GetPropertyValue("BackColorSelectedTab")
    ElseIf PropertyExists("BackColorTabSel") Then
        nCtl.BackColorSelectedTab = GetPropertyValue("BackColorTabSel")
    Else
        ' -1: a theme that says nothing about the selected tab wants it derived
        nCtl.BackColorSelectedTab = -1
    End If
    
    
    If PropertyExists("ForeColorSelectedTab") Then
        nCtl.ForeColorSelectedTab = GetPropertyValue("ForeColorSelectedTab")
    ElseIf PropertyExists("ForeColorTabSel") Then
        nCtl.ForeColorSelectedTab = GetPropertyValue("ForeColorTabSel")
    Else
        nCtl.ForeColorSelectedTab = -1
    End If
    
    If PropertyExists("ForeColorHighlighted") Then nCtl.ForeColorHighlighted = GetPropertyValue("ForeColorHighlighted") Else nCtl.ForeColorHighlighted = -1
    If PropertyExists("FlatTabBorderColorHighlighted") Then
        nCtl.FlatTabBorderColorHighlighted = GetPropertyValue("FlatTabBorderColorHighlighted")
    ElseIf PropertyExists("FlatTabBorderColorHighlight") Then
        nCtl.FlatTabBorderColorHighlighted = GetPropertyValue("FlatTabBorderColorHighlight")
    Else
        nCtl.FlatTabBorderColorHighlighted = -1
    End If
    
    If PropertyExists("FlatTabBorderColorSelectedTab") Then
        nCtl.FlatTabBorderColorSelectedTab = GetPropertyValue("FlatTabBorderColorSelectedTab")
    ElseIf PropertyExists("FlatTabBorderColorTabSel") Then
        nCtl.FlatTabBorderColorSelectedTab = GetPropertyValue("FlatTabBorderColorTabSel")
    Else
        nCtl.FlatTabBorderColorSelectedTab = -1
    End If
    If PropertyExists("FlatTabBorderColorInactive") Then nCtl.FlatTabBorderColorInactive = GetPropertyValue("FlatTabBorderColorInactive") Else nCtl.FlatTabBorderColorInactive = -1
    
    If PropertyExists("IconColor") Then nCtl.IconColor = GetPropertyValue("IconColor") Else nCtl.IconColor = -1
    
    ' -1 where the theme says nothing: each icon color then falls back, inside the control, to the
    ' caption color of the same state
    If PropertyExists("IconColorSelectedTab") Then
        nCtl.IconColorSelectedTab = GetPropertyValue("IconColorSelectedTab")
    ElseIf PropertyExists("IconColorTabSel") Then
        nCtl.IconColorSelectedTab = GetPropertyValue("IconColorTabSel")
    Else
        nCtl.IconColorSelectedTab = -1
    End If
    
    If PropertyExists("IconColorTabHighlighted") Then nCtl.IconColorTabHighlighted = GetPropertyValue("IconColorTabHighlighted") Else nCtl.IconColorTabHighlighted = -1
    
    ' The mouse hover pair falls back to the icon color of the analogous state
    If PropertyExists("IconColorMouseHover") Then nCtl.IconColorMouseHover = GetPropertyValue("IconColorMouseHover") Else nCtl.IconColorMouseHover = -1
    If PropertyExists("IconColorMouseHoverSelectedTab") Then
        nCtl.IconColorMouseHoverSelectedTab = GetPropertyValue("IconColorMouseHoverSelectedTab")
    ElseIf PropertyExists("IconColorMouseHoverTabSel") Then
        nCtl.IconColorMouseHoverSelectedTab = GetPropertyValue("IconColorMouseHoverTabSel")
    Else
        nCtl.IconColorMouseHoverSelectedTab = -1
    End If
    
    ' Left automatic, the close button follows the caption color of its state and not the icon color.
    ' The two mouse hover slots fall back to the red they are born with
    If PropertyExists("CloseButtonColor") Then nCtl.CloseButtonColor = GetPropertyValue("CloseButtonColor") Else nCtl.CloseButtonColor = -1
    If PropertyExists("CloseButtonColorSelectedTab") Then nCtl.CloseButtonColorSelectedTab = GetPropertyValue("CloseButtonColorSelectedTab") Else nCtl.CloseButtonColorSelectedTab = -1
    If PropertyExists("CloseButtonColorTabHighlighted") Then nCtl.CloseButtonColorTabHighlighted = GetPropertyValue("CloseButtonColorTabHighlighted") Else nCtl.CloseButtonColorTabHighlighted = -1
    If PropertyExists("CloseButtonColorMouseHover") Then nCtl.CloseButtonColorMouseHover = GetPropertyValue("CloseButtonColorMouseHover") Else nCtl.CloseButtonColorMouseHover = cPropDef_CloseButtonColorMouseHover
    If PropertyExists("CloseButtonColorMouseHoverSelectedTab") Then nCtl.CloseButtonColorMouseHoverSelectedTab = GetPropertyValue("CloseButtonColorMouseHoverSelectedTab") Else nCtl.CloseButtonColorMouseHoverSelectedTab = cPropDef_CloseButtonColorMouseHover
    ' The '+' as a tab sits on the tabs, so it is themed like them and follows the caption when the
    ' theme says nothing. The strip buttons sit on BackColor, which is not themed, so no Else: a theme
    ' saved before they left the themed set still carries them, one without them leaves them alone
    If PropertyExists("NewTabButtonColor") Then nCtl.NewTabButtonColor = GetPropertyValue("NewTabButtonColor") Else nCtl.NewTabButtonColor = -1
    If PropertyExists("NewTabButtonColorHighlighted") Then nCtl.NewTabButtonColorHighlighted = GetPropertyValue("NewTabButtonColorHighlighted") Else nCtl.NewTabButtonColorHighlighted = -1
    If PropertyExists("StripButtonsColor") Then nCtl.StripButtonsColor = GetPropertyValue("StripButtonsColor")
    If PropertyExists("StripButtonsColorHighlighted") Then nCtl.StripButtonsColorHighlighted = GetPropertyValue("StripButtonsColorHighlighted")
    
    If PropertyExists("FlatBarColorSelectedTab") Then
        nCtl.FlatBarColorSelectedTab = GetPropertyValue("FlatBarColorSelectedTab")
    ElseIf PropertyExists("FlatBarColorTabSel") Then
        nCtl.FlatBarColorSelectedTab = GetPropertyValue("FlatBarColorTabSel")
    End If
    
    If PropertyExists("FlatBarColorInactive") Then nCtl.FlatBarColorInactive = GetPropertyValue("FlatBarColorInactive")
    If PropertyExists("FlatTabsSeparationLineColor") Then nCtl.FlatTabsSeparationLineColor = GetPropertyValue("FlatTabsSeparationLineColor")
    If PropertyExists("FlatBodySeparationLineColor") Then nCtl.FlatBodySeparationLineColor = GetPropertyValue("FlatBodySeparationLineColor")
    If PropertyExists("FlatBorderColor") Then nCtl.FlatBorderColor = GetPropertyValue("FlatBorderColor")
    If PropertyExists("FlatBarColorHighlighted") Then
        nCtl.FlatBarColorHighlighted = GetPropertyValue("FlatBarColorHighlighted")
    ElseIf PropertyExists("FlatBarColorHighlight") Then
        nCtl.FlatBarColorHighlighted = GetPropertyValue("FlatBarColorHighlight")
    End If
    If PropertyExists("HighlightColor") Then nCtl.HighlightColor = GetPropertyValue("HighlightColor")
    
    If PropertyExists("HighlightColorSelectedTab") Then
        nCtl.HighlightColorSelectedTab = GetPropertyValue("HighlightColorSelectedTab")
    ElseIf PropertyExists("HighlightColorTabSel") Then
        nCtl.HighlightColorSelectedTab = GetPropertyValue("HighlightColorTabSel")
    End If
    If PropertyExists("HighlightGradientColor") Then nCtl.HighlightGradientColor = GetPropertyValue("HighlightGradientColor")
    If PropertyExists("HighlightGradientColorSelectedTab") Then nCtl.HighlightGradientColorSelectedTab = GetPropertyValue("HighlightGradientColorSelectedTab")
    
    If PropertyExists("Style") Then nCtl.Style = GetPropertyValue("Style")
    If PropertyExists("TabAppearance") Then nCtl.TabAppearance = GetPropertyValue("TabAppearance")
    
    For Each iProp In nThemeData
        Select Case iProp.Name
            Case "TabWidthStyle"
                If nCtl.TDIMode = ntTDIModeNone Then
                    nCtl.TabWidthStyle = Val(iProp.Value)
                End If
            Case "TabHeightStyleForHorizontal"
                nCtl.TabHeightStyleForHorizontal = Val(iProp.Value)
            Case "ShowRowsInPerspective"
                nCtl.ShowRowsInPerspective = Val(iProp.Value)
            Case "HighlightMode"
                nCtl.HighlightMode = Val(iProp.Value)
            Case "HighlightModeSelectedTab", "HighlightModeTabSel"
                nCtl.HighlightModeSelectedTab = Val(iProp.Value)
            Case "FlatBorderMode"
                nCtl.FlatBorderMode = Val(iProp.Value)
            Case "FlatBarPosition"
                nCtl.FlatBarPosition = Val(iProp.Value)
            Case "FlatBarHeight"
                nCtl.FlatBarHeight = Val(iProp.Value)
            Case "FlatBarGripHeight"
                nCtl.FlatBarGripHeight = Val(iProp.Value)
            Case "FlatTabBodyGap"
                nCtl.FlatTabBodyGap = Val(iProp.Value)
            Case "FlatBodySeparationLineHeight"
                nCtl.FlatBodySeparationLineHeight = Val(iProp.Value)
            Case "TabSeparation"
                nCtl.TabSeparation = Val(iProp.Value)
            Case "TabPaddingWidth"
                nCtl.TabPaddingWidth = Val(iProp.Value)
            Case "TabPaddingHeight"
                nCtl.TabPaddingHeight = Val(iProp.Value)
            Case "HighlightTabExtraHeight"
                nCtl.HighlightTabExtraHeight = Val(iProp.Value)
            Case "HighlightTabCaptionShift"
                nCtl.HighlightTabCaptionShift = Val(iProp.Value)
            Case "TabMaxWidth"
                nCtl.TabMaxWidth = Val(iProp.Value)
            Case "TabMinWidth"
                nCtl.TabMinWidth = Val(iProp.Value)
            Case "TabGeneralFixedWidth"
                nCtl.TabGeneralFixedWidth = Val(iProp.Value)
            Case "FlatRoundnessTop"
                nCtl.FlatRoundnessTop = Val(iProp.Value)
            Case "FlatRoundnessBottom"
                nCtl.FlatRoundnessBottom = Val(iProp.Value)
            Case "FlatRoundnessTabs"
                nCtl.FlatRoundnessTabs = Val(iProp.Value)
            Case "ShowFocusRect"
                nCtl.ShowFocusRect = CBool(Val(iProp.Value))
            Case "HighlightEffect"
                nCtl.HighlightEffect = AnimationSpeedFromStoredValue(Val(iProp.Value))
            Case "HighlightEffectOut"
                nCtl.HighlightEffectOut = AnimationSpeedFromStoredValue(Val(iProp.Value))
            Case "SoftEdges"
                nCtl.SoftEdges = CBool(Val(iProp.Value))
            Case "TabMousePointerHand"
                nCtl.TabMousePointerHand = CBool(Val(iProp.Value))
        End Select
    Next
    nCtl.Redraw = iRedraw
End Sub

' No ambient parameters on purpose: nothing in here compares against the container
Public Function GetThemeStringFromControl(nCtl As NewTab, Optional nHash As String) As String
    Dim iPropsStr() As String
    Dim ub As Long
    Dim iTheme As NewTabTheme
    Dim c As Long
    
    c = -1
    ub = 100
    ReDim iPropsStr(ub)
    
    If nCtl.Style <> cPropDef_Style Then AddPropStrToArray iPropsStr, c, "Style", nCtl.Style
    If nCtl.Style <> ntStyleWindows Then
        If nCtl.TabAppearance <> cPropDef_TabAppearance Then AddPropStrToArray iPropsStr, c, "TabAppearance", nCtl.TabAppearance
        If nCtl.HighlightEffect <> cPropDef_HighlightEffect Then AddPropStrToArray iPropsStr, c, "HighlightEffect", Val(Str$(CLng(nCtl.HighlightEffect)))
        If nCtl.HighlightEffectOut <> cPropDef_HighlightEffectOut Then AddPropStrToArray iPropsStr, c, "HighlightEffectOut", Val(Str$(CLng(nCtl.HighlightEffectOut)))
    End If
    If nCtl.HighlightMode <> cPropDef_HighlightMode Then AddPropStrToArray iPropsStr, c, "HighlightMode", nCtl.HighlightMode
    If nCtl.HighlightModeSelectedTab <> cPropDef_HighlightModeSelectedTab Then AddPropStrToArray iPropsStr, c, "HighlightModeSelectedTab", nCtl.HighlightModeSelectedTab
    If nCtl.TabWidthStyle <> cPropDef_TabWidthStyle Then AddPropStrToArray iPropsStr, c, "TabWidthStyle", nCtl.TabWidthStyle
    If nCtl.TabHeightStyleForHorizontal <> cPropDef_TabHeightStyleForHorizontal Then AddPropStrToArray iPropsStr, c, "TabHeightStyleForHorizontal", nCtl.TabHeightStyleForHorizontal
    If nCtl.ShowRowsInPerspective <> cPropDef_ShowRowsInPerspective Then AddPropStrToArray iPropsStr, c, "ShowRowsInPerspective", nCtl.ShowRowsInPerspective
    If nCtl.TabSeparation <> cPropDef_TabSeparation Then AddPropStrToArray iPropsStr, c, "TabSeparation", Val(Str$(nCtl.TabSeparation))
    If nCtl.TabPaddingWidth <> cPropDef_TabPaddingWidth Then AddPropStrToArray iPropsStr, c, "TabPaddingWidth", Val(Str$(nCtl.TabPaddingWidth))
    If nCtl.TabPaddingHeight <> cPropDef_TabPaddingHeight Then AddPropStrToArray iPropsStr, c, "TabPaddingHeight", Val(Str$(nCtl.TabPaddingHeight))
    If Not nCtl.HighlightTabExtraHeightIsDefault Then AddPropStrToArray iPropsStr, c, "HighlightTabExtraHeight", Val(Str$(nCtl.HighlightTabExtraHeight))
    If Not nCtl.HighlightTabCaptionShiftIsDefault Then AddPropStrToArray iPropsStr, c, "HighlightTabCaptionShift", Val(Str$(nCtl.HighlightTabCaptionShift))
    If nCtl.TabMaxWidth <> cPropDef_TabMaxWidth Then AddPropStrToArray iPropsStr, c, "TabMaxWidth", Val(Str$(nCtl.TabMaxWidth))
    If nCtl.TabMinWidth <> cPropDef_TabMinWidth Then AddPropStrToArray iPropsStr, c, "TabMinWidth", Val(Str$(nCtl.TabMinWidth))
    If Not nCtl.TabGeneralFixedWidthIsDefault Then AddPropStrToArray iPropsStr, c, "TabGeneralFixedWidth", Val(Str$(nCtl.TabGeneralFixedWidth))
    If nCtl.ShowFocusRect <> cPropDef_ShowFocusRect Then AddPropStrToArray iPropsStr, c, "ShowFocusRect", Val(Str$(CLng(nCtl.ShowFocusRect)))
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
        If nCtl.FlatTabBodyGap <> cPropDef_FlatTabBodyGap Then AddPropStrToArray iPropsStr, c, "FlatTabBodyGap", nCtl.FlatTabBodyGap
        If nCtl.FlatBodySeparationLineHeight <> cPropDef_FlatBodySeparationLineHeight Then AddPropStrToArray iPropsStr, c, "FlatBodySeparationLineHeight", nCtl.FlatBodySeparationLineHeight
        If nCtl.FlatRoundnessTop <> cPropDef_FlatRoundnessTop Then AddPropStrToArray iPropsStr, c, "FlatRoundnessTop", Val(Str$(nCtl.FlatRoundnessTop))
        If nCtl.FlatRoundnessBottom <> cPropDef_FlatRoundnessBottom Then AddPropStrToArray iPropsStr, c, "FlatRoundnessBottom", Val(Str$(nCtl.FlatRoundnessBottom))
        If nCtl.FlatRoundnessTabs <> cPropDef_FlatRoundnessTabs Then AddPropStrToArray iPropsStr, c, "FlatRoundnessTabs", Val(Str$(nCtl.FlatRoundnessTabs))
    End If
    
    ' colors
    ' One rule: a slot that is inheriting (its Ambient or _IsAutomatic flag is on) writes nothing, and
    ' one that is not writes its number. Never deduced by comparing against the container.
    ' Only ForeColor and BackColorTabs can be hooked to the container
    If Not nCtl.ForeColor_IsFromAmbient Then
        AddPropStrToArray iPropsStr, c, "ForeColor", nCtl.ForeColor
    End If
    ' Paired with the Else in ApplyThemeToControl. The two sides are changed together or a theme
    ' starts carrying a value that is already implied, or losing one that is not
    If Not nCtl.BackColorTabs_IsFromAmbient Then
        AddPropStrToArray iPropsStr, c, "BackColorTabs", nCtl.BackColorTabs
    End If
    If Not nCtl.BackColorSelectedTab_IsAutomatic Then
        AddPropStrToArray iPropsStr, c, "BackColorSelectedTab", nCtl.BackColorSelectedTab
    End If
    If Not nCtl.ForeColorSelectedTab_IsAutomatic Then
        AddPropStrToArray iPropsStr, c, "ForeColorSelectedTab", nCtl.ForeColorSelectedTab
    End If
    If Not nCtl.ForeColorHighlighted_IsAutomatic Then
        AddPropStrToArray iPropsStr, c, "ForeColorHighlighted", nCtl.ForeColorHighlighted
    End If
    If Not nCtl.FlatTabBorderColorHighlighted_IsAutomatic Then
        AddPropStrToArray iPropsStr, c, "FlatTabBorderColorHighlighted", nCtl.FlatTabBorderColorHighlighted
    End If
    If Not nCtl.FlatTabBorderColorSelectedTab_IsAutomatic Then
        AddPropStrToArray iPropsStr, c, "FlatTabBorderColorSelectedTab", nCtl.FlatTabBorderColorSelectedTab
    End If
    If Not nCtl.FlatTabBorderColorInactive_IsAutomatic Then
        AddPropStrToArray iPropsStr, c, "FlatTabBorderColorInactive", nCtl.FlatTabBorderColorInactive
    End If
    If Not nCtl.IconColor_IsAutomatic Then
        AddPropStrToArray iPropsStr, c, "IconColor", nCtl.IconColor
    End If
    If Not nCtl.IconColorSelectedTab_IsAutomatic Then
        AddPropStrToArray iPropsStr, c, "IconColorSelectedTab", nCtl.IconColorSelectedTab
    End If
    If Not nCtl.IconColorTabHighlighted_IsAutomatic Then
        AddPropStrToArray iPropsStr, c, "IconColorTabHighlighted", nCtl.IconColorTabHighlighted
    End If
    If Not nCtl.IconColorMouseHover_IsAutomatic Then
        AddPropStrToArray iPropsStr, c, "IconColorMouseHover", nCtl.IconColorMouseHover
    End If
    If Not nCtl.IconColorMouseHoverSelectedTab_IsAutomatic Then
        AddPropStrToArray iPropsStr, c, "IconColorMouseHoverSelectedTab", nCtl.IconColorMouseHoverSelectedTab
    End If
    If Not nCtl.CloseButtonColor_IsAutomatic Then
        AddPropStrToArray iPropsStr, c, "CloseButtonColor", nCtl.CloseButtonColor
    End If
    If Not nCtl.CloseButtonColorSelectedTab_IsAutomatic Then
        AddPropStrToArray iPropsStr, c, "CloseButtonColorSelectedTab", nCtl.CloseButtonColorSelectedTab
    End If
    If Not nCtl.CloseButtonColorTabHighlighted_IsAutomatic Then
        AddPropStrToArray iPropsStr, c, "CloseButtonColorTabHighlighted", nCtl.CloseButtonColorTabHighlighted
    End If
    ' Written against the red they fall back to, not against the icon colors
    If nCtl.CloseButtonColorMouseHover <> cPropDef_CloseButtonColorMouseHover Then
        AddPropStrToArray iPropsStr, c, "CloseButtonColorMouseHover", nCtl.CloseButtonColorMouseHover
    End If
    If nCtl.CloseButtonColorMouseHoverSelectedTab <> cPropDef_CloseButtonColorMouseHover Then
        AddPropStrToArray iPropsStr, c, "CloseButtonColorMouseHoverSelectedTab", nCtl.CloseButtonColorMouseHoverSelectedTab
    End If
    If Not nCtl.NewTabButtonColor_IsAutomatic Then
        AddPropStrToArray iPropsStr, c, "NewTabButtonColor", nCtl.NewTabButtonColor
    End If
    If Not nCtl.NewTabButtonColorHighlighted_IsAutomatic Then
        AddPropStrToArray iPropsStr, c, "NewTabButtonColorHighlighted", nCtl.NewTabButtonColorHighlighted
    End If
    If Not nCtl.HighlightColor_IsAutomatic Then
        AddPropStrToArray iPropsStr, c, "HighlightColor", nCtl.HighlightColor
    End If
    If Not nCtl.HighlightColorSelectedTab_IsAutomatic Then
        AddPropStrToArray iPropsStr, c, "HighlightColorSelectedTab", nCtl.HighlightColorSelectedTab
    End If
    If Not nCtl.HighlightGradientColor_IsAutomatic Then
        AddPropStrToArray iPropsStr, c, "HighlightGradientColor", nCtl.HighlightGradientColor
    End If
    If Not nCtl.HighlightGradientColorSelectedTab_IsAutomatic Then
        AddPropStrToArray iPropsStr, c, "HighlightGradientColorSelectedTab", nCtl.HighlightGradientColorSelectedTab
    End If
    
    If nCtl.Style = ntStyleFlat Then
        If Not nCtl.FlatBarColorInactive_IsAutomatic Then
            AddPropStrToArray iPropsStr, c, "FlatBarColorInactive", nCtl.FlatBarColorInactive
        End If
        If Not nCtl.FlatTabsSeparationLineColor_IsAutomatic Then
            AddPropStrToArray iPropsStr, c, "FlatTabsSeparationLineColor", nCtl.FlatTabsSeparationLineColor
        End If
        If Not nCtl.FlatBodySeparationLineColor_IsAutomatic Then
            AddPropStrToArray iPropsStr, c, "FlatBodySeparationLineColor", nCtl.FlatBodySeparationLineColor
        End If
        If Not nCtl.FlatBorderColor_IsAutomatic Then
            AddPropStrToArray iPropsStr, c, "FlatBorderColor", nCtl.FlatBorderColor
        End If
        If Not nCtl.FlatBarColorHighlighted_IsAutomatic Then
            AddPropStrToArray iPropsStr, c, "FlatBarColorHighlighted", nCtl.FlatBarColorHighlighted
        End If
        ' The one flat colour with no automatic state: its reference is the fixed default
        If nCtl.FlatBarColorSelectedTab <> cPropDef_FlatBarColorSelectedTab Then
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
    AddDefaultTheme "Sea Green", cThemeString_FlatSeaGreen
    AddDefaultTheme "Cherry", cThemeString_FlatCherry
    AddDefaultTheme "Amber", cThemeString_FlatAmber
    AddDefaultTheme "Violet", cThemeString_FlatViolet
    AddDefaultTheme "Teal", cThemeString_FlatTeal
    AddDefaultTheme "Graphite", cThemeString_FlatGraphite
    AddDefaultTheme "Clay", cThemeString_FlatClay
    AddDefaultTheme "Lavender", cThemeString_FlatLavender
    AddDefaultTheme "Emerald", cThemeString_FlatEmerald
    AddDefaultTheme "Red Wine", cThemeString_FlatRedWine
    AddDefaultTheme "Deep Waters", cThemeString_FlatDeepWaters
    AddDefaultTheme "Candlelight", cThemeString_FlatCandlelight
    AddDefaultTheme "Moonlight", cThemeString_FlatMoonlight
    AddDefaultTheme "Midnight", cThemeString_FlatMidnight
    AddDefaultTheme "Espresso", cThemeString_FlatEspresso
    AddDefaultTheme "Twilight", cThemeString_FlatTwilight
    AddDefaultTheme "Pitch Black", cThemeString_FlatPitchBlack
    AddDefaultTheme "Ink", cThemeString_FlatInk
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
    AddDefaultTheme "Buttons 12", cThemeString_Buttons12
    AddDefaultTheme "Web Links", cThemeString_WebLinks
    AddDefaultTheme "Web Links 2", cThemeString_WebLinks2
    AddDefaultTheme "Another Button Light", cThemeString_AnotherButtonLight
    AddDefaultTheme "Another Button Dark", cThemeString_AnotherButtonDark
    AddDefaultTheme "Slate", cThemeString_Slate
    AddDefaultTheme "Brick", cThemeString_Brick
    AddDefaultTheme "Raised", cThemeString_FlatRaised
    AddDefaultTheme "Marker", cThemeString_FlatMarker
    AddDefaultTheme "Outline", cThemeString_FlatOutline
    AddDefaultTheme "Blue Ribbon", cThemeString_FlatBlueRibbon
    AddDefaultTheme "Red Ribbon", cThemeString_FlatRedRibbon
    AddDefaultTheme "Green Ribbon", cThemeString_FlatGreenRibbon
    AddDefaultTheme "Neon", cThemeString_FlatNeon
    AddDefaultTheme "Chips", cThemeString_FlatChips
    AddDefaultTheme "Cards", cThemeString_FlatCards
    AddDefaultTheme "Segments", cThemeString_FlatSegments
    AddDefaultTheme "Carbon", cThemeString_FlatCarbon
    AddDefaultTheme "Carbon Blue", cThemeString_FlatCarbonBlue
    AddDefaultTheme "Carbon Green", cThemeString_FlatCarbonGreen
    AddDefaultTheme "Carbon Violet", cThemeString_FlatCarbonViolet
    AddDefaultTheme "Spacious", cThemeString_FlatSpacious
    AddDefaultTheme "Pills", cThemeString_FlatPills
    AddDefaultTheme "Mint Pills", cThemeString_FlatMintPills
    AddDefaultTheme "Coral Pills", cThemeString_FlatCoralPills
    AddDefaultTheme "Graphite Pills", cThemeString_FlatGraphitePills
    AddDefaultTheme "Capsules", cThemeString_FlatCapsules
    AddDefaultTheme "Tinted", cThemeString_FlatTinted
    AddDefaultTheme "Floating Mint", cThemeString_FlatFloatingMint
    AddDefaultTheme "Floating Blue", cThemeString_FlatFloatingBlue
    AddDefaultTheme "Floating Coral", cThemeString_FlatFloatingCoral
    AddDefaultTheme "Floating Graphite", cThemeString_FlatFloatingGraphite
    AddDefaultTheme "Tags Mint", cThemeString_FlatTagsMint
    AddDefaultTheme "Tags Blue", cThemeString_FlatTagsBlue
    AddDefaultTheme "Tags Coral", cThemeString_FlatTagsCoral
    AddDefaultTheme "Tags Graphite", cThemeString_FlatTagsGraphite
    AddDefaultTheme "Badges Mint", cThemeString_FlatBadgesMint
    AddDefaultTheme "Badges Blue", cThemeString_FlatBadgesBlue
    AddDefaultTheme "Badges Coral", cThemeString_FlatBadgesCoral
    AddDefaultTheme "Badges Graphite", cThemeString_FlatBadgesGraphite
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

