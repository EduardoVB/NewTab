VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCommonDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const DC_PAPERS = 2
'Const DC_PAPERNAMES = 16
Const DC_PAPERSIZE = 3

Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, ByRef lpdwProcessId As Long) As Long

'Private Declare Function ResetDC Lib "gdi32" Alias "ResetDCA" (ByVal hDC As Long, lpInitData As Any) As Long
Private Declare Function DeviceCapabilities Lib "winspool.drv" Alias "DeviceCapabilitiesA" (ByVal lpDeviceName As String, ByVal lpPort As String, ByVal iIndex As Long, lpOutput As Any, lpDevMode As Any) As Long

Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long
Private Const CLR_INVALID = -1

'API function called by ChooseColor method
Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As T_ChooseColor) As Long

'API function called by ChooseFont method
Private Declare Function CHOOSEFONT Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As T_CHOOSEFONT) As Long

'API function inside ShowHelp method
Private Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long

'API function called by ShowOpen method
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As T_OpenFilename) As Long

'API function called by ShowSave method
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As T_OpenFilename) As Long

'API function called by ShowPrint method
Private Declare Function PrintDlg Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As T_PrintDlg) As Long


'API function to retrieve extended error information
Private Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long

Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Const LOCALE_USER_DEFAULT As Long = &H400&
Private Const LOCALE_IMEASURE As Long = &HD&

'API memory functions
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function lstrlenPtr Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long

Private Type BrowseInfo
   hwndOwner As Long
   pIDLRoot As Long
   pszDisplayName As String
   lpszTitle As String
   ulFlags As Long
   lpfnCallback As Long
   lParam As Long
   iImage As Long
End Type

Private Type SHITEMID
    cb As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As SHITEMID
End Type

Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BrowseInfo) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Private Declare Function SHGetFolderLocation Lib "shell32" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwReserved As Long, pidl As ITEMIDLIST) As Long
Private Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long
Private Declare Function SHParseDisplayName Lib "shell32.dll" (ByVal pszName As Long, ByVal pbc As Long, ByRef ppidl As Long, ByVal sfgaoIn As Long, ByRef psfgaoOut As Long) As Long

Private Const S_OK As Long = 0&
Private Const CSIDL_PERSONAL As Long = &H5
Private Const CSIDL_DRIVES As Long = &H11
Private Const CSIDL_PRINTERS As Long = &H4

'constants for API memory functions
Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_ZEROINIT = &H40
Private Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)

'data buffer for the ChooseColor function
Private Type T_ChooseColor
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        RGBResult As Long
        lpCustColors As Long
        Flags As Long
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type

'constants for LOGFONT
Private Const LF_FACESIZE As Long = 32
'Private Const LF_FULLFACESIZE As Long = 64

'Private Const FW_THIN As Long = 100
'Private Const FW_EXTRALIGHT As Long = 200
'Private Const FW_LIGHT As Long = 300
Private Const FW_NORMAL As Long = 400
'Private Const FW_MEDIUM As Long = 500
'Private Const FW_SEMIBOLD As Long = 600
Private Const FW_BOLD As Long = 700
'Private Const FW_EXTRABOLD As Long = 800
'Private Const FW_HEAVY As Long = 900


'data buffer for the ChooseFont function
Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(0 To LF_FACESIZE - 1) As Byte
End Type

'data buffer for the ChooseFont function
Private Type T_CHOOSEFONT
        lStructSize As Long
        hwndOwner As Long
        hDC As Long
        lpLogFont As Long
        iPointSize As Long
        Flags As Long
        rgbColors As Long
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
        hInstance As Long
        lpszStyle As String
        nFontType As Integer
        MISSING_ALIGNMENT As Integer
        nSizeMin As Long
        nSizeMax As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type T_PAGESETUPDLG
    lStructSize                 As Long
    hwndOwner                   As Long
    hDevMode                    As Long
    hDevNames                   As Long
    Flags                       As Long
    ptPaperSize                 As POINTAPI
    rtMinMargin                 As RECT
    rtMargin                    As RECT
    hInstance                   As Long
    lCustData                   As Long
    lpfnPageSetupHook           As Long
    lpfnPagePaintHook           As Long
    lpPageSetupTemplateName     As Long
    hPageSetupTemplate          As Long
End Type

Private Declare Function PageSetupDlg Lib "COMDLG32" Alias "PageSetupDlgA" (lppage As T_PAGESETUPDLG) As Boolean
    
Private Const CC_ENABLEHOOK = &H10&
Private Const CC_ENABLETEMPLATE = &H20&
Private Const CC_RGBINIT = &H1&

Private Const CF_PRINTERFONTS As Long = &H2
Private Const CF_SCREENFONTS As Long = &H1
Private Const CF_INITTOLOGFONTSTRUCT As Long = &H40&
'Private Const CF_BOTH As Long = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Private Const CF_EFFECTS As Long = &H100
Private Const CF_APPLY As Long = &H200
Private Const CF_ENABLEHOOK As Long = &H8
Private Const CF_ENABLETEMPLATE As Long = &H10
Private Const CF_LIMITSIZE As Long = &H2000

Private Const PD_ENABLEPRINTHOOK = &H1000&
Private Const PD_ENABLESETUPHOOK = &H2000&
Private Const PD_ENABLEPRINTTEMPLATE = &H4000&
Private Const PD_ENABLESETUPTEMPLATE = &H8000&

Private Const PSD_ENABLEPAGEPAINTHOOK As Long = &H40000
Private Const PSD_ENABLEPAGESETUPHOOK As Long = &H2000&
Private Const PSD_ENABLEPAGESETUPTEMPLATE As Long = &H8000&


'data buffer for the GetOpenFileName and GetSaveFileName functions
Private Type T_OpenFilename
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        mFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        Flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type


'data buffer for the PrintDlg function
Private Type T_PrintDlg
        lStructSize As Long
        hwndOwner As Long
        hDevMode As Long
        hDevNames As Long
        hDC As Long
        Flags As Long
        nFromPage As Integer
        nToPage As Integer
        nMinPage As Integer
        nMaxPage As Integer
        nCopies As Integer
        hInstance As Long
        lCustData As Long
        lpfnPrintHook As Long
        lpfnSetupHook As Long
        lpPrintTemplateName As String
        lpSetupTemplateName As String
        hPrintTemplate As Long
        hSetupTemplate As Long
End Type

'constants for color dialog

'Private Const CDERR_DIALOGFAILURE = &HFFFF
'Private Const CDERR_FINDRESFAILURE = &H6
'Private Const CDERR_GENERALCODES = &H0
'Private Const CDERR_INITIALIZATION = &H2
'Private Const CDERR_LOADRESFAILURE = &H7
'Private Const CDERR_LOADSTRFAILURE = &H5
'Private Const CDERR_LOCKRESFAILURE = &H8
'Private Const CDERR_MEMALLOCFAILURE = &H9
'Private Const CDERR_MEMLOCKFAILURE = &HA
'Private Const CDERR_NOHINSTANCE = &H4
'Private Const CDERR_NOHOOK = &HB
'Private Const CDERR_NOTEMPLATE = &H3
'Private Const CDERR_REGISTERMSGFAIL = &HC
'Private Const CDERR_STRUCTSIZE = &H1


'constants for file dialog

'Private Const FNERR_BUFFERTOOSMALL As Long = &H3003
'Private Const FNERR_FILENAMECODES As Long = &H3000
'Private Const FNERR_INVALIDFILENAME As Long = &H3002
'Private Const FNERR_SUBCLASSFAILURE As Long = &H3001

' other constants
Private Const CLEARTYPE_QUALITY As Byte = 6
Private Const DEFAULT_CHARSET As Byte = 1

Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32

Private Const PSD_MARGINS As Long = &H2&
Private Const PSD_INHUNDREDTHSOFMILLIMETERS = &H8&
Private Const PSD_INTHOUSANDTHSOFINCHES = &H4&

Private Type DEVMODE
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

Private Type DEVNAMES
    wDriverOffset As Integer
    wDeviceOffset As Integer
    wOutputOffset As Integer
    wDefault As Integer
End Type


Public Enum CommonDialogsFlags

' File Open, Save as
    cdlbOFNAllowMultiselect = &H200&
    cdlbOFNCreatePrompt = &H2000&
    cdlbOFNExplorer = &H80000
    cdlbOFNExtensionDifferent = &H400&
    cdlbOFNFileMustExist = &H1000&
    cdlbOFNHelpButton = &H10&
    cdlbOFNHideReadOnly = &H4&
    cdlbOFNLongNames = &H200000
    cdlbOFNNoChangeDir = &H8&
    cdlbOFNNoDereferenceLinks = &H100000
    cdlbOFNNoLongNames = &H40000
    cdlbOFNNoReadOnlyReturn = &H8000&
    cdlbOFNNoValidate = &H100&
    cdlbOFNOverwritePrompt = &H2&
    cdlbOFNPathMustExist = &H800&
    cdlbOFNReadOnly = &H1&
    cdlbOFNShareAware = &H4000&
 
'PrintFlags
    cdlbPDAllPages = &H0&
    cdlbPDSelection = &H1&
    cdlbPDPageNums = &H2&
    cdlbPDNoSelection = &H4&
    cdlbPDNoPageNums = &H8&
    cdlbPDCollate = &H10&
    cdlbPDPrintToFile = &H20&
    cdlbPDPrintSetup = &H40&
    cdlbPDNoWarning = &H80&
    cdlbPDReturnDC = &H100&
    cdlbPDReturnIC = &H200&
    cdlbPDReturnDefault = &H400&
    cdlbPDHelpButton = &H800&
    cdlbPDUseDevModeCopies = &H40000
    cdlbPDDisablePrintToFile = &H80000
    cdlbPDHidePrintToFile = &H100000
 
'ColorFlags
    cdlbCCFullOpen = &H2&
    cdlbCCShowHelpButton = &H8&
    cdlbCCPreventFullOpen = &H4&
    cdlbCCRGBInit = &H1&
 
'FontFlags
    cdlbCFANSIOnly = &H400&
    cdlbCFApply = &H200&
    cdlbCFBoth = &H3&
    cdlbCFEffects = &H100&
    cdlbCFFixedPitchOnly = &H4000&
    cdlbCFForceFontExist = &H10000
    cdlbCFHelpButton = &H4&
    cdlbCFLimitSize = &H2000&
    cdlbCFNoFaceSel = &H80000
    cdlbCFNoSimulations = &H1000&
    cdlbCFNoSizeSel = &H200000
    cdlbCFNoStyleSel = &H100000
    cdlbCFNoVectorFonts = &H800&
    cdlbCFPrinterFonts = &H2&
    cdlbCFScalableOnly = &H20000
    cdlbCFScreenFonts = &H1&
    cdlbCFTTOnly = &H40000
    cdlbCFWYSIWYG = &H8000&
    
' Page setup flags
    cdlbPSDefaultMinMargins = &H0&
    cdlbPSDisableMargins = &H10&
    cdlbPSDisableOrientation = &H100&
    cdlbPSDisablePagePainting = &H80000
    cdlbPSDisablePaper = &H200&
    cdlbPSDisablePrinter = &H20&
    'cdlbPSMargins = &H2&
    cdlbPSMinMargins = &H1&
    cdlbPSNoNetworkButton = &H200000
    cdlbPSNoWarning = &H80&
    cdlbPSReturnDefault = &H400&
    cdlbPSShowHelp = &H800&
    
' Show Folder
    cdlbSFReturnOnlyFSDirs = &H1&
    cdlbSFDontGoBelowDomain = &H2&
    cdlbSFStatusText = &H4&
    cdlbSFReturnFSAncestors = &H8&
    cdlbSFEditBox = &H10&
    cdlbSFValidate = &H20&
    cdlbSFNewDialogStyle = &H40&
    cdlbSFUseNewUI = (cdlbSFNewDialogStyle Or cdlbSFEditBox)
    cdlbSFBrowseIncludeURLs = &H80&
    cdlbSFUAHint = &H100&
    cdlbSFNoNewFolderButton = &H200&
    cdlbSFNoTranslateTargets = &H400&
    cdlbSFNoReturnOnlyFSDirs = &H800&
    cdlbSFBrowseForComputer = &H1000&
    cdlbSFBrowseForPrinter = &H2000&
    cdlbSFBrowseIncludeFiles = &H4000&
    cdlbSFShareable = &H8000&
End Enum

Public Enum cdeCommonDialogExFileFlagsConstants
    cdeOFNAllowMultiselect = CommonDialogsFlags.cdlbOFNAllowMultiselect
    cdeOFNCreatePrompt = CommonDialogsFlags.cdlbOFNCreatePrompt
    cdeOFNExplorer = CommonDialogsFlags.cdlbOFNExplorer
    cdeOFNExtensionDifferent = CommonDialogsFlags.cdlbOFNExtensionDifferent
    cdeOFNFileMustExist = CommonDialogsFlags.cdlbOFNFileMustExist
    cdeOFNHelpButton = CommonDialogsFlags.cdlbOFNHelpButton
    cdeOFNHideReadOnly = CommonDialogsFlags.cdlbOFNHideReadOnly
    cdeOFNLongNames = CommonDialogsFlags.cdlbOFNLongNames
    cdeOFNNoChangeDir = CommonDialogsFlags.cdlbOFNNoChangeDir
    cdeOFNNoDereferenceLinks = CommonDialogsFlags.cdlbOFNNoDereferenceLinks
    cdeOFNNoLongNames = CommonDialogsFlags.cdlbOFNNoLongNames
    cdeOFNNoReadOnlyReturn = CommonDialogsFlags.cdlbOFNNoReadOnlyReturn
    cdeOFNNoValidate = CommonDialogsFlags.cdlbOFNNoValidate
    cdeOFNOverwritePrompt = CommonDialogsFlags.cdlbOFNOverwritePrompt
    cdeOFNPathMustExist = CommonDialogsFlags.cdlbOFNPathMustExist
    cdeOFNReadOnly = CommonDialogsFlags.cdlbOFNReadOnly
    cdeOFNShareAware = CommonDialogsFlags.cdlbOFNShareAware
End Enum

Public Enum cdeCommonDialogExPrinterFlagsConstants
    cdePDAllPages = CommonDialogsFlags.cdlbPDAllPages
    cdePDSelection = CommonDialogsFlags.cdlbPDSelection
    cdePDPageNums = CommonDialogsFlags.cdlbPDPageNums
    cdePDNoSelection = CommonDialogsFlags.cdlbPDNoSelection
    cdePDNoPageNums = CommonDialogsFlags.cdlbPDNoPageNums
    cdePDCollate = CommonDialogsFlags.cdlbPDCollate
    cdePDPrintToFile = CommonDialogsFlags.cdlbPDPrintToFile
    cdePDPrintSetup = CommonDialogsFlags.cdlbPDPrintSetup
    cdePDNoWarning = CommonDialogsFlags.cdlbPDNoWarning
    cdePDReturnDC = CommonDialogsFlags.cdlbPDReturnDC
    cdePDReturnIC = CommonDialogsFlags.cdlbPDReturnIC
    cdePDReturnDefault = CommonDialogsFlags.cdlbPDReturnDefault
    cdePDHelpButton = CommonDialogsFlags.cdlbPDHelpButton
    cdePDUseDevModeCopies = CommonDialogsFlags.cdlbPDUseDevModeCopies
    cdePDDisablePrintToFile = CommonDialogsFlags.cdlbPDDisablePrintToFile
    cdePDHidePrintToFile = CommonDialogsFlags.cdlbPDHidePrintToFile
End Enum

Public Enum cdeCommonDialogExColorFlagsConstants
    cdeCCFullOpen = CommonDialogsFlags.cdlbCCFullOpen
    cdeCCShowHelpButton = CommonDialogsFlags.cdlbCCShowHelpButton
    cdeCCPreventFullOpen = CommonDialogsFlags.cdlbCCPreventFullOpen
'    cdeCCRGBInit = CommonDialogsFlags.cdlbCCRGBInit
End Enum

Public Enum cdeCommonDialogExFontFlagsConstants
    cdeCFANSIOnly = CommonDialogsFlags.cdlbCFANSIOnly
    cdeCFApply = CommonDialogsFlags.cdlbCFApply
    cdeCFBoth = CommonDialogsFlags.cdlbCFBoth
    cdeCFEffects = CommonDialogsFlags.cdlbCFEffects
    cdeCFFixedPitchOnly = CommonDialogsFlags.cdlbCFFixedPitchOnly
    cdeCFForceFontExist = CommonDialogsFlags.cdlbCFForceFontExist
    cdeCFHelpButton = CommonDialogsFlags.cdlbCFHelpButton
    cdeCFLimitSize = CommonDialogsFlags.cdlbCFLimitSize
    cdeCFNoFaceSel = CommonDialogsFlags.cdlbCFNoFaceSel
    cdeCFNoSimulations = CommonDialogsFlags.cdlbCFNoSimulations
    cdeCFNoSizeSel = CommonDialogsFlags.cdlbCFNoSizeSel
    cdeCFNoStyleSel = CommonDialogsFlags.cdlbCFNoStyleSel
    cdeCFNoVectorFonts = CommonDialogsFlags.cdlbCFNoVectorFonts
    cdeCFPrinterFonts = CommonDialogsFlags.cdlbCFPrinterFonts
    cdeCFScalableOnly = CommonDialogsFlags.cdlbCFScalableOnly
    cdeCFScreenFonts = CommonDialogsFlags.cdlbCFScreenFonts
    cdeCFTTOnly = CommonDialogsFlags.cdlbCFTTOnly
    cdeCFWYSIWYG = CommonDialogsFlags.cdlbCFWYSIWYG
End Enum

Public Enum cdeCommonDialogExPageSetupFlagsConstants
    cdePSDefaultMinMargins = CommonDialogsFlags.cdlbPSDefaultMinMargins
    cdePSDisableMargins = CommonDialogsFlags.cdlbPSDisableMargins
    cdePSDisableOrientation = CommonDialogsFlags.cdlbPSDisableOrientation
    cdePSDisablePagePainting = CommonDialogsFlags.cdlbPSDisablePagePainting
    cdePSDisablePaper = CommonDialogsFlags.cdlbPSDisablePaper
    cdePSDisablePrinter = CommonDialogsFlags.cdlbPSDisablePrinter
'    cdePSMargins = CommonDialogsFlags.cdlbPSMargins
    cdePSMinMargins = CommonDialogsFlags.cdlbPSMinMargins
    cdePSNoNetworkButton = CommonDialogsFlags.cdlbPSNoNetworkButton
    cdePSNoWarning = CommonDialogsFlags.cdlbPSNoWarning
    cdePSReturnDefault = CommonDialogsFlags.cdlbPSReturnDefault
    cdePSShowHelp = CommonDialogsFlags.cdlbPSShowHelp
End Enum

Public Enum cdeCommonDialogExFolderFlagsConstants
    cdeSFReturnOnlyFSDirs = CommonDialogsFlags.cdlbSFReturnOnlyFSDirs
    cdeSFDontGoBelowDomain = CommonDialogsFlags.cdlbSFDontGoBelowDomain
    cdeSFStatusText = CommonDialogsFlags.cdlbSFStatusText
    cdeSFReturnFSAncestors = CommonDialogsFlags.cdlbSFReturnFSAncestors
    cdeSFEditBox = CommonDialogsFlags.cdlbSFEditBox
    cdeSFValidate = CommonDialogsFlags.cdlbSFValidate
    cdeSFNewDialogStyle = CommonDialogsFlags.cdlbSFNewDialogStyle
    cdeSFUseNewUI = CommonDialogsFlags.cdlbSFUseNewUI
    cdeSFBrowseIncludeURLs = CommonDialogsFlags.cdlbSFBrowseIncludeURLs
    cdeSFUAHint = CommonDialogsFlags.cdlbSFUAHint
    cdeSFNoNewFolderButton = CommonDialogsFlags.cdlbSFNoNewFolderButton
    cdeSFNoTranslateTargets = CommonDialogsFlags.cdlbSFNoTranslateTargets
    cdeSFNoReturnOnlyFSDirs = CommonDialogsFlags.cdlbSFNoReturnOnlyFSDirs
    cdeSFBrowseForComputer = CommonDialogsFlags.cdlbSFBrowseForComputer
    cdeSFBrowseForPrinter = CommonDialogsFlags.cdlbSFBrowseForPrinter
    cdeSFBrowseIncludeFiles = CommonDialogsFlags.cdlbSFBrowseIncludeFiles
    cdeSFShareable = CommonDialogsFlags.cdlbSFShareable
    
End Enum


' Other Public enumerations
Public Enum cdePageOrientationConstants
    vbPRORPrinterDefault = 0&
    vbPRORPortrait = 1&
    vbPRORLandscape = 2&
End Enum

Public Enum cdeColorModeConstants
    vbPRCMPrinterDefault = 0&
    vbPRCMColor = 2&
    vbPRCMMonochrome = 1&
End Enum

Public Enum cdePaperBinConstants
    vbPRBNPrinterDefault = 0&
    vbPRBNUpper = 1&
    vbPRBNLower = 2&
    vbPRBNMiddle = 3&
    vbPRBNManual = 4&
    vbPRBNEnvelope = 5&
    vbPRBNEnvManual = 6&
    vbPRBNAuto = 7&
    vbPRBNTractor = 8&
    vbPRBNSmallFmt = 9&
    vbPRBNLargeFmt = 10&
    vbPRBNLargeCapacity = 11&
    vbPRBNCassette = 14&
End Enum

Public Enum cdePaperSizeConstants
    vbPRPSPrinterDefault = 0&
    vbPRPSLetter = 1&
    vbPRPSLetterSmall = 2&
    vbPRPSTabloid = 3&
    vbPRPSLedger = 4&
    vbPRPSLegal = 5&
    vbPRPSStatement = 6&
    vbPRPSExecutive = 7&
    vbPRPSA3 = 8&
    vbPRPSA4 = 9&
    vbPRPSA4Small = 10&
    vbPRPSA5 = 11&
    vbPRPSB4 = 12&
    vbPRPSB5 = 13&
    vbPRPSFolio = 14&
    vbPRPSQuarto = 15&
    vbPRPS10x14 = 16&
    vbPRPS11x17 = 17&
    vbPRPSNote = 18&
    vbPRPSEnv9 = 19&
    vbPRPSEnv10 = 20&
    vbPRPSEnv11 = 21&
    vbPRPSEnv12 = 22&
    vbPRPSEnv14 = 23&
    vbPRPSCSheet = 24&
    vbPRPSDSheet = 25&
    vbPRPSESheet = 26&
    vbPRPSEnvDL = 27&
    vbPRPSEnvC3 = 29&
    vbPRPSEnvC4 = 30&
    vbPRPSEnvC5 = 28&
    vbPRPSEnvC6 = 31&
    vbPRPSEnvC65 = 32&
    vbPRPSEnvB4 = 33&
    vbPRPSEnvB5 = 34&
    vbPRPSEnvB6 = 35&
    vbPRPSEnvItaly = 36&
    vbPRPSEnvMonarch = 37&
    vbPRPSEnvPersonal = 38&
    vbPRPSFanfoldUS = 39&
    vbPRPSFanfoldStdGerman = 40&
    vbPRPSFanfoldLglGerman = 41&
    vbPRPSUser = 256&
End Enum

Public Enum cdePrintQualityConstants
    vbPRPQPrinterDefault = 0&
    vbPRPQDraft = -1&
    vbPRPQLow = -2&
    vbPRPQMedium = -3&
    vbPRPQHigh = -4&
End Enum

Public Enum cdeDuplexConstants
    vbPRDPPrinterDefault = 0&
    vbPRDPSimplex = 1&
    vbPRDPHorizontal = 2&
    vbPRDPVertical = 3&
End Enum

Public Enum cdeUnits
    vbMillimeters = 6&
    vbInches = 5&
End Enum

Public Enum cdeUnitsForUser
    cdeMUUserLocale = 0&
    cdeMUMillimeters = 6&
    cdeMUInches = 5&
End Enum

'Properties

Private mAction As Integer
Private mCancelError As Boolean
Private mColor As Long
Private mCopies As Long
Private mCollate As Boolean
Private mDefaultExt As String
Private mDialogTitle As String
Private mFileName As String
Private mFileTitle As String
Private mFilter As String
Private mFilterIndex As Integer
Private mFlags As Long
Private mFont As StdFont
'Private mFont.Bold As Boolean
'Private mFont.Italic As Boolean
'Private mFont.Name As String
'Private mFont.Size As Long
'Private mFont.Strikethrough As Boolean
'Private mFont.Underline As Boolean
Private mFromPage As Long
Private mhDc As Long
Private mHelpCommand As Long
Private mHelpContext As Long
Private mHelpFile As String
Private mHelpKey As String
Private mInitDir As String
Private mMax As Long
Private mMaxFileSize As Long
Private mMin As Long
'Private mPrinterDefault As Integer
Private mToPage As Long
Private mOrientation As Long
Private mCustomColors(0 To 15) As Long
Private mAutoSaveCustomColors As Boolean
Private mCustomColorsLoaded As Boolean

Private mApiReturn As Long
Private mExtendedError As Long
Private mCancelled As Boolean

' Added properties special for printer and page setup
Private mPaperSize As Long
Private mPaperBin As Long
Private mPrintQuality As Long
Private mColorMode As Long
Private mDuplex As Long
' read only
Private mDeviceName As String
Private mDriverName As String
Private mPort As String
Private mPaperWidth As Long
Private mPaperHeight As Long
Private mDefaultPaperWidth As Long
Private mDefaultPaperHeight As Long

' added properties special for page setup
Private mLeftMargin As Single
Private mRightMargin As Single
Private mTopMargin As Single
Private mBottomMargin As Single
Private mMinLeftMargin As Single
Private mMinRightMargin As Single
Private mMinTopMargin As Single
Private mMinBottomMargin As Single
'Private mMarginSet As Boolean
Private mUnits As Long
Private mUnitsForUser As Long

Private Const cLeftMarginDefault As Single = 20
Private Const cRightMarginDefault As Single = 15
Private Const cTopMarginDefault As Single = 20
Private Const cBottomMarginDefault As Single = 20


' added property for show folder
Private mRootFolder As String
Private mDialogHeader As String

' auxiliary variables
Private mDevMode As DEVMODE
Private mDevNames As DEVNAMES
Private mDevModePtr As Long
Private mDevNamesPtr As Long
Private mhDevNames As Long
Private mhDevMode As Long
Private mPageSet As Boolean

Private mAmbientUserMode As Boolean

Public Sub ShowColor(Optional ByVal nFlags As cdeCommonDialogExColorFlagsConstants = -1)
    'display the color dialog box
    
    Dim tChooseColor As T_ChooseColor
    Dim lCustomColorSize As Long
    Dim lCustomColorAddress As Long
    Dim lMemHandle As Long
    Dim iHwndOwner As Long
    Dim iFlags As Long
    
    On Error GoTo ShowColorError
    
    iHwndOwner = GetActiveWindowHwnd
    iFlags = mFlags
    If nFlags <> -1 Then
        iFlags = iFlags Or nFlags
    End If
    
    mAction = 3  'Action property - ShowColor
    mApiReturn = 0  'APIReturn property
    mExtendedError = 0  'ExtendedError property
    
'    If mAutoSaveCustomColors Then
'        If Not mCustomColorsLoaded Then LoadCustomColors
'    End If
    
    tChooseColor.lStructSize = Len(tChooseColor)
    
    tChooseColor.hwndOwner = iHwndOwner
    
    ' Get size of memory needed for custom colors
    lCustomColorSize = Len(mCustomColors(0)) * 16
    ' Get a global memory block to hold a copy of the custom colors
    lMemHandle = GlobalAlloc(GHND, lCustomColorSize)
    
    If lMemHandle = 0 Then
        Exit Sub
    End If
    ' Lock the custom color's global memory block
    lCustomColorAddress = GlobalLock(lMemHandle)
    If lCustomColorAddress = 0 Then
        Exit Sub
    End If
    ' Copy custom colors to the global memory block
    CopyMemory ByVal lCustomColorAddress, mCustomColors(0), lCustomColorSize
 
    tChooseColor.lpCustColors = lCustomColorAddress
    
    tChooseColor.Flags = iFlags And Not (CC_ENABLEHOOK Or CC_ENABLETEMPLATE) Or CC_RGBINIT
        
    tChooseColor.RGBResult = TranslateAColor(mColor)
        
    
    '***    call the ChooseColor API function
    mApiReturn = ChooseColor(tChooseColor)
    
    '***    handle return from ChooseColor API function
    mCancelled = False
    Select Case mApiReturn
        
        Case 0  'user canceled
        mCancelled = True
        If mCancelError Then
            'generate an error
            On Error GoTo 0
            Err.Raise 32755, "Cancel Pressed"
            Exit Sub
        End If
        
        Case 1  'user selected a color
            'update property buffer
            mColor = tChooseColor.RGBResult
        
'            CopyMemory mCustomColors(0), ByVal lCustomColorAddress, lCustomColorSize
'            If mAutoSaveCustomColors Then
'                SaveCustomColors
'            End If
        
        Case Else   'an error occured
            'call CommDlgExtendedError
            mExtendedError = CommDlgExtendedError
        
    End Select
    
    GlobalFree lMemHandle
    
Exit Sub

ShowColorError:
    Exit Sub
End Sub

Public Sub ShowFont(Optional ByVal nFlags As cdeCommonDialogExFontFlagsConstants = -1)
    'display the font dialog box
    
    Dim tLogFont As LOGFONT
    Dim tChooseFont As T_CHOOSEFONT
    Dim lLogFontSize As Long
    Dim lLogFontAddress As Long
    Dim lMemHandle As Long
    Dim i As Long
    Dim iHwndOwner As Long
    Dim iFlags As Long
    
    On Error GoTo ShowFontError
    
    iHwndOwner = GetActiveWindowHwnd
    iFlags = mFlags
    If nFlags <> -1 Then
        iFlags = iFlags Or nFlags
    End If
    
    '***    init property buffers
    
    mAction = 4  'Action property - ShowFont
    mApiReturn = 0  'APIReturn property
    mExtendedError = 0  'ExtendedError property

    
    '***    prepare tChooseFont data
        
    
    tLogFont.lfHeight = mFont.Size * 20 / Screen.TwipsPerPixelY * -1 ' PointsPerTwip = 1440 / 72 = 20
    'tLogFont.lfWidth As Long
    'tLogFont.lfEscapement As Long
    'tLogFont.lfOrientation As Long
    
    'tLogFont.lfWeight As Long - init from FontBold property
    If mFont.Bold Then
        tLogFont.lfWeight = FW_BOLD
    Else
        tLogFont.lfWeight = FW_NORMAL
    End If
    
    'tLogFont.lfItalic As Byte - init from FontItalic property
    If mFont.Italic Then
        tLogFont.lfItalic = 1
    End If
    
    'tLogFont.lfUnderline As Byte - init from FontUnderline property
    If mFont.Underline Then
        tLogFont.lfUnderline = 1
    End If

    'tLogFont.lfStrikeOut As Byte - init from FontStrikethru property
    If mFont.Strikethrough Then
        tLogFont.lfStrikeOut = 1
    End If


    tLogFont.lfCharSet = DEFAULT_CHARSET
    'tLogFont.lfOutPrecision As Byte
    'tLogFont.lfClipPrecision As Byte
    tLogFont.lfQuality = CLEARTYPE_QUALITY
    'tLogFont.lfPitchAndFamily As Byte
    
    For i = 0 To Len(mFont.Name) - 1
        If i > 31 Then Exit For
        tLogFont.lfFaceName(i) = Asc(Mid$(mFont.Name, i + 1, 1))
    Next
    
    lLogFontSize = Len(tLogFont)
    
    
    tChooseFont.lStructSize = Len(tChooseFont)
    tChooseFont.hwndOwner = iHwndOwner
    tChooseFont.rgbColors = mColor
'    tChooseFont.lCustData = 0
'    tChooseFont.lpfnHook = 0
    tChooseFont.lpTemplateName = Space$(2048)
    'tChooseFont.hInstance As Long
    
    'tChooseFont.lpszStyle As String
    'sFont = Chr$(0) & Space$(20) & Chr$(0)
    'tChooseFont.lpszStyle = sFont
    
 '   tChooseFont.nFontType = Screen.FontCount
    'tChooseFont.MISSING_ALIGNMENT As Integer
    tChooseFont.nSizeMin = mMin
    tChooseFont.nSizeMax = mMax
                    
    
    'tChooseFont.lpLogFont As Long
    
    ' Get a global memory block to hold a copy of tLogFont - exit on failure
    lMemHandle = GlobalAlloc(GHND, lLogFontSize)
    If lMemHandle = 0 Then
        Exit Sub
    End If
    
    ' Lock tLogFont's global memory block - exit on failure
    lLogFontAddress = GlobalLock(lMemHandle)
    If lLogFontAddress = 0 Then
        Exit Sub
    End If
    
    ' Copy tLogFont to the global memory block
    Call CopyMemory(ByVal lLogFontAddress, tLogFont, lLogFontSize)
 
    tChooseFont.lpLogFont = lLogFontAddress
    
    'tChooseFont.iPointSize As Long - init from FontSize property
    tChooseFont.iPointSize = mFont.Size * 10
    
    ' Flags can get reference variable or constant with bit flags
    tChooseFont.Flags = iFlags Or CF_INITTOLOGFONTSTRUCT And Not (CF_APPLY Or CF_ENABLEHOOK Or CF_ENABLETEMPLATE)
    
    If mhDc = 0 Then
        If Not Printer Is Nothing Then
            If ((tChooseFont.Flags And CF_PRINTERFONTS) = CF_PRINTERFONTS) Then tChooseFont.hDC = Printer.hDC
        End If
    Else
        tChooseFont.Flags = tChooseFont.Flags Or CF_PRINTERFONTS
        tChooseFont.hDC = mhDc
    End If
    ' Must have some fonts
    If (tChooseFont.Flags And CF_PRINTERFONTS) = 0 Then tChooseFont.Flags = tChooseFont.Flags Or CF_SCREENFONTS
    ' Color can take initial color, receive chosen color
    If mColor <> 0 Then tChooseFont.Flags = tChooseFont.Flags Or CF_EFFECTS
    ' MinSize can be minimum size accepted
    If (mMin <> 0) Then tChooseFont.Flags = tChooseFont.Flags Or CF_LIMITSIZE
    ' MaxSize can be maximum size accepted
    If (mMax <> 0) Then tChooseFont.Flags = tChooseFont.Flags Or CF_LIMITSIZE
    
    'tChooseFont.Flags = tChooseFont.Flags And Not CF_LIMITSIZE
    
    '***    call the CHOOSEFONT API function
    mApiReturn = CHOOSEFONT(tChooseFont)    'store to APIReturn property
    
    
    '***    handle return from CHOOSEFONT API function
    mCancelled = False
    Select Case mApiReturn
        
        Case 0  'user canceled
        mCancelled = True
        If mCancelError Then
            'generate an error
            Err.Raise 32755, "Cancel Pressed"
            Exit Sub
        End If
        
        Case 1  'user selected a font
            ' Copy global memory block to tLogFont
            Call CopyMemory(tLogFont, ByVal lLogFontAddress, lLogFontSize)
            
            'tLogFont.lfWeight As Long  - store to FontBold property
            If tLogFont.lfWeight >= FW_BOLD Then
                mFont.Bold = True
            Else
                mFont.Bold = False
            End If
                        
            'tLogFont.lfItalic As Byte - store to FontItalic property
            If tLogFont.lfItalic <> 0 Then
                mFont.Italic = True
            Else
                mFont.Italic = False
            End If
            
            'tLogFont.lfUnderline As Byte - store to FontUnderline property
            If tLogFont.lfUnderline <> 0 Then
                mFont.Underline = True
            Else
                mFont.Underline = False
            End If
        
            'tLogFont.lfStrikeOut As Byte - store to FontStrikethru property
            If tLogFont.lfStrikeOut <> 0 Then
                mFont.Strikethrough = True
            Else
                mFont.Strikethrough = False
            End If
            
            mColor = tChooseFont.rgbColors
            
            'tLogFont.lfFaceName(LF_FACESIZE) As Byte - store to FontName property
            FontName = sByteArrayToString(tLogFont.lfFaceName())
            
            'tChooseFont.iPointSize As Long - store to FontSize property
            mFont.Size = CLng(tChooseFont.iPointSize / 10)
        
        Case Else   'an error occured
            'call CommDlgExtendedError
            mExtendedError = CommDlgExtendedError   'store to ExtendedError property
        
    End Select
Exit Sub

ShowFontError:
    Exit Sub
End Sub


Public Property Set Font(nFont As StdFont)
    If nFont Is Nothing Then Exit Property
    Set mFont = CloneFont(nFont)
End Property

Public Property Get Font() As StdFont
    Set Font = CloneFont(mFont)
End Property


Public Sub ShowHelp()
    'run winhelp.exe with the specified help file
    Dim sHelpFileBuff As String
    Dim lData As Long
    
    On Error GoTo ShowHelpError
    
    '***    init Private properties
    mAction = 6  'Action property - ShowHelp
    mApiReturn = 0  'APIReturn property
    mExtendedError = 0  'ExtendedError property

    '***    prepare the buffers and parameters for the API function
    'mHelpFile is a null terminated string
    sHelpFileBuff = mHelpFile & Chr$(0)
    
    'sData is dependent on mHelpCommand
    Select Case mHelpCommand
        Case 0
            lData = 0
        Case Else
            lData = 0
    End Select
    
    '***    call the API function
    mApiReturn = WinHelp(0&, mHelpFile, mHelpCommand, lData)     ' - Store to APIReturn property
    
    Select Case mApiReturn
        
        Case 0  '
            'call CommDlgExtendedError
            mExtendedError = CommDlgExtendedError   ' - store to ExtendedError property
            
        Case Else   '
            'call CommDlgExtendedError
            mExtendedError = CommDlgExtendedError
        
    End Select
        
Exit Sub

ShowHelpError:
    Exit Sub
End Sub


Public Sub ShowOpen(Optional ByVal nFlags As cdeCommonDialogExFileFlagsConstants = -1)
    Dim iHwndOwner As Long
    Dim iFlags As Long
    
    'display the file open dialog box
    iHwndOwner = GetActiveWindowHwnd
    iFlags = mFlags
    If nFlags <> -1 Then
        iFlags = iFlags Or nFlags
    End If
    ShowFileDialog 1, iFlags, iHwndOwner    'Action property - ShowOpen
    
End Sub

Public Sub ShowPrinter(Optional ByVal nFlags As cdeCommonDialogExPrinterFlagsConstants = -1)
    Dim iHwndOwner As Long
    Dim iFlags As Long
    
    'display the print dialog
    Dim tPrintDlg As T_PrintDlg
    
    On Error GoTo ShowPrinterError
    
    iHwndOwner = GetActiveWindowHwnd
    iFlags = mFlags
    If nFlags <> -1 Then
        iFlags = iFlags Or nFlags
    End If
    
    iFlags = iFlags And Not PD_ENABLEPRINTHOOK
    iFlags = iFlags And Not PD_ENABLEPRINTTEMPLATE
    iFlags = iFlags And Not PD_ENABLESETUPHOOK
    iFlags = iFlags And Not PD_ENABLESETUPTEMPLATE
    
    
    mAction = 5  'Action property - ShowPrint
    mApiReturn = 0  'APIReturn property
    mExtendedError = 0  'ExtendedError property

    
    tPrintDlg.lStructSize = Len(tPrintDlg)
    
    tPrintDlg.hwndOwner = iHwndOwner
    
    If (iFlags And cdlbPDReturnDefault) <> 0 Then
        tPrintDlg.hDevMode = 0
        tPrintDlg.hDevNames = 0
        If mDevModePtr <> 0 Then
            mDevModePtr = 0
            mDevNamesPtr = 0
            GlobalUnlock mhDevMode
            GlobalUnlock mhDevNames
            mhDevMode = 0
            mhDevNames = 0
        End If
    Else
        If mDevModePtr <> 0 Then
            tPrintDlg.hDevMode = mhDevMode
            UpdateDevModeWithCurrentSettings
            tPrintDlg.hDevNames = mhDevNames
            GlobalUnlock mhDevMode
            GlobalUnlock mhDevNames
            mDevModePtr = 0
            mDevNamesPtr = 0
        End If
    End If
    
    'hDevMode As Long
    
    'hDevNames As Long
    
    'flags As Long - init from Flags property
    tPrintDlg.Flags = iFlags Or cdlbPDUseDevModeCopies Or cdlbPDReturnDC ' it requires these two flags to return the number of copies properly
    
    'nFromPage As Integer - init from FromPage property
    tPrintDlg.nFromPage = mFromPage
    
    'nToPage As Integer - init from ToPage property
    tPrintDlg.nToPage = mToPage
    
    'nMinPage As Integer - init from Min property
    tPrintDlg.nMinPage = mMin
    
    'nMaxPage As Integer - init from Max property
    tPrintDlg.nMaxPage = mMax
    
    'nCopies As Integer - init from Copies property
    tPrintDlg.nCopies = mCopies
    
    'hInstance As Long
    
    'lCustData As Long
    
    
    '***    Call the PrintDlg API function
    mApiReturn = PrintDlg(tPrintDlg)
    
    '***    handle return from PrintDlg API function
    mCancelled = False
'    If (tPrintDlg.flags And cdlbPDReturnDefault) <> 0 Then
'        mApiReturn = 1
'    End If
    Select Case mApiReturn
        
        Case 0  'user canceled
            If tPrintDlg.hDevMode <> 0 Then
                mDevModePtr = GlobalLock(tPrintDlg.hDevMode)
                mhDevMode = tPrintDlg.hDevMode
                mDevNamesPtr = GlobalLock(tPrintDlg.hDevNames)
                mhDevNames = tPrintDlg.hDevNames
            End If
            
            mCancelled = True
            If mCancelError Then
                'generate an error
                Err.Raise 32755, "Cancel Pressed"
                Exit Sub
            End If
        
        Case 1  'user selected OK
            'nFromPage As Integer - store to FromPage property
            
            mFlags = mFlags And Not (cdlbPDSelection Or cdlbPDPageNums)
            mFlags = mFlags Or (tPrintDlg.Flags And (cdlbPDSelection Or cdlbPDPageNums))

            mFromPage = tPrintDlg.nFromPage
            
            'nToPage As Integer - store to ToPage property
            mToPage = tPrintDlg.nToPage
            
            'nMinPage As Integer - store to Min property
            mMin = tPrintDlg.nMinPage
            
            'nMaxPage As Integer - store to Max property
            mMax = tPrintDlg.nMaxPage
            
            mhDc = tPrintDlg.hDC
            
            If mDevModePtr = 0 Then
                mDevModePtr = GlobalLock(tPrintDlg.hDevMode)
                mhDevMode = tPrintDlg.hDevMode
            Else
                If mhDevMode <> tPrintDlg.hDevMode Then
                    GlobalUnlock mhDevMode
                    mDevModePtr = GlobalLock(tPrintDlg.hDevMode)
                    mhDevMode = tPrintDlg.hDevMode
                End If
            End If
            CopyMemory mDevMode, ByVal mDevModePtr, Len(mDevMode)
            mPaperSize = mDevMode.dmPaperSize
            mPaperBin = mDevMode.dmDefaultSource
            mDuplex = mDevMode.dmDuplex
            mOrientation = mDevMode.dmOrientation
            mPrintQuality = mDevMode.dmPrintQuality
            mColorMode = mDevMode.dmColor
            mDefaultPaperWidth = mDevMode.dmPaperWidth
            mDefaultPaperHeight = mDevMode.dmPaperLength
            mCopies = mDevMode.dmCopies
            If mCopies < 1 Then
                mCopies = 1
            End If
            mCollate = CBool(mDevMode.dmCollate)
            
            If mDevNamesPtr = 0 Then
                mDevNamesPtr = GlobalLock(tPrintDlg.hDevNames)
                mhDevNames = tPrintDlg.hDevNames
            Else
                If mhDevNames <> tPrintDlg.hDevNames Then
                    GlobalUnlock mhDevNames
                    mDevNamesPtr = GlobalLock(tPrintDlg.hDevNames)
                    mhDevNames = tPrintDlg.hDevNames
                End If
            End If
            CopyMemory mDevNames, ByVal mDevNamesPtr, Len(mDevNames)
            mDriverName = GetDevNameString(mDevNamesPtr, mDevNames.wDriverOffset)
            mDeviceName = GetDevNameString(mDevNamesPtr, mDevNames.wDeviceOffset)
            mPort = GetDevNameString(mDevNamesPtr, mDevNames.wOutputOffset)
            
'            Debug.Print mDevMode.dmDeviceName
            
            PutPaperSize
            mPageSet = True
            
        Case Else   'an error occured
            'call CommDlgExtendedError
            mExtendedError = CommDlgExtendedError   'store to ExtendedError property
    
    End Select

Exit Sub

ShowPrinterError:
    
    Exit Sub
    
End Sub


Public Sub ShowSave(Optional ByVal nFlags As cdeCommonDialogExFileFlagsConstants = -1)
    Dim iHwndOwner As Long
    Dim iFlags As Long
    
    'display the file save dialog box
    iHwndOwner = GetActiveWindowHwnd
    iFlags = mFlags
    If nFlags <> -1 Then
        iFlags = iFlags Or nFlags
    End If
    
    ShowFileDialog 2, iFlags, iHwndOwner   'Action property - ShowSave

End Sub


Public Property Get FileName() As String
    FileName = mFileName
End Property

Public Property Let FileName(nValue As String)
    mFileName = nValue
End Property


Public Property Get Filter() As String
    Filter = mFilter
End Property

Public Property Let Filter(nValue As String)
    mFilter = nValue
End Property


Private Function sLeftOfNull(ByVal sIn As String)
    'returns the part of sIn preceding Chr$(0)
    Dim lNullPos As Long
    
    'init output
    sLeftOfNull = sIn
    
    'get position of first Chr$(0) in sIn
    lNullPos = InStr(sIn, Chr$(0))
    
    'return part of sIn to Left of first Chr$(0) if found
    If lNullPos > 0 Then
        sLeftOfNull = Mid$(sIn, 1, lNullPos - 1)
    End If
    
End Function

Private Function sLeftOfLastNull(ByVal sIn As String)
    'returns the part of sIn preceding Chr$(0)
    Dim iNullPos As Long
    Dim iChr As String
    
    'init output
    sLeftOfLastNull = sIn
    
    'get position of first Chr$(0) in sIn
    iNullPos = InStrRev(sIn, Chr$(0))
    
    'return part of sIn to Left of first Chr$(0) if found
    If iNullPos > 0 Then
        iChr = Mid$(sIn, iNullPos, 1)
        Do Until (iChr <> Chr$(0)) And (iChr <> " ")
            iNullPos = iNullPos - 1
            If iNullPos = 0 Then
                iNullPos = InStrRev(sIn, Chr$(0))
                Exit Do
            End If
            iChr = Mid$(sIn, iNullPos, 1)
        Loop
        iNullPos = iNullPos + 1
        sLeftOfLastNull = Mid$(sIn, 1, iNullPos - 1)
    End If
    
End Function


Private Function sAPIFilter(sIn)
    'prepares sIn for use as a filter string in API common dialog functions
    Dim lChrNdx As Long
    Dim sOneChr As String
    Dim sOutStr As String
    
    'convert any | characters to nulls
    For lChrNdx = 1 To Len(sIn)
        sOneChr = Mid$(sIn, lChrNdx, 1)
        If sOneChr = "|" Then
            sOutStr = sOutStr & Chr$(0)
        Else
            sOutStr = sOutStr & sOneChr
        End If
    Next
    
    'add a null to the end
    sOutStr = sOutStr & Chr$(0)
    
    'return sOutStr
    sAPIFilter = sOutStr
    
End Function

Public Property Get FilterIndex() As Integer
    FilterIndex = mFilterIndex
End Property

Public Property Let FilterIndex(nValue As Integer)
    mFilterIndex = nValue
End Property

Public Property Get CancelError() As Boolean
    CancelError = mCancelError
End Property

Public Property Let CancelError(nValue As Boolean)
    mCancelError = nValue
End Property

Public Property Get Color() As Long
    Color = mColor
End Property

Public Property Let Color(nValue As Long)
    mColor = nValue
End Property

Public Property Get Copies() As Long
    Copies = mCopies
End Property

Public Property Let Copies(nValue As Long)
    mCopies = nValue
End Property

Public Property Get Collate() As Boolean
    Collate = mCollate
End Property

Public Property Let Collate(nValue As Boolean)
    mCollate = nValue
End Property

Public Property Get DefaultExt() As String
    DefaultExt = mDefaultExt
End Property

Public Property Let DefaultExt(nValue As String)
    mDefaultExt = nValue
End Property

Public Property Get DialogTitle() As String
    DialogTitle = mDialogTitle
End Property

Public Property Let DialogTitle(nValue As String)
    mDialogTitle = nValue
End Property

Public Property Get Flags() As CommonDialogsFlags
    Flags = mFlags
End Property

Public Property Let Flags(nValue As CommonDialogsFlags)
    mFlags = nValue
End Property

Public Property Get FontBold() As Boolean
    FontBold = mFont.Bold
End Property

Public Property Let FontBold(nValue As Boolean)
    mFont.Bold = nValue
End Property

Public Property Get FontItalic() As Boolean
    FontItalic = mFont.Italic
End Property

Public Property Let FontItalic(nValue As Boolean)
    mFont.Italic = nValue
End Property

Public Property Get FontName() As String
    FontName = mFont.Name
End Property

Public Property Let FontName(nValue As String)
    mFont.Name = nValue
End Property

Public Property Get FontSize() As Long
    FontSize = mFont.Size
End Property

Public Property Let FontSize(nValue As Long)
    mFont.Size = nValue
End Property

Public Property Get FontStrikeThru() As Boolean
    FontStrikeThru = mFont.Strikethrough
End Property

Public Property Let FontStrikeThru(nValue As Boolean)
    mFont.Strikethrough = nValue
End Property

Public Property Get FontUnderLine() As Boolean
    FontUnderLine = mFont.Underline
End Property

Public Property Let FontUnderLine(nValue As Boolean)
    mFont.Underline = nValue
End Property

Public Property Get FromPage() As Long
Attribute FromPage.VB_MemberFlags = "40"
    FromPage = mFromPage
End Property

Public Property Let FromPage(nValue As Long)
    mFromPage = nValue
End Property

Public Property Get hDC() As Long
    EnsurePageSet
    hDC = mhDc
End Property

Public Property Let hDC(nValue As Long)
    mhDc = nValue
End Property


Public Property Get HelpCommand() As Long
    HelpCommand = mHelpCommand
End Property

Public Property Let HelpCommand(nValue As Long)
    mHelpCommand = nValue
End Property

Public Property Get HelpContext() As Long
    HelpContext = mHelpContext
End Property

Public Property Let HelpContext(nValue As Long)
    mHelpContext = nValue
End Property

Public Property Get HelpFile() As String
    HelpFile = mHelpFile
End Property

Public Property Let HelpFile(nValue As String)
    mHelpFile = nValue
End Property

Public Property Get HelpKey() As String
    HelpKey = mHelpKey
End Property

Public Property Let HelpKey(nValue As String)
    mHelpKey = nValue
End Property

Public Property Get InitDir() As String
    InitDir = mInitDir
End Property

Public Property Let InitDir(nValue As String)
    mInitDir = Trim$(nValue)
End Property

Public Property Get Max() As Long
    Max = mMax
End Property

Public Property Let Max(nValue As Long)
    mMax = nValue
    If mMin > mMax Then
        mMin = mMax
    End If
End Property

Public Property Get MaxFileSize() As Long
    MaxFileSize = mMaxFileSize
End Property

Public Property Let MaxFileSize(nValue As Long)
    mMaxFileSize = nValue
End Property

Public Property Get Min() As Long
    Min = mMin
End Property

Public Property Let Min(nValue As Long)
    mMin = nValue
    If mMax < mMin Then
        mMax = mMin
    End If
End Property

Public Property Get Object() As Object
    Set Object = Me
End Property

'Public Property Get PrinterDefault() As Integer
'    PrinterDefault = mPrinterDefault
'End Property
'
'Public Property Let PrinterDefault(nValue As Integer)
'    mPrinterDefault = nValue
'End Property

Public Property Get ToPage() As Long
    ToPage = mToPage
End Property

Public Property Let ToPage(nValue As Long)
    mToPage = nValue
End Property

Public Property Get FileTitle() As String
    FileTitle = mFileTitle
End Property

Public Property Let FileTitle(nValue As String)
    mFileTitle = nValue
End Property

'Private Property Get APIReturn() As Long
'    APIReturn = mApiReturn
'End Property
'
'Private Property Get ExtendedError() As Long
'    ExtendedError = mExtendedError
'End Property


Private Function sByteArrayToString(abBytes() As Byte) As String
    'return a string from a byte array
    Dim lBytePoint As Long
    Dim lByteVal As Long
    Dim sOut As String
    
    'init array pointer
    lBytePoint = LBound(abBytes)
    
    'fill sOut with characters in array
    While lBytePoint <= UBound(abBytes)
        
        lByteVal = abBytes(lBytePoint)
        
        'return sOut and stop if Chr$(0) is encountered
        If lByteVal = 0 Then
            sByteArrayToString = sOut
            Exit Function
        Else
            sOut = sOut & Chr$(lByteVal)
        End If
        
        lBytePoint = lBytePoint + 1
    
    Wend
    
    'return sOut if Chr$(0) wasn't encountered
    sByteArrayToString = sOut
    
End Function
Private Sub ShowFileDialog(ByVal mAction As Integer, nFlags As Long, Optional ByVal nHwndOwner As Long)
    
    'display the file dialog for ShowOpen or ShowSave
    
    Dim tOpenFile As T_OpenFilename
    Dim lMaxSize As Long
    Dim sFileNameBuff As String
    Dim sFileTitleBuff As String
    Dim iPos As Long
    
    On Error GoTo ShowFileDialogError
    
    mCancelled = False
    
    '***    init property buffers
    
    mAction = mAction  'Action property
    mApiReturn = 0  'APIReturn property
    mExtendedError = 0  'ExtendedError property
        
    
    '***    prepare tOpenFile data
    
    'tOpenFile.lStructSize As Long
    tOpenFile.lStructSize = Len(tOpenFile)
    
    'tOpenFile.hWndOwner As Long - init from hdc property
    tOpenFile.hwndOwner = nHwndOwner
    
    'tOpenFile.lpstrFilter As String - init from Filter property
    tOpenFile.lpstrFilter = sAPIFilter(mFilter)
        
    'tOpenFile.mFilterIndex As Long - init from FilterIndex property
    tOpenFile.mFilterIndex = mFilterIndex
    
    'tOpenFile.lpstrFile As String
        'determine size of buffer from MaxFileSize property
        If mMaxFileSize > 0 Then
            lMaxSize = mMaxFileSize
        Else
            lMaxSize = 255
        End If
    
    'tOpenFile.lpstrFile As Long - init from FileName property
        'prepare sFileNameBuff
        sFileNameBuff = mFileName
        If InStr(sFileNameBuff, "//") > 0 Then
            Do Until InStr(sFileNameBuff, "//") = 0
                sFileNameBuff = Replace(sFileNameBuff, "//", "/")
            Loop
        End If
        If InStr(sFileNameBuff, "\\") > 0 Then
            Do Until InStr(sFileNameBuff, "\\") = 0
                sFileNameBuff = Replace(sFileNameBuff, "\\", "\")
            Loop
        End If
        If InStr(sFileNameBuff, "\\") > 0 Then
            Do Until InStr(sFileNameBuff, "\\") = 0
                sFileNameBuff = Replace(sFileNameBuff, "\\", "\")
            Loop
        End If
        iPos = InStr(sFileNameBuff, "\")
        If InStr(sFileNameBuff, "/") < iPos Then
            iPos = InStr(sFileNameBuff, "\")
        End If
        If iPos = 0 Then
            iPos = Len(sFileNameBuff)
        End If
        
        If InStr(iPos, sFileNameBuff, ":") > 0 Then
            sFileNameBuff = Replace(sFileNameBuff, ":", "")
        End If
        If InStr(iPos, sFileNameBuff, "?") > 0 Then
            sFileNameBuff = Replace(sFileNameBuff, "?", "")
        End If
        If InStr(iPos, sFileNameBuff, "*") > 0 Then
            sFileNameBuff = Replace(sFileNameBuff, "*", "")
        End If
        If InStr(iPos, sFileNameBuff, """") > 0 Then
            sFileNameBuff = Replace(sFileNameBuff, """", "")
        End If
        If InStr(iPos, sFileNameBuff, ">") > 0 Then
            sFileNameBuff = Replace(sFileNameBuff, ">", "")
        End If
        If InStr(iPos, sFileNameBuff, "<") > 0 Then
            sFileNameBuff = Replace(sFileNameBuff, "<", "")
        End If
        If InStr(iPos, sFileNameBuff, "|") > 0 Then
            sFileNameBuff = Replace(sFileNameBuff, "|", "")
        End If
        
        'pad with spaces
        If (nFlags And cdlbOFNAllowMultiselect) <> 0 Then
            sFileNameBuff = sFileNameBuff & Space$(100000 - Len(sFileNameBuff) - 1)
            sFileNameBuff = Mid$(sFileNameBuff, 1, 99999)
        Else
            While Len(sFileNameBuff) < lMaxSize - 1
                sFileNameBuff = sFileNameBuff & " "
            Wend
        'Trim$ to length of mMaxFileSize - 1
            sFileNameBuff = Mid$(sFileNameBuff, 1, lMaxSize - 1)
        End If
        'null terminate
        sFileNameBuff = sFileNameBuff & Chr$(0)
    tOpenFile.lpstrFile = sFileNameBuff
    
    'nMaxFile As Long - init from MaxFileSize property
    'If mMaxFileSize <> 255 Then  'default is 255
    If (nFlags And cdlbOFNAllowMultiselect) <> 0 Then
        tOpenFile.nMaxFile = 100000
    Else
        tOpenFile.nMaxFile = lMaxSize
    End If
    'End If
            
    'lpstrFileTitle As String - init from FileTitle property
        'prepare sFileTitleBuff
        sFileTitleBuff = mFileTitle
        'pad with spaces
        While Len(sFileTitleBuff) < lMaxSize - 1
            sFileTitleBuff = sFileTitleBuff & " "
        Wend
        'Trim$ to length of mMaxFileSize - 1
        sFileTitleBuff = Mid$(sFileTitleBuff, 1, lMaxSize - 1)
        'null terminate
        sFileTitleBuff = sFileTitleBuff & Chr$(0)
    tOpenFile.lpstrFileTitle = sFileTitleBuff
        
    'tOpenFile.lpstrInitialDir As String - init from InitDir property
    tOpenFile.lpstrInitialDir = mInitDir
    
    'tOpenFile.lpstrTitle As String - init from DialogTitle property
    tOpenFile.lpstrTitle = mDialogTitle
    
    'tOpenFile.flags As Long - init from Flags property
    tOpenFile.Flags = nFlags Or cdlbOFNNoChangeDir
        
    'tOpenFile.lpstrDefExt As String - init from DefaultExt property
    tOpenFile.lpstrDefExt = mDefaultExt
    
    
    '***    call the GetOpenFileName API function
    Select Case mAction
        Case 1  'ShowOpen
            mApiReturn = GetOpenFileName(tOpenFile)
        Case 2  'ShowSave
            mApiReturn = GetSaveFileName(tOpenFile)
        Case Else   'unknown action
            Exit Sub
    End Select
    
    
    '***    handle return from GetOpenFileName API function
    mCancelled = False
    Select Case mApiReturn
        
        Case 0  'user canceled
            mCancelled = True
            mFileName = ""
            mFileTitle = ""
            If mCancelError Then
                'generate an error
                On Error Resume Next
                Err.Raise 32755, "Cancel Pressed"
                Exit Sub
            End If
        
        Case 1  'user selected or entered a file
            'mFileName gets part of tOpenFile.lpstrFile to the Left of first Chr$(0)
            If (tOpenFile.Flags And cdlbOFNAllowMultiselect) <> 0 Then
                mFileName = sLeftOfLastNull(tOpenFile.lpstrFile)
            Else
                mFileName = sLeftOfNull(tOpenFile.lpstrFile)
            End If
            mFileTitle = sLeftOfNull(tOpenFile.lpstrFileTitle)
        
        Case Else   'an error occured
            'call CommDlgExtendedError
            mExtendedError = CommDlgExtendedError
        
    End Select
    

Exit Sub

ShowFileDialogError:
    
    Exit Sub

End Sub


Public Property Get Canceled() As Boolean
    Canceled = mCancelled
End Property

Private Function TranslateAColor(ByVal clr As OLE_COLOR, _
                        Optional hPal As Long = 0) As Long
   If OleTranslateColor(clr, hPal, TranslateAColor) Then
      TranslateAColor = CLR_INVALID
   End If
End Function


Private Sub Class_Initialize()
    mMaxFileSize = 255
    mUnits = vbMillimeters
    mAmbientUserMode = True
    mCopies = 1
    mLeftMargin = cLeftMarginDefault
    mRightMargin = cRightMarginDefault
    mTopMargin = cTopMarginDefault
    mBottomMargin = cBottomMarginDefault
    mAutoSaveCustomColors = True
    InitCustomColors
    
    Set mFont = New StdFont
End Sub

Public Property Let Action(nValue As Integer)
    Select Case nValue
        Case 0
            '
        Case 1
            ShowOpen
        Case 2
            ShowSave
        Case 3
            
        Case 4
            
        Case 5
            ShowPrinter
        Case 6
            ShowHelp
    End Select
End Property


Public Property Get Orientation() As cdePageOrientationConstants
    EnsurePageSet
    Orientation = mOrientation
End Property

Public Property Let Orientation(nValue As cdePageOrientationConstants)
    If (nValue = 0) And mAmbientUserMode Then Exit Property 'If (nValue = 0) And mAmbientUserMode Then Err.Raise 1769, , "At run time this property must have a value other than 0."
    EnsurePageSet
    mOrientation = nValue
End Property


Public Property Get PaperSize() As cdePaperSizeConstants
    EnsurePageSet
    PaperSize = mPaperSize
End Property

Public Property Let PaperSize(nValue As cdePaperSizeConstants)
    If (nValue = 0) And mAmbientUserMode Then Exit Property 'If (nValue = 0) And mAmbientUserMode Then Err.Raise 1769, , "At run time this property must have a value other than 0."
    EnsurePageSet
    mPaperSize = nValue
End Property


Public Property Get PrintQuality() As cdePrintQualityConstants
    EnsurePageSet
    PrintQuality = mPrintQuality
End Property

Public Property Let PrintQuality(nValue As cdePrintQualityConstants)
    If (nValue = 0) And mAmbientUserMode Then Exit Property 'If (nValue = 0) And mAmbientUserMode Then Err.Raise 1769, , "At run time this property must have a value other than 0."
    EnsurePageSet
    mPrintQuality = nValue
End Property


Public Property Get ColorMode() As cdeColorModeConstants
    EnsurePageSet
    ColorMode = mColorMode
End Property

Public Property Let ColorMode(nValue As cdeColorModeConstants)
    If (nValue = 0) And mAmbientUserMode Then Exit Property 'If (nValue = 0) And mAmbientUserMode Then Err.Raise 1769, , "At run time this property must have a value other than 0."
    EnsurePageSet
    mColorMode = nValue
End Property


Public Property Get DriverName() As String
    EnsurePageSet
    DriverName = mDriverName
End Property


Public Property Get Duplex() As cdeDuplexConstants
    EnsurePageSet
    Duplex = mDuplex
End Property

Public Property Let Duplex(nValue As cdeDuplexConstants)
    If (nValue = 0) And mAmbientUserMode Then Exit Property 'If (nValue = 0) And mAmbientUserMode Then Err.Raise 1769, , "At run time this property must have a value other than 0."
    EnsurePageSet
    mDuplex = nValue
End Property


Public Property Get PaperBin() As cdePaperBinConstants
    EnsurePageSet
    PaperBin = mPaperBin
End Property

Public Property Let PaperBin(nValue As cdePaperBinConstants)
    If (nValue = 0) And mAmbientUserMode Then Exit Property 'If (nValue = 0) And mAmbientUserMode Then Err.Raise 1769, , "At run time this property must have a value other than 0."
    EnsurePageSet
    mPaperBin = nValue
End Property


Public Property Get Port() As String
    EnsurePageSet
    Port = mPort
End Property


Public Property Get DeviceName() As String
    EnsurePageSet
    DeviceName = mDeviceName
End Property


Public Property Get PaperWidth() As Single
    EnsurePageSet
    If Units = vbInches Then
        PaperWidth = mPaperWidth / 254
    Else
        PaperWidth = mPaperWidth / 10
    End If
End Property


Public Property Get PaperHeight() As Single
    EnsurePageSet
    If Units = vbInches Then
        PaperHeight = mPaperHeight / 254
    Else
        PaperHeight = mPaperHeight / 10
    End If
End Property


Public Property Get LeftMargin() As Single
    EnsurePageSet
    LeftMargin = mLeftMargin
End Property

Public Property Let LeftMargin(nValue As Single)
    EnsurePageSet
    mLeftMargin = nValue
'    'mMarginSet = True
    If mLeftMargin < mMinLeftMargin Then
        mLeftMargin = mMinLeftMargin
    End If
End Property


Public Property Get MinLeftMargin() As Single
    EnsurePageSet
    MinLeftMargin = mMinLeftMargin
End Property

Public Property Let MinLeftMargin(nValue As Single)
    EnsurePageSet
    mMinLeftMargin = nValue
    'mMarginSet = True
    If mLeftMargin < mMinLeftMargin Then
        mLeftMargin = mMinLeftMargin
    End If
End Property


Public Property Get RightMargin() As Single
    EnsurePageSet
    RightMargin = mRightMargin
End Property

Public Property Let RightMargin(nValue As Single)
    EnsurePageSet
    mRightMargin = nValue
    'mMarginSet = True
    If mRightMargin < mMinRightMargin Then
        mRightMargin = mMinRightMargin
    End If
End Property


Public Property Get MinRightMargin() As Single
    EnsurePageSet
    MinRightMargin = mMinRightMargin
End Property

Public Property Let MinRightMargin(nValue As Single)
    EnsurePageSet
    mMinRightMargin = nValue
    'mMarginSet = True
    If mRightMargin < mMinRightMargin Then
        mRightMargin = mMinRightMargin
    End If
End Property


Public Property Get TopMargin() As Single
    EnsurePageSet
    TopMargin = mTopMargin
End Property

Public Property Let TopMargin(nValue As Single)
    EnsurePageSet
    mTopMargin = nValue
    'mMarginSet = True
    If mTopMargin < mMinTopMargin Then
        mTopMargin = mMinTopMargin
    End If
End Property


Public Property Get MinTopMargin() As Single
    EnsurePageSet
    MinTopMargin = mMinTopMargin
End Property

Public Property Let MinTopMargin(nValue As Single)
    EnsurePageSet
    mMinTopMargin = nValue
    'mMarginSet = True
    If mTopMargin < mMinTopMargin Then
        mTopMargin = mMinTopMargin
    End If
End Property


Public Property Get BottomMargin() As Single
    EnsurePageSet
    BottomMargin = mBottomMargin
End Property

Public Property Let BottomMargin(nValue As Single)
    EnsurePageSet
    mBottomMargin = nValue
    'mMarginSet = True
    If mBottomMargin < mMinBottomMargin Then
        mBottomMargin = mMinBottomMargin
    End If
End Property


Public Property Get MinBottomMargin() As Single
    EnsurePageSet
    MinBottomMargin = mMinBottomMargin
End Property

Public Property Let MinBottomMargin(nValue As Single)
    EnsurePageSet
    mMinBottomMargin = nValue
    'mMarginSet = True
    If mBottomMargin < mMinBottomMargin Then
        mBottomMargin = mMinBottomMargin
    End If
End Property


Public Property Get Units() As cdeUnits
    Units = mUnits
End Property

Public Property Let Units(nValue As cdeUnits)
    If (nValue <> vbInches) And (nValue <> vbMillimeters) Then Exit Property
    If nValue <> mUnits Then
        mUnits = nValue
        ConvertMarginValues
    End If
End Property


Public Property Get UnitsForUser() As cdeUnitsForUser
    UnitsForUser = mUnitsForUser
End Property

Public Property Let UnitsForUser(nValue As cdeUnitsForUser)
    If (nValue <> cdeMUInches) And (nValue <> cdeMUMillimeters) And (nValue <> cdeMUUserLocale) Then Exit Property
    If nValue <> mUnitsForUser Then
        mUnitsForUser = nValue
    End If
End Property


Public Sub ShowPageSetup(Optional ByVal nFlags As cdeCommonDialogExPageSetupFlagsConstants = -1)
    Dim iHwndOwner As Long
    Dim iPsd As T_PAGESETUPDLG
    Dim iUnitsMultiplier As Single
    Dim iFlags As Long
    Dim iUnits_Ant  As Long
    
    On Error GoTo ShowPageSetupError
    
    iHwndOwner = GetActiveWindowHwnd
    iFlags = mFlags
    If nFlags <> -1 Then
        iFlags = iFlags Or nFlags
    End If
    
    iUnits_Ant = mUnits
    If mUnits <> vbMillimeters Then
        Units = vbMillimeters
    End If
    
    iPsd.lStructSize = Len(iPsd)
    
    iPsd.Flags = iFlags And Not (PSD_ENABLEPAGEPAINTHOOK Or PSD_ENABLEPAGESETUPHOOK Or PSD_ENABLEPAGESETUPTEMPLATE)
    
    Select Case mUnitsForUser
        Case cdeMUInches
            iPsd.Flags = iPsd.Flags Or PSD_INTHOUSANDTHSOFINCHES
        Case cdeMUMillimeters
            iPsd.Flags = iPsd.Flags Or PSD_INHUNDREDTHSOFMILLIMETERS
        Case Else
            If GetLocaleMeasureSystem = 1 Then ' EEUU system, Inches
                iPsd.Flags = iPsd.Flags Or PSD_INTHOUSANDTHSOFINCHES
            Else
                iPsd.Flags = iPsd.Flags Or PSD_INHUNDREDTHSOFMILLIMETERS
            End If
    End Select
    
    If (iPsd.Flags And PSD_INTHOUSANDTHSOFINCHES) <> 0 Then
        iUnitsMultiplier = 1000
    Else
        iUnitsMultiplier = 100
    End If
    Select Case True
        Case (Units = vbInches) And ((iPsd.Flags And PSD_INTHOUSANDTHSOFINCHES) = 0)
            iUnitsMultiplier = iUnitsMultiplier * 25.4
        Case (Units = vbInches) And ((iPsd.Flags And PSD_INTHOUSANDTHSOFINCHES) <> 0)
            ' OK
        Case (Units = vbMillimeters) And ((iPsd.Flags And PSD_INTHOUSANDTHSOFINCHES) = 0)
            ' OK
        Case (Units = vbMillimeters) And ((iPsd.Flags And PSD_INTHOUSANDTHSOFINCHES) <> 0)
            iUnitsMultiplier = iUnitsMultiplier / 25.4
    End Select
    
    
'    If mMarginSet Then
    iPsd.Flags = iPsd.Flags Or PSD_MARGINS
'    End If
    If (mMinTopMargin <> 0) Or (mMinLeftMargin <> 0) Or (mMinBottomMargin <> 0) Or (mMinRightMargin <> 0) Then
        iPsd.Flags = iPsd.Flags Or cdlbPSMinMargins
    End If
    
    iPsd.hwndOwner = iHwndOwner
    iPsd.rtMargin.Top = mTopMargin * iUnitsMultiplier
    iPsd.rtMargin.Left = mLeftMargin * iUnitsMultiplier
    iPsd.rtMargin.Bottom = mBottomMargin * iUnitsMultiplier
    iPsd.rtMargin.Right = mRightMargin * iUnitsMultiplier
    iPsd.rtMinMargin.Top = mMinTopMargin * iUnitsMultiplier
    iPsd.rtMinMargin.Left = mMinLeftMargin * iUnitsMultiplier
    iPsd.rtMinMargin.Bottom = mMinBottomMargin * iUnitsMultiplier
    iPsd.rtMinMargin.Right = mMinRightMargin * iUnitsMultiplier
    
    If mDevModePtr <> 0 Then
        iPsd.hDevMode = mhDevMode
        UpdateDevModeWithCurrentSettings
        iPsd.hDevNames = mhDevNames
        GlobalUnlock mhDevMode
        GlobalUnlock mhDevNames
        mDevModePtr = 0
        mDevNamesPtr = 0
    End If
    
    mApiReturn = PageSetupDlg(iPsd)
    mCancelled = False
    
    Select Case mApiReturn
        
        Case 0  'user canceled
        
        If iPsd.hDevMode <> 0 Then
            mDevModePtr = GlobalLock(iPsd.hDevMode)
            mhDevMode = iPsd.hDevMode
            mDevNamesPtr = GlobalLock(iPsd.hDevNames)
            mhDevNames = iPsd.hDevNames
        End If
        
        mCancelled = True
        If mCancelError Then
            'generate an error
            On Error GoTo 0
            Err.Raise 32755, "Cancel Pressed"
            Exit Sub
        End If
        
        Case 1  'user clicked OK
            If (iPsd.Flags And PSD_INTHOUSANDTHSOFINCHES) <> 0 Then
                iUnitsMultiplier = 1000
            Else
                iUnitsMultiplier = 100
            End If
            Select Case True
                Case (Units = vbInches) And ((iPsd.Flags And PSD_INTHOUSANDTHSOFINCHES) = 0)
                    iUnitsMultiplier = iUnitsMultiplier * 25.4
                Case (Units = vbInches) And ((iPsd.Flags And PSD_INTHOUSANDTHSOFINCHES) <> 0)
                    ' OK
                Case (Units = vbMillimeters) And ((iPsd.Flags And PSD_INTHOUSANDTHSOFINCHES) = 0)
                    ' OK
                Case (Units = vbMillimeters) And ((iPsd.Flags And PSD_INTHOUSANDTHSOFINCHES) <> 0)
                    iUnitsMultiplier = iUnitsMultiplier / 25.4
            End Select
            
            mTopMargin = Int((iPsd.rtMargin.Top + 0.49) / iUnitsMultiplier * 100) / 100
            mLeftMargin = Int((iPsd.rtMargin.Left + 0.49) / iUnitsMultiplier * 100) / 100
            mBottomMargin = Int((iPsd.rtMargin.Bottom + 0.49) / iUnitsMultiplier * 100) / 100
            mRightMargin = Int((iPsd.rtMargin.Right + 0.49) / iUnitsMultiplier * 100) / 100
            mMinTopMargin = Int((iPsd.rtMinMargin.Top + 0.49) / iUnitsMultiplier * 100) / 100
            mMinLeftMargin = Int((iPsd.rtMinMargin.Left + 0.49) / iUnitsMultiplier * 100) / 100
            mMinBottomMargin = Int((iPsd.rtMinMargin.Bottom + 0.49) / iUnitsMultiplier * 100) / 100
            mMinRightMargin = Int((iPsd.rtMinMargin.Right + 0.49) / iUnitsMultiplier * 100) / 100
            
            If mDevModePtr = 0 Then
                mDevModePtr = GlobalLock(iPsd.hDevMode)
                mhDevMode = iPsd.hDevMode
            End If
            CopyMemory mDevMode, ByVal mDevModePtr, Len(mDevMode)
            mPaperSize = mDevMode.dmPaperSize
            mPaperBin = mDevMode.dmDefaultSource
            mDuplex = mDevMode.dmDuplex
            mOrientation = mDevMode.dmOrientation
            mPrintQuality = mDevMode.dmPrintQuality
            mColorMode = mDevMode.dmColor
            mDefaultPaperWidth = mDevMode.dmPaperWidth
            mDefaultPaperHeight = mDevMode.dmPaperLength
            PutPaperSize
            
            If mDevNamesPtr = 0 Then
                mDevNamesPtr = GlobalLock(iPsd.hDevNames)
                mhDevNames = iPsd.hDevNames
            End If
            CopyMemory mDevNames, ByVal mDevNamesPtr, Len(mDevNames)
            mDriverName = GetDevNameString(mDevNamesPtr, mDevNames.wDriverOffset)
            mDeviceName = GetDevNameString(mDevNamesPtr, mDevNames.wDeviceOffset)
            mPort = GetDevNameString(mDevNamesPtr, mDevNames.wOutputOffset)
            
            mPageSet = True
        Case Else   'an error occured
            'call CommDlgExtendedError
            mExtendedError = CommDlgExtendedError
        
    End Select
    
    If mUnits <> iUnits_Ant Then
        Units = iUnits_Ant
    End If
    
'    If Not mCancelled Then
'        If mhDc <> 0 Then
'            Dim iDevMode As DEVMODE
'
'            CopyMemory iDevMode, ByVal mDevModePtr, Len(iDevMode)
'            ShowPrinter cdlbPDReturnDefault Or cdePDReturnDC
'            ResetDC mhDc, iDevMode
'            mPaperSize = iDevMode.dmPaperSize
'            mPaperBin = iDevMode.dmDefaultSource
'            mDuplex = iDevMode.dmDuplex
'            mOrientation = iDevMode.dmOrientation
'            mPrintQuality = iDevMode.dmPrintQuality
'            mColorMode = iDevMode.dmColor
'            mDefaultPaperWidth = iDevMode.dmPaperWidth
'            mDefaultPaperHeight = iDevMode.dmPaperLength
'            PutPaperSize
'            UpdateDevModeWithCurrentSettings
'        End If
'    End If
    
ShowPageSetupError:
End Sub

Private Function GetDevNameString( _
      ByVal ptrDevNames As Long, _
      ByVal ptrOffset As Long _
   )
   Dim Ptr As Long
   Dim lSize As Long
   Dim B() As Byte
      
   Ptr = UnsignedAdd(ptrDevNames, ptrOffset)
   lSize = lstrlenPtr(Ptr)
   If (lSize > 0) Then
      ReDim B(0 To lSize - 1) As Byte
      CopyMemory B(0), ByVal Ptr, lSize
      GetDevNameString = StrConv(B, vbUnicode)
   End If
End Function

Private Function UnsignedAdd(Start As Long, Incr As Long) As Long
' This function is useful when doing pointer arithmetic,
' but note it only works for positive values of Incr

   If Start And &H80000000 Then 'Start < 0
      UnsignedAdd = Start + Incr
   ElseIf (Start Or &H80000000) < -Incr Then
      UnsignedAdd = Start + Incr
   Else
      UnsignedAdd = (Start + &H80000000) + (Incr + &H80000000)
   End If
   
End Function

Private Sub Class_Terminate()
    If mDevModePtr <> 0 Then
        GlobalUnlock mhDevMode
        GlobalUnlock mhDevNames
    End If
End Sub

Private Function GetLocaleMeasureSystem() As Long
    Dim Buffer As String * 100
    Dim nullpos&
    Dim dl&
    
    dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_IMEASURE, Buffer, 99)
    nullpos& = InStr(Buffer, Chr$(0))
    GetLocaleMeasureSystem = Val(Left$(Buffer, nullpos - 1))
End Function

Private Sub EnsurePageSet()
    If mAmbientUserMode Then
        If Not mPageSet Then
            ShowPrinter cdlbPDReturnDefault Or cdePDReturnDC
        End If
    End If
End Sub


Public Property Get FolderName() As String
    FolderName = mFileName
End Property

Public Property Let FolderName(nValue As String)
    mFileName = nValue
End Property


Public Property Get FolderTitle() As String
    FolderTitle = mFileTitle
End Property

Public Property Let FolderTitle(nValue As String)
    mFileTitle = nValue
End Property


Public Property Get RootFolder() As String
    RootFolder = mRootFolder
End Property

Public Property Let RootFolder(nValue As String)
    mRootFolder = Trim$(nValue)
End Property


Public Property Get DialogHeader() As String
    DialogHeader = mDialogHeader
End Property

Public Property Let DialogHeader(nValue As String)
    mDialogHeader = nValue
End Property


'Private Function GetFileFolderPath(nFileFullPath As String) As String
'    Dim iFolderPath As String
'
'    SeparatePathAndFileName nFileFullPath, iFolderPath
'    GetFileFolderPath = iFolderPath
'    AddDirSep GetFileFolderPath
'End Function

'Private Function GetFolderName(nFullFolderPath As String) As String
'    Dim iPath As String
'    Dim iFolder As String
'
'    SeparatePathAndFileName nFullFolderPath, iPath, iFolder
'    GetFolderName = iFolder
'End Function

Private Function GetAddressofFunction(Add As Long) As Long
    GetAddressofFunction = Add
End Function


Private Function IsPrinter(nName As String) As Boolean
    Dim iPrn As Printer
    
    For Each iPrn In Printers
        If LCase$(iPrn.DeviceName) = LCase$(nName) Then
            IsPrinter = True
            Exit For
        End If
    Next iPrn
    
End Function

Public Property Let AmbientUserMode(nValue As Boolean)
Attribute AmbientUserMode.VB_MemberFlags = "40"
    mAmbientUserMode = nValue
End Property

Private Sub UpdateDevModeWithCurrentSettings()
    
'    Debug.Print mDevMode.dmDeviceName, mDeviceName & Chr(0)
    mDevMode.dmPaperSize = mPaperSize
    mDevMode.dmDefaultSource = mPaperBin
    mDevMode.dmDuplex = mDuplex
    mDevMode.dmOrientation = mOrientation
    mDevMode.dmPrintQuality = mPrintQuality
    mDevMode.dmColor = mColorMode
    mDevMode.dmCopies = mCopies
    mDevMode.dmCollate = Abs(CLng(mCollate))
    
    CopyMemory ByVal mDevModePtr, mDevMode, Len(mDevMode)
    
End Sub

Private Sub ConvertMarginValues()
    Dim iMultiplier As Single
    
    If mUnits = vbMillimeters Then
        iMultiplier = 25.4
    Else
        iMultiplier = 1 / 25.4
    End If
    mLeftMargin = mLeftMargin * iMultiplier
    mRightMargin = mRightMargin * iMultiplier
    mTopMargin = mTopMargin * iMultiplier
    mBottomMargin = mBottomMargin * iMultiplier
    mMinLeftMargin = mMinLeftMargin * iMultiplier
    mMinRightMargin = mMinRightMargin * iMultiplier
    mMinTopMargin = mMinTopMargin * iMultiplier
    mMinBottomMargin = mMinBottomMargin * iMultiplier
    
End Sub


Private Sub PutPaperSize()
    Dim iPs As POINTAPI
    
    iPs = GetPaperSize(mPaperSize)
    
    If iPs.X = 0 Then
        mPaperWidth = mDefaultPaperWidth
        mPaperHeight = mDefaultPaperHeight
    Else
        mPaperWidth = iPs.X
        mPaperHeight = iPs.Y
    End If
    
End Sub

Private Function GetPaperSize(nPaperSizeNumber As Long) As POINTAPI
    Dim ret As Long
    Dim iPaperSizesNumbers() As Integer
    Dim iPaperSizes() As POINTAPI
    Dim c As Long
    Dim iLng As Long
    
    ret = DeviceCapabilities(mDeviceName, mPort, DC_PAPERS, ByVal 0&, ByVal 0&)
    
    ReDim iPaperSizesNumbers(1 To ret)
    ReDim iPaperSizes(1 To ret)
    
    Call DeviceCapabilities(mDeviceName, mPort, DC_PAPERS, iPaperSizesNumbers(1), ByVal 0&)
    Call DeviceCapabilities(mDeviceName, mPort, DC_PAPERSIZE, iPaperSizes(1), ByVal 0&)
    
    For c = 1 To UBound(iPaperSizesNumbers)
        If iPaperSizesNumbers(c) = nPaperSizeNumber Then
            GetPaperSize.X = iPaperSizes(c).X
            GetPaperSize.Y = iPaperSizes(c).Y
            If GetPaperSize.X > GetPaperSize.Y Then
                iLng = GetPaperSize.X
                GetPaperSize.X = GetPaperSize.Y
                GetPaperSize.Y = iLng
            End If
            Exit Function
        End If
    Next c
End Function

Public Property Get DevModePtr() As Long
    DevModePtr = mDevModePtr
End Property


Public Property Let CustomColors(Index As Integer, nValue As Long)
    If (Index < 0) Or (Index > 15) Then Exit Property
    mCustomColors(Index) = nValue
End Property

Public Property Get CustomColors(Index As Integer) As Long
    CustomColors = mCustomColors(Index)
End Property


Public Property Let AutoSaveCustomColors(nValue As Boolean)
    mAutoSaveCustomColors = nValue
End Property

Public Property Get AutoSaveCustomColors() As Boolean
    AutoSaveCustomColors = mAutoSaveCustomColors
End Property

Private Sub InitCustomColors()
    Dim c As Long
    Dim iByte As Byte
    
    For c = 0 To 15
        iByte = 255 - c * 16
        mCustomColors(c) = RGB(iByte, iByte, iByte)
    Next c
End Sub

Public Property Get PrinterDefault() As Variant
Attribute PrinterDefault.VB_MemberFlags = "40"
    '
End Property

Public Property Let PrinterDefault(ByVal vNewValue As Variant)
    '
End Property

Private Function GetActiveWindowHwnd() As Long
    GetActiveWindowHwnd = GetForegroundWindow
    If GetWindowThreadProcessId(GetActiveWindowHwnd, 0&) <> App.ThreadID Then
        GetActiveWindowHwnd = 0
    End If
End Function

Private Function CloneFont(nOrigFont) As StdFont
    Dim iFont As New StdFont
    
    If nOrigFont Is Nothing Then Exit Function
    If Not TypeOf nOrigFont Is StdFont Then Exit Function
    
    iFont.Name = nOrigFont.Name
    iFont.Size = nOrigFont.Size
    iFont.Bold = nOrigFont.Bold
    iFont.Italic = nOrigFont.Italic
    iFont.Strikethrough = nOrigFont.Strikethrough
    iFont.Underline = nOrigFont.Underline
    iFont.Weight = nOrigFont.Weight
    iFont.Charset = nOrigFont.Charset
    
    Set CloneFont = iFont
End Function

