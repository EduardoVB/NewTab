Attribute VB_Name = "mNewTabTexts"
Option Explicit

' The texts a developer saves as defaults for the NewTab controls he places from then on. They live in
' the registry and are only read at design time. Registry strings go through the ANSI code page, so
' everything outside printable ASCII is stored as \uXXXX

Public Const cTextsLast As Long = 5

Private Const cTextsSection As String = "Texts"
' Never produced by EncodeText, which escapes every backslash
Private Const cTextMissing As String = "\"

Public Function TextPropertyName(ByVal nIndex As Long) As String
    Select Case nIndex
        Case 0: TextPropertyName = "NewTabButtonToolTipText"
        Case 1: TextPropertyName = "StripButtonBackToolTipText"
        Case 2: TextPropertyName = "StripButtonForwardToolTipText"
        Case 3: TextPropertyName = "StripButtonListToolTipText"
        Case 4: TextPropertyName = "TDIDefaultTabCaption"
        Case 5: TextPropertyName = "TDINewTabCaption"
    End Select
End Function

Public Function TextFactoryValue(ByVal nIndex As Long) As String
    Select Case nIndex
        Case 4: TextFactoryValue = cPropDef_TDIDefaultTabCaption
        Case 5: TextFactoryValue = cPropDef_TDINewTabCaption
    End Select
End Function

Public Function SavedTextsExist(ByVal nAppName As String) As Boolean
    Dim iTexts() As String
    
    SavedTextsExist = GetSavedTexts(nAppName, iTexts)
End Function

' Fills all of nTexts, with the factory value where nothing was saved. False when nothing was
Public Function GetSavedTexts(ByVal nAppName As String, ByRef nTexts() As String) As Boolean
    Dim iStored As String
    Dim c As Long
    
    ReDim nTexts(cTextsLast)
    For c = 0 To cTextsLast
        iStored = GetSetting(nAppName, cTextsSection, TextPropertyName(c), cTextMissing)
        If iStored = cTextMissing Then
            nTexts(c) = TextFactoryValue(c)
        Else
            nTexts(c) = DecodeText(iStored)
            GetSavedTexts = True
        End If
    Next c
End Function

' Saving the factory texts removes the key, so that nothing is left behind that means nothing
Public Sub SaveTexts(ByVal nAppName As String, ByRef nTexts() As String)
    Dim iAllFactory As Boolean
    Dim c As Long
    
    iAllFactory = True
    For c = 0 To cTextsLast
        If nTexts(c) <> TextFactoryValue(c) Then iAllFactory = False
    Next c
    If iAllFactory Then
        On Error Resume Next
        DeleteSetting nAppName, cTextsSection
        On Error GoTo 0
    Else
        For c = 0 To cTextsLast
            SaveSetting nAppName, cTextsSection, TextPropertyName(c), EncodeText(nTexts(c))
        Next c
    End If
End Sub

Private Function EncodeText(ByVal nText As String) As String
    Dim iCode As Long
    Dim c As Long
    
    For c = 1 To Len(nText)
        iCode = AscW(Mid$(nText, c, 1)) And &HFFFF&
        If (iCode >= 32) And (iCode <= 126) And (iCode <> 92) Then
            EncodeText = EncodeText & ChrW$(iCode)
        Else
            EncodeText = EncodeText & "\u" & Right$("000" & Hex$(iCode), 4)
        End If
    Next c
End Function

Private Function DecodeText(ByVal nText As String) As String
    Dim iHex As String
    Dim c As Long
    
    c = 1
    Do While c <= Len(nText)
        iHex = Mid$(nText, c + 2, 4)
        If (Mid$(nText, c, 2) = "\u") And (Len(iHex) = 4) And Not (UCase$(iHex) Like "*[!0-9A-F]*") Then
            DecodeText = DecodeText & ChrW$(Val("&H" & iHex & "&"))
            c = c + 6
        Else
            DecodeText = DecodeText & Mid$(nText, c, 1)
            c = c + 1
        End If
    Loop
End Function
