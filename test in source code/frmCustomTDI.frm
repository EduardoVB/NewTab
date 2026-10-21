VERSION 5.00
Object = "*\A..\control-source\NewTabCtl.vbp"
Begin VB.Form frmCustomTDI 
   Caption         =   "Your own TDI"
   ClientHeight    =   5400
   ClientLeft      =   2316
   ClientTop       =   2148
   ClientWidth     =   7800
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   7800
   Begin VB.TextBox txtDoc 
      Height          =   600
      Index           =   0
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   4500
      Visible         =   0   'False
      Width           =   1200
   End
   Begin NewTabCtl.NewTab NewTab1 
      Height          =   4740
      Left            =   0
      TabIndex        =   2
      Top             =   660
      Width           =   7800
      _ExtentX        =   13758
      _ExtentY        =   8361
      Tabs            =   1
      ControlsBackColorApplied=   -2147483633
      NewTabButton    =   1
      CloseButton     =   -1  'True
      TabsOverflow    =   2
      CanReorderTabs  =   -1  'True
      ControlVersion  =   19
      TabCaption(0)   =   "Document 1"
   End
   Begin VB.Label lblHelp 
      Caption         =   $"frmCustomTDI.frx":0000
      Height          =   480
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   7560
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmCustomTDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' A document is a TextBox loaded from the txtDoc(0) template, which stays hidden on the form, and
' put in a tab of its own. The tab's key ties the tab to its TextBox, and it travels with the tab
' when the user drags it to another position.

Private mDocumentCount As Long
Private mModified() As Boolean
Private mTitle() As String

Private Sub Form_Load()
    ReDim mModified(0)
    ReDim mTitle(0)
    ' The tab placed at design time takes the first document
    NewDocument 0
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    lblHelp.Width = Me.ScaleWidth - lblHelp.Left * 2
    If Me.ScaleHeight > NewTab1.Top Then NewTab1.Move 0, NewTab1.Top, Me.ScaleWidth, Me.ScaleHeight - NewTab1.Top
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim iTab As Long
    
    For iTab = NewTab1.Tabs - 1 To 0 Step -1
        If Not CloseDocument(iTab) Then
            Cancel = 1
            Exit Sub
        End If
    Next
End Sub

Private Function NewDocument(Optional ByVal nExistingTab As Long = -1) As Long
    Dim iDoc As Long
    Dim iTab As Long
    
    mDocumentCount = mDocumentCount + 1
    iDoc = txtDoc.UBound + 1
    ReDim Preserve mModified(iDoc)
    ReDim Preserve mTitle(iDoc)
    mTitle(iDoc) = "Document " & mDocumentCount
    
    If nExistingTab > -1 Then
        iTab = nExistingTab
        NewTab1.TabCaption(iTab) = mTitle(iDoc)
    Else
        iTab = NewTab1.AddTab(mTitle(iDoc))
    End If
    NewTab1.TabKey(iTab) = DocumentKey(iDoc)
    
    Load txtDoc(iDoc)
    Set txtDoc(iDoc).Container = NewTab1
    NewTab1.ControlSetTab txtDoc(iDoc), iTab
    NewTab1.TabSel = iTab
    ' Keeps the TextBox filling the tab when the form is resized
    NewTab1.FitToClientArea txtDoc(iDoc)
    txtDoc(iDoc).Visible = True
    If Me.Visible Then txtDoc(iDoc).SetFocus
    NewDocument = iDoc
End Function

' Returns False if the user chose to keep the document open
Private Function CloseDocument(ByVal nTab As Long) As Boolean
    Dim iDoc As Long
    
    iDoc = DocumentOfTab(nTab)
    If mModified(iDoc) Then
        NewTab1.TabSel = nTab
        Select Case MsgBox("Save the changes to " & mTitle(iDoc) & "?", vbYesNoCancel Or vbQuestion)
            Case vbCancel
                Exit Function
            Case vbYes
                ' A real program saves the document here
        End Select
    End If
    NewTab1.ReleaseFromClientArea txtDoc(iDoc)
    Unload txtDoc(iDoc)
    ' A tab can only be removed once it holds no controls
    NewTab1.RemoveTab nTab
    CloseDocument = True
End Function

Private Function DocumentKey(ByVal nDoc As Long) As String
    DocumentKey = "Doc" & nDoc
End Function

Private Function DocumentOfTab(ByVal nTab As Long) As Long
    DocumentOfTab = Val(Mid$(NewTab1.TabKey(nTab), 4))
End Function

Private Sub NewTab1_NewTabButtonClick(Cancel As Boolean)
    NewDocument
End Sub

Private Sub NewTab1_CloseButtonClick(ByVal nTab As Long, Cancel As Boolean)
    ' Left alone, the control would only hide the tab
    Cancel = True
    CloseDocument nTab
End Sub

Private Sub txtDoc_Change(Index As Integer)
    If Not mModified(Index) Then
        mModified(Index) = True
        ' The key finds the tab wherever the user has dragged it
        NewTab1.TabCaption(DocumentKey(Index)) = mTitle(Index) & " *"
    End If
End Sub

Private Sub mnuNew_Click()
    NewDocument
End Sub

Private Sub mnuClose_Click()
    If NewTab1.TabSel > -1 Then CloseDocument NewTab1.TabSel
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub
