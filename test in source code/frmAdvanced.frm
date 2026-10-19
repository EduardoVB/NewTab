VERSION 5.00
Object = "{66E63055-5A66-4C79-9327-4BC077858695}#16.0#0"; "NewTab01.ocx"
Begin VB.Form frmAdvanced 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Advanced"
   ClientHeight    =   5160
   ClientLeft      =   2280
   ClientTop       =   1836
   ClientWidth     =   8400
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   8400
   Begin NewTabCtl.NewTab NewTab1 
      Height          =   3840
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   6773
      Tabs            =   5
      Style           =   3
      ControlsBackColorApplied=   16777215
      TabTransition   =   2
      CanReorderTabs  =   -1  'True
      ControlVersion  =   19
      TabCaption(0)   =   "Inbox"
      Tab(0).ControlCount=   1
      Tab(0).Control(0)=   "lblTabInbox"
      TabCaption(1)   =   "Calendar"
      Tab(1).ControlCount=   1
      Tab(1).Control(0)=   "lblTabCalendar"
      TabCaption(2)   =   "Contacts"
      Tab(2).ControlCount=   1
      Tab(2).Control(0)=   "lblTabContacts"
      TabCaption(3)   =   "Notes"
      Tab(3).ControlCount=   1
      Tab(3).Control(0)=   "lblTabNotes"
      TabCaption(4)   =   "Archive"
      Tab(4).ControlCount=   1
      Tab(4).Control(0)=   "lblTabArchive"
      Begin VB.Label lblTabInbox 
         BackColor       =   &H00FFFFFF&
         Caption         =   $"frmAdvanced.frx":0000
         Height          =   1500
         Left            =   360
         TabIndex        =   0
         Top             =   900
         Width           =   7200
      End
      Begin VB.Label lblTabCalendar 
         BackColor       =   &H00FFFFFF&
         Caption         =   $"frmAdvanced.frx":00C4
         Height          =   1500
         Left            =   -74640
         TabIndex        =   1
         Top             =   900
         Width           =   7200
      End
      Begin VB.Label lblTabContacts 
         BackColor       =   &H00FFFFFF&
         Caption         =   $"frmAdvanced.frx":0156
         Height          =   1500
         Left            =   -74640
         TabIndex        =   2
         Top             =   900
         Width           =   7200
      End
      Begin VB.Label lblTabNotes 
         BackColor       =   &H00FFFFFF&
         Caption         =   "A tab can be disabled with TabEnabled: try the check box below the control."
         Height          =   1500
         Left            =   -74640
         TabIndex        =   3
         Top             =   900
         Width           =   7200
      End
      Begin VB.Label lblTabArchive 
         BackColor       =   &H00FFFFFF&
         Caption         =   "This is the tab that the check box disables."
         Height          =   1500
         Left            =   -74640
         TabIndex        =   4
         Top             =   900
         Width           =   7200
      End
   End
   Begin VB.CheckBox chkDisableArchive 
      Caption         =   "Disable the Archive tab"
      Height          =   300
      Left            =   240
      TabIndex        =   6
      Top             =   4260
      Width           =   3600
   End
   Begin VB.CommandButton cmdResetOrder 
      Caption         =   "Restore the default order"
      Height          =   420
      Left            =   5400
      TabIndex        =   7
      Top             =   4200
      Width           =   2760
   End
End
Attribute VB_Name = "frmAdvanced"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const cRegApp As String = "NewTab samples"
Private Const cRegSection As String = "Advanced"
Private Const cRegKey As String = "TabOrder"

Private mDefaultOrder As Variant
Private mRestoringOrder As Boolean

Private Sub Form_Load()
    Dim iIcons As Variant
    Dim iBackColors As Variant
    Dim iBarColors As Variant
    Dim iFont As StdFont
    Dim t As Long
    
    mDefaultOrder = Array("inbox", "calendar", "contacts", "notes", "archive")
    iIcons = Array("1F4E5", "1F4C5", "1F465", "1F4DD", "1F5C4")
    iBackColors = Array(RGB(232, 240, 254), RGB(230, 244, 234), RGB(254, 247, 224), RGB(243, 232, 253), RGB(238, 238, 238))
    iBarColors = Array(RGB(26, 115, 232), RGB(30, 142, 62), RGB(242, 153, 0), RGB(161, 66, 244), RGB(95, 99, 104))
    
    For t = 0 To UBound(mDefaultOrder)
        NewTab1.TabKey(t) = mDefaultOrder(t)
        ' Each tab keeps its own icon font, so it is given one per tab
        Set iFont = New StdFont
        iFont.Name = "Segoe UI Emoji"
        iFont.Size = 16
        Set NewTab1.TabIconFont(t) = iFont
        NewTab1.TabIconCharHex(t) = iIcons(t)
        NewTab1.TabIconInColor(t) = True
        NewTab1.TabCustomColor(t, ntCCBackColorTab) = iBackColors(t)
        NewTab1.TabCustomColor(t, ntCCBackColorSelectedTab) = iBackColors(t)
        NewTab1.TabCustomColor(t, ntCCFlatBarColorSelectedTab) = iBarColors(t)
    Next
    RestoreTabOrder
End Sub

Private Sub NewTab1_TabReordered(ByVal CurrentIndex As Long, ByVal OldIndex As Long)
    If Not mRestoringOrder Then SaveTabOrder
End Sub

Private Sub chkDisableArchive_Click()
    NewTab1.TabEnabled("archive") = (chkDisableArchive.Value = vbUnchecked)
End Sub

Private Sub cmdResetOrder_Click()
    On Error Resume Next
    DeleteSetting cRegApp, cRegSection, cRegKey
    On Error GoTo 0
    ApplyTabOrder mDefaultOrder
End Sub

' The order is saved as the tab keys separated by commas
Private Sub SaveTabOrder()
    Dim iKeys() As String
    Dim t As Long
    
    ReDim iKeys(NewTab1.Tabs - 1)
    For t = 0 To NewTab1.Tabs - 1
        iKeys(t) = NewTab1.TabKey(t)
    Next
    SaveSetting cRegApp, cRegSection, cRegKey, Join(iKeys, ",")
End Sub

Private Sub RestoreTabOrder()
    Dim iSaved As String
    
    iSaved = GetSetting(cRegApp, cRegSection, cRegKey, "")
    If iSaved <> "" Then ApplyTabOrder Split(iSaved, ",")
End Sub

' A key saved by an older version of the program that no longer exists is skipped
Private Sub ApplyTabOrder(nKeys As Variant)
    Dim iPos As Long
    Dim iCurrent As Long
    Dim c As Long
    
    mRestoringOrder = True
    For c = 0 To UBound(nKeys)
        iCurrent = TabIndexOfKey(nKeys(c))
        If iCurrent > -1 Then
            If iCurrent <> iPos Then NewTab1.MoveTab iCurrent, iPos
            iPos = iPos + 1
        End If
    Next
    mRestoringOrder = False
End Sub

Private Function TabIndexOfKey(ByVal nKey As String) As Long
    Dim t As Long
    
    For t = 0 To NewTab1.Tabs - 1
        If NewTab1.TabKey(t) = nKey Then
            TabIndexOfKey = t
            Exit Function
        End If
    Next
    TabIndexOfKey = -1
End Function
