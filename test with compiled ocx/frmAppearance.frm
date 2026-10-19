VERSION 5.00
Object = "{66E63055-5A66-4C79-9327-4BC077858695}#16.0#0"; "NewTab01.ocx"
Begin VB.Form frmAppearance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Themes and more"
   ClientHeight    =   5880
   ClientLeft      =   2304
   ClientTop       =   1776
   ClientWidth     =   9840
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
   ScaleHeight     =   5880
   ScaleWidth      =   9840
   Begin NewTabCtl.NewTab NewTab1 
      Height          =   5400
      Left            =   240
      TabIndex        =   14
      Top             =   240
      Width           =   6600
      _ExtentX        =   11642
      _ExtentY        =   9525
      ControlJustAdded=   0   'False
      ControlsBackColorApplied=   -2147483633
      ControlVersion  =   19
      TabCaption(0)   =   "General"
      Tab(0).ControlCount=   6
      Tab(0).Control(0)=   "txtName"
      Tab(0).Control(1)=   "txtEmail"
      Tab(0).Control(2)=   "chkNews"
      Tab(0).Control(3)=   "chkRemember"
      Tab(0).Control(4)=   "lblName"
      Tab(0).Control(5)=   "lblEmail"
      TabCaption(1)   =   "Options"
      Tab(1).ControlCount=   4
      Tab(1).Control(0)=   "fraSize"
      Tab(1).Control(1)=   "lblColor"
      Tab(1).Control(2)=   "cboColor"
      Tab(1).Control(3)=   "chkAutoSave"
      TabCaption(2)   =   "Notes"
      Tab(2).ControlCount=   1
      Tab(2).Control(0)=   "txtNotes"
      Begin VB.TextBox txtName 
         Height          =   324
         Left            =   1440
         TabIndex        =   1
         Top             =   720
         Width           =   3600
      End
      Begin VB.TextBox txtEmail 
         Height          =   324
         Left            =   1440
         TabIndex        =   3
         Top             =   1170
         Width           =   3600
      End
      Begin VB.CheckBox chkNews 
         Caption         =   "Send me the newsletter"
         Height          =   300
         Left            =   1440
         TabIndex        =   4
         Top             =   1650
         Width           =   3600
      End
      Begin VB.CheckBox chkRemember 
         Caption         =   "Remember my settings"
         Height          =   300
         Left            =   1440
         TabIndex        =   5
         Top             =   2010
         Value           =   1  'Checked
         Width           =   3600
      End
      Begin VB.Frame fraSize 
         Caption         =   "Text size"
         Height          =   1500
         Left            =   -74760
         TabIndex        =   9
         Top             =   720
         Width           =   2400
         Begin VB.OptionButton optSmall 
            Caption         =   "Small"
            Height          =   300
            Left            =   180
            TabIndex        =   6
            Top             =   330
            Width           =   2000
         End
         Begin VB.OptionButton optMedium 
            Caption         =   "Medium"
            Height          =   300
            Left            =   180
            TabIndex        =   7
            Top             =   690
            Value           =   -1  'True
            Width           =   2000
         End
         Begin VB.OptionButton optLarge 
            Caption         =   "Large"
            Height          =   300
            Left            =   180
            TabIndex        =   8
            Top             =   1050
            Width           =   2000
         End
      End
      Begin VB.ComboBox cboColor 
         Height          =   315
         Left            =   -72120
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1050
         Width           =   2160
      End
      Begin VB.CheckBox chkAutoSave 
         Caption         =   "Save automatically"
         Height          =   300
         Left            =   -74760
         TabIndex        =   12
         Top             =   2460
         Width           =   4000
      End
      Begin VB.TextBox txtNotes 
         BorderStyle     =   0  'None
         Height          =   3300
         Left            =   -74760
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Text            =   "frmAppearance.frx":0000
         Top             =   720
         Width           =   4800
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Top             =   750
         Width           =   1140
      End
      Begin VB.Label lblEmail 
         BackStyle       =   0  'Transparent
         Caption         =   "E-mail:"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   1200
         Width           =   1140
      End
      Begin VB.Label lblColor 
         Caption         =   "Color:"
         Height          =   255
         Left            =   -72120
         TabIndex        =   10
         Top             =   780
         Width           =   2160
      End
   End
   Begin VB.ComboBox cboStyle 
      Height          =   315
      Left            =   7080
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   570
      Width           =   2520
   End
   Begin VB.ComboBox cboOrientation 
      Height          =   315
      Left            =   7080
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   1290
      Width           =   2520
   End
   Begin VB.ComboBox cboTheme 
      Height          =   315
      Left            =   7080
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   2010
      Width           =   2520
   End
   Begin VB.CheckBox chkRightToLeft 
      Caption         =   "RightToLeft"
      Height          =   300
      Left            =   7080
      TabIndex        =   21
      Top             =   2520
      Width           =   2520
   End
   Begin VB.CheckBox chkTransition 
      Caption         =   "Fade between tabs"
      Height          =   300
      Left            =   7080
      TabIndex        =   22
      Top             =   2880
      Width           =   2520
   End
   Begin VB.Label lblStyle 
      Caption         =   "Style:"
      Height          =   255
      Left            =   7080
      TabIndex        =   15
      Top             =   300
      Width           =   2520
   End
   Begin VB.Label lblOrientation 
      Caption         =   "TabOrientation:"
      Height          =   255
      Left            =   7080
      TabIndex        =   17
      Top             =   1020
      Width           =   2520
   End
   Begin VB.Label lblTheme 
      Caption         =   "Theme:"
      Height          =   255
      Left            =   7080
      TabIndex        =   19
      Top             =   1740
      Width           =   2520
   End
   Begin VB.Label lblHint 
      Caption         =   "The controls in the tabs follow the body when the orientation or the style changes, and take the colors of the theme."
      ForeColor       =   &H00808080&
      Height          =   1200
      Left            =   7080
      TabIndex        =   23
      Top             =   3360
      Width           =   2520
   End
End
Attribute VB_Name = "frmAppearance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mSyncing As Boolean

Private Sub Form_Load()
    Dim iTheme As NewTabTheme
    
    mSyncing = True
    AddItem cboStyle, "Tabbed dialog", ssStyleTabbedDialog
    AddItem cboStyle, "Property page", ssStylePropertyPage
    AddItem cboStyle, "TabStrip", ntStyleTabStrip
    AddItem cboStyle, "Flat", ntStyleFlat
    AddItem cboStyle, "Windows", ntStyleWindows
    
    AddItem cboOrientation, "Top", ssTabOrientationTop
    AddItem cboOrientation, "Bottom", ssTabOrientationBottom
    AddItem cboOrientation, "Left", ssTabOrientationLeft
    AddItem cboOrientation, "Right", ssTabOrientationRight
    AddItem cboOrientation, "TopRight", ntTabOrientationTopRight
    AddItem cboOrientation, "BottomRight", ntTabOrientationBottomRight
    AddItem cboOrientation, "LeftTop", ntTabOrientationLeftTop
    AddItem cboOrientation, "RightBottom", ntTabOrientationRightBottom
    AddItem cboOrientation, "LeftHorizontal", ntTabOrientationLeftHorizontal
    AddItem cboOrientation, "RightHorizontal", ntTabOrientationRightHorizontal
    AddItem cboOrientation, "LeftBottomHorizontal", ntTabOrientationLeftBottomHorizontal
    AddItem cboOrientation, "RightBottomHorizontal", ntTabOrientationRightBottomHorizontal
    
    For Each iTheme In NewTab1.Themes
        cboTheme.AddItem iTheme.Name
    Next
    
    cboColor.AddItem "Black"
    cboColor.AddItem "Blue"
    cboColor.AddItem "Green"
    cboColor.ListIndex = 0
    
    chkRightToLeft.Value = IIf(NewTab1.RightToLeft, vbChecked, vbUnchecked)
    chkTransition.Value = IIf(NewTab1.TabTransition <> ntAnimationNone, vbChecked, vbUnchecked)
    mSyncing = False
    SyncCombos
End Sub

Private Sub AddItem(nCombo As ComboBox, ByVal nText As String, ByVal nValue As Long)
    nCombo.AddItem nText
    nCombo.ItemData(nCombo.NewIndex) = nValue
End Sub

' A theme sets the style, and a style change can leave the control on no named theme ("Custom"),
' so after any change the three combos are made to show what the control has
Private Sub SyncCombos()
    Dim c As Long
    
    mSyncing = True
    SelectItemData cboStyle, NewTab1.Style
    SelectItemData cboOrientation, NewTab1.TabOrientation
    cboTheme.ListIndex = -1
    For c = 0 To cboTheme.ListCount - 1
        If cboTheme.List(c) = NewTab1.Theme Then
            cboTheme.ListIndex = c
            Exit For
        End If
    Next
    mSyncing = False
End Sub

Private Sub SelectItemData(nCombo As ComboBox, ByVal nValue As Long)
    Dim c As Long
    
    For c = 0 To nCombo.ListCount - 1
        If nCombo.ItemData(c) = nValue Then
            nCombo.ListIndex = c
            Exit Sub
        End If
    Next
End Sub

Private Sub cboStyle_Click()
    If mSyncing Then Exit Sub
    NewTab1.Style = cboStyle.ItemData(cboStyle.ListIndex)
    SyncCombos
End Sub

Private Sub cboOrientation_Click()
    If mSyncing Then Exit Sub
    NewTab1.TabOrientation = cboOrientation.ItemData(cboOrientation.ListIndex)
    SyncCombos
End Sub

Private Sub cboTheme_Click()
    If mSyncing Then Exit Sub
    If cboTheme.ListIndex < 0 Then Exit Sub
    NewTab1.Theme = cboTheme.Text
    SyncCombos
End Sub

Private Sub chkRightToLeft_Click()
    NewTab1.RightToLeft = (chkRightToLeft.Value = vbChecked)
End Sub

Private Sub chkTransition_Click()
    If chkTransition.Value = vbChecked Then
        NewTab1.TabTransition = ntAnimationMedium
    Else
        NewTab1.TabTransition = ntAnimationNone
    End If
End Sub

Private Sub NewTab1_Resize()
    Const cPadding As Long = 15
    
    NewTab1.ControlMove txtNotes, NewTab1.ClientLeft + cPadding, NewTab1.ClientTop + cPadding, NewTab1.ClientWidth - cPadding * 2, NewTab1.ClientHeight - cPadding * 2
End Sub
