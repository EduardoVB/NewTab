VERSION 5.00
Object = "*\A..\control-source\NewTabCtl.vbp"
Begin VB.Form frmTDIControlsTest 
   Caption         =   "TDI"
   ClientHeight    =   5748
   ClientLeft      =   2964
   ClientTop       =   2568
   ClientWidth     =   7680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   5748
   ScaleWidth      =   7680
   StartUpPosition =   2  'CenterScreen
   Begin NewTabCtl.NewTab NewTab1 
      Height          =   2530
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5290
      _ExtentX        =   9335
      _ExtentY        =   4466
      Tabs            =   2
      ForeColorTabSel =   10184001
      ForeColorHighlighted=   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   3
      TabHeight       =   600
      Themed          =   0   'False
      BackColorTabSel =   16250871
      FlatBarColorInactive=   14211288
      FlatBorderColor =   10184001
      HighlightColor  =   16477710
      IconAlignment   =   5
      AutoTabHeight   =   -1  'True
      FlatRoundnessTabs=   8
      TabMousePointerHand=   -1  'True
      IconColorMouseHover=   255
      IconColorMouseHoverTabSel=   255
      IconColorTabHighlighted=   16777215
      HighlightMode   =   12
      HighlightModeTabSel=   10
      FlatBorderMode  =   1
      FlatBarHeight   =   0
      CanReorderTabs  =   -1  'True
      TDIMode         =   1
      TabIconChar(0)  =   57606
      TabIconLeftOffset(0)=   -3
      TabIconTopOffset(0)=   1
      TabCaption(0)   =   "New tab template   "
      Tab(0).ControlCount=   4
      Tab(0).Control(0)=   "txtDoc(0)"
      Tab(0).Control(1)=   "Command1(0)"
      Tab(0).Control(2)=   "txtSearch(0)"
      Tab(0).Control(3)=   "Label1(0)"
      BeginProperty IconFont(0) {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe MDL2 Assets"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconFontName(0) =   "Segoe MDL2 Assets"
      TabIconChar(1)  =   63658
      TabIconLeftOffset(1)=   -2
      TabIconTopOffset(1)=   1
      TabToolTipText(1)=   "Add a new tab"
      Tab(1).ControlCount=   0
      BeginProperty IconFont(1) {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe MDL2 Assets"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconFontName(1) =   "Segoe MDL2 Assets"
      Begin VB.TextBox txtDoc 
         Appearance      =   0  'Flat
         Height          =   1210
         Index           =   0
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   1116
         Width           =   5050
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00F7F7F7&
         Caption         =   "Do something"
         Height          =   370
         Index           =   0
         Left            =   3840
         TabIndex        =   3
         Top             =   516
         Width           =   1330
      End
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         Height          =   370
         Index           =   0
         Left            =   1320
         TabIndex        =   2
         Top             =   516
         Width           =   2410
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F7F7F7&
         Caption         =   "Search:"
         ForeColor       =   &H009B6541&
         Height          =   370
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   566
         Width           =   970
      End
   End
   Begin VB.Label Label2 
      Caption         =   "TDI: Tabbed Document Interface. To use the TDI Mode 'Controls', set the TDIMode property to 'ntTDIModeControls'."
      Height          =   972
      Left            =   600
      TabIndex        =   5
      Top             =   3000
      Visible         =   0   'False
      Width           =   4452
   End
End
Attribute VB_Name = "frmTDIControlsTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
    Dim c As Long
    
    NewTab1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    
    For c = 0 To NewTab1.Tabs - 2
        If NewTab1.TabVisible(c) Then
            PositionAndResizeControlsInTab c
        End If
    Next
End Sub

Private Sub NewTab1_TDINewTabAdded(ByVal TabNumber As Long)
    PositionAndResizeControlsInTab TabNumber
End Sub

Private Sub PositionAndResizeControlsInTab(TabNumber As Long)
    NewTab1.ControlMove "txtDoc(" & TabNumber & ")", Screen.TwipsPerPixelX, NewTab1.ClientTop + 700, NewTab1.ClientWidth, NewTab1.ClientHeight - 700 - Screen.TwipsPerPixelY
End Sub

