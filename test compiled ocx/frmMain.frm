VERSION 5.00
Object = "{66E63055-5A66-4C79-9327-4BC077858695}#8.0#0"; "NewTab01.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NewTab sample"
   ClientHeight    =   5040
   ClientLeft      =   1524
   ClientTop       =   1932
   ClientWidth     =   5844
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
   ScaleHeight     =   5040
   ScaleWidth      =   5844
   Begin VB.CommandButton Command3 
      Caption         =   "All current preset themes at once"
      Height          =   370
      Left            =   360
      TabIndex        =   5
      Top             =   4560
      Width           =   5130
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Check some tabs with fonts icons"
      Height          =   370
      Left            =   360
      TabIndex        =   4
      Top             =   4080
      Width           =   5130
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Test TDI Controls mode (TDI: Tabbed Document Interface)"
      Height          =   370
      Left            =   360
      TabIndex        =   3
      Top             =   3600
      Width           =   5130
   End
   Begin VB.ComboBox cboThemes 
      Height          =   312
      Left            =   1050
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   4440
   End
   Begin NewTabCtl.NewTab NewTab1 
      Height          =   2530
      Left            =   360
      TabIndex        =   2
      Top             =   720
      Width           =   5130
      _ExtentX        =   9059
      _ExtentY        =   4466
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabHeight       =   600
      Themed          =   0   'False
      AutoTabHeight   =   -1  'True
      TabCaption(0)   =   "Tab 0"
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlCount=   0
   End
   Begin VB.Label Label1 
      Caption         =   "Theme:"
      Height          =   250
      Left            =   360
      TabIndex        =   0
      Top             =   150
      Width           =   730
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboThemes_Click()
    NewTab1.Theme = cboThemes.Text
End Sub

Private Sub Command1_Click()
    frmTDIControlsTest.Show vbModal
End Sub

Private Sub Command2_Click()
    frmIcons.Show vbModal
End Sub

Private Sub Command3_Click()
    frmAllThemes.Show vbModal
End Sub

Private Sub Form_Load()
    Dim iTheme As NewTabTheme
    
    For Each iTheme In NewTab1.Themes
        cboThemes.AddItem iTheme.Name
    Next
    cboThemes.ListIndex = 0
End Sub
