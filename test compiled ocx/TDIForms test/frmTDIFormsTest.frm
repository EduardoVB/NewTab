VERSION 5.00
Object = "{66E63055-5A66-4C79-9327-4BC077858695}#11.0#0"; "NewTab01.ocx"
Begin VB.Form frmTDIFormsTest 
   Caption         =   "TDI forms test"
   ClientHeight    =   8112
   ClientLeft      =   4308
   ClientTop       =   2520
   ClientWidth     =   13632
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8112
   ScaleWidth      =   13632
   Begin NewTabCtl.NewTab NewTab1 
      Height          =   4332
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   5892
      _ExtentX        =   10393
      _ExtentY        =   7641
      ControlJustAdded=   0   'False
      Tabs            =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabHeight       =   783
      IconAlignment   =   5
      AutoTabHeight   =   -1  'True
      IconColorMouseHover=   255
      IconColorMouseHoverSelectedTab=   255
      CanReorderTabs  =   -1  'True
      TDIMode         =   2
      ControlVersion  =   11
      TabCaption(0)   =   "Home"
      Tab(0).ControlCount=   6
      Tab(0).Control(0)=   "Command6"
      Tab(0).Control(1)=   "Command5"
      Tab(0).Control(2)=   "Command4"
      Tab(0).Control(3)=   "Command1"
      Tab(0).Control(4)=   "Command2"
      Tab(0).Control(5)=   "Command3"
      BeginProperty IconFont(0) {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe MDL2 Assets"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconFontName(0) =   "Segoe MDL2 Assets"
      Begin VB.CommandButton Command6 
         Caption         =   "Show Form4 (non-TDI-child)"
         Height          =   612
         Left            =   600
         TabIndex        =   7
         Top             =   3240
         Width           =   1892
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Close all forms (tabs)"
         Height          =   612
         Left            =   2880
         TabIndex        =   5
         Top             =   2436
         Width           =   1892
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Launch new Form2 instance"
         Height          =   612
         Left            =   2880
         TabIndex        =   4
         Top             =   1596
         Width           =   1892
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Show Form1"
         Height          =   612
         Left            =   600
         TabIndex        =   3
         Top             =   756
         Width           =   1892
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Show Form2 (default Form2)"
         Height          =   612
         Left            =   600
         TabIndex        =   2
         Top             =   1596
         Width           =   1892
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Show Form3 modal"
         Height          =   612
         Left            =   600
         TabIndex        =   1
         Top             =   2436
         Width           =   1892
      End
   End
   Begin VB.Image Image1 
      Height          =   8028
      Left            =   6960
      Picture         =   "frmTDIFormsTest.frx":0000
      Top             =   1680
      Visible         =   0   'False
      Width           =   6312
   End
   Begin VB.Label Label1 
      Caption         =   $"frmTDIFormsTest.frx":10213E
      Height          =   1212
      Left            =   7200
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   3852
   End
End
Attribute VB_Name = "frmTDIFormsTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command6_Click()
    Form4.Show
End Sub

Private Sub Form_Load()
    NewTab1.TabCaption(0) = "Home, Menu or Start"
End Sub

Private Sub Form_Resize()
    NewTab1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub Command1_Click()
    Form1.Show
End Sub

Private Sub Command2_Click()
    Form2.Show
End Sub

Private Sub Command4_Click()
    Dim frm As New Form2
    
    frm.Show
End Sub

Private Sub Command3_Click()
    Form3Modal.Show vbModal
End Sub

Private Sub Command5_Click()
    Dim frm As Form
    
    For Each frm In Forms
        If Not frm Is Me And Not frm Is Form4 Then
            Unload frm
        End If
    Next
End Sub

Private Sub NewTab1_TDIBeforeNewTab(ByVal TabType As NewTabCtl.NTTDINewTabTypeConstants, ByVal TabNumber As Long, TabCaption As String, LoadControls As Boolean, ShowTabCloseButton As Boolean, Cancel As Boolean)
    If TabCaption = "Form4" Then
        Cancel = True
    End If
End Sub
