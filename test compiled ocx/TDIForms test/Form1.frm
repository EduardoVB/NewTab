VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9084
   ClientLeft      =   2340
   ClientTop       =   2064
   ClientWidth     =   5604
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
   ScaleHeight     =   9084
   ScaleWidth      =   5604
   Begin VB.PictureBox picContainer 
      BorderStyle     =   0  'None
      Height          =   5172
      Left            =   600
      ScaleHeight     =   5172
      ScaleWidth      =   3972
      TabIndex        =   0
      Top             =   480
      Width           =   3972
      Begin VB.CommandButton Command1 
         Caption         =   "Close"
         Height          =   612
         Left            =   1200
         TabIndex        =   4
         Top             =   4080
         Width           =   1892
      End
      Begin VB.PictureBox Picture1 
         Height          =   1212
         Left            =   1080
         ScaleHeight     =   1164
         ScaleWidth      =   2004
         TabIndex        =   3
         Top             =   0
         Width           =   2052
      End
      Begin VB.ComboBox Combo1 
         Height          =   336
         Left            =   1320
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   2040
         Width           =   1452
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Show Form3 modal"
         Height          =   612
         Left            =   1200
         TabIndex        =   1
         Top             =   3360
         Width           =   1892
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    Form3Modal.Show vbModal
End Sub

Private Sub Form_Resize()
    picContainer.Move (Me.ScaleWidth - picContainer.Height) / 2, (Me.ScaleHeight - picContainer.Height) / 2
End Sub
