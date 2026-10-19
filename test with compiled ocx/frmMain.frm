VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NewTab samples"
   ClientHeight    =   5712
   ClientLeft      =   2244
   ClientTop       =   1812
   ClientWidth     =   7080
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
   ScaleHeight     =   5712
   ScaleWidth      =   7080
   Begin VB.CommandButton cmdAppearance 
      Caption         =   "Themes and more"
      Height          =   420
      Left            =   240
      TabIndex        =   2
      Top             =   1260
      Width           =   2160
   End
   Begin VB.CommandButton cmdIcons 
      Caption         =   "Icons"
      Height          =   420
      Left            =   240
      TabIndex        =   4
      Top             =   2040
      Width           =   2160
   End
   Begin VB.CommandButton cmdCustomTDI 
      Caption         =   "Your own TDI"
      Height          =   420
      Left            =   240
      TabIndex        =   6
      Top             =   2820
      Width           =   2160
   End
   Begin VB.CommandButton cmdAdvanced 
      Caption         =   "Advanced"
      Height          =   420
      Left            =   240
      TabIndex        =   8
      Top             =   3600
      Width           =   2160
   End
   Begin VB.CommandButton cmdTDIControls 
      Caption         =   "TDI Controls mode"
      Height          =   420
      Left            =   240
      TabIndex        =   10
      Top             =   4380
      Width           =   2160
   End
   Begin VB.Label lblTitle 
      Caption         =   "NewTab samples"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   0
      Top             =   180
      Width           =   6600
   End
   Begin VB.Label lblIntro 
      Caption         =   "Each sample shows one area of the control. Open the forms in the IDE to see how they are set up."
      Height          =   480
      Left            =   240
      TabIndex        =   1
      Top             =   540
      Width           =   6600
   End
   Begin VB.Label lblAppearance 
      Caption         =   "Controls placed in the tabs at design time, as with SSTab. Change the style, the orientation and the theme while it runs."
      Height          =   720
      Left            =   2580
      TabIndex        =   3
      Top             =   1296
      Width           =   4260
   End
   Begin VB.Label lblIcons 
      Caption         =   "Icons from fonts and images, in their different positions."
      Height          =   600
      Left            =   2580
      TabIndex        =   5
      Top             =   2070
      Width           =   4260
   End
   Begin VB.Label lblCustomTDI 
      Caption         =   "A document program built with AddTab, RemoveTab and two events: you decide how the tabs behave."
      Height          =   600
      Left            =   2580
      TabIndex        =   7
      Top             =   2850
      Width           =   4260
   End
   Begin VB.Label lblAdvanced 
      Caption         =   "Colors per tab, color emoji icons, and a tab order that the user changes by dragging, saved and restored."
      Height          =   600
      Left            =   2580
      TabIndex        =   9
      Top             =   3630
      Width           =   4260
   End
   Begin VB.Label lblTDIControls 
      Caption         =   "The built-in mode: the control creates and closes the tabs and their controls by itself."
      Height          =   600
      Left            =   2580
      TabIndex        =   11
      Top             =   4410
      Width           =   4260
   End
   Begin VB.Label lblTDIForms 
      Caption         =   "TDI Forms mode, where each form of the program opens in a tab, is in the TDIFormsTest project of this group."
      ForeColor       =   &H00808080&
      Height          =   480
      Left            =   240
      TabIndex        =   12
      Top             =   5220
      Width           =   6600
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAppearance_Click()
    frmAppearance.Show vbModal
End Sub

Private Sub cmdIcons_Click()
    frmIcons.Show vbModal
End Sub

Private Sub cmdCustomTDI_Click()
    frmCustomTDI.Show vbModal
End Sub

Private Sub cmdAdvanced_Click()
    frmAdvanced.Show vbModal
End Sub

Private Sub cmdTDIControls_Click()
    frmTDIControlsTest.Show vbModal
End Sub
