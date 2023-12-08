VERSION 5.00
Begin VB.Form Form3Modal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form3 Modal"
   ClientHeight    =   3444
   ClientLeft      =   2340
   ClientTop       =   2064
   ClientWidth     =   5148
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3444
   ScaleWidth      =   5148
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   492
      Left            =   3360
      TabIndex        =   1
      Top             =   2640
      Width           =   1332
   End
   Begin VB.Label Label1 
      Caption         =   "This is a modal form."
      Height          =   1092
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   3132
   End
End
Attribute VB_Name = "Form3Modal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Unload Me
End Sub
