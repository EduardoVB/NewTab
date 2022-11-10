VERSION 5.00
Begin VB.Form frmNewTabSelectControl 
   Caption         =   "Select control"
   ClientHeight    =   3444
   ClientLeft      =   8688
   ClientTop       =   3612
   ClientWidth     =   5856
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   7.8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNewTabSelectControl.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3444
   ScaleWidth      =   5856
   Begin VB.ListBox lstControls 
      Height          =   2736
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5844
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   435
      Left            =   4044
      TabIndex        =   1
      Top             =   2880
      Width           =   1515
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   435
      Left            =   2448
      TabIndex        =   0
      Top             =   2880
      Width           =   1515
   End
End
Attribute VB_Name = "frmNewTabSelectControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public SelectedControl As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim iLng As Long
    
    If lstControls.ListIndex > -1 Then
        SelectedControl = lstControls.Text
        iLng = InStr(SelectedControl, "[now")
        If iLng > 0 Then
            SelectedControl = Trim(Left(SelectedControl, iLng - 1))
        End If
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Move Screen.Width / 2 - Me.Width / 2, Screen.Height / 2 - Me.Height / 2
End Sub
