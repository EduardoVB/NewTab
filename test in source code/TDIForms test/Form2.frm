VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
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
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   9084
   ScaleWidth      =   5604
   Begin VB.TextBox Text2 
      Height          =   7212
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Text            =   "Form2.frx":030A
      Top             =   1200
      Width           =   5052
   End
   Begin VB.TextBox Text1 
      Height          =   372
      Left            =   960
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   360
      Width           =   3852
   End
   Begin VB.Label Label1 
      Caption         =   "Find:"
      Height          =   372
      Left            =   480
      TabIndex        =   2
      Top             =   420
      Width           =   492
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
    If (Me.WindowState <> vbMinimized) And (Me.ScaleWidth <> 0) Then
        Text2.Move 0, Text1.Top + Text1.Height + 200, Me.ScaleWidth, Me.ScaleHeight - Text1.Top - Text1.Height - 200
    End If
End Sub
