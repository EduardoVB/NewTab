VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFontEventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mFont As StdFont
Attribute mFont.VB_VarHelpID = -1
Private mParent As cFontEventHandlers

Public Sub SetFont(ByRef nFont As StdFont, ByRef nParent As cFontEventHandlers)
    Set mFont = nFont
    Set mParent = nParent
End Sub

Public Sub Release()
    Set mParent = Nothing
End Sub

Private Sub mFont_FontChanged(ByVal PropertyName As String)
    mParent.RaiseFontEvent PropertyName
End Sub
