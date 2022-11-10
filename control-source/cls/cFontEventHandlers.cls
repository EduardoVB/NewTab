VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFontEventHandlers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Event FontChanged(ByVal PropertyName As String)

Private mFontsEventHandlers As Collection

Public Sub RaiseFontEvent(ByRef PropertyName As String)
    RaiseEvent FontChanged(PropertyName)
End Sub

Private Sub Class_Initialize()
    Set mFontsEventHandlers = New Collection
End Sub

Private Sub Class_Terminate()
    Set mFontsEventHandlers = Nothing
End Sub

Public Sub AddFont(nFont As StdFont, nTabIndex As Long)
    Dim iFeh As cFontEventHandler
    
    Set iFeh = New cFontEventHandler
    iFeh.SetFont nFont, Me
    mFontsEventHandlers.Add iFeh, CStr(ObjPtr(nFont)) & "_" & CStr(nTabIndex)
End Sub

Public Sub RemoveFont(nFont As StdFont, nTabIndex As Long)
    Dim iFeh As cFontEventHandler
    Dim iKey As String
    
    iKey = CStr(ObjPtr(nFont)) & "_" & CStr(nTabIndex)
    Set iFeh = mFontsEventHandlers(iKey)
    iFeh.Release
    mFontsEventHandlers.Remove iKey
End Sub

Public Sub Release()
    Dim iFeh As cFontEventHandler
    
    For Each iFeh In mFontsEventHandlers
        iFeh.Release
    Next
    Set mFontsEventHandlers = New Collection
End Sub
