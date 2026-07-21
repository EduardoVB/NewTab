VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IBSSubclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Shim so classic VB6 understands the LongPtr type (twinBASIC already has it natively).
' Has to live in a public class (not a standard module), otherwise VB6 throws
' "Only public objects and types can be used as parameters... for public procedures of public classes"
' when it's used in this interface's parameters.
#Const VBA7 = 0

#If VBA7 = 0 Then
    Public Enum LongPtr
        [_]
    End Enum
#End If

Public Function MsgResponse(ByVal hWnd As LongPtr, ByVal iMsg As Long) As Long
    '
End Function

Function WindowProc(ByVal hWnd As LongPtr, ByVal iMsg As Long, ByRef wParam As LongPtr, ByRef lParam As LongPtr, ByRef bConsume As Boolean) As LongPtr
Attribute WindowProc.VB_Description = "Raised whenever a message you have subclassed is sent."
    '
End Function

Public Sub UnsubclassIt()

End Sub
