Attribute VB_Name = "mBSPropsDB"
Option Explicit

Public gWinPropsDB As New Collection
       
Public Function MySetProp(nHwnd As LongPtr, nPropertyName As String, nData As LongPtr) As Long
    
    On Error Resume Next
    Err.Clear
    gWinPropsDB.Add nData, CStr(nHwnd) & "|" & nPropertyName
    If Err.Number = 457 Then
        On Error GoTo ErrorExit:
        gWinPropsDB.Remove CStr(nHwnd) & "|" & nPropertyName
        gWinPropsDB.Add nData, CStr(nHwnd) & "|" & nPropertyName
    End If
    'Debug.Print gWinPropsDB.count
    MySetProp = 1
    Exit Function

ErrorExit:
    MySetProp = 0
    Err.Clear
End Function

Public Function MyGetProp(nHwnd As LongPtr, nPropertyName As String) As LongPtr
    
    On Error GoTo ErrorExit:
    MyGetProp = gWinPropsDB(CStr(nHwnd) & "|" & nPropertyName)
    Exit Function

ErrorExit:
    MyGetProp = 0
    'Debug.Print "Error MyGetProp, hWnd" & nHwnd & ", PropertyName: " & nPropertyName
    Err.Clear
End Function

Public Function MyRemoveProp(nHwnd As LongPtr, nPropertyName As String) As LongPtr
    
    On Error Resume Next
    MyRemoveProp = gWinPropsDB(CStr(nHwnd) & "|" & nPropertyName)
    On Error GoTo ErrorExit:
    gWinPropsDB.Remove CStr(nHwnd) & "|" & nPropertyName
    Exit Function

ErrorExit:
    MyRemoveProp = 0
    Err.Clear
End Function

