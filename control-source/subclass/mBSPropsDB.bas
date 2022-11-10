Attribute VB_Name = "mBSPropsDB"
Option Explicit

Public gWinPropsDB As New Collection
       
Public Function MySetProp(nHwnd As Long, nPropertyName As String, nData As Long) As Long
    
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

Public Function MyGetProp(nHwnd As Long, nPropertyName As String) As Long
    
    On Error GoTo ErrorExit:
    MyGetProp = gWinPropsDB(CStr(nHwnd) & "|" & nPropertyName)
    Exit Function

ErrorExit:
    MyGetProp = 0
    'Debug.Print "Error MyGetProp, hWnd" & nHwnd & ", PropertyName: " & nPropertyName
    Err.Clear
End Function

Public Function MyRemoveProp(nHwnd As Long, nPropertyName As String) As Long
    
    On Error Resume Next
    MyRemoveProp = gWinPropsDB(CStr(nHwnd) & "|" & nPropertyName)
    On Error GoTo ErrorExit:
    gWinPropsDB.Remove CStr(nHwnd) & "|" & nPropertyName
    Exit Function

ErrorExit:
    MyRemoveProp = 0
    Err.Clear
End Function

