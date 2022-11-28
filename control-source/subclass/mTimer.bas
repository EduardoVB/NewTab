Attribute VB_Name = "mTimer"
Option Explicit

' declares:
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long

Const cTimerMax = 100

' Array of timers
Public aTimers(1 To cTimerMax) As cTimer
' Added SPM to prevent excessive searching through aTimers array:
Private m_cTimerCount As Integer

Function TimerCreate(Timer As cTimer) As Boolean
    ' Create the Timer
    'Debug.Print m_cTimerCount
    Timer.TimerID = SetTimer(0&, 0&, Timer.Interval, AddressOf TimerProc)
    If Timer.TimerID Then
        TimerCreate = True
        Dim i As Integer
        For i = 1 To cTimerMax
            If aTimers(i) Is Nothing Then
                Set aTimers(i) = Timer
                If (i > m_cTimerCount) Then
                    m_cTimerCount = i
                End If
                TimerCreate = True
                Exit Function
            End If
        Next
        Timer.ErrRaise eeTooManyTimers
    Else
        ' TimerCreate = False
        Timer.TimerID = 0
        Timer.Interval = 0
    End If
End Function

Public Function TimerDestroy(Timer As cTimer) As Long
    ' TimerDestroy = False
    ' Find and remove this Timer
    Dim i As Integer, f As Boolean
    ' SPM - no need to count past the last Timer set up in the
    ' aTimer array:
    For i = 1 To m_cTimerCount
        ' Find Timer in array
        If Not aTimers(i) Is Nothing Then
            If Timer.TimerID = aTimers(i).TimerID Then
                f = KillTimer(0, Timer.TimerID)
                ' Remove Timer and set reference to nothing
                Set aTimers(i) = Nothing
                TimerDestroy = True
                Exit Function
            End If
        ' SPM: aTimers(1) could well be nothing before
        ' aTimers(2) is.  This original [else] would leave
        ' Timer 2 still running when the class terminates -
        ' not very nice!  Causes serious GPF in IE and VB design
        ' mode...
        'Else
        '    TimerDestroy = True
        '    Exit Function
        End If
    Next
End Function


Sub TimerProc(ByVal hWnd As Long, ByVal uMsg As Long, _
                     ByVal idEvent As Long, ByVal dwTime As Long)
    Dim i As Integer
    ' Find the Timer with this ID
    If m_cTimerCount = 0 Then
        Call KillTimer(0, idEvent)
    Else
        For i = 1 To m_cTimerCount
            ' SPM: Add a check to ensure aTimers(i) is not nothing!
            ' This would occur if we had two timers declared from
            ' the same thread and we terminated the first one before
            ' the second!  Causes serious GPF if we don't do this...
            If Not (aTimers(i) Is Nothing) Then
                If idEvent = aTimers(i).TimerID Then
                    ' Generate the event
                    aTimers(i).PulseTimer
                    Exit Sub
                End If
            End If
        Next
    End If
End Sub


'Private Function StoreTimer(Timer As cTimer)
'    Dim i As Integer
'    For i = 1 To m_cTimerCount
'        If aTimers(i) Is Nothing Then
'            Set aTimers(i) = Timer
'            StoreTimer = True
'            Exit Function
'        End If
'    Next
'End Function




