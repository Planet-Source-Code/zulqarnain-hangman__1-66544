Attribute VB_Name = "modTimer"
' ********************************************
' Copyright ©Zulqarnain F. Sarani, 2006
' All Rights Reserved.
' ********************************************

'-------------------------------------------------------------------------------
' This module works hand-in-hand with the DropDownHelper class.
'-------------------------------------------------------------------------------

Option Explicit

'-------------------------------------------------------------------------------
'Timer APIs:

Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, _
    ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) _
    As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, _
    ByVal nIDEvent As Long) As Long

'-------------------------------------------------------------------------------
'A list of pointers to timer objects. The list uses timer IDs as the keys.

Public gcTimerObjects As SortedList

'-------------------------------------------------------------------------------
'The timer code:

Private Sub TimerProc(ByVal lHwnd As Long, ByVal lMsg As Long, _
    ByVal lTimerID As Long, ByVal lTime As Long)

    Dim nPtr As Long
    Dim oTimerObject As objTimer

'Debug.Print "TimerProc is firing"

    'Create a Timer object from the pointer
    nPtr = gcTimerObjects.ItemByKey(lTimerID)
    CopyMemory oTimerObject, nPtr, 4
    'Call a method which will fire the Timer event
    oTimerObject.Tick
    'Get rid of the Timer object so that VB will not try to release it
    CopyMemory oTimerObject, 0&, 4
End Sub

Public Function StartTimer(lInterval As Long) As Long
    StartTimer = SetTimer(0, 0, lInterval, AddressOf TimerProc)
End Function

Public Sub StopTimer(lTimerID As Long)
    KillTimer 0, lTimerID
End Sub

Public Sub SetInterval(lInterval As Long, lTimerID As Long)
    SetTimer 0, lTimerID, lInterval, AddressOf TimerProc
End Sub
