Attribute VB_Name = "GlobalAPI"
' ********************************************
' Copyright ©Zulqarnain F. Sarani, 2006
' All Rights Reserved.
' ********************************************

Public Type POINTAPI
    x As Long
    y As Long
End Type
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, _
    ByVal yPoint As Long) As Long


Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    pDest As Any, pSource As Any, ByVal ByteLen As Long)


Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (lpszSoundName As Any, ByVal uFlags As Long) As Long
Public Const SND_ASYNC = &H1
Public Const SND_LOOP = &H8
Public Const SND_MEMORY = &H4
Public Const SND_NODEFAULT = &H2
Public Const SND_NOSTOP = &H10
Public Const SND_SYNC = &H0

