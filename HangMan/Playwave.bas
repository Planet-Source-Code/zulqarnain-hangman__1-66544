Attribute VB_Name = "modPlayWave"
' ********************************************
' Copyright ©Zulqarnain F. Sarani, 2006
' All Rights Reserved.
' ********************************************

'Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" _
(ByVal lpSound As String, ByVal flag As Long) As Long

Sub PlayWave(sFileName As String)
    On Error GoTo Play_Err
    
    Dim iReturn As Integer
    
    'Make sure something was passed to the Play Function
    If sFileName > "" Then
        'Make sure a WAV filename was passed
        If UCase$(Right$(sFileName, 3)) = "WAV" Then
            'Make sure the file exists
            If Dir(sFileName) > "" Then
                iReturn = sndPlaySound(sFileName, 0)
            End If
        End If
    End If
    
    'Wav file play successful
    Exit Sub

Play_Err:
    'If there was an error then exit without playing
    Exit Sub
End Sub

