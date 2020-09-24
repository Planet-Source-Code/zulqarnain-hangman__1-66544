Attribute VB_Name = "modHangman"
' ********************************************
' Copyright ©Zulqarnain F. Sarani, 2006
' All Rights Reserved.
' ********************************************

Option Explicit

Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_SHOWNORMAL = 1

Public gsCategory As String
Public gclsCategories As New CCategories
Public gclsJukebox As New CJukebox
Public glPoints As Long
Public glCurrentRound As Long
Public gbSolveCanceled As Boolean

Public Const gcCorrectGuess = 50
Public Const gcBonus = 200
Public Const gcVowel = 100
Public Const gcWrongGuess = 100
Public Const gcInitialCash = 0
Public Const gcGameWon = 200
Public Const gcGameLost = 200
Public gsShareware(299) As String
Public gsSWCat(299) As String
Public giSWCount As Integer

Public Const HANGMAN_CONTENTS = 0
Public Const HOW_TO_PLAY_HANGMAN = 1
Public Const HOW_TO_REGISTER = 3
Public Const WHY_REGISTER = 2

Public gsEditCat As String
Public Animate As CAnimate
