Attribute VB_Name = "modHyperlink"
' ********************************************
' Copyright ©Zulqarnain F. Sarani, 2006
' All Rights Reserved.
' ********************************************

Option Explicit

Declare Function ShellExecute _
   Lib "shell32.dll" Alias "ShellExecuteA" _
   (ByVal hWnd As Long, _
   ByVal lpOperation As String, _
   ByVal lpFile As String, _
   ByVal lpParameters As String, _
   ByVal lpDirectory As String, _
   ByVal nShowCmd As Long) As Long

Public Const SW_SHOWNORMAL = 1


Public Sub GotoDesertwareWeb()
    
    Dim lReturn As Long
    
    lReturn = ShellExecute(frmMain.hWnd, "open", "http://www.desertware.com/games/index.htm", vbNull, vbNull, SW_SHOWNORMAL)

End Sub

