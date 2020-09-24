VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CJukebox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ********************************************
' Copyright ©Zulqarnain F. Sarani, 2006
' All Rights Reserved.
' ********************************************

Option Explicit
Private msCurrentCategory As String
Private msCurrentPhrase As String
Private mCol As New Collection

Public Property Get Category() As String
    Category = msCurrentCategory
End Property

Public Property Get Phrase() As String
    Phrase = msCurrentPhrase
End Property

Public Sub GetNewPhrase()
    Dim iNumber As Integer
    Dim iCount As Integer
    
    If mCol.Count = 0 Then
        ResetList
    End If
    
    iCount = mCol.Count
    
    iNumber = Int((iCount) * Rnd) + 1
    
    msCurrentCategory = mCol.Item(iNumber).Category
    msCurrentPhrase = mCol.Item(iNumber).Phrase
    
    mCol.Remove iNumber
End Sub

Public Sub ResetList()
    Dim clsItem As CJukeboxItem
    Dim aPhrases() As String
    Dim iCounter As Integer
    Dim iPhrases As Integer
    
    Do While mCol.Count > 0
        mCol.Remove 1
    Loop
    
    For iCounter = 1 To gclsCategories.Count
        If gclsCategories.Item(iCounter).Enabled And gclsCategories.Item(iCounter).Count > 0 Then
            aPhrases = gclsCategories.Item(iCounter).GetPhrases
            For iPhrases = 0 To UBound(aPhrases)
                Set clsItem = New CJukeboxItem
                clsItem.Category = gclsCategories.Item(iCounter).CategoryName
                clsItem.Phrase = aPhrases(iPhrases)
                mCol.Add clsItem
            Next
        End If
    Next
    
End Sub
