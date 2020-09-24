VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCategories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"CCategory"
Attribute VB_Ext_KEY = "Member0" ,"CCategory"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ********************************************
' Copyright ©Zulqarnain F. Sarani, 2006
' All Rights Reserved.
' ********************************************

Option Explicit

'local variable to hold collection
Private mCol As Collection
'local variable(s) to hold property value(s)
Private mvarCurrentCategory As String 'local copy

Private maCategorys() As String
Private maPhrases() As String

Public Property Let CurrentCategory(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CurrentCategory = 5
    mvarCurrentCategory = vData
End Property

Public Sub LoadSelections()
    Dim iCounter As Integer
    Dim iPhrases As Integer
    Dim aPhrases() As String
    Dim iPhraseCounter As Integer
    
    iPhraseCounter = 0
    For iCounter = 1 To mCol.Count
        aPhrases = mCol(iCounter).GetPhrases
        For iPhrases = 0 To UBound(aPhrases)
            ReDim Preserve maCategorys(iPhraseCounter)
            ReDim Preserve maPhrases(iPhraseCounter)
            maCategorys(iPhraseCounter) = mCol(iCounter).CategoryName
            maPhrases(iPhraseCounter) = aPhrases(iPhrases)
        Next
    Next
    
End Sub

Public Property Get CurrentCategory() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CurrentCategory
    CurrentCategory = mvarCurrentCategory
End Property

Public Function GetNewPhrase() As String
    On Error Resume Next
    Dim mCategories() As String
    Dim iCounter As Integer
    Dim iCatCount As Integer
    
    iCatCount = -1
    For iCounter = 1 To mCol.Count
        If mCol.Item(iCounter).Enabled Then
            iCatCount = iCatCount + 1
            ReDim Preserve mCategories(iCatCount)
            mCategories(iCatCount) = mCol.Item(iCounter).CategoryName
        End If
    Next
    
    If iCatCount = -1 Then
        iCatCount = iCatCount + 1
        ReDim Preserve mCategories(iCatCount)
        mCategories(iCatCount) = mCol.Item(1).CategoryName
    End If
    
    mvarCurrentCategory = mCategories(Int((iCatCount + 1) * Rnd))
    GetNewPhrase = mCol.Item(mvarCurrentCategory).GetNewPhrase
    
End Function

Public Function Add(FileName As String, Optional sKey As String) As CCategory
    'create a new object
    Dim objNewMember As CCategory
    Set objNewMember = New CCategory


'    objNewMember.Enabled = Enabled
    objNewMember.FileName = FileName
    objNewMember.CategoryName = sKey
    objNewMember.LoadPhrases
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As CCategory
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)
    Dim sFileName As String
    
    sFileName = App.Path & "\" & mCol(vntIndexKey).FileName
    
    mCol.Remove vntIndexKey
    
    Kill sFileName

End Sub

Public Sub Rename(vntIndexKey As Variant, sNewName As String)

    Dim sFileName As String
    Dim sNewFileName As String
    
    sFileName = App.Path & "\" & mCol(vntIndexKey).FileName
    sNewFileName = App.Path & "\" & sNewName & ".hm"
    
    mCol.Remove vntIndexKey
    
    Name sFileName As sNewFileName
    
    Add sNewFileName, sNewName
    
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

