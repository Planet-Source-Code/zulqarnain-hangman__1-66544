VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ********************************************
' Copyright ©Zulqarnain F. Sarani, 2006
' All Rights Reserved.
' ********************************************

Option Explicit

'local variable(s) to hold property value(s)
Private mvarEnabled As Boolean 'local copy
Private mvarFilename As String 'local copy
Private mPhrases() As String
'local variable(s) to hold property value(s)
Private mvarCategoryName As String 'local copy
'local variable(s) to hold property value(s)
Private mvarCategoryChanged As Boolean 'local copy
Public Property Let CategoryChanged(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CategoryChanged = 5
    mvarCategoryChanged = vData
End Property

Public Function GetPhrases() As String()
    'Dim iCounter As Integer
    'Dim sPhrases As String
    
    'For iCounter = 0 To UBound(mPhrases) - 1
    '    sPhrases = sPhrases & mPhrase(iCounter) & vbCrLf
    'Next
    'sPhrases = Join(mPhrases, vbCrLf)
    
    GetPhrases = mPhrases
End Function
Public Sub SetPhrases(sPhrases() As String)
    mPhrases = sPhrases 'Split(sPhrases, vbCrLf)
    mvarCategoryChanged = True
End Sub
Public Property Get CategoryChanged() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CategoryChanged
    CategoryChanged = mvarCategoryChanged
End Property



Public Property Let CategoryName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CategoryName = 5
    mvarCategoryName = vData
    mvarCategoryChanged = True
    mvarFilename = mvarCategoryName & ".hm"
End Property

Public Property Get CategoryName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CategoryName
    CategoryName = mvarCategoryName
End Property

Public Property Let FileName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Filename = 5
    mvarFilename = vData
End Property


Public Property Get FileName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Filename
    FileName = mvarFilename
End Property

Public Property Get Count() As Integer
    On Error Resume Next
    If mPhrases(0) = "" Then
        Count = 0
    Else
        Count = UBound(mPhrases) + 1
    End If
    
    
End Property

Public Sub SavePhrases()
    On Error Resume Next
    
    Dim iFnum As Integer
    Dim sLine As String
    Dim iPhraseCount As Integer
    Dim iCounter As Integer
    
    If mvarCategoryChanged Then
        iFnum = FreeFile
        
        Open App.Path & "\" & mvarFilename For Output As iFnum
        
        For iCounter = 0 To UBound(mPhrases)
            If mPhrases(iCounter) > "" Then
                Print #iFnum, mPhrases(iCounter)
            End If
        Next
        
        If mvarEnabled Then
            Print #iFnum, "CATEGORY=ON"
        Else
            Print #iFnum, "CATEGORY=OFF"
        End If
        
        Close iFnum
    End If
    
    mvarCategoryChanged = False
    
End Sub

Public Sub LoadPhrases()
    On Error Resume Next
    
    Dim iFnum As Integer
    Dim sLine As String
    Dim iPhraseCount As Integer
    
    iFnum = FreeFile
    
    'ReDim mPhrases()
    
    If Dir$(App.Path & "\" & mvarFilename) = "" Then
        Open App.Path & "\" & mvarFilename For Output As iFnum
        Print #iFnum, "CATEGORY=ON"
        Close iFnum
    End If
    
    Open App.Path & "\" & mvarFilename For Input As iFnum
    
    iPhraseCount = -1
    Do While EOF(iFnum) = False
        Line Input #iFnum, sLine
        If EOF(iFnum) = True Then
            If UCase$(sLine) = "CATEGORY=ON" Then
                mvarEnabled = True
                Exit Do
            ElseIf UCase$(sLine) = "CATEGORY=OFF" Then
                mvarEnabled = False
                Exit Do
            Else
                mvarEnabled = True
            End If
        End If
        If sLine > "" Then
            iPhraseCount = iPhraseCount + 1
            ReDim Preserve mPhrases(iPhraseCount)
            mPhrases(iPhraseCount) = sLine
        End If
    Loop
    
    
    Close iFnum
    
End Sub

Public Function GetNewPhrase() As String
    Dim iPhraseCount As Integer
    
    iPhraseCount = UBound(mPhrases) + 1
    GetNewPhrase = mPhrases(Int(iPhraseCount * Rnd))

End Function

Public Property Let Enabled(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Enabled = 5
    If vData <> mvarEnabled Then
        mvarEnabled = vData
        mvarCategoryChanged = True
    End If
End Property


Public Property Get Enabled() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Enabled
    Enabled = mvarEnabled
End Property
