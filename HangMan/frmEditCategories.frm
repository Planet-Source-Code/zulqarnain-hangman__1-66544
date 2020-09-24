VERSION 5.00
Begin VB.Form frmEditCategories 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Categories"
   ClientHeight    =   3885
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   5055
   Icon            =   "frmEditCategories.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstCat 
      Height          =   450
      Left            =   3420
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   2160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdRename 
      Caption         =   "&Rename Category"
      Height          =   375
      Left            =   3300
      TabIndex        =   6
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete Category"
      Height          =   375
      Left            =   3300
      TabIndex        =   5
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit Category"
      Height          =   375
      Left            =   3300
      TabIndex        =   4
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add Category"
      Height          =   375
      Left            =   3300
      TabIndex        =   3
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3300
      TabIndex        =   2
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3300
      TabIndex        =   1
      Top             =   2760
      Width           =   1695
   End
   Begin VB.ListBox lstCategories 
      Height          =   3375
      Left            =   180
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label lblReg 
      BackStyle       =   0  'Transparent
      Caption         =   "* Only registered users may modify the list of categories."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   4875
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   3660
      Width           =   2775
   End
End
Attribute VB_Name = "frmEditCategories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ********************************************
' Copyright ©Zulqarnain F. Sarani, 2006
' All Rights Reserved.
' ********************************************

Option Explicit
Dim miCatIndex As Integer

Private Sub cmdAdd_Click()
    Dim sNewCat As String
    
    sNewCat = InputBox("Enter the name of the category", "New Category")
    
    If sNewCat > "" Then
        If ValidName(sNewCat) Then
            gclsCategories.Add App.Path & "\" & sNewCat & ".hm", sNewCat
            lstCategories.AddItem sNewCat & " (Empty)"
            lstCat.AddItem sNewCat
            lstCategories.ListIndex = lstCategories.NewIndex
        Else
            MsgBox "The name you typed is invalid." & vbCrLf & vbCrLf & "A category name can contain up to 20 characters, including spaces." & vbCrLf & "But, it cannot contain any of the following characters: \ / : * ? " & Chr$(32) & " < > |", vbInformation + vbOKOnly, "Invalid Name"
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    Dim iRet As Integer
    
    iRet = MsgBox("Are you sure you want to delete the " & gsEditCat & " category and all of its puzzles?", vbQuestion + vbYesNo, "Delete Category?")
    
    If iRet = vbYes Then
        gclsCategories.Remove gsEditCat
        LoadCategories
    End If
End Sub

Private Sub cmdEdit_Click()
    gsEditCat = lstCat.List(lstCategories.ListIndex)
    frmEdit.Show vbModal, Me
    LoadCategories
End Sub

Private Sub cmdOK_Click()
    Dim iCounter As Integer
    
    'With gclsCategories
    '    For iCounter = 1 To .Count
    '        .Item(iCounter).SavePhrases
    '    Next
    'End With
    Unload Me
End Sub

Private Sub cmdRename_Click()
    Dim sNewName As String
    
    sNewName = InputBox("Enter the new name of the category", "Rename Category", gsEditCat)
    
    If sNewName > "" Then
        If ValidName(sNewName) Then
            gclsCategories.Rename gsEditCat, sNewName
            LoadCategories
        Else
            MsgBox "The name you typed is invalid." & vbCrLf & vbCrLf & "A category name can contain up to 20 characters, including spaces." & vbCrLf & "But, it cannot contain any of the following characters: \ / : * ? " & Chr$(32) & " < > |", vbInformation + vbOKOnly, "Invalid Name"
        End If
    End If
End Sub

Private Sub Form_Load()
    
    LoadCategories
    
End Sub

Private Function ValidName(sName As String) As String
    Dim bValid As Boolean
    Dim iCounter As Integer
    
    If Len(sName) > 20 Then
        ValidName = False
        Exit Function
    End If
    
    bValid = True
    
    For iCounter = 1 To Len(sName)
        bValid = ValidChar(Mid$(sName, iCounter, 1))
        If Not bValid Then
            Exit For
        End If
    Next
    
    ValidName = bValid
End Function

Private Function ValidChar(sChar As String) As Boolean
    Select Case sChar
        Case "\", "/", ":", "*", "?", "<", ">", ">", "|"
            ValidChar = False
        Case Chr$(34)
            ValidChar = False
        Case Else
            ValidChar = True
    End Select
End Function
Private Sub LoadCategories()
    Dim iCounter As Integer
    Dim sCat As String
    Dim sCount As String
    Dim lPhraseCount As Long
    
    lstCategories.Clear
    lstCat.Clear
    
    With gclsCategories
        For iCounter = 1 To .Count
            If .Item(iCounter).Count = 0 Then
                sCount = " (Empty)"
            Else
                sCount = " (" & CStr(.Item(iCounter).Count) & ")"
                lPhraseCount = lPhraseCount + .Item(iCounter).Count
            End If
            lstCategories.AddItem .Item(iCounter).CategoryName & sCount
            lstCat.AddItem .Item(iCounter).CategoryName
        Next
    End With
    lblTotal.Caption = lPhraseCount & " Total Puzzles"
    If lstCategories.ListCount > 0 Then
        lstCategories.ListIndex = 0
    End If

End Sub
Private Sub lstCategories_Click()
    gsEditCat = lstCat.List(lstCategories.ListIndex)
    'If lstCategories.ListIndex > -1 Then
    '    txtPhrases = gclsCategories(lstCategories.List(lstCategories.ListIndex)).GetPhrases
    'End If
End Sub

Private Sub lstCategories_DblClick()
    cmdEdit_Click
End Sub

Private Sub txtPhrases_Change()
    miCatIndex = lstCategories.ListIndex

End Sub

Private Sub txtPhrases_GotFocus()
    miCatIndex = lstCategories.ListIndex
    
End Sub

Private Sub txtPhrases_LostFocus()
   ' gclsCategories(miCatIndex + 1).SetPhrases txtPhrases
End Sub
