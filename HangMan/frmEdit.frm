VERSION 5.00
Begin VB.Form frmEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit"
   ClientHeight    =   5715
   ClientLeft      =   2010
   ClientTop       =   2130
   ClientWidth     =   6585
   Icon            =   "frmEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4800
      TabIndex        =   6
      Top             =   5160
      Width           =   1695
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   4800
      TabIndex        =   5
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton cmdDeletePuzzle 
      Caption         =   "&Delete Puzzle"
      Height          =   375
      Left            =   4800
      TabIndex        =   4
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit Puzzle"
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add Puzzle"
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin VB.ListBox lstPhrases 
      Height          =   4740
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   840
      Width           =   4335
   End
   Begin VB.Label lblReg 
      BackStyle       =   0  'Transparent
      Caption         =   "* Only registered users may modify the list of puzzles."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   5235
   End
   Begin VB.Label lblCategory 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   60
      Width           =   6015
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ********************************************
' Copyright ©Zulqarnain F. Sarani, 2006
' All Rights Reserved.
' ********************************************

Option Explicit

Private Sub cmdAdd_Click()
    Dim sPhrase As String
    
    sPhrase = InputBox("Enter your new puzzle.", "New Puzzle")
    
    If sPhrase > "" Then
        If ValidPhrase(sPhrase) Then
            lstPhrases.AddItem sPhrase
            lstPhrases.ListIndex = lstPhrases.NewIndex
        Else
            MsgBox "There are too many letters in one of the words in your puzzle.  Each word may be no larger than 20 characters.", vbInformation + vbOKOnly, "Invalid Puzzle"
        End If
    End If
End Sub

Private Function ValidPhrase(sNew As String) As Boolean
    Dim aPhrase() As String
    Dim iCounter As Integer
    Dim bValid As Boolean
    
    aPhrase = Split(sNew, " ")
    
    bValid = True
    For iCounter = 0 To UBound(aPhrase)
        If Len(aPhrase(iCounter)) > 20 Then
            bValid = False
            Exit For
        End If
    Next
    
    ValidPhrase = bValid
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDeletePuzzle_Click()
    lstPhrases.RemoveItem lstPhrases.ListIndex

End Sub

Private Sub cmdEdit_Click()
    Dim sPhrase As String
    Dim sOldPhrase As String
    
    sPhrase = InputBox("Edit the text of the current puzzle.", "Edit Puzzle", lstPhrases.List(lstPhrases.ListIndex))
    
    If sPhrase > "" Then
        If ValidPhrase(sPhrase) Then
            lstPhrases.List(lstPhrases.ListIndex) = sPhrase
        Else
            MsgBox "There are too many letters in one of the words in your puzzle.  Each word may be no larger than 20 characters.", vbInformation + vbOKOnly, "Invalid Puzzle"
        End If
    Else
        MsgBox "You cannot have a blank puzzle.", vbOKOnly + vbInformation, "Blank Puzzle"
    End If
    
End Sub

Private Sub cmdOK_Click()
    Dim aPhrases() As String
    Dim iCounter As Integer
    
    If lstPhrases.ListCount > 0 Then
        ReDim aPhrases(lstPhrases.ListCount - 1)
        
        For iCounter = 0 To lstPhrases.ListCount - 1
            aPhrases(iCounter) = lstPhrases.List(iCounter)
        Next
    Else
        ReDim aPhrases(0)
    End If
    
    gclsCategories(gsEditCat).SetPhrases aPhrases
    gclsCategories(gsEditCat).SavePhrases
    
    Unload Me
End Sub

Private Sub Form_Load()
    Dim aPhrases() As String
    Dim iCounter As Integer
    
    aPhrases = gclsCategories(gsEditCat).GetPhrases
    
    For iCounter = 0 To gclsCategories(gsEditCat).Count - 1
        lstPhrases.AddItem aPhrases(iCounter)
    Next
       
    Me.Caption = "Edit " & gsEditCat
    lblCategory = gsEditCat
End Sub
