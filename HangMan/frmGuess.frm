VERSION 5.00
Begin VB.Form frmGuess 
   AutoRedraw      =   -1  'True
   Caption         =   "Guess the remaining letters"
   ClientHeight    =   3525
   ClientLeft      =   2730
   ClientTop       =   2955
   ClientWidth     =   6645
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   15.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGuess.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   235
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   443
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2160
      TabIndex        =   1
      Top             =   3120
      Width           =   2655
   End
   Begin VB.Image imgArrow 
      Height          =   480
      Left            =   300
      Picture         =   "frmGuess.frx":030A
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "With your keyboard, type in each letter as the arrow points to it."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   60
      TabIndex        =   0
      Top             =   2640
      Width           =   6495
   End
End
Attribute VB_Name = "frmGuess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ********************************************
' Copyright ©Zulqarnain F. Sarani, 2006
' All Rights Reserved.
' ********************************************

'Arrays to store X and Y positions of
'each letter in the phrase
Private LettersLeft() As Integer
Private LettersTop() As Integer
Private LettersLetter() As String
Private LettersShown() As String
Private LettersRed() As Boolean

Private miPhraseLength As Integer
Private miCurrentLetter As Integer


Private Sub cmdCancel_Click()
    gbSolveCanceled = True
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    Dim sLetter As String
    Dim bLetterFound As Boolean
    Dim bGameOver As Boolean
    Dim iCounter As Integer
    
    sLetter = UCase(Chr$(KeyAscii))
    
    'Make sure the key that was pressed is a letter
    'and call the cmdLetter_Click event to process the letter
    If sLetter >= "A" And sLetter <= "Z" Then
        If cmdCancel.Enabled Then
            cmdCancel.Caption = "Too late to cancel now."
            cmdCancel.Enabled = False
        End If
        If sLetter = LettersLetter(miCurrentLetter) Then
            LettersShown(miCurrentLetter) = sLetter
            DrawPhrase
            SetNextLetter
        Else
            LettersShown(miCurrentLetter) = LettersLetter(miCurrentLetter)
            LettersRed(miCurrentLetter) = True
            DrawPhrase
            SetNextLetter
        End If
    End If

End Sub

Private Sub Form_Load()
    With frmMain
        miPhraseLength = .PhraseLength
        LettersLeft = .PhraseLettersLeft
        LettersTop = .PhraseLettersTop
        LettersLetter = .PhraseLettersLetter
        LettersShown = .PhraseLettersShown
        LettersRed = .PhraseLettersRed
    End With
    
    miCurrentLetter = 1
    SetNextLetter
    DrawPhrase
End Sub

Private Sub SetNextLetter()
    Dim iCounter As Integer
    Dim bLetterFound As Boolean
    
    bLetterFound = False
    For iCounter = miCurrentLetter To miPhraseLength
        If LettersShown(iCounter) = "" And (LCase$(LettersLetter(iCounter)) >= "a" And LCase$(LettersLetter(iCounter)) <= "z") Then
            miCurrentLetter = iCounter
            bLetterFound = True
            Exit For
        End If
    Next
    
    If Not bLetterFound Then
        With frmMain
            .PhraseLettersRed = LettersRed
            .PhraseLettersShown = LettersShown
        End With
        Unload Me
    Else
        imgArrow.Top = LettersTop(miCurrentLetter) - Me.TextHeight("X") - 5
        imgArrow.Left = LettersLeft(miCurrentLetter) - Me.TextWidth("X") + 3
    End If
    
End Sub

Private Sub DrawPhrase()

    'This sub draws all of the Underlines for each letter
    'and the letter itself if it has already been guessed.

    Dim iCounter As Integer
    Dim iLetterWidth As Integer
    
    iLetterWidth = TextWidth("X")
    
    For iCounter = 1 To miPhraseLength
        If LettersRed(iCounter) Then
            Me.ForeColor = &HC0&      'cmdLetter(0).ForeColor
        End If
        If LettersShown(iCounter) > "" Then
            CurrentX = LettersLeft(iCounter) + ((iLetterWidth / 2) - (TextWidth(LettersLetter(iCounter)) / 2))
            CurrentY = LettersTop(iCounter)
            Print LettersShown(iCounter)
        End If
        Me.ForeColor = vbBlack
        If (LCase$(LettersLetter(iCounter)) >= "a" And LCase$(LettersLetter(iCounter)) <= "z") Then
            CurrentX = LettersLeft(iCounter)
            CurrentY = LettersTop(iCounter)
            Print "_"
        End If
    Next
    
End Sub

