VERSION 5.00
Begin VB.Form frmCategoryOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Category Options"
   ClientHeight    =   4470
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   3495
   Icon            =   "frmCategoryOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstCategories 
      Height          =   2760
      Left            =   300
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   1080
      Width           =   2895
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   4020
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   4020
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Place a check mark next to each category that you want to play.  If you don't care for a category, then  uncheck that category."
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmCategoryOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ********************************************
' Copyright ©Zulqarnain F. Sarani, 2006
' All Rights Reserved.
' ********************************************

Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim iCounter As Integer
    Dim bSomethingEnabled As Boolean
    
    gsCategory = ""
    For iCounter = 0 To lstCategories.ListCount - 1
        gclsCategories(iCounter + 1).Enabled = lstCategories.Selected(iCounter)
        gclsCategories(iCounter + 1).SavePhrases
        bSomethingEnabled = True
    Next
    
    If Not bSomethingEnabled Then
        MsgBox "At least one category needs to be checked.", vbInformation, "Nothing Checked"
        Exit Sub
    End If
    
    Unload Me
End Sub

Private Sub Form_Load()
    Dim iCounter As Integer
    Dim sCat As String
    
    With gclsCategories
        For iCounter = 1 To .Count
            lstCategories.AddItem .Item(iCounter).CategoryName
            If .Item(iCounter).Enabled Then
                lstCategories.Selected(lstCategories.NewIndex) = True
            End If
        Next
    End With
    
End Sub
