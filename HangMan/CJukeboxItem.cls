VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CJukeboxItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ********************************************
' Copyright ©Zulqarnain F. Sarani, 2006
' All Rights Reserved.
' ********************************************

Option Explicit
Private msCategory As String
Private msPhrase As String

Public Property Get Category() As String
    Category = msCategory
End Property

Public Property Let Category(sCat As String)
    msCategory = sCat
End Property

Public Property Get Phrase() As String
    Phrase = msPhrase
End Property

Public Property Let Phrase(sPhrase As String)
    msPhrase = sPhrase
End Property
