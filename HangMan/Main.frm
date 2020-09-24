VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hangman"
   ClientHeight    =   6000
   ClientLeft      =   840
   ClientTop       =   1425
   ClientWidth     =   9510
   DrawWidth       =   2
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   15.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Main.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   400
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   634
   Begin VB.CommandButton cmdSolve 
      Caption         =   "&Solve Puzzle"
      Default         =   -1  'True
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
      Left            =   840
      TabIndex        =   33
      Top             =   3840
      Width           =   3555
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   4395
      Left            =   5220
      ScaleHeight     =   4335
      ScaleWidth      =   4200
      TabIndex        =   4
      Top             =   660
      Width           =   4260
      Begin VB.Label lblNewGame 
         BackStyle       =   0  'Transparent
         Caption         =   "Press F2 for a new game"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   6
         Top             =   3660
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label lblGameOver 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Game Over"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   495
         Left            =   360
         TabIndex        =   5
         Top             =   3180
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Line shpRope 
         BorderWidth     =   3
         X1              =   1620
         X2              =   1620
         Y1              =   1020
         Y2              =   300
      End
      Begin VB.Shape shpHead 
         BorderWidth     =   2
         Height          =   375
         Left            =   1440
         Shape           =   3  'Circle
         Top             =   1020
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Line lnBody 
         BorderWidth     =   2
         Index           =   1
         Visible         =   0   'False
         X1              =   1320
         X2              =   1620
         Y1              =   1800
         Y2              =   1500
      End
      Begin VB.Line lnBody 
         BorderWidth     =   3
         Index           =   0
         Visible         =   0   'False
         X1              =   1620
         X2              =   1620
         Y1              =   2100
         Y2              =   1380
      End
      Begin VB.Line lnBody 
         BorderWidth     =   2
         Index           =   2
         Visible         =   0   'False
         X1              =   1920
         X2              =   1620
         Y1              =   1800
         Y2              =   1500
      End
      Begin VB.Line lnBody 
         BorderWidth     =   2
         Index           =   3
         Visible         =   0   'False
         X1              =   1320
         X2              =   1620
         Y1              =   2700
         Y2              =   2100
      End
      Begin VB.Line lnBody 
         BorderWidth     =   2
         Index           =   4
         Visible         =   0   'False
         X1              =   1920
         X2              =   1620
         Y1              =   2700
         Y2              =   2100
      End
   End
   Begin VB.FileListBox lstFiles 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   6300
      TabIndex        =   1
      Top             =   3780
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.PictureBox picLetters 
      Align           =   2  'Align Bottom
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   9450
      TabIndex        =   0
      Top             =   5325
      Width           =   9510
      Begin Hangman.VBUSoftButton cmdLetter 
         Height          =   375
         Index           =   0
         Left            =   60
         TabIndex        =   7
         Top             =   120
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   661
         ForeColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "A"
         Picture         =   "Main.frx":030A
      End
      Begin Hangman.VBUSoftButton cmdLetter 
         Height          =   375
         Index           =   1
         Left            =   420
         TabIndex        =   8
         Top             =   120
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   661
         ForeColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "B"
         Picture         =   "Main.frx":0326
      End
      Begin Hangman.VBUSoftButton cmdLetter 
         Height          =   375
         Index           =   2
         Left            =   780
         TabIndex        =   9
         Top             =   120
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   661
         ForeColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "C"
         Picture         =   "Main.frx":0342
      End
      Begin Hangman.VBUSoftButton cmdLetter 
         Height          =   375
         Index           =   3
         Left            =   1140
         TabIndex        =   10
         Top             =   120
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   661
         ForeColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "D"
         Picture         =   "Main.frx":035E
      End
      Begin Hangman.VBUSoftButton cmdLetter 
         Height          =   375
         Index           =   4
         Left            =   1500
         TabIndex        =   11
         Top             =   120
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   661
         ForeColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "E"
         Picture         =   "Main.frx":037A
      End
      Begin Hangman.VBUSoftButton cmdLetter 
         Height          =   375
         Index           =   5
         Left            =   1860
         TabIndex        =   12
         Top             =   120
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   661
         ForeColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "F"
         Picture         =   "Main.frx":0396
      End
      Begin Hangman.VBUSoftButton cmdLetter 
         Height          =   375
         Index           =   6
         Left            =   2220
         TabIndex        =   13
         Top             =   120
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   661
         ForeColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "G"
         Picture         =   "Main.frx":03B2
      End
      Begin Hangman.VBUSoftButton cmdLetter 
         Height          =   375
         Index           =   7
         Left            =   2580
         TabIndex        =   14
         Top             =   120
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   661
         ForeColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "H"
         Picture         =   "Main.frx":03CE
      End
      Begin Hangman.VBUSoftButton cmdLetter 
         Height          =   375
         Index           =   8
         Left            =   2940
         TabIndex        =   15
         Top             =   120
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   661
         ForeColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "I"
         Picture         =   "Main.frx":03EA
      End
      Begin Hangman.VBUSoftButton cmdLetter 
         Height          =   375
         Index           =   9
         Left            =   3300
         TabIndex        =   16
         Top             =   120
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   661
         ForeColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "J"
         Picture         =   "Main.frx":0406
      End
      Begin Hangman.VBUSoftButton cmdLetter 
         Height          =   375
         Index           =   10
         Left            =   3660
         TabIndex        =   17
         Top             =   120
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   661
         ForeColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "K"
         Picture         =   "Main.frx":0422
      End
      Begin Hangman.VBUSoftButton cmdLetter 
         Height          =   375
         Index           =   11
         Left            =   4020
         TabIndex        =   18
         Top             =   120
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   661
         ForeColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "L"
         Picture         =   "Main.frx":043E
      End
      Begin Hangman.VBUSoftButton cmdLetter 
         Height          =   375
         Index           =   12
         Left            =   4380
         TabIndex        =   19
         Top             =   120
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   661
         ForeColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "M"
         Picture         =   "Main.frx":045A
      End
      Begin Hangman.VBUSoftButton cmdLetter 
         Height          =   375
         Index           =   13
         Left            =   4740
         TabIndex        =   20
         Top             =   120
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   661
         ForeColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "N"
         Picture         =   "Main.frx":0476
      End
      Begin Hangman.VBUSoftButton cmdLetter 
         Height          =   375
         Index           =   14
         Left            =   5100
         TabIndex        =   21
         Top             =   120
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   661
         ForeColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "O"
         Picture         =   "Main.frx":0492
      End
      Begin Hangman.VBUSoftButton cmdLetter 
         Height          =   375
         Index           =   15
         Left            =   5460
         TabIndex        =   22
         Top             =   120
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   661
         ForeColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "P"
         Picture         =   "Main.frx":04AE
      End
      Begin Hangman.VBUSoftButton cmdLetter 
         Height          =   375
         Index           =   16
         Left            =   5820
         TabIndex        =   23
         Top             =   120
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   661
         ForeColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Q"
         Picture         =   "Main.frx":04CA
      End
      Begin Hangman.VBUSoftButton cmdLetter 
         Height          =   375
         Index           =   17
         Left            =   6180
         TabIndex        =   24
         Top             =   120
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   661
         ForeColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "R"
         Picture         =   "Main.frx":04E6
      End
      Begin Hangman.VBUSoftButton cmdLetter 
         Height          =   375
         Index           =   18
         Left            =   6540
         TabIndex        =   25
         Top             =   120
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   661
         ForeColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "S"
         Picture         =   "Main.frx":0502
      End
      Begin Hangman.VBUSoftButton cmdLetter 
         Height          =   375
         Index           =   19
         Left            =   6900
         TabIndex        =   26
         Top             =   120
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   661
         ForeColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "T"
         Picture         =   "Main.frx":051E
      End
      Begin Hangman.VBUSoftButton cmdLetter 
         Height          =   375
         Index           =   20
         Left            =   7260
         TabIndex        =   27
         Top             =   120
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   661
         ForeColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "U"
         Picture         =   "Main.frx":053A
      End
      Begin Hangman.VBUSoftButton cmdLetter 
         Height          =   375
         Index           =   21
         Left            =   7620
         TabIndex        =   28
         Top             =   120
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   661
         ForeColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "V"
         Picture         =   "Main.frx":0556
      End
      Begin Hangman.VBUSoftButton cmdLetter 
         Height          =   375
         Index           =   22
         Left            =   7980
         TabIndex        =   29
         Top             =   120
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   661
         ForeColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "W"
         Picture         =   "Main.frx":0572
      End
      Begin Hangman.VBUSoftButton cmdLetter 
         Height          =   375
         Index           =   23
         Left            =   8340
         TabIndex        =   30
         Top             =   120
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   661
         ForeColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "X"
         Picture         =   "Main.frx":058E
      End
      Begin Hangman.VBUSoftButton cmdLetter 
         Height          =   375
         Index           =   24
         Left            =   8700
         TabIndex        =   31
         Top             =   120
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   661
         ForeColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Y"
         Picture         =   "Main.frx":05AA
      End
      Begin Hangman.VBUSoftButton cmdLetter 
         Height          =   375
         Index           =   25
         Left            =   9060
         TabIndex        =   32
         Top             =   120
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   661
         ForeColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Z"
         Picture         =   "Main.frx":05C6
      End
   End
   Begin VB.Label lblPoints 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "$100"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   2640
      TabIndex        =   35
      Top             =   4440
      Width           =   1755
   End
   Begin VB.Label lblThisRound 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   315
      Left            =   840
      TabIndex        =   37
      Top             =   4440
      Width           =   1755
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "$100"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   120
      TabIndex        =   38
      Top             =   4740
      Width           =   4935
   End
   Begin VB.Image imgTree 
      Height          =   4335
      Left            =   1620
      Picture         =   "Main.frx":05E2
      Top             =   -1080
      Visible         =   0   'False
      Width           =   4200
   End
   Begin VB.Label lblCategory 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   5220
      TabIndex        =   3
      Top             =   240
      Width           =   4215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   2
      Top             =   -60
      Width           =   1635
   End
   Begin VB.Label Label2 
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   34
      Top             =   4200
      Width           =   1755
   End
   Begin VB.Label Label3 
      Caption         =   "This Round"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   36
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Game"
      Begin VB.Menu mnuFileNewGame 
         Caption         =   "&New Game"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuDash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit Hangman"
      End
   End
   Begin VB.Menu mnuCat 
      Caption         =   "&Categories"
      Begin VB.Menu mnuCatEdit 
         Caption         =   "&Edit Categories..."
      End
      Begin VB.Menu mnuDash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCatOptions 
         Caption         =   "Category &Options..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ********************************************
' Copyright ©Zulqarnain F. Sarani, 2006
' All Rights Reserved.
' ********************************************

Option Explicit

'Arrays to store X and Y positions of
'each letter in the phrase
Private LettersLeft() As Integer
Private LettersTop() As Integer
Private LettersLetter() As String
Private LettersShown() As String
Private LettersRed() As Boolean

'Module level variables
Private msCurrentPhrase As String
Private miPhraseLength As Integer
Private miNumberWrong As Integer
Private mbPhraseCompleted As Boolean
Private mbVowel As Boolean

'Bondaries for the phrase
Private Const cCanvasLeft = 20
Private Const cCanvasTop = 20
Private Const cCanvasWidth = 261
Private Const cCanvasHeight = 200

'Sound byte arrays
Private sndWrongLetter() As Byte
Private sndGoodLetter() As Byte

Public Property Get PhraseLettersLeft() As Integer()
    PhraseLettersLeft = LettersLeft
End Property
Public Property Let PhraseLettersLeft(lLeft() As Integer)
    LettersLeft = lLeft
End Property
Public Property Get PhraseLettersTop() As Integer()
    PhraseLettersTop = LettersTop
End Property
Public Property Let PhraseLettersTop(lTop() As Integer)
    LettersTop = lTop
End Property
Public Property Get PhraseLettersLetter() As String()
    PhraseLettersLetter = LettersLetter
End Property
Public Property Let PhraseLettersLetter(LLetter() As String)
    LettersLetter = LLetter
End Property
Public Property Get PhraseLettersShown() As String()
    PhraseLettersShown = LettersShown
End Property
Public Property Let PhraseLettersShown(LShown() As String)
    LettersShown = LShown
End Property
Public Property Get PhraseLettersRed() As Boolean()
    PhraseLettersRed = LettersRed
End Property
Public Property Let PhraseLettersRed(LRed() As Boolean)
    LettersRed = LRed
End Property
Public Property Get PhraseLength() As Integer
    PhraseLength = miPhraseLength
End Property

Private Sub ParseLetters()
    
    'This function sets the coordinates for each
    'letter in the phrase
    
    Dim iCounter As Integer
    Dim iLastLetterLeft As Integer
    Dim iLastLetterTop As Integer
    Dim iFillerWidth As Integer
    Dim iFillerHeight As Integer
    Dim sPhrase As String
    
    'iFillerWidth and iFillerHeight is used to
    'space the letters apart.
    iFillerWidth = 2
    iFillerHeight = 3
    
    miPhraseLength = Len(msCurrentPhrase)
    
    ReDim LettersLeft(miPhraseLength)
    ReDim LettersTop(miPhraseLength)
    ReDim LettersLetter(miPhraseLength)
    ReDim LettersShown(miPhraseLength)
    ReDim LettersRed(miPhraseLength)
    
    iLastLetterLeft = cCanvasLeft
    iLastLetterTop = cCanvasTop + TextHeight("X")
    
    For iCounter = 1 To miPhraseLength
        If Mid$(msCurrentPhrase, iCounter, 1) = " " Or Mid$(msCurrentPhrase, iCounter, 1) = "-" And iLastLetterLeft > cCanvasLeft Then
            If iLastLetterLeft + TextWidth(Space(GetNextWordLength(iCounter))) + (iFillerWidth * GetNextWordLength(iCounter)) > cCanvasWidth Then
                LettersLeft(iCounter) = cCanvasLeft
                LettersTop(iCounter) = iLastLetterTop + iFillerHeight + TextHeight("X")
                iLastLetterTop = iLastLetterTop + iFillerHeight + TextHeight("X")
                iLastLetterLeft = cCanvasLeft
            Else
                LettersLeft(iCounter) = iLastLetterLeft + iFillerWidth
                LettersTop(iCounter) = iLastLetterTop + iFillerHeight
                iLastLetterLeft = iLastLetterLeft + iFillerWidth + TextWidth("X")
            End If
        Else
            LettersLeft(iCounter) = iLastLetterLeft + iFillerWidth
            LettersTop(iCounter) = iLastLetterTop + iFillerHeight
            iLastLetterLeft = iLastLetterLeft + iFillerWidth + TextWidth("X")
        End If
        LettersLetter(iCounter) = UCase$(Mid$(msCurrentPhrase, iCounter, 1))
        If (LCase$(LettersLetter(iCounter)) < "a" Or LCase$(LettersLetter(iCounter)) > "z") Then
            If LettersLetter(iCounter) <> " " Then
                LettersShown(iCounter) = LettersLetter(iCounter)
            End If
        End If
    Next
    
End Sub
Private Function GetNextWordLength(iBeginWord) As Integer
    
    'This function returns the width of the next word.
    'This is used to figure out where to word wrap.
    
    Dim iNextSpace As Integer
    Dim iNextHyphen As Integer
    
    iNextSpace = InStr(iBeginWord + 1, msCurrentPhrase, " ")
    iNextHyphen = InStr(iBeginWord + 1, msCurrentPhrase, "-")
    
    If iNextSpace = 0 And iNextHyphen = 0 Then
        If iBeginWord < Len(msCurrentPhrase) Then
            GetNextWordLength = Len(msCurrentPhrase) - iBeginWord
        Else
            GetNextWordLength = 0
        End If
    ElseIf iNextSpace = 0 Then
        GetNextWordLength = iNextHyphen - iBeginWord
    ElseIf iNextHyphen = 0 Then
        GetNextWordLength = iNextSpace - iBeginWord + 1
    Else
        If iNextSpace < iNextHyphen Then
            GetNextWordLength = iNextSpace - iBeginWord
        Else
            GetNextWordLength = iNextHyphen - iBeginWord + 1
        End If
    End If
End Function
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
Private Function GuessLetter(sLetter As String) As Boolean
    
    'This function get the letter that was guessed and returns
    'True if the letter exists in the phrase and returns
    'False if it does not exist.
    
    Dim iCounter As Integer
    Dim bLetterFound As Boolean
    Dim iNumLetters As Integer
    Dim lReturn As Long
    
    'Set this initially to true. The for...loop will
    'Check to see if it is false.
    mbPhraseCompleted = True
    
    For iCounter = 1 To miPhraseLength
        If LCase$(LettersLetter(iCounter)) = LCase$(sLetter) Then
            bLetterFound = True
            If Not mbVowel Then
                glPoints = glPoints + gcCorrectGuess
                glCurrentRound = glCurrentRound + gcCorrectGuess
            End If
            iNumLetters = iNumLetters + 1
            LettersShown(iCounter) = LettersLetter(iCounter)
            LettersRed(iCounter) = False
        End If
        If LettersShown(iCounter) = "" And (LettersLetter(iCounter) >= "A" And LettersLetter(iCounter) <= "Z") Then
            mbPhraseCompleted = False
        End If
    Next
    
    'If the letter has been found then redraw
    'the phrase.
    If bLetterFound Then
        lReturn = sndPlaySound(sndGoodLetter(0), SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY)
        DrawPhrase
    End If
    If mbVowel Then
        lblStatus.Caption = "No points received for vowels."
    Else
        If iNumLetters > 0 Then
            lblStatus.Caption = CStr(iNumLetters) & " Letters found! You get " & CStr(iNumLetters * gcCorrectGuess) & " points."
        End If
    End If
    GuessLetter = bLetterFound
    
End Function

Private Sub cmdLetter_Click(Index As Integer)
    
    Dim bLetterFound As Boolean
    Dim bGameOver As Boolean
    Dim iCounter As Integer
    Dim lReturn As Long
    
    Select Case cmdLetter(Index).Caption
        Case "A", "E", "I", "O", "U"
            mbVowel = True
        Case Else
            mbVowel = False
    End Select
    
    'Check to see if the phrase contains that letter
    bLetterFound = GuessLetter(cmdLetter(Index).Caption)
    
    SetBonusLabel
    If Not bLetterFound Then
        'The letter does not exist. Draw a body part.
        bGameOver = DrawBodyPart
        lReturn = sndPlaySound(sndWrongLetter(0), SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY)
        If Not mbVowel Then
            glPoints = glPoints - gcWrongGuess
            glCurrentRound = glCurrentRound - gcWrongGuess
            lblStatus.Caption = "Letter not found. -" & gcWrongGuess & " points"
            SetPointsLabel
        End If
        'If the game is over then inform the user and start over.
        If bGameOver Then
            'Show the rest of the letters
            For iCounter = 1 To miPhraseLength
                If LettersShown(iCounter) = "" Then
                    LettersRed(iCounter) = True
                End If
                LettersShown(iCounter) = LettersLetter(iCounter)
            Next
            DrawPhrase
            DoEvents
            glPoints = glPoints - gcGameLost
            glCurrentRound = glCurrentRound - gcGameLost
            SetPointsLabel
            cmdSolve.Caption = "New Game {F2}"
            lblGameOver.Visible = True
            lblNewGame.Visible = True
            Exit Sub
        End If
    End If
    
    SetPointsLabel
    
    'Check to see if the phrase has been completed. This
    'variable will have been set by the GuessLetter function.
    If mbPhraseCompleted Then
        glPoints = glPoints + gcGameWon
        glCurrentRound = glCurrentRound + gcGameWon
        lblStatus = "Way to GO! You get an extra 200 points for solving the puzzle!"
        SetPointsLabel
        DrawEndGame
        Exit Sub
    End If
    
    cmdLetter(Index).ForeColor = vbGrayText
    cmdLetter(Index).Enabled = False
    
End Sub
Private Function NumLettersLeft() As Integer
    Dim iCounter As Integer
    Dim iLettersLeft As Integer
    
    For iCounter = 1 To UBound(LettersShown)
        If LettersShown(iCounter) = "" And (LCase$(LettersLetter(iCounter)) >= "a" And LCase$(LettersLetter(iCounter)) <= "z") Then
            iLettersLeft = iLettersLeft + 1
        End If
    Next
    
    NumLettersLeft = iLettersLeft
End Function
Private Sub SetBonusLabel()
    Dim iCount As Integer
    
    iCount = NumLettersLeft
    
    cmdSolve.Caption = "&Solve puzzle now for " & CStr(iCount * gcBonus) & " points"
End Sub
Private Sub SetThisRoundLabel()
    lblThisRound = " " & CStr(glCurrentRound)
    
End Sub

Private Sub DrawEndGame()
    Dim iCounter As Integer
        
    
    Set Picture1.Picture = Nothing
    cmdSolve.Caption = "New Game {F2}"
    shpRope.Visible = False
    shpHead.Visible = False
    For iCounter = 0 To 4
        lnBody(iCounter).Visible = False
    Next
    Set Animate = New CAnimate
    With Animate
        .Parent = Picture1.hWnd
        .Move 10, 10
        .AutoPlay = True
        .Center = False
        .Transparent = False
        .ResourceID = 102
    End With
    
    lblNewGame.Visible = True
    For iCounter = 0 To 25
        cmdLetter(iCounter).Enabled = False
    Next
End Sub

Private Sub cmdSolve_Click()
    Dim iCounter As Integer
    Dim bGameLost As Boolean
    Dim lBonus As Long
    
    If cmdSolve.Caption <> "New Game {F2}" Then
        lBonus = NumLettersLeft * gcBonus
        cmdSolve.Enabled = False
        gbSolveCanceled = False
        
        frmGuess.Show vbModal, Me
        
        Set frmGuess = Nothing
        
        If gbSolveCanceled Then
            cmdSolve.Enabled = True
            Exit Sub
        End If
        
        DrawPhrase
        mbPhraseCompleted = True
        
        For iCounter = 1 To miPhraseLength
            If LettersRed(iCounter) Then
                bGameLost = True
                Exit For
            End If
        Next
        
        cmdSolve.Enabled = True
        cmdSolve.Caption = "New Game {F2}"
        If bGameLost Then
            glPoints = glPoints - lBonus
            glCurrentRound = glCurrentRound - lBonus
            lblStatus.Caption = "Unable to complete the Puzzle -" & lBonus & " points."
            SetPointsLabel
            lblGameOver.Visible = True
            lblNewGame.Visible = True
        Else
            glPoints = glPoints + lBonus
            glCurrentRound = glCurrentRound + lBonus
            glPoints = glPoints + gcGameWon
            glCurrentRound = glCurrentRound + gcGameWon
            lblStatus.Caption = "Way to GO! You get " & lBonus & " points and an extra 200 points for solving the puzzle!"
            SetPointsLabel
            DrawEndGame
        End If
        Me.SetFocus
        Me.Show
        Me.ZOrder 0
        cmdSolve.SetFocus
    Else
        mnuFileNewGame_Click
    End If
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
        If cmdLetter(Asc(sLetter) - 65).Enabled Then
            cmdLetter_Click Asc(sLetter) - 65
        End If
    End If

End Sub
Private Function DrawBodyPart() As Boolean
    
    'Draw the body part
    If miNumberWrong = 0 Then
        shpHead.Visible = True
    Else
        lnBody(miNumberWrong - 1).Visible = True
    End If
    
    'Increment the number of wrong guesses
    miNumberWrong = miNumberWrong + 1
    
    'If we just drew the last body part then we want to
    'return a True value to signify the game is over.
    If miNumberWrong = 6 Then
        DrawBodyPart = True
    Else
        DrawBodyPart = False
    End If
    
End Function
Private Sub SetPointsLabel()
    lblPoints = " " & CStr(glPoints)
    SetThisRoundLabel
End Sub
Private Sub Form_Load()
    Dim iCounter As Integer
    Dim sFile As String
    
    'Seed the random number generator
    Randomize
        
        
    App.HelpFile = App.Path & "\Hangman.hlp"
    Set Picture1.Picture = imgTree.Picture
    
    sndWrongLetter = LoadResData(105, "WAVE")
    sndGoodLetter = LoadResData(104, "WAVE")
    
    lstFiles.Path = App.Path
    lstFiles.Pattern = "*.hm"
    
    glPoints = gcInitialCash
    glCurrentRound = glPoints
    lblStatus.Caption = ""
    SetPointsLabel
    
    Me.Caption = "Hangman"
    'Load the Categories
    
    For iCounter = 0 To lstFiles.ListCount - 1
        sFile = lstFiles.List(iCounter)
        gclsCategories.Add sFile, Left$(sFile, Len(sFile) - 3)
    Next
    
    gclsJukebox.ResetList
    mnuFileNewGame_Click
    
End Sub


Private Sub mnuCatEdit_Click()
    frmEditCategories.Show vbModal, Me
    gclsJukebox.ResetList
End Sub

Private Sub mnuCatOptions_Click()
    frmCategoryOptions.Show vbModal, Me
    gclsJukebox.ResetList
End Sub

Private Sub mnuFileExit_Click()
    
    Unload Me
    
End Sub

Private Sub mnuFileNewGame_Click()
        
    On Error Resume Next
        
    Dim iCounter As Integer
    
    Set Animate = Nothing
    
    shpRope.Visible = True
    glCurrentRound = 0
    SetThisRoundLabel
    Set Picture1.Picture = imgTree.Picture
    lblGameOver.Visible = False
    lblNewGame.Visible = False
    lblStatus.Caption = ""
    'Reset each of the letters
    For iCounter = 0 To 25
        cmdLetter(iCounter).Enabled = True
        cmdLetter(iCounter).ForeColor = &HC0&
    Next
    
    'Clear the window
    Cls
    
    'Reset the number of wrong guesses
    miNumberWrong = 0
    
    'Hide all of the body parts
    shpHead.Visible = False
    For iCounter = 0 To 4
        lnBody(iCounter).Visible = False
    Next
    
    'Get a new phrase and draw it in the window
    'GetNewCategory
    'msCurrentPhrase = GetNewPhrase
    gclsJukebox.GetNewPhrase
    msCurrentPhrase = gclsJukebox.Phrase
    lblCategory = gclsJukebox.Category
    ParseLetters
    DrawPhrase
    SetBonusLabel
    cmdSolve.Enabled = True
    cmdSolve.SetFocus

End Sub
Private Sub GetNewCategory()
    Dim i As Integer
    Dim sFile As String
    
    i = Int(lstFiles.ListCount * Rnd)
    lstFiles.ListIndex = i
    
    sFile = lstFiles.List(lstFiles.ListIndex)
    lblCategory = Left$(sFile, Len(sFile) - 3)
    
End Sub
Public Function GetNewPhrase() As String
    Dim mPhrases() As String
    Dim iFnum As Integer
    Dim sFile As String
    Dim sLine As String
    Dim iPhraseCount As Integer
    
    sFile = lblCategory & ".hm"
    iFnum = FreeFile
    
    Open lstFiles.Path & "\" & sFile For Input As iFnum
    
    iPhraseCount = -1
    Do While EOF(iFnum) = False
        Line Input #iFnum, sLine
        iPhraseCount = iPhraseCount + 1
        ReDim Preserve mPhrases(iPhraseCount)
        mPhrases(iPhraseCount) = sLine
    Loop
    
    Close iFnum
    
    GetNewPhrase = mPhrases(Int(iPhraseCount * Rnd))
    
End Function

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub
