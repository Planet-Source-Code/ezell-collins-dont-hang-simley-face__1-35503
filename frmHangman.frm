VERSION 5.00
Begin VB.Form frmHangman 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hang Man"
   ClientHeight    =   5460
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   8310
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHangman.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   8310
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frameLetters 
      Caption         =   "Remaing Letters"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   1095
      Left            =   720
      TabIndex        =   18
      Top             =   3840
      Width           =   3375
      Begin VB.Label lblLetter 
         Alignment       =   2  'Center
         Caption         =   "Z"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   25
         Left            =   3000
         TabIndex        =   44
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lblLetter 
         Alignment       =   2  'Center
         Caption         =   "Y"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   24
         Left            =   2760
         TabIndex        =   43
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lblLetter 
         Alignment       =   2  'Center
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   23
         Left            =   2520
         TabIndex        =   42
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lblLetter 
         Alignment       =   2  'Center
         Caption         =   "W"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   22
         Left            =   2280
         TabIndex        =   41
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lblLetter 
         Alignment       =   2  'Center
         Caption         =   "V"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   21
         Left            =   2040
         TabIndex        =   40
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lblLetter 
         Alignment       =   2  'Center
         Caption         =   "U"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   20
         Left            =   1800
         TabIndex        =   39
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lblLetter 
         Alignment       =   2  'Center
         Caption         =   "T"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   19
         Left            =   1560
         TabIndex        =   38
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lblLetter 
         Alignment       =   2  'Center
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   18
         Left            =   1320
         TabIndex        =   37
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lblLetter 
         Alignment       =   2  'Center
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   17
         Left            =   1080
         TabIndex        =   36
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lblLetter 
         Alignment       =   2  'Center
         Caption         =   "Q"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   16
         Left            =   840
         TabIndex        =   35
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lblLetter 
         Alignment       =   2  'Center
         Caption         =   "P"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   15
         Left            =   600
         TabIndex        =   34
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lblLetter 
         Alignment       =   2  'Center
         Caption         =   "O"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   14
         Left            =   360
         TabIndex        =   33
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lblLetter 
         Alignment       =   2  'Center
         Caption         =   "N"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   13
         Left            =   120
         TabIndex        =   32
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lblLetter 
         Alignment       =   2  'Center
         Caption         =   "M"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   12
         Left            =   3000
         TabIndex        =   31
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblLetter 
         Alignment       =   2  'Center
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   11
         Left            =   2760
         TabIndex        =   30
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblLetter 
         Alignment       =   2  'Center
         Caption         =   "K"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   2520
         TabIndex        =   29
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblLetter 
         Alignment       =   2  'Center
         Caption         =   "J"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   2280
         TabIndex        =   28
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblLetter 
         Alignment       =   2  'Center
         Caption         =   "I"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   2040
         TabIndex        =   27
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblLetter 
         Alignment       =   2  'Center
         Caption         =   "H"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   1800
         TabIndex        =   26
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblLetter 
         Alignment       =   2  'Center
         Caption         =   "G"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   1560
         TabIndex        =   25
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblLetter 
         Alignment       =   2  'Center
         Caption         =   "F"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   1320
         TabIndex        =   24
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblLetter 
         Alignment       =   2  'Center
         Caption         =   "E"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   1080
         TabIndex        =   23
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblLetter 
         Alignment       =   2  'Center
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   840
         TabIndex        =   22
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblLetter 
         Alignment       =   2  'Center
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   600
         TabIndex        =   21
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblLetter 
         Alignment       =   2  'Center
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   20
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblLetter 
         Alignment       =   2  'Center
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.Frame frameWord 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   1200
      Width           =   4935
      Begin VB.Label lblLetterOfWord 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   14
         Left            =   4440
         TabIndex        =   15
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblLetterOfWord 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   13
         Left            =   4137
         TabIndex        =   14
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblLetterOfWord 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   12
         Left            =   3828
         TabIndex        =   13
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblLetterOfWord 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   11
         Left            =   3519
         TabIndex        =   12
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblLetterOfWord 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   3210
         TabIndex        =   11
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblLetterOfWord 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   2901
         TabIndex        =   10
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblLetterOfWord 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   2592
         TabIndex        =   9
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblLetterOfWord 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   2283
         TabIndex        =   8
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblLetterOfWord 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   1974
         TabIndex        =   7
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblLetterOfWord 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   1665
         TabIndex        =   6
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblLetterOfWord 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   1356
         TabIndex        =   5
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblLetterOfWord 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   1047
         TabIndex        =   4
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblLetterOfWord 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   738
         TabIndex        =   3
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblLetterOfWord 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   429
         TabIndex        =   2
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblLetterOfWord 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Label lblHelp 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   47
      Top             =   480
      Width           =   3375
   End
   Begin VB.Label lblGueses 
      BackColor       =   &H80000012&
      Caption         =   "Chances  Remaining"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   375
      Left            =   5760
      TabIndex        =   16
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label lblRemainingChances 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   7680
      TabIndex        =   17
      Top             =   4440
      Width           =   375
   End
   Begin VB.Shape shpPlateform 
      BorderColor     =   &H80000001&
      BorderWidth     =   4
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   5640
      Top             =   4320
      Width           =   2535
   End
   Begin VB.Label lblTalk 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "I Knew you could do it "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   46
      Top             =   240
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      FillColor       =   &H80000005&
      Height          =   255
      Left            =   5520
      Shape           =   2  'Oval
      Top             =   1320
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      FillColor       =   &H80000005&
      Height          =   135
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   1800
      Width           =   135
   End
   Begin VB.Image imgHead 
      Appearance      =   0  'Flat
      Height          =   675
      Left            =   5950
      Picture         =   "frmHangman.frx":0442
      Stretch         =   -1  'True
      Top             =   1680
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label lblOutCome 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   855
      Left            =   960
      TabIndex        =   45
      Top             =   2520
      Width           =   2895
   End
   Begin VB.Line lineRightFoot 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   5880
      X2              =   5640
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line lineLeftFoot 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   6600
      X2              =   6840
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line lineLeftLeg 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   6600
      X2              =   6240
      Y1              =   3840
      Y2              =   3120
   End
   Begin VB.Line lineRightLeg 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   6240
      X2              =   5880
      Y1              =   3120
      Y2              =   3840
   End
   Begin VB.Line lineLeftArm 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   6240
      X2              =   6840
      Y1              =   2520
      Y2              =   2880
   End
   Begin VB.Line lineRightArm 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   6240
      X2              =   5640
      Y1              =   2520
      Y2              =   2880
   End
   Begin VB.Line lineBody 
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   6240
      X2              =   6240
      Y1              =   2280
      Y2              =   3120
   End
   Begin VB.Line Line4 
      BorderWidth     =   3
      X1              =   6240
      X2              =   6240
      Y1              =   1440
      Y2              =   1680
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   7080
      X2              =   7680
      Y1              =   1440
      Y2              =   2040
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   7680
      X2              =   6240
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   7680
      X2              =   7680
      Y1              =   4440
      Y2              =   1440
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H80000005&
      Height          =   1080
      Left            =   3960
      Shape           =   2  'Oval
      Top             =   0
      Width           =   1935
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuFileNewGame 
      Caption         =   "&New Game"
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmHangman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sWordsList(600) As String
Private iNmbrOfWords As Integer
Private bGameOver As Boolean
Private sWord As String
Private iChancesRemaining As Integer
Private Sub Form_Load()

    '-- Load Words into Array from dat file
    LoadWords
    
    StartNewGame
        
End Sub
Public Sub LoadWords()

    Dim iFNum As Integer
    
On Err GoTo LocalError
    
    iFNum = FreeFile
    iNmbrOfWords = 0
    
    Open App.Path & "/" & "Data" & "/" & "WordList.dat" For Input As iFNum
    
    Do While Not EOF(iFNum)
        iNmbrOfWords = iNmbrOfWords + 1
        Input #iFNum, sWordsList(iNmbrOfWords)
    Loop
    
    Close #iFNum
    
Exit Sub

LocalError:

    MsgBox Err.Number & "Unable to load Wordlist.dat, check to see if file exist" & _
        vbCrLf & "AppPath/Data/Wordlist.dat", vbOKOnly + vbCritical
    Resume Next
    
End Sub
Private Sub StartNewGame()

    InilizeLetterPicks
    ClearWordDisplay
    SetGameVariables
    HideBalloon
    HideHangman
    sWord = RandomWord
    lblHelp = "Type a letter or click on a letter in remaining letter box."
    DisplayBlanks

End Sub
Private Sub InilizeLetterPicks()

    Dim iLetter As Integer

    '-- Display Letters from A to Z
    For iLetter = 0 To 25
        lblLetter(iLetter).Visible = True
        lblLetter(iLetter) = Chr$(iLetter + 65)
    Next iLetter
    
End Sub
Private Sub ClearWordDisplay()
    
    Dim iLetter As Integer
    
    '-- Clear Computer word pick , erase each letter in word display
    For iLetter = 0 To 14
        lblLetterOfWord(iLetter) = ""
    Next iLetter
    
End Sub
Public Sub SetGameVariables()
    
    lblOutCome = ""
    iChancesRemaining = 7
    lblRemainingChances = iChancesRemaining
    bGameOver = False
    
End Sub
Private Sub HideBalloon()

    Shape1.Visible = False
    Shape2.Visible = False
    Shape3.Visible = False
    lblTalk = ""
    lblTalk.Visible = False
    
End Sub

Private Sub HideHangman()

    imgHead.Picture = LoadPicture(App.Path + "/Images/FACE01.ICO")
    imgHead.Visible = False
    lineBody.Visible = False
    lineLeftArm.Visible = False
    lineRightArm.Visible = False
    lineLeftLeg.Visible = False
    lineRightLeg.Visible = False
    lineLeftFoot.Visible = False
    lineRightFoot.Visible = False
    
End Sub
Private Function RandomWord() As String
'-- Choose A Random Word Form WordList Array

    Dim iComputerPick As Integer
    
    Randomize
    'iComputerPick = 23  'Debug
    iComputerPick = Int((iNmbrOfWords - 1 + 1) * Rnd + 1)

    RandomWord = sWordsList(iComputerPick)

End Function
Private Sub DisplayBlanks()
'-- Display A Underscore for each letter in word, Diplay a
'-- space if letter is a space

    Dim iLetter As Integer
    Dim iNumberOfLetters As Integer
    
    iNumberOfLetters = Len(sWord)
    
    For iLetter = 0 To iNumberOfLetters - 1
        If Asc(Mid$(sWord, iLetter + 1, 1)) = vbKeySpace Then
            lblLetterOfWord(iLetter) = " "
        Else
            lblLetterOfWord(iLetter) = "_"
        End If
    Next iLetter

End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
'-- If keyPress is a letter from a to z Upper or Lower case then pass Uppercase
'-- Index value to lblLetter_Click

    If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Then
    
        Select Case KeyAscii
            Case 78 And (bGameOver = True)
                StartNewGame
            Case 110 And (bGameOver = True)
                StartNewGame
            Case Else
                lblLetter_Click ((KeyAscii And Not 32) - 65)    '(KeyAscii And Not 32) Convert to Uppercase
        End Select
        
    End If

End Sub


Private Sub lblLetter_Click(Index As Integer)
'-- Manage Game base on key entered by player

    Dim iLetterPicked As Integer
    Dim iLettersFound As Integer
    Dim sReply As String
    
    If lblLetter(Index).Visible = False Then Exit Sub
    
    iLetterPicked = Index + 65          'Convert index to Ascii value
    lblLetter(Index).Visible = False    'Erase letter player picked
        
    iLettersFound = FindLetters(iLetterPicked)
    
    If iLettersFound = 0 Then
        iChancesRemaining = iChancesRemaining - 1
        ShowBodyPart
        lblRemainingChances = iChancesRemaining
    Else
        bGameOver = DoPlayerWin         'If Player win then game is over
        If bGameOver Then
            ShowPlayerWins
        End If
    End If
    
    If iChancesRemaining = 0 Then
        bGameOver = True
        ShowPlayerLose
    End If
        
End Sub
Private Function FindLetters(sLetterPick) As Integer

    Dim iLettersFound As Integer
    Dim iLetter As Integer
    
    For iLetter = 1 To Len(sWord)
        If lblLetterOfWord(iLetter - 1) = "_" And Mid$(sWord, iLetter, 1) = Chr$(sLetterPick) Then
            lblLetterOfWord(iLetter - 1) = Mid$(sWord, iLetter, 1)
            iLettersFound = iLettersFound + 1
        End If
    Next iLetter
    
    FindLetters = iLettersFound
        
End Function
Private Function DoPlayerWin() As Boolean

    Dim iIndex As Integer
    
    DoPlayerWin = True
    For iIndex = 1 To Len(Trim(sWord))
        If lblLetterOfWord(iIndex - 1) <> Mid$(sWord, iIndex, 1) Then
            DoPlayerWin = False
        End If
    Next iIndex
    

End Function
Private Sub ShowBodyPart()

    Select Case iChancesRemaining
        Case 6
            imgHead.Visible = True
            ShowBalloon
            ShowHelpMeMessage
        Case 5
            lineBody.Visible = True
            ShowHelpMeMessage
        Case 4
            lineLeftArm.Visible = True
            ShowHelpMeMessage
        Case 3
            lineRightArm.Visible = True
            ShowHelpMeMessage
        Case 2
            lineLeftLeg.Visible = True
            ShowHelpMeMessage
        Case 1
            lineRightLeg.Visible = True
            ShowHelpMeMessage
        Case 0
            lineRightFoot.Visible = True
            lineLeftFoot.Visible = True
            HideBalloon
    End Select
    
End Sub
Private Sub ShowBalloon()
    
    Shape1.Visible = True
    Shape2.Visible = True
    Shape3.Visible = True
    lblTalk = ""
    lblTalk.Visible = True
    
End Sub

Private Sub ShowPlayerWins()
    
    Dim iIndex As Integer
    
    ShowWinMessage
    
    lblOutCome = "YOU WIN!!!"
    lblHelp = "Press (N) to start a new game"
    imgHead = LoadPicture(App.Path & "\Images\FACE02.ICO")
    
    For iIndex = 0 To 25
        lblLetter(iIndex).Visible = False
    Next iIndex
End Sub
Private Sub ShowPlayerLose()
    
    Dim iIndex As Integer
    
    lblOutCome = "YOU LOSE!!"
    lblHelp = "Press (N) to start a new game"
    imgHead = LoadPicture(App.Path & "\Images\FACE04.ICO")
    
    For iIndex = 0 To 25
        lblLetter(iIndex).Visible = False
    Next iIndex
        
End Sub
Private Sub mnuAbout_Click()
    
    frmAbout.Show vbModal
    
End Sub

Private Sub mnuFileNewGame_Click()
    Dim iAnwser As Integer
    
    If Not bGameOver Then
        iAnwser = MsgBox("Are you sure you want to quit the present game" & vbCrLf & "and start a new game?", vbYesNo + vbInformation, "Continue")
        If iAnwser = vbYes Then
            StartNewGame
        End If
    Else
        StartNewGame
    End If
    
End Sub
Private Sub mnuFileExit_Click()

    Unload Me
    
End Sub
Private Sub ShowWinMessage(Optional ByVal iNumber As Integer)

    Dim sMessage(5) As String
    
    sMessage(1) = "Thanks Partner"
    sMessage(2) = "Alrighty Then"
    sMessage(3) = "I was not worried"
    sMessage(4) = "That was close"
    sMessage(5) = "Bless You"
    
    If iNumber < 1 Or iNumber > 5 Then
        iNumber = Int((5 * Rnd) + 1)
    End If
    
    lblTalk = sMessage(iNumber)
End Sub

Private Sub ShowHelpMeMessage(Optional ByVal iNumber As Integer)
    
    Dim sMessage(5) As String
    
    sMessage(1) = "Help Me!"
    sMessage(2) = "Save Me!"
    sMessage(3) = "Please Save Me!"
    sMessage(4) = "Try Harder Please!"
    sMessage(5) = "I'm At the end of my rope!"
    
    If iNumber < 1 Or iNumber > 5 Then
        iNumber = Int((5 * Rnd) + 1)
    End If
    
    lblTalk = sMessage(iNumber)
End Sub
