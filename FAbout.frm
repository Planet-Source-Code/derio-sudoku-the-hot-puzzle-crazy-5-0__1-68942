VERSION 5.00
Begin VB.Form FAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2520
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   3825
   ControlBox      =   0   'False
   Icon            =   "FAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   3825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrTransparent 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   180
      Top             =   1860
   End
   Begin Sudoku.Button cmdOK 
      Height          =   390
      Left            =   2580
      Top             =   1980
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   688
      Caption         =   "OK"
      Enabled         =   0   'False
   End
   Begin VB.Timer tmrHide 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   180
      Top             =   1380
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1680
      Left            =   720
      TabIndex        =   0
      Top             =   60
      Width           =   2955
      Begin VB.Label lblMotto 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "The Hot Puzzle Craze"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   1920
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sudoku 5.0"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   1
         Left            =   135
         TabIndex        =   2
         Top             =   165
         Width           =   2640
      End
      Begin VB.Label lblCopyright 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright (C) Derio 2006 - 2007"
         Height          =   195
         Left            =   540
         TabIndex        =   1
         Top             =   1380
         Width           =   2220
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sudoku 5.0"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000010&
         Height          =   645
         Index           =   0
         Left            =   165
         TabIndex        =   3
         Top             =   180
         Width           =   2640
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "FAbout.frx":030A
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "FAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'***********************************************
'* Title: FAbout                               *
'* Stamp:                                      *
'* Auth : Derio                                *
'* Desc : Show the Information about My Sudoku *
'***********************************************

Private Sub cmdOK_Click()
  FadeOut Me
End Sub

Private Sub tmrHide_Timer()
'** Unload form after ?? second

  FadeOut Me
  Unload Me
End Sub

Private Sub tmrTransparent_Timer()
  FadeIn Me
End Sub
