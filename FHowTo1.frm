VERSION 5.00
Begin VB.Form FHowTo1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " How to - The Rule"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7095
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrTransparent 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4080
      Top             =   3120
   End
   Begin Sudoku.Button cmdCommand 
      Height          =   390
      Index           =   0
      Left            =   4740
      Top             =   3180
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   688
      Caption         =   "Close"
      Enabled         =   -1  'True
   End
   Begin Sudoku.Button cmdCommand 
      Height          =   390
      Index           =   1
      Left            =   5880
      Top             =   3180
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   688
      Caption         =   "Next"
      Enabled         =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"FHowTo1.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   2340
      Width           =   6735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"FHowTo1.frx":00C0
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   1
      Left            =   2100
      TabIndex        =   1
      Top             =   1020
      Width           =   4935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"FHowTo1.frx":019F
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Index           =   0
      Left            =   2100
      TabIndex        =   0
      Top             =   120
      Width           =   4875
   End
   Begin VB.Image Image1 
      Height          =   1680
      Left            =   180
      Picture         =   "FHowTo1.frx":0239
      Top             =   240
      Width           =   1680
   End
End
Attribute VB_Name = "FHowTo1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCommand_Click(Index As Integer)
  Tag = Me.cmdCommand(Index).Caption
  FadeOut Me
End Sub

Private Sub tmrTransparent_Timer()
  FadeIn Me
End Sub
