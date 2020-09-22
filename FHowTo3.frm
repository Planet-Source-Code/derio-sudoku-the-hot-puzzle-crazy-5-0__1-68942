VERSION 5.00
Begin VB.Form FHowTo3 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " How to - Penciling"
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
      Left            =   3000
      Top             =   3180
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
   Begin Sudoku.Button cmdCommand 
      Height          =   390
      Index           =   2
      Left            =   3600
      Top             =   3180
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   688
      Caption         =   "Prev"
      Enabled         =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "For each box, You can have up to five numbers as penciling notes."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   2
      Left            =   2820
      TabIndex        =   2
      Top             =   2460
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Right-Click one of them (i.e. 8), and the chosen one will appear as a small number on top of that box.           "
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
      Index           =   1
      Left            =   2820
      TabIndex        =   1
      Top             =   1620
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Left-Click the empty box, and the list option appears."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   2820
      TabIndex        =   0
      Top             =   180
      Width           =   3855
   End
   Begin VB.Image Image1 
      Height          =   1110
      Index           =   3
      Left            =   1440
      Picture         =   "FHowTo3.frx":0000
      Top             =   1620
      Width           =   1125
   End
   Begin VB.Image Image1 
      Height          =   1110
      Index           =   2
      Left            =   120
      Picture         =   "FHowTo3.frx":422A
      Top             =   1620
      Width           =   1125
   End
   Begin VB.Image Image1 
      Height          =   1110
      Index           =   1
      Left            =   1440
      Picture         =   "FHowTo3.frx":8454
      Top             =   120
      Width           =   1125
   End
   Begin VB.Image Image1 
      Height          =   1110
      Index           =   0
      Left            =   120
      Picture         =   "FHowTo3.frx":C67E
      Top             =   120
      Width           =   1125
   End
End
Attribute VB_Name = "FHowTo3"
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
