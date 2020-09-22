VERSION 5.00
Begin VB.Form FLayer 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1440
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1440
   ScaleWidth      =   1680
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrShow 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   960
      Top             =   300
   End
   Begin VB.Timer tmrHide 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   180
      Top             =   240
   End
End
Attribute VB_Name = "FLayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub tmrHide_Timer()
Static I As Integer

  I = I + 2
  If I < 225 Then
    MakeTransparent Me.hWnd, 225 - I
  Else
    Me.tmrHide.Enabled = False
    Unload Me
  End If
End Sub

Private Sub tmrShow_Timer()
Static I As Integer

  I = I + 2
  If I < 255 Then
    MakeTransparent Me.hWnd, I
  Else
    Me.tmrShow.Enabled = False
  End If
End Sub
