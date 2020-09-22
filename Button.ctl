VERSION 5.00
Begin VB.UserControl Button 
   AutoRedraw      =   -1  'True
   BackStyle       =   0  'Transparent
   CanGetFocus     =   0   'False
   ClientHeight    =   405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1080
   MaskColor       =   &H00C0E0FF&
   MaskPicture     =   "Button.ctx":0000
   MouseIcon       =   "Button.ctx":1632
   MousePointer    =   99  'Custom
   Picture         =   "Button.ctx":1784
   ScaleHeight     =   405
   ScaleWidth      =   1080
   ToolboxBitmap   =   "Button.ctx":2DB6
   Begin VB.Timer tmrHiLight 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   360
      Top             =   600
   End
   Begin VB.Image imgLib 
      Height          =   390
      Index           =   2
      Left            =   0
      Picture         =   "Button.ctx":30C8
      Top             =   1740
      Width           =   1080
   End
   Begin VB.Image imgLib 
      Height          =   390
      Index           =   1
      Left            =   0
      Picture         =   "Button.ctx":46FA
      Top             =   1260
      Width           =   1080
   End
   Begin VB.Image imgLib 
      Height          =   390
      Index           =   0
      Left            =   0
      Picture         =   "Button.ctx":5D2C
      Top             =   780
      Width           =   1080
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Caption"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   240
      Left            =   45
      MouseIcon       =   "Button.ctx":735E
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   60
      Width           =   1005
   End
End
Attribute VB_Name = "Button"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private MouseOver As Boolean

Public Event Click()
Public Event MouseIn()
Public Event MouseOut()


Private Sub lblCaption_Click()
  If Me.Enabled Then RaiseEvent Click
End Sub

Private Sub lblCaption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Me.Enabled Then
    If Not MouseOver Then HiLight
  End If
End Sub

Private Sub tmrHiLight_Timer()
Dim Cur As POINT

  GetCursorPos Cur
  Cur.X = Cur.X - (Extender.Parent.Left + (Extender.Parent.Width - Extender.Parent.ScaleWidth) / 2) \ Screen.TwipsPerPixelX
  Cur.Y = Cur.Y - (Extender.Parent.Top + Extender.Parent.Height - Extender.Parent.ScaleHeight - 30) \ Screen.TwipsPerPixelY
  
  Cur.X = Cur.X * Screen.TwipsPerPixelX - Extender.Left
  Cur.Y = Cur.Y * Screen.TwipsPerPixelY - Extender.Top
  If Not (Cur.X >= 0 And Cur.X <= UserControl.Width _
          And Cur.Y >= 0 And Cur.Y <= UserControl.Height) Then
    tmrHiLight.Enabled = False
    MouseOver = False
    UserControl.Picture = UserControl.imgLib(1).Picture
    UserControl.lblCaption.ForeColor = RGB(196, 196, 196)
    RaiseEvent MouseOut
  End If
End Sub

Private Sub UserControl_Click()
  If Me.Enabled Then RaiseEvent Click
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Me.Enabled Then
    If Not MouseOver Then HiLight
  End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  Me.Caption = PropBag.ReadProperty("Caption", "")
  Me.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

Private Sub UserControl_Resize()
  UserControl.Width = UserControl.imgLib(0).Width
  UserControl.Height = UserControl.imgLib(0).Height
End Sub

Public Property Get Caption() As String
  Caption = UserControl.lblCaption
End Property

Public Property Let Caption(ByVal vNewValue As String)
  UserControl.lblCaption = vNewValue
  PropertyChanged "Caption"
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  PropBag.WriteProperty "Caption", UserControl.lblCaption.Caption
  PropBag.WriteProperty "Enabled", UserControl.Enabled
End Sub

Private Sub HiLight()
  MouseOver = True
  UserControl.Picture = UserControl.imgLib(0)
  UserControl.lblCaption.ForeColor = RGB(255, 255, 255)
  RaiseEvent MouseIn
  If Not tmrHiLight.Enabled Then tmrHiLight.Enabled = True
End Sub

Public Property Get Enabled() As Boolean
  Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal vNewValue As Boolean)
  UserControl.Enabled = vNewValue
  If Me.Enabled Then
    UserControl.Picture = UserControl.imgLib(1)
  Else
    UserControl.Picture = UserControl.imgLib(2)
  End If
  PropertyChanged "Enabled"
End Property
