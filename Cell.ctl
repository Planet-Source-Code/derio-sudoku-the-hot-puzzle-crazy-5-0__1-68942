VERSION 5.00
Begin VB.UserControl Cell 
   CanGetFocus     =   0   'False
   ClientHeight    =   735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   720
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   HasDC           =   0   'False
   Picture         =   "Cell.ctx":0000
   ScaleHeight     =   735
   ScaleWidth      =   720
   ToolboxBitmap   =   "Cell.ctx":2376
   Begin VB.Timer tmrHiLight 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   900
      Top             =   120
   End
   Begin VB.Image imgLib 
      Height          =   720
      Index           =   2
      Left            =   1650
      Picture         =   "Cell.ctx":2688
      Top             =   1155
      Width           =   720
   End
   Begin VB.Image imgLib 
      Height          =   720
      Index           =   1
      Left            =   855
      Picture         =   "Cell.ctx":49B4
      Top             =   1140
      Width           =   720
   End
   Begin VB.Image imgLib 
      Height          =   720
      Index           =   0
      Left            =   60
      Picture         =   "Cell.ctx":6CA8
      Top             =   1140
      Width           =   720
   End
   Begin VB.Label lblNote 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   260
      Index           =   4
      Left            =   570
      TabIndex        =   4
      Top             =   30
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label lblNote 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   260
      Index           =   3
      Left            =   435
      TabIndex        =   3
      Top             =   30
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label lblNote 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   260
      Index           =   2
      Left            =   300
      TabIndex        =   2
      Top             =   30
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label lblNote 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   260
      Index           =   1
      Left            =   165
      TabIndex        =   1
      Top             =   30
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label lblNote 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   260
      Index           =   0
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label lblNumber 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Harrington"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   685
      Index           =   1
      Left            =   30
      TabIndex        =   6
      Top             =   45
      Width           =   615
   End
   Begin VB.Label lblNumber 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Harrington"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   685
      Index           =   0
      Left            =   45
      TabIndex        =   5
      Top             =   60
      Width           =   615
   End
End
Attribute VB_Name = "Cell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'*******************************
'* Title  : Cell               *
'* Type   : ActiveX OCX        *
'* Author : Derio              *
'* Stamp  : 31 Dec 2006        *
'* Desc   : UI for Sudoku Cell *
'*******************************

Private vCaption As String
Private vNoteIndex As Integer
Private vCurrentNoteIndex As Integer
Private vForeColor As OLE_COLOR
Private vProtectedColor As OLE_COLOR

Public AddNoteSuccess As Boolean

Public Enum SUDOKU_MODE
  Protected = 0
  LightButton = 1
  DarkButton = 2
End Enum
Private vMode As SUDOKU_MODE

Public Event LeftClick()
Public Event RightClick()
Public Event NoteClick(ByVal LastCaption As String, ByVal NoteList As String)
Public Event NoteRemove(ByVal Note As String)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)


Private Sub lblNote_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim NoteList As String
Dim LastCaption As String

  'protected cell raises nothing
  If vMode = Protected Then Exit Sub
  
  If Button = vbLeftButton Then
    If lblNote(Index).Caption <> "" Then
      
      NoteList = GetNoteList()
      LastCaption = Me.Caption
      Me.Caption = lblNote(Index).Caption
      Me.ClearNote
      RaiseEvent NoteClick(LastCaption, NoteList)
      
    Else
      RaiseEvent LeftClick
    End If
    
  Else
    If lblNote(Index).Caption <> "" Then
      LastCaption = lblNote(Index).Caption
      Me.RemoveNote lblNote(Index).Caption
      RaiseEvent NoteRemove(LastCaption)
    Else
      RaiseEvent RightClick
    End If
  End If
End Sub

Private Sub lblNote_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If vNoteIndex >= 0 Then
    If vCurrentNoteIndex <> Index Then
      If vCurrentNoteIndex <> -1 Then
        With lblNote(vCurrentNoteIndex)
          .ForeColor = RGB(64, 64, 64)
          .BackStyle = 0 'opaque
        End With
      End If
      
      vCurrentNoteIndex = Index
      With lblNote(vCurrentNoteIndex)
        .ForeColor = vbYellow
        .BackStyle = 1 'Opaque
      End With
      If Not tmrHiLight.Enabled Then tmrHiLight.Enabled = True
    End If
  End If
End Sub

Private Sub lblNumber_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  'protected cell raises nothing
  If vMode = Protected Then Exit Sub
  
  If Button = vbLeftButton Then
    RaiseEvent LeftClick
  Else
    RaiseEvent RightClick
  End If
End Sub

Private Sub lblNumber_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  'protected cell raises nothing
  If vMode = Protected Then Exit Sub
  
  RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub tmrHiLight_Timer()
Dim Cur As POINT

  GetCursorPos Cur
  Cur.X = Cur.X - (Extender.Parent.Left + (Extender.Parent.Width - Extender.Parent.ScaleWidth) / 2) \ Screen.TwipsPerPixelX
  Cur.Y = Cur.Y - (Extender.Parent.Top + Extender.Parent.Height - Extender.Parent.ScaleHeight - 30) \ Screen.TwipsPerPixelY
  
  Cur.X = Cur.X * Screen.TwipsPerPixelX - Extender.Left
  Cur.Y = Cur.Y * Screen.TwipsPerPixelY - Extender.Top
  If Not (Cur.X >= lblNote(vCurrentNoteIndex).Left _
          And Cur.X <= lblNote(vCurrentNoteIndex).Left + lblNote(vCurrentNoteIndex).Width _
          And Cur.Y >= lblNote(vCurrentNoteIndex).Top _
          And Cur.Y <= lblNote(vCurrentNoteIndex).Top + lblNote(vCurrentNoteIndex).Height) Then
    tmrHiLight.Enabled = False
    With lblNote(vCurrentNoteIndex)
      .ForeColor = RGB(64, 64, 64)
      .BackStyle = 0
    End With
    
    vCurrentNoteIndex = -1
  End If
End Sub

Private Sub UserControl_Initialize()
  vNoteIndex = -1 'no note
  vCurrentNoteIndex = -1 'no selected note
  vProtectedColor = RGB(8 * 16, 4 * 16, 0)
  vForeColor = RGB(255, 255, 0)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  'protected cell raises nothing
  If vMode = Protected Then Exit Sub
  
  RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  Caption = PropBag.ReadProperty("Caption", "")
  Mode = PropBag.ReadProperty("Mode", SUDOKU_MODE.Protected)
  ForeColor = PropBag.ReadProperty("ForeColor", vbBlack)
  ProtectedColor = PropBag.ReadProperty("ProtectedColor", vbBlack)
End Sub

Private Sub UserControl_Resize()
  UserControl.Width = Screen.TwipsPerPixelX * 48
  UserControl.Height = Screen.TwipsPerPixelY * 48
End Sub

Public Property Get Caption() As String
Attribute Caption.VB_UserMemId = 0
  Caption = vCaption
End Property

Public Property Let Caption(ByVal vNewValue As String)
  vCaption = vNewValue
  UserControl.lblNumber(0).Caption = vCaption
  UserControl.lblNumber(1).Caption = vCaption
  Me.ClearNote
  PropertyChanged "Caption"
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  PropBag.WriteProperty "Caption", vCaption
  PropBag.WriteProperty "Mode", vMode
  PropBag.WriteProperty "ForeColor", vForeColor
  PropBag.WriteProperty "ProtectedColor", vProtectedColor
End Sub

Public Sub AddNote(MyNote As String)
'** Add note to Sudoku Cell

Dim I As Integer
Dim J As Integer
Dim strTemp As String

  AddNoteSuccess = False
  
  'check note capacity
  If vNoteIndex = UserControl.lblNote.Count - 1 Then Exit Sub
  
  'check validitas
  If InStr("123456789", MyNote) = 0 Then Exit Sub
  
  'check if number exist
  If Me.Caption <> "" Then
    strTemp = Me.Caption
    Me.Caption = ""
    Me.AddNote strTemp
  End If
  
  'check if the same note exist
  For I = 0 To vNoteIndex
    If UserControl.lblNote(I).Caption = MyNote Then
      Exit Sub
    End If
  Next I
  
  'check note possition
  For I = 0 To vNoteIndex
    If UserControl.lblNote(I).Caption > MyNote Then
    
      'insert new note if the number smaller than other
      For J = vNoteIndex To I Step -1
        UserControl.lblNote(J + 1).Caption = UserControl.lblNote(J).Caption
        UserControl.lblNote(J + 1).Visible = True
      Next J
      
      Exit For
    End If
  Next I
  
  'insert note
  UserControl.lblNote(I).Caption = MyNote
  UserControl.lblNote(I).Visible = True
  vNoteIndex = vNoteIndex + 1
  AddNoteSuccess = True
End Sub

Public Sub RemoveNote(ByVal Note As String)
'** Remove note form Sudoku Cell

Dim I As Integer
Dim J As Integer

  For I = 0 To vNoteIndex
    If UserControl.lblNote(I).Caption = Note Then
      For J = I To vNoteIndex - 1
        UserControl.lblNote(J).Caption = UserControl.lblNote(J + 1).Caption
      Next J
      UserControl.lblNote(vNoteIndex).Caption = ""
      UserControl.lblNote(vNoteIndex).Visible = False
      vNoteIndex = vNoteIndex - 1
      Exit For
    End If
  Next I
End Sub

Public Sub ClearNote()
'** Clear entire note from Sudoku Cell

Dim I As Integer

  For I = 0 To vNoteIndex
    UserControl.lblNote(I).Caption = ""
    UserControl.lblNote(I).Visible = False
  Next I
  vNoteIndex = -1
End Sub

Public Property Get Mode() As SUDOKU_MODE
  Mode = vMode
End Property

Public Property Let Mode(ByVal vNewValue As SUDOKU_MODE)
  vMode = vNewValue
  Select Case vMode
  Case SUDOKU_MODE.Protected
    UserControl.Picture = UserControl.imgLib(0).Picture
    UserControl.lblNumber(1).ForeColor = vProtectedColor
    If UserControl.lblNumber(0).Visible Then UserControl.lblNumber(0).Visible = False
  
  Case SUDOKU_MODE.LightButton
    UserControl.Picture = UserControl.imgLib(1).Picture
    UserControl.lblNumber(1).ForeColor = vForeColor
    If Not UserControl.lblNumber(0).Visible Then UserControl.lblNumber(0).Visible = True
  
  Case SUDOKU_MODE.DarkButton
    UserControl.Picture = UserControl.imgLib(2).Picture
    UserControl.lblNumber(1).ForeColor = vForeColor
    If Not UserControl.lblNumber(0).Visible Then UserControl.lblNumber(0).Visible = True
  End Select
  
  PropertyChanged "Mode"
End Property

Public Property Get ForeColor() As OLE_COLOR
  ForeColor = vForeColor
End Property

Public Property Let ForeColor(ByVal vNewValue As OLE_COLOR)
  vForeColor = vNewValue
  
  If Me.Mode <> Protected Then
    UserControl.lblNumber(1).ForeColor = vNewValue
  End If
  
  PropertyChanged "ForeColor"
End Property

Public Property Get ProtectedColor() As OLE_COLOR
  ProtectedColor = vProtectedColor
End Property

Public Property Let ProtectedColor(ByVal vNewValue As OLE_COLOR)
  vProtectedColor = vNewValue
  If Me.Mode = Protected Then
    UserControl.lblNumber(1).ForeColor = vProtectedColor
  End If
  
  PropertyChanged "ProtectedColor"
End Property

Public Function GetNoteList() As String
Dim I As Integer
Dim NoteList As String

  For I = 0 To vNoteIndex
    NoteList = NoteList & UserControl.lblNote(I)
  Next I
  GetNoteList = NoteList
End Function
