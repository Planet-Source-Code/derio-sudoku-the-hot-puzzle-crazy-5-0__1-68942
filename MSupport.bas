Attribute VB_Name = "MSupport"
Option Explicit

Public Type POINT
  X As Long
  Y As Long
End Type

Public Declare Sub Sleep _
  Lib "kernel32" (ByVal Milliseconds As Long)

Public Declare Function GetTempPath _
  Lib "kernel32" _
  Alias "GetTempPathA" (ByVal nBufferLength As Long, _
                        ByVal lpBuffer As String) As Long
                        
Public Declare Function GetCursorPos _
  Lib "user32" (lpPoint As POINT) As Long

Private Declare Function GetWindowLong _
  Lib "user32" _
  Alias "GetWindowLongA" (ByVal hWnd As Long, _
                          ByVal nIndex As Long) As Long
                          
Private Declare Function SetLayeredWindowAttributes _
  Lib "user32" (ByVal hWnd As Long, _
                ByVal crKey As Long, _
                ByVal bAlpha As Byte, _
                ByVal dwFlags As Long) As Long
                
Private Declare Function SetWindowLong _
  Lib "user32" _
  Alias "SetWindowLongA" (ByVal hWnd As Long, _
                          ByVal nIndex As Long, _
                          ByVal dwNewLong As Long) As Long

Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_ALPHA = &H2

Public Sub Main()
Dim FTemp As FMain

  Set FTemp = New FMain
  With FTemp
    MakeTransparent .hWnd, 0
    .Show
  End With
End Sub

Public Function IsYes(ByVal Message As String, _
                      Optional CaptionYes As String = "Yes", _
                      Optional CaptionNo As String = "No") As Boolean
'** Show message box with question Yes or No

Dim FTemp As FQuestion
Dim actWidth As Single

  Set FTemp = New FQuestion
  With FTemp
    .Caption = " " & App.Title
    .lblMessage = Message
    .imgLogo(0).Visible = True
    .imgLogo(1).Visible = False
    .imgLogo(2).Visible = False
    
    With .btnCommand(0)
      .Caption = CaptionNo
      .Top = FTemp.lblMessage.Top + FTemp.lblMessage.Height
      If .Top < FTemp.imgLogo(0).Top + FTemp.imgLogo(0).Height Then
        .Top = FTemp.imgLogo(0).Top + FTemp.imgLogo(0).Height
      End If
      
      .Top = .Top + 300
      .Visible = True
    End With
    
    With .btnCommand(1)
      .Caption = CaptionYes
      .Top = FTemp.btnCommand(0).Top
      .Visible = True
    End With
    
    actWidth = .lblMessage.Width
    If .TextWidth(Message) < .lblMessage.Width Then
      .lblMessage.Width = .TextWidth(Message) + 60
      If .lblMessage.Width < 2 * .btnCommand(0).Width + 30 Then
        .lblMessage.Width = 2 * .btnCommand(0).Width + 30
      End If
    End If
    actWidth = actWidth - .lblMessage.Width
    .Width = .Width - actWidth
    .btnCommand(0).Left = .btnCommand(0).Left - actWidth
    .btnCommand(1).Left = .btnCommand(1).Left - actWidth
    
    .Height = .Height - .ScaleHeight + .btnCommand(0).Top + .btnCommand(0).Height + 60
    
    HideForm FTemp
    
    IsYes = (.Tag = "Yes")
  End With
  Unload FTemp
  Set FTemp = Nothing
End Function

Public Sub ShowInfo(ByVal Message As String, _
                    Optional Warning As Boolean = False)
'** Show message box just for information

Dim FTemp As FQuestion
Dim actWidth As Single

  Set FTemp = New FQuestion
  With FTemp
    .Caption = " " & App.Title
    .imgLogo(0).Visible = False
    If Not Warning Then
      .imgLogo(1).Visible = True
      .imgLogo(2).Visible = False
    Else
      .imgLogo(1).Visible = False
      .imgLogo(2).Visible = True
    End If
    
    .lblMessage = Message
    
    With .btnCommand(0)
      .Caption = "OK"
      .Top = FTemp.lblMessage.Top + FTemp.lblMessage.Height
      If .Top < FTemp.imgLogo(0).Top + FTemp.imgLogo(0).Height Then
        .Top = FTemp.imgLogo(0).Top + FTemp.imgLogo(0).Height
      End If
      
      .Top = .Top + 300
      .Visible = True
    End With
    
    actWidth = .lblMessage.Width
    If .TextWidth(Message) < .lblMessage.Width Then
      .lblMessage.Width = .TextWidth(Message) + 60
      If .lblMessage.Width < .btnCommand(0).Width + 30 Then
        .lblMessage.Width = .btnCommand(0).Width + 30
      End If
    End If
    actWidth = actWidth - .lblMessage.Width
    .Width = .Width - actWidth
    .btnCommand(0).Left = .btnCommand(0).Left - actWidth
    .btnCommand(1).Left = .btnCommand(1).Left - actWidth
    .btnCommand(1).Visible = False
    .Height = .Height - .ScaleHeight + .btnCommand(0).Top + .btnCommand(0).Height + 60
    
    HideForm FTemp
  End With
  Unload FTemp
  Set FTemp = Nothing
End Sub

Public Sub HideForm(MyForm As Form, _
                    Optional Modal As Boolean = True, _
                    Optional ParentForm As Form)
'** Show form transparent

  MakeTransparent MyForm.hWnd, 0
  MyForm.tmrTransparent.Enabled = True
  If Modal Then
    MyForm.Show vbModal
  Else
    If ParentForm Is Nothing Then
      MyForm.Show
    Else
      MyForm.Show , ParentForm
    End If
  End If
End Sub

Public Sub MakeTransparent(ByVal hWnd As Long, _
                           Opacity As Integer)
'** Make form transparent base on Opacity

Dim Msg As Long
    
  Msg = GetWindowLong(hWnd, GWL_EXSTYLE)
  Msg = Msg Or WS_EX_LAYERED
  SetWindowLong hWnd, GWL_EXSTYLE, Msg
  SetLayeredWindowAttributes hWnd, 0, Opacity, LWA_ALPHA
End Sub

Public Sub FadeIn(MyForm As Form)
'** Make form transparent from 0 to 225

Dim I As Integer

  For I = 0 To 225
    MakeTransparent MyForm.hWnd, I
    DoEvents
  Next I
  MyForm.tmrTransparent.Enabled = False
End Sub

Public Sub FadeOut(MyForm As Form, Optional StartingPoint As Integer = 225)
'** Make form transparent from 225 (base on starting point) to 0

Dim I As Integer
  
  For I = StartingPoint To 0 Step -1
    MakeTransparent MyForm.hWnd, I
    DoEvents
  Next I
  MyForm.Hide
End Sub

