VERSION 5.00
Begin VB.UserControl OnOffButton 
   BackColor       =   &H00FFFFFF&
   BackStyle       =   0  '透明
   ClientHeight    =   585
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1080
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   ScaleHeight     =   585
   ScaleWidth      =   1080
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '平面
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   195
      TabIndex        =   0
      Top             =   75
      Width           =   220
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  '平面
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  '沒有框線
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   150
      ScaleHeight     =   315
      ScaleWidth      =   795
      TabIndex        =   1
      Top             =   150
      Width           =   800
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   375
      Left            =   120
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "OnOffButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim Switched As Boolean

Private Sub Picture1_Click()
  Picture2_Click
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Shape1.BorderColor = vbBlack
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Shape1.BorderColor = &H808080
End Sub

Private Sub Picture2_Click()
  If Picture2.BackColor = &HC0C0C0 Then
    Picture2.BackColor = &HCC7A00
    Picture1.Left = 745
    Switched = True
  Else
    Picture2.BackColor = &HC0C0C0
    Picture1.Left = 120
    Switched = False
  End If
End Sub

Public Property Get Switch() As Boolean
  Switch = Switched
End Property

Public Property Let Switch(ByVal vNewValue As Boolean)
  Switched = vNewValue
  If Switched Then
    Picture2.BackColor = &HCC7A00
    Picture1.Left = 745
    Switched = True
  Else
    Picture2.BackColor = &HC0C0C0
    Picture1.Left = 120
    Switched = False
  End If
End Property

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Shape1.BorderColor = vbBlack
End Sub

Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Shape1.BorderColor = &H808080
End Sub
