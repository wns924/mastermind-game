VERSION 5.00
Begin VB.Form frmSignIn 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  '沒有框線
   Caption         =   "Sign In"
   ClientHeight    =   2520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5670
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmSignIn.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   5670
   StartUpPosition =   1  '所屬視窗中央
   Begin VB.TextBox TextSetFocus 
      Appearance      =   0  '平面
      Height          =   270
      Left            =   240
      TabIndex        =   1
      Top             =   2640
      Width           =   1935
   End
   Begin VB.TextBox Username 
      Appearance      =   0  '平面
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '沒有框線
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   360
      TabIndex        =   3
      Text            =   "Username"
      Top             =   1080
      Width           =   3375
   End
   Begin VB.TextBox Password 
      Appearance      =   0  '平面
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '沒有框線
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   285
      IMEMode         =   3  '暫止
      Left            =   360
      TabIndex        =   5
      Text            =   "Password"
      Top             =   1680
      Width           =   3375
   End
   Begin VB.Timer tmrMouseOver 
      Interval        =   1
      Left            =   4800
      Top             =   480
   End
   Begin VB.Label lblEnd 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      BackColor       =   &H00202020&
      BackStyle       =   0  '透明
      Caption         =   "離開"
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   4950
      TabIndex        =   4
      Top             =   2160
      Width           =   390
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      Height          =   525
      Left            =   240
      Top             =   960
      Width           =   3615
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      Height          =   525
      Left            =   240
      Top             =   1560
      Width           =   3615
   End
   Begin VB.Label Sign_In 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "Sign in"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   510
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1275
   End
   Begin VB.Shape shp1 
      BackColor       =   &H8000000D&
      BorderColor     =   &H8000000D&
      BorderWidth     =   2
      FillColor       =   &H8000000D&
      Height          =   2475
      Left            =   30
      Top             =   30
      Width           =   5625
   End
   Begin VB.Image SignIn3 
      Height          =   465
      Left            =   3600
      Picture         =   "frmSignIn.frx":000C
      Top             =   3120
      Width           =   1260
   End
   Begin VB.Image SignIn2 
      Height          =   465
      Left            =   3600
      Picture         =   "frmSignIn.frx":1ED2
      Top             =   2640
      Width           =   1260
   End
   Begin VB.Image SignIn 
      Height          =   465
      Left            =   4080
      Picture         =   "frmSignIn.frx":3D98
      Top             =   1320
      Width           =   1260
   End
   Begin VB.Label lblRegister 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      BackColor       =   &H00202020&
      BackStyle       =   0  '透明
      Caption         =   "註冊帳號"
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   4020
      TabIndex        =   0
      Top             =   2160
      Width           =   780
   End
End
Attribute VB_Name = "frmSignIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  Dim fso As Object
  If Dir(App.Path & "\Data", vbDirectory) = "" Then
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CreateFolder App.Path & "\Data"
    Set fso = Nothing
  End If
  If Dir(App.Path & "\Data\Record", vbNormal) = "" Then
    Open App.Path & "\Data\Record" For Output As #1
      Dim i As Long
        For i = 1 To 10
          Print #1, "Null"
          Print #1, "00:00"
        Next i
    Close #1
  End If
End Sub

Private Sub lblEnd_Click()
  End
End Sub

Private Sub lblRegister_Click()
  Dim AC, PW As String
  AC = InputBox("請輸入帳號。")
  If LCase(AC) = "admin" Then MsgBox "不可使用！", vbCritical: Exit Sub
  If AC = "" Then MsgBox "請輸入正確文字。", vbCritical: Exit Sub
  If GetSetting(Appname:="MasterMind", Section:="Name", Key:=LCase(AC)) <> "" Then MsgBox "此帳號已被登記！", vbCritical: Exit Sub
  PW = InputBox("請輸入密碼。")
  If PW = "" Then MsgBox "請輸入正確文字。", vbCritical: Exit Sub
  MsgBox "註冊成功！", vbOKOnly
  SaveSetting "MasterMind", "Name", LCase(AC), PW
  Open App.Path & "\Data\" & LCase(AC) & ".data" For Output As #1
    Print #1, "0" & vbCrLf & "0" & vbCrLf & "00:00"
  Close #1
End Sub

Private Sub SignIn_Click()
  On Error Resume Next
  If Username.Text = "" Or Username.Text = "Username" Or Password.Text = "" Or Password.Text = "Password" Then
    MsgBox "請輸入帳號及密碼。", vbCritical
    TextSetFocus.SetFocus
    Exit Sub
  End If
  If (LCase(Username.Text) = "admin" And Password.Text = "123") Or (GetSetting(Appname:="MasterMind", Section:="Name", Key:=Username.Text) = Password.Text) Then
    If LCase(Username.Text) <> "admin" Then
      With frmSettings
        .lblSeeAns.Visible = False
        .buttonSeeAns.Visible = False
        .lblRowsNColumns.Top = .lblRowsNColumns.Top - 570
        .ComboColumns.Top = .ComboColumns.Top - 570
        .ComboRows.Top = .ComboRows.Top - 570
        .lblYes.Top = .lblYes.Top - 570
        .shp1.Height = .shp1.Height - 570
        .Height = .Height - 570
      End With
    End If
    
    Dim Win, Lose, Record As String
    Dim Load As String
    Open App.Path & "\Data\" & LCase(Username.Text) & ".data" For Input As #1
    Line Input #1, Win
    Line Input #1, Lose
    Line Input #1, Record
    Close #1
    
    If Win < 0 Or _
    Not IsNumeric(Win) Or _
    Lose < 0 Or _
    Not IsNumeric(Lose) Or _
    InStr(1, Record, ":") = 0 Or _
    Not IsNumeric(Replace(Record, ":", "")) Then
      frmMajor.lblUsername.Caption = LCase(Username.Text)
      frmMsgBox.lblInfo.Caption = "文件損失！" & vbCrLf & "將重置帳號資料。。。"
      frmMsgBox.lblYes.Left = 3360
      frmMsgBox.Show 1, Me
      frmMsgBox.lblYes.Left = 2040
      Open App.Path & "\Data\" & LCase(Username.Text) & ".data" For Output As #1
        Print #1, "0" & vbCrLf & "0" & vbCrLf & "00:00"
      Close #1
      WinTimes = 0: LoseTimes = 0: FastestRecord = "00:00"
    Else
      WinTimes = Win: LoseTimes = Lose: FastestRecord = Record
    End If
    
    Dim str, str2 As String
    Dim i As Long
    Open App.Path & "\Data\Record" For Input As #1
      For i = 1 To 10
        Line Input #1, str
        Line Input #1, str2
        RankingName(i) = str
        RankingTime(i) = str2
      Next i
    Close #1
    
    frmMajor.lblUsername.Caption = LCase(Username.Text)
    frmMsgBox.lblInfo.Caption = "登入成功！"
    frmMsgBox.lblYes.Left = 3360
    frmMsgBox.Show 1, Me
    frmMsgBox.lblYes.Left = 2040
    Unload Me
  Else
    frmMsgBox.lblInfo.Caption = "登入失敗！"
    frmMsgBox.lblYes.Left = 3360
    frmMsgBox.Show 1, Me
    frmMsgBox.lblYes.Left = 2040
  End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  TextSetFocus.SetFocus
End Sub

Private Sub tmrMouseOver_Timer()
  If MouseOver(Me, SignIn) Then
    SignIn.Picture = SignIn2.Picture
  Else
    SignIn.Picture = SignIn3.Picture
  End If
End Sub

Private Sub Username_GotFocus()
  Shape1.BorderColor = &H0
  If Username.Text = "Username" Then Username.Text = ""
  Username.ForeColor = &H0
End Sub

Private Sub Username_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then SignIn_Click
End Sub

Private Sub Password_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then SignIn_Click
End Sub

Private Sub Username_LostFocus()
  Shape1.BorderColor = &HC0C0C0
  If Username.Text = "" Or Username.Text = "Username" Then Username.Text = "Username": Username.ForeColor = &H808080
End Sub

Private Sub Password_GotFocus()
  Shape2.BorderColor = &H0
  If Password.Text = "" Or Password.Text = "Password" Then Password.Text = ""
  Password.ForeColor = &H0
  Password.PasswordChar = "*"
End Sub

Private Sub Password_LostFocus()
  Shape2.BorderColor = &HC0C0C0
  If Password.Text = "" Or Password.Text = "Password" Then Password.Text = "Password": Password.ForeColor = &H808080: Password.PasswordChar = ""
End Sub
