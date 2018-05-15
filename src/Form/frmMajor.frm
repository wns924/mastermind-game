VERSION 5.00
Begin VB.Form frmMajor 
   Appearance      =   0  '平面
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  '沒有框線
   Caption         =   "MasterMind"
   ClientHeight    =   5730
   ClientLeft      =   18150
   ClientTop       =   2805
   ClientWidth     =   6255
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmMajor.frx":0000
   LinkTopic       =   "frmMajor"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   6255
   StartUpPosition =   2  '螢幕中央
   Begin VB.Timer tmrShowForm 
      Interval        =   1
      Left            =   4440
      Top             =   480
   End
   Begin VB.PictureBox FrmMenu_bg 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   4560
      Picture         =   "frmMajor.frx":7F6A
      ScaleHeight     =   300
      ScaleWidth      =   1470
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   30
      Width           =   1470
      Begin VB.Image FrmMenu_03 
         Height          =   315
         Left            =   795
         Top             =   0
         Width           =   675
      End
      Begin VB.Image FrmMenu_01 
         Height          =   315
         Left            =   0
         Top             =   0
         Width           =   400
      End
   End
   Begin VB.Timer tmrMouseOver 
      Interval        =   1
      Left            =   3600
      Top             =   5160
   End
   Begin VB.Timer tmrTime 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5640
      Top             =   1200
   End
   Begin VB.Image imgProfile 
      Height          =   465
      Left            =   5160
      Picture         =   "frmMajor.frx":96CC
      Stretch         =   -1  'True
      ToolTipText     =   "個人檔案/排行榜"
      Top             =   1080
      Width           =   465
   End
   Begin VB.Label lblGameVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "MASTERMIND v"
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   720
      TabIndex        =   0
      Top             =   180
      Width           =   2370
   End
   Begin VB.Label lblUsername 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "(username)"
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3330
      TabIndex        =   7
      Top             =   420
      Width           =   2730
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgSolution 
      Height          =   405
      Index           =   1
      Left            =   4080
      Stretch         =   -1  'True
      Top             =   5160
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Label lblGameStatus 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00FFFFFF&
      Caption         =   "STATUS"
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   450
      TabIndex        =   6
      Top             =   4080
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Image imgSad 
      Height          =   255
      Left            =   4080
      Picture         =   "frmMajor.frx":F372
      Stretch         =   -1  'True
      Top             =   4755
      Width           =   255
   End
   Begin VB.Image imgSmile 
      Height          =   255
      Left            =   2880
      Picture         =   "frmMajor.frx":1DD02
      Stretch         =   -1  'True
      Top             =   4755
      Width           =   255
   End
   Begin VB.Label lblSolution 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "答案"
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3960
      TabIndex        =   3
      Top             =   5160
      Width           =   2160
   End
   Begin VB.Shape shp6 
      BackColor       =   &H8000000D&
      BorderColor     =   &H8000000D&
      FillColor       =   &H8000000D&
      Height          =   690
      Left            =   3960
      Top             =   5040
      Width           =   2175
   End
   Begin VB.Shape shpAttempt 
      Height          =   405
      Index           =   1
      Left            =   480
      Shape           =   3  '圓形
      Top             =   840
      Width           =   405
   End
   Begin VB.Shape shpBall 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   405
      Index           =   5
      Left            =   2880
      Shape           =   3  '圓形
      Top             =   5160
      Width           =   405
   End
   Begin VB.Shape shpBall 
      BorderWidth     =   2
      Height          =   405
      Index           =   4
      Left            =   2400
      Shape           =   3  '圓形
      Top             =   5160
      Width           =   405
   End
   Begin VB.Shape shpBall 
      BorderWidth     =   2
      Height          =   405
      Index           =   3
      Left            =   1920
      Shape           =   3  '圓形
      Top             =   5160
      Width           =   405
   End
   Begin VB.Shape shpBall 
      BorderWidth     =   2
      Height          =   420
      Index           =   2
      Left            =   1440
      Shape           =   3  '圓形
      Top             =   5160
      Width           =   420
   End
   Begin VB.Shape shpBall 
      BorderWidth     =   2
      Height          =   405
      Index           =   1
      Left            =   960
      Shape           =   3  '圓形
      Top             =   5160
      Width           =   405
   End
   Begin VB.Shape shpBall 
      BorderWidth     =   2
      Height          =   405
      Index           =   0
      Left            =   480
      Shape           =   3  '圓形
      Top             =   5160
      Width           =   405
   End
   Begin VB.Image imgBall 
      Height          =   405
      Index           =   5
      Left            =   2880
      Picture         =   "frmMajor.frx":254B2
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   405
   End
   Begin VB.Image imgBall 
      Height          =   405
      Index           =   4
      Left            =   2400
      Picture         =   "frmMajor.frx":25FBC
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   405
   End
   Begin VB.Image imgBall 
      Height          =   405
      Index           =   3
      Left            =   1920
      Picture         =   "frmMajor.frx":26AC6
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   405
   End
   Begin VB.Image imgBall 
      Height          =   405
      Index           =   2
      Left            =   1440
      Picture         =   "frmMajor.frx":275D0
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   405
   End
   Begin VB.Image imgBall 
      Height          =   405
      Index           =   1
      Left            =   960
      Picture         =   "frmMajor.frx":280DA
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   405
   End
   Begin VB.Image imgBall 
      Height          =   405
      Index           =   0
      Left            =   480
      Picture         =   "frmMajor.frx":28BE4
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   405
   End
   Begin VB.Shape shp5 
      BackColor       =   &H8000000D&
      BorderColor     =   &H8000000D&
      FillColor       =   &H8000000D&
      Height          =   690
      Left            =   360
      Top             =   5040
      Width           =   3135
   End
   Begin VB.Image imgScore 
      Height          =   255
      Index           =   1
      Left            =   3000
      Stretch         =   -1  'True
      Top             =   900
      Width           =   255
   End
   Begin VB.Shape shp3 
      BackColor       =   &H8000000D&
      BorderColor     =   &H8000000D&
      FillColor       =   &H8000000D&
      Height          =   3975
      Left            =   2880
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label lblLine 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   240
      Index           =   0
      Left            =   2520
      TabIndex        =   2
      Top             =   960
      Width           =   105
   End
   Begin VB.Shape shp2 
      BackColor       =   &H8000000D&
      BorderColor     =   &H8000000D&
      FillColor       =   &H8000000D&
      Height          =   3975
      Left            =   360
      Top             =   720
      Width           =   2085
   End
   Begin VB.Shape shp1 
      BackColor       =   &H8000000D&
      BorderColor     =   &H8000000D&
      BorderWidth     =   2
      FillColor       =   &H8000000D&
      Height          =   5715
      Left            =   15
      Top             =   15
      Width           =   6240
   End
   Begin VB.Shape shp4 
      BackColor       =   &H8000000D&
      BorderColor     =   &H8000000D&
      FillColor       =   &H8000000D&
      Height          =   3615
      Left            =   5040
      Top             =   840
      Width           =   735
   End
   Begin VB.Image imgSettings 
      Height          =   465
      Left            =   5160
      Picture         =   "frmMajor.frx":296EE
      Stretch         =   -1  'True
      ToolTipText     =   "遊玩設定"
      Top             =   3840
      Width           =   465
   End
   Begin VB.Image imgTutorial 
      Height          =   465
      Left            =   5160
      Picture         =   "frmMajor.frx":30DC7
      Stretch         =   -1  'True
      ToolTipText     =   "查看教程"
      Top             =   3240
      Width           =   465
   End
   Begin VB.Image imgDelete 
      Height          =   465
      Left            =   5160
      Picture         =   "frmMajor.frx":33A30
      Stretch         =   -1  'True
      ToolTipText     =   "重新開始"
      Top             =   2640
      Width           =   465
   End
   Begin VB.Label lblTimer 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   5160
      TabIndex        =   1
      ToolTipText     =   "遊玩時間"
      Top             =   2280
      Width           =   465
   End
   Begin VB.Image imgChronometer 
      Height          =   465
      Left            =   5160
      Picture         =   "frmMajor.frx":38D01
      Stretch         =   -1  'True
      ToolTipText     =   "遊玩時間"
      Top             =   1680
      Width           =   465
   End
   Begin VB.Image imgIcon 
      Height          =   465
      Left            =   180
      Picture         =   "frmMajor.frx":3F40E
      Stretch         =   -1  'True
      Top             =   120
      Width           =   465
   End
   Begin VB.Image imgAttempt 
      Height          =   405
      Index           =   1
      Left            =   480
      Stretch         =   -1  'True
      Top             =   840
      Width           =   405
   End
   Begin VB.Label lblDescription 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "完全正確           位置錯誤"
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   3240
      TabIndex        =   5
      Top             =   4755
      Width           =   1935
   End
End
Attribute VB_Name = "frmMajor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Initialize()
  If App.PrevInstance = True Then End '//當程式已被開啟，則結束程式
  Call 載入物件(Me)
  Call 初始化(Me, True)
  lblGameStatus.Top = 5160
  lblGameVersion.Caption = lblGameVersion.Caption & App.Major & "." & App.Minor & "." & App.Revision '//顯示版本
  Columns = 4: Rows = 8
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  移動視窗 hwnd
End Sub

Private Sub FrmMenu_bg_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  FrmMenu_bg.Picture = LoadResPicture("000000", vbResBitmap) '//轉換圖片
End Sub

Private Sub imgBall_Click(Index As Integer)
  If GameOver = True Then Exit Sub
  Dim CorrectTimes As Byte
  ClickTimes = ClickTimes + 1 '//遞增按令牌次數
  
  If ClickTimes <= Rows * Columns Then '//未猜完所有組合
    tmrTime.Enabled = True '//開始計時
    imgAttempt(ClickTimes).Picture = imgBall(Index).Picture '//設置嘗試令牌的圖片
    imgAttempt(ClickTimes).Tag = Index '//設置第幾個令牌的顏色
    If NoRepetitiveColor Then imgBall(Index).Visible = False: shpBall(Index).Visible = False
    If ClickTimes Mod Columns = 0 Then '//當完成一次猜測
      CorrectTimes = 檢查猜測狀況
      Dim i As Byte
      For i = 0 To 5
        imgBall(i).Visible = True: shpBall(i).Visible = True
      Next i
      If CorrectTimes = Columns Then '//猜測正確
        Call 猜測正確或失敗(Me, True)
      End If
      
    End If
    
  End If
  
  If ClickTimes >= Rows * Columns And GameOver = False Then '//用盡所有機會
    Call 猜測正確或失敗(Me, False)
  End If
End Sub

Private Sub imgDelete_Click()
  If tmrTime.Enabled Or GameOver Then  '//遊戲已開始或輸左
    frmMsgBox.lblInfo.Caption = "你是否要重新開始遊戲？"
    frmMsgBox.Show 1, Me
    If frmMsgBox.Tag = "1" Then '//用戶選擇確定
      Call 初始化(Me, False)
    End If
  ElseIf lblTimer.Caption = "00:00" Then
    frmMsgBox.lblInfo.Caption = "遊戲未開始！"
    frmMsgBox.lblYes.Left = 3360
    frmMsgBox.Show 1, Me
    frmMsgBox.lblYes.Left = 2040
  End If
End Sub

Private Sub imgIcon_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  移動視窗 hwnd
End Sub

Private Sub imgProfile_Click()
  frmProfile.Show 1, Me
End Sub

Private Sub imgSad_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  移動視窗 hwnd
End Sub

Private Sub imgScore_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  移動視窗 hwnd
End Sub

Private Sub imgSettings_Click()
  frmSettings.Show 1, Me
End Sub

Private Sub imgSmile_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  移動視窗 hwnd
End Sub

Private Sub imgSolution_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  移動視窗 hwnd
End Sub

Private Sub imgTutorial_Click()
  Dim Status As Boolean
  Status = tmrTime.Enabled
  tmrTime.Enabled = False '//停止計時
  frmMsgBox.lblInfo.Caption = "確定要查看教程嗎？"
  frmMsgBox.Show 1, Me
  If frmMsgBox.Tag = "1" Then '//用戶選擇確定
    frmTutorial.Show 1, Me
  End If
  tmrTime.Enabled = Status  '//如果開始了遊戲，則繼續計時
End Sub

Private Sub lblDescription_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  移動視窗 hwnd
End Sub

Private Sub lblGameVersion_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  移動視窗 hwnd
End Sub

Private Sub lblGameStatus_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  移動視窗 hwnd
End Sub

Private Sub lblSolution_Click()
  If tmrTime.Enabled Then '//遊戲已開始
    frmMsgBox.lblInfo.Caption = "你是否要放棄並觀看答案？"
    frmMsgBox.Show 1, Me
    If frmMsgBox.Tag = "1" Then '//用戶選擇確定
      Call 猜測正確或失敗(Me, False)
    End If
  Else
    frmMsgBox.lblInfo.Caption = "遊戲未開始！"
    frmMsgBox.lblYes.Left = 3360
    frmMsgBox.Show 1, Me
    frmMsgBox.lblYes.Left = 2040
  End If
End Sub

Private Sub FrmMenu_03_Click()
  frmMsgBox.lblInfo.Caption = "你是否要離開遊戲？"
  frmMsgBox.Show 1, Me
  If frmMsgBox.Tag = "1" Then End
End Sub

Private Sub FrmMenu_01_Click()
  FrmMenu_bg.Picture = LoadResPicture("000000", vbResBitmap) '//轉換圖片
  Me.WindowState = 1 '//縮小視窗
End Sub

Private Sub FrmMenu_01_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 0 Then FrmMenu_bg.Picture = LoadResPicture("010000", vbResBitmap) '//轉換圖片
End Sub

Private Sub FrmMenu_01_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  FrmMenu_bg.Picture = LoadResPicture("020000", vbResBitmap) '//轉換圖片
End Sub

Private Sub FrmMenu_03_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 0 Then FrmMenu_bg.Picture = LoadResPicture("000001", vbResBitmap) '//轉換圖片
End Sub

Private Sub FrmMenu_03_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  FrmMenu_bg.Picture = LoadResPicture("000002", vbResBitmap) '//轉換圖片
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  FrmMenu_bg.Picture = LoadResPicture("000000", vbResBitmap) '//轉換圖片
End Sub

Private Sub lblUsername_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  移動視窗 hwnd
End Sub

Private Sub tmrMouseOver_Timer()
  Dim i As Integer
  For i = 0 To 5
    shpBall(i).BorderColor = IIf(MouseOver(Me, imgBall(i)), vbBlack, &H808080) '//檢查滑鼠是否滑過控制項
  Next i
End Sub

Private Sub tmrShowForm_Timer()
  frmSignIn.Show 1, Me
  tmrShowForm.Enabled = False
End Sub

Private Sub tmrTime_Timer()
  Dim Min, Sec As String
  Time = Time + 1 '//遞增時間
  Min = IIf(Time / 60 < 10, "0" & Round(Time / 60), Round(Time / 60)) '//如果分鐘是單位數，則補0在前
  Sec = IIf(Time Mod 60 < 10, "0" & Round(Time Mod 60), Round(Time Mod 60)) '//如果秒鐘是單位數，則補0在前
  lblTimer.Caption = Min & ":" & Sec '//顯示已玩的時間
End Sub
