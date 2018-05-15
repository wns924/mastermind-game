Attribute VB_Name = "modFunction"
Option Explicit
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Type POINTAPI
  X As Long
  Y As Long
End Type
Private Const HTCAPTION           As Long = 2
Private Const WM_NCLBUTTONDOWN    As Long = &HA1
Public Time As Long '//儲存遊玩時間
Public ClickTimes As Integer '//儲存按第幾次令牌(Token)
Public GameOver As Boolean '//儲存是否玩完
Public BallNumbers(0 To 5) As Integer '//儲存答案中不同顏色的令牌的數量
Public NoRepetitiveColor, SeeAns As Boolean '//沒有重複顏色 最上顯示 睇答案
Public Rows, Columns As Byte '//自訂行數和列數
Public WinTimes, LoseTimes As Long, FastestRecord As String
Public RankingName(1 To 10) As String, RankingTime(1 To 10) As Long

Public Sub Delay(DelayTime As Single)
  Dim ST As Single
  Dim Dummy As Integer
  ST = Timer
  Do While Timer - ST < DelayTime
    Dummy = DoEvents()
    If Timer < ST Then ST = ST - 24 * 60 Or ST - 86400
  Loop
End Sub

Public Sub 移動視窗(hwnd As Long)
  ReleaseCapture
  SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Public Function MouseOver(Form As Object, 控制項名稱 As Object) As Boolean
  Dim MousePoint As POINTAPI
  GetCursorPos MousePoint
  ScreenToClient Form.hwnd, MousePoint
  With 控制項名稱
    MouseOver = (MousePoint.X * Screen.TwipsPerPixelX < .Left Or MousePoint.X * Screen.TwipsPerPixelX > .Left + .Width Or MousePoint.Y * Screen.TwipsPerPixelY < .Top Or MousePoint.Y * Screen.TwipsPerPixelY > .Top + .Height)
  End With
End Function

Public Sub SetFrm(hwnd As Long, ByVal V1 As Byte, V2 As Byte, Optional Vstep As Integer = 1)
  Dim OldStyle As Long: OldStyle = GetWindowLong(hwnd, -20)
  SetWindowLong hwnd, -20, OldStyle Or &H80000
  Dim Value As Integer
  For Value = V1 To V2 Step Vstep
    SetLayeredWindowAttributes hwnd, 0, Value, &H2
    DoEvents
  Next
End Sub

Public Sub 載入物件(Form As Object)
  On Error Resume Next
  Dim i, j As Integer
  With Form
    For j = 0 To 11
      Load .lblLine(j) '                                       ┐
      With .lblLine(j) '                                       ｜
        .Caption = j + 1 '                                     ｜——建立行數1-12
        If j <= 7 Then .Visible = True '                       ｜
        .Top = 960 + j * 480 '                                 ｜
      End With '                                               ┘
      For i = 1 To 6 '　　　　　  　                           ┐
        Load .shpAttempt(j * 6 + i) '                          ｜
        With .shpAttempt(j * 4 + i) '                          ｜
          .Top = 840 + j * 480 '                               ｜——建立72個Shape
          .Left = i * 480 '                                    ｜
          If i <= 4 And j <= 7 Then .Visible = True '          ｜
        End With '                                             ┘
        Load .imgAttempt(j * 6 + i) '                          ┐
        With .imgAttempt(j * 4 + i) '                          ｜
          .Top = 840 + j * 480 '                               ｜
          .Left = i * 480 '                                    ｜——建立72個空白令牌
          If i <= 4 And j <= 7 Then .Visible = True '          ｜
        End With '                                             ┘
        Load .imgScore(j * 6 + i) '                            ┐
        With .imgScore(j * 4 + i) '                            ｜
          .Top = 900 + j * 480 '                               ｜——建立72個用來顯示猜測是否正確的空白圖案
          .Left = 2640 + i * 360 '                             ｜
          If i <= 4 And j <= 7 Then .Visible = True '          ｜
        End With '                                             ┘
      Next i
    Next j
  End With
End Sub

Public Sub 初始化(Form As Object, FirstStart As Boolean)
  On Error Resume Next
  Dim i, RndNum As Byte
  Randomize '//產生亂數
  With Form
    If FirstStart Then
    
    '//第一次執行
      For i = 1 To 6
        Load .imgSolution(i) '//建立答案圖案
        RndNum = Round(Rnd * 5) '//產生0-5的亂數
        If i <= 4 Then BallNumbers(RndNum) = BallNumbers(RndNum) + 1 '//增加不同顏色的令牌的數量，以供檢測
        With .imgSolution(i)
          .Top = 5160
          .Left = 3600 + i * 480
         If i <= 4 Then .Picture = Form.imgBall(RndNum).Picture '//設置答案圖片
        End With
      Next i

    Else
    
    '//不是第一次執行
      .tmrTime.Enabled = False '//停止繼續計時
      .lblTimer.Caption = "00:00" '//時間歸0
      Time = 0 '//遊玩時間歸0
      For i = 1 To 72
        .imgAttempt(i).Picture = Nothing '//初始化圖像
        .imgScore(i).Picture = Nothing '//初始化圖像
      Next i
      ClickTimes = 0 '//按第令牌次數歸0
      If SeeAns = False Then .lblSolution.Visible = True '//顯示看答案按鈕
      For i = 1 To Columns
        If .imgSolution(i).Visible And SeeAns = False Then .imgSolution(i).Visible = False '//隱藏答案
      Next i
      For i = 0 To 5
        .imgBall(i).Visible = True: .shpBall(i).Visible = True '//顯示令牌
        BallNumbers(i) = 0 '//答案中不同顏色的令牌的數量歸0
      Next i
      i = 1
      Do While i <= Columns
        RndNum = Round(Rnd * 5) '//產生0-5的亂數
        If (BallNumbers(RndNum) = 0 And NoRepetitiveColor) Or (NoRepetitiveColor = False) Then '//開啟顏色不可相同時確保沒有重複顏色或沒有開啟
          BallNumbers(RndNum) = BallNumbers(RndNum) + 1 '//增加不同顏色的令牌的數量，以供檢測
          .imgSolution(i).Picture = .imgBall(RndNum).Picture '//設置答案圖片
          i = i + 1
        End If
      Loop
      GameOver = False '回到未玩完的狀態
      .lblGameStatus.Visible = False
      
    End If
  End With
End Sub

Public Sub 猜測正確或失敗(Form As Object, Win As Boolean)
  Dim i As Byte
  With Form
    .lblSolution.Visible = False '//隱藏看答案的按鈕
    For i = 1 To Columns '
      .imgSolution(i).Visible = True '//顯示答案
    Next i
    .lblGameStatus.Visible = True
    .lblGameStatus.Caption = IIf(Win, "你贏了！", "你輸了！")
    .tmrTime.Enabled = False '//停止計時
    frmMsgBox.lblInfo.Caption = IIf(Win, "你在第" & ClickTimes / Columns & "列贏了！用時" & Replace(.lblTimer.Caption, ":", "分") & "秒！", "你輸了！")
    If Win Then
      WinTimes = WinTimes + 1
      If FastestRecord = "00:00" Or _
      Time < Val(Left(FastestRecord, 2)) * 60 + Val(Right(FastestRecord, 2)) Then '//有更快的紀錄或未贏過
        FastestRecord = Int(Time / 60) & ":" & Time - Int(Time / 60) * 60
      End If
    
      If Len(Split(FastestRecord, ":")(0)) = 1 Then FastestRecord = "0" & FastestRecord
      If Len(Split(FastestRecord, ":")(1)) = 1 Then FastestRecord = Left(FastestRecord, Len(FastestRecord) - 1) & "0" & Right(FastestRecord, 1)

    
    
      Dim j As Long
      For i = 1 To 10
        If Time <= RankingTime(i) Or (RankingName(i) = "Null" And RankingTime(i) = 0) Then
          If i < 10 Then
            For j = 10 To i + 1 Step -1
              RankingName(j) = RankingName(j - 1)
              RankingTime(j) = RankingTime(j - 1)
            Next j
          End If
          RankingName(i) = frmMajor.lblUsername.Caption
          RankingTime(i) = Time
          Exit For
        End If
      Next i
      Open App.Path & "\Data\Record" For Output As #1
        For i = 1 To 10
          Print #1, RankingName(i)
          Print #1, RankingTime(i)
        Next i
      Close #1
      
    Else
      LoseTimes = LoseTimes + 1
    End If
    Open App.Path & "\Data\" & frmMajor.lblUsername.Caption & ".data" For Output As #1
      Print #1, WinTimes & vbCrLf & LoseTimes & vbCrLf & FastestRecord
    Close #1
    frmMsgBox.lblYes.Left = 3360
    frmMsgBox.Show 1, Form
    frmMsgBox.lblYes.Left = 2040
    GameOver = True
  End With
End Sub

Public Function 檢查猜測狀況() As Byte
  Dim i, Ballnum(0 To 5), ScorePosition, CorrectTimes As Byte, CorrectToken(1 To 6) As Boolean
  For i = 0 To 5
    Ballnum(i) = BallNumbers(i) '//不直接使用BallNumbers，避免下一次錯誤，因為會修改
  Next
  With frmMajor
    ScorePosition = 1
    For i = 1 To Columns
      If .imgAttempt(ClickTimes - Columns + i).Picture = .imgSolution(i).Picture Then '//完全正確
        .imgScore(ClickTimes - Columns + ScorePosition).Picture = .imgSmile.Picture '//畫哈哈笑
        CorrectTimes = CorrectTimes + 1 '//遞增正確次數
        Ballnum(.imgAttempt(ClickTimes - Columns + i).Tag) = Ballnum(.imgAttempt(ClickTimes - Columns + i).Tag) - 1 '//減少不同顏色的令牌剩餘的數量
        CorrectToken(i) = True ''//防止被誤判為位置錯誤
        ScorePosition = ScorePosition + 1 '//遞增分數位置
      End If
    Next i
    For i = 1 To Columns
        If Ballnum(.imgAttempt(ClickTimes - Columns + i).Tag) > 0 And CorrectToken(i) = False Then '//位置錯誤且目標答案並沒有被正確猜測
          .imgScore(ClickTimes - Columns + ScorePosition).Picture = .imgSad.Picture '//畫不高興
          Ballnum(.imgAttempt(ClickTimes - Columns + i).Tag) = Ballnum(.imgAttempt(ClickTimes - Columns + i).Tag) - 1 '//減少不同顏色的令牌剩餘的數量
          ScorePosition = ScorePosition + 1   '//遞增分數位置
        End If
    Next i
  End With
  檢查猜測狀況 = CorrectTimes
End Function

Public Sub 重置界面(ByVal Rows, Columns As Byte)
  Dim j, i As Byte
  With frmMajor
    If Columns - 1 < 6 Then
      For i = Columns + 1 To 6
        .imgSolution(i).Visible = False
      Next i
    End If
    For i = 1 To 72
      If i <= 12 Then .lblLine(i - 1).Visible = (i <= Rows)
      .shpAttempt(i).Visible = (i <= Rows * Columns)
      .imgAttempt(i).Visible = (i <= Rows * Columns)
      .imgScore(i).Visible = (i <= Rows * Columns)
    Next i
    .shp2.Width = Columns * 525 - 15
    .shp2.Height = 135 + Rows * 480
    .shp3.Height = .shp2.Height
    .shp5.Top = .shp2.Top + .shp2.Height + 345
    .Height = .shp5.Top + .shp5.Height
    .shp1.Height = .Height - 15
    .shp6.Top = .Height - .shp6.Height
    .lblSolution.Top = .shp6.Top + 120
    For i = 0 To 5
      .shpBall(i).Top = .shp5.Top + 120
      .imgBall(i).Top = .shp5.Top + 120
      .imgSolution(i + 1).Top = .shp6.Top + 120
    Next i
    .lblGameStatus.Top = .shpBall(0).Top
    .lblDescription.Top = .shp6.Top - 285
    .imgSmile.Top = .lblDescription.Top
    .imgSad.Top = .lblDescription.Top
    .shp3.Left = .shp2.Left + .shp2.Width + 435
    .shp3.Width = 135 + 360 * Columns
    .shp4.Left = .shp3.Left + .shp3.Width + 585
    .imgChronometer.Left = .shp4.Left + 120
    .imgProfile.Left = .imgChronometer.Left
    .lblTimer.Left = .imgChronometer.Left
    .imgDelete.Left = .imgChronometer.Left
    .imgTutorial.Left = .imgChronometer.Left
    .imgSettings.Left = .imgChronometer.Left
    .Width = .shp4.Left + 1200
    .shp1.Width = .Width - 15
    .FrmMenu_bg.Left = .Width - 1695
    .lblUsername.Left = .Width - 195 - .lblUsername.Width
    .lblDescription.Left = .shp3.Left + 360
    .imgSmile.Left = .shp3.Left
    .imgSad.Left = .shp3.Left + 1200
    .shp6.Width = 255 + 480 * Columns
    .lblSolution.Width = .shp6.Width - 15
    For j = 0 To Rows - 1
      .lblLine(j).Left = .shp2.Left + .shp2.Width + 75
      For i = 1 To Columns
        .shpAttempt(j * Columns + i).Top = 840 + j * 480
        .shpAttempt(j * Columns + i).Left = i * 480
        
        .imgAttempt(j * Columns + i).Top = 840 + j * 480 '
        .imgAttempt(j * Columns + i).Left = i * 480

        .imgScore(j * Columns + i).Top = 900 + j * 480
        .imgScore(j * Columns + i).Left = 540 + Columns * 525 + i * 360
      Next i
    Next j
    End With
End Sub
