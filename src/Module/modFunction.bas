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
Public Time As Long '//�x�s�C���ɶ�
Public ClickTimes As Integer '//�x�s���ĴX���O�P(Token)
Public GameOver As Boolean '//�x�s�O�_����
Public BallNumbers(0 To 5) As Integer '//�x�s���פ����P�C�⪺�O�P���ƶq
Public NoRepetitiveColor, SeeAns As Boolean '//�S�������C�� �̤W��� ڻ����
Public Rows, Columns As Byte '//�ۭq��ƩM�C��
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

Public Sub ���ʵ���(hwnd As Long)
  ReleaseCapture
  SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Public Function MouseOver(Form As Object, ����W�� As Object) As Boolean
  Dim MousePoint As POINTAPI
  GetCursorPos MousePoint
  ScreenToClient Form.hwnd, MousePoint
  With ����W��
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

Public Sub ���J����(Form As Object)
  On Error Resume Next
  Dim i, j As Integer
  With Form
    For j = 0 To 11
      Load .lblLine(j) '                                       �{
      With .lblLine(j) '                                       �U
        .Caption = j + 1 '                                     �U�X�X�إߦ��1-12
        If j <= 7 Then .Visible = True '                       �U
        .Top = 960 + j * 480 '                                 �U
      End With '                                               �}
      For i = 1 To 6 '�@�@�@�@�@  �@                           �{
        Load .shpAttempt(j * 6 + i) '                          �U
        With .shpAttempt(j * 4 + i) '                          �U
          .Top = 840 + j * 480 '                               �U�X�X�إ�72��Shape
          .Left = i * 480 '                                    �U
          If i <= 4 And j <= 7 Then .Visible = True '          �U
        End With '                                             �}
        Load .imgAttempt(j * 6 + i) '                          �{
        With .imgAttempt(j * 4 + i) '                          �U
          .Top = 840 + j * 480 '                               �U
          .Left = i * 480 '                                    �U�X�X�إ�72�ӪťեO�P
          If i <= 4 And j <= 7 Then .Visible = True '          �U
        End With '                                             �}
        Load .imgScore(j * 6 + i) '                            �{
        With .imgScore(j * 4 + i) '                            �U
          .Top = 900 + j * 480 '                               �U�X�X�إ�72�ӥΨ���ܲq���O�_���T���ťչϮ�
          .Left = 2640 + i * 360 '                             �U
          If i <= 4 And j <= 7 Then .Visible = True '          �U
        End With '                                             �}
      Next i
    Next j
  End With
End Sub

Public Sub ��l��(Form As Object, FirstStart As Boolean)
  On Error Resume Next
  Dim i, RndNum As Byte
  Randomize '//���Ͷü�
  With Form
    If FirstStart Then
    
    '//�Ĥ@������
      For i = 1 To 6
        Load .imgSolution(i) '//�إߵ��׹Ϯ�
        RndNum = Round(Rnd * 5) '//����0-5���ü�
        If i <= 4 Then BallNumbers(RndNum) = BallNumbers(RndNum) + 1 '//�W�[���P�C�⪺�O�P���ƶq�A�H���˴�
        With .imgSolution(i)
          .Top = 5160
          .Left = 3600 + i * 480
         If i <= 4 Then .Picture = Form.imgBall(RndNum).Picture '//�]�m���׹Ϥ�
        End With
      Next i

    Else
    
    '//���O�Ĥ@������
      .tmrTime.Enabled = False '//�����~��p��
      .lblTimer.Caption = "00:00" '//�ɶ��k0
      Time = 0 '//�C���ɶ��k0
      For i = 1 To 72
        .imgAttempt(i).Picture = Nothing '//��l�ƹϹ�
        .imgScore(i).Picture = Nothing '//��l�ƹϹ�
      Next i
      ClickTimes = 0 '//���ĥO�P�����k0
      If SeeAns = False Then .lblSolution.Visible = True '//��ܬݵ��׫��s
      For i = 1 To Columns
        If .imgSolution(i).Visible And SeeAns = False Then .imgSolution(i).Visible = False '//���õ���
      Next i
      For i = 0 To 5
        .imgBall(i).Visible = True: .shpBall(i).Visible = True '//��ܥO�P
        BallNumbers(i) = 0 '//���פ����P�C�⪺�O�P���ƶq�k0
      Next i
      i = 1
      Do While i <= Columns
        RndNum = Round(Rnd * 5) '//����0-5���ü�
        If (BallNumbers(RndNum) = 0 And NoRepetitiveColor) Or (NoRepetitiveColor = False) Then '//�}���C�⤣�i�ۦP�ɽT�O�S�������C��ΨS���}��
          BallNumbers(RndNum) = BallNumbers(RndNum) + 1 '//�W�[���P�C�⪺�O�P���ƶq�A�H���˴�
          .imgSolution(i).Picture = .imgBall(RndNum).Picture '//�]�m���׹Ϥ�
          i = i + 1
        End If
      Loop
      GameOver = False '�^�쥼���������A
      .lblGameStatus.Visible = False
      
    End If
  End With
End Sub

Public Sub �q�����T�Υ���(Form As Object, Win As Boolean)
  Dim i As Byte
  With Form
    .lblSolution.Visible = False '//���ìݵ��ת����s
    For i = 1 To Columns '
      .imgSolution(i).Visible = True '//��ܵ���
    Next i
    .lblGameStatus.Visible = True
    .lblGameStatus.Caption = IIf(Win, "�AĹ�F�I", "�A��F�I")
    .tmrTime.Enabled = False '//����p��
    frmMsgBox.lblInfo.Caption = IIf(Win, "�A�b��" & ClickTimes / Columns & "�CĹ�F�I�ή�" & Replace(.lblTimer.Caption, ":", "��") & "��I", "�A��F�I")
    If Win Then
      WinTimes = WinTimes + 1
      If FastestRecord = "00:00" Or _
      Time < Val(Left(FastestRecord, 2)) * 60 + Val(Right(FastestRecord, 2)) Then '//����֪������Υ�Ĺ�L
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

Public Function �ˬd�q�����p() As Byte
  Dim i, Ballnum(0 To 5), ScorePosition, CorrectTimes As Byte, CorrectToken(1 To 6) As Boolean
  For i = 0 To 5
    Ballnum(i) = BallNumbers(i) '//�������ϥ�BallNumbers�A�קK�U�@�����~�A�]���|�ק�
  Next
  With frmMajor
    ScorePosition = 1
    For i = 1 To Columns
      If .imgAttempt(ClickTimes - Columns + i).Picture = .imgSolution(i).Picture Then '//�������T
        .imgScore(ClickTimes - Columns + ScorePosition).Picture = .imgSmile.Picture '//�e������
        CorrectTimes = CorrectTimes + 1 '//���W���T����
        Ballnum(.imgAttempt(ClickTimes - Columns + i).Tag) = Ballnum(.imgAttempt(ClickTimes - Columns + i).Tag) - 1 '//��֤��P�C�⪺�O�P�Ѿl���ƶq
        CorrectToken(i) = True ''//����Q�~�P����m���~
        ScorePosition = ScorePosition + 1 '//���W���Ʀ�m
      End If
    Next i
    For i = 1 To Columns
        If Ballnum(.imgAttempt(ClickTimes - Columns + i).Tag) > 0 And CorrectToken(i) = False Then '//��m���~�B�ؼе��רèS���Q���T�q��
          .imgScore(ClickTimes - Columns + ScorePosition).Picture = .imgSad.Picture '//�e������
          Ballnum(.imgAttempt(ClickTimes - Columns + i).Tag) = Ballnum(.imgAttempt(ClickTimes - Columns + i).Tag) - 1 '//��֤��P�C�⪺�O�P�Ѿl���ƶq
          ScorePosition = ScorePosition + 1   '//���W���Ʀ�m
        End If
    Next i
  End With
  �ˬd�q�����p = CorrectTimes
End Function

Public Sub ���m�ɭ�(ByVal Rows, Columns As Byte)
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
