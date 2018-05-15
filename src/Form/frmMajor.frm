VERSION 5.00
Begin VB.Form frmMajor 
   Appearance      =   0  '����
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  '�S���ؽu
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
   StartUpPosition =   2  '�ù�����
   Begin VB.Timer tmrShowForm 
      Interval        =   1
      Left            =   4440
      Top             =   480
   End
   Begin VB.PictureBox FrmMenu_bg 
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      BorderStyle     =   0  '�S���ؽu
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
      ToolTipText     =   "�ӤH�ɮ�/�Ʀ�]"
      Top             =   1080
      Width           =   465
   End
   Begin VB.Label lblGameVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "MASTERMIND v"
      BeginProperty Font 
         Name            =   "�L�n������"
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
      Alignment       =   1  '�a�k���
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "(username)"
      BeginProperty Font 
         Name            =   "�L�n������"
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
      Alignment       =   2  '�m�����
      BackColor       =   &H00FFFFFF&
      Caption         =   "STATUS"
      BeginProperty Font 
         Name            =   "�L�n������"
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
      Alignment       =   2  '�m�����
      BackStyle       =   0  '�z��
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "�L�n������"
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
      Shape           =   3  '���
      Top             =   840
      Width           =   405
   End
   Begin VB.Shape shpBall 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   405
      Index           =   5
      Left            =   2880
      Shape           =   3  '���
      Top             =   5160
      Width           =   405
   End
   Begin VB.Shape shpBall 
      BorderWidth     =   2
      Height          =   405
      Index           =   4
      Left            =   2400
      Shape           =   3  '���
      Top             =   5160
      Width           =   405
   End
   Begin VB.Shape shpBall 
      BorderWidth     =   2
      Height          =   405
      Index           =   3
      Left            =   1920
      Shape           =   3  '���
      Top             =   5160
      Width           =   405
   End
   Begin VB.Shape shpBall 
      BorderWidth     =   2
      Height          =   420
      Index           =   2
      Left            =   1440
      Shape           =   3  '���
      Top             =   5160
      Width           =   420
   End
   Begin VB.Shape shpBall 
      BorderWidth     =   2
      Height          =   405
      Index           =   1
      Left            =   960
      Shape           =   3  '���
      Top             =   5160
      Width           =   405
   End
   Begin VB.Shape shpBall 
      BorderWidth     =   2
      Height          =   405
      Index           =   0
      Left            =   480
      Shape           =   3  '���
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
      BackStyle       =   0  '�z��
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "�L�n������"
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
      ToolTipText     =   "�C���]�w"
      Top             =   3840
      Width           =   465
   End
   Begin VB.Image imgTutorial 
      Height          =   465
      Left            =   5160
      Picture         =   "frmMajor.frx":30DC7
      Stretch         =   -1  'True
      ToolTipText     =   "�d�ݱе{"
      Top             =   3240
      Width           =   465
   End
   Begin VB.Image imgDelete 
      Height          =   465
      Left            =   5160
      Picture         =   "frmMajor.frx":33A30
      Stretch         =   -1  'True
      ToolTipText     =   "���s�}�l"
      Top             =   2640
      Width           =   465
   End
   Begin VB.Label lblTimer 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "�L�n������"
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
      ToolTipText     =   "�C���ɶ�"
      Top             =   2280
      Width           =   465
   End
   Begin VB.Image imgChronometer 
      Height          =   465
      Left            =   5160
      Picture         =   "frmMajor.frx":38D01
      Stretch         =   -1  'True
      ToolTipText     =   "�C���ɶ�"
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
      BackStyle       =   0  '�z��
      Caption         =   "�������T           ��m���~"
      BeginProperty Font 
         Name            =   "�L�n������"
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
  If App.PrevInstance = True Then End '//��{���w�Q�}�ҡA�h�����{��
  Call ���J����(Me)
  Call ��l��(Me, True)
  lblGameStatus.Top = 5160
  lblGameVersion.Caption = lblGameVersion.Caption & App.Major & "." & App.Minor & "." & App.Revision '//��ܪ���
  Columns = 4: Rows = 8
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ���ʵ��� hwnd
End Sub

Private Sub FrmMenu_bg_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  FrmMenu_bg.Picture = LoadResPicture("000000", vbResBitmap) '//�ഫ�Ϥ�
End Sub

Private Sub imgBall_Click(Index As Integer)
  If GameOver = True Then Exit Sub
  Dim CorrectTimes As Byte
  ClickTimes = ClickTimes + 1 '//���W���O�P����
  
  If ClickTimes <= Rows * Columns Then '//���q���Ҧ��զX
    tmrTime.Enabled = True '//�}�l�p��
    imgAttempt(ClickTimes).Picture = imgBall(Index).Picture '//�]�m���եO�P���Ϥ�
    imgAttempt(ClickTimes).Tag = Index '//�]�m�ĴX�ӥO�P���C��
    If NoRepetitiveColor Then imgBall(Index).Visible = False: shpBall(Index).Visible = False
    If ClickTimes Mod Columns = 0 Then '//�����@���q��
      CorrectTimes = �ˬd�q�����p
      Dim i As Byte
      For i = 0 To 5
        imgBall(i).Visible = True: shpBall(i).Visible = True
      Next i
      If CorrectTimes = Columns Then '//�q�����T
        Call �q�����T�Υ���(Me, True)
      End If
      
    End If
    
  End If
  
  If ClickTimes >= Rows * Columns And GameOver = False Then '//�κɩҦ����|
    Call �q�����T�Υ���(Me, False)
  End If
End Sub

Private Sub imgDelete_Click()
  If tmrTime.Enabled Or GameOver Then  '//�C���w�}�l�ο饪
    frmMsgBox.lblInfo.Caption = "�A�O�_�n���s�}�l�C���H"
    frmMsgBox.Show 1, Me
    If frmMsgBox.Tag = "1" Then '//�Τ��ܽT�w
      Call ��l��(Me, False)
    End If
  ElseIf lblTimer.Caption = "00:00" Then
    frmMsgBox.lblInfo.Caption = "�C�����}�l�I"
    frmMsgBox.lblYes.Left = 3360
    frmMsgBox.Show 1, Me
    frmMsgBox.lblYes.Left = 2040
  End If
End Sub

Private Sub imgIcon_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ���ʵ��� hwnd
End Sub

Private Sub imgProfile_Click()
  frmProfile.Show 1, Me
End Sub

Private Sub imgSad_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ���ʵ��� hwnd
End Sub

Private Sub imgScore_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ���ʵ��� hwnd
End Sub

Private Sub imgSettings_Click()
  frmSettings.Show 1, Me
End Sub

Private Sub imgSmile_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ���ʵ��� hwnd
End Sub

Private Sub imgSolution_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ���ʵ��� hwnd
End Sub

Private Sub imgTutorial_Click()
  Dim Status As Boolean
  Status = tmrTime.Enabled
  tmrTime.Enabled = False '//����p��
  frmMsgBox.lblInfo.Caption = "�T�w�n�d�ݱе{�ܡH"
  frmMsgBox.Show 1, Me
  If frmMsgBox.Tag = "1" Then '//�Τ��ܽT�w
    frmTutorial.Show 1, Me
  End If
  tmrTime.Enabled = Status  '//�p�G�}�l�F�C���A�h�~��p��
End Sub

Private Sub lblDescription_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ���ʵ��� hwnd
End Sub

Private Sub lblGameVersion_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ���ʵ��� hwnd
End Sub

Private Sub lblGameStatus_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ���ʵ��� hwnd
End Sub

Private Sub lblSolution_Click()
  If tmrTime.Enabled Then '//�C���w�}�l
    frmMsgBox.lblInfo.Caption = "�A�O�_�n�����[�ݵ��סH"
    frmMsgBox.Show 1, Me
    If frmMsgBox.Tag = "1" Then '//�Τ��ܽT�w
      Call �q�����T�Υ���(Me, False)
    End If
  Else
    frmMsgBox.lblInfo.Caption = "�C�����}�l�I"
    frmMsgBox.lblYes.Left = 3360
    frmMsgBox.Show 1, Me
    frmMsgBox.lblYes.Left = 2040
  End If
End Sub

Private Sub FrmMenu_03_Click()
  frmMsgBox.lblInfo.Caption = "�A�O�_�n���}�C���H"
  frmMsgBox.Show 1, Me
  If frmMsgBox.Tag = "1" Then End
End Sub

Private Sub FrmMenu_01_Click()
  FrmMenu_bg.Picture = LoadResPicture("000000", vbResBitmap) '//�ഫ�Ϥ�
  Me.WindowState = 1 '//�Y�p����
End Sub

Private Sub FrmMenu_01_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 0 Then FrmMenu_bg.Picture = LoadResPicture("010000", vbResBitmap) '//�ഫ�Ϥ�
End Sub

Private Sub FrmMenu_01_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  FrmMenu_bg.Picture = LoadResPicture("020000", vbResBitmap) '//�ഫ�Ϥ�
End Sub

Private Sub FrmMenu_03_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 0 Then FrmMenu_bg.Picture = LoadResPicture("000001", vbResBitmap) '//�ഫ�Ϥ�
End Sub

Private Sub FrmMenu_03_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  FrmMenu_bg.Picture = LoadResPicture("000002", vbResBitmap) '//�ഫ�Ϥ�
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  FrmMenu_bg.Picture = LoadResPicture("000000", vbResBitmap) '//�ഫ�Ϥ�
End Sub

Private Sub lblUsername_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ���ʵ��� hwnd
End Sub

Private Sub tmrMouseOver_Timer()
  Dim i As Integer
  For i = 0 To 5
    shpBall(i).BorderColor = IIf(MouseOver(Me, imgBall(i)), vbBlack, &H808080) '//�ˬd�ƹ��O�_�ƹL���
  Next i
End Sub

Private Sub tmrShowForm_Timer()
  frmSignIn.Show 1, Me
  tmrShowForm.Enabled = False
End Sub

Private Sub tmrTime_Timer()
  Dim Min, Sec As String
  Time = Time + 1 '//���W�ɶ�
  Min = IIf(Time / 60 < 10, "0" & Round(Time / 60), Round(Time / 60)) '//�p�G�����O���ơA�h��0�b�e
  Sec = IIf(Time Mod 60 < 10, "0" & Round(Time Mod 60), Round(Time Mod 60)) '//�p�G�����O���ơA�h��0�b�e
  lblTimer.Caption = Min & ":" & Sec '//��ܤw�����ɶ�
End Sub
