VERSION 5.00
Begin VB.Form frmSettings 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  '�S���ؽu
   ClientHeight    =   3735
   ClientLeft      =   7935
   ClientTop       =   4800
   ClientWidth     =   3735
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmSettings.frx":0000
   ScaleHeight     =   3735
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '���ݵ�������
   Begin VB.ComboBox ComboRows 
      BeginProperty Font 
         Name            =   "�L�n������"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmSettings.frx":7F6A
      Left            =   2160
      List            =   "frmSettings.frx":7F7D
      Style           =   2  '��¤U�Ԧ�
      TabIndex        =   6
      Top             =   1890
      Width           =   1215
   End
   Begin VB.ComboBox ComboColumns 
      BeginProperty Font 
         Name            =   "�L�n������"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmSettings.frx":7F93
      Left            =   2160
      List            =   "frmSettings.frx":7FA0
      Style           =   2  '��¤U�Ԧ�
      TabIndex        =   5
      Top             =   2370
      Width           =   1215
   End
   Begin MasterMind.OnOffButton buttonColor 
      Height          =   615
      Left            =   2400
      TabIndex        =   3
      Top             =   600
      Width           =   1095
      _ExtentX        =   4048
      _ExtentY        =   1508
   End
   Begin MasterMind.OnOffButton buttonSeeAns 
      Height          =   615
      Left            =   2400
      TabIndex        =   7
      Top             =   1200
      Width           =   1095
      _ExtentX        =   4048
      _ExtentY        =   1508
   End
   Begin VB.Label lblSeeAns 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "��ܵ���"
      BeginProperty Font 
         Name            =   "�L�n������"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   8
      Top             =   1320
      Width           =   960
   End
   Begin VB.Label lblRowsNColumns 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "�C����ƩM�C��"
      BeginProperty Font 
         Name            =   "�L�n������"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   4
      Top             =   1890
      Width           =   1680
   End
   Begin VB.Label lblColor 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "�C�⤣�i�ۦP"
      BeginProperty Font 
         Name            =   "�L�n������"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   1440
   End
   Begin VB.Label lblYes 
      Alignment       =   2  '�m�����
      BackColor       =   &H00CC7A00&
      Caption         =   "�T�w"
      BeginProperty Font 
         Name            =   "�L�n������"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label lblTopic 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "�]�w"
      BeginProperty Font 
         Name            =   "�L�n������"
         Size            =   18
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009A572B&
      Height          =   465
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   720
   End
   Begin VB.Shape shp1 
      BackColor       =   &H8000000D&
      BorderColor     =   &H8000000D&
      BorderWidth     =   2
      FillColor       =   &H8000000D&
      Height          =   3720
      Left            =   15
      Top             =   15
      Width           =   3720
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  buttonColor.Switch = NoRepetitiveColor
  buttonSeeAns.Switch = SeeAns
  ComboRows.ListIndex = (Rows - 8)
  ComboColumns.ListIndex = (Columns - 4)
End Sub

Private Sub lblYes_Click()
  If (buttonColor.Switch And NoRepetitiveColor = False) Or (NoRepetitiveColor And buttonColor.Switch = False) Then '//�ﶵ���P
    NoRepetitiveColor = buttonColor.Switch
    Call ��l��(frmMajor, False)
  End If
  If (Rows <> ComboRows.Text) Or (Columns <> ComboColumns.Text) Then
    Rows = ComboRows.Text
    Columns = ComboColumns.Text
    Call ��l��(frmMajor, False)
    Call ���m�ɭ�(Rows, Columns)
  End If
  SeeAns = buttonSeeAns.Switch
  For i = 1 To Columns
    If SeeAns Then
      frmMajor.lblSolution.Visible = False
      frmMajor.imgSolution(i).Visible = True
    ElseIf GameOver = False Then
      frmMajor.lblSolution.Visible = True
      frmMajor.imgSolution(i).Visible = False
    End If
  Next i
  Unload Me
End Sub
