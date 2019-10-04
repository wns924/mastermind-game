VERSION 5.00
Begin VB.Form frmTutorial 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  '沒有框線
   ClientHeight    =   3645
   ClientLeft      =   7935
   ClientTop       =   4800
   ClientWidth     =   5325
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmTutorial.frx":0000
   ScaleHeight     =   3645
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所屬視窗中央
   Begin VB.Label lblInfo2 
      BackStyle       =   0  '透明
      Caption         =   "哈哈笑圖案是指一個令牌已經在其組合中的正確位置，而不高興圖案是指在組合中含有該令牌但位置不正確。"
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   4920
   End
   Begin VB.Label lblYes 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00CC7A00&
      Caption         =   "確定"
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label lblTopic 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "How to play MasterMind？"
      BeginProperty Font 
         Name            =   "微軟正黑體"
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
      TabIndex        =   2
      Top             =   120
      Width           =   4650
   End
   Begin VB.Image imgInfo 
      Height          =   1320
      Left            =   240
      Picture         =   "frmTutorial.frx":7F6A
      Stretch         =   -1  'True
      Top             =   600
      Width           =   1320
   End
   Begin VB.Shape shp1 
      BackColor       =   &H8000000D&
      BorderColor     =   &H8000000D&
      BorderWidth     =   2
      FillColor       =   &H8000000D&
      Height          =   3600
      Left            =   15
      Top             =   15
      Width           =   5280
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  '透明
      Caption         =   $"frmTutorial.frx":ABD3
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   1560
      TabIndex        =   0
      Top             =   720
      Width           =   3600
   End
End
Attribute VB_Name = "frmTutorial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  SetFrm Me.hwnd, 50, 200
End Sub

Private Sub lblYes_Click()
  Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then lblYes_Click
End Sub
