VERSION 5.00
Begin VB.Form frmMsgBox 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  '沒有框線
   ClientHeight    =   2175
   ClientLeft      =   7935
   ClientTop       =   4800
   ClientWidth     =   4695
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmMsgBox.frx":0000
   ScaleHeight     =   2175
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所屬視窗中央
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
      Left            =   2040
      TabIndex        =   1
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label lblTopic 
      BackStyle       =   0  '透明
      Caption         =   "提示"
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
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblNo 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00CC7A00&
      Caption         =   "取消"
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
      Left            =   3360
      TabIndex        =   2
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Image imgInfo 
      Height          =   750
      Left            =   240
      Picture         =   "frmMsgBox.frx":7F6A
      Top             =   720
      Width           =   750
   End
   Begin VB.Shape shp1 
      BackColor       =   &H8000000D&
      BorderColor     =   &H8000000D&
      BorderWidth     =   2
      FillColor       =   &H8000000D&
      Height          =   2160
      Left            =   10
      Top             =   15
      Width           =   4680
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  '透明
      Caption         =   "提示內容"
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
      Height          =   735
      Left            =   1080
      TabIndex        =   0
      Top             =   720
      Width           =   3135
   End
End
Attribute VB_Name = "frmMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then lblYes_Click
End Sub

Private Sub Form_Load()
  SetFrm Me.hwnd, 50, 200
End Sub

Private Sub lblYes_Click()
  Unload Me
  Me.Tag = 1 '//用於被識別選項是否"確定"
End Sub

Private Sub lblno_Click()
  Unload Me
End Sub
