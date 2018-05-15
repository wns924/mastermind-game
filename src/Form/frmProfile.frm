VERSION 5.00
Begin VB.Form frmProfile 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  '沒有框線
   ClientHeight    =   3645
   ClientLeft      =   7935
   ClientTop       =   4800
   ClientWidth     =   6165
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmProfile.frx":0000
   ScaleHeight     =   3645
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所屬視窗中央
   Begin VB.Frame Ranking 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '沒有框線
      Height          =   2535
      Left            =   3240
      TabIndex        =   6
      Top             =   240
      Visible         =   0   'False
      Width           =   2775
      Begin VB.ListBox ListRanking 
         Appearance      =   0  '平面
         Height          =   2010
         Left            =   120
         MultiSelect     =   1  '簡易多重選取
         TabIndex        =   8
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label lblRanking 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "排行榜"
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
         Left            =   120
         TabIndex        =   7
         Top             =   0
         Width           =   945
      End
   End
   Begin VB.Image imgChange 
      Height          =   705
      Left            =   2280
      Picture         =   "frmProfile.frx":7F6A
      Stretch         =   -1  'True
      ToolTipText     =   "排行榜"
      Top             =   2835
      Width           =   705
   End
   Begin VB.Label lblFastestRecord 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "最快紀錄："
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   360
      TabIndex        =   5
      Top             =   2280
      Width           =   1200
   End
   Begin VB.Label lblLoseTimes 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "輸的次數："
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   360
      TabIndex        =   4
      Top             =   1800
      Width           =   1200
   End
   Begin VB.Label lblWinTimes 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "贏的次數："
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   1200
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "名稱："
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   720
   End
   Begin VB.Label lblProfile 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "個人檔案"
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
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   1260
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
      Left            =   360
      TabIndex        =   0
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Shape shp1 
      BackColor       =   &H8000000D&
      BorderColor     =   &H8000000D&
      BorderWidth     =   2
      FillColor       =   &H8000000D&
      Height          =   3600
      Left            =   15
      Top             =   15
      Width           =   3120
   End
End
Attribute VB_Name = "frmProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  On Error Resume Next
  SetFrm Me.hwnd, 50, 200
  Ranking.Left = 240
  Me.Width = 3140
  Dim i As Long, j As String
    For i = 1 To 10
      j = Int(RankingTime(i) / 60) & ":" & RankingTime(i) - Int(RankingTime(i) / 60) * 60
      If Int(RankingTime(i) / 60) < 10 Then j = "0" & j
      If RankingTime(i) - Int(RankingTime(i) / 60) * 60 < 10 Then j = Replace(j, ":", ":0")
      ListRanking.AddItem i & ". " & RankingName(i) & " " & j
    Next i
  Close #1
  lblName.Caption = "名稱：" & frmMajor.lblUsername.Caption
  lblWinTimes.Caption = "贏的次數：" & WinTimes
  lblLoseTimes.Caption = "輸的次數：" & LoseTimes
  j = Int(FastestRecord / 60) & "分" & FastestRecord - Int(FastestRecord / 60) * 60 & "秒'"
  If Int(FastestRecord / 60) < 10 Then j = "0" & j
  If FastestRecord - Int(FastestRecord / 60) * 60 < 10 Then j = Replace(j, "分", "分0")
  lblFastestRecord.Caption = "最快紀錄：" & Replace(FastestRecord, ":", "分") & "秒"
End Sub

Private Sub imgChange_Click()
  Ranking.Visible = IIf(Ranking.Visible, False, True)
  imgChange.ToolTipText = IIf(Ranking.Visible, "個人檔案", "排行榜")
End Sub

Private Sub lblYes_Click()
  Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then lblYes_Click
End Sub
