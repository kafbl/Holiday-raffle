VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{82351433-9094-11D1-A24B-00A0C932C7DF}#1.5#0"; "AniGIF.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "！元旦快乐！"
   ClientHeight    =   6270
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11265
   ForeColor       =   &H000000FF&
   Icon            =   "元旦抽奖.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "元旦抽奖.frx":25CA
   ScaleHeight     =   6270
   ScaleWidth      =   11265
   StartUpPosition =   3  '窗口缺省
   Begin AniGIFCtrl.AniGIF AniGIF2 
      Height          =   1455
      Left            =   8520
      TabIndex        =   7
      Top             =   2400
      Width           =   1455
      BackColor       =   12632256
      PLaying         =   -1  'True
      Transparent     =   -1  'True
      Speed           =   1
      Stretch         =   0
      AutoSize        =   0   'False
      SequenceString  =   ""
      Sequence        =   0
      HTTPProxy       =   ""
      HTTPUserName    =   ""
      HTTPPassword    =   ""
      MousePointer    =   0
      GIF             =   "元旦抽奖.frx":E9DA
      ExtendWidth     =   2566
      ExtendHeight    =   2566
      Loop            =   0
      AutoRewind      =   0   'False
      Synchronized    =   -1  'True
   End
   Begin AniGIFCtrl.AniGIF AniGIF1 
      Height          =   1575
      Left            =   600
      TabIndex        =   6
      Top             =   2400
      Width           =   1695
      BackColor       =   12632256
      PLaying         =   -1  'True
      Transparent     =   -1  'True
      Speed           =   1
      Stretch         =   0
      AutoSize        =   0   'False
      SequenceString  =   ""
      Sequence        =   0
      HTTPProxy       =   ""
      HTTPUserName    =   ""
      HTTPPassword    =   ""
      MousePointer    =   0
      GIF             =   "元旦抽奖.frx":4B2F4
      ExtendWidth     =   2990
      ExtendHeight    =   2778
      Loop            =   0
      AutoRewind      =   0   'False
      Synchronized    =   -1  'True
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   10200
      Top             =   1080
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "开始"
      BeginProperty Font 
         Name            =   "新宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1440
      MaskColor       =   &H0080FFFF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4680
      Width           =   3495
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080C0FF&
      Caption         =   "停止"
      BeginProperty Font 
         Name            =   "新宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5280
      MaskColor       =   &H0080FFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4680
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H000000C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FFFF&
      Height          =   855
      Left            =   4320
      TabIndex        =   3
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   10200
      Top             =   480
   End
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Left            =   9120
      Picture         =   "元旦抽奖.frx":87C0E
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   1
      Top             =   4320
      Width           =   1575
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      DragMode        =   1  'Automatic
      Height          =   1215
      Left            =   2760
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   5655
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   0   'False
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   9975
      _cy             =   2143
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "！！！元旦快乐！！！"
      BeginProperty Font 
         Name            =   "新宋体"
         Size            =   36
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1815
      Left            =   1440
      TabIndex        =   0
      Top             =   1320
      Width           =   8055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim s
Dim a()


Private Sub Command1_Click()

Timer2.Enabled = True



End Sub

Private Sub Command2_Click()

Timer2.Enabled = False

MsgBox "恭喜" + a(s) + "同学"




End Sub


Private Sub Form_Load()

Timer1.Enabled = True

Timer2.Enabled = False
Timer2.Interval = 50

WindowsMediaPlayer1.URL = "C:\Users\Administrator\Desktop\抽奖程序（勿删）\成品\好运来（高品）.mp3"

WindowsMediaPlayer1.Enabled = ture
Text1.FontSize = 25
Text1.FontBold = ture


Form1.Caption = "！！！元旦快乐！！！"

a = Array("董安宁", "郝泉浩", "苏纬振", "梁欣雨", "赵文龙", "牛向征", "孙一鸣", "刘尚武", "高家旺", "赵金酉", "卢玉龙", "牛金帅", "陈瀚", "姚家欣", "徐博", "于海宽", "贾姜华", "宿明丽", "郭凡", "赵迪", "贾军杰", "许世超", "张浩", "杜佳坤", "杜鹏飞", "赵一曼", "刘洁林", "王兴原", "赵珈玲", "管明明", "崔焱", "安世吉", "车轩", "赵志禹", "于鑫童", "王田锦", "耿家蕊", "郝塞娅", "赵建鑫", "戴文杰", "刘爱珠", "崔楠", "靳子晴", "马魏", "赵九童", "李玉娇", "张东阳", "张佳豪", "杜佳莹", "张学强", "韩思琪", "刘可", "孙泽宇", "李钰", "王欣然", "宋昊威", "王瑶")


End Sub

Private Sub Label2_Click()

End Sub

Private Sub Timer1_Timer()
If Me.WindowsMediaPlayer1.playState = 1 Then '
WindowsMediaPlayer1.URL = "C:\Users\Administrator\Desktop\抽奖程序（勿删）\成品\好运来（高品）.mp3"
WindowsMediaPlayer1.Controls.play
End If

End Sub

Private Sub Timer2_Timer()
s = Int(Rnd * 57)
Text1.Text = a(s)

End Sub
