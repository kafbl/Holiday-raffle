VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5010
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6240
   LinkTopic       =   "Form1"
   ScaleHeight     =   5010
   ScaleWidth      =   6240
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text2 
      Height          =   1215
      Left            =   3120
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   960
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   1215
      Left            =   480
      TabIndex        =   3
      Top             =   960
      Width           =   2535
   End
   Begin VB.Timer Timer1 
      Left            =   5640
      Top             =   360
   End
   Begin VB.CommandButton Command3 
      Caption         =   "退出&E"
      Height          =   735
      Left            =   4200
      TabIndex        =   2
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "停止&S"
      Height          =   735
      Left            =   2280
      TabIndex        =   1
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始&O"
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   3600
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim j As Integer
Dim a() As String

Private Sub Command1_Click()

Timer1.Enabled = True

End Sub

Private Sub Command2_Click()
Dim b()

Dim i, k As Integer

Timer1.Enabled = False

For i = 0 To UBound(a)


If Text1.Text = a(i) Then

Else

ReDim Preserve b(k)

b(k) = a(i)

k = k + 1

End If

Next

ReDim a(j - 1)

For i = 0 To UBound(b)

a(i) = b(i)

Next

j = j - 1

Text2.Text = Text2.Text & Text1.Text & vbCrLf

End Sub

Private Sub Command3_Click()

End

End Sub

Private Sub Form_Load()
Timer1.Enabled = False
Text1.Text = ""
Command1.Caption = "开始"
Command2.Caption = "停止"
Command3.Caption = "退出"
Timer1.Interval = 500
j = 5
ReDim a(57)
a(0) = "董安宁"
a(1) = "郝泉浩"
a(2) = "苏纬振"
a(3) = "梁欣雨"
a(4) = "赵文龙"
a(5) = "牛向征"
a(6) = "孙一鸣"
a(7) = "刘尚武"
a(8) = "高家旺"
a(9) = "赵金酉"
a(10) = "卢玉龙"
a(11) = "牛金帅"
a(12) = "陈瀚"
a(13) = "姚家欣"
a(14) = "徐  博"
a(15) = "于海宽"
a(16) = "贾姜华"
a(17) = "宿明丽"
a(18) = "郭  凡"
a(19) = "赵  迪"
a(20) = "贾军杰"
a(21) = "许世超"
a(22) = "张  浩"
a(23) = "杜佳坤"
a(24) = "杜鹏飞"
a(25) = "赵一曼"
a(26) = "刘洁林"
a(27) = "王兴原"
a(28) = "赵珈玲"
a(29) = "管明明"
a(30) = "崔  焱"
a(31) = "安世吉"
a(32) = "车  轩"
a(33) = "赵志禹"
a(34) = "于鑫童"
a(35) = "王田锦"
a(36) = "耿家蕊"
a(37) = "郝塞娅"
a(38) = "赵建鑫"
a(39) = "戴文杰"
a(40) = "刘爱珠"
a(41) = "崔  楠"
a(42) = "靳子晴"
a(43) = "马  魏"
a(44) = "赵九童"
a(45) = "李玉娇"
a(46) = "张东阳"
a(47) = "张佳豪"
a(48) = "杜佳莹"
a(49) = "张学强"
a(50) = "韩思琪"
a(51) = "刘  可"
a(52) = "孙泽宇"
a(53) = "李  钰"
a(54) = "王欣然"
a(55) = "宋昊威"
a(56) = "王  瑶"




Text2.Text = "已抽出名单" & vbCrLf




End Sub
