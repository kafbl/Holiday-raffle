VERSION 5.00
Begin VB.Form Form1 
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
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Left            =   8640
      Picture         =   "元旦抽奖.frx":E9DA
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   2
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "开始抽奖&A"
      BeginProperty Font 
         Name            =   "华文彩云"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3720
      MaskColor       =   &H0080FFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4680
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "！！！元旦快乐！！！"
      BeginProperty Font 
         Name            =   "新宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1815
      Left            =   1440
      TabIndex        =   1
      Top             =   1320
      Width           =   8055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim a
Dim b(4) As Integer
Dim i, j, k
Dim bl As Boolean
a = Array("董安宁", "郝泉浩", "苏纬振", "梁欣雨", "赵文龙", "牛向征", "孙一鸣", "刘尚武", "高家旺", "赵金酉", "卢玉龙", "牛金帅", "陈瀚", "姚家欣", "徐博", "于海宽", "贾姜华", "宿明丽", "郭凡", "赵迪", "贾军杰", "许世超", "张浩", "杜佳坤", "杜鹏飞", "赵一曼", "刘洁林", "王兴原", "赵珈玲", "管明明", "崔焱", "安世吉", "车轩", "赵志禹", "于鑫童", "王田锦", "耿家蕊", "郝塞娅", "赵建鑫", "戴文杰", "刘爱珠", "崔楠", "靳子晴", "马魏", "赵九童", "李玉娇", "张东阳", "张佳豪", "杜佳莹", "张学强", "韩思琪", "刘可", "孙泽宇", "李钰", "王欣然", "宋昊威", "王瑶")

 
Randomize
k = Int(UBound(a) * Rnd)
 b(j) = k
Do
bl = False
Randomize
k = Int(UBound(a) * Rnd)
For i = 0 To 4
If k = b(i) Then
bl = True
Exit For
End If
Next
If bl = False Then
j = j + 1
b(j) = k
End If
Loop Until j = 4
For i = 0 To 0
s = s & a(b(i)) & vbCrLf
Next

MsgBox "！！！！恭喜！！！！" & vbCrLf & s



End Sub

Private Sub Form_Load()

Form1.Caption = "！！！元旦快乐！！！"
End Sub

