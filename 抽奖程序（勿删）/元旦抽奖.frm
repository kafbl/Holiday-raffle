VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form1 
   Caption         =   "��Ԫ�����֣�"
   ClientHeight    =   6270
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11265
   ForeColor       =   &H000000FF&
   Icon            =   "Ԫ���齱.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Ԫ���齱.frx":25CA
   ScaleHeight     =   6270
   ScaleWidth      =   11265
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Timer Timer2 
      Left            =   10200
      Top             =   1080
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   3120
      TabIndex        =   4
      Top             =   2640
      Width           =   4455
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   10200
      Top             =   480
   End
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Left            =   8640
      Picture         =   "Ԫ���齱.frx":E9DA
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   2
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "��ʼ�齱&A"
      BeginProperty Font 
         Name            =   "���Ĳ���"
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
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      DragMode        =   1  'Automatic
      Height          =   1215
      Left            =   2760
      TabIndex        =   3
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
      Caption         =   "������Ԫ�����֣�����"
      BeginProperty Font 
         Name            =   "������"
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
Dim a(57)


Private Sub Command1_Click()
Dim a
Dim b(4) As Integer
Dim i, j, k
Dim bl As Boolean
a = Array("������", "��Ȫ��", "��γ��", "������", "������", "ţ����", "��һ��", "������", "�߼���", "�Խ���", "¬����", "ţ��˧", "���", "Ҧ����", "�첩", "�ں���", "�ֽ���", "������", "����", "�Ե�", "�־���", "������", "�ź�", "�ż���", "������", "��һ��", "������", "����ԭ", "������", "������", "����", "������", "����", "��־��", "����ͯ", "�����", "������", "�����", "�Խ���", "���Ľ�", "������", "���", "������", "��κ", "�Ծ�ͯ", "����", "�Ŷ���", "�żѺ�", "�ż�Ө", "��ѧǿ", "��˼��", "����", "������", "����", "����Ȼ", "�����", "����")

 
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

MsgBox "����������ϲ��������" & vbCrLf & s



End Sub

Private Sub dcButton1_Click()

End Sub

Private Sub Form_Load()
Timer1.Enabled = True
Timer2.Enabled = True

WindowsMediaPlayer1.URL = "C:\Users\Administrator\Desktop\����������Ʒ��.mp3"
WindowsMediaPlayer1.Enabled = ture

Form1.Caption = "������Ԫ�����֣�����"
End Sub

Private Sub Label2_Click()

End Sub

Private Sub Timer1_Timer()
If Me.WindowsMediaPlayer1.playState = 1 Then '
WindowsMediaPlayer1.URL = "C:\Users\Administrator\Desktop\����������Ʒ��.mp3"
WindowsMediaPlayer1.Controls.play
End If
End Sub

Private Sub Timer2_Timer()
Text1.Text = Int(rand() * 57)

End Sub
