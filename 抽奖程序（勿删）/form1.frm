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
   StartUpPosition =   3  '����ȱʡ
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
      Caption         =   "�˳�&E"
      Height          =   735
      Left            =   4200
      TabIndex        =   2
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ֹͣ&S"
      Height          =   735
      Left            =   2280
      TabIndex        =   1
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��ʼ&O"
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
Command1.Caption = "��ʼ"
Command2.Caption = "ֹͣ"
Command3.Caption = "�˳�"
Timer1.Interval = 500
j = 5
ReDim a(57)
a(0) = "������"
a(1) = "��Ȫ��"
a(2) = "��γ��"
a(3) = "������"
a(4) = "������"
a(5) = "ţ����"
a(6) = "��һ��"
a(7) = "������"
a(8) = "�߼���"
a(9) = "�Խ���"
a(10) = "¬����"
a(11) = "ţ��˧"
a(12) = "���"
a(13) = "Ҧ����"
a(14) = "��  ��"
a(15) = "�ں���"
a(16) = "�ֽ���"
a(17) = "������"
a(18) = "��  ��"
a(19) = "��  ��"
a(20) = "�־���"
a(21) = "������"
a(22) = "��  ��"
a(23) = "�ż���"
a(24) = "������"
a(25) = "��һ��"
a(26) = "������"
a(27) = "����ԭ"
a(28) = "������"
a(29) = "������"
a(30) = "��  ��"
a(31) = "������"
a(32) = "��  ��"
a(33) = "��־��"
a(34) = "����ͯ"
a(35) = "�����"
a(36) = "������"
a(37) = "�����"
a(38) = "�Խ���"
a(39) = "���Ľ�"
a(40) = "������"
a(41) = "��  �"
a(42) = "������"
a(43) = "��  κ"
a(44) = "�Ծ�ͯ"
a(45) = "����"
a(46) = "�Ŷ���"
a(47) = "�żѺ�"
a(48) = "�ż�Ө"
a(49) = "��ѧǿ"
a(50) = "��˼��"
a(51) = "��  ��"
a(52) = "������"
a(53) = "��  ��"
a(54) = "����Ȼ"
a(55) = "�����"
a(56) = "��  ��"




Text2.Text = "�ѳ������" & vbCrLf




End Sub
