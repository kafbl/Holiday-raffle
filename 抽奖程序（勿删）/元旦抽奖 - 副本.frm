VERSION 5.00
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "������Ԫ�����֣�����"
      BeginProperty Font 
         Name            =   "������"
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

Private Sub Form_Load()

Form1.Caption = "������Ԫ�����֣�����"
End Sub

