VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6285
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   9270
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command5 
      Caption         =   "�������"
      Height          =   495
      Left            =   3600
      TabIndex        =   6
      Top             =   3720
      Width           =   2655
   End
   Begin VB.CommandButton Command4 
      Caption         =   "���ʹ���"
      Height          =   495
      Left            =   6720
      TabIndex        =   5
      Top             =   5400
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "���"
      Height          =   495
      Left            =   3840
      TabIndex        =   4
      Top             =   5400
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�˳�"
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   5400
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��������"
      Height          =   495
      Left            =   480
      MaskColor       =   &H000000FF&
      TabIndex        =   2
      Top             =   3720
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      Height          =   3135
      Left            =   3960
      ScaleHeight     =   3075
      ScaleWidth      =   5115
      TabIndex        =   0
      Top             =   120
      Width           =   5175
   End
   Begin VB.Label Label1 
      Caption         =   "spWSN   ר�û�Ա������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim sum As Double, t As Double
sum = 0: t = 1
Do
sum = sum + (-1) ^ (t + 1) / (2 * t - 1)
t = t + 1
Loop Until Abs((-1) ^ (t + 1) / (2 * t - 1)) < 10 ^ (-8)
MsgBox 4 * sum

End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()
a = MsgBox("�Ѿ�����Ʒ 100 �� �ǳ���л����֧�֣�����", vbDefaultButton3, "���׳ɹ�")
End Sub

Private Sub Command4_Click()
Dim x As Single
Dim y As Single

b = MsgBox("������ʼ���ӣ��밴��ȷ��������", vbDefaultButton2, "ȷ������")

x = 1
y = 1
Do Until x > 10
  Do Until y = 10
  y = y + 0.1
  Loop
x = x + 0.1
Loop
  
a = MsgBox("�������������ӳ�ʱ��", vbDefaultButton3, "����ʧ��")

End Sub

Private Sub Command5_Click()
Form2.Visible = True

End Sub
