VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   6315
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8910
   LinkTopic       =   "Form2"
   ScaleHeight     =   6315
   ScaleWidth      =   8910
   StartUpPosition =   3  '窗口缺省
   Visible         =   0   'False
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   27.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      TabIndex        =   10
      Text            =   "0"
      Top             =   4080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "下一盘"
      Height          =   615
      Left            =   7320
      TabIndex        =   8
      Top             =   5280
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "输完啦！！！"
      Height          =   615
      Left            =   5880
      TabIndex        =   7
      Top             =   5280
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   510
      Left            =   5280
      TabIndex        =   6
      Top             =   4320
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   360
      TabIndex        =   5
      Top             =   5160
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   360
      Top             =   3600
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   360
      Top             =   2880
   End
   Begin VB.Frame Frame1 
      Caption         =   "Time"
      Height          =   1575
      Left            =   2280
      TabIndex        =   2
      Top             =   2160
      Visible         =   0   'False
      Width           =   4095
      Begin VB.Label Label3 
         Caption         =   "5"
         Height          =   615
         Left            =   2880
         TabIndex        =   4
         Top             =   720
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "剩余时间（s）"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   3
         Top             =   600
         Visible         =   0   'False
         Width           =   2055
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始决战！！！"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   1560
      Width           =   3975
   End
   Begin VB.Label Label4 
      Caption         =   "得分"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   9
      Top             =   3960
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "运算比赛：                请在指定时间内输入指定位数"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   7215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Private Sub Command2_Click()

b = Val(Text2.Text)

c = Int(Log(Val(Text1.Text)) / Log(10))

d = c + 1

If d = b Then

MsgBox ("太棒了，你胜利了！！！加 3 分！！！")

Timer1.Enabled = False
Timer2.Enabled = False

Command3.Visible = True
Command2.Visible = False

e = Val(Text3.Text)
e = e + 3
Text3.Text = e

Else

MsgBox ("你输啦！！！准备接受惩罚吧！！！！算π算死你！！！！！！！！！！！！")

Dim sum As Double, t As Double
sum = 0: t = 1
Do
sum = sum + (-1) ^ (t + 1) / (2 * t - 1)
t = t + 1
Loop Until Abs((-1) ^ (t + 1) / (2 * t - 1)) < 10 ^ (-10)
MsgBox 4 * sum

End If

End Sub


Private Sub Command3_Click()

Timer1.Enabled = True
Timer2.Enabled = True

Text2.Text = "0"

Text1.Text = ""
Text2.Text = Fix(Rnd() * 12)

Command3.Visible = False
Command2.Visible = True

x = Val(Label3.Caption)
x = x + 3
Label3.Caption = x


End Sub

Private Sub Timer1_Timer()

a = Val(Label3.Caption)

a = a - 1

Label3.Caption = a

End Sub


Private Sub Command1_Click()

Timer1.Enabled = True
Timer2.Enabled = True

Frame1.Visible = True
Label2.Visible = True
Label3.Visible = True
Label4.Visible = True

Text1.Visible = True
Text2.Visible = True
Text3.Visible = True

Command2.Visible = True


Text2.Text = Fix(Rnd() * 12)


End Sub

Private Sub Timer2_Timer()

If Val(Label3.Caption) = 0 Then

MsgBox ("时间到！准备接受惩罚吧！！！算π算死你！！！！！！！！！！！")
Timer1.Enabled = False
Timer2.Enabled = False

Dim sum As Double, t As Double
sum = 0: t = 1
Do
sum = sum + (-1) ^ (t + 1) / (2 * t - 1)
t = t + 1
Loop Until Abs((-1) ^ (t + 1) / (2 * t - 1)) < 10 ^ (-10)
MsgBox 4 * sum

End If


End Sub
