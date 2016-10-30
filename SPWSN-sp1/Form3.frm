VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   5565
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8160
   LinkTopic       =   "Form3"
   ScaleHeight     =   5565
   ScaleWidth      =   8160
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      Caption         =   "Time"
      Height          =   2175
      Left            =   2280
      TabIndex        =   1
      Top             =   1200
      Width           =   3255
      Begin VB.Label Label2 
         Caption         =   "1"
         Height          =   615
         Left            =   2040
         TabIndex        =   3
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "剩余时间"
         Height          =   735
         Left            =   840
         TabIndex        =   2
         Top             =   1080
         Width           =   1215
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   1320
      Top             =   600
   End
   Begin VB.TextBox Text1 
      Height          =   975
      Left            =   960
      TabIndex        =   0
      Top             =   3840
      Width           =   6255
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
