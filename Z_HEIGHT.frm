VERSION 5.00
Begin VB.Form Z_HEIGHT 
   Caption         =   "INK TEST"
   ClientHeight    =   2655
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3045
   Icon            =   "Z_HEIGHT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   3045
   StartUpPosition =   1  '소유자 가운데
   Begin VB.CommandButton Command2 
      Caption         =   "INK RUN && Close"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   2775
   End
   Begin VB.CommandButton Command6 
      Caption         =   "INK Dot Test"
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Z Height Set"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  '가운데 맞춤
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Text            =   "3000"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "2000 ~ 4000"
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   4920
      Width           =   1095
   End
End
Attribute VB_Name = "Z_HEIGHT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    INK_OFF_TEST = False            '2015.11.30
    Unload Me
End Sub

Private Sub Command2_Click()
    INK_OFF_TEST = True             '2015.11.30
    Unload Me
End Sub

Private Sub Command5_Click()
'    StarProbe_Z_Height (val(Text7.Text))
End Sub

Private Sub Command6_Click()
    Call StarProbe_Left_Ink_Dot(StarProbe.Ink_LeftPort)
End Sub

Private Sub Form_Load()
'    Text7.Text = StarProbe_Z_Current
End Sub
