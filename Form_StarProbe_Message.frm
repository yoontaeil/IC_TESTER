VERSION 5.00
Begin VB.Form Form_StarProbe_Message 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "Form2"
   ClientHeight    =   600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5250
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   40
   ScaleMode       =   3  '픽셀
   ScaleWidth      =   350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   3900
      Top             =   60
   End
   Begin VB.Label Label2 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  '투명
      Caption         =   "123456789012345678901234567890"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   90
      TabIndex        =   1
      Top             =   90
      Width           =   4995
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00C0FFFF&
      Caption         =   "123456789012345678901234567890"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4995
   End
End
Attribute VB_Name = "Form_StarProbe_Message"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    Me.Caption = StarProbeMessage.Title
    
    Label1 = StarProbeMessage.Message
    Label2 = StarProbeMessage.Message
    
    Me.Refresh

End Sub

Private Sub Timer1_Timer()
    Label2.Visible = Not Label2.Visible
End Sub
