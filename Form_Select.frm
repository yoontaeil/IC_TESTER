VERSION 5.00
Begin VB.Form Form_Select 
   BorderStyle     =   1  '단일 고정
   Caption         =   "Select Program"
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6315
   Icon            =   "Form_Select.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   6315
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select"
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      Begin VB.OptionButton Option1 
         Caption         =   "EAGLE(E4090)"
         Height          =   375
         Index           =   3
         Left            =   360
         TabIndex        =   5
         Top             =   1440
         Width           =   5415
      End
      Begin VB.OptionButton Option1 
         Caption         =   "EAGLE(2001X)"
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   4
         Top             =   1080
         Width           =   5415
      End
      Begin VB.OptionButton Option1 
         Caption         =   "AMT-88"
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   3
         Top             =   720
         Width           =   5415
      End
      Begin VB.OptionButton Option1 
         Caption         =   "ACCO Tester"
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   5415
      End
   End
End
Attribute VB_Name = "Form_Select"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Form_Login.Show
    
    Unload Me
End Sub

Private Sub Option1_Click(Index As Integer)
    Select Case Index
        Case 0:
            Model_Select = 1                   'ACCO
        Case 1:
            Model_Select = 2                   'AMT-88
        Case 2:
            Model_Select = 3                   'EAGLE(2001X)
        Case 3:
            Model_Select = 4                   'EAGLE(E4090)
    End Select
End Sub
