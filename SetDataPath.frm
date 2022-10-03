VERSION 5.00
Begin VB.Form SetDataPath 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "Data Server Config"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4320
   Icon            =   "SetDataPath.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   3000
      TabIndex        =   4
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   375
      Index           =   0
      Left            =   1680
      TabIndex        =   3
      Top             =   3840
      Width           =   1215
   End
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4095
   End
   Begin VB.DirListBox Dir1 
      Height          =   3030
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "Select Data Server Path"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "SetDataPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
    If Index = 0 Then
        If Path_Check = 1 Then
            SelectExt.Text4.Text = Dir1.Path
        ElseIf Path_Check = 2 Then
            SelectExt.Text5.Text = Dir1.Path
        ElseIf Path_Check = 3 Then
            SelectExt.Text10.Text = Dir1.Path
        End If
    End If
    Unload Me
End Sub

Private Sub Drive1_Change()
    On Error GoTo DriveErr
    Dir1.Path = Drive1.Drive
    Exit Sub
    
DriveErr:
    MsgBox "Can't Use drive!", 16
    Exit Sub
End Sub
