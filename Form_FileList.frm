VERSION 5.00
Begin VB.Form Form_FileList 
   BorderStyle     =   1  '단일 고정
   Caption         =   "Map List"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5670
   Icon            =   "Form_FileList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   5670
   StartUpPosition =   1  '소유자 가운데
   Begin VB.CommandButton Command1 
      Caption         =   "확인"
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   1920
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   1680
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "Form_FileList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim idx As Integer
    
'    If List1.Selected = False Then
'        MsgBox "파일을 선택하세요."
'        Exit Sub
'    End If
    idx = List1.ListIndex
    Load_MAP = List1.list(idx)
    Unload Me
End Sub
