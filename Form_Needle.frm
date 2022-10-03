VERSION 5.00
Begin VB.Form Form_Needle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "침적 체크"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2955
   Icon            =   "Form_Needle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   2955
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check2 
      Caption         =   "ALL CHECK"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   3840
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   375
      Left            =   240
      TabIndex        =   25
      Top             =   4320
      Width           =   2535
   End
   Begin VB.CheckBox Check1 
      Caption         =   "25"
      Height          =   255
      Index           =   24
      Left            =   2160
      TabIndex        =   24
      Top             =   1680
      Width           =   495
   End
   Begin VB.CheckBox Check1 
      Caption         =   "24"
      Height          =   255
      Index           =   23
      Left            =   2160
      TabIndex        =   23
      Top             =   1320
      Width           =   495
   End
   Begin VB.CheckBox Check1 
      Caption         =   "23"
      Height          =   255
      Index           =   22
      Left            =   2160
      TabIndex        =   22
      Top             =   960
      Width           =   495
   End
   Begin VB.CheckBox Check1 
      Caption         =   "22"
      Height          =   255
      Index           =   21
      Left            =   2160
      TabIndex        =   21
      Top             =   600
      Width           =   495
   End
   Begin VB.CheckBox Check1 
      Caption         =   "21"
      Height          =   255
      Index           =   20
      Left            =   2160
      TabIndex        =   20
      Top             =   240
      Width           =   495
   End
   Begin VB.CheckBox Check1 
      Caption         =   "20"
      Height          =   255
      Index           =   19
      Left            =   1200
      TabIndex        =   19
      Top             =   3480
      Width           =   495
   End
   Begin VB.CheckBox Check1 
      Caption         =   "19"
      Height          =   255
      Index           =   18
      Left            =   1200
      TabIndex        =   18
      Top             =   3120
      Width           =   495
   End
   Begin VB.CheckBox Check1 
      Caption         =   "18"
      Height          =   255
      Index           =   17
      Left            =   1200
      TabIndex        =   17
      Top             =   2760
      Width           =   495
   End
   Begin VB.CheckBox Check1 
      Caption         =   "17"
      Height          =   255
      Index           =   16
      Left            =   1200
      TabIndex        =   16
      Top             =   2400
      Width           =   495
   End
   Begin VB.CheckBox Check1 
      Caption         =   "16"
      Height          =   255
      Index           =   15
      Left            =   1200
      TabIndex        =   15
      Top             =   2040
      Width           =   495
   End
   Begin VB.CheckBox Check1 
      Caption         =   "15"
      Height          =   255
      Index           =   14
      Left            =   1200
      TabIndex        =   14
      Top             =   1680
      Width           =   495
   End
   Begin VB.CheckBox Check1 
      Caption         =   "14"
      Height          =   255
      Index           =   13
      Left            =   1200
      TabIndex        =   13
      Top             =   1320
      Width           =   495
   End
   Begin VB.CheckBox Check1 
      Caption         =   "13"
      Height          =   255
      Index           =   12
      Left            =   1200
      TabIndex        =   12
      Top             =   960
      Width           =   495
   End
   Begin VB.CheckBox Check1 
      Caption         =   "12"
      Height          =   255
      Index           =   11
      Left            =   1200
      TabIndex        =   11
      Top             =   600
      Width           =   495
   End
   Begin VB.CheckBox Check1 
      Caption         =   "11"
      Height          =   255
      Index           =   10
      Left            =   1200
      TabIndex        =   10
      Top             =   240
      Width           =   495
   End
   Begin VB.CheckBox Check1 
      Caption         =   "10"
      Height          =   255
      Index           =   9
      Left            =   240
      TabIndex        =   9
      Top             =   3480
      Width           =   495
   End
   Begin VB.CheckBox Check1 
      Caption         =   "9"
      Height          =   255
      Index           =   8
      Left            =   240
      TabIndex        =   8
      Top             =   3120
      Width           =   495
   End
   Begin VB.CheckBox Check1 
      Caption         =   "8"
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   7
      Top             =   2760
      Width           =   495
   End
   Begin VB.CheckBox Check1 
      Caption         =   "7"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   6
      Top             =   2400
      Width           =   495
   End
   Begin VB.CheckBox Check1 
      Caption         =   "6"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   5
      Top             =   2040
      Width           =   495
   End
   Begin VB.CheckBox Check1 
      Caption         =   "5"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   495
   End
   Begin VB.CheckBox Check1 
      Caption         =   "4"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   495
   End
   Begin VB.CheckBox Check1 
      Caption         =   "3"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   495
   End
   Begin VB.CheckBox Check1 
      Caption         =   "2"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   495
   End
   Begin VB.CheckBox Check1 
      Caption         =   "1"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "Form_Needle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'[ 2022.07.29 ] : 침적확인 관련 추가
Private Sub Check2_Click()
    If Check2.value = 1 Then                '침적체크 on
        For i = 0 To 24
            Check1(i).value = 1
        Next i
    Else                                    '침적체크 off
        For i = 0 To 24
            Check1(i).value = 0
        Next i
    End If
End Sub

Private Sub Command2_Click()
    Dim i As Integer
    Dim sfilename As String             '파일이름
    Dim ifreefile As Integer            '파일처리열 확인
    Dim chk_cnt As Integer
    
    '[ 2022.09.30 ] : 체크를 하나도 안한 경우 경고 메시지 출력
    chk_cnt = 0
    For i = 0 To 24
        If Check1(i).value = 1 Then
            chk_cnt = chk_cnt + 1
        End If
    Next i
    If chk_cnt = 0 Then
        MsgBox "침적확인 넘버가 설정되지 않았습니다.", vbCritical
        Exit Sub
    End If
       
    '저장파일경로
    sfilename = "c:\star probe\Needle_Chk.dat"
    ifreefile = FreeFile
    
    '설정내용을 파일로 저장한다.
    Open sfilename For Output As ifreefile
        For i = 0 To 24
            If Check1(i).value = 1 Then
                Needle_Chk(i) = True
            Else
                Needle_Chk(i) = False
            End If
            Print #ifreefile, Needle_Chk(i)
        Next i
    Close ifreefile
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    For i = 0 To 24
        If Needle_Chk(i) = True Then
            Check1(i).value = 1
        Else
            Check1(i).value = 0
        End If
    Next i
End Sub
