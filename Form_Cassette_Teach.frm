VERSION 5.00
Begin VB.Form Form_Cassette_Teach 
   BorderStyle     =   1  '단일 고정
   Caption         =   "Cassette"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6990
   Icon            =   "Form_Cassette_Teach.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   6990
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CheckBox Check8 
      BackColor       =   &H00FFC0C0&
      Caption         =   "ALL CHECK"
      Height          =   615
      Left            =   3720
      Style           =   1  '그래픽
      TabIndex        =   58
      Top             =   4800
      Width           =   1290
   End
   Begin VB.CheckBox Check7 
      Caption         =   "AUTO ALIGN OFF"
      Height          =   375
      Left            =   240
      TabIndex        =   57
      Top             =   4920
      Value           =   1  '확인
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      Caption         =   "Slot Empty Check"
      Height          =   4575
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6735
      Begin VB.CheckBox Check6 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Line check"
         Height          =   465
         Left            =   5400
         Style           =   1  '그래픽
         TabIndex        =   56
         Top             =   3960
         Width           =   1170
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 50"
         Height          =   255
         Index           =   49
         Left            =   5640
         TabIndex        =   55
         Top             =   3600
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 49"
         Height          =   255
         Index           =   48
         Left            =   5640
         TabIndex        =   54
         Top             =   3240
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 48"
         Height          =   255
         Index           =   47
         Left            =   5640
         TabIndex        =   53
         Top             =   2880
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 47"
         Height          =   255
         Index           =   46
         Left            =   5640
         TabIndex        =   52
         Top             =   2520
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 46"
         Height          =   255
         Index           =   45
         Left            =   5640
         TabIndex        =   51
         Top             =   2160
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 45"
         Height          =   255
         Index           =   44
         Left            =   5640
         TabIndex        =   50
         Top             =   1800
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 44"
         Height          =   255
         Index           =   43
         Left            =   5640
         TabIndex        =   49
         Top             =   1440
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 43"
         Height          =   255
         Index           =   42
         Left            =   5640
         TabIndex        =   48
         Top             =   1080
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 42"
         Height          =   255
         Index           =   41
         Left            =   5640
         TabIndex        =   47
         Top             =   720
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 41"
         Height          =   255
         Index           =   40
         Left            =   5640
         TabIndex        =   46
         Top             =   360
         Width           =   615
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Line check"
         Height          =   465
         Left            =   4080
         Style           =   1  '그래픽
         TabIndex        =   45
         Top             =   3960
         Width           =   1170
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 40"
         Height          =   255
         Index           =   39
         Left            =   4320
         TabIndex        =   44
         Top             =   3600
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 39"
         Height          =   255
         Index           =   38
         Left            =   4320
         TabIndex        =   43
         Top             =   3240
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 38"
         Height          =   255
         Index           =   37
         Left            =   4320
         TabIndex        =   42
         Top             =   2880
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 37"
         Height          =   255
         Index           =   36
         Left            =   4320
         TabIndex        =   41
         Top             =   2520
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 36"
         Height          =   255
         Index           =   35
         Left            =   4320
         TabIndex        =   40
         Top             =   2160
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 35"
         Height          =   255
         Index           =   34
         Left            =   4320
         TabIndex        =   39
         Top             =   1800
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 34"
         Height          =   255
         Index           =   33
         Left            =   4320
         TabIndex        =   38
         Top             =   1440
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 33"
         Height          =   255
         Index           =   32
         Left            =   4320
         TabIndex        =   37
         Top             =   1080
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 32"
         Height          =   255
         Index           =   31
         Left            =   4320
         TabIndex        =   36
         Top             =   720
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 31"
         Height          =   255
         Index           =   30
         Left            =   4320
         TabIndex        =   35
         Top             =   360
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 10"
         Height          =   255
         Index           =   9
         Left            =   360
         TabIndex        =   34
         Top             =   3600
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 9"
         Height          =   255
         Index           =   8
         Left            =   360
         TabIndex        =   33
         Top             =   3240
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 8"
         Height          =   255
         Index           =   7
         Left            =   360
         TabIndex        =   32
         Top             =   2880
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 7"
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   31
         Top             =   2520
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 6"
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   30
         Top             =   2160
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 5"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   29
         Top             =   1800
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 4"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   28
         Top             =   1440
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 3"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   27
         Top             =   1080
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 2"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   26
         Top             =   720
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 1"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   25
         Top             =   360
         Width           =   615
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Line Check"
         Height          =   465
         Left            =   120
         Style           =   1  '그래픽
         TabIndex        =   24
         Top             =   3960
         Width           =   1170
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 12"
         Height          =   255
         Index           =   11
         Left            =   1680
         TabIndex        =   23
         Top             =   720
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 11"
         Height          =   255
         Index           =   10
         Left            =   1680
         TabIndex        =   22
         Top             =   360
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 20"
         Height          =   255
         Index           =   19
         Left            =   1680
         TabIndex        =   21
         Top             =   3600
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 19"
         Height          =   255
         Index           =   18
         Left            =   1680
         TabIndex        =   20
         Top             =   3240
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 18"
         Height          =   255
         Index           =   17
         Left            =   1680
         TabIndex        =   19
         Top             =   2880
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 17"
         Height          =   255
         Index           =   16
         Left            =   1680
         TabIndex        =   18
         Top             =   2520
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 16"
         Height          =   255
         Index           =   15
         Left            =   1680
         TabIndex        =   17
         Top             =   2160
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 15"
         Height          =   255
         Index           =   14
         Left            =   1680
         TabIndex        =   16
         Top             =   1800
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 14"
         Height          =   255
         Index           =   13
         Left            =   1680
         TabIndex        =   15
         Top             =   1440
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 13"
         Height          =   255
         Index           =   12
         Left            =   1680
         TabIndex        =   14
         Top             =   1080
         Width           =   615
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Line check"
         Height          =   465
         Left            =   1440
         Style           =   1  '그래픽
         TabIndex        =   13
         Top             =   3960
         Width           =   1170
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 24"
         Height          =   255
         Index           =   23
         Left            =   3000
         TabIndex        =   12
         Top             =   1440
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 23"
         Height          =   255
         Index           =   22
         Left            =   3000
         TabIndex        =   11
         Top             =   1080
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 22"
         Height          =   255
         Index           =   21
         Left            =   3000
         TabIndex        =   10
         Top             =   720
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 21"
         Height          =   255
         Index           =   20
         Left            =   3000
         TabIndex        =   9
         Top             =   360
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 25"
         Height          =   255
         Index           =   24
         Left            =   3000
         TabIndex        =   8
         Top             =   1800
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 26"
         Height          =   255
         Index           =   25
         Left            =   3000
         TabIndex        =   7
         Top             =   2160
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 27"
         Height          =   255
         Index           =   26
         Left            =   3000
         TabIndex        =   6
         Top             =   2520
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 28"
         Height          =   255
         Index           =   27
         Left            =   3000
         TabIndex        =   5
         Top             =   2880
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 29"
         Height          =   255
         Index           =   28
         Left            =   3000
         TabIndex        =   4
         Top             =   3240
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 30"
         Height          =   255
         Index           =   29
         Left            =   3000
         TabIndex        =   3
         Top             =   3600
         Width           =   615
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Line check"
         Height          =   465
         Left            =   2760
         Style           =   1  '그래픽
         TabIndex        =   2
         Top             =   3960
         Width           =   1170
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      Style           =   1  '그래픽
      TabIndex        =   0
      Top             =   4800
      Width           =   1695
   End
End
Attribute VB_Name = "Form_Cassette_Teach"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check2_Click()
    If Check2.value = 1 Then
        For i = 0 To 9
            Check1(i).value = 1
        Next i
    Else
        For i = 0 To 9
            Check1(i).value = 0
        Next i
    End If
End Sub

Private Sub Check3_Click()
    If Check3.value = 1 Then
        For i = 10 To 19
            Check1(i).value = 1
        Next i
    Else
        For i = 10 To 19
            Check1(i).value = 0
        Next i
    End If
End Sub

Private Sub Check4_Click()
    If Check4.value = 1 Then
        For i = 20 To 29
            Check1(i).value = 1
        Next i
    Else
        For i = 20 To 29
            Check1(i).value = 0
        Next i
    End If
End Sub

Private Sub Check5_Click()
    If Check5.value = 1 Then
        For i = 30 To 39
            Check1(i).value = 1
        Next i
    Else
        For i = 30 To 39
            Check1(i).value = 0
        Next i
    End If
End Sub

Private Sub Check6_Click()
    If Check6.value = 1 Then
        For i = 40 To 49
            Check1(i).value = 1
        Next i
    Else
        For i = 40 To 49
            Check1(i).value = 0
        Next i
    End If
End Sub

Private Sub Check7_Click()
    If Check7.value = 0 Then
        AutoAlign_Flag = False
        Check7.Caption = "AUTO ALIGH OFF"
    Else
        AutoAlign_Flag = True
        Check7.Caption = "AUTO ALIGH ON"
    End If
End Sub

Private Sub Check8_Click()
    If Check8.value = 1 Then
        For i = 0 To 49
            Check1(i).value = 1
        Next i
    Else
        For i = 0 To 49
            Check1(i).value = 0
        Next i
    End If
End Sub

Private Sub Command1_Click()
    For i = 0 To 49
        If Check1(i).value = 1 Then
            Slot_No(i + 1) = True
        Else
            Slot_No(i + 1) = False
            Slot_Max_Count = i + 1
        End If
    Next i
    Unload Me
End Sub

Private Sub Form_Load()
    For i = 0 To 49
        If Slot_No(i + 1) = True Then
            Check1(i).value = 1
        Else
            Check1(i).value = 0
        End If
    Next i
    
    If AutoAlign_Flag = True Then
        Check7.value = 1
    Else
        Check7.value = 0
    End If
End Sub
