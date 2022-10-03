VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form_StarProbe_BINColor 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "BIN Color"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6810
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form_StarProbe_BINColor.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   457
   ScaleMode       =   3  '픽셀
   ScaleWidth      =   454
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.CommandButton Command_BINColor 
      Height          =   255
      Index           =   26
      Left            =   2280
      Style           =   1  '그래픽
      TabIndex        =   49
      Top             =   6480
      Width           =   735
   End
   Begin VB.CommandButton Command_BINColor 
      Height          =   255
      Index           =   25
      Left            =   2280
      Style           =   1  '그래픽
      TabIndex        =   48
      Top             =   6240
      Width           =   735
   End
   Begin VB.CommandButton Command_BINColor 
      Height          =   255
      Index           =   24
      Left            =   2280
      Style           =   1  '그래픽
      TabIndex        =   47
      Top             =   6000
      Width           =   735
   End
   Begin VB.CommandButton Command_BINColor 
      Height          =   255
      Index           =   23
      Left            =   2280
      Style           =   1  '그래픽
      TabIndex        =   46
      Top             =   5760
      Width           =   735
   End
   Begin VB.CommandButton Command_BINColor 
      Height          =   255
      Index           =   22
      Left            =   2280
      Style           =   1  '그래픽
      TabIndex        =   45
      Top             =   5520
      Width           =   735
   End
   Begin VB.CommandButton Command_BINColor 
      Height          =   255
      Index           =   21
      Left            =   2280
      Style           =   1  '그래픽
      TabIndex        =   44
      Top             =   5280
      Width           =   735
   End
   Begin VB.CommandButton Command_BINColor 
      Height          =   255
      Index           =   20
      Left            =   2280
      Style           =   1  '그래픽
      TabIndex        =   43
      Top             =   5040
      Width           =   735
   End
   Begin VB.CommandButton Command_BINColor 
      Height          =   255
      Index           =   19
      Left            =   2280
      Style           =   1  '그래픽
      TabIndex        =   42
      Top             =   4800
      Width           =   735
   End
   Begin VB.CommandButton Command_BINColor 
      Height          =   255
      Index           =   18
      Left            =   2280
      Style           =   1  '그래픽
      TabIndex        =   41
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton Command_BINColor 
      Height          =   255
      Index           =   17
      Left            =   2280
      Style           =   1  '그래픽
      TabIndex        =   40
      Top             =   4320
      Width           =   735
   End
   Begin VB.CommandButton Command_BINColor 
      Height          =   255
      Index           =   16
      Left            =   2280
      Style           =   1  '그래픽
      TabIndex        =   39
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton Command_BINColor 
      Height          =   255
      Index           =   15
      Left            =   2280
      Style           =   1  '그래픽
      TabIndex        =   38
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton Command_BINColor 
      Height          =   255
      Index           =   14
      Left            =   2280
      Style           =   1  '그래픽
      TabIndex        =   37
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton Command_BINColor 
      Height          =   255
      Index           =   13
      Left            =   2280
      Style           =   1  '그래픽
      TabIndex        =   36
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton Command_BINColor 
      Height          =   255
      Index           =   12
      Left            =   2280
      Style           =   1  '그래픽
      TabIndex        =   35
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton Command_BINColor 
      Height          =   255
      Index           =   11
      Left            =   2280
      Style           =   1  '그래픽
      TabIndex        =   34
      Top             =   2880
      Width           =   735
   End
   Begin VB.CommandButton Command_BINColor 
      Height          =   255
      Index           =   10
      Left            =   2280
      Style           =   1  '그래픽
      TabIndex        =   33
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton Command_BINColor 
      Height          =   255
      Index           =   9
      Left            =   2280
      Style           =   1  '그래픽
      TabIndex        =   32
      Top             =   2400
      Width           =   735
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4680
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label_BinNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Ink"
      Height          =   240
      Index           =   32
      Left            =   4920
      TabIndex        =   105
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label Label_BinCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "0"
      Height          =   240
      Index           =   26
      Left            =   3120
      TabIndex        =   104
      Top             =   6480
      Width           =   105
   End
   Begin VB.Label Label_BinCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "0"
      Height          =   240
      Index           =   25
      Left            =   3120
      TabIndex        =   103
      Top             =   6240
      Width           =   105
   End
   Begin VB.Label Label_BinCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "0"
      Height          =   240
      Index           =   24
      Left            =   3120
      TabIndex        =   102
      Top             =   6000
      Width           =   105
   End
   Begin VB.Label Label_BinCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "0"
      Height          =   240
      Index           =   23
      Left            =   3120
      TabIndex        =   101
      Top             =   5760
      Width           =   105
   End
   Begin VB.Label Label_BinCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "0"
      Height          =   240
      Index           =   22
      Left            =   3120
      TabIndex        =   100
      Top             =   5520
      Width           =   105
   End
   Begin VB.Label Label_BinCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "0"
      Height          =   240
      Index           =   21
      Left            =   3120
      TabIndex        =   99
      Top             =   5280
      Width           =   105
   End
   Begin VB.Label Label_BinCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "0"
      Height          =   240
      Index           =   20
      Left            =   3120
      TabIndex        =   98
      Top             =   5040
      Width           =   105
   End
   Begin VB.Label Label_BinCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "0"
      Height          =   240
      Index           =   19
      Left            =   3120
      TabIndex        =   97
      Top             =   4800
      Width           =   105
   End
   Begin VB.Label Label_BinCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "0"
      Height          =   240
      Index           =   18
      Left            =   3120
      TabIndex        =   96
      Top             =   4560
      Width           =   105
   End
   Begin VB.Label Label_BinCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "0"
      Height          =   240
      Index           =   17
      Left            =   3120
      TabIndex        =   95
      Top             =   4320
      Width           =   105
   End
   Begin VB.Label Label_BinCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "0"
      Height          =   240
      Index           =   16
      Left            =   3120
      TabIndex        =   94
      Top             =   4080
      Width           =   105
   End
   Begin VB.Label Label_BinCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "0"
      Height          =   240
      Index           =   15
      Left            =   3120
      TabIndex        =   93
      Top             =   3840
      Width           =   105
   End
   Begin VB.Label Label_BinCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "0"
      Height          =   240
      Index           =   14
      Left            =   3120
      TabIndex        =   92
      Top             =   3600
      Width           =   105
   End
   Begin VB.Label Label_BinCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "0"
      Height          =   240
      Index           =   13
      Left            =   3120
      TabIndex        =   91
      Top             =   3360
      Width           =   105
   End
   Begin VB.Label Label_BinCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "0"
      Height          =   240
      Index           =   12
      Left            =   3120
      TabIndex        =   90
      Top             =   3120
      Width           =   105
   End
   Begin VB.Label Label_BinCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "0"
      Height          =   240
      Index           =   11
      Left            =   3120
      TabIndex        =   89
      Top             =   2880
      Width           =   105
   End
   Begin VB.Label Label_BinCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "0"
      Height          =   240
      Index           =   10
      Left            =   3120
      TabIndex        =   88
      Top             =   2640
      Width           =   105
   End
   Begin VB.Label Label_BinCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "0"
      Height          =   240
      Index           =   9
      Left            =   3120
      TabIndex        =   87
      Top             =   2400
      Width           =   105
   End
   Begin VB.Label Label_BinCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "0"
      Height          =   240
      Index           =   8
      Left            =   3120
      TabIndex        =   86
      Top             =   2160
      Width           =   105
   End
   Begin VB.Label Label_BinCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "0"
      Height          =   240
      Index           =   7
      Left            =   3120
      TabIndex        =   85
      Top             =   1920
      Width           =   105
   End
   Begin VB.Label Label_BinCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "0"
      Height          =   240
      Index           =   6
      Left            =   3120
      TabIndex        =   84
      Top             =   1680
      Width           =   105
   End
   Begin VB.Label Label_BinCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "0"
      Height          =   240
      Index           =   5
      Left            =   3120
      TabIndex        =   83
      Top             =   1440
      Width           =   105
   End
   Begin VB.Label Label_BinCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "0"
      Height          =   240
      Index           =   4
      Left            =   3120
      TabIndex        =   82
      Top             =   1200
      Width           =   105
   End
   Begin VB.Label Label_BinCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "0"
      Height          =   240
      Index           =   3
      Left            =   3120
      TabIndex        =   81
      Top             =   960
      Width           =   105
   End
   Begin VB.Label Label_BinCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "0"
      Height          =   240
      Index           =   2
      Left            =   3120
      TabIndex        =   80
      Top             =   720
      Width           =   105
   End
   Begin VB.Label Label_BinCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "0"
      Height          =   240
      Index           =   1
      Left            =   3120
      TabIndex        =   79
      Top             =   480
      Width           =   105
   End
   Begin VB.Label Label_BinCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "0"
      Height          =   240
      Index           =   0
      Left            =   3120
      TabIndex        =   78
      Top             =   120
      Width           =   105
   End
   Begin VB.Label Label_BinCommand 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Height          =   240
      Index           =   26
      Left            =   960
      TabIndex        =   77
      Top             =   6480
      Width           =   60
   End
   Begin VB.Label Label_BinCommand 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Height          =   240
      Index           =   25
      Left            =   960
      TabIndex        =   76
      Top             =   6240
      Width           =   60
   End
   Begin VB.Label Label_BinCommand 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Height          =   240
      Index           =   24
      Left            =   960
      TabIndex        =   75
      Top             =   6000
      Width           =   60
   End
   Begin VB.Label Label_BinCommand 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Height          =   240
      Index           =   23
      Left            =   960
      TabIndex        =   74
      Top             =   5760
      Width           =   60
   End
   Begin VB.Label Label_BinCommand 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Height          =   240
      Index           =   22
      Left            =   960
      TabIndex        =   73
      Top             =   5520
      Width           =   540
   End
   Begin VB.Label Label_BinCommand 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Height          =   240
      Index           =   21
      Left            =   960
      TabIndex        =   72
      Top             =   5280
      Width           =   660
   End
   Begin VB.Label Label_BinCommand 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Height          =   240
      Index           =   20
      Left            =   960
      TabIndex        =   71
      Top             =   5040
      Width           =   60
   End
   Begin VB.Label Label_BinCommand 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Height          =   240
      Index           =   19
      Left            =   960
      TabIndex        =   70
      Top             =   4800
      Width           =   60
   End
   Begin VB.Label Label_BinCommand 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Height          =   240
      Index           =   18
      Left            =   960
      TabIndex        =   69
      Top             =   4560
      Width           =   60
   End
   Begin VB.Label Label_BinCommand 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Height          =   240
      Index           =   17
      Left            =   960
      TabIndex        =   68
      Top             =   4320
      Width           =   60
   End
   Begin VB.Label Label_BinCommand 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Height          =   240
      Index           =   16
      Left            =   960
      TabIndex        =   67
      Top             =   4080
      Width           =   60
   End
   Begin VB.Label Label_BinCommand 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Height          =   240
      Index           =   15
      Left            =   960
      TabIndex        =   66
      Top             =   3840
      Width           =   60
   End
   Begin VB.Label Label_BinCommand 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Height          =   240
      Index           =   14
      Left            =   960
      TabIndex        =   65
      Top             =   3600
      Width           =   60
   End
   Begin VB.Label Label_BinCommand 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Height          =   240
      Index           =   13
      Left            =   960
      TabIndex        =   64
      Top             =   3360
      Width           =   60
   End
   Begin VB.Label Label_BinCommand 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Height          =   240
      Index           =   12
      Left            =   960
      TabIndex        =   63
      Top             =   3120
      Width           =   60
   End
   Begin VB.Label Label_BinCommand 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Height          =   240
      Index           =   11
      Left            =   960
      TabIndex        =   62
      Top             =   2880
      Width           =   60
   End
   Begin VB.Label Label_BinCommand 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Height          =   240
      Index           =   10
      Left            =   960
      TabIndex        =   61
      Top             =   2640
      Width           =   60
   End
   Begin VB.Label Label_BinCommand 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Height          =   240
      Index           =   9
      Left            =   960
      TabIndex        =   60
      Top             =   2400
      Width           =   60
   End
   Begin VB.Label Label_BinCommand 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Height          =   240
      Index           =   8
      Left            =   960
      TabIndex        =   59
      Top             =   2160
      Width           =   60
   End
   Begin VB.Label Label_BinCommand 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Height          =   240
      Index           =   7
      Left            =   960
      TabIndex        =   58
      Top             =   1920
      Width           =   60
   End
   Begin VB.Label Label_BinCommand 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Height          =   240
      Index           =   6
      Left            =   960
      TabIndex        =   57
      Top             =   1680
      Width           =   60
   End
   Begin VB.Label Label_BinCommand 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Height          =   240
      Index           =   5
      Left            =   960
      TabIndex        =   56
      Top             =   1440
      Width           =   60
   End
   Begin VB.Label Label_BinCommand 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Height          =   240
      Index           =   4
      Left            =   960
      TabIndex        =   55
      Top             =   1200
      Width           =   60
   End
   Begin VB.Label Label_BinCommand 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Height          =   240
      Index           =   3
      Left            =   960
      TabIndex        =   54
      Top             =   960
      Width           =   60
   End
   Begin VB.Label Label_BinCommand 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Height          =   240
      Index           =   2
      Left            =   960
      TabIndex        =   53
      Top             =   720
      Width           =   60
   End
   Begin VB.Label Label_BinCommand 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Height          =   240
      Index           =   1
      Left            =   960
      TabIndex        =   52
      Top             =   480
      Width           =   60
   End
   Begin VB.Label Label_BinCommand 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Height          =   240
      Index           =   0
      Left            =   960
      TabIndex        =   51
      Top             =   120
      Width           =   540
   End
   Begin VB.Label Label_BinNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Plate Zone"
      Height          =   240
      Index           =   31
      Left            =   4920
      TabIndex        =   50
      Top             =   1560
      Width           =   930
   End
   Begin VB.Label Label_BinNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Skip Die"
      Height          =   240
      Index           =   30
      Left            =   4920
      TabIndex        =   31
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label_BinNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Measure"
      Height          =   240
      Index           =   29
      Left            =   4920
      TabIndex        =   30
      Top             =   840
      Width           =   750
   End
   Begin VB.Label Label_BinNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Pattern"
      Height          =   240
      Index           =   28
      Left            =   4920
      TabIndex        =   29
      Top             =   480
      Width           =   630
   End
   Begin VB.Label Label_BinNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Normal"
      Height          =   240
      Index           =   27
      Left            =   4920
      TabIndex        =   28
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label_BinNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "BIN #0"
      Height          =   240
      Index           =   26
      Left            =   180
      TabIndex        =   27
      Top             =   6487
      Width           =   585
   End
   Begin VB.Label Label_BinNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "BIN #0"
      Height          =   240
      Index           =   25
      Left            =   180
      TabIndex        =   26
      Top             =   6247
      Width           =   585
   End
   Begin VB.Label Label_BinNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "BIN #0"
      Height          =   240
      Index           =   24
      Left            =   180
      TabIndex        =   25
      Top             =   6007
      Width           =   585
   End
   Begin VB.Label Label_BinNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "BIN #0"
      Height          =   240
      Index           =   23
      Left            =   180
      TabIndex        =   24
      Top             =   5767
      Width           =   585
   End
   Begin VB.Label Label_BinNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "BIN #0"
      Height          =   240
      Index           =   22
      Left            =   180
      TabIndex        =   23
      Top             =   5527
      Width           =   585
   End
   Begin VB.Label Label_BinNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "BIN #0"
      Height          =   240
      Index           =   21
      Left            =   180
      TabIndex        =   22
      Top             =   5287
      Width           =   585
   End
   Begin VB.Label Label_BinNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "BIN #0"
      Height          =   240
      Index           =   20
      Left            =   180
      TabIndex        =   21
      Top             =   5047
      Width           =   585
   End
   Begin VB.Label Label_BinNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "BIN #0"
      Height          =   240
      Index           =   19
      Left            =   180
      TabIndex        =   20
      Top             =   4807
      Width           =   585
   End
   Begin VB.Label Label_BinNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "BIN #0"
      Height          =   240
      Index           =   18
      Left            =   180
      TabIndex        =   19
      Top             =   4567
      Width           =   585
   End
   Begin VB.Label Label_BinNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "BIN #0"
      Height          =   240
      Index           =   17
      Left            =   180
      TabIndex        =   18
      Top             =   4327
      Width           =   585
   End
   Begin VB.Label Label_BinNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "BIN #0"
      Height          =   240
      Index           =   16
      Left            =   180
      TabIndex        =   17
      Top             =   4087
      Width           =   585
   End
   Begin VB.Label Label_BinNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "BIN #0"
      Height          =   240
      Index           =   15
      Left            =   180
      TabIndex        =   16
      Top             =   3847
      Width           =   585
   End
   Begin VB.Label Label_BinNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "BIN #0"
      Height          =   240
      Index           =   14
      Left            =   180
      TabIndex        =   15
      Top             =   3607
      Width           =   585
   End
   Begin VB.Label Label_BinNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "BIN #0"
      Height          =   240
      Index           =   13
      Left            =   180
      TabIndex        =   14
      Top             =   3367
      Width           =   585
   End
   Begin VB.Label Label_BinNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "BIN #0"
      Height          =   240
      Index           =   12
      Left            =   180
      TabIndex        =   13
      Top             =   3127
      Width           =   585
   End
   Begin VB.Label Label_BinNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "BIN #0"
      Height          =   240
      Index           =   11
      Left            =   180
      TabIndex        =   12
      Top             =   2887
      Width           =   585
   End
   Begin VB.Label Label_BinNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "BIN #0"
      Height          =   240
      Index           =   10
      Left            =   180
      TabIndex        =   11
      Top             =   2647
      Width           =   585
   End
   Begin VB.Label Label_BinNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "BIN #0"
      Height          =   240
      Index           =   9
      Left            =   180
      TabIndex        =   10
      Top             =   2407
      Width           =   585
   End
   Begin VB.Label Label_BinNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "BIN #0"
      Height          =   240
      Index           =   8
      Left            =   180
      TabIndex        =   9
      Top             =   2167
      Width           =   585
   End
   Begin VB.Label Label_BinNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "BIN #0"
      Height          =   240
      Index           =   7
      Left            =   180
      TabIndex        =   8
      Top             =   1927
      Width           =   585
   End
   Begin VB.Label Label_BinNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "BIN #0"
      Height          =   240
      Index           =   6
      Left            =   180
      TabIndex        =   7
      Top             =   1687
      Width           =   585
   End
   Begin VB.Label Label_BinNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "BIN #0"
      Height          =   240
      Index           =   5
      Left            =   180
      TabIndex        =   6
      Top             =   1447
      Width           =   585
   End
   Begin VB.Label Label_BinNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "BIN #0"
      Height          =   240
      Index           =   4
      Left            =   180
      TabIndex        =   5
      Top             =   1207
      Width           =   585
   End
   Begin VB.Label Label_BinNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "BIN #0"
      Height          =   240
      Index           =   3
      Left            =   180
      TabIndex        =   4
      Top             =   967
      Width           =   585
   End
   Begin VB.Label Label_BinNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "BIN #0"
      Height          =   240
      Index           =   2
      Left            =   180
      TabIndex        =   3
      Top             =   727
      Width           =   585
   End
   Begin VB.Label Label_BinNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "BIN #0"
      Height          =   240
      Index           =   1
      Left            =   180
      TabIndex        =   2
      Top             =   487
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "(Bad)"
      Height          =   240
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   465
   End
   Begin VB.Label Label_BinNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "BIN #0"
      Height          =   240
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   585
   End
End
Attribute VB_Name = "Form_StarProbe_BINColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command_BINColor_Click(Index As Integer)

'    Dim lColor As Long
'
'    lColor = Command_BINColor(Index).BackColor
'
'    CommonDialog1.Flags = cdlCCRGBInit
'    CommonDialog1.Color = lColor
'    CommonDialog1.ShowColor
'
'    lColor = CommonDialog1.Color
'
'    Command_BINColor(Index).BackColor = lColor

End Sub

Private Sub Command_ChipColor_Click(Index As Integer)

'    Dim lColor As Long
'
'    lColor = Command_ChipColor(Index).BackColor
'
'    CommonDialog1.Flags = cdlCCRGBInit
'    CommonDialog1.Color = lColor
'    CommonDialog1.ShowColor
'
'    lColor = CommonDialog1.Color
'
'    Command_ChipColor(Index).BackColor = lColor

End Sub

Private Sub Form_Load()

'    Dim i, j As Integer
'    Dim buf As String
'
'    For i = 0 To 26
'
'        Label_BinNo(i) = "BIN #" & i
'
'        Command_BINColor(i).BackColor = BINColor(i)
'
'        For j = 0 To 24                                      '추가
'
''                 If i = PROD.Bin_Data(j).Bin_no Then
''                      Label_BinCommand(i) = UCase(PROD.Bin_Data(j).Bin_Comment)
''                      Label_BinCommand(i).ForeColor = vbBlue
''                 End If
'
'        Next j
'    Next
'
'
'    For i = 0 To 24
'        buf = ""
'
'        If Test_Cnt > 0 Then
'
'        End If
'        Label_BinCount(i).Caption = buf
'    Next
'
'
'
'    For i = 0 To 5
'        Command_ChipColor(i).BackColor = ChipColor(i)
'    Next
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

'    Dim i As Integer
'
'    For i = 0 To 26
'
'        BINColor(i) = Command_BINColor(i).BackColor
'
'    Next
'
'    For i = 0 To 5
'        ChipColor(i) = Command_ChipColor(i).BackColor
'    Next
'
'    ' 2005.08.11
'    Call StarProbe_FileSave_SystemInfo

End Sub

