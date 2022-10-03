VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form Form_Cassette_New 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "Cassette"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13770
   Icon            =   "Cassette_New.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   13770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.CommandButton Command3 
      Caption         =   "FILE LOAD"
      Height          =   615
      Left            =   8400
      TabIndex        =   127
      Top             =   4200
      Width           =   5175
   End
   Begin VB.TextBox Text10 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   6
      Left            =   10200
      TabIndex        =   121
      Top             =   1680
      Width           =   3375
   End
   Begin VB.TextBox Text10 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   10200
      TabIndex        =   120
      Top             =   960
      Width           =   3375
   End
   Begin VB.TextBox Text10 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   10200
      TabIndex        =   119
      Top             =   240
      Width           =   3375
   End
   Begin VB.TextBox Text10 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   10200
      TabIndex        =   118
      Top             =   1320
      Width           =   3375
   End
   Begin VB.TextBox Text10 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   5
      Left            =   10200
      TabIndex        =   117
      Top             =   600
      Width           =   3375
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   9720
      Top             =   5160
   End
   Begin VB.TextBox txt_Barcode 
      Height          =   375
      Left            =   8400
      TabIndex        =   0
      Top             =   3480
      Width           =   5175
   End
   Begin VB.Frame Frame26 
      Height          =   5415
      Left            =   6960
      TabIndex        =   115
      Top             =   6360
      Width           =   30
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Apply Default Value"
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
      Left            =   3360
      Style           =   1  '그래픽
      TabIndex        =   114
      Top             =   6600
      Width           =   2295
   End
   Begin VB.TextBox text1 
      Alignment       =   2  '가운데 맞춤
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1440
      TabIndex        =   112
      Text            =   "3"
      Top             =   6720
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  '가운데 맞춤
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2400
      TabIndex        =   111
      Text            =   "1"
      Top             =   6720
      Width           =   855
   End
   Begin VB.Frame Frame18 
      Height          =   30
      Left            =   240
      TabIndex        =   104
      Top             =   2400
      Width           =   7935
   End
   Begin VB.Frame Frame6 
      Height          =   5415
      Left            =   5520
      TabIndex        =   83
      Top             =   360
      Width           =   30
   End
   Begin VB.ComboBox Combo_Y 
      Height          =   300
      Index           =   3
      ItemData        =   "Cassette_New.frx":08CA
      Left            =   1920
      List            =   "Cassette_New.frx":08FE
      Style           =   2  '드롭다운 목록
      TabIndex        =   45
      Top             =   2520
      Width           =   855
   End
   Begin VB.ComboBox Combo_X 
      Height          =   300
      Index           =   3
      ItemData        =   "Cassette_New.frx":0942
      Left            =   1080
      List            =   "Cassette_New.frx":0976
      Style           =   2  '드롭다운 목록
      TabIndex        =   44
      Top             =   2520
      Width           =   855
   End
   Begin VB.ComboBox Combo_Y 
      Height          =   300
      Index           =   2
      ItemData        =   "Cassette_New.frx":09BA
      Left            =   1920
      List            =   "Cassette_New.frx":09EE
      Style           =   2  '드롭다운 목록
      TabIndex        =   43
      Top             =   2040
      Width           =   855
   End
   Begin VB.ComboBox Combo_X 
      Height          =   300
      Index           =   2
      ItemData        =   "Cassette_New.frx":0A32
      Left            =   1080
      List            =   "Cassette_New.frx":0A66
      Style           =   2  '드롭다운 목록
      TabIndex        =   42
      Top             =   2040
      Width           =   855
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
      Left            =   11520
      Style           =   1  '그래픽
      TabIndex        =   31
      Top             =   6600
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Caption         =   "Slot (NEW LOT작업시 ALL CHECK를 한번 실행 후 체크상태를 초기화 해주세요)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   8175
      Begin VB.Frame Frame24 
         Height          =   30
         Left            =   120
         TabIndex        =   110
         Top             =   5160
         Width           =   7935
      End
      Begin VB.Frame Frame23 
         Height          =   30
         Left            =   120
         TabIndex        =   109
         Top             =   4680
         Width           =   7935
      End
      Begin VB.Frame Frame22 
         Height          =   30
         Left            =   120
         TabIndex        =   108
         Top             =   4200
         Width           =   7935
      End
      Begin VB.Frame Frame21 
         Height          =   30
         Left            =   120
         TabIndex        =   107
         Top             =   3720
         Width           =   7935
      End
      Begin VB.Frame Frame20 
         Height          =   30
         Left            =   120
         TabIndex        =   106
         Top             =   3240
         Width           =   7935
      End
      Begin VB.Frame Frame19 
         Height          =   30
         Left            =   120
         TabIndex        =   105
         Top             =   2760
         Width           =   7935
      End
      Begin VB.Frame Frame17 
         Height          =   30
         Left            =   120
         TabIndex        =   103
         Top             =   1800
         Width           =   7935
      End
      Begin VB.Frame Frame16 
         Height          =   30
         Left            =   120
         TabIndex        =   102
         Top             =   1320
         Width           =   7935
      End
      Begin VB.Frame Frame15 
         Height          =   30
         Left            =   120
         TabIndex        =   101
         Top             =   5640
         Width           =   7935
      End
      Begin VB.Frame Frame14 
         Height          =   30
         Left            =   120
         TabIndex        =   100
         Top             =   360
         Width           =   7935
      End
      Begin VB.Frame Frame12 
         Height          =   5415
         Left            =   120
         TabIndex        =   99
         Top             =   240
         Width           =   30
      End
      Begin VB.Frame Frame8 
         Height          =   5415
         Left            =   8040
         TabIndex        =   98
         Top             =   240
         Width           =   30
      End
      Begin VB.ComboBox Combo_Y 
         Height          =   300
         Index           =   24
         ItemData        =   "Cassette_New.frx":0AAA
         Left            =   7080
         List            =   "Cassette_New.frx":0ADE
         Style           =   2  '드롭다운 목록
         TabIndex        =   97
         Top             =   2880
         Width           =   855
      End
      Begin VB.ComboBox Combo_Y 
         Height          =   300
         Index           =   23
         ItemData        =   "Cassette_New.frx":0B22
         Left            =   7080
         List            =   "Cassette_New.frx":0B56
         Style           =   2  '드롭다운 목록
         TabIndex        =   96
         Top             =   2400
         Width           =   855
      End
      Begin VB.ComboBox Combo_Y 
         Height          =   300
         Index           =   22
         ItemData        =   "Cassette_New.frx":0B9A
         Left            =   7080
         List            =   "Cassette_New.frx":0BCE
         Style           =   2  '드롭다운 목록
         TabIndex        =   95
         Top             =   1920
         Width           =   855
      End
      Begin VB.ComboBox Combo_Y 
         Height          =   300
         Index           =   21
         ItemData        =   "Cassette_New.frx":0C12
         Left            =   7080
         List            =   "Cassette_New.frx":0C46
         Style           =   2  '드롭다운 목록
         TabIndex        =   94
         Top             =   1440
         Width           =   855
      End
      Begin VB.ComboBox Combo_Y 
         Height          =   300
         Index           =   20
         ItemData        =   "Cassette_New.frx":0C8A
         Left            =   7080
         List            =   "Cassette_New.frx":0CBE
         Style           =   2  '드롭다운 목록
         TabIndex        =   93
         Top             =   960
         Width           =   855
      End
      Begin VB.ComboBox Combo_X 
         Height          =   300
         Index           =   24
         ItemData        =   "Cassette_New.frx":0D02
         Left            =   6240
         List            =   "Cassette_New.frx":0D36
         Style           =   2  '드롭다운 목록
         TabIndex        =   92
         Top             =   2880
         Width           =   855
      End
      Begin VB.ComboBox Combo_X 
         Height          =   300
         Index           =   23
         ItemData        =   "Cassette_New.frx":0D7A
         Left            =   6240
         List            =   "Cassette_New.frx":0DAE
         Style           =   2  '드롭다운 목록
         TabIndex        =   91
         Top             =   2400
         Width           =   855
      End
      Begin VB.ComboBox Combo_X 
         Height          =   300
         Index           =   22
         ItemData        =   "Cassette_New.frx":0DF2
         Left            =   6240
         List            =   "Cassette_New.frx":0E26
         Style           =   2  '드롭다운 목록
         TabIndex        =   90
         Top             =   1920
         Width           =   855
      End
      Begin VB.ComboBox Combo_X 
         Height          =   300
         Index           =   21
         ItemData        =   "Cassette_New.frx":0E6A
         Left            =   6240
         List            =   "Cassette_New.frx":0E9E
         Style           =   2  '드롭다운 목록
         TabIndex        =   89
         Top             =   1440
         Width           =   855
      End
      Begin VB.ComboBox Combo_X 
         Height          =   300
         Index           =   20
         ItemData        =   "Cassette_New.frx":0EE2
         Left            =   6240
         List            =   "Cassette_New.frx":0F16
         Style           =   2  '드롭다운 목록
         TabIndex        =   88
         Top             =   960
         Width           =   855
      End
      Begin VB.Frame Frame7 
         Height          =   5415
         Left            =   6120
         TabIndex        =   85
         Top             =   240
         Width           =   30
      End
      Begin VB.ComboBox Combo_Y 
         Height          =   300
         Index           =   19
         ItemData        =   "Cassette_New.frx":0F5A
         Left            =   4440
         List            =   "Cassette_New.frx":0F8E
         Style           =   2  '드롭다운 목록
         TabIndex        =   82
         Top             =   5280
         Width           =   855
      End
      Begin VB.ComboBox Combo_X 
         Height          =   300
         Index           =   19
         ItemData        =   "Cassette_New.frx":0FD2
         Left            =   3600
         List            =   "Cassette_New.frx":1006
         Style           =   2  '드롭다운 목록
         TabIndex        =   81
         Top             =   5280
         Width           =   855
      End
      Begin VB.ComboBox Combo_Y 
         Height          =   300
         Index           =   18
         ItemData        =   "Cassette_New.frx":104A
         Left            =   4440
         List            =   "Cassette_New.frx":107E
         Style           =   2  '드롭다운 목록
         TabIndex        =   80
         Top             =   4800
         Width           =   855
      End
      Begin VB.ComboBox Combo_X 
         Height          =   300
         Index           =   18
         ItemData        =   "Cassette_New.frx":10C2
         Left            =   3600
         List            =   "Cassette_New.frx":10F6
         Style           =   2  '드롭다운 목록
         TabIndex        =   79
         Top             =   4800
         Width           =   855
      End
      Begin VB.ComboBox Combo_Y 
         Height          =   300
         Index           =   17
         ItemData        =   "Cassette_New.frx":113A
         Left            =   4440
         List            =   "Cassette_New.frx":116E
         Style           =   2  '드롭다운 목록
         TabIndex        =   78
         Top             =   4320
         Width           =   855
      End
      Begin VB.ComboBox Combo_X 
         Height          =   300
         Index           =   17
         ItemData        =   "Cassette_New.frx":11B2
         Left            =   3600
         List            =   "Cassette_New.frx":11E6
         Style           =   2  '드롭다운 목록
         TabIndex        =   77
         Top             =   4320
         Width           =   855
      End
      Begin VB.ComboBox Combo_Y 
         Height          =   300
         Index           =   16
         ItemData        =   "Cassette_New.frx":122A
         Left            =   4440
         List            =   "Cassette_New.frx":125E
         Style           =   2  '드롭다운 목록
         TabIndex        =   76
         Top             =   3840
         Width           =   855
      End
      Begin VB.ComboBox Combo_X 
         Height          =   300
         Index           =   16
         ItemData        =   "Cassette_New.frx":12A2
         Left            =   3600
         List            =   "Cassette_New.frx":12D6
         Style           =   2  '드롭다운 목록
         TabIndex        =   75
         Top             =   3840
         Width           =   855
      End
      Begin VB.ComboBox Combo_Y 
         Height          =   300
         Index           =   15
         ItemData        =   "Cassette_New.frx":131A
         Left            =   4440
         List            =   "Cassette_New.frx":134E
         Style           =   2  '드롭다운 목록
         TabIndex        =   74
         Top             =   3360
         Width           =   855
      End
      Begin VB.ComboBox Combo_X 
         Height          =   300
         Index           =   15
         ItemData        =   "Cassette_New.frx":1392
         Left            =   3600
         List            =   "Cassette_New.frx":13C6
         Style           =   2  '드롭다운 목록
         TabIndex        =   73
         Top             =   3360
         Width           =   855
      End
      Begin VB.ComboBox Combo_Y 
         Height          =   300
         Index           =   14
         ItemData        =   "Cassette_New.frx":140A
         Left            =   4440
         List            =   "Cassette_New.frx":143E
         Style           =   2  '드롭다운 목록
         TabIndex        =   72
         Top             =   2880
         Width           =   855
      End
      Begin VB.ComboBox Combo_X 
         Height          =   300
         Index           =   14
         ItemData        =   "Cassette_New.frx":1482
         Left            =   3600
         List            =   "Cassette_New.frx":14B6
         Style           =   2  '드롭다운 목록
         TabIndex        =   71
         Top             =   2880
         Width           =   855
      End
      Begin VB.ComboBox Combo_Y 
         Height          =   300
         Index           =   13
         ItemData        =   "Cassette_New.frx":14FA
         Left            =   4440
         List            =   "Cassette_New.frx":152E
         Style           =   2  '드롭다운 목록
         TabIndex        =   70
         Top             =   2400
         Width           =   855
      End
      Begin VB.ComboBox Combo_X 
         Height          =   300
         Index           =   13
         ItemData        =   "Cassette_New.frx":1572
         Left            =   3600
         List            =   "Cassette_New.frx":15A6
         Style           =   2  '드롭다운 목록
         TabIndex        =   69
         Top             =   2400
         Width           =   855
      End
      Begin VB.ComboBox Combo_Y 
         Height          =   300
         Index           =   12
         ItemData        =   "Cassette_New.frx":15EA
         Left            =   4440
         List            =   "Cassette_New.frx":161E
         Style           =   2  '드롭다운 목록
         TabIndex        =   68
         Top             =   1920
         Width           =   855
      End
      Begin VB.ComboBox Combo_X 
         Height          =   300
         Index           =   12
         ItemData        =   "Cassette_New.frx":1662
         Left            =   3600
         List            =   "Cassette_New.frx":1696
         Style           =   2  '드롭다운 목록
         TabIndex        =   67
         Top             =   1920
         Width           =   855
      End
      Begin VB.ComboBox Combo_Y 
         Height          =   300
         Index           =   11
         ItemData        =   "Cassette_New.frx":16DA
         Left            =   4440
         List            =   "Cassette_New.frx":170E
         Style           =   2  '드롭다운 목록
         TabIndex        =   66
         Top             =   1440
         Width           =   855
      End
      Begin VB.ComboBox Combo_X 
         Height          =   300
         Index           =   11
         ItemData        =   "Cassette_New.frx":1752
         Left            =   3600
         List            =   "Cassette_New.frx":1786
         Style           =   2  '드롭다운 목록
         TabIndex        =   65
         Top             =   1440
         Width           =   855
      End
      Begin VB.ComboBox Combo_Y 
         Height          =   300
         Index           =   10
         ItemData        =   "Cassette_New.frx":17CA
         Left            =   4440
         List            =   "Cassette_New.frx":17FE
         Style           =   2  '드롭다운 목록
         TabIndex        =   64
         Top             =   960
         Width           =   855
      End
      Begin VB.ComboBox Combo_X 
         Height          =   300
         Index           =   10
         ItemData        =   "Cassette_New.frx":1842
         Left            =   3600
         List            =   "Cassette_New.frx":1876
         Style           =   2  '드롭다운 목록
         TabIndex        =   63
         Top             =   960
         Width           =   855
      End
      Begin VB.Frame Frame5 
         Height          =   5415
         Left            =   3480
         TabIndex        =   62
         Top             =   240
         Width           =   30
      End
      Begin VB.Frame Frame4 
         Height          =   5415
         Left            =   2760
         TabIndex        =   58
         Top             =   240
         Width           =   30
      End
      Begin VB.ComboBox Combo_Y 
         Height          =   300
         Index           =   9
         ItemData        =   "Cassette_New.frx":18BA
         Left            =   1800
         List            =   "Cassette_New.frx":18EE
         Style           =   2  '드롭다운 목록
         TabIndex        =   57
         Top             =   5280
         Width           =   855
      End
      Begin VB.ComboBox Combo_X 
         Height          =   300
         Index           =   9
         ItemData        =   "Cassette_New.frx":1932
         Left            =   960
         List            =   "Cassette_New.frx":1966
         Style           =   2  '드롭다운 목록
         TabIndex        =   56
         Top             =   5280
         Width           =   855
      End
      Begin VB.ComboBox Combo_Y 
         Height          =   300
         Index           =   8
         ItemData        =   "Cassette_New.frx":19AA
         Left            =   1800
         List            =   "Cassette_New.frx":19DE
         Style           =   2  '드롭다운 목록
         TabIndex        =   55
         Top             =   4800
         Width           =   855
      End
      Begin VB.ComboBox Combo_X 
         Height          =   300
         Index           =   8
         ItemData        =   "Cassette_New.frx":1A22
         Left            =   960
         List            =   "Cassette_New.frx":1A56
         Style           =   2  '드롭다운 목록
         TabIndex        =   54
         Top             =   4800
         Width           =   855
      End
      Begin VB.ComboBox Combo_Y 
         Height          =   300
         Index           =   7
         ItemData        =   "Cassette_New.frx":1A9A
         Left            =   1800
         List            =   "Cassette_New.frx":1ACE
         Style           =   2  '드롭다운 목록
         TabIndex        =   53
         Top             =   4320
         Width           =   855
      End
      Begin VB.ComboBox Combo_X 
         Height          =   300
         Index           =   7
         ItemData        =   "Cassette_New.frx":1B12
         Left            =   960
         List            =   "Cassette_New.frx":1B46
         Style           =   2  '드롭다운 목록
         TabIndex        =   52
         Top             =   4320
         Width           =   855
      End
      Begin VB.ComboBox Combo_Y 
         Height          =   300
         Index           =   6
         ItemData        =   "Cassette_New.frx":1B8A
         Left            =   1800
         List            =   "Cassette_New.frx":1BBE
         Style           =   2  '드롭다운 목록
         TabIndex        =   51
         Top             =   3840
         Width           =   855
      End
      Begin VB.ComboBox Combo_X 
         Height          =   300
         Index           =   6
         ItemData        =   "Cassette_New.frx":1C02
         Left            =   960
         List            =   "Cassette_New.frx":1C36
         Style           =   2  '드롭다운 목록
         TabIndex        =   50
         Top             =   3840
         Width           =   855
      End
      Begin VB.ComboBox Combo_Y 
         Height          =   300
         Index           =   5
         ItemData        =   "Cassette_New.frx":1C7A
         Left            =   1800
         List            =   "Cassette_New.frx":1CAE
         Style           =   2  '드롭다운 목록
         TabIndex        =   49
         Top             =   3360
         Width           =   855
      End
      Begin VB.ComboBox Combo_X 
         Height          =   300
         Index           =   5
         ItemData        =   "Cassette_New.frx":1CF2
         Left            =   960
         List            =   "Cassette_New.frx":1D26
         Style           =   2  '드롭다운 목록
         TabIndex        =   48
         Top             =   3360
         Width           =   855
      End
      Begin VB.ComboBox Combo_Y 
         Height          =   300
         Index           =   4
         ItemData        =   "Cassette_New.frx":1D6A
         Left            =   1800
         List            =   "Cassette_New.frx":1D9E
         Style           =   2  '드롭다운 목록
         TabIndex        =   47
         Top             =   2880
         Width           =   855
      End
      Begin VB.ComboBox Combo_X 
         Height          =   300
         Index           =   4
         ItemData        =   "Cassette_New.frx":1DE2
         Left            =   960
         List            =   "Cassette_New.frx":1E16
         Style           =   2  '드롭다운 목록
         TabIndex        =   46
         Top             =   2880
         Width           =   855
      End
      Begin VB.ComboBox Combo_Y 
         Height          =   300
         Index           =   1
         ItemData        =   "Cassette_New.frx":1E5A
         Left            =   1800
         List            =   "Cassette_New.frx":1E8E
         Style           =   2  '드롭다운 목록
         TabIndex        =   41
         Top             =   1440
         Width           =   855
      End
      Begin VB.ComboBox Combo_X 
         Height          =   300
         Index           =   1
         ItemData        =   "Cassette_New.frx":1ED2
         Left            =   960
         List            =   "Cassette_New.frx":1F06
         Style           =   2  '드롭다운 목록
         TabIndex        =   40
         Top             =   1440
         Width           =   855
      End
      Begin VB.Frame Frame3 
         Height          =   5415
         Left            =   840
         TabIndex        =   39
         Top             =   240
         Width           =   30
      End
      Begin VB.Frame Frame1 
         Height          =   30
         Left            =   120
         TabIndex        =   38
         Top             =   840
         Width           =   7935
      End
      Begin VB.ComboBox Combo_Y 
         Height          =   300
         Index           =   0
         ItemData        =   "Cassette_New.frx":1F4A
         Left            =   1800
         List            =   "Cassette_New.frx":1F7E
         Style           =   2  '드롭다운 목록
         TabIndex        =   34
         Top             =   960
         Width           =   855
      End
      Begin VB.ComboBox Combo_X 
         Height          =   300
         Index           =   0
         ItemData        =   "Cassette_New.frx":1FC2
         Left            =   960
         List            =   "Cassette_New.frx":1FF6
         Style           =   2  '드롭다운 목록
         TabIndex        =   33
         Top             =   960
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 1"
         Height          =   300
         Index           =   0
         Left            =   240
         TabIndex        =   32
         Top             =   960
         Width           =   615
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Line check"
         Height          =   465
         Left            =   5400
         Style           =   1  '그래픽
         TabIndex        =   30
         Top             =   5760
         Width           =   2610
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 25"
         Height          =   300
         Index           =   24
         Left            =   5520
         TabIndex        =   29
         Top             =   2880
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 21"
         Height          =   300
         Index           =   20
         Left            =   5520
         TabIndex        =   28
         Top             =   960
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 22"
         Height          =   300
         Index           =   21
         Left            =   5520
         TabIndex        =   27
         Top             =   1440
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 23"
         Height          =   300
         Index           =   22
         Left            =   5520
         TabIndex        =   26
         Top             =   1920
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 24"
         Height          =   300
         Index           =   23
         Left            =   5520
         TabIndex        =   25
         Top             =   2400
         Width           =   615
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Line check"
         Height          =   465
         Left            =   2760
         Style           =   1  '그래픽
         TabIndex        =   24
         Top             =   5760
         Width           =   2610
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 13"
         Height          =   300
         Index           =   12
         Left            =   2880
         TabIndex        =   23
         Top             =   1920
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 14"
         Height          =   300
         Index           =   13
         Left            =   2880
         TabIndex        =   22
         Top             =   2400
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 15"
         Height          =   300
         Index           =   14
         Left            =   2880
         TabIndex        =   21
         Top             =   2880
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 16"
         Height          =   300
         Index           =   15
         Left            =   2880
         TabIndex        =   20
         Top             =   3360
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 17"
         Height          =   300
         Index           =   16
         Left            =   2880
         TabIndex        =   19
         Top             =   3840
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 18"
         Height          =   300
         Index           =   17
         Left            =   2880
         TabIndex        =   18
         Top             =   4320
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 19"
         Height          =   300
         Index           =   18
         Left            =   2880
         TabIndex        =   17
         Top             =   4800
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 20"
         Height          =   300
         Index           =   19
         Left            =   2880
         TabIndex        =   16
         Top             =   5280
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 11"
         Height          =   300
         Index           =   10
         Left            =   2880
         TabIndex        =   15
         Top             =   960
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 12"
         Height          =   255
         Index           =   11
         Left            =   2880
         TabIndex        =   14
         Top             =   1440
         Width           =   615
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Line Check"
         Height          =   465
         Left            =   120
         Style           =   1  '그래픽
         TabIndex        =   13
         Top             =   5760
         Width           =   2610
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 2"
         Height          =   300
         Index           =   1
         Left            =   240
         TabIndex        =   12
         Top             =   1440
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 3"
         Height          =   300
         Index           =   2
         Left            =   240
         TabIndex        =   11
         Top             =   1920
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 4"
         Height          =   300
         Index           =   3
         Left            =   240
         TabIndex        =   10
         Top             =   2400
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 5"
         Height          =   300
         Index           =   4
         Left            =   240
         TabIndex        =   9
         Top             =   2880
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 6"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   8
         Top             =   3360
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 7"
         Height          =   300
         Index           =   6
         Left            =   240
         TabIndex        =   7
         Top             =   3840
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 8"
         Height          =   300
         Index           =   7
         Left            =   240
         TabIndex        =   6
         Top             =   4320
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 9"
         Height          =   300
         Index           =   8
         Left            =   240
         TabIndex        =   5
         Top             =   4800
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 10"
         Height          =   300
         Index           =   9
         Left            =   240
         TabIndex        =   4
         Top             =   5280
         Width           =   615
      End
      Begin VB.Label Label12 
         Caption         =   "Y step"
         Height          =   255
         Left            =   7200
         TabIndex        =   87
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "X step"
         Height          =   255
         Left            =   6360
         TabIndex        =   86
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "No"
         Height          =   255
         Left            =   5640
         TabIndex        =   84
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label8 
         Caption         =   "Y step"
         Height          =   255
         Left            =   4560
         TabIndex        =   61
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "X step"
         Height          =   255
         Left            =   3720
         TabIndex        =   60
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "No"
         Height          =   255
         Left            =   3000
         TabIndex        =   59
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label4 
         Caption         =   "Y step"
         Height          =   255
         Left            =   1920
         TabIndex        =   37
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "X step"
         Height          =   255
         Left            =   1080
         TabIndex        =   36
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "No"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   35
         Top             =   480
         Width           =   255
      End
   End
   Begin VB.CheckBox Check7 
      Caption         =   "AUTO ALIGN OFF"
      Height          =   375
      Left            =   7920
      TabIndex        =   2
      Top             =   6720
      Value           =   1  '확인
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CheckBox Check8 
      BackColor       =   &H00FFC0C0&
      Caption         =   "ALL CHECK"
      Height          =   615
      Left            =   10080
      Style           =   1  '그래픽
      TabIndex        =   1
      Top             =   6600
      Width           =   1290
   End
   Begin Threed.SSPanel SSPanel4 
      Height          =   375
      Index           =   1
      Left            =   8400
      TabIndex        =   116
      Top             =   3120
      Width           =   5175
      _Version        =   65536
      _ExtentX        =   9128
      _ExtentY        =   661
      _StockProps     =   15
      Caption         =   "BARCODE"
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.74
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelInner      =   1
      Font3D          =   3
   End
   Begin Threed.SSPanel SSPanel4 
      Height          =   375
      Index           =   0
      Left            =   8400
      TabIndex        =   122
      Top             =   1680
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   661
      _StockProps     =   15
      Caption         =   "PROBE CARD NO"
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.74
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelInner      =   1
      Font3D          =   3
   End
   Begin Threed.SSPanel SSPanel4 
      Height          =   375
      Index           =   13
      Left            =   8400
      TabIndex        =   123
      Top             =   1320
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   661
      _StockProps     =   15
      Caption         =   "EQUIPMENT ID"
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.74
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelInner      =   1
      Font3D          =   3
   End
   Begin Threed.SSPanel SSPanel9 
      Height          =   375
      Index           =   8
      Left            =   8400
      TabIndex        =   124
      Top             =   960
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   661
      _StockProps     =   15
      Caption         =   "LOT NO"
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.74
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelInner      =   1
      Font3D          =   3
   End
   Begin Threed.SSPanel SSPanel4 
      Height          =   375
      Index           =   11
      Left            =   8400
      TabIndex        =   125
      Top             =   600
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   661
      _StockProps     =   15
      Caption         =   "OPERATER NO"
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.74
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelInner      =   1
      Font3D          =   3
   End
   Begin Threed.SSPanel SSPanel18 
      Height          =   375
      Index           =   1
      Left            =   8400
      TabIndex        =   126
      Top             =   240
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   661
      _StockProps     =   15
      Caption         =   "DEVICE NAME"
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.74
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelInner      =   1
      Font3D          =   3
   End
   Begin VB.Label Label2 
      Caption         =   "- AOI(O) : Cassette No가 자동으로 설정 됩니다."
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8520
      TabIndex        =   129
      Top             =   5520
      Width           =   4935
   End
   Begin VB.Label Label6 
      Caption         =   "- AOI(X) : Cassette No를 수동으로 설정 해야 됩니다."
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8520
      TabIndex        =   128
      Top             =   6000
      Width           =   5175
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "X ,Y PITCH"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   3
      Left            =   240
      TabIndex        =   113
      Top             =   6840
      Width           =   870
   End
End
Attribute VB_Name = "Form_Cassette_New"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click(Index As Integer)
    If Check1(Index).value = 0 Then     'wafer use
        NOW_NO(Index) = True
    Else                                'wafer not use
        NOW_NO(Index) = False
    End If
End Sub

Private Sub Check2_Click()
    Dim i As Integer
    
    If Check2.value = 1 Then
        For i = 0 To 9
            Check1(i).value = 1         'wafer not use
        Next i
    Else
        For i = 0 To 9
            Check1(i).value = 0         'wafer use
        Next i
    End If
End Sub

Private Sub Check3_Click()
    Dim i As Integer
    
    If Check3.value = 1 Then
        For i = 10 To 19
            Check1(i).value = 1         'wafer not use
        Next i
    Else
        For i = 10 To 19
            Check1(i).value = 0         'wafer use
        Next i
    End If
End Sub

Private Sub Check4_Click()
    Dim i As Integer
    
    If Check4.value = 1 Then
        For i = 20 To 24
            Check1(i).value = 1         'wafer not use
        Next i
    Else
        For i = 20 To 24
            Check1(i).value = 0         'wafer use
        Next i
    End If
End Sub

Private Sub Check7_Click()
    If Check7.value = 0 Then
        AutoAlign_Flag = False
        Check7.Caption = "AUTO ALIGN OFF"
    Else
        AutoAlign_Flag = True
        Check7.Caption = "AUTO ALIGN ON"
    End If
End Sub

Private Sub Check8_Click()
    Dim i As Integer
    
    If Check8.value = 1 Then
        For i = 0 To 24
            Check1(i).value = 1     'wafer not use
        Next i
    Else
        For i = 0 To 24
            Check1(i).value = 0     'wafer use
        Next i
    End If
End Sub

Private Sub Combo_X_Click(Index As Integer)
    If Combo_X(Index).ListIndex = 0 Then                            '0을 선택한 경우 처리
        Combo_Y(Index).ListIndex = 0                                'Y도 0으로 변경해준다
    Else                                                            '0이 아닌 경우
        If Combo_Y(Index).ListIndex = 0 Then                        'Y가0이면 default값으로 변경해준다
            Combo_Y(Index).ListIndex = val(Text2)                   'default값 적용
        End If
    End If
End Sub

Private Sub Combo_Y_Click(Index As Integer)
    If Combo_Y(Index).ListIndex = 0 Then                            '0을 선택한 경우 처리
        Combo_X(Index).ListIndex = 0                                'X도 0으로 변경해준다
    Else                                                            '0이 아닌 경우
        If Combo_X(Index).ListIndex = 0 Then                        'X가 0인 경우
            Combo_X(Index).ListIndex = val(Text1)                   'default값 적용
        End If
    End If
End Sub

Private Sub Command1_Click()
    Dim CHK_FLAG As Integer                                         '사용할 wafer의 수를 저장하는 변수
    Dim i As Integer
    
    CHK_FLAG = 0                                                    '변수 초기화
    For i = 0 To 24                                                 'total 25매 wafer
        If NOW_NO(i) = False Then                                   'check를 하지 않은 경우 사용하는 것으로 설정
            CHK_FLAG = CHK_FLAG + 1                                 '변수를 증가
        End If
    Next i
    
    If CHK_FLAG = 25 Then                                           '아무 설정도 하지 않은 경우 처리
        MsgBox "Wafer 설정상태를 다시 확인해 주세요!!"              '경고 메시지로 다시 한번 확인
        Exit Sub
    End If
    
    If Barcode_Use = True Then
        If Text10(2).Text = "" Then
            MsgBox "DEVICE NAME을 입력해주세요!!"                   '경고 메시지로 다시 한번 확인
            Exit Sub
        End If
        If Text10(5).Text = "" Then
            MsgBox "OPERATOR NO를 입력해주세요!!"                   '경고 메시지로 다시 한번 확인
            Exit Sub
        End If
        If Text10(1).Text = "" Then
            MsgBox "LOT NO를 입력해주세요!!"                        '경고 메시지로 다시 한번 확인
            Exit Sub
        End If
        If Text10(3).Text = "" Then
            MsgBox "EQUIPMENT ID를 입력해주세요!!"                  '경고 메시지로 다시 한번 확인
            Exit Sub
        End If
        If Text10(6).Text = "" Then
            MsgBox "PROBE CARD NO를 입력해주세요!!"                 '경고 메시지로 다시 한번 확인
            Exit Sub
        End If
    Else
        If Text10(2).Text = "" Then
            MsgBox "DEVICE NAME을 입력해주세요!!"                   '경고 메시지로 다시 한번 확인
            Exit Sub
        End If
        If Text10(1).Text = "" Then
            MsgBox "LOT NO를 입력해주세요!!"                        '경고 메시지로 다시 한번 확인
            Exit Sub
        End If
    End If
        
    For i = 0 To 24
        If Check1(i).value = 1 Then                                 '사용안함
            Slot_No(i + 1) = True
        Else                                                        '사용함
            Slot_No(i + 1) = False
            Slot_Max_Count = i + 1                                  '측정할 wafer의 수를 증가한다.
        End If
        If Combo_X(i).Text <> Empty Then XPitch(i) = Combo_X(i).Text       '설정값을 변수에 적용 X step value
        If Combo_Y(i).Text <> Empty Then YPitch(i) = Combo_Y(i).Text       '설정값을 변수에 적용 Y step value
    Next i
    
    '[ 2017.08.18 ] : 기준 x,y step값을 cassette에서 설정하도록 수정.
    XPitch_MAIN = Text1.Text                                         'default X step value set
    YPitch_MAIN = Text2.Text                                         'default Y step value set
    StarProbe.MeasureStepX = XPitch_MAIN                            'X step value save값 설정
    StarProbe.MeasureStepY = YPitch_MAIN                            'Y step value save값 설정
    
    MT2000.Text1(0).Text = Text10(1).Text
    MT2000.Text1(1).Text = Text10(2).Text
    MT2000.Text1(3).Text = Text10(3).Text
    
    DEV = Text10(2).Text
    OPE = Text10(5).Text
    LOT = Text10(1).Text
    EQU = Text10(3).Text
    PRO = Text10(6).Text
    
    Log_Cnt = 0
    Log_Prn_Cnt = 0
    DataLog_flag = True
    
    Call StarProbe_FileSave_SystemInfo                              '설정내용을 starporbe.ifo 파일에 저장
    '''''
    AutoAlign_Flag = True
    Timer1.Enabled = False                                          '[ 바코드 ]
    Unload Me                                                       '폼 종료
End Sub

Private Sub Command2_Click()
    Dim i As Integer
    
    If Text1.Text <> Empty Then                                  'default X step value
        If IsNumeric(Text1.Text) = False Then                    '숫자 값인지를 체크
            MsgBox "Invalid .This blank Inputed to Number !", vbExclamation, "Error"
            Text1.Text = ""
            Exit Sub
        End If
        If val(Text1.Text) > 15 Or val(Text1.Text) < 0 Then                              '15보다 큰수를 입력한 경우
            MsgBox "Invalid value.Use only 1 ~ 15 here!", vbExclamation, "Error"
            Text1.Text = ""
            Exit Sub
        End If
    Else                                                        '공백인 경우
        MsgBox "Invalid value.Use only 1 ~ 15 here!", vbExclamation, "Error"
        Text1.Text = ""
        Exit Sub
    End If
    If Text2.Text <> Empty Then                                  'default Y step value
        If IsNumeric(Text2.Text) = False Then                    '숫자 값인지를 체크
            MsgBox "Invalid .This blank Inputed to Number !", vbExclamation, "Error"
            Text2 = ""
            Exit Sub
        End If
        If val(Text2.Text) > 15 Or val(Text2.Text) < 0 Then                              '15보다 큰수를 입력한 경우
            MsgBox "Invalid value.Use only 1 ~ 15 here!", vbExclamation, "Error"
            Text2.Text = ""
            Exit Sub
        End If
    Else
        MsgBox "Invalid value.Use only 1 ~ 15 here!", vbExclamation, "Error"
        Text2.Text = ""
        Exit Sub
    End If
    
    '[ 2017.08.18 ] : 기준 x,y step값을 cassette에서 설정하도록 수정.
    XPitch_MAIN = Text1.Text                                         'default X step value를 변수에 적용
    YPitch_MAIN = Text2.Text                                         'default Y step value를 변수에 적용
    StarProbe.MeasureStepX = XPitch_MAIN
    StarProbe.MeasureStepY = YPitch_MAIN
    
    For i = 0 To 24
        Combo_X(i).ListIndex = XPitch_MAIN
        Combo_Y(i).ListIndex = YPitch_MAIN
    Next i
End Sub

Private Sub Command3_Click()
'    Form_AOI_LIST.Show 1
    Dim CHK_FLAG As Integer                                         '사용할 wafer의 수를 저장하는 변수
    Dim i As Integer
    
    CHK_FLAG = 0                                                    '변수 초기화
    For i = 0 To 24                                                 'total 25매 wafer
        If NOW_NO(i) = False Then                                   'check를 하지 않은 경우 사용하는 것으로 설정
            CHK_FLAG = CHK_FLAG + 1                                 '변수를 증가
        End If
    Next i
    
    If CHK_FLAG = 25 Then                                           '아무 설정도 하지 않은 경우 처리
        MsgBox "Wafer 설정상태를 다시 확인해 주세요!!"              '경고 메시지로 다시 한번 확인
        Exit Sub
    End If
    
    If Barcode_Use = True Then
        If Text10(2).Text = "" Then
            MsgBox "DEVICE NAME을 입력해주세요!!"                   '경고 메시지로 다시 한번 확인
            Exit Sub
        End If
        If Text10(5).Text = "" Then
            MsgBox "OPERATOR NO를 입력해주세요!!"                   '경고 메시지로 다시 한번 확인
            Exit Sub
        End If
        If Text10(1).Text = "" Then
            MsgBox "LOT NO를 입력해주세요!!"                        '경고 메시지로 다시 한번 확인
            Exit Sub
        End If
        If Text10(3).Text = "" Then
            MsgBox "EQUIPMENT ID를 입력해주세요!!"                  '경고 메시지로 다시 한번 확인
            Exit Sub
        End If
        If Text10(6).Text = "" Then
            MsgBox "PROBE CARD NO를 입력해주세요!!"                 '경고 메시지로 다시 한번 확인
            Exit Sub
        End If
    Else
        If Text10(2).Text = "" Then
            MsgBox "DEVICE NAME을 입력해주세요!!"                   '경고 메시지로 다시 한번 확인
            Exit Sub
        End If
        If Text10(1).Text = "" Then
            MsgBox "LOT NO를 입력해주세요!!"                        '경고 메시지로 다시 한번 확인
            Exit Sub
        End If
    End If
    
        
    For i = 0 To 24
        If Check1(i).value = 1 Then                                 '사용안함
            Slot_No(i + 1) = True
        Else                                                        '사용함
            Slot_No(i + 1) = False
            Slot_Max_Count = i + 1                                  '측정할 wafer의 수를 증가한다.
        End If
        If Combo_X(i).Text <> Empty Then XPitch(i) = Combo_X(i)      '설정값을 변수에 적용 X step value
        If Combo_Y(i).Text <> Empty Then YPitch(i) = Combo_Y(i)      '설정값을 변수에 적용 Y step value
    Next i
    
    '[ 2017.08.18 ] : 기준 x,y step값을 cassette에서 설정하도록 수정.
    XPitch_MAIN = Text1.Text                                         'default X step value set
    YPitch_MAIN = Text2.Text                                         'default Y step value set
    StarProbe.MeasureStepX = XPitch_MAIN                            'X step value save값 설정
    StarProbe.MeasureStepY = YPitch_MAIN                            'Y step value save값 설정
    
    MT2000.Text1(0).Text = Text10(1).Text
    MT2000.Text1(1).Text = Text10(2).Text
    
    DEV = Text10(2).Text
    OPE = Text10(5).Text
    LOT = Text10(1).Text
    EQU = Text10(3).Text
    PRO = Text10(6).Text
    
    Log_Cnt = 0
    Log_Prn_Cnt = 0
    DataLog_flag = True
    
    Call StarProbe_FileSave_SystemInfo                              '설정내용을 starporbe.ifo 파일에 저장
    '''''
    
    Timer1.Enabled = False                                          '[ 바코드 ]
    
    MT2000.LOAD_CONTROL
End Sub

'[ 2018.01.29 ] : 4채널 사용시 0~4까지 선택 할 수 있도록 내용 수정.
Private Sub Form_Load()
    Dim i As Integer
    Dim j As Integer
    
    For i = 0 To 24
        If Slot_No(i + 1) = True Then
            Check1(i).value = 1                         '측정X
        Else
            Check1(i).value = 0                         '측정O
        End If
    Next i
    AutoAlign_Flag = True
    If AutoAlign_Flag = True Then
        Check7.value = 1                                'auto alignment on
    Else
        Check7.value = 0                                'auto alignment off
    End If
        
    '''''''''''''''[ 2017.08.18 ] : 기준 x,y step값을 cassette에서 설정하도록 수정.
    Text1.Text = StarProbe.MeasureStepX
    Text2.Text = StarProbe.MeasureStepY
    
    XPitch_MAIN = StarProbe.MeasureStepX
    YPitch_MAIN = StarProbe.MeasureStepY
    '''''''''''''''
    
    For i = 0 To 24
        If (XPitch(i) <> XPitch_MAIN) Then          'main step과 다른 경우
            Combo_X(i).ListIndex = XPitch(i)
        Else                                        'main step과 같은 경우
            Combo_X(i).ListIndex = XPitch_MAIN
        End If
        
        If (YPitch(i) <> YPitch_MAIN) Then          'main step과 다른 경우
            Combo_Y(i).ListIndex = YPitch(i)
        Else                                        'main step과 같은 경우
            Combo_Y(i).ListIndex = YPitch_MAIN
        End If
    Next i
    
    Text10(2).Text = DEV
    Text10(5).Text = OPE
    Text10(1).Text = LOT
    Text10(3).Text = EQU
    Text10(6).Text = PRO
    
    '[ 2020.09.17 ] : barcode use/not use
    If Barcode_Use = True Then
        SSPanel4(1).Visible = True
        txt_Barcode.Visible = True
        Command3.Visible = True
        
        Label2.Visible = True
        Label6.Visible = True
    Else
        SSPanel4(1).Visible = False
        txt_Barcode.Visible = False
        Command3.Visible = False
        
        Label2.Visible = False
        Label6.Visible = False
    End If
    
    If Mode_Set = False Then
        Text10(2).Enabled = False
        Text10(5).Enabled = False
        Text10(1).Enabled = False
        Text10(3).Enabled = False
    Else
        Text10(2).Enabled = True
        Text10(5).Enabled = True
        Text10(1).Enabled = True
        Text10(3).Enabled = True
    End If
    
    Timer1.Enabled = True       '[ 바코드 ]
End Sub

Private Sub Text10_Change(Index As Integer)
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
On Error GoTo ErrHandler
    Dim Array_str1() As String                                  '[ 2020.12.16 ] : 디바이스 이름에 "/" 유무 확인 변수
    Dim Find_Option_List  As Boolean
    Dim find_slash As Boolean
        
    Dim fso As Object                                           '[ 2020.10.29 ] : aoi
    Dim FolderList As Object                                    '[ 2020.10.29 ] : aoi
        
    If txt_Barcode.Text <> "" And Len(txt_Barcode.Text) > 3 Then       'barcode가 한글자라도 찍히면 delay를 주어서 글자입력 시간을 준다.
        find_slash = False
        
        DoEvents
        Sleep 10
        DoEvents
        txt_Barcode.Text = UCase(txt_Barcode.Text)

        'K1335-JM
        'K2727A-JM-Y
        'K5003A-PZ/SL
        If Left(UCase(txt_Barcode.Text), 1) = "K" Then
            For i = Len(txt_Barcode.Text) To 1 Step -1            '"-"의 수량을 파악한다.
                If Mid(txt_Barcode.Text, i, 1) = "-" Then
                    cnt = cnt + 1
                End If
            Next i
                        
            Array_str = Split(txt_Barcode.Text, "-")                  '"-"를 기준으로 나눈다.
            Array_tmp1 = Replace(Array_str(0), "K", "")
            Array_tmp1 = Mid(Array_str(0), 2, Len(Array_str(0)))
            Array_tmp2 = Array_str(1)
            Array_tmp3 = Array_str(2)
            
            '"/"이 있는 경우와 아닌 경우 구분 (5003,5004만 적용)
            If InStr(txt_Barcode.Text, "/") <> 0 Then                 '"/"있는 경우
                Array_str1 = Split(Array_str(1), "/")           '"/"를 기준으로 나눈다.
                Array_tmp2 = Array_str1(0)
                Array_tmp3 = Array_str1(1)
                                
                If InStr(txt_Barcode.Text, "5003") <> 0 Or InStr(txt_Barcode.Text, "5004") <> 0 Then
                    Text10(2).Text = Array_tmp1 & Array_tmp3     'xxxxzz
                Else
                    Text10(2).Text = Array_tmp1                  'xxxx
                End If
            Else                                                '"/"없는 경우
                If cnt = 2 Then                                 'Kxxxx-yy-z
                    Text10(2).Text = Array_tmp1 & Array_tmp3     'xxxxz
                Else                                            'Kxxxx-xx
                    Text10(2).Text = Array_tmp1                  'xxxx
                End If
            End If
            
            'option file list비교
            FileName = "C:\Star Probe\OPTION_LIST.txt"
            
            '파일존재여부확인
            If LenB(Dir$(FileName)) Then
                Open FileName For Input As #1                   'option list file이 없는 경우 에러처리 필요하다.
                    Find_Option_List = False
                    While Not EOF(1)
                        Line Input #1, tmp
                        If Array_tmp1 = tmp Then
                            Find_Option_List = True
                        End If
                    Wend
                Close #1
                
                If Find_Option_List = True Then
                    MsgBox "Option List detect 수동으로 파일을 열어주세요!!"
                    Text10(2).Text = ""
                End If
            Else
                MsgBox "C:\Star Probe\OPTION_LIST.txt 파일이 없습니다."
            End If
            txt_Barcode.Text = ""
        End If
        
        If InStr(UCase(txt_Barcode.Text), "OP") Then                                                     'operator : OP00123456 -> OP00빼고 나머지를 표시한다.
            Text10(5).Text = Trim(UCase(Mid(txt_Barcode.Text, 5, Len(txt_Barcode.Text))))
            txt_Barcode.Text = ""
        ElseIf InStr(UCase(txt_Barcode.Text), "TEPR") Then                                               'equipment : TEPR00000 -> 전부표시한다.
            Text10(3).Text = UCase(Mid(txt_Barcode.Text, 1, Len(txt_Barcode.Text)))
            txt_Barcode.Text = ""
        ElseIf InStr(UCase(txt_Barcode.Text), "PC") Then                                                 'probe card id : PC00000 -> 전부표시한다.
            Text10(6).Text = UCase(Mid(txt_Barcode.Text, 1, Len(txt_Barcode.Text)))
            txt_Barcode.Text = ""
        Else                                                                                            'lot no : 기본형과 aoi 두가지가 구분이 되어야 한다.
            If Len(txt_Barcode.Text) >= 8 Then
                DoEvents
                Sleep 10
                DoEvents
                Text10(1).Text = Left(UCase(txt_Barcode.Text), 8)
                'Text10(1) = Mid(Trim(UCase(txt_Barcode)), 1, Len(txt_Barcode))
                txt_Barcode.Text = ""
            End If
        End If
    End If
    
    If AOI_MODE = 1 Then
        'AOI는 모든내용이 입력이 되면 작동한다.
        If Text10(2) <> "" And Text10(5) <> "" And Text10(1) <> "" And Text10(3) <> "" And Text10(6) <> "" Then
            If Array_tmp1 = "" Then                 '[  2021.02.04 ]
                Array_tmp1 = Text10(2)
            End If
            If Array_tmp1 = "" Then
                AOI_Use = False
            Else
                If Len(Text10(1).Text) > 2 Then             '[ 2021.10.12 ] : 7 -> 2로 수정  LOT
                    '==================================================================================================================
                    '[ AOI 관련 추가 부분 ]
                    '==================================================================================================================
                    'ex)Z:\AOI\TR\2029K166-01.aoi
                    If Right(AOI_path, 1) = "\" Then
                        FileName = AOI_path
                    Else
                        FileName = AOI_path & "\"
                    End If
                    
                    '파일존재여부확인 (지정파일이 있으면 AOI모드 없으면 일반모드 적용한다.
                    If Dir(FileName, vbDirectory) <> "" Then
                        AOI_Use = True
                        '[ 2020.09.14 ] : map이름 찾기 함수
                        Set fso = CreateObject("Scripting.FileSystemObject")
                        
                        'AOI파일위치
                        Set FolderList = fso.GetFolder(FileName)                                    '(1)지정폴더에 모두 들어가는 경우
                        Call Get_Folder(FolderList)
                    Else
                        AOI_Use = False
                    End If
                    Timer1.Enabled = False
                    '==================================================================================================================
                End If
            End If
        End If
    End If
    Exit Sub
    
ErrHandler:
    Resume Next
End Sub

'[ 2020.09.14 ] : 폴더에서 aoi 파일을 찾는 함수 추가
Public Sub Get_Folder(folder As Object)
On Error GoTo ErrorSub
    Dim find_aoi As Boolean
    Dim f As Object
    Dim Array_backup() As String
    Dim i As Integer
    Dim strTmp As String
    
    find_aoi = False
    
    'AOI파일 리스트 클리어
    For i = 0 To 23
        Form_AOI_LIST.Text1(i).Text = ""
    Next i
    
    '카세트 체크 클리어
    If AOI_Use = False Then
        For i = 0 To 24
            Check1(i).value = 1
        Next i
    End If
    
    strTmp = UCase(Form_Cassette_New.Text10(1).Text)                                                     'strTmp 변수에 lot no값을 할당한다.
        
    For Each f In folder.Files
        If InStr(UCase(f), strTmp) <> 0 And Right(UCase(f), 3) = "AOI" Then                             'LOT NO 가 포함되고 확장자명이 "AOI"인 경우
            For i = 0 To 24
                Check1(i).value = 1
            Next i
            Exit For
        End If
    Next
    
    For Each f In folder.Files
        If InStr(UCase(f), strTmp) <> 0 And Right(UCase(f), 3) = "AOI" Then                             'LOT NO 가 포함되고 확장자명이 "AOI"인 경우
            Array_backup = Split(f, "-")                                                                '파일 이름을 나눈다.
            i = Left(Array_backup(1), 2)                                                                'wafer no를 추출한다. : [ LOT-01.AOI ] Array_backup(0) = LOT, Array_backup(1) = NO
            AOI_MAP(i) = f
            Form_AOI_LIST.Text1(i - 1).Text = AOI_MAP(i)                                                 '임시저장 확인 용.
            Check1(i - 1).value = False                                                                 '웨이퍼가 있는 경우 체크박스 해제한다.(측정전이라는 표시)
            find_aoi = True
        End If
    Next
    
    If find_aoi = False Then
        AOI_Use = False
    End If
    Exit Sub

ErrorSub:
    Load_MAP = ""
End Sub

Private Sub txt_Barcode_Change()
    Timer1.Enabled = True
End Sub
