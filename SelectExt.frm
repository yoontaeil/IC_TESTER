VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form SelectExt 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "Star Probe Option"
   ClientHeight    =   12900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8040
   Icon            =   "SelectExt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12900
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Visible         =   0   'False
   Begin VB.CheckBox Check4 
      Caption         =   "Tester On"
      Height          =   180
      Left            =   4200
      TabIndex        =   95
      Top             =   12480
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "침적확인"
      Height          =   615
      Left            =   4200
      TabIndex        =   94
      Top             =   11640
      Width           =   1095
   End
   Begin VB.Frame Frame12 
      Height          =   1215
      Left            =   0
      TabIndex        =   86
      Top             =   12960
      Width           =   4095
      Begin VB.TextBox txt_X 
         Alignment       =   2  '가운데 맞춤
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   89
         Text            =   "1"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txt_Y 
         Alignment       =   2  '가운데 맞춤
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   88
         Text            =   "1"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txt_Z 
         Alignment       =   2  '가운데 맞춤
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   87
         Text            =   "1"
         Top             =   720
         Width           =   735
      End
      Begin Threed.SSPanel SSPanel7 
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   90
         Top             =   240
         Width           =   3735
         _Version        =   65536
         _ExtentX        =   6588
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "CLEAN TIP POSITION"
         ForeColor       =   4210752
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FloodColor      =   12582912
      End
      Begin VB.Label Label2 
         Caption         =   "X : "
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   93
         Top             =   800
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Y : "
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   1440
         TabIndex        =   92
         Top             =   800
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Z : "
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   2760
         TabIndex        =   91
         Top             =   800
         Width           =   495
      End
   End
   Begin VB.Frame Frame11 
      Height          =   1215
      Left            =   0
      TabIndex        =   81
      Top             =   11520
      Width           =   4095
      Begin VB.OptionButton Option3 
         Caption         =   "1CH"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   85
         Top             =   840
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton Option3 
         Caption         =   "2CH"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   1560
         TabIndex        =   84
         Top             =   840
         Width           =   855
      End
      Begin VB.OptionButton Option3 
         Caption         =   "4CH"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   3000
         TabIndex        =   83
         Top             =   840
         Width           =   855
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   375
         Left            =   120
         TabIndex        =   82
         Top             =   240
         Width           =   3735
         _Version        =   65536
         _ExtentX        =   6588
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "CHANNEL SELECT"
         ForeColor       =   4210752
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FloodColor      =   12582912
      End
   End
   Begin VB.Frame Frame10 
      Height          =   1215
      Left            =   4200
      TabIndex        =   69
      Top             =   10320
      Width           =   3735
      Begin VB.CheckBox Check9 
         Caption         =   "Barcode USE"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   74
         Top             =   720
         Value           =   1  '확인
         Width           =   2535
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Re-Measure"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   73
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  '가운데 맞춤
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   72
         Text            =   "1"
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  '가운데 맞춤
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   71
         Text            =   "1"
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox Text9 
         Alignment       =   2  '가운데 맞춤
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   70
         Text            =   "1"
         Top             =   2160
         Width           =   615
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   375
         Left            =   120
         TabIndex        =   75
         Top             =   240
         Width           =   3495
         _Version        =   65536
         _ExtentX        =   6165
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "BARCODE OPTION"
         ForeColor       =   4210752
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FloodColor      =   12582912
      End
      Begin VB.Label Label2 
         Caption         =   "Rotation Count"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   77
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Rotation Count(Sub)"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   76
         Top             =   2280
         Width           =   2295
      End
   End
   Begin VB.Frame Frame9 
      Height          =   3135
      Left            =   0
      TabIndex        =   62
      Top             =   7200
      Width           =   7935
      Begin VB.CommandButton Command5 
         Caption         =   "Search"
         Height          =   390
         Left            =   6720
         TabIndex        =   79
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox Text10 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   120
         TabIndex        =   78
         Top             =   2640
         Width           =   6615
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   120
         TabIndex        =   66
         Top             =   1680
         Width           =   6615
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Search"
         Height          =   390
         Left            =   6720
         TabIndex        =   65
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Search"
         Height          =   390
         Left            =   6720
         TabIndex        =   64
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   120
         TabIndex        =   63
         Top             =   720
         Width           =   6615
      End
      Begin Threed.SSPanel SSPanel7 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   67
         Top             =   240
         Width           =   7695
         _Version        =   65536
         _ExtentX        =   13573
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "DATA SERVER PATH"
         ForeColor       =   4210752
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FloodColor      =   12582912
      End
      Begin Threed.SSPanel SSPanel7 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   68
         Top             =   1200
         Width           =   7695
         _Version        =   65536
         _ExtentX        =   13573
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "MAP DATA PATH"
         ForeColor       =   4210752
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FloodColor      =   12582912
      End
      Begin Threed.SSPanel SSPanel7 
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   80
         Top             =   2160
         Width           =   7695
         _Version        =   65536
         _ExtentX        =   13573
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "AOI PATH"
         ForeColor       =   4210752
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FloodColor      =   12582912
      End
   End
   Begin VB.Frame Frame8 
      Height          =   1215
      Left            =   0
      TabIndex        =   55
      Top             =   10320
      Width           =   4095
      Begin VB.CheckBox Check_Test_Fail_count 
         BackColor       =   &H008080FF&
         Caption         =   "TEST FAIL COUNT :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   57
         Top             =   720
         Width           =   2415
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  '가운데 맞춤
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   56
         Text            =   "1"
         Top             =   720
         Width           =   1095
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   375
         Left            =   120
         TabIndex        =   58
         Top             =   240
         Width           =   3735
         _Version        =   65536
         _ExtentX        =   6588
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "PROBE STOP TEST FAIL COUNT"
         ForeColor       =   4210752
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FloodColor      =   12582912
      End
   End
   Begin VB.TextBox Text_DelayTime 
      Alignment       =   2  '가운데 맞춤
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   53
      Text            =   "0"
      Top             =   12960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CheckBox Check_WaferTest 
      Caption         =   "Wafer Test"
      Height          =   375
      Left            =   120
      TabIndex        =   51
      Top             =   10320
      Value           =   1  '확인
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame5 
      Height          =   2895
      Left            =   4200
      TabIndex        =   35
      Top             =   4200
      Width           =   3735
      Begin VB.OptionButton Option1 
         Caption         =   "OFF"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   2520
         TabIndex        =   49
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "DIRECT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   48
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "AFTER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   45
         Top             =   720
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.Frame Frame7 
         Height          =   975
         Left            =   120
         TabIndex        =   39
         Top             =   960
         Width           =   3495
         Begin VB.TextBox Text6 
            Alignment       =   2  '가운데 맞춤
            Enabled         =   0   'False
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
            Left            =   2280
            TabIndex        =   44
            Text            =   "2"
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox Text6 
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
            Index           =   0
            Left            =   1080
            TabIndex        =   43
            Text            =   "1"
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "LEFT INK"
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
            Index           =   8
            Left            =   1080
            TabIndex        =   42
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "RIGHT PORT"
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
            Index           =   7
            Left            =   2160
            TabIndex        =   41
            Top             =   240
            Width           =   1035
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "INK PORT"
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
            Index           =   6
            Left            =   120
            TabIndex        =   40
            Top             =   480
            Width           =   810
         End
      End
      Begin VB.Frame Frame6 
         Height          =   855
         Left            =   120
         TabIndex        =   37
         Top             =   1920
         Width           =   3495
         Begin VB.CheckBox Check2 
            Caption         =   "CENTER INK"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   2
            Left            =   1080
            TabIndex        =   61
            Top             =   480
            Width           =   1335
         End
         Begin VB.CheckBox Check2 
            Caption         =   "RIGHT INK"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   1
            Left            =   2280
            TabIndex        =   47
            Top             =   240
            Width           =   1095
         End
         Begin VB.CheckBox Check2 
            Caption         =   "LEFT INK"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   0
            Left            =   1080
            TabIndex        =   46
            Top             =   240
            Value           =   1  '확인
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "INK PORT"
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
            Index           =   5
            Left            =   120
            TabIndex        =   38
            Top             =   240
            Width           =   810
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   375
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   3495
         _Version        =   65536
         _ExtentX        =   6165
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "INK OPTION"
         ForeColor       =   4210752
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FloodColor      =   12582912
      End
   End
   Begin VB.Frame Frame4 
      Height          =   2775
      Left            =   4200
      TabIndex        =   27
      Top             =   1440
      Width           =   3735
      Begin VB.TextBox Text_Tipclean_Count_Limit 
         Alignment       =   1  '오른쪽 맞춤
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   60
         Text            =   "100000000"
         Top             =   2280
         Width           =   1695
      End
      Begin VB.CheckBox Check_Tip_Clean 
         Caption         =   "Tip-Clean"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   59
         Top             =   2280
         Width           =   1815
      End
      Begin VB.CheckBox Check_AllTest 
         Caption         =   "All Measure"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   52
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox Text_rcount_sub 
         Alignment       =   2  '가운데 맞춤
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   32
         Text            =   "1"
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox Text_rcount 
         Alignment       =   2  '가운데 맞춤
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   31
         Text            =   "1"
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox Text_LineOk 
         Alignment       =   2  '가운데 맞춤
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   30
         Text            =   "1"
         Top             =   1080
         Width           =   615
      End
      Begin VB.CheckBox Check_ReMeasure 
         Caption         =   "Re-Measure"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   1080
         Width           =   1815
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   3495
         _Version        =   65536
         _ExtentX        =   6165
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "MEASURE OPTION"
         ForeColor       =   4210752
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FloodColor      =   12582912
      End
      Begin VB.Label Label2 
         Caption         =   "Rotation Count(Sub)"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   34
         Top             =   1920
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "Rotation Count"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   33
         Top             =   1560
         Width           =   1935
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1335
      Left            =   4200
      TabIndex        =   22
      Top             =   0
      Width           =   3735
      Begin VB.CheckBox Check3 
         Caption         =   "Limit Area Select"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   50
         Top             =   720
         Width           =   2655
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
         Left            =   3960
         TabIndex        =   25
         Text            =   "1"
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox Text1 
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
         Left            =   3000
         TabIndex        =   24
         Text            =   "3"
         Top             =   1440
         Width           =   855
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   3495
         _Version        =   65536
         _ExtentX        =   6165
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "X,Y STEP MODE"
         ForeColor       =   4210752
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FloodColor      =   12582912
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
         Left            =   1800
         TabIndex        =   26
         Top             =   1320
         Width           =   870
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1575
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   4095
      Begin VB.OptionButton Option2 
         Caption         =   "GPIB Use TO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   20
         Top             =   720
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         Caption         =   "GPIB Not Use TO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   19
         Top             =   1080
         Width           =   2535
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   3855
         _Version        =   65536
         _ExtentX        =   6800
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "PROBE USE SELECT"
         ForeColor       =   4210752
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FloodColor      =   12582912
      End
      Begin VB.Label Label1 
         Caption         =   "(2001X .....Addr: 1)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   21
         Top             =   720
         Width           =   1680
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5535
      Left            =   0
      TabIndex        =   2
      Top             =   1560
      Width           =   4095
      Begin VB.CheckBox Check1 
         Caption         =   "1. MF/MC ON X-Y MOTION"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   16
         Top             =   720
         Width           =   3015
      End
      Begin VB.CheckBox Check1 
         Caption         =   "2. MF/MC ON Z MOTION"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   15
         Top             =   1080
         Width           =   2775
      End
      Begin VB.CheckBox Check1 
         Caption         =   "3. MF/MC ON OPT DEVICES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   14
         Top             =   1440
         Width           =   3015
      End
      Begin VB.CheckBox Check1 
         Caption         =   "4. MF/MC ON REST OF CMDS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   13
         Top             =   1800
         Width           =   3495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "5. TEST START MESSAGE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   12
         Top             =   2160
         Width           =   3255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "6. TEST COMPLETE MESSAGE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   11
         Top             =   2520
         Width           =   3255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "7. PATTERN COMPLETE MESSAGE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   10
         Top             =   2880
         Width           =   3495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "8. PAUSE/CONTINUE MESSAGE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   9
         Top             =   3240
         Width           =   3375
      End
      Begin VB.CheckBox Check1 
         Caption         =   "9. ALARM MESSAGE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   8
         Top             =   3600
         Width           =   2295
      End
      Begin VB.CheckBox Check1 
         Caption         =   "10. WAFER COMPLETE MESSAGE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   7
         Top             =   3960
         Width           =   3375
      End
      Begin VB.CheckBox Check1 
         Caption         =   "11. ENHANCED PC MESSAGE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   240
         TabIndex        =   6
         Top             =   4320
         Width           =   3135
      End
      Begin VB.CheckBox Check1 
         Caption         =   "12. ENHANCED TS MESSAGE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   240
         TabIndex        =   5
         Top             =   4680
         Width           =   3015
      End
      Begin VB.CheckBox Check1 
         Caption         =   "13. PAUSE PENDING MESSAGE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   240
         TabIndex        =   4
         Top             =   5040
         Width           =   3135
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   3855
         _Version        =   65536
         _ExtentX        =   6800
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "EXTERNAL I/O MODE MENU"
         ForeColor       =   4210752
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FloodColor      =   12582912
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6840
      Picture         =   "SelectExt.frx":08CA
      Style           =   1  '그래픽
      TabIndex        =   1
      Top             =   11640
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5640
      Picture         =   "SelectExt.frx":108C
      Style           =   1  '그래픽
      TabIndex        =   0
      Top             =   11640
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Measure Delay Time"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   2160
      TabIndex        =   54
      Top             =   13200
      Visible         =   0   'False
      Width           =   2295
   End
End
Attribute VB_Name = "SelectExt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hmenu As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hmenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Private Const MF_BYPOSITION = &H400&
Private Const MF_DISABLED = &H2&

Dim OptCategory As String
Dim Readbuf  As String
Dim GpibTmp As Long

Private Sub Check_AllTest_Click()
    Check3.value = IIf(Check_AllTest.value = 0, 1, 0)
End Sub

Private Sub Check_ReMeasure_Click()
    Text_LineOk.Enabled = IIf(Check_ReMeasure.value = 1, True, False)
End Sub

Private Sub Check_Test_Fail_count_Click()
    If Check_Test_Fail_count.value = 1 Then
        '[ 2021.12.31 ] : mode
        If Mode_Set = False Then
            Text3.Enabled = False
        Else
            Text3.Enabled = True
        End If
    Else
        Text3.Enabled = False
    End If
End Sub

Private Sub Check_Tip_Clean_Click()
    Text_Tipclean_Count_Limit.Enabled = IIf(Check_Tip_Clean.value = 1, True, False)
End Sub

Private Sub Check2_Click(Index As Integer)
    Dim i, j As Integer

    If Check2(2).value = 1 Then
        Label1(7).Caption = ""
        Label1(8).Caption = "CENTER INK"
        Text6(1).Enabled = False
    ElseIf Check2(2).value = 0 Then
        Label1(7).Caption = "RIGHT PORT"
        Label1(8).Caption = "LEFT INK"
        Text6(1).Enabled = True
    End If
    
    Text6(0).Enabled = IIf(Check2(0).value = 1, True, False)
    Text6(1).Enabled = IIf(Check2(1).value = 1, True, False)
    
    j = 0
    For i = 0 To 1
        If Check2(i).value = 1 And Check2(2).value = 1 Then j = j + 1
    Next i

    If j > 0 Then
        MsgBox "Only One Check !!!!", vbOKOnly
        Check2(0).value = 0
        Check2(1).value = 0
        Check2(2).value = 0
    End If
End Sub

Private Sub Check3_Click()
    Check_AllTest.value = IIf(Check3.value = 0, 1, 0)
End Sub

Private Sub Check4_Click()
    If Check4.Caption = "Tester Off" Then
        Check4.Caption = "Tester On"
        TESTER_OFF = False
    Else
        Check4.Caption = "Tester Off"
        TESTER_OFF = True
    End If
End Sub

Private Sub Check9_Click()
    If Check9.value = 0 Then
        Barcode_Use = False
        Check9.Caption = "Barcode not use"
    Else
        Barcode_Use = True
        Check9.Caption = "Barcode use"
    End If
End Sub

Private Sub Command1_Click()
    Server_path = Text4.Text
    MAP_path = Text5.Text                   '2020.09.07 : map path add
    If AOI_MODE = 1 Then AOI_path = Text10.Text
    
    '[ 2021.12.13 ] : ch select
    If Option3(0).value = True Then
        CH_SET = 1
    ElseIf Option3(1).value = True Then
        CH_SET = 2
    Else
        CH_SET = 4
    End If
    MT2000.Label15.Caption = CH_SET & "CH"
    
    If Option2(0).value = True Then
'        XPitch = Text1.Text
'        YPitch = Text2.Text
            
        StarProbe.MeasureStepX = val(Text1)
        StarProbe.MeasureStepY = val(Text2)
            
        StarProbe.Ink_LeftPort = Text6(0)
        StarProbe.Ink_RightPort = Text6(1)
        StarProbe.LineOk = val(Text_LineOk)
        StarProbe.RCount = val(Text_rcount)
        StarProbe.RCount_Sub = val(Text_rcount_sub)
        StarProbe.MeasureSleep = val(Text_DelayTime)
        StarProbe.ReMeasure = Check_ReMeasure.value
        StarProbe.Ink_After_LeftPort = IIf(Check2(0).value = 1, 1, 0)
        StarProbe.Ink_After_RightPort = IIf(Check2(1).value = 1, 1, 0)
        StarProbe.Ink_After_CenterPort = IIf(Check2(2).value = 1, 1, 0)
        StarProbe.LimitArea = IIf(Check3.value = 1, 1, 0)
            
        StarProbe.Tip_Clean = Check_Tip_Clean.value
        StarProbe.Tipclean_Count_Limit = val(Text_Tipclean_Count_Limit)
            
        If StarProbe.Tipclean_Count_Limit > 100000000 Then StarProbe.Tipclean_Count_Limit = 100000000
            
        StarProbe.WaferTest = Check_WaferTest.value
        StarProbe.MeasureAll = Check_AllTest.value
            
        StarProbe.Probe_Stop_Tfail_Count = val(Text3)
        StarProbe.Probe_Stop = IIf(Check_Test_Fail_count.value = 1, 1, 0)
            
        If Option1(0).value = True Then StarProbe.Ink_After = 0
        If Option1(1).value = True Then StarProbe.Ink_After = 1
        If Option1(2).value = True Then StarProbe.Ink_After = 2
                                    
        For i = 1 To 13
            If Check1(i).value = 1 Then
                OptCategory = OptCategory & "1"
            Else
                OptCategory = OptCategory & "0"
            End If
            ProbeioMsg(i) = Check1(i).value
        Next i
                        
        If DemoMode = 0 Then
            If IO_2001X = 0 Then
                GpibAdd = StarProbe_Address()
                If OptCategory <> Empty Then Call StarProbe_Option_Set(OptCategory)
            End If
        End If
        OptCategory = " "
        Opt_Select_Flag = True
        
    ElseIf Option2(0).value = False Then
        Opt_Select_Flag = False
        Exit Sub
    End If
    Call StarProbe_FileSave_SystemInfo
    Unload Me
End Sub

Private Sub Command2_Click()
    Opt_Select_Flag = False
    Unload Me
    Exit Sub
End Sub

Private Sub Command3_Click()
    Path_Check = 1
    SetDataPath.Show 1
End Sub

Private Sub Command4_Click()
    Path_Check = 2
    SetDataPath.Show 1
End Sub

Private Sub Command5_Click()
    Path_Check = 3
    SetDataPath.Show 1
End Sub

'[ 2022.07.29 ]
Private Sub Command6_Click()
    '[ 2022.07.20 ]
    If LOG_FILE_ON = 1 Then SelectExt.Log_Data_Save (3)
    Form_Needle.Show 1
End Sub

Private Sub Form_Load()
    Dim hmenu As Long
    Dim nCount As Long
        
    OptCategory = " "
    
    hmenu = GetSystemMenu(Me.hwnd, 0)
    nCount = GetMenuItemCount(hmenu)
 
    Text1.Text = StarProbe.MeasureStepX
    Text2.Text = StarProbe.MeasureStepY
    
    Text_LineOk = StarProbe.LineOk
    Text_rcount = StarProbe.RCount
    Text_rcount_sub = StarProbe.RCount_Sub
    Text_DelayTime = StarProbe.MeasureSleep
    
    Check_ReMeasure.value = StarProbe.ReMeasure
    Check2(0).value = IIf(StarProbe.Ink_After_LeftPort = 1, 1, 0)
    Check2(1).value = IIf(StarProbe.Ink_After_RightPort = 1, 1, 0)
    Check2(2).value = IIf(StarProbe.Ink_After_CenterPort = 1, 1, 0)
    Check3.value = IIf(StarProbe.LimitArea = 1, 1, 0)
    
    Check_Tip_Clean.value = StarProbe.Tip_Clean
    Text_Tipclean_Count_Limit = StarProbe.Tipclean_Count_Limit
      
    Check_WaferTest.value = StarProbe.WaferTest
    Check_AllTest.value = StarProbe.MeasureAll
    
    Text3.Text = StarProbe.Probe_Stop_Tfail_Count
    Check_Test_Fail_count.value = IIf(StarProbe.Probe_Stop = 1, 1, 0)
        
    Text3.Enabled = IIf(Check_Test_Fail_count.value = 1, True, False)
    
    For i = 1 To 13
        Check1(i).value = ProbeioMsg(i)
    Next
    
    If StarProbe.Ink_After = 0 Then              '추가
        Option1(0).value = True
        Option1(1).value = False
        Option1(2).value = False
    ElseIf StarProbe.Ink_After = 1 Then
        Option1(0).value = False
        Option1(1).value = True
        Option1(2).value = False
    ElseIf StarProbe.Ink_After = 2 Then
        Option1(0).value = False
        Option1(1).value = False
        Option1(2).value = True
    End If
    
    Text4.Text = Server_path
    Text5.Text = MAP_path               '2020.09.07 : map path add
    
    If AOI_MODE = 1 Then Text10.Text = AOI_path
    
    '[ 2020.09.17 ] : barcode use,not use select
    Check9.value = IIf(Barcode_Use = True, 1, 0)
    
    If AOI_MODE = 0 Then
        SSPanel7(2).Visible = False
        Text10.Visible = False
        Command5.Visible = False
    Else
        SSPanel7(2).Visible = True
        Text10.Visible = True
        Command5.Visible = True
    End If
    
    If CH_SET = 1 Then
        SelectExt.Option3(0).value = True
    ElseIf CH_SET = 2 Then
        SelectExt.Option3(1).value = True
    Else
        SelectExt.Option3(2).value = True
    End If
    
    Call RemoveMenu(hmenu, nCount - 1, MF_DISABLED Or MF_BYPOSITION)
    DrawMenuBar Me.hwnd
    
    '[ 2021.12.31 ] : mode
    If Mode_Set = False Then
        Text3.Enabled = False            '[ 2022.06.15 ] : false->true임시 변경
        Check4.Visible = False          '[ 2022.08.10 ] : tester on/off visible
        TESTER_OFF = False
    Else
        If DemoMode = 1 Then
            Check4.Visible = True           '[ 2022.08.10 ] : tester on/off visible
        End If
        Text3.Enabled = True
    End If
    
    '[ 2022.06.24 ]
    txt_X.Text = StarProbe.Tipclean_X
    txt_Y.Text = StarProbe.Tipclean_Y
    txt_Z.Text = StarProbe.Tipclean_Z
End Sub

Private Sub Option1_Click(Index As Integer)
    Select Case Index
        Case 0
            Check2(0).Enabled = False
            Check2(1).Enabled = False
            Text6(0).Enabled = True
            Text6(1).Enabled = True
        Case 1
            Check2(0).Enabled = True
            Check2(1).Enabled = True
            Text6(0).Enabled = True
            Text6(1).Enabled = True
        Case 2
            Check2(0).Enabled = False
            Check2(1).Enabled = False
            Text6(0).Enabled = False
            Text6(1).Enabled = False
    End Select
End Sub

Private Sub Option2_Click(Index As Integer)
    Label1(2).Visible = IIf(Index = 0, True, False)
End Sub

'[ 2022.07.20 ] : 주요동작의 사용 로그파일을 만드는 서브루틴
Public Sub Log_Data_Save(no As Integer)
    Dim myFSO As Object                                                     '[ 2022.09.02 ] : File create day 관련
    Dim f As Object                                                         '[ 2022.09.02 ] : File create day 관련
    Set myFSO = CreateObject("Scripting.FileSystemObject")                  '[ 2022.09.02 ] : File create day 관련
    
    Dim sfilename As String             '파일이름
    Dim ifreefile As Integer            '파일처리열 확인
    Dim strTmp As String                '파일내용 임시저장
    Dim Start_Day As String             '파일최초 생성일 저장
    Dim Diff As Integer                 '현재-파일최초생성일 차이를 저장
        
    strTmp = ""
    
    '저장파일경로
    sfilename = "c:\Star Probe\Starprobe_Log.dat"
    
    '파일 존재유무를 확인
    If LenB(Dir$(sfilename)) Then
        '파일의 최초 생성일 추출
        Set f = myFSO.GetFile(sfilename)                                    '[ 2022.09.02 ] : File create day 관련
        Start_Day = f.DateCreated                                           '[ 2022.09.02 ] : File create day 관련
        
        '현재와 파일최초 생성일간의 차이를 날짜로 구한다.
        Diff = DateDiff("d", Start_Day, Now)
        
        '용량이 커지는것과 검색의 어려움을 막기위해서 30일 주기로 파일 이름을 백업한다.
        If Diff >= 30 Then
            '파일이름을 변경해서 저장하고 기존파일은 새로시작하기 위해서 삭제한다.
            FileCopy sfilename, "c:\Star Probe\Starprobe_Log( " & Mid(Start_Day, 1, 10) & " ~ " & Mid(Now, 1, 10) & " ).dat"   'ex)c:\Star Probe\Starprobe_Log(2022-07-01~2022.07.30).dat
            Kill sfilename
        End If
    End If
    
    ifreefile = FreeFile
    
    '시간과 함께 주요 동작을 log파일로 저장한다.
    Open sfilename For Append As ifreefile
        strTmp = "[ " & Now & " ] : "
        Select Case no
            Case 1:
                strTmp = strTmp & "Engineer mode on"                                    'engineer mode
            Case 2:
                strTmp = strTmp & "Operator mode on"                                    'operator mode
            Case 3:
                strTmp = strTmp & "침적확인 설정에 접근하였습니다."
            Case 4:
                strTmp = strTmp & "AUTO ON" & "," & "MAP:" & MT2000.SSPanel2(0).Caption & "," & "LOTNO:" & MT2000.Text1(0).Text
            Case 5:
                strTmp = strTmp & "AUTO OFF" & "," & "MAP:" & MT2000.SSPanel2(0).Caption & "," & "LOTNO:" & MT2000.Text1(0).Text
            Case 6:
                strTmp = strTmp & "침적확인 메시지 발생" & "," & "MAP:" & MT2000.SSPanel2(0).Caption & "," & "LOTNO:" & MT2000.Text1(0).Text & "," & "WaferNo:" & MT2000.lblWafer.Caption
            Case 7:
                strTmp = strTmp & "Sample test ON & Inking ON"
            Case 8:
                strTmp = strTmp & "Sample test OFF & Inking OFF"
            Case 9:
                strTmp = strTmp & "Only edge ink ON"
            Case 10:
                strTmp = strTmp & "Only edge ink OFF"
            Case 11:
                strTmp = strTmp & "Sample test no ink ON"
            Case 12:
                strTmp = strTmp & "Sample test no ink OFF"
            Case 13:
                strTmp = strTmp & "Tip Clean을 실시하였습니다."
            Case 14:
                strTmp = strTmp & "침적 위치 변화검사를 체크 하였습니다."
            Case 15:
                strTmp = strTmp & "침적 세기(Chuck 높이)검사를 체크 하였습니다."
            Case 16:
                strTmp = strTmp & "침적 세기(Probe Card)검사를 체크 하였습니다."
            Case 17:
                strTmp = strTmp & "편마모 검사를 체크 하였습니다."
            Case 18:
                strTmp = strTmp & MSG_DATA
            Case 19:
                strTmp = strTmp & "작업진행을 클릭 하였습니다."
            Case 20:
                strTmp = strTmp & "침적 벗어남(불량처리)를 확인 하였습니다."
            Case Else
                strTmp = strTmp & ""
        End Select
        Print #ifreefile, strTmp
    Close ifreefile
End Sub



