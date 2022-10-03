VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form MT2000 
   BackColor       =   &H00E0E0E0&
   Caption         =   "SPC-2001X "
   ClientHeight    =   14850
   ClientLeft      =   1920
   ClientTop       =   2370
   ClientWidth     =   19080
   FillColor       =   &H00C0C0C0&
   BeginProperty Font 
      Name            =   "µ¸¿ò"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C0C0C0&
   Icon            =   "mt2000.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1815.078
   ScaleMode       =   0  '»ç¿ëÀÚ
   ScaleWidth      =   1272
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "FLAT"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1001
      Left            =   13800
      TabIndex        =   358
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "»ó"
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   362
         Top             =   200
         Width           =   495
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ÇÏ"
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   361
         Top             =   680
         Value           =   -1  'True
         Width           =   495
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ÁÂ"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   360
         Top             =   440
         Width           =   495
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "¿ì"
         Height          =   255
         Index           =   3
         Left            =   1080
         TabIndex        =   359
         Top             =   440
         Width           =   495
      End
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Inker Position"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1001
      Left            =   15600
      TabIndex        =   322
      Top             =   120
      Width           =   2175
      Begin VB.OptionButton Option6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "»ó"
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   327
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ÇÏ"
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   326
         Top             =   720
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ÁÂ"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   325
         Top             =   480
         Width           =   735
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Áß"
         Height          =   255
         Index           =   3
         Left            =   840
         TabIndex        =   324
         Top             =   480
         Width           =   735
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "¿ì"
         Height          =   255
         Index           =   4
         Left            =   1560
         TabIndex        =   323
         Top             =   480
         Width           =   495
      End
   End
   Begin VB.TextBox Text11 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10920
      TabIndex        =   317
      Text            =   "3"
      Top             =   240
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Wafer No"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1001
      Left            =   17880
      TabIndex        =   314
      Top             =   120
      Width           =   1095
      Begin VB.Label lblWafer 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BackColor       =   &H00E0E0E0&
         Caption         =   "11"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   315
         Top             =   340
         Width           =   855
      End
   End
   Begin VB.CommandButton Command_Stop 
      BackColor       =   &H000000FF&
      Caption         =   "STOP"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   542
      Left            =   120
      Style           =   1  '±×·¡ÇÈ
      TabIndex        =   281
      TabStop         =   0   'False
      Top             =   13920
      Width           =   1215
   End
   Begin VB.TextBox Text10 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10920
      TabIndex        =   270
      Text            =   "10000"
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   15000
      TabIndex        =   269
      Top             =   9720
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   14640
      TabIndex        =   268
      Top             =   10320
      Visible         =   0   'False
      Width           =   3855
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   4800
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab4 
      Height          =   13725
      Left            =   120
      TabIndex        =   130
      Top             =   120
      Width           =   5805
      _ExtentX        =   10239
      _ExtentY        =   24209
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "WORK FORM"
      TabPicture(0)   =   "mt2000.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "SSPanel5(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "BIN FORM"
      TabPicture(1)   =   "mt2000.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "SSPanel5(1)"
      Tab(1).ControlCount=   1
      Begin Threed.SSPanel SSPanel5 
         Height          =   13095
         Index           =   0
         Left            =   120
         TabIndex        =   131
         Top             =   480
         Width           =   5535
         _Version        =   65536
         _ExtentX        =   9763
         _ExtentY        =   23098
         _StockProps     =   15
         ForeColor       =   -2147483630
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelInner      =   1
         Begin VB.TextBox Text7 
            Alignment       =   1  '¿À¸¥ÂÊ ¸ÂÃã
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   12
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3120
            TabIndex        =   366
            Top             =   12000
            Width           =   975
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Clean Tip"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3120
            TabIndex        =   365
            Top             =   11400
            Width           =   975
         End
         Begin VB.CheckBox Check_Yes_Ink 
            Caption         =   "Sample Test (Ink)"
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   364
            TabStop         =   0   'False
            Top             =   8880
            Width           =   1695
         End
         Begin Threed.SSPanel SSPanel_DateTime 
            Height          =   375
            Left            =   2040
            TabIndex        =   162
            Top             =   6120
            Width           =   3255
            _Version        =   65536
            _ExtentX        =   5741
            _ExtentY        =   661
            _StockProps     =   15
            ForeColor       =   0
            BackColor       =   16777215
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
         End
         Begin Threed.SSPanel SSPanel14 
            Height          =   375
            Left            =   120
            TabIndex        =   160
            Top             =   6120
            Width           =   1935
            _Version        =   65536
            _ExtentX        =   3413
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "WORK TIME"
            ForeColor       =   -2147483630
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
         Begin Threed.SSPanel SSPanel_Yield 
            Height          =   375
            Left            =   2040
            TabIndex        =   357
            Top             =   5760
            Visible         =   0   'False
            Width           =   3255
            _Version        =   65536
            _ExtentX        =   5741
            _ExtentY        =   661
            _StockProps     =   15
            ForeColor       =   0
            BackColor       =   16777215
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
         End
         Begin Threed.SSPanel SSPanel4 
            Height          =   375
            Index           =   18
            Left            =   120
            TabIndex        =   356
            Top             =   5760
            Visible         =   0   'False
            Width           =   1935
            _Version        =   65536
            _ExtentX        =   3413
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "YIELD"
            ForeColor       =   -2147483630
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
         Begin VB.CheckBox Check_Crack_wafer 
            Caption         =   "Crack Wafer"
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   1860
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   355
            TabStop         =   0   'False
            Top             =   9600
            Width           =   1695
         End
         Begin VB.CommandButton Command13 
            Caption         =   $"mt2000.frx":0902
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   3600
            TabIndex        =   354
            Top             =   8880
            Width           =   1695
         End
         Begin VB.CommandButton Command12 
            Caption         =   "Chip Fail Set"
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   6300
            TabIndex        =   353
            Top             =   8880
            Width           =   1695
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Ink Start Pos Off"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   4200
            MaskColor       =   &H8000000F&
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   313
            TabStop         =   0   'False
            Top             =   12480
            Width           =   1215
         End
         Begin VB.CommandButton Command_MaskMove 
            Caption         =   "Mask Shift"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   4200
            TabIndex        =   300
            Top             =   11940
            Width           =   1215
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Skip to Ink"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   4200
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   320
            Top             =   11400
            Width           =   1215
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Sampling Off"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   5640
            MaskColor       =   &H8000000F&
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   311
            TabStop         =   0   'False
            Top             =   11880
            Width           =   1215
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00C0FFFF&
            Caption         =   " Wafer Direction "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   240
            TabIndex        =   305
            Top             =   11520
            Width           =   2775
            Begin VB.OptionButton Option_WaferDirection 
               BackColor       =   &H00C0FFFF&
               Caption         =   "270"
               Height          =   255
               Index           =   3
               Left            =   2040
               TabIndex        =   310
               TabStop         =   0   'False
               Top             =   360
               Width           =   615
            End
            Begin VB.OptionButton Option_WaferDirection 
               BackColor       =   &H00C0FFFF&
               Caption         =   "-90"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   309
               TabStop         =   0   'False
               Top             =   360
               Width           =   615
            End
            Begin VB.OptionButton Option_WaferDirection 
               BackColor       =   &H00C0FFFF&
               Caption         =   "90"
               Height          =   255
               Index           =   1
               Left            =   720
               TabIndex        =   308
               TabStop         =   0   'False
               Top             =   360
               Width           =   615
            End
            Begin VB.OptionButton Option_WaferDirection 
               BackColor       =   &H00C0FFFF&
               Caption         =   "180"
               Height          =   255
               Index           =   2
               Left            =   1320
               TabIndex        =   307
               TabStop         =   0   'False
               Top             =   360
               Width           =   615
            End
            Begin VB.CommandButton Command_WaferDirection 
               Caption         =   "Wafer Direction"
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
               Left            =   120
               TabIndex        =   306
               TabStop         =   0   'False
               Top             =   0
               Width           =   2535
            End
         End
         Begin VB.CommandButton Command_WaferDraw 
            Caption         =   $"mt2000.frx":0915
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
            Height          =   495
            Left            =   120
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   304
            TabStop         =   0   'False
            Top             =   10800
            Width           =   1215
         End
         Begin VB.CommandButton Command_Map_Clear 
            Caption         =   $"mt2000.frx":0924
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
            Height          =   495
            Left            =   3120
            TabIndex        =   303
            TabStop         =   0   'False
            Top             =   10800
            Width           =   1095
         End
         Begin VB.CommandButton Command_DisplayWafer 
            Caption         =   $"mt2000.frx":0932
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
            Height          =   495
            Left            =   1680
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   302
            TabStop         =   0   'False
            Top             =   10800
            Width           =   1095
         End
         Begin VB.CommandButton Command_ImageSave 
            BackColor       =   &H00C0C0FF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   4440
            Picture         =   "mt2000.frx":0944
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   301
            TabStop         =   0   'False
            Top             =   10800
            Width           =   915
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Frame2"
            Height          =   615
            Left            =   240
            TabIndex        =   294
            Top             =   12360
            Width           =   2775
            Begin VB.OptionButton Option3 
               BackColor       =   &H00C0FFFF&
               Caption         =   "X4"
               Height          =   255
               Index           =   3
               Left            =   2040
               TabIndex        =   299
               Top             =   300
               Width           =   615
            End
            Begin VB.OptionButton Option3 
               BackColor       =   &H00C0FFFF&
               Caption         =   "X3"
               Height          =   255
               Index           =   2
               Left            =   1440
               TabIndex        =   298
               Top             =   300
               Value           =   -1  'True
               Width           =   615
            End
            Begin VB.OptionButton Option3 
               BackColor       =   &H00C0FFFF&
               Caption         =   "X2"
               Height          =   255
               Index           =   1
               Left            =   840
               TabIndex        =   297
               Top             =   300
               Width           =   735
            End
            Begin VB.OptionButton Option3 
               BackColor       =   &H00C0FFFF&
               Caption         =   "X1"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   296
               Top             =   300
               Width           =   735
            End
            Begin VB.CommandButton Command9 
               Caption         =   "Size Change"
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
               Left            =   120
               TabIndex        =   295
               TabStop         =   0   'False
               Top             =   0
               Width           =   2535
            End
         End
         Begin VB.CommandButton Command_MapMove 
            Caption         =   "Map Shift"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   6360
            TabIndex        =   293
            Top             =   11640
            Width           =   1215
         End
         Begin VB.CheckBox Check_No_Probe 
            Caption         =   "Only Edge Ink"
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   1860
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   291
            TabStop         =   0   'False
            Top             =   8880
            Width           =   1695
         End
         Begin VB.CheckBox Check_No_Ink 
            Caption         =   "Sample Test (No Ink)"
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   290
            TabStop         =   0   'False
            Top             =   9600
            Width           =   1695
         End
         Begin VB.CommandButton Command7 
            Caption         =   "Ink Test"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   3600
            TabIndex        =   289
            Top             =   9600
            Width           =   1695
         End
         Begin VB.TextBox Text1 
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
            Left            =   2040
            TabIndex        =   173
            Top             =   7920
            Width           =   3255
         End
         Begin Threed.SSPanel SSPanel4 
            Height          =   375
            Index           =   6
            Left            =   120
            TabIndex        =   174
            Top             =   7920
            Width           =   1935
            _Version        =   65536
            _ExtentX        =   3413
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "EQUIPMENT ID"
            ForeColor       =   -2147483630
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
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
         Begin VB.TextBox Text1 
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
            Left            =   2040
            TabIndex        =   171
            Top             =   7200
            Width           =   1935
         End
         Begin VB.TextBox Text1 
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
            Left            =   2040
            TabIndex        =   169
            Top             =   7560
            Width           =   3255
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Cassette"
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   11.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3960
            TabIndex        =   168
            Top             =   7200
            Width           =   1335
         End
         Begin VB.TextBox Text6 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            BackColor       =   &H00E0E0E0&
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
            Left            =   3960
            TabIndex        =   154
            Text            =   "0"
            Top             =   3720
            Width           =   1335
         End
         Begin VB.TextBox Text5 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            BackColor       =   &H00E0E0E0&
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
            Left            =   2040
            TabIndex        =   148
            Text            =   "0"
            Top             =   3720
            Width           =   1335
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H0080FFFF&
            Caption         =   "TEST"
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   0
            Left            =   -1200
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   132
            TabStop         =   0   'False
            Top             =   5880
            Visible         =   0   'False
            Width           =   1095
         End
         Begin Threed.SSPanel SSPanel2 
            Height          =   375
            Index           =   9
            Left            =   1680
            TabIndex        =   133
            Top             =   2400
            Width           =   3615
            _Version        =   65536
            _ExtentX        =   6376
            _ExtentY        =   661
            _StockProps     =   15
            ForeColor       =   4210752
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
         End
         Begin Threed.SSPanel SSPanel4 
            Height          =   375
            Index           =   22
            Left            =   120
            TabIndex        =   134
            Top             =   2400
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "RESULT TIME"
            ForeColor       =   -2147483630
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
         Begin Threed.SSPanel SSPanel2 
            Height          =   375
            Index           =   4
            Left            =   1680
            TabIndex        =   135
            Top             =   2040
            Width           =   3615
            _Version        =   65536
            _ExtentX        =   6376
            _ExtentY        =   661
            _StockProps     =   15
            ForeColor       =   4210752
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
         End
         Begin Threed.SSPanel SSPanel4 
            Height          =   375
            Index           =   38
            Left            =   120
            TabIndex        =   136
            Top             =   2040
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "RESULT"
            ForeColor       =   -2147483630
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
         Begin Threed.SSPanel SSPanel2 
            Height          =   375
            Index           =   3
            Left            =   1680
            TabIndex        =   137
            Top             =   1680
            Width           =   3615
            _Version        =   65536
            _ExtentX        =   6376
            _ExtentY        =   661
            _StockProps     =   15
            ForeColor       =   4210752
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
         End
         Begin Threed.SSPanel SSPanel4 
            Height          =   375
            Index           =   40
            Left            =   120
            TabIndex        =   138
            Top             =   1680
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "BIN"
            ForeColor       =   -2147483630
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
         Begin Threed.SSPanel SSPanel2 
            Height          =   375
            Index           =   2
            Left            =   1680
            TabIndex        =   139
            Top             =   1320
            Width           =   3615
            _Version        =   65536
            _ExtentX        =   6376
            _ExtentY        =   661
            _StockProps     =   15
            ForeColor       =   4210752
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
         End
         Begin Threed.SSPanel SSPanel4 
            Height          =   375
            Index           =   42
            Left            =   120
            TabIndex        =   140
            Top             =   1320
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "OUTPUT"
            ForeColor       =   -2147483630
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
         Begin Threed.SSPanel SSPanel2 
            Height          =   375
            Index           =   1
            Left            =   1680
            TabIndex        =   141
            Top             =   960
            Width           =   3615
            _Version        =   65536
            _ExtentX        =   6376
            _ExtentY        =   661
            _StockProps     =   15
            ForeColor       =   4210752
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
         End
         Begin Threed.SSPanel NO_CHEAT 
            Height          =   375
            Left            =   120
            TabIndex        =   142
            Top             =   960
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "COUNT"
            ForeColor       =   -2147483630
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
         Begin Threed.SSPanel SSPanel2 
            Height          =   375
            Index           =   0
            Left            =   1680
            TabIndex        =   143
            Top             =   600
            Width           =   3615
            _Version        =   65536
            _ExtentX        =   6376
            _ExtentY        =   661
            _StockProps     =   15
            ForeColor       =   4210752
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
         End
         Begin Threed.SSPanel Cheat_Debug 
            Height          =   375
            Left            =   120
            TabIndex        =   144
            Top             =   600
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "MAP"
            ForeColor       =   -2147483630
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
         Begin Threed.SSPanel SSPanel25 
            Height          =   375
            Left            =   120
            TabIndex        =   145
            Top             =   120
            Width           =   5175
            _Version        =   65536
            _ExtentX        =   9128
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "WORK INFORMATION"
            ForeColor       =   12648447
            BackColor       =   4210752
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±¼¸²"
               Size            =   11.99
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FloodColor      =   0
         End
         Begin Threed.SSPanel SSPanel6 
            Height          =   375
            Left            =   120
            TabIndex        =   146
            Top             =   3720
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   " X,Y Pos"
            ForeColor       =   -2147483630
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
         Begin Threed.SSPanel SSPanel_WaferSize 
            Height          =   375
            Left            =   1680
            TabIndex        =   149
            Top             =   3360
            Width           =   3615
            _Version        =   65536
            _ExtentX        =   6376
            _ExtentY        =   661
            _StockProps     =   15
            ForeColor       =   0
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
         End
         Begin Threed.SSPanel SSPanel3 
            Height          =   375
            Left            =   120
            TabIndex        =   150
            Top             =   3360
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   " Wafer Size"
            ForeColor       =   -2147483630
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
         Begin Threed.SSPanel SSPanel_ChipYSize 
            Height          =   375
            Left            =   3960
            TabIndex        =   151
            Top             =   4080
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   661
            _StockProps     =   15
            ForeColor       =   0
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
         End
         Begin Threed.SSPanel SSPanel_ChipXSize 
            Height          =   375
            Left            =   2040
            TabIndex        =   152
            Top             =   4080
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   661
            _StockProps     =   15
            ForeColor       =   0
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
         End
         Begin Threed.SSPanel SSPanel4 
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   153
            Top             =   4080
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   " Chip Size"
            ForeColor       =   -2147483630
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
         Begin Threed.SSPanel SSPanel15 
            Height          =   375
            Left            =   120
            TabIndex        =   158
            Top             =   5040
            Width           =   1935
            _Version        =   65536
            _ExtentX        =   3413
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "TOTAL COUNT"
            ForeColor       =   -2147483630
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
            Index           =   24
            Left            =   120
            TabIndex        =   159
            Top             =   5400
            Width           =   1935
            _Version        =   65536
            _ExtentX        =   3413
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "GOOD/BAD/SKIP"
            ForeColor       =   -2147483630
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
         Begin Threed.SSPanel SSPanel_TotalCount 
            Height          =   375
            Left            =   2040
            TabIndex        =   161
            Top             =   5040
            Width           =   3255
            _Version        =   65536
            _ExtentX        =   5741
            _ExtentY        =   661
            _StockProps     =   15
            ForeColor       =   0
            BackColor       =   16777215
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
         End
         Begin Threed.SSPanel SSPanel_SkipCount 
            Height          =   375
            Left            =   4200
            TabIndex        =   163
            Top             =   5400
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   661
            _StockProps     =   15
            ForeColor       =   0
            BackColor       =   16777215
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
         End
         Begin Threed.SSPanel SSPanel_BadCount 
            Height          =   375
            Left            =   3120
            TabIndex        =   164
            Top             =   5400
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   661
            _StockProps     =   15
            ForeColor       =   0
            BackColor       =   16777215
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
         End
         Begin Threed.SSPanel SSPanel_GoodCount 
            Height          =   375
            Left            =   2040
            TabIndex        =   165
            Top             =   5400
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   661
            _StockProps     =   15
            ForeColor       =   0
            BackColor       =   16777215
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
         End
         Begin Threed.SSPanel SSPanel10 
            Height          =   375
            Left            =   120
            TabIndex        =   166
            Top             =   2880
            Width           =   5175
            _Version        =   65536
            _ExtentX        =   9128
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "PACKAGE INFORMATION"
            ForeColor       =   12648447
            BackColor       =   4210752
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "@±¼¸²"
               Size            =   11.99
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FloodColor      =   0
         End
         Begin Threed.SSPanel SSPanel11 
            Height          =   375
            Left            =   120
            TabIndex        =   167
            Top             =   4560
            Width           =   5175
            _Version        =   65536
            _ExtentX        =   9128
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "COUNT INFORMATION"
            ForeColor       =   12648447
            BackColor       =   4210752
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "@±¼¸²"
               Size            =   11.99
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FloodColor      =   0
         End
         Begin Threed.SSPanel SSPanel8 
            Height          =   375
            Left            =   120
            TabIndex        =   170
            Top             =   7560
            Width           =   1935
            _Version        =   65536
            _ExtentX        =   3413
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "DEVICE NAME"
            ForeColor       =   -2147483630
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
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
            Left            =   120
            TabIndex        =   172
            Top             =   7200
            Width           =   1935
            _Version        =   65536
            _ExtentX        =   3413
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "LOT NO"
            ForeColor       =   -2147483630
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
         Begin Threed.SSPanel SSPanel12 
            Height          =   375
            Left            =   120
            TabIndex        =   175
            Top             =   6720
            Width           =   5175
            _Version        =   65536
            _ExtentX        =   9128
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "DATA TEACH"
            ForeColor       =   12648447
            BackColor       =   4210752
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "@±¼¸²"
               Size            =   11.99
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FloodColor      =   0
         End
         Begin Threed.SSPanel SSPanel18 
            Height          =   375
            Left            =   120
            TabIndex        =   288
            Top             =   8400
            Width           =   5175
            _Version        =   65536
            _ExtentX        =   9128
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "PROBE FUNCTION"
            ForeColor       =   12648447
            BackColor       =   4210752
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "@±¼¸²"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FloodColor      =   0
         End
         Begin Threed.SSPanel SSPanel19 
            Height          =   375
            Left            =   120
            TabIndex        =   292
            Top             =   10320
            Width           =   5295
            _Version        =   65536
            _ExtentX        =   9340
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "MAP FUNCTION"
            ForeColor       =   12648447
            BackColor       =   4210752
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "@±¼¸²"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FloodColor      =   8454143
         End
         Begin VB.Label Label2 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            BackColor       =   &H00C0FFFF&
            Caption         =   "Y"
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   11.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   3600
            TabIndex        =   157
            Top             =   3840
            Width           =   255
         End
         Begin VB.Label Label2 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            BackColor       =   &H00C0FFFF&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   11.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   1680
            TabIndex        =   156
            Top             =   4200
            Width           =   255
         End
         Begin VB.Label Label2 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            BackColor       =   &H00C0FFFF&
            Caption         =   "Y"
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   11.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   3600
            TabIndex        =   155
            Top             =   4200
            Width           =   255
         End
         Begin VB.Label Label2 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            BackColor       =   &H00C0FFFF&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   11.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   1680
            TabIndex        =   147
            Top             =   3840
            Width           =   255
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   9255
         Index           =   1
         Left            =   -74880
         TabIndex        =   176
         Top             =   480
         Width           =   5655
         _Version        =   65536
         _ExtentX        =   9975
         _ExtentY        =   16325
         _StockProps     =   15
         ForeColor       =   -2147483630
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelInner      =   1
         Begin VB.CommandButton Command_BINClear 
            Caption         =   "Clear"
            Height          =   375
            Index           =   24
            Left            =   720
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   352
            Top             =   8760
            Width           =   615
         End
         Begin VB.CommandButton Command_BINClear 
            Caption         =   "Clear"
            Height          =   375
            Index           =   23
            Left            =   720
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   351
            Top             =   8400
            Width           =   615
         End
         Begin VB.CommandButton Command_BINClear 
            Caption         =   "Clear"
            Height          =   375
            Index           =   22
            Left            =   720
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   350
            Top             =   8040
            Width           =   615
         End
         Begin VB.CommandButton Command_BINClear 
            Caption         =   "Clear"
            Height          =   375
            Index           =   21
            Left            =   720
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   349
            Top             =   7680
            Width           =   615
         End
         Begin VB.CommandButton Command_BINClear 
            Caption         =   "Clear"
            Height          =   375
            Index           =   20
            Left            =   720
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   348
            Top             =   7320
            Width           =   615
         End
         Begin VB.CommandButton Command_BINClear 
            Caption         =   "Clear"
            Height          =   375
            Index           =   19
            Left            =   720
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   347
            Top             =   6960
            Width           =   615
         End
         Begin VB.CommandButton Command_BINClear 
            Caption         =   "Clear"
            Height          =   375
            Index           =   18
            Left            =   720
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   346
            Top             =   6600
            Width           =   615
         End
         Begin VB.CommandButton Command_BINClear 
            Caption         =   "Clear"
            Height          =   375
            Index           =   17
            Left            =   720
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   345
            Top             =   6240
            Width           =   615
         End
         Begin VB.CommandButton Command_BINClear 
            Caption         =   "Clear"
            Height          =   375
            Index           =   16
            Left            =   720
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   344
            Top             =   5880
            Width           =   615
         End
         Begin VB.CommandButton Command_BINClear 
            Caption         =   "Clear"
            Height          =   375
            Index           =   15
            Left            =   720
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   343
            Top             =   5520
            Width           =   615
         End
         Begin VB.CommandButton Command_BINClear 
            Caption         =   "Clear"
            Height          =   375
            Index           =   14
            Left            =   720
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   342
            Top             =   5160
            Width           =   615
         End
         Begin VB.CommandButton Command_BINClear 
            Caption         =   "Clear"
            Height          =   375
            Index           =   13
            Left            =   720
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   341
            Top             =   4800
            Width           =   615
         End
         Begin VB.CommandButton Command_BINClear 
            Caption         =   "Clear"
            Height          =   375
            Index           =   12
            Left            =   720
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   340
            Top             =   4440
            Width           =   615
         End
         Begin VB.CommandButton Command_BINClear 
            Caption         =   "Clear"
            Height          =   375
            Index           =   11
            Left            =   720
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   339
            Top             =   4080
            Width           =   615
         End
         Begin VB.CommandButton Command_BINClear 
            Caption         =   "Clear"
            Height          =   375
            Index           =   10
            Left            =   720
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   338
            Top             =   3720
            Width           =   615
         End
         Begin VB.CommandButton Command_BINClear 
            Caption         =   "Clear"
            Height          =   375
            Index           =   9
            Left            =   720
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   337
            Top             =   3360
            Width           =   615
         End
         Begin VB.CommandButton Command_BINClear 
            Caption         =   "Clear"
            Height          =   375
            Index           =   8
            Left            =   720
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   336
            Top             =   3000
            Width           =   615
         End
         Begin VB.CommandButton Command_BINClear 
            Caption         =   "Clear"
            Height          =   375
            Index           =   7
            Left            =   720
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   335
            Top             =   2640
            Width           =   615
         End
         Begin VB.CommandButton Command_BINClear 
            Caption         =   "Clear"
            Height          =   375
            Index           =   6
            Left            =   720
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   334
            Top             =   2280
            Width           =   615
         End
         Begin VB.CommandButton Command_BINClear 
            Caption         =   "Clear"
            Height          =   375
            Index           =   5
            Left            =   720
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   333
            Top             =   1920
            Width           =   615
         End
         Begin VB.CommandButton Command_BINClear 
            Caption         =   "Clear"
            Height          =   375
            Index           =   4
            Left            =   720
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   332
            Top             =   1560
            Width           =   615
         End
         Begin VB.CommandButton Command_BINClear 
            Caption         =   "Clear"
            Height          =   375
            Index           =   3
            Left            =   720
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   331
            Top             =   1200
            Width           =   615
         End
         Begin VB.CommandButton Command_BINClear 
            Caption         =   "Clear"
            Height          =   375
            Index           =   2
            Left            =   720
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   330
            Top             =   840
            Width           =   615
         End
         Begin VB.CommandButton Command_BINClear 
            Caption         =   "Clear"
            Height          =   375
            Index           =   1
            Left            =   720
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   329
            Top             =   480
            Width           =   615
         End
         Begin VB.CommandButton Command_BINClear 
            Caption         =   "Clear"
            Height          =   375
            Index           =   0
            Left            =   720
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   328
            Top             =   120
            Width           =   615
         End
         Begin VB.CommandButton Command_ChipColor 
            Caption         =   "Ink2"
            Height          =   375
            Index           =   6
            Left            =   4080
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   321
            Top             =   3000
            Width           =   1335
         End
         Begin Threed.SSPanel SSPanel13 
            Height          =   855
            Left            =   3960
            TabIndex        =   265
            Top             =   5760
            Width           =   1575
            _Version        =   65536
            _ExtentX        =   2778
            _ExtentY        =   1508
            _StockProps     =   15
            BackColor       =   12648384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "µ¸¿ò"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin VB.TextBox Text_SkipCount 
               Height          =   375
               Left            =   120
               TabIndex        =   267
               Top             =   360
               Width           =   1335
            End
            Begin VB.Label Label10 
               BackColor       =   &H00C0FFC0&
               Caption         =   "UGLY DIE"
               Height          =   255
               Left            =   120
               TabIndex        =   266
               Top             =   120
               Width           =   975
            End
         End
         Begin VB.TextBox Text_BadCount 
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4080
            TabIndex        =   264
            Top             =   5160
            Width           =   1335
         End
         Begin VB.TextBox Text_GoodCount 
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4080
            TabIndex        =   262
            Top             =   4560
            Width           =   1335
         End
         Begin VB.TextBox Text_TotalCount 
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4080
            TabIndex        =   260
            Top             =   3960
            Width           =   1335
         End
         Begin VB.CommandButton Command_ChipColor 
            Caption         =   "Normal"
            Height          =   375
            Index           =   0
            Left            =   4080
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   258
            Top             =   120
            Width           =   1335
         End
         Begin VB.CommandButton Command_ChipColor 
            Caption         =   "Pattern"
            Height          =   375
            Index           =   1
            Left            =   4080
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   257
            Top             =   600
            Width           =   1335
         End
         Begin VB.CommandButton Command_ChipColor 
            Caption         =   "Measure"
            Height          =   375
            Index           =   2
            Left            =   4080
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   256
            Top             =   1080
            Width           =   1335
         End
         Begin VB.CommandButton Command_ChipColor 
            Caption         =   "Skip Die"
            Height          =   375
            Index           =   3
            Left            =   4080
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   255
            Top             =   1560
            Width           =   1335
         End
         Begin VB.CommandButton Command_ChipColor 
            Caption         =   "Plate Zone"
            Height          =   375
            Index           =   4
            Left            =   4080
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   254
            Top             =   2040
            Width           =   1335
         End
         Begin VB.CommandButton Command_ChipColor 
            Caption         =   "Ink1"
            Height          =   375
            Index           =   5
            Left            =   4080
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   253
            Top             =   2520
            Width           =   1335
         End
         Begin VB.TextBox Text_BinCount 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   24
            Left            =   2760
            TabIndex        =   252
            Top             =   8760
            Width           =   1095
         End
         Begin VB.TextBox Text_BinCount 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   23
            Left            =   2760
            TabIndex        =   251
            Top             =   8400
            Width           =   1095
         End
         Begin VB.TextBox Text_BinCount 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   22
            Left            =   2760
            TabIndex        =   250
            Top             =   8040
            Width           =   1095
         End
         Begin VB.TextBox Text_BinCount 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   21
            Left            =   2760
            TabIndex        =   249
            Top             =   7680
            Width           =   1095
         End
         Begin VB.TextBox Text_BinCount 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   20
            Left            =   2760
            TabIndex        =   248
            Top             =   7320
            Width           =   1095
         End
         Begin VB.TextBox Text_BinCount 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   19
            Left            =   2760
            TabIndex        =   247
            Top             =   6960
            Width           =   1095
         End
         Begin VB.TextBox Text_BinCount 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   18
            Left            =   2760
            TabIndex        =   246
            Top             =   6600
            Width           =   1095
         End
         Begin VB.TextBox Text_BinCount 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   17
            Left            =   2760
            TabIndex        =   245
            Top             =   6240
            Width           =   1095
         End
         Begin VB.TextBox Text_BinCount 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   16
            Left            =   2760
            TabIndex        =   244
            Top             =   5880
            Width           =   1095
         End
         Begin VB.TextBox Text_BinCount 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   15
            Left            =   2760
            TabIndex        =   243
            Top             =   5520
            Width           =   1095
         End
         Begin VB.TextBox Text_BinCount 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   14
            Left            =   2760
            TabIndex        =   242
            Top             =   5160
            Width           =   1095
         End
         Begin VB.TextBox Text_BinCount 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   13
            Left            =   2760
            TabIndex        =   241
            Top             =   4800
            Width           =   1095
         End
         Begin VB.TextBox Text_BinCount 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   12
            Left            =   2760
            TabIndex        =   240
            Top             =   4440
            Width           =   1095
         End
         Begin VB.TextBox Text_BinCount 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   11
            Left            =   2760
            TabIndex        =   239
            Top             =   4080
            Width           =   1095
         End
         Begin VB.TextBox Text_BinCount 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   10
            Left            =   2760
            TabIndex        =   238
            Top             =   3720
            Width           =   1095
         End
         Begin VB.TextBox Text_BinCount 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   9
            Left            =   2760
            TabIndex        =   237
            Top             =   3360
            Width           =   1095
         End
         Begin VB.TextBox Text_BinCount 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   8
            Left            =   2760
            TabIndex        =   236
            Top             =   3000
            Width           =   1095
         End
         Begin VB.TextBox Text_BinCount 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   7
            Left            =   2760
            TabIndex        =   235
            Top             =   2640
            Width           =   1095
         End
         Begin VB.TextBox Text_BinCount 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   6
            Left            =   2760
            TabIndex        =   234
            Top             =   2280
            Width           =   1095
         End
         Begin VB.TextBox Text_BinCount 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   5
            Left            =   2760
            TabIndex        =   233
            Top             =   1920
            Width           =   1095
         End
         Begin VB.TextBox Text_BinCount 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   4
            Left            =   2760
            TabIndex        =   232
            Top             =   1560
            Width           =   1095
         End
         Begin VB.TextBox Text_BinCount 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   2760
            TabIndex        =   231
            Top             =   1200
            Width           =   1095
         End
         Begin VB.TextBox Text_BinCount 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   2760
            TabIndex        =   230
            Top             =   840
            Width           =   1095
         End
         Begin VB.TextBox Text_BinCount 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   2760
            TabIndex        =   229
            Top             =   480
            Width           =   1095
         End
         Begin VB.TextBox Text_BinCount 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   2760
            TabIndex        =   228
            Top             =   120
            Width           =   1095
         End
         Begin VB.TextBox Text_BinCommand 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   24
            Left            =   1440
            TabIndex        =   227
            Top             =   8760
            Width           =   1215
         End
         Begin VB.TextBox Text_BinCommand 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   23
            Left            =   1440
            TabIndex        =   226
            Top             =   8400
            Width           =   1215
         End
         Begin VB.TextBox Text_BinCommand 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   22
            Left            =   1440
            TabIndex        =   225
            Top             =   8040
            Width           =   1215
         End
         Begin VB.TextBox Text_BinCommand 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   21
            Left            =   1440
            TabIndex        =   224
            Top             =   7680
            Width           =   1215
         End
         Begin VB.TextBox Text_BinCommand 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   20
            Left            =   1440
            TabIndex        =   223
            Top             =   7320
            Width           =   1215
         End
         Begin VB.TextBox Text_BinCommand 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   19
            Left            =   1440
            TabIndex        =   222
            Top             =   6960
            Width           =   1215
         End
         Begin VB.TextBox Text_BinCommand 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   18
            Left            =   1440
            TabIndex        =   221
            Top             =   6600
            Width           =   1215
         End
         Begin VB.TextBox Text_BinCommand 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   17
            Left            =   1440
            TabIndex        =   220
            Top             =   6240
            Width           =   1215
         End
         Begin VB.TextBox Text_BinCommand 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   16
            Left            =   1440
            TabIndex        =   219
            Top             =   5880
            Width           =   1215
         End
         Begin VB.TextBox Text_BinCommand 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   15
            Left            =   1440
            TabIndex        =   218
            Top             =   5520
            Width           =   1215
         End
         Begin VB.TextBox Text_BinCommand 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   14
            Left            =   1440
            TabIndex        =   217
            Top             =   5160
            Width           =   1215
         End
         Begin VB.TextBox Text_BinCommand 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   13
            Left            =   1440
            TabIndex        =   216
            Top             =   4800
            Width           =   1215
         End
         Begin VB.TextBox Text_BinCommand 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   12
            Left            =   1440
            TabIndex        =   215
            Top             =   4440
            Width           =   1215
         End
         Begin VB.TextBox Text_BinCommand 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   11
            Left            =   1440
            TabIndex        =   214
            Top             =   4080
            Width           =   1215
         End
         Begin VB.TextBox Text_BinCommand 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   10
            Left            =   1440
            TabIndex        =   213
            Top             =   3720
            Width           =   1215
         End
         Begin VB.TextBox Text_BinCommand 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   9
            Left            =   1440
            TabIndex        =   212
            Top             =   3360
            Width           =   1215
         End
         Begin VB.TextBox Text_BinCommand 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   8
            Left            =   1440
            TabIndex        =   211
            Top             =   3000
            Width           =   1215
         End
         Begin VB.TextBox Text_BinCommand 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   7
            Left            =   1440
            TabIndex        =   210
            Top             =   2640
            Width           =   1215
         End
         Begin VB.TextBox Text_BinCommand 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   6
            Left            =   1440
            TabIndex        =   209
            Top             =   2280
            Width           =   1215
         End
         Begin VB.TextBox Text_BinCommand 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   5
            Left            =   1440
            TabIndex        =   208
            Top             =   1920
            Width           =   1215
         End
         Begin VB.TextBox Text_BinCommand 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   4
            Left            =   1440
            TabIndex        =   207
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox Text_BinCommand 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   1440
            TabIndex        =   206
            Top             =   1200
            Width           =   1215
         End
         Begin VB.TextBox Text_BinCommand 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   1440
            TabIndex        =   205
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox Text_BinCommand 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   1440
            TabIndex        =   204
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox Text_BinCommand 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            BeginProperty Font 
               Name            =   "µ¸¿ò"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   1440
            TabIndex        =   203
            Top             =   120
            Width           =   1215
         End
         Begin VB.CommandButton Command_BINColor 
            Caption         =   "BIN24"
            Height          =   375
            Index           =   24
            Left            =   120
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   202
            Top             =   8760
            Width           =   615
         End
         Begin VB.CommandButton Command_BINColor 
            Caption         =   "BIN23"
            Height          =   375
            Index           =   23
            Left            =   120
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   201
            Top             =   8400
            Width           =   615
         End
         Begin VB.CommandButton Command_BINColor 
            Caption         =   "BIN22"
            Height          =   375
            Index           =   22
            Left            =   120
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   200
            Top             =   8040
            Width           =   615
         End
         Begin VB.CommandButton Command_BINColor 
            Caption         =   "BIN21"
            Height          =   375
            Index           =   21
            Left            =   120
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   199
            Top             =   7680
            Width           =   615
         End
         Begin VB.CommandButton Command_BINColor 
            Caption         =   "BIN20"
            Height          =   375
            Index           =   20
            Left            =   120
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   198
            Top             =   7320
            Width           =   615
         End
         Begin VB.CommandButton Command_BINColor 
            Caption         =   "BIN19"
            Height          =   375
            Index           =   19
            Left            =   120
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   197
            Top             =   6960
            Width           =   615
         End
         Begin VB.CommandButton Command_BINColor 
            Caption         =   "BIN18"
            Height          =   375
            Index           =   18
            Left            =   120
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   196
            Top             =   6600
            Width           =   615
         End
         Begin VB.CommandButton Command_BINColor 
            Caption         =   "BIN17"
            Height          =   375
            Index           =   17
            Left            =   120
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   195
            Top             =   6240
            Width           =   615
         End
         Begin VB.CommandButton Command_BINColor 
            Caption         =   "BIN16"
            Height          =   375
            Index           =   16
            Left            =   120
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   194
            Top             =   5880
            Width           =   615
         End
         Begin VB.CommandButton Command_BINColor 
            Caption         =   "BIN15"
            Height          =   375
            Index           =   15
            Left            =   120
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   193
            Top             =   5520
            Width           =   615
         End
         Begin VB.CommandButton Command_BINColor 
            Caption         =   "BIN14"
            Height          =   375
            Index           =   14
            Left            =   120
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   192
            Top             =   5160
            Width           =   615
         End
         Begin VB.CommandButton Command_BINColor 
            Caption         =   "BIN13"
            Height          =   375
            Index           =   13
            Left            =   120
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   191
            Top             =   4800
            Width           =   615
         End
         Begin VB.CommandButton Command_BINColor 
            Caption         =   "BIN12"
            Height          =   375
            Index           =   12
            Left            =   120
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   190
            Top             =   4440
            Width           =   615
         End
         Begin VB.CommandButton Command_BINColor 
            Caption         =   "BIN11"
            Height          =   375
            Index           =   11
            Left            =   120
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   189
            Top             =   4080
            Width           =   615
         End
         Begin VB.CommandButton Command_BINColor 
            Caption         =   "BIN10"
            Height          =   375
            Index           =   10
            Left            =   120
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   188
            Top             =   3720
            Width           =   615
         End
         Begin VB.CommandButton Command_BINColor 
            Caption         =   "BIN9"
            Height          =   375
            Index           =   9
            Left            =   120
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   187
            Top             =   3360
            Width           =   615
         End
         Begin VB.CommandButton Command_BINColor 
            Caption         =   "BIN8"
            Height          =   375
            Index           =   8
            Left            =   120
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   186
            Top             =   3000
            Width           =   615
         End
         Begin VB.CommandButton Command_BINColor 
            Caption         =   "BIN7"
            Height          =   375
            Index           =   7
            Left            =   120
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   185
            Top             =   2640
            Width           =   615
         End
         Begin VB.CommandButton Command_BINColor 
            Caption         =   "BIN6"
            Height          =   375
            Index           =   6
            Left            =   120
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   184
            Top             =   2280
            Width           =   615
         End
         Begin VB.CommandButton Command_BINColor 
            Caption         =   "BIN5"
            Height          =   375
            Index           =   5
            Left            =   120
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   183
            Top             =   1920
            Width           =   615
         End
         Begin VB.CommandButton Command_BINColor 
            Caption         =   "BIN4"
            Height          =   375
            Index           =   4
            Left            =   120
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   182
            Top             =   1560
            Width           =   615
         End
         Begin VB.CommandButton Command_BINColor 
            Caption         =   "BIN3"
            Height          =   375
            Index           =   3
            Left            =   120
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   181
            Top             =   1200
            Width           =   615
         End
         Begin VB.CommandButton Command_BINColor 
            Caption         =   "BIN2"
            Height          =   375
            Index           =   2
            Left            =   120
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   180
            Top             =   840
            Width           =   615
         End
         Begin VB.CommandButton Command_BINColor 
            Caption         =   "BIN1"
            Height          =   375
            Index           =   1
            Left            =   120
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   179
            Top             =   480
            Width           =   615
         End
         Begin VB.CommandButton Command_BINColor 
            Caption         =   "BIN0"
            Height          =   375
            Index           =   0
            Left            =   120
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   178
            Top             =   120
            Width           =   615
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H0080FFFF&
            Caption         =   "TEST"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   2
            Left            =   -960
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   177
            TabStop         =   0   'False
            Top             =   6000
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H000000FF&
            Height          =   2055
            Left            =   3960
            Top             =   3600
            Width           =   1575
         End
         Begin VB.Label Label9 
            BackColor       =   &H00C0FFFF&
            Caption         =   "BAD COUNT"
            Height          =   255
            Left            =   4080
            TabIndex        =   263
            Top             =   4920
            Width           =   1335
         End
         Begin VB.Label Label8 
            BackColor       =   &H00C0FFFF&
            Caption         =   "GOOD COUNT"
            Height          =   255
            Left            =   4080
            TabIndex        =   261
            Top             =   4320
            Width           =   1335
         End
         Begin VB.Label Label7 
            BackColor       =   &H00C0FFFF&
            Caption         =   "TOTAL COUNT"
            Height          =   255
            Left            =   4080
            TabIndex        =   259
            Top             =   3720
            Width           =   1335
         End
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "RESET"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "±¼¸²Ã¼"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   542
      Index           =   1
      Left            =   1800
      Style           =   1  '±×·¡ÇÈ
      TabIndex        =   129
      TabStop         =   0   'False
      Top             =   13920
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFF80&
      Caption         =   "TEST ID"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   542
      Left            =   3120
      Style           =   1  '±×·¡ÇÈ
      TabIndex        =   128
      Top             =   13920
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "AUTO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   542
      Index           =   3
      Left            =   4560
      MaskColor       =   &H8000000F&
      Style           =   1  '±×·¡ÇÈ
      TabIndex        =   127
      TabStop         =   0   'False
      Top             =   13920
      Width           =   1335
   End
   Begin VB.Frame Frame5 
      Caption         =   " Multi Probe "
      Height          =   5655
      Left            =   13920
      TabIndex        =   16
      Top             =   7680
      Visible         =   0   'False
      Width           =   4935
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   600
         TabIndex        =   125
         Text            =   "Text4"
         Top             =   3480
         Width           =   1935
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Recive Data"
         Height          =   375
         Left            =   3000
         TabIndex        =   103
         Top             =   3480
         Width           =   1215
      End
      Begin VB.TextBox Text_ChipRecive 
         Height          =   285
         Index           =   15
         Left            =   8040
         TabIndex        =   102
         Top             =   5880
         Width           =   1095
      End
      Begin VB.TextBox Text_ChipRecive 
         Height          =   285
         Index           =   14
         Left            =   8040
         TabIndex        =   101
         Top             =   5520
         Width           =   1095
      End
      Begin VB.TextBox Text_ChipRecive 
         Height          =   285
         Index           =   13
         Left            =   8040
         TabIndex        =   100
         Top             =   5160
         Width           =   1095
      End
      Begin VB.TextBox Text_ChipRecive 
         Height          =   285
         Index           =   12
         Left            =   8040
         TabIndex        =   99
         Top             =   4800
         Width           =   1095
      End
      Begin VB.TextBox Text_ChipRecive 
         Height          =   285
         Index           =   11
         Left            =   8040
         TabIndex        =   98
         Top             =   4440
         Width           =   1095
      End
      Begin VB.TextBox Text_ChipRecive 
         Height          =   285
         Index           =   10
         Left            =   8040
         TabIndex        =   97
         Top             =   4080
         Width           =   1095
      End
      Begin VB.TextBox Text_ChipRecive 
         Height          =   285
         Index           =   9
         Left            =   8040
         TabIndex        =   96
         Top             =   3720
         Width           =   1095
      End
      Begin VB.TextBox Text_ChipRecive 
         Height          =   285
         Index           =   8
         Left            =   8040
         TabIndex        =   95
         Top             =   3360
         Width           =   1095
      End
      Begin VB.TextBox Text_ChipRecive 
         Height          =   285
         Index           =   7
         Left            =   8040
         TabIndex        =   94
         Top             =   3000
         Width           =   1095
      End
      Begin VB.TextBox Text_ChipRecive 
         Height          =   285
         Index           =   6
         Left            =   8040
         TabIndex        =   93
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox Text_ChipRecive 
         Height          =   285
         Index           =   5
         Left            =   8040
         TabIndex        =   92
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox Text_ChipRecive 
         Height          =   285
         Index           =   4
         Left            =   8040
         TabIndex        =   91
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox Text_ChipRecive 
         Height          =   285
         Index           =   3
         Left            =   3000
         TabIndex        =   90
         Top             =   4080
         Width           =   1095
      End
      Begin VB.TextBox Text_ChipRecive 
         Height          =   285
         Index           =   2
         Left            =   3000
         TabIndex        =   89
         Top             =   4440
         Width           =   1095
      End
      Begin VB.TextBox Text_ChipRecive 
         Height          =   285
         Index           =   1
         Left            =   3000
         TabIndex        =   88
         Top             =   4800
         Width           =   1095
      End
      Begin VB.TextBox Text_ChipRecive 
         Height          =   285
         Index           =   0
         Left            =   3000
         TabIndex        =   87
         Top             =   5160
         Width           =   1095
      End
      Begin VB.TextBox Text_ReciveData 
         Height          =   375
         Left            =   600
         TabIndex        =   86
         Top             =   3000
         Width           =   3615
      End
      Begin VB.TextBox Text_ChipTest 
         Height          =   375
         Left            =   600
         TabIndex        =   85
         Top             =   2520
         Width           =   3615
      End
      Begin VB.TextBox Text_Chip 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   15
         Left            =   5640
         TabIndex        =   84
         Text            =   "Chip"
         Top             =   5880
         Width           =   615
      End
      Begin VB.TextBox Text_ChipX 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   15
         Left            =   6240
         TabIndex        =   83
         Text            =   "-100"
         Top             =   5880
         Width           =   495
      End
      Begin VB.TextBox Text_ChipBIN 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   15
         Left            =   7200
         TabIndex        =   82
         Text            =   "0"
         Top             =   5880
         Width           =   615
      End
      Begin VB.TextBox Text_ChipY 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   15
         Left            =   6720
         TabIndex        =   81
         Text            =   "-100"
         Top             =   5880
         Width           =   495
      End
      Begin VB.TextBox Text_Chip 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   14
         Left            =   5640
         TabIndex        =   80
         Text            =   "Chip"
         Top             =   5520
         Width           =   615
      End
      Begin VB.TextBox Text_ChipX 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   14
         Left            =   6240
         TabIndex        =   79
         Text            =   "-100"
         Top             =   5520
         Width           =   495
      End
      Begin VB.TextBox Text_ChipBIN 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   14
         Left            =   7200
         TabIndex        =   78
         Text            =   "0"
         Top             =   5520
         Width           =   615
      End
      Begin VB.TextBox Text_ChipY 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   14
         Left            =   6720
         TabIndex        =   77
         Text            =   "-100"
         Top             =   5520
         Width           =   495
      End
      Begin VB.TextBox Text_Chip 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   13
         Left            =   5640
         TabIndex        =   76
         Text            =   "Chip"
         Top             =   5160
         Width           =   615
      End
      Begin VB.TextBox Text_ChipX 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   13
         Left            =   6240
         TabIndex        =   75
         Text            =   "-100"
         Top             =   5160
         Width           =   495
      End
      Begin VB.TextBox Text_ChipBIN 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   13
         Left            =   7200
         TabIndex        =   74
         Text            =   "0"
         Top             =   5160
         Width           =   615
      End
      Begin VB.TextBox Text_ChipY 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   13
         Left            =   6720
         TabIndex        =   73
         Text            =   "-100"
         Top             =   5160
         Width           =   495
      End
      Begin VB.TextBox Text_Chip 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   12
         Left            =   5640
         TabIndex        =   72
         Text            =   "Chip"
         Top             =   4800
         Width           =   615
      End
      Begin VB.TextBox Text_ChipX 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   12
         Left            =   6240
         TabIndex        =   71
         Text            =   "-100"
         Top             =   4800
         Width           =   495
      End
      Begin VB.TextBox Text_ChipBIN 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   12
         Left            =   7200
         TabIndex        =   70
         Text            =   "0"
         Top             =   4800
         Width           =   615
      End
      Begin VB.TextBox Text_ChipY 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   12
         Left            =   6720
         TabIndex        =   69
         Text            =   "-100"
         Top             =   4800
         Width           =   495
      End
      Begin VB.TextBox Text_Chip 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   11
         Left            =   5640
         TabIndex        =   68
         Text            =   "Chip"
         Top             =   4440
         Width           =   615
      End
      Begin VB.TextBox Text_ChipX 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   11
         Left            =   6240
         TabIndex        =   67
         Text            =   "-100"
         Top             =   4440
         Width           =   495
      End
      Begin VB.TextBox Text_ChipBIN 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   11
         Left            =   7200
         TabIndex        =   66
         Text            =   "0"
         Top             =   4440
         Width           =   615
      End
      Begin VB.TextBox Text_ChipY 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   11
         Left            =   6720
         TabIndex        =   65
         Text            =   "-100"
         Top             =   4440
         Width           =   495
      End
      Begin VB.TextBox Text_Chip 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   10
         Left            =   5640
         TabIndex        =   64
         Text            =   "Chip"
         Top             =   4080
         Width           =   615
      End
      Begin VB.TextBox Text_ChipX 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   10
         Left            =   6240
         TabIndex        =   63
         Text            =   "-100"
         Top             =   4080
         Width           =   495
      End
      Begin VB.TextBox Text_ChipBIN 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   10
         Left            =   7200
         TabIndex        =   62
         Text            =   "0"
         Top             =   4080
         Width           =   615
      End
      Begin VB.TextBox Text_ChipY 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   10
         Left            =   6720
         TabIndex        =   61
         Text            =   "-100"
         Top             =   4080
         Width           =   495
      End
      Begin VB.TextBox Text_Chip 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   9
         Left            =   5640
         TabIndex        =   60
         Text            =   "Chip"
         Top             =   3720
         Width           =   615
      End
      Begin VB.TextBox Text_ChipX 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   9
         Left            =   6240
         TabIndex        =   59
         Text            =   "-100"
         Top             =   3720
         Width           =   495
      End
      Begin VB.TextBox Text_ChipBIN 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   9
         Left            =   7200
         TabIndex        =   58
         Text            =   "0"
         Top             =   3720
         Width           =   615
      End
      Begin VB.TextBox Text_ChipY 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   9
         Left            =   6720
         TabIndex        =   57
         Text            =   "-100"
         Top             =   3720
         Width           =   495
      End
      Begin VB.TextBox Text_Chip 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   8
         Left            =   5640
         TabIndex        =   56
         Text            =   "Chip"
         Top             =   3360
         Width           =   615
      End
      Begin VB.TextBox Text_ChipX 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   8
         Left            =   6240
         TabIndex        =   55
         Text            =   "-100"
         Top             =   3360
         Width           =   495
      End
      Begin VB.TextBox Text_ChipBIN 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   8
         Left            =   7200
         TabIndex        =   54
         Text            =   "0"
         Top             =   3360
         Width           =   615
      End
      Begin VB.TextBox Text_ChipY 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   8
         Left            =   6720
         TabIndex        =   53
         Text            =   "-100"
         Top             =   3360
         Width           =   495
      End
      Begin VB.TextBox Text_Chip 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   7
         Left            =   5640
         TabIndex        =   52
         Text            =   "Chip"
         Top             =   3000
         Width           =   615
      End
      Begin VB.TextBox Text_ChipX 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   7
         Left            =   6240
         TabIndex        =   51
         Text            =   "-100"
         Top             =   3000
         Width           =   495
      End
      Begin VB.TextBox Text_ChipBIN 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   7
         Left            =   7200
         TabIndex        =   50
         Text            =   "0"
         Top             =   3000
         Width           =   615
      End
      Begin VB.TextBox Text_ChipY 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   7
         Left            =   6720
         TabIndex        =   49
         Text            =   "-100"
         Top             =   3000
         Width           =   495
      End
      Begin VB.TextBox Text_Chip 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   6
         Left            =   5640
         TabIndex        =   48
         Text            =   "Chip"
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox Text_ChipX 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   6
         Left            =   6240
         TabIndex        =   47
         Text            =   "-100"
         Top             =   2640
         Width           =   495
      End
      Begin VB.TextBox Text_ChipBIN 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   6
         Left            =   7200
         TabIndex        =   46
         Text            =   "0"
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox Text_ChipY 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   6
         Left            =   6720
         TabIndex        =   45
         Text            =   "-100"
         Top             =   2640
         Width           =   495
      End
      Begin VB.TextBox Text_Chip 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   5
         Left            =   5640
         TabIndex        =   44
         Text            =   "Chip"
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox Text_ChipX 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   5
         Left            =   6240
         TabIndex        =   43
         Text            =   "-100"
         Top             =   2280
         Width           =   495
      End
      Begin VB.TextBox Text_ChipBIN 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   5
         Left            =   7200
         TabIndex        =   42
         Text            =   "0"
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox Text_ChipY 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   5
         Left            =   6720
         TabIndex        =   41
         Text            =   "-100"
         Top             =   2280
         Width           =   495
      End
      Begin VB.TextBox Text_Chip 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   4
         Left            =   5640
         TabIndex        =   40
         Text            =   "Chip"
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox Text_ChipX 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   4
         Left            =   6240
         TabIndex        =   39
         Text            =   "-100"
         Top             =   1920
         Width           =   495
      End
      Begin VB.TextBox Text_ChipBIN 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   4
         Left            =   7200
         TabIndex        =   38
         Text            =   "0"
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox Text_ChipY 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   4
         Left            =   6720
         TabIndex        =   37
         Text            =   "-100"
         Top             =   1920
         Width           =   495
      End
      Begin VB.TextBox Text_Chip 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   600
         TabIndex        =   36
         Text            =   "Chip"
         Top             =   4080
         Width           =   615
      End
      Begin VB.TextBox Text_ChipX 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   1200
         TabIndex        =   35
         Text            =   "-100"
         Top             =   4080
         Width           =   495
      End
      Begin VB.TextBox Text_ChipBIN 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   2160
         TabIndex        =   34
         Text            =   "0"
         Top             =   4080
         Width           =   615
      End
      Begin VB.TextBox Text_ChipY 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   1680
         TabIndex        =   33
         Text            =   "-100"
         Top             =   4080
         Width           =   495
      End
      Begin VB.TextBox Text_Chip 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   600
         TabIndex        =   32
         Text            =   "Chip"
         Top             =   4440
         Width           =   615
      End
      Begin VB.TextBox Text_ChipX 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   1200
         TabIndex        =   31
         Text            =   "-100"
         Top             =   4440
         Width           =   495
      End
      Begin VB.TextBox Text_ChipBIN 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   2160
         TabIndex        =   30
         Text            =   "0"
         Top             =   4440
         Width           =   615
      End
      Begin VB.TextBox Text_ChipY 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   1680
         TabIndex        =   29
         Text            =   "-100"
         Top             =   4440
         Width           =   495
      End
      Begin VB.TextBox Text_Chip 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   600
         TabIndex        =   28
         Text            =   "Chip"
         Top             =   4800
         Width           =   615
      End
      Begin VB.TextBox Text_ChipX 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   1200
         TabIndex        =   27
         Text            =   "-100"
         Top             =   4800
         Width           =   495
      End
      Begin VB.TextBox Text_ChipBIN 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   2160
         TabIndex        =   26
         Text            =   "0"
         Top             =   4800
         Width           =   615
      End
      Begin VB.TextBox Text_ChipY 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   1680
         TabIndex        =   25
         Text            =   "-100"
         Top             =   4800
         Width           =   495
      End
      Begin VB.TextBox Text_AreaX 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         Height          =   285
         Left            =   480
         TabIndex        =   24
         Text            =   "1"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox Text_AreaY 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         Height          =   285
         Left            =   1680
         TabIndex        =   23
         Text            =   "4"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox Text_Chip 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   600
         TabIndex        =   22
         Text            =   "Chip"
         Top             =   5160
         Width           =   615
      End
      Begin VB.TextBox Text_ChipX 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   1200
         TabIndex        =   21
         Text            =   "-100"
         Top             =   5160
         Width           =   495
      End
      Begin VB.TextBox Text_ChipBIN 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   2160
         TabIndex        =   20
         Text            =   "0"
         Top             =   5160
         Width           =   615
      End
      Begin VB.TextBox Text_StartX 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         Height          =   285
         Left            =   3240
         TabIndex        =   19
         Text            =   "4"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox Text_StartY 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         Height          =   285
         Left            =   3960
         TabIndex        =   18
         Text            =   "4"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox Text_ChipY 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   1680
         TabIndex        =   17
         Text            =   "-100"
         Top             =   5160
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Wafer X"
         Height          =   255
         Left            =   1080
         TabIndex        =   124
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "CHIP "
         Height          =   255
         Left            =   600
         TabIndex        =   123
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "BIN"
         Height          =   255
         Left            =   2280
         TabIndex        =   122
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label_Count 
         AutoSize        =   -1  'True
         Caption         =   "1"
         Height          =   195
         Index           =   15
         Left            =   5280
         TabIndex        =   121
         Top             =   5880
         Width           =   90
      End
      Begin VB.Label Label_Count 
         AutoSize        =   -1  'True
         Caption         =   "1"
         Height          =   195
         Index           =   14
         Left            =   5280
         TabIndex        =   120
         Top             =   5520
         Width           =   90
      End
      Begin VB.Label Label_Count 
         AutoSize        =   -1  'True
         Caption         =   "1"
         Height          =   195
         Index           =   13
         Left            =   5280
         TabIndex        =   119
         Top             =   5160
         Width           =   90
      End
      Begin VB.Label Label_Count 
         AutoSize        =   -1  'True
         Caption         =   "1"
         Height          =   195
         Index           =   12
         Left            =   5280
         TabIndex        =   118
         Top             =   4800
         Width           =   90
      End
      Begin VB.Label Label_Count 
         AutoSize        =   -1  'True
         Caption         =   "1"
         Height          =   195
         Index           =   11
         Left            =   5280
         TabIndex        =   117
         Top             =   4440
         Width           =   90
      End
      Begin VB.Label Label_Count 
         AutoSize        =   -1  'True
         Caption         =   "1"
         Height          =   195
         Index           =   10
         Left            =   5280
         TabIndex        =   116
         Top             =   4080
         Width           =   90
      End
      Begin VB.Label Label_Count 
         AutoSize        =   -1  'True
         Caption         =   "1"
         Height          =   195
         Index           =   9
         Left            =   5280
         TabIndex        =   115
         Top             =   3720
         Width           =   90
      End
      Begin VB.Label Label_Count 
         AutoSize        =   -1  'True
         Caption         =   "1"
         Height          =   195
         Index           =   8
         Left            =   5280
         TabIndex        =   114
         Top             =   3360
         Width           =   90
      End
      Begin VB.Label Label_Count 
         AutoSize        =   -1  'True
         Caption         =   "1"
         Height          =   195
         Index           =   7
         Left            =   5280
         TabIndex        =   113
         Top             =   3000
         Width           =   90
      End
      Begin VB.Label Label_Count 
         AutoSize        =   -1  'True
         Caption         =   "1"
         Height          =   195
         Index           =   6
         Left            =   5280
         TabIndex        =   112
         Top             =   2640
         Width           =   90
      End
      Begin VB.Label Label_Count 
         AutoSize        =   -1  'True
         Caption         =   "1"
         Height          =   195
         Index           =   5
         Left            =   5280
         TabIndex        =   111
         Top             =   2280
         Width           =   90
      End
      Begin VB.Label Label_Count 
         AutoSize        =   -1  'True
         Caption         =   "1"
         Height          =   195
         Index           =   4
         Left            =   5280
         TabIndex        =   110
         Top             =   1920
         Width           =   90
      End
      Begin VB.Label Label_Count 
         AutoSize        =   -1  'True
         Caption         =   "1"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   109
         Top             =   4080
         Width           =   90
      End
      Begin VB.Label Label_Count 
         AutoSize        =   -1  'True
         Caption         =   "1"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   108
         Top             =   4440
         Width           =   90
      End
      Begin VB.Label Label_Count 
         AutoSize        =   -1  'True
         Caption         =   "1"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   107
         Top             =   4800
         Width           =   90
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "X"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   106
         Top             =   360
         Width           =   120
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Y"
         Height          =   195
         Left            =   1440
         TabIndex        =   105
         Top             =   360
         Width           =   120
      End
      Begin VB.Label Label_Count 
         AutoSize        =   -1  'True
         Caption         =   "1"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   104
         Top             =   5160
         Width           =   90
      End
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   720
      Top             =   8640
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton Command_Option 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8160
      Picture         =   "mt2000.frx":0CCA
      Style           =   1  '±×·¡ÇÈ
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "OPTION"
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   1
      Left            =   6240
      Picture         =   "mt2000.frx":0E1A
      Style           =   1  '±×·¡ÇÈ
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton Command_Save 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   19920
      Picture         =   "mt2000.frx":1426
      Style           =   1  '±×·¡ÇÈ
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command_SaveAs 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7200
      Picture         =   "mt2000.frx":19CD
      Style           =   1  '±×·¡ÇÈ
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   240
      Width           =   855
   End
   Begin TabDlg.SSTab SSTab3 
      Height          =   13170
      Left            =   6000
      TabIndex        =   2
      Top             =   1200
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   23230
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   4
      TabHeight       =   529
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Map"
      TabPicture(0)   =   "mt2000.frx":1F8D
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label_BIN"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label_ChipPosition"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label_Move"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label6"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label_Ink"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "SSPanel1(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "HScroll_Zoom"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "VScroll_Zoom"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Original"
      TabPicture(1)   =   "mt2000.frx":1FA9
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "SSPanel7"
      Tab(1).ControlCount=   1
      Begin Threed.SSPanel SSPanel7 
         Height          =   13215
         Left            =   -74880
         TabIndex        =   7
         Top             =   360
         Width           =   12855
         _Version        =   65536
         _ExtentX        =   22675
         _ExtentY        =   23310
         _StockProps     =   15
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "µ¸¿ò"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         Begin VB.PictureBox pOriginal 
            Appearance      =   0  'Æò¸é
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  '¾øÀ½
            ForeColor       =   &H80000008&
            Height          =   13215
            Left            =   0
            ScaleHeight     =   881
            ScaleMode       =   3  'ÇÈ¼¿
            ScaleWidth      =   857
            TabIndex        =   8
            Top             =   0
            Width           =   12855
            Begin VB.Shape Shape_OInk 
               BorderColor     =   &H000080FF&
               BorderWidth     =   2
               FillColor       =   &H0000FFFF&
               Height          =   180
               Left            =   1080
               Top             =   120
               Width           =   180
            End
            Begin VB.Shape Shape_OMove 
               BorderColor     =   &H00FF00FF&
               BorderWidth     =   2
               FillColor       =   &H0000FFFF&
               Height          =   180
               Left            =   1800
               Top             =   120
               Width           =   180
            End
            Begin VB.Shape Shape_OChip 
               BorderColor     =   &H0000FFFF&
               BorderWidth     =   2
               FillColor       =   &H0000FFFF&
               Height          =   180
               Left            =   1440
               Top             =   180
               Width           =   180
            End
            Begin VB.Shape Shape_OFirstChip 
               BorderColor     =   &H00FF0000&
               BorderWidth     =   2
               FillColor       =   &H0000FFFF&
               Height          =   180
               Left            =   2100
               Top             =   120
               Width           =   180
            End
         End
      End
      Begin VB.VScrollBar VScroll_Zoom 
         Height          =   12495
         LargeChange     =   100
         Left            =   12720
         Max             =   1000
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   360
         Width           =   255
      End
      Begin VB.HScrollBar HScroll_Zoom 
         Height          =   255
         LargeChange     =   100
         Left            =   120
         Max             =   1000
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   12840
         Width           =   12615
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   12495
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   12615
         _Version        =   65536
         _ExtentX        =   22251
         _ExtentY        =   22040
         _StockProps     =   15
         ForeColor       =   -2147483630
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.PictureBox pZoom 
            Appearance      =   0  'Æò¸é
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  '¾øÀ½
            ForeColor       =   &H00C0FFFF&
            Height          =   12615
            Left            =   0
            ScaleHeight     =   841
            ScaleMode       =   3  'ÇÈ¼¿
            ScaleWidth      =   849
            TabIndex        =   6
            Top             =   0
            Width           =   12735
            Begin VB.CommandButton Command_Skip2Ink 
               Caption         =   "Skip to Ink"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   360
               TabIndex        =   319
               Top             =   3960
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.CommandButton Command10 
               Caption         =   "Mask to Ink"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   360
               TabIndex        =   318
               Top             =   4500
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.TextBox Text1 
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
               Left            =   7800
               TabIndex        =   278
               Top             =   9720
               Visible         =   0   'False
               Width           =   3255
            End
            Begin VB.TextBox Text1 
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
               Index           =   4
               Left            =   7800
               TabIndex        =   277
               Top             =   10080
               Visible         =   0   'False
               Width           =   3255
            End
            Begin VB.TextBox Text8 
               Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
               Height          =   375
               Left            =   1350
               TabIndex        =   274
               Top             =   480
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.TextBox Text9 
               Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
               Height          =   375
               Left            =   3150
               TabIndex        =   273
               Top             =   480
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.CommandButton Command8 
               Caption         =   "SEND"
               Height          =   375
               Left            =   3960
               TabIndex        =   272
               Top             =   480
               Visible         =   0   'False
               Width           =   1215
            End
            Begin Threed.SSPanel SSPanel16 
               Height          =   375
               Left            =   360
               TabIndex        =   275
               Top             =   480
               Visible         =   0   'False
               Width           =   975
               _Version        =   65536
               _ExtentX        =   1720
               _ExtentY        =   661
               _StockProps     =   15
               Caption         =   "X Size"
               ForeColor       =   0
               BackColor       =   15790320
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "µ¸¿ò"
                  Size            =   9.74
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin Threed.SSPanel SSPanel17 
               Height          =   375
               Left            =   2160
               TabIndex        =   276
               Top             =   480
               Visible         =   0   'False
               Width           =   975
               _Version        =   65536
               _ExtentX        =   1720
               _ExtentY        =   661
               _StockProps     =   15
               Caption         =   "Y Size"
               ForeColor       =   0
               BackColor       =   15790320
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "µ¸¿ò"
                  Size            =   9.74
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin Threed.SSPanel SSPanel4 
               Height          =   375
               Index           =   7
               Left            =   5880
               TabIndex        =   279
               Top             =   9720
               Visible         =   0   'False
               Width           =   1935
               _Version        =   65536
               _ExtentX        =   3413
               _ExtentY        =   661
               _StockProps     =   15
               Caption         =   "TYPE"
               ForeColor       =   -2147483630
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9.75
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
               Left            =   5880
               TabIndex        =   280
               Top             =   10080
               Visible         =   0   'False
               Width           =   1935
               _Version        =   65536
               _ExtentX        =   3413
               _ExtentY        =   661
               _StockProps     =   15
               Caption         =   "OPERATER NO."
               ForeColor       =   -2147483630
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9.75
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
            Begin TabDlg.SSTab SSTab2 
               Height          =   3975
               Left            =   6360
               TabIndex        =   282
               TabStop         =   0   'False
               Top             =   10800
               Visible         =   0   'False
               Width           =   5775
               _ExtentX        =   10186
               _ExtentY        =   7011
               _Version        =   393216
               Tabs            =   2
               TabsPerRow      =   2
               TabHeight       =   520
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TabCaption(0)   =   "PROBE FUNCTION"
               TabPicture(0)   =   "mt2000.frx":1FC5
               Tab(0).ControlEnabled=   -1  'True
               Tab(0).Control(0)=   "Command3"
               Tab(0).Control(0).Enabled=   0   'False
               Tab(0).Control(1)=   "Check_OldSPFile"
               Tab(0).Control(1).Enabled=   0   'False
               Tab(0).Control(2)=   "Command_Wafer_Center_Set"
               Tab(0).Control(2).Enabled=   0   'False
               Tab(0).Control(3)=   "Check_Pause"
               Tab(0).Control(3).Enabled=   0   'False
               Tab(0).Control(4)=   "Command_First_Move"
               Tab(0).Control(4).Enabled=   0   'False
               Tab(0).ControlCount=   5
               TabCaption(1)   =   "MAP FUNCTION"
               TabPicture(1)   =   "mt2000.frx":1FE1
               Tab(1).ControlEnabled=   0   'False
               Tab(1).ControlCount=   0
               Begin VB.CommandButton Command_First_Move 
                  Caption         =   $"mt2000.frx":1FFD
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "µ¸¿ò"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   615
                  Left            =   2640
                  TabIndex        =   287
                  TabStop         =   0   'False
                  Top             =   3000
                  Visible         =   0   'False
                  Width           =   1215
               End
               Begin VB.CheckBox Check_Pause 
                  Caption         =   "CONTINUE"
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "µ¸¿ò"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   615
                  Left            =   3960
                  Style           =   1  '±×·¡ÇÈ
                  TabIndex        =   286
                  TabStop         =   0   'False
                  Top             =   3000
                  Visible         =   0   'False
                  Width           =   1215
               End
               Begin VB.CommandButton Command_Wafer_Center_Set 
                  Caption         =   $"mt2000.frx":2010
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
                  Height          =   615
                  Left            =   360
                  TabIndex        =   285
                  TabStop         =   0   'False
                  Top             =   3000
                  Visible         =   0   'False
                  Width           =   1215
               End
               Begin VB.CheckBox Check_OldSPFile 
                  Caption         =   "Old SP File"
                  Height          =   255
                  Left            =   840
                  TabIndex        =   284
                  Top             =   3240
                  Visible         =   0   'False
                  Width           =   1335
               End
               Begin VB.CommandButton Command3 
                  Caption         =   "Chip Fail Set"
                  BeginProperty Font 
                     Name            =   "µ¸¿ò"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   615
                  Left            =   1320
                  TabIndex        =   283
                  Top             =   3000
                  Visible         =   0   'False
                  Width           =   1215
               End
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   127
               Left            =   240
               Top             =   480
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   126
               Left            =   0
               Top             =   480
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   125
               Left            =   480
               Top             =   240
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   124
               Left            =   240
               Top             =   240
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   123
               Left            =   0
               Top             =   240
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   122
               Left            =   480
               Top             =   0
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   121
               Left            =   240
               Top             =   0
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   120
               Left            =   0
               Top             =   0
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   119
               Left            =   240
               Top             =   480
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   118
               Left            =   0
               Top             =   480
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   117
               Left            =   480
               Top             =   240
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   116
               Left            =   240
               Top             =   240
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   115
               Left            =   0
               Top             =   240
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   114
               Left            =   480
               Top             =   0
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   113
               Left            =   240
               Top             =   0
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   112
               Left            =   0
               Top             =   0
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   111
               Left            =   240
               Top             =   480
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   110
               Left            =   0
               Top             =   480
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   109
               Left            =   480
               Top             =   240
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   108
               Left            =   240
               Top             =   240
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   107
               Left            =   0
               Top             =   240
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   106
               Left            =   480
               Top             =   0
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   105
               Left            =   240
               Top             =   0
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   104
               Left            =   0
               Top             =   0
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   103
               Left            =   240
               Top             =   480
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   102
               Left            =   0
               Top             =   480
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   101
               Left            =   480
               Top             =   240
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   100
               Left            =   240
               Top             =   240
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   99
               Left            =   0
               Top             =   240
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   98
               Left            =   480
               Top             =   0
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   97
               Left            =   240
               Top             =   0
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   96
               Left            =   0
               Top             =   0
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   95
               Left            =   240
               Top             =   480
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   94
               Left            =   0
               Top             =   480
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   93
               Left            =   480
               Top             =   240
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   92
               Left            =   240
               Top             =   240
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   91
               Left            =   0
               Top             =   240
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   90
               Left            =   480
               Top             =   0
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   89
               Left            =   240
               Top             =   0
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   88
               Left            =   0
               Top             =   0
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   87
               Left            =   240
               Top             =   480
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   86
               Left            =   0
               Top             =   480
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   85
               Left            =   480
               Top             =   240
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   84
               Left            =   240
               Top             =   240
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   83
               Left            =   0
               Top             =   240
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   82
               Left            =   480
               Top             =   0
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   81
               Left            =   240
               Top             =   0
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   80
               Left            =   0
               Top             =   0
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Ink 
               BorderColor     =   &H000080FF&
               BorderWidth     =   2
               FillColor       =   &H0000FFFF&
               Height          =   180
               Left            =   2400
               Top             =   1200
               Width           =   180
            End
            Begin VB.Shape Shape_FirstChip 
               BorderColor     =   &H00FF0000&
               BorderWidth     =   2
               FillColor       =   &H0000FFFF&
               Height          =   180
               Left            =   4320
               Top             =   1140
               Width           =   180
            End
            Begin VB.Shape Shape_Chip 
               BorderColor     =   &H0000FFFF&
               BorderWidth     =   2
               FillColor       =   &H0000FFFF&
               Height          =   180
               Left            =   2340
               Top             =   2160
               Width           =   180
            End
            Begin VB.Shape Shape_Move 
               BorderColor     =   &H00FF00FF&
               BorderWidth     =   2
               FillColor       =   &H0000FFFF&
               Height          =   180
               Left            =   3300
               Top             =   1740
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   7
               Left            =   4800
               Top             =   3960
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   6
               Left            =   4560
               Top             =   3960
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   5
               Left            =   5040
               Top             =   3720
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   4
               Left            =   4800
               Top             =   3720
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   3
               Left            =   4560
               Top             =   3720
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   2
               Left            =   5040
               Top             =   3480
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   1
               Left            =   4800
               Top             =   3480
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   0
               Left            =   4560
               Top             =   3480
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   79
               Left            =   7320
               Top             =   5760
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   78
               Left            =   7560
               Top             =   5760
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   77
               Left            =   7800
               Top             =   5760
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   76
               Left            =   7320
               Top             =   6000
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   75
               Left            =   7560
               Top             =   6000
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   74
               Left            =   7800
               Top             =   6000
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   73
               Left            =   7320
               Top             =   6240
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   72
               Left            =   7560
               Top             =   6240
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   71
               Left            =   6480
               Top             =   5760
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   70
               Left            =   6720
               Top             =   5760
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   69
               Left            =   6960
               Top             =   5760
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   68
               Left            =   6480
               Top             =   6000
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   67
               Left            =   6720
               Top             =   6000
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   66
               Left            =   6960
               Top             =   6000
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   65
               Left            =   6480
               Top             =   6240
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   64
               Left            =   6720
               Top             =   6240
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   63
               Left            =   5640
               Top             =   5760
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   62
               Left            =   5880
               Top             =   5760
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   61
               Left            =   6120
               Top             =   5760
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   60
               Left            =   5640
               Top             =   6000
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   59
               Left            =   5880
               Top             =   6000
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   58
               Left            =   6120
               Top             =   6000
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   57
               Left            =   5640
               Top             =   6240
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   56
               Left            =   5880
               Top             =   6240
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   55
               Left            =   4800
               Top             =   5760
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   54
               Left            =   5040
               Top             =   5760
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   53
               Left            =   5280
               Top             =   5760
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   52
               Left            =   4800
               Top             =   6000
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   51
               Left            =   5040
               Top             =   6000
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   50
               Left            =   5280
               Top             =   6000
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   49
               Left            =   4800
               Top             =   6240
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   48
               Left            =   5040
               Top             =   6240
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   47
               Left            =   6480
               Top             =   4800
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   46
               Left            =   6720
               Top             =   4800
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   45
               Left            =   6960
               Top             =   4800
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   44
               Left            =   6480
               Top             =   5040
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   43
               Left            =   6720
               Top             =   5040
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   42
               Left            =   6960
               Top             =   5040
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   41
               Left            =   6480
               Top             =   5280
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   40
               Left            =   6720
               Top             =   5280
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   39
               Left            =   5640
               Top             =   4800
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   38
               Left            =   5880
               Top             =   4800
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   37
               Left            =   6120
               Top             =   4800
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   36
               Left            =   5640
               Top             =   5040
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   35
               Left            =   5880
               Top             =   5040
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   34
               Left            =   6120
               Top             =   5040
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   33
               Left            =   5640
               Top             =   5280
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   32
               Left            =   5880
               Top             =   5280
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   31
               Left            =   4800
               Top             =   4800
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   30
               Left            =   5040
               Top             =   4800
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   29
               Left            =   5280
               Top             =   4800
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   28
               Left            =   4800
               Top             =   5040
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   27
               Left            =   5040
               Top             =   5040
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   26
               Left            =   5280
               Top             =   5040
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   25
               Left            =   4800
               Top             =   5280
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   24
               Left            =   5040
               Top             =   5280
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   23
               Left            =   6480
               Top             =   3600
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   22
               Left            =   6720
               Top             =   3600
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   21
               Left            =   6960
               Top             =   3600
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   20
               Left            =   6480
               Top             =   3840
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   19
               Left            =   6720
               Top             =   3840
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   18
               Left            =   6960
               Top             =   3840
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   17
               Left            =   6480
               Top             =   4080
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   16
               Left            =   6720
               Top             =   4080
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   15
               Left            =   5640
               Top             =   3600
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   14
               Left            =   5880
               Top             =   3600
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   13
               Left            =   6120
               Top             =   3600
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   12
               Left            =   5640
               Top             =   3840
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   11
               Left            =   5880
               Top             =   3840
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   10
               Left            =   6120
               Top             =   3840
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   9
               Left            =   5640
               Top             =   4080
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape Shape_Mea 
               BorderColor     =   &H000000FF&
               FillColor       =   &H0000FFFF&
               Height          =   180
               Index           =   8
               Left            =   5880
               Top             =   4080
               Visible         =   0   'False
               Width           =   180
            End
         End
      End
      Begin VB.Label Label_Ink 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BackColor       =   &H000080FF&
         BorderStyle     =   1  '´ÜÀÏ °íÁ¤
         Height          =   255
         Left            =   7200
         TabIndex        =   312
         Top             =   0
         Width           =   855
      End
      Begin VB.Label Label6 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  '´ÜÀÏ °íÁ¤
         Height          =   255
         Left            =   10080
         TabIndex        =   126
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label Label_Move 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  '´ÜÀÏ °íÁ¤
         Height          =   255
         Left            =   12240
         TabIndex        =   10
         Top             =   0
         Width           =   855
      End
      Begin VB.Label Label_ChipPosition 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  '´ÜÀÏ °íÁ¤
         Height          =   255
         Left            =   11400
         TabIndex        =   9
         Top             =   0
         Width           =   855
      End
      Begin VB.Label Label_BIN 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  '´ÜÀÏ °íÁ¤
         Height          =   255
         Left            =   8280
         TabIndex        =   11
         Top             =   0
         Width           =   1815
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   8520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  '¾Æ·¡ ¸ÂÃã
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   14475
      Width           =   19080
      _ExtentX        =   33655
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   25426
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "2022-09-30"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   1
            Text            =   "time"
            TextSave        =   "¿ÀÈÄ 5:42"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   480
      Top             =   6360
   End
   Begin VB.Label Label15 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackColor       =   &H00E0E0E0&
      Caption         =   "1CH"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12480
      TabIndex        =   363
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label14 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Fail Step :"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9120
      TabIndex        =   316
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label13 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Auto save count :"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9120
      TabIndex        =   271
      Top             =   795
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Åõ¸í
      Caption         =   "Copyright (c) 2008.11 by TSPS MECHATRONICS CO., LTD. All rights reserved."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   210
      Index           =   0
      Left            =   12840
      TabIndex        =   1
      Top             =   14160
      Width           =   6165
   End
End
Attribute VB_Name = "MT2000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Auto_flag As Boolean

Dim frequency As Currency
Dim startTime As Currency
Dim endTime As Currency
Dim result As Double

Dim Starprobe_Auto_probing_Flag As Boolean

Dim Test_End_Delay As ccrpStopWatch
Dim Test_Start_Delay As ccrpStopWatch
Dim WaitTime_Delay As ccrpStopWatch
Dim File_Time_Delay As ccrpStopWatch

Private Sub Check_Crack_wafer_Click()
    If Check_Crack_wafer.value = 1 Then
        Check_Crack_wafer.BackColor = &H40C0&
        Crack_Wafer = True
    Else
        Check_Crack_wafer.BackColor = &H8000000F
        Crack_Wafer = False
    End If
End Sub

Private Sub Check_No_Ink_Click()
    If Check_No_Ink.value = 1 Then
        Check_No_Ink.BackColor = &HC0C000
        Sample_No_Ink = True
        '[ 2022.07.20 ]
        If LOG_FILE_ON = 1 Then SelectExt.Log_Data_Save (11)
    Else
        Check_No_Ink.BackColor = &H8000000F
        Sample_No_Ink = False
        '[ 2022.07.20 ]
        If LOG_FILE_ON = 1 Then SelectExt.Log_Data_Save (12)
    End If
End Sub

Private Sub Check_No_Probe_Click()
    If Check_No_Probe.value = 1 Then
        Check_No_Probe.BackColor = &HC0C000
        No_Probe = True
        '[ 2022.07.20 ]
        If LOG_FILE_ON = 1 Then SelectExt.Log_Data_Save (9)
    Else
        Check_No_Probe.BackColor = &H8000000F
        No_Probe = False
        bStarprobe_AfterInk = False
        '[ 2022.07.20 ]
        If LOG_FILE_ON = 1 Then SelectExt.Log_Data_Save (10)
    End If
End Sub

Private Sub Check_Yes_Ink_Click()
    '[ 2021.12.31 ] : Sampling + ink(o)
    If Check_Yes_Ink.value = 1 Then
        Check_Yes_Ink.BackColor = &HC0C000
        Sample_No_Ink = True
        Sample_Yes_Ink = True
        '[ 2022.07.20 ]
        If LOG_FILE_ON = 1 Then SelectExt.Log_Data_Save (7)
    Else
        Check_Yes_Ink.BackColor = &H8000000F
        Sample_No_Ink = False
        Sample_Yes_Ink = False
        '[ 2022.07.20 ]
        If LOG_FILE_ON = 1 Then SelectExt.Log_Data_Save (8)
    End If
End Sub

Private Sub Check1_Click(Index As Integer)
    Select Case Index
        Case 1
            If Check1(1).value = 0 Then
                Check1(1).Caption = "Ink Start Pos Off"
                Ink_Start_Flag = 0
            Else
                Check1(1).Caption = "Ink Start Pos On"
                Ink_Start_Flag = 1
            End If
        Case 3
            If Check1(3).value = 1 Then
                Auto_flag = True
                
                If Server_path = "" Or Server_path = Empty Or InStr(Server_path, ":") = 0 Then
                    MsgBox "Input Data server path...", 16, "Error"
                    Check1(3).value = 0
                ElseIf Text1(0) <> LOT And Barcode_Use = True Then      '[ 2021.12.16 ] : ¹ÙÄÚµå¸¦ »ç¿ëÇÏ´Â °æ¿ì¿¡¸¸ Àû¿ëÇÑ´Ù.
                    If InStr(SSPanel2(0).Caption, ".SP") = 0 Then
                        MsgBox "LOT no check please!", 16, "Error"
                        Check1(3).value = 0
                    Else
                        '====================================================================================================================================================================
                        'New_Lot = Text1(0)
                        LOT = Text1(0)             '[ 2021.01.19 ]
                         
                        If Right(Server_path, 1) = "\" Then                 'ex)c:\ -> ±×³É»ç¿ë
                            R_File_Name = Server_path
                        Else                                                'ex)c:\test -> \Ãß°¡
                            R_File_Name = Server_path & "\"
                        End If
                        If Dir(R_File_Name, vbDirectory) = "" Then MkDir (R_File_Name)
                        
                        If SaveDrive = 0 Then
                            If Dir("C:\data\", vbDirectory) = "" Then MkDir ("C:\data\")
                            If Dir("C:\data\" & UCase(Text1(0)), vbDirectory) = "" Then MkDir ("C:\data\" & UCase(Text1(0)))
                        Else
                            If Dir("D:\data\", vbDirectory) = "" Then MkDir ("D:\data\")
                            If Dir("D:\data\" & UCase(Text1(0)), vbDirectory) = "" Then MkDir ("D:\data\" & UCase(Text1(0)))
                        End If
                        '====================================================================================================================================================================
                        AutoTest True
                    End If
                ElseIf Text1(0) = "" Then
                    MsgBox "Input Lot number..", 16, "Error"
                    Check1(3).value = 0
                Else
                    '====================================================================================================================================================================
                    'New_Lot = Text1(0)
                    LOT = Text1(0)             '[ 2021.01.19 ]
                     
                    If Right(Server_path, 1) = "\" Then                 'ex)c:\ -> ±×³É»ç¿ë
                        R_File_Name = Server_path
                    Else                                                'ex)c:\test -> \Ãß°¡
                        R_File_Name = Server_path & "\"
                    End If
                    If Dir(R_File_Name, vbDirectory) = "" Then MkDir (R_File_Name)
                    
                    If SaveDrive = 0 Then
                        If Dir("C:\data\", vbDirectory) = "" Then MkDir ("C:\data\")
                        If Dir("C:\data\" & UCase(Text1(0)), vbDirectory) = "" Then MkDir ("C:\data\" & UCase(Text1(0)))
                    Else
                        If Dir("D:\data\", vbDirectory) = "" Then MkDir ("D:\data\")
                        If Dir("D:\data\" & UCase(Text1(0)), vbDirectory) = "" Then MkDir ("D:\data\" & UCase(Text1(0)))
                    End If
                    '====================================================================================================================================================================
                    AutoTest True
                End If
            Else
                Auto_flag = False
                AutoTest False
            End If
    End Select
End Sub

Private Sub AutoTest(flag As Boolean)
    Dim s, val, sval, File_Name As String
    Dim ETS_LOT As String

    If flag = True Then
        '[ 2022.07.20 ]
        If LOG_FILE_ON = 1 Then SelectExt.Log_Data_Save (4)
        Text11.Enabled = False
        PrevTest = True                                                  'true
        Check1(3).BackColor = &HFF00&
        Command2(1).Enabled = False
        Command_Save.Enabled = False
        Command_SaveAs.Enabled = False
        
        Command_WaferDraw.Enabled = False
        Command_DisplayWafer.Enabled = False
        Command_Map_Clear.Enabled = False
        Command_First_Move.Enabled = False
        Command_ImageSave.Enabled = False
        Command_Option.Enabled = False

        '[ 2021.01.12 ] : inker position »ç¿ëºÒ°¡´É
        For i = 0 To 4
            Option6(i).Enabled = False
        Next i

        Check1(3).Enabled = False
        bStop = False
        Command7.Enabled = False
               
        Do
            DoEvents
            If Crack_Wafer = False Then
                If AutoAlign_Flag = False Then
                    XAxis = 0
                    YAxis = 0
                Else
                    If DemoMode = 0 Then
                        Call StarProbe_Auto_Probing
                        If Starprobe_Auto_probing_Flag = False Then
                            MsgBox "Probe Profile & Auto Aligned Fail !", 16, "STAR PROBE"
                            Check1(3).Enabled = True                '2021.09.02 : auto probing½ÇÆÐ·Î auto ¹öÆ°À» È°¼ºÈ­ ÇØÁØ´Ù.
                            Exit Do
                        End If
                    End If
                End If
            End If
            
            If DemoMode = 0 Then
                val = MSComm1.Input
                If ACCO = 1 Or Model_Select = 1 Then                            'ACCO only
                    sval = Replace(val, vbCrLf, "")
                Else
                    sval = Replace(val, vbLf, "")
                End If
                sval = Trim(sval)
                If sval <> Empty And Mid(sval, 1, 2) = "BA" Then
                    If ACCO = 1 Or Model_Select = 1 Then                            'ACCO only
                    Else
                        If DemoMode = 0 Then MSComm1.Output = ">" & vbLf              '¾ø¾îµµ »ó°ü ¾øÀ½.
                    End If
                End If
                
                If Wafer_Start = False Then                                                                 'Wafer ÃøÁ¤½ÃÀÛ
                    ETS_Count = File_Count + 1
                    If ACCO = 1 Or Model_Select = 1 Then                            'ACCO only
                        ETS_LOT = Trim(Str(ETS_Count))                                                      'WBI/1
                    Else
                        ETS_LOT = New_Lot & "-" & Trim(Str(ETS_Count))                                      'WBI/-1
                    End If
                    
                    s = "WBI" & "/" & ETS_LOT
                    If DemoMode = 0 Then
                        If ACCO = 1 Or Model_Select = 1 Then                            'ACCO only
                            MSComm1.Output = s & vbCrLf
                        Else
                            MSComm1.Output = s & vbLf
                        End If
                    End If
                    Wafer_Start = True
                End If
                If bStop = True Then Exit Do
                If bStarprobe_AfterInk = False Then
                    If TESTING_flag = False Then        '16.12.13
                        Call StarProbe_FistDie
                        If StarProbe_Motor_End_check Then MsgBox "Motor not end check !", 16, "STAR PROBE"
                    End If
                End If
                
                If No_Probe = False Or bStarprobe_AfterInk = False Then
                    Call StarProbe_Z_UP
                End If
            End If
    
            Test_Ready = True
            Command1(0).Enabled = False
            
            bStop = False
            
            StarProbe_WorkDateTime_From = CDate(Date$ & " " & Time$)
           
            If bStarprobe_AfterInk = False Then
                '============================================================ [ 2017.08.22 ]
                For XXX = 0 To 24
                    If NOW_NO(XXX) = True Then          'cassette¿¡¼­ ¼³Á¤ÇÑ ¼ø¼­´ë·Î tt_no¸¦ ¼³Á¤ÇÑ´Ù.
                        TT_NO = XXX
                        Exit For
                    End If
                Next XXX
                '============================================================
                lblWafer.Caption = TT_NO + 1            'mainÈ­¸é¿¡ wafer no¸¦ Ç¥½ÃÇÑ´Ù.
                
                If No_Probe = False Then
                    If DemoMode = 0 Then Call StarProbe_LAMP_OFF               '2015.12.07 : lamp off
                    If CH_SET = 1 Then
                        Call StarProbe_Auto_Test_CH1
                    ElseIf CH_SET = 2 Then
                        Call StarProbe_Auto_Test_CH2
                    Else
                        Call StarProbe_Auto_Test_CH4
                    End If
                Else
                    Call StarProbe_After_Ink_Dot
                End If
            Else
                If Ink_Start_Flag = 0 Then              '2016.09.27 : Normal Ink Mode
                    Call StarProbe_After_Ink_Dot
                Else                                    '2016.09.27 : Ink Start Position Set Mode
                    Call StarProbe_After_Ink_Dot_STT
                End If
            End If
            
            If bStop = True Then
                StarProbe_WorkDateTime_Total = StarProbe_WorkDateTime_Total + DateDiff("S", StarProbe_WorkDateTime_From, StarProbe_WorkDateTime_To)
                Exit Do
            Else
                StarProbe_WorkDateTime_Total = 0
            End If
        Loop
    Else
        '[ 2022.07.20 ]
        If LOG_FILE_ON = 1 Then SelectExt.Log_Data_Save (5)
        Command1(0).Enabled = True
        Check1(3).BackColor = &H8000000F
        Check1(3).value = 0                 ' 2004.05.11
        Check1(3).Enabled = True
 
        Command2(1).Enabled = True
        
        Command_Save.Enabled = True
        Command_SaveAs.Enabled = True
        Command_WaferDraw.Enabled = True
        Command_DisplayWafer.Enabled = True
        Command_Map_Clear.Enabled = True
        Command_First_Move.Enabled = True
        Command_ImageSave.Enabled = True
        Command_Option.Enabled = True
        
        '[ 2021.01.12 ]
        For i = 0 To 4
            Option6(i).Enabled = True
        Next i
        Test_Ready = False
    End If
    Exit Sub
End Sub

Private Sub Check2_Click()
    Dim forx As Integer, fory As Integer
    
    If Check2.value = 1 Then            'mask to ink
'        Check2.Caption = "Mask to Ink On"
        Check2.BackColor = &HFF00&
        For fory = 0 To StarProbe.ChipCountY
            For forx = 0 To StarProbe.ChipCountX
                If Wafer(forx, fory).Chip And Wafer(forx, fory).ChipMask_Backup Then
                    Wafer(forx, fory).ChipSkipDie = True
                    Wafer(forx, fory).ChipMask = False
                    Wafer(forx, fory).ChipInk = True
                    Call Display_Chip(pZoom, pOriginal, (forx - StarProbe.StartChip.x), (fory - StarProbe.StartChip.y))
                End If
            Next
        Next
    Else                                'ink to mask
'        Check2.Caption = "Mask to Ink Off"
        Check2.BackColor = &H8000000F
        For fory = 0 To StarProbe.ChipCountY
            For forx = 0 To StarProbe.ChipCountX
                If Wafer(forx, fory).Chip And Wafer(forx, fory).ChipMask_Backup Then
                    Wafer(forx, fory).ChipSkipDie = True
                    Wafer(forx, fory).ChipMask = True
                    Wafer(forx, fory).ChipInk = False
                    Call Display_Chip(pZoom, pOriginal, (forx - StarProbe.StartChip.x), (fory - StarProbe.StartChip.y))
                End If
            Next
        Next
    End If
End Sub

Private Sub Check3_Click()
    Dim forx As Integer, fory As Integer
    
    If Check3.value = 1 Then            'skip to ink
'        Check3.Caption = "Skip to Ink On"
        Check3.BackColor = &HFF00&
        For fory = 0 To StarProbe.ChipCountY
            For forx = 0 To StarProbe.ChipCountX
                If Wafer(forx, fory).Chip And Wafer(forx, fory).ChipSkipDie Then
                    Wafer(forx, fory).ChipInk = True
                    Call Display_Chip(pZoom, pOriginal, (forx - StarProbe.StartChip.x), (fory - StarProbe.StartChip.y))
                End If
            Next
        Next
    Else                                'ink to skip
'        Check3.Caption = "Skip to Ink Off"
        Check3.BackColor = &H8000000F
        For fory = 0 To StarProbe.ChipCountY
            For forx = 0 To StarProbe.ChipCountX
                If Wafer(forx, fory).Chip And Wafer(forx, fory).ChipSkipDie Then
                    Wafer(forx, fory).ChipInk = False
                    Call Display_Chip(pZoom, pOriginal, (forx - StarProbe.StartChip.x), (fory - StarProbe.StartChip.y))
                End If
            Next
        Next
    End If
End Sub

Private Sub Command_BINClear_Click(Index As Integer)
    Call StarProbe_Clear_Test(Index)
End Sub

Private Sub Command_BINColor_Click(Index As Integer)
    Dim lColor As Long

    lColor = Command_BINColor(Index).BackColor

    CommonDialog2.Flags = cdlCCRGBInit
    CommonDialog2.Color = lColor
    CommonDialog2.ShowColor
    
    lColor = CommonDialog2.Color
    
    Command_BINColor(Index).BackColor = lColor
    BINColor(Index) = lColor
    
    Call StarProbe_FileSave_SystemInfo
End Sub

Private Sub Command_ChipColor_Click(Index As Integer)
    Dim lColor As Long

    lColor = Command_ChipColor(Index).BackColor

    CommonDialog2.Flags = cdlCCRGBInit
    CommonDialog2.Color = lColor
    CommonDialog2.ShowColor
    
    lColor = CommonDialog2.Color
    
    Command_ChipColor(Index).BackColor = lColor
    ChipColor(Index) = lColor
    
    Call StarProbe_FileSave_SystemInfo
End Sub

Private Sub Command_DisplayWafer_Click()
    Dim i As Integer
    Dim oldflag As Boolean

    oldflag = Command_DisplayWafer.Enabled
    Command_DisplayWafer.Enabled = False
    
    Call Display_Wafer(pZoom, pOriginal, _
                       Shape_Chip, Shape_FirstChip, Shape_Ink, Shape_Move, _
                       Shape_OChip, Shape_OFirstChip, Shape_OInk, Shape_OMove, _
                       VScroll_Zoom, HScroll_Zoom)
                       
    SSPanel_WaferSize.Caption = StarProbe.WaferSizemm & "mm"
    SSPanel_ChipXSize.Caption = Format(StarProbe.ChipSizeX, "0.00000") & "mm"
    SSPanel_ChipYSize.Caption = Format(StarProbe.ChipSizeY, "0.00000") & "mm"
    
    If SP_FLAG = False Then         '[ 2017.03.23 ] : SPÆÄÀÏ ·ÎµùÀÌ ¾Æ´Ñ °æ¿ì¿¡¸¸ Àû¿ëÇÏµµ·Ï ¼öÁ¤.
        If AOI_MODE = 1 Then
            '[ 2020.11.02 ] : aoi
            Bin_Count(AOI_BIN) = AOI_FAIL_COUNT
            'ItemFailCount(AOI_BIN) = AOI_FAIL_COUNT
            Test_Cnt = AOI_FAIL_COUNT
            SSPanel2(1).Caption = Test_Cnt
            StarProbe.CountGoodDie = StarProbe.CountGoodDie ' - AOI_FAIL_COUNT
            StarProbe.CountBadDie = StarProbe.CountBadDie + AOI_FAIL_COUNT
        End If
        SSPanel_TotalCount.Caption = StarProbe.CountTotalChip
        SSPanel_GoodCount.Caption = StarProbe.CountGoodDie
        SSPanel_BadCount.Caption = StarProbe.CountBadDie
        Text_TotalCount.Text = StarProbe.CountTotalChip
        Text_GoodCount.Text = StarProbe.CountGoodDie
        Text_BadCount.Text = StarProbe.CountBadDie
        SSPanel_SkipCount.Caption = StarProbe.CountSkipDie
        Text_SkipCount.Text = StarProbe.CountSkipDie
        
        SSPanel_Yield.Caption = " "
    End If
    
    Text_BadCount.Text = StarProbe.CountBadDie
    Text_SkipCount.Text = StarProbe.CountSkipDie
       
    '2015.10.27
    If SP_FLAG = False Then StarProbe_WorkDateTime_Total = 0
    
    For i = 0 To 72
        Shape_Mea(i).width = StarProbe.DisplayChipSizeX + 1
        Shape_Mea(i).Height = StarProbe.DisplayChipSizeY + 1
    Next
    
    Command_First_Move.Enabled = True
    Check_Pause.Enabled = True
    Command_Stop.Enabled = True
    Command_DisplayWafer.Enabled = oldflag
    TESTING_flag = False
End Sub

Private Sub Command_First_Move_Click()
    Dim x, y As Integer    'Ãß°¡

    x = 0
    y = 0

    Call StarProbe_XY_Moving((x), (y))
    If StarProbe_Motor_End_check Then MsgBox "Motor not end check !", 16, "STAR PROBE"
End Sub

Private Sub Command_ImageSave_Click()
    Form_StarProbe_Save.Show vbModal, Me
    Exit Sub

    CommonDialog1.CancelError = True

    On Error GoTo ErrorSub
    
    Dim sfilename As String
    
    sfilename = "c:\Star Probe\Image\test.bmp"
    
    sfilename = Mid(sfilename, 1, InStr(1, sfilename, ".") - 1) & ".BMP"

    CommonDialog1.DialogTitle = "Wafer Image File BMP Save"
    CommonDialog1.Filter = "DB Files(*.BMP)|*.BMP"
    CommonDialog1.Flags = cdlOFNFileMustExist + cdlOFNOverwritePrompt + cdlOFNHideReadOnly
    CommonDialog1.FileName = IIf(Trim(sfilename) = Empty, "c:\Star Probe\Image\*.bmp", sfilename)
    CommonDialog1.ShowSave
    
    If CommonDialog1.FileName <> "" Then
        sfilename = CommonDialog1.FileName
        SavePicture pZoom.Image, sfilename
    End If
    
ErrorSub:
    CommonDialog1.CancelError = False
End Sub

Sub StarProbe_After_Ink_Dot_noink()
    Dim forx As Integer, fory As Integer
    Dim xx As Integer, yy As Integer
    Dim bRight As Boolean
    Dim bEnd As Boolean
    Dim File_Name, SP_File_Name As String
    Dim FindX As Integer, FindY As Integer
    Dim vx As Integer, vy As Integer
    Dim FindStep As Integer
    Dim bFind As Boolean
    Dim FindCount As Long
    Dim FindForX As Integer, FindForY As Integer
    Dim FindInk As Boolean
    Dim iStepX As Integer, iStepY As Integer
    
    '[ 2020.02.07 ] : ink needle position set
    If Option6(1).value = True Then                     'top
        NEEDLE_POS_X = 0
        NEEDLE_POS_Y = -1
    ElseIf Option6(0).value = True Then                 'bottom
        NEEDLE_POS_X = 0
        NEEDLE_POS_Y = 1
    ElseIf Option6(4).value = True Then                 'left
        NEEDLE_POS_X = -1
        NEEDLE_POS_Y = 0
    ElseIf Option6(3).value = True Then                 'center
        NEEDLE_POS_X = 0
        NEEDLE_POS_Y = 0
    ElseIf Option6(2).value = True Then                 'right
        NEEDLE_POS_X = 1
        NEEDLE_POS_Y = 0
    End If
    
    Set File_Time_Delay = New ccrpStopWatch
    
    bStarProbeStart = True
    bRight = True
    
    bStarprobe_AfterInk = True
    
    xx = StarProbe.StartChip.x
    yy = StarProbe.StartChip.y
                 
    '2015.11.06
    StarProbe_WorkDateTime_From = CDate(Date$ & " " & Time$)
    StarProbe_WorkDateTime_To = CDate(Date$ & " " & Time$)
    StarProbe_WorkDateTime_Total = StarProbe_WorkDateTime_Total + DateDiff("S", StarProbe_WorkDateTime_From, StarProbe_WorkDateTime_To)
            
    If DemoMode = 0 Then
        If StarProbe.Ink_After = 1 Or INK_OFF_TEST = True Then
            Call StarProbe_Z_Down
            Sleep 500
        
            Z = StarProbe_Z_Position
        
            If Z <> "D" Then
                MsgBox " Z Down Fail", vbOKOnly
                Exit Sub
            End If
        End If
    End If

    ''''''''''
    ' 2-Position Skip Die Ink
    bRight = True
    If Door1 = 0 Then bStop = False
    
    If Not bStop Then
        bStarprobe_AfterInk = False
        File_Time_Delay.Reset
        
        Do
            DoEvents
            If File_Time_Delay.Elapsed > 1000 Then Exit Do
        Loop
            
        Call Command_Map_Clear_Click
        Call Command_DisplayWafer_Click
            
        File_Time_Delay.Reset
            
        If DemoMode = 0 Then
            If ACCO = 1 Or Model_Select = 1 Then                            'ACCO only
                MSComm1.Output = "PC" & vbCrLf
            Else
                MSComm1.Output = "PC" & vbLf
            End If
        End If
            
        Do
            DoEvents
            If File_Time_Delay.Elapsed > 100 Then Exit Do
        Loop
        
        File_Time_Delay.Reset
            
        If Slot_Max_Count = File_Count Then
            File_Count = 0
            ETS_Count = 0
            If DemoMode = 0 Then
                If ACCO = 1 Or Model_Select = 1 Then                            'ACCO only
                    MSComm1.Output = "PCP1L0S0R0W/-1" & vbCrLf
                Else
                    MSComm1.Output = "PCP1L0S0R0W/-1" & vbLf
                End If
            End If
        Else
            If DemoMode = 0 Then
                If ACCO = 1 Or Model_Select = 1 Then                            'ACCO only
                    MSComm1.Output = "PCP1L0S1R1W/-1" & vbCrLf
                Else
                    MSComm1.Output = "PCP1L0S1R1W/-1" & vbLf
                End If
            End If
        End If
                        
        Old_Lot = New_Lot
                        
        Do
            DoEvents
            If File_Time_Delay.Elapsed > 100 Then Exit Do
        Loop
        File_Time_Delay.Reset
            
        If DemoMode = 0 Then
            If ACCO = 1 Or Model_Select = 1 Then                            'ACCO only
            
            Else
                MSComm1.Output = ">" & vbLf
            End If
        End If
        Do
            DoEvents
            If File_Time_Delay.Elapsed > 100 Then Exit Do
        Loop
               
        Wafer_Start = False
        
        RESET_DATA
  
        Call StarProbe_Unlode_Wafer_New_Wafer
        If Crack_Wafer = False Then
        Else
            Check1(3).value = 0
            bStop = True
        End If
        
        File_Time_Delay.Reset
        Do
            DoEvents
            If File_Time_Delay.Elapsed > 30000 Then Exit Do     '[ 2017.03.27 ] : wafer unloadingÈÄ delay ¼öÁ¤ 10000->30000
        Loop
        Set File_Time_Delay = Nothing
        
        '2016.06.14 Å¸ÀÓÅ¬¸®¾îÃß°¡
        StarProbe_WorkDateTime_Total = 0
        StarProbe_WorkDateTime_From = 0
        StarProbe_WorkDateTime_To = 0
        ''''''''''''''''''''''''''''''''''''''''''''''''''''2016.09.22
        If AutoAlign_Flag = False Then
            Check1(3).value = 0
            Check1(3).Enabled = True
            bStop = True
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''
        bStarprobe_AfterInk = False         '2016.06.29
        Check1(1).value = 0                 '2016.09.27 : Ink Start Position Check Box Off
        Ink_Start_Flag = 0                  '2016.09.27 : Ink Start Position Flag Clear
        
        '[ 2020.10.29 ] : ÃøÁ¤ÀÌ ³¡³ª¸é Ä«¼¼Æ®ÀÇ Ã¼Å©¸¦ Ç¥½ÃÇÑ´Ù.
        NOW_NO(val(lblWafer.Caption) - 1) = False
        
        If AOI_MODE = 1 Then
            '[ 2020.10.29 ] : AOI°ü·Ã ¼öÁ¤     --> ÆÄÀÏÀÌ¸§¹Ù²Ù±â -> ´ÙÀ½ÆÄÀÏ·Îµù -> ¹æÇâÀüÈ¯ -> first die ¼³Á¤
            If UCase(Right(SSPanel2(0).Caption, 3)) = "AOI" Then
                Name AOI_MAP(lblWafer.Caption) As AOI_MAP(lblWafer.Caption) & "1"                 'ÆÄÀÏ ÀÌ¸§À» ¹Ù²ãÁØ´Ù. *.aoi --> *.aoi1
                If AutoAlign_Flag = True Then
                    Dim load_no As Integer
                                
                    For II = 0 To 24
                        If NOW_NO(II) = True Then
                            load_no = II
                            Exit For
                        End If
                    Next II
                    
                    'lblWafer°ú ÀÏÄ¡ÇÏ´Â ***.aoiÆÄÀÏÀ» ºÒ·¯¿Â´Ù.
                    'AOIÆÄÀÏÀÌ¸§ÀÌ wafer no¿Í °°Àº °æ¿ì´Â ´Ù½ÃºÎ¸£Áö ¾Ê°í ³Ñ¾î°£´Ù.
                    If load_no <> 0 Then
                        Command_Map_Clear.Enabled = False
            
                        For xx = 0 To StarProbe.ChipCountX                      'map º¯¼ö ÃÊ±âÈ­
                            For yy = 0 To StarProbe.ChipCountY
                                Wafer(xx, yy).flag = False
                                Wafer(xx, yy).FlagBad = False
                                Wafer(xx, yy).MeasureWait = False
                                Wafer(xx, yy).InkDot = False
                                Wafer(xx, yy).ChipMeasure = False
                                Wafer(xx, yy).BIN = 0
                            Next
                        Next
                        Fail_Loop = False
                        Stop_Measure = False
                        Command_Map_Clear.Enabled = True
                        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        Dim map_first As String                                 'Ã³À½ loadÇÒ ÆÄÀÏ ÀÌ¸§ ÁöÁ¤ º¯¼ö
                        For i = 0 To 24
                            If NOW_NO(i) = True Then
                                map_first = AOI_MAP(i + 1)
                                Exit For
                            End If
                        Next i
                        
                        If map_first = "" Then                              '[ 2021.09.17 ] : AOIÆÄÀÏÀÌ ¾ø´Â °æ¿ì , lot endÀÎ °æ¿ì ºüÁ®³ª°£´Ù.
                            Exit Sub
                        End If
                        
                        '============================================== 2015.11.30 : load map name display
                        For i = Len(map_first) To 1 Step -1
                            If Mid(map_first, i, 1) = "\" Then
                                SSPanel2(0).Caption = UCase(Mid(map_first, i + 1, Len(map_first) - i + 1))
                                Exit For
                            End If
                        Next i
                        '==============================================
                        
                        
                        bStarprobe_AfterInk = False                             'INK MODE OFF
                        For i = 0 To 8
                            Shape_Mea(i).Visible = False
                        Next i
                        StarProbe.FileName_Map = Load_MAP
                
                        StarProbeTemp = StarProbe
                        
                        Call StarProbe_FileLoad_ControlMap(map_first)
                                                                                                  
                        StarProbe.CountBadDie = StarProbeTemp.CountBadDie
                        StarProbe.CountGoodDie = StarProbeTemp.CountGoodDie
                        StarProbe.CountSkipDie = StarProbeTemp.CountSkipDie
                        StarProbe.CountTotalChip = StarProbeTemp.CountTotalChip
                
                        '[ 2020.03.27 ] : DB°ü·Ã ¼öÁ¤
                        ''''''''''''''''''''''''''''''''''''''''''''''''
                        Map_Total_Backup = StarProbe.CountTotalChip
                        Map_Good_Backup = StarProbe.CountGoodDie
                        Map_Bad_Backup = StarProbe.CountBadDie
                        Map_Skip_Backup = StarProbe.CountSkipDie
                        '''''''''''''''''''''''''''''''''''''''''''''''''
                
                        'Call Command_Map_Clear_Click                '2016.06.02 : map clearÃß°¡
                
                        Call Command_DisplayWafer_Click
                
                        StarProbe.MeasureStepX = StarProbeTemp.MeasureStepX
                        StarProbe.MeasureStepY = StarProbeTemp.MeasureStepY
                
                        StarProbe.Ink_LeftPort = StarProbeTemp.Ink_LeftPort
                        StarProbe.Ink_RightPort = StarProbeTemp.Ink_RightPort
                        StarProbe.LineOk = StarProbeTemp.LineOk
                        StarProbe.RCount = StarProbeTemp.RCount
                        StarProbe.RCount_Sub = StarProbeTemp.RCount_Sub
                        StarProbe.MeasureSleep = StarProbeTemp.MeasureSleep
                        StarProbe.ReMeasure = StarProbeTemp.ReMeasure
                        StarProbe.Ink_After_LeftPort = StarProbeTemp.Ink_After_LeftPort
                        StarProbe.Ink_After_RightPort = StarProbeTemp.Ink_After_RightPort
                        StarProbe.Ink_After_CenterPort = StarProbeTemp.Ink_After_CenterPort
                        StarProbe.LimitArea = StarProbeTemp.LimitArea
                
                        StarProbe.WaferTest = StarProbeTemp.WaferTest
'                        StarProbe.MeasureAll = StarProbeTemp.MeasureAll
                
                        StarProbe.Ink_After = StarProbeTemp.Ink_After
                
                        Me.MousePointer = 0
                        Command2(1).Enabled = True
                        
                        Call Command_WaferDirection_Click               '[ 2020.12.17 ] : ¹æÇâ¼³Á¤ Ãß°¡
                        
                        '[ 2021.06.25 ] first die
                        StarProbe.StartChip.x = First_X
                        StarProbe.StartChip.y = First_Y
                        
                        Shape_FirstChip.Top = First_Zoom_TOP
                        Shape_FirstChip.Left = First_Zoom_LEFT

                        Shape_OFirstChip.Top = First_Original_TOP
                        Shape_OFirstChip.Left = First_Original_LEFT
                    End If
                End If
            End If
        End If
    End If
    Exit Sub
End Sub

Sub StarProbe_After_Ink_Dot()
    Dim forx As Integer, fory As Integer
    Dim xx As Integer, yy As Integer
    Dim bRight As Boolean
    Dim bEnd As Boolean
    Dim File_Name, SP_File_Name As String
    Dim FindX As Integer, FindY As Integer
    Dim vx As Integer, vy As Integer
    Dim FindStep As Integer
    Dim bFind As Boolean
    Dim FindCount As Long
    Dim FindForX As Integer, FindForY As Integer
    Dim FindInk As Boolean
    Dim iStepX As Integer, iStepY As Integer
    
    '[ 2020.02.07 ] : ink needle position set
    If Option6(1).value = True Then                     'top
        NEEDLE_POS_X = 0
        NEEDLE_POS_Y = -1
    ElseIf Option6(0).value = True Then                 'bottom
        NEEDLE_POS_X = 0
        NEEDLE_POS_Y = 1
    ElseIf Option6(4).value = True Then                 'left
        NEEDLE_POS_X = -1
        NEEDLE_POS_Y = 0
    ElseIf Option6(3).value = True Then                 'center
        NEEDLE_POS_X = 0
        NEEDLE_POS_Y = 0
    ElseIf Option6(2).value = True Then                 'right
        NEEDLE_POS_X = 1
        NEEDLE_POS_Y = 0
    End If
    
    Set File_Time_Delay = New ccrpStopWatch
    
    bStarProbeStart = True
    bRight = True
    
    bStarprobe_AfterInk = True
    
    xx = StarProbe.StartChip.x
    yy = StarProbe.StartChip.y
                 
    '2015.11.06
    StarProbe_WorkDateTime_From = CDate(Date$ & " " & Time$)
    StarProbe_WorkDateTime_To = CDate(Date$ & " " & Time$)
    StarProbe_WorkDateTime_Total = StarProbe_WorkDateTime_Total + DateDiff("S", StarProbe_WorkDateTime_From, StarProbe_WorkDateTime_To)
            
    If DemoMode = 0 Then
        If StarProbe.Ink_After = 1 Or INK_OFF_TEST = True Then
            Call StarProbe_Z_Down
            Sleep 500
        
            Z = StarProbe_Z_Position
        
            If Z <> "D" Then
                MsgBox " Z Down Fail", vbOKOnly
                Exit Sub
            End If
        End If
    End If

    ''''''''''
    ' 2-Position Skip Die Ink
    bRight = True
    If Door1 = 0 Then bStop = False
    
    'If Sample_No_Ink = False Then
    If Sample_No_Ink = False Or Sample_Yes_Ink = True Then          '[ 2021.12.31 ] : Sampling + ink(o)
        ' 1
        forx = StarProbe.ChipCountX \ 2
        For fory = StarProbe.ChipCountY To 0 Step -1
            DoEvents
            Do While Not bStop
                DoEvents
                If InkRun((forx), (fory)) Then
                    If StarProbe.Ink_After_CenterPort = 0 Then      'direct ink
                        xx = forx - StarProbe.StartChip.x + IIf(bRight, 1, -1)
                    Else
                        xx = forx - StarProbe.StartChip.x          'after ink
                    End If
                    yy = fory - StarProbe.StartChip.y
                
                    VScroll_Zoom.value = (fory / (Abs(StarProbe.Max.y) + Abs(StarProbe.Min.y) + 1)) * 1000
                    HScroll_Zoom.value = (forx / (Abs(StarProbe.Max.x) + Abs(StarProbe.Min.x) + 1)) * 1000
                
                    StarProbe.CurrentChip.x = xx
                    StarProbe.CurrentChip.y = yy
                
                    Text5 = StarProbe.CurrentChip.x
                    Text6 = StarProbe.CurrentChip.y

                    Shape_Chip.Top = (fory * StarProbe.DisplayChipSizeY) - 2
                    Shape_Chip.Left = (forx * StarProbe.DisplayChipSizeX) - 2
                
                    Label_ChipPosition = StarProbe.CurrentChip.x & "/" & StarProbe.CurrentChip.y
                
                    'Call StarProbe_XY_Moving((xx), (yy - 4))
                    '====================================================================================================================================
                    If CH_SET = 1 Then
                        Call StarProbe_XY_Moving((xx + NEEDLE_POS_X), (yy + (NEEDLE_POS_Y * 1)))                         '[ 2020.01.20 ] : ink needle position
                    ElseIf CH_SET = 2 Then
                        Call StarProbe_XY_Moving((xx + NEEDLE_POS_X), (yy + (NEEDLE_POS_Y * 1)))                         '[ 2020.01.20 ] : ink needle position
                    Else
                        Call StarProbe_XY_Moving((xx + NEEDLE_POS_X), (yy + (NEEDLE_POS_Y * 4)))                         '[ 2020.01.20 ] : ink needle position
                    End If
                    '====================================================================================================================================
                                          
                    bEnd = False
                           
                    Do While (Not bEnd)
                        DoEvents
                        If Not StarProbe_Motor_End_check Then
                            bEnd = True
                        Else
                            MsgBox "Motor not end check !", 16, "STAR PROBE"
                            bEnd = True
                        End If
                        If ErrorStop = True Or bEnd = False Then bEnd = True
                    Loop
                           
                    If StarProbe.Ink_After_LeftPort = 1 Or StarProbe.Ink_After_RightPort = 1 Or StarProbe.Ink_After_CenterPort = 1 Then
                        Sleep 10
                    End If

                    If bRight And StarProbe.Ink_After_LeftPort = 1 And InkRun_Left(xx, yy) Then
                        Call StarProbe_Left_Ink_Dot(StarProbe.Ink_LeftPort)
                        Call InkRun_LeftOk((xx), (yy))
                        FindCount = FindCount - 1
                        If DemoMode = 1 Then Call Display_Chip_demo(pZoom, pOriginal, xx, yy)     'È­¸é¿¡ Ç¥½Ã
                    End If
                
                    If Not bRight And StarProbe.Ink_After_RightPort = 1 And InkRun_Right(xx, yy) Then
                        Call StarProbe_Right_Ink_Dot(StarProbe.Ink_RightPort)
                        Call InkRun_RightOk((xx), (yy))
                        FindCount = FindCount - 1
                        If DemoMode = 1 Then Call Display_Chip_demo(pZoom, pOriginal, xx, yy)     'È­¸é¿¡ Ç¥½Ã
                    End If
                
                    If StarProbe.Ink_After_CenterPort = 1 And InkRun_Center(xx, yy) Then   ' after center
                        Call StarProbe_Left_Ink_Dot(StarProbe.Ink_LeftPort)
                        Call InkRun_CenterOk((xx), (yy))
                        FindCount = FindCount - 1
                        If DemoMode = 1 Then Call Display_Chip_demo(pZoom, pOriginal, xx, yy)     'È­¸é¿¡ Ç¥½Ã
                    End If
                    StarProbe_WorkDateTime_To = CDate(Date$ & " " & Time$)
                    Call StarProbe_WorkDateTime_HMS(StarProbe_WorkDateTime_Total + DateDiff("S", StarProbe_WorkDateTime_From, StarProbe_WorkDateTime_To))
                    SSPanel_DateTime.Caption = StarProbe_WorkDateTime.D & " Day " & StarProbe_WorkDateTime.h & ":" & StarProbe_WorkDateTime.M & ":" & StarProbe_WorkDateTime.s
                End If
            
                If bRight Then
                    forx = forx + 1
                    If forx > (StarProbe.ChipCountX + 1) Then
                        forx = (StarProbe.ChipCountX)
                        bRight = False
                        Exit Do
                    End If
                Else
                    forx = forx - 1
                    If forx < (StarProbe.ChipCountX \ 2) Then
                        forx = (StarProbe.ChipCountX \ 2)
                        bRight = True
                        Exit Do
                    End If
                End If
            Loop
        Next

        bRight = False

        ' 2
        forx = StarProbe.ChipCountX \ 2
        For fory = 0 To StarProbe.ChipCountY
            DoEvents
            Do While Not bStop
                DoEvents
                If InkRun((forx), (fory)) Then
                    If StarProbe.Ink_After_CenterPort = 0 Then
                        xx = forx - StarProbe.StartChip.x + IIf(bRight, 1, -1)
                    Else
                        xx = forx - StarProbe.StartChip.x
                    End If
                    yy = fory - StarProbe.StartChip.y

                    VScroll_Zoom.value = (fory / (Abs(StarProbe.Max.y) + Abs(StarProbe.Min.y) + 1)) * 1000
                    HScroll_Zoom.value = (forx / (Abs(StarProbe.Max.x) + Abs(StarProbe.Min.x) + 1)) * 1000

                    StarProbe.CurrentChip.x = xx
                    StarProbe.CurrentChip.y = yy

                    Text5 = StarProbe.CurrentChip.x
                    Text6 = StarProbe.CurrentChip.y

                    Shape_Chip.Top = (fory * StarProbe.DisplayChipSizeY) - 2
                    Shape_Chip.Left = (forx * StarProbe.DisplayChipSizeX) - 2

                    Label_ChipPosition = StarProbe.CurrentChip.x & "/" & StarProbe.CurrentChip.y

                    'Call StarProbe_XY_Moving((xx), (yy - 4))
                    '====================================================================================================================================
                    If CH_SET = 1 Then
                        Call StarProbe_XY_Moving((xx + NEEDLE_POS_X), (yy + (NEEDLE_POS_Y * 1)))                         '[ 2020.01.20 ] : ink needle position
                    ElseIf CH_SET = 2 Then
                        Call StarProbe_XY_Moving((xx + NEEDLE_POS_X), (yy + (NEEDLE_POS_Y * 1)))                         '[ 2020.01.20 ] : ink needle position
                    Else
                        Call StarProbe_XY_Moving((xx + NEEDLE_POS_X), (yy + (NEEDLE_POS_Y * 4)))                         '[ 2020.01.20 ] : ink needle position
                    End If
                    '====================================================================================================================================

                    bEnd = False

                    Do While (Not bEnd)
                        DoEvents
                        If Not StarProbe_Motor_End_check Then
                            bEnd = True
                        Else
                            MsgBox "Motor not end check !", 16, "STAR PROBE"
                            bEnd = True
                        End If
                        If ErrorStop = True Or bEnd = False Then bEnd = True
                    Loop
                    
                    If StarProbe.Ink_After_LeftPort = 1 Or StarProbe.Ink_After_RightPort = 1 Or StarProbe.Ink_After_CenterPort = 1 Then
                        Sleep 10
                    End If

                    If bRight And StarProbe.Ink_After_LeftPort = 1 And InkRun_Left(xx, yy) Then
                        Call StarProbe_Left_Ink_Dot(StarProbe.Ink_LeftPort)
                        Call InkRun_LeftOk((xx), (yy))
                        FindCount = FindCount - 1
                        If DemoMode = 1 Then Call Display_Chip_demo(pZoom, pOriginal, xx, yy)     'È­¸é¿¡ Ç¥½Ã
                    End If

                    If Not bRight And StarProbe.Ink_After_RightPort = 1 And InkRun_Right(xx, yy) Then
                        Call StarProbe_Right_Ink_Dot(StarProbe.Ink_RightPort)
                        Call InkRun_RightOk((xx), (yy))
                        FindCount = FindCount - 1
                        If DemoMode = 1 Then Call Display_Chip_demo(pZoom, pOriginal, xx, yy)     'È­¸é¿¡ Ç¥½Ã
                    End If

                    If StarProbe.Ink_After_CenterPort = 1 And InkRun_Center(xx, yy) Then
                        Call StarProbe_Left_Ink_Dot(StarProbe.Ink_LeftPort)
                        Call InkRun_CenterOk((xx), (yy))
                        FindCount = FindCount - 1
                        If DemoMode = 1 Then Call Display_Chip_demo(pZoom, pOriginal, xx, yy)     'È­¸é¿¡ Ç¥½Ã
                    End If
                    StarProbe_WorkDateTime_To = CDate(Date$ & " " & Time$)
                    Call StarProbe_WorkDateTime_HMS(StarProbe_WorkDateTime_Total + DateDiff("S", StarProbe_WorkDateTime_From, StarProbe_WorkDateTime_To))
                    SSPanel_DateTime.Caption = StarProbe_WorkDateTime.D & " Day " & StarProbe_WorkDateTime.h & ":" & StarProbe_WorkDateTime.M & ":" & StarProbe_WorkDateTime.s
                End If
            
                If bRight Then
                    forx = forx + 1
                    If forx > StarProbe.ChipCountX \ 2 Then
                        forx = (StarProbe.ChipCountX \ 2)
                        bRight = False
                        Exit Do
                    End If
                Else
                    forx = forx - 1
                    If forx < -1 Then
                        forx = 0
                        bRight = True
                        Exit Do
                    End If
                End If
            Loop
        Next

' 2-Position Skip Die Ink
''''''''''

''''''''''
' Ink

        FindCount = 0
        
        iStepX = StarProbe.CurrentChip.x + StarProbe.StartChip.x
        iStepY = StarProbe.CurrentChip.y + StarProbe.StartChip.y

        FindStep = 1
          
        Do While Not bStop
            DoEvents
            If FindCount <= 0 Then Exit Do
            Exit Do
            
            bFind = False
        
            If InkRun_Center((iStepX), (iStepY)) Then
                FindX = iStepX
                FindY = iStepY
                bFind = True
            End If
        
            FindStep = 1
        
            Do While Not bStop
                DoEvents
                If bFind Then Exit Do
            
                ' 2
                vx = iStepX + FindStep
                vy = iStepY
                If vx > StarProbe.ChipCountX Then vx = StarProbe.ChipCountX
                If vx < 0 Then vx = 0
                If vy > StarProbe.ChipCountY Then vy = StarProbe.ChipCountY
                If vy < 0 Then vy = 0
                
                If Not bFind And InkRun((vx), (vy)) Then
                    FindX = vx: FindY = vy: bFind = True
                End If
                
                ' 3
                vx = iStepX - FindStep
                vy = iStepY
                If vx > StarProbe.ChipCountX Then vx = StarProbe.ChipCountX
                If vx < 0 Then vx = 0
                If vy > StarProbe.ChipCountY Then vy = StarProbe.ChipCountY
                If vy < 0 Then vy = 0
                If Not bFind And InkRun_Center((vx), (vy)) Then
                    FindX = vx: FindY = vy: bFind = True
                End If
                
                ' 1
                vx = iStepX
                vy = iStepY - FindStep
                If vx > StarProbe.ChipCountX Then vx = StarProbe.ChipCountX
                If vx < 0 Then vx = 0
                If vy > StarProbe.ChipCountY Then vy = StarProbe.ChipCountY
                If vy < 0 Then vy = 0
                If Not bFind And InkRun_Center((vx), (vy)) Then
                    FindX = vx: FindY = vy: bFind = True
                End If
                    
                ' 4
                vx = iStepX
                vy = iStepY + FindStep
                If vx > StarProbe.ChipCountX Then vx = StarProbe.ChipCountX
                If vx < 0 Then vx = 0
                If vy > StarProbe.ChipCountY Then vy = StarProbe.ChipCountY
                If vy < 0 Then vy = 0
                If Not bFind And InkRun((vx), (vy)) Then
                    FindX = vx: FindY = vy: bFind = True
                End If
                    
                ' 5
                vx = iStepX + FindStep
                vy = iStepY - FindStep
                If vx > StarProbe.ChipCountX Then vx = StarProbe.ChipCountX
                If vx < 0 Then vx = 0
                If vy > StarProbe.ChipCountY Then vy = StarProbe.ChipCountY
                If vy < 0 Then vy = 0
                If Not bFind And InkRun((vx), (vy)) Then
                    FindX = vx: FindY = vy: bFind = True
                End If
                    
                ' 6
                vx = iStepX - FindStep
                vy = iStepY - FindStep
                If vx > StarProbe.ChipCountX Then vx = StarProbe.ChipCountX
                If vx < 0 Then vx = 0
                If vy > StarProbe.ChipCountY Then vy = StarProbe.ChipCountY
                If vy < 0 Then vy = 0
                If Not bFind And InkRun((vx), (vy)) Then
                    FindX = vx: FindY = vy: bFind = True
                End If
                    
                ' 7
                vx = iStepX + FindStep
                vy = iStepY + FindStep
                If vx > StarProbe.ChipCountX Then vx = StarProbe.ChipCountX
                If vx < 0 Then vx = 0
                If vy > StarProbe.ChipCountY Then vy = StarProbe.ChipCountY
                If vy < 0 Then vy = 0
                If Not bFind And InkRun((vx), (vy)) Then
                    FindX = vx: FindY = vy: bFind = True
                End If
                    
                ' 8
                vx = iStepX - FindStep
                vy = iStepY + FindStep
                If vx > StarProbe.ChipCountX Then vx = StarProbe.ChipCountX
                If vx < 0 Then vx = 0
                If vy > StarProbe.ChipCountY Then vy = StarProbe.ChipCountY
                If vy < 0 Then vy = 0
                If Not bFind And InkRun((vx), (vy)) Then
                    FindX = vx: FindY = vy: bFind = True
                End If
                
                If Not bFind And FindStep > 1 Then
                    ' 9
                    vy = iStepY - FindStep
                    For FindForX = (iStepX + 1) To ((iStepX + 1) + (FindStep - 2))
                        vx = FindForX
                        If vx > StarProbe.ChipCountX Then vx = StarProbe.ChipCountX
                        If vx < 0 Then vx = 0
                        If vy > StarProbe.ChipCountY Then vy = StarProbe.ChipCountY
                        If vy < 0 Then vy = 0
                        If Not bFind And InkRun((vx), (vy)) Then
                            FindX = vx: FindY = vy: bFind = True
                            Exit For
                        End If
                    Next
            
                    ' 10
                    vy = iStepY - FindStep
                    For FindForX = (iStepX - 1) To ((iStepX - 1) - (FindStep - 2)) Step -1
                        vx = FindForX
                        If vx > StarProbe.ChipCountX Then vx = StarProbe.ChipCountX
                        If vx < 0 Then vx = 0
                        If vy > StarProbe.ChipCountY Then vy = StarProbe.ChipCountY
                        If vy < 0 Then vy = 0
                        If Not bFind And InkRun((vx), (vy)) Then
                            FindX = vx: FindY = vy: bFind = True
                            Exit For
                        End If
                    Next
            
                    ' 11
                    vx = iStepX + FindStep
                    For FindForY = (iStepY - 1) To ((iStepY - 1) - (FindStep - 2)) Step -1
                        vy = FindForY
                        If vx > StarProbe.ChipCountX Then vx = StarProbe.ChipCountX
                        If vx < 0 Then vx = 0
                        If vy > StarProbe.ChipCountY Then vy = StarProbe.ChipCountY
                        If vy < 0 Then vy = 0
                        If Not bFind And InkRun((vx), (vy)) Then
                            FindX = vx: FindY = vy: bFind = True
                            Exit For
                        End If
                    Next
            
                    ' 12
                    vx = iStepX - FindStep
                    For FindForY = (iStepY - 1) To ((iStepY - 1) - (FindStep - 2)) Step -1
                        vy = FindForY
                        If vx > StarProbe.ChipCountX Then vx = StarProbe.ChipCountX
                        If vx < 0 Then vx = 0
                        If vy > StarProbe.ChipCountY Then vy = StarProbe.ChipCountY
                        If vy < 0 Then vy = 0
                        If Not bFind And InkRun((vx), (vy)) Then
                            FindX = vx: FindY = vy: bFind = True
                            Exit For
                        End If
                    Next
            
                    ' 13
                    vy = iStepY + FindStep
                    For FindForX = (iStepX + 1) To ((iStepX + 1) + (FindStep - 2))
                        vx = FindForX
                        If vx > StarProbe.ChipCountX Then vx = StarProbe.ChipCountX
                        If vx < 0 Then vx = 0
                        If vy > StarProbe.ChipCountY Then vy = StarProbe.ChipCountY
                        If vy < 0 Then vy = 0
                        If Not bFind And InkRun((vx), (vy)) Then
                            FindX = vx: FindY = vy: bFind = True
                            Exit For
                        End If
                    Next
            
                    ' 14
                    vy = iStepY + FindStep
                    For FindForX = (iStepX - 1) To ((iStepX - 1) - (FindStep - 2)) Step -1
                        vx = FindForX
                        If vx > StarProbe.ChipCountX Then vx = StarProbe.ChipCountX
                        If vx < 0 Then vx = 0
                        If vy > StarProbe.ChipCountY Then vy = StarProbe.ChipCountY
                        If vy < 0 Then vy = 0
                        If Not bFind And InkRun((vx), (vy)) Then
                            FindX = vx: FindY = vy: bFind = True
                            Exit For
                        End If
                    Next
            
                    ' 15
                    vx = iStepX + FindStep
                    For FindForY = (iStepY + 1) To ((iStepY + 1) - (FindStep - 2))
                        vy = FindForY
                        If vx > StarProbe.ChipCountX Then vx = StarProbe.ChipCountX
                        If vx < 0 Then vx = 0
                        If vy > StarProbe.ChipCountY Then vy = StarProbe.ChipCountY
                        If vy < 0 Then vy = 0
                        If Not bFind And InkRun((vx), (vy)) Then
                            FindX = vx: FindY = vy: bFind = True
                            Exit For
                        End If
                    Next
            
                    ' 16
                    vx = iStepX - FindStep
                    For FindForY = (iStepY + 1) To ((iStepY + 1) - (FindStep - 2))
                        vy = FindForY
                        If vx > StarProbe.ChipCountX Then vx = StarProbe.ChipCountX
                        If vx < 0 Then vx = 0
                        If vy > StarProbe.ChipCountY Then vy = StarProbe.ChipCountY
                        If vy < 0 Then vy = 0
                        If Not bFind And InkRun((vx), (vy)) Then
                            FindX = vx: FindY = vy: bFind = True
                            Exit For
                        End If
                    Next
                End If
            
                If bFind Then
                    Exit Do
                Else
                    FindStep = FindStep + 1
                End If
            Loop
        
            If bFind Then
                If FindX >= (StarProbe.ChipCountX / 2) Then
                    iStepX = FindX + 1
                    iStepY = FindY
            
                    xx = FindX - StarProbe.StartChip.x - 1
                    yy = FindY - StarProbe.StartChip.y
                Else
                    iStepX = FindX - 1
                    iStepY = FindY
            
                    xx = FindX - StarProbe.StartChip.x + 1
                    yy = FindY - StarProbe.StartChip.y
                End If
        
                VScroll_Zoom.value = (FindY / (Abs(StarProbe.Max.y) + Abs(StarProbe.Min.y) + 1)) * 1000
                HScroll_Zoom.value = (FindX / (Abs(StarProbe.Max.x) + Abs(StarProbe.Min.x) + 1)) * 1000
            
                StarProbe.CurrentChip.x = xx
                StarProbe.CurrentChip.y = yy
            
                Text5 = StarProbe.CurrentChip.x
                Text6 = StarProbe.CurrentChip.y
            
                Shape_Chip.Top = (FindY * StarProbe.DisplayChipSizeY) - 2
                Shape_Chip.Left = (FindX * StarProbe.DisplayChipSizeX) - 2
            
                Label_ChipPosition = StarProbe.CurrentChip.x & "/" & StarProbe.CurrentChip.y
            
                'Call StarProbe_XY_Moving((xx), (yy))
                '====================================================================================================================================
                If CH_SET = 1 Then
                    Call StarProbe_XY_Moving((xx + NEEDLE_POS_X), (yy + (NEEDLE_POS_Y * 1)))                         '[ 2020.01.20 ] : ink needle position
                ElseIf CH_SET = 2 Then
                    Call StarProbe_XY_Moving((xx + NEEDLE_POS_X), (yy + (NEEDLE_POS_Y * 1)))                         '[ 2020.01.20 ] : ink needle position
                Else
                    Call StarProbe_XY_Moving((xx + NEEDLE_POS_X), (yy + (NEEDLE_POS_Y * 4)))                         '[ 2020.01.20 ] : ink needle position
                End If
                '====================================================================================================================================
        
                bEnd = False
                       
                Do While (Not bEnd)
                    DoEvents
                          
                    If Not StarProbe_Motor_End_check Then
                        bEnd = True
                    Else
                        MsgBox "Motor not end check !", 16, "STAR PROBE"
                        bEnd = True
                    End If
                           
                    If ErrorStop = True Or bEnd = False Then
                        bEnd = True
                    End If
                Loop
                       
                If StarProbe.Ink_After_LeftPort = 1 And InkRun_Center(xx, yy) Then
                    Call StarProbe_Left_Ink_Dot(StarProbe.Ink_LeftPort)
                    Call InkRun_CenterOk((xx), (yy))
                    FindCount = FindCount - 1
                End If
            
                If StarProbe.Ink_After_RightPort = 1 And InkRun_Center(xx, yy) Then
                    Call StarProbe_Right_Ink_Dot(StarProbe.Ink_RightPort)
                    Call InkRun_CenterOk((xx), (yy))
                    FindCount = FindCount - 1
                End If
        
                If StarProbe.Ink_After_LeftPort = 1 Or _
                    StarProbe.Ink_After_RightPort = 1 Then
                    Sleep 30
                End If
       
                StarProbe_WorkDateTime_To = CDate(Date$ & " " & Time$)
                Call StarProbe_WorkDateTime_HMS(StarProbe_WorkDateTime_Total + DateDiff("S", StarProbe_WorkDateTime_From, StarProbe_WorkDateTime_To))
                SSPanel_DateTime.Caption = StarProbe_WorkDateTime.D & " Day " & StarProbe_WorkDateTime.h & ":" & StarProbe_WorkDateTime.M & ":" & StarProbe_WorkDateTime.s
            End If
        Loop
    End If
    ' Ink
    ''''''''''
    
    If Not bStop Then
        If Right(Load_MAP, 2) = "SP" And (UCase(Load_MAP) <> "TEMP.SP") Then
            'hdd
            If SaveDrive = 0 Then
                BMP_file = "C:\data\" & LOT & "\" & LOT & "_" & SP_CNT & "(INK)" & ".PNG"
            Else
                BMP_file = "D:\data\" & LOT & "\" & LOT & "_" & SP_CNT & "(INK)" & ".PNG"
            End If
        Else
            'hdd
            If No_Probe = True Then
                If SaveDrive = 0 Then
                    BMP_file = "C:\data\" & LOT & "\" & LOT & "_" & TT_NO + 1 & "(EDGEINK)" & ".PNG"
                Else
                    BMP_file = "D:\data\" & LOT & "\" & LOT & "_" & TT_NO + 1 & "(EDGEINK)" & ".PNG"
                End If
            Else
                If SaveDrive = 0 Then
                    BMP_file = "C:\data\" & LOT & "\" & LOT & "_" & TT_NO + 1 & "(INK)" & ".PNG"
                Else
                    BMP_file = "D:\data\" & LOT & "\" & LOT & "_" & TT_NO + 1 & "(INK)" & ".PNG"
                End If
            End If
        End If
        Form_StarProbe_MeasureDataSave.Display_View
        SP_CNT = 0
        
        bStarprobe_AfterInk = False
        File_Time_Delay.Reset
        
        Do
            DoEvents
            If File_Time_Delay.Elapsed > 1000 Then Exit Do
        Loop
            
        Call Command_Map_Clear_Click
        Call Command_DisplayWafer_Click
            
        File_Time_Delay.Reset
            
        If DemoMode = 0 Then
            If ACCO = 1 Or Model_Select = 1 Then                            'ACCO only
                MSComm1.Output = "PC" & vbCrLf
            Else
                MSComm1.Output = "PC" & vbLf
            End If
        End If
            
        Do
            DoEvents
            If File_Time_Delay.Elapsed > 100 Then Exit Do
        Loop
        
        File_Time_Delay.Reset
            
        If Slot_Max_Count = File_Count Then
            File_Count = 0
            ETS_Count = 0
            If DemoMode = 0 Then
                If ACCO = 1 Or Model_Select = 1 Then                            'ACCO only
                    MSComm1.Output = "PCP1L0S0R0W/-1" & vbCrLf
                Else
                    MSComm1.Output = "PCP1L0S0R0W/-1" & vbLf
                End If
            End If
        Else
            If DemoMode = 0 Then
                If ACCO = 1 Or Model_Select = 1 Then                            'ACCO only
                    MSComm1.Output = "PCP1L0S1R1W/-1" & vbCrLf
                Else
                    MSComm1.Output = "PCP1L0S1R1W/-1" & vbLf
                End If
            End If
        End If
                        
        Old_Lot = New_Lot
                        
        Do
            DoEvents
            If File_Time_Delay.Elapsed > 100 Then Exit Do
        Loop
        File_Time_Delay.Reset
            
        If DemoMode = 0 Then
            If ACCO = 1 Or Model_Select = 1 Then                            'ACCO only
                
            Else
                MSComm1.Output = ">" & vbLf
            End If
        End If
        Do
            DoEvents
            If File_Time_Delay.Elapsed > 100 Then Exit Do
        Loop
               
        Wafer_Start = False
        
        RESET_DATA
  
        Call StarProbe_Unlode_Wafer_New_Wafer
        
        If Crack_Wafer = True Then
            Check1(3).value = 0
            bStop = True
        End If
        
        If DemoMode = 0 Then
            File_Time_Delay.Reset
            Do
                DoEvents
                If File_Time_Delay.Elapsed > 30000 Then Exit Do     '[ 2017.03.27 ] : wafer unloadingÈÄ delay ¼öÁ¤ 10000->30000
            Loop
            Set File_Time_Delay = Nothing
        End If
        
        '2016.06.14 Å¸ÀÓÅ¬¸®¾îÃß°¡
        StarProbe_WorkDateTime_Total = 0
        StarProbe_WorkDateTime_From = 0
        StarProbe_WorkDateTime_To = 0
        ''''''''''''''''''''''''''''''''''''''''''''''''''''2016.09.22
        If AutoAlign_Flag = False Then
            Check1(3).value = 0
            Check1(3).Enabled = True
            bStop = True
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''
        bStarprobe_AfterInk = False         '2016.06.29
        Check1(1).value = 0                 '2016.09.27 : Ink Start Position Check Box Off
        Ink_Start_Flag = 0                  '2016.09.27 : Ink Start Position Flag Clear
        
        '[ 2020.10.29 ] : ÃøÁ¤ÀÌ ³¡³ª¸é Ä«¼¼Æ®ÀÇ Ã¼Å©¸¦ Ç¥½ÃÇÑ´Ù.
        NOW_NO(val(lblWafer.Caption) - 1) = False

        If AOI_MODE = 1 Then
            '[ 2020.10.29 ] : AOI°ü·Ã ¼öÁ¤     --> ÆÄÀÏÀÌ¸§¹Ù²Ù±â -> ´ÙÀ½ÆÄÀÏ·Îµù -> ¹æÇâÀüÈ¯ -> first die ¼³Á¤
            If UCase(Right(SSPanel2(0).Caption, 3)) = "AOI" Then
                Name AOI_MAP(lblWafer.Caption) As AOI_MAP(lblWafer.Caption) & "1"               'ÆÄÀÏ ÀÌ¸§À» ¹Ù²ãÁØ´Ù. *.aoi --> *.aoi1
                If AutoAlign_Flag = True Then
                    Dim load_no As Integer
                                
                    For II = 0 To 24
                        If NOW_NO(II) = True Then
                            load_no = II
                            Exit For
                        End If
                    Next II
                    
                    'lblWafer.caption°ú ÀÏÄ¡ÇÏ´Â ***.aoiÆÄÀÏÀ» ºÒ·¯¿Â´Ù.
                    'AOIÆÄÀÏÀÌ¸§ÀÌ wafer no¿Í °°Àº °æ¿ì´Â ´Ù½ÃºÎ¸£Áö ¾Ê°í ³Ñ¾î°£´Ù.
                    If load_no <> 0 Then
                        Command_Map_Clear.Enabled = False
            
                        For xx = 0 To StarProbe.ChipCountX                      'map º¯¼ö ÃÊ±âÈ­
                            For yy = 0 To StarProbe.ChipCountY
                                Wafer(xx, yy).flag = False
                                Wafer(xx, yy).FlagBad = False
                                Wafer(xx, yy).MeasureWait = False
                                Wafer(xx, yy).InkDot = False
                                Wafer(xx, yy).ChipMeasure = False
                                Wafer(xx, yy).BIN = 0
                            Next
                        Next
                        Fail_Loop = False
                        Stop_Measure = False
                        Command_Map_Clear.Enabled = True
                        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        Dim map_first As String                                 'Ã³À½ loadÇÒ ÆÄÀÏ ÀÌ¸§ ÁöÁ¤ º¯¼ö
                        For i = 0 To 24
                            If NOW_NO(i) = True Then
                                map_first = AOI_MAP(i + 1)
                                Exit For
                            End If
                        Next i
                        
                        If map_first = "" Then                              '[ 2021.09.17 ] : AOIÆÄÀÏÀÌ ¾ø´Â °æ¿ì , lot endÀÎ °æ¿ì ºüÁ®³ª°£´Ù.
                            Exit Sub
                        End If
                        '============================================== 2015.11.30 : load map name display
                        For i = Len(map_first) To 1 Step -1
                            If Mid(map_first, i, 1) = "\" Then
                                'map_first = UCase(Mid(map_first, i + 1, Len(map_first) - i + 1))
                                SSPanel2(5).Caption = UCase(Mid(map_first, i + 1, Len(map_first) - i + 1))
                                Exit For
                            End If
                        Next i
                        '==============================================
                        
                        
                        bStarprobe_AfterInk = False                             'INK MODE OFF
                        For xxxx = 0 To 8
                            Shape_Mea(xxxx).Visible = False
                        Next xxxx
                        StarProbe.FileName_Map = Load_MAP
                
                        StarProbeTemp = StarProbe
                        
                        Call StarProbe_FileLoad_ControlMap(map_first)
                                                                                                  
                        StarProbe.CountBadDie = StarProbeTemp.CountBadDie
                        StarProbe.CountGoodDie = StarProbeTemp.CountGoodDie
                        StarProbe.CountSkipDie = StarProbeTemp.CountSkipDie
                        StarProbe.CountTotalChip = StarProbeTemp.CountTotalChip
                
                        '[ 2020.03.27 ] : DB°ü·Ã ¼öÁ¤
                        ''''''''''''''''''''''''''''''''''''''''''''''''
                        Map_Total_Backup = StarProbe.CountTotalChip
                        Map_Good_Backup = StarProbe.CountGoodDie
                        Map_Bad_Backup = StarProbe.CountBadDie
                        Map_Skip_Backup = StarProbe.CountSkipDie
                        '''''''''''''''''''''''''''''''''''''''''''''''''
                
                        'Call Command_Map_Clear_Click                '2016.06.02 : map clearÃß°¡
                
                        Call Command_DisplayWafer_Click
                
                        StarProbe.MeasureStepX = StarProbeTemp.MeasureStepX
                        StarProbe.MeasureStepY = StarProbeTemp.MeasureStepY
                
                        StarProbe.Ink_LeftPort = StarProbeTemp.Ink_LeftPort
                        StarProbe.Ink_RightPort = StarProbeTemp.Ink_RightPort
                        StarProbe.LineOk = StarProbeTemp.LineOk
                        StarProbe.RCount = StarProbeTemp.RCount
                        StarProbe.RCount_Sub = StarProbeTemp.RCount_Sub
                        StarProbe.MeasureSleep = StarProbeTemp.MeasureSleep
                        StarProbe.ReMeasure = StarProbeTemp.ReMeasure
                        StarProbe.Ink_After_LeftPort = StarProbeTemp.Ink_After_LeftPort
                        StarProbe.Ink_After_RightPort = StarProbeTemp.Ink_After_RightPort
                        StarProbe.Ink_After_CenterPort = StarProbeTemp.Ink_After_CenterPort
                        StarProbe.LimitArea = StarProbeTemp.LimitArea
                
                        StarProbe.WaferTest = StarProbeTemp.WaferTest
'                        StarProbe.MeasureAll = StarProbeTemp.MeasureAll
                
                        StarProbe.Ink_After = StarProbeTemp.Ink_After
                
                        Me.MousePointer = 0
                        Command2(1).Enabled = True
                                                            
                        Call Command_WaferDirection_Click               '[ 2020.12.17 ] : ¹æÇâ¼³Á¤ Ãß°¡
                        'Call StarProbe_Set_Die_Size(val(StarProbe.ChipSizeX), val(StarProbe.ChipSizeY))
                        
                        '[ 2021.06.25 ] first die
                        StarProbe.StartChip.x = First_X
                        StarProbe.StartChip.y = First_Y
                        
                        Shape_FirstChip.Top = First_Zoom_TOP
                        Shape_FirstChip.Left = First_Zoom_LEFT

                        Shape_OFirstChip.Top = First_Original_TOP
                        Shape_OFirstChip.Left = First_Original_LEFT
                    End If
                End If
            End If
        End If
    End If
    Exit Sub
End Sub

Sub StarProbe_After_Ink_Dot_STT()
    Dim forx As Integer, fory As Integer
    Dim xx As Integer, yy As Integer
    Dim bRight As Boolean
    Dim bEnd As Boolean
    Dim File_Name, SP_File_Name As String
    Dim FindX As Integer, FindY As Integer
    Dim vx As Integer, vy As Integer
    Dim FindStep As Integer
    Dim bFind As Boolean
    Dim FindCount As Long
    Dim FindForX As Integer, FindForY As Integer
    Dim FindInk As Boolean
    Dim iStepX As Integer, iStepY As Integer
    
    '[ 2020.02.07 ] : ink needle position set
    If Option6(1).value = True Then                     'top
        NEEDLE_POS_X = 0
        NEEDLE_POS_Y = -1
    ElseIf Option6(0).value = True Then                 'bottom
        NEEDLE_POS_X = 0
        NEEDLE_POS_Y = 1
    ElseIf Option6(4).value = True Then                 'left
        NEEDLE_POS_X = -1
        NEEDLE_POS_Y = 0
    ElseIf Option6(3).value = True Then                 'center
        NEEDLE_POS_X = 0
        NEEDLE_POS_Y = 0
    ElseIf Option6(2).value = True Then                 'right
        NEEDLE_POS_X = 1
        NEEDLE_POS_Y = 0
    End If
    
    Set File_Time_Delay = New ccrpStopWatch
    
    bStarProbeStart = True
    bRight = True
    
    bStarprobe_AfterInk = True
    
    xx = StarProbe.StartChip.x
    yy = StarProbe.StartChip.y
                 
    '2015.11.06
    StarProbe_WorkDateTime_From = CDate(Date$ & " " & Time$)
    StarProbe_WorkDateTime_To = CDate(Date$ & " " & Time$)
    StarProbe_WorkDateTime_Total = StarProbe_WorkDateTime_Total + DateDiff("S", StarProbe_WorkDateTime_From, StarProbe_WorkDateTime_To)
            
    If DemoMode = 0 Then
        If StarProbe.Ink_After = 1 Or INK_OFF_TEST = True Then
            Call StarProbe_Z_Down
            Sleep 500
        
            Z = StarProbe_Z_Position
        
            If Z <> "D" Then
                MsgBox " Z Down Fail", vbOKOnly
                Exit Sub
            End If
        End If
    End If

    ''''''''''
    ' 2-Position Skip Die Ink
    bRight = True
    If Door1 = 0 Then bStop = False
    
    'If Sample_No_Ink = False Then
    If Sample_No_Ink = False Or Sample_Yes_Ink = True Then          '[ 2021.12.31 ] : Sampling + ink(o)
        If StarProbe.InkStart.x > StarProbe.ChipCountX \ 2 Then            'Wafer ¿ìÃø¿¡ ÁöÁ¤ÇÑ °æ¿ì
            ' 1
            forx = StarProbe.InkStart.x
            For fory = StarProbe.InkStart.y To 0 Step -1
                DoEvents
                Do While Not bStop
                    DoEvents
                    If InkRun((forx), (fory)) Then
                        If StarProbe.Ink_After_CenterPort = 0 Then      'direct ink
                            xx = forx - StarProbe.StartChip.x + IIf(bRight, 1, -1)
                        Else
                            xx = forx - StarProbe.StartChip.x          'after ink
                        End If
                        yy = fory - StarProbe.StartChip.y
                    
                        VScroll_Zoom.value = (fory / (Abs(StarProbe.Max.y) + Abs(StarProbe.Min.y) + 1)) * 1000
                        HScroll_Zoom.value = (forx / (Abs(StarProbe.Max.x) + Abs(StarProbe.Min.x) + 1)) * 1000
                    
                        StarProbe.CurrentChip.x = xx
                        StarProbe.CurrentChip.y = yy
                    
                        Text5 = StarProbe.CurrentChip.x
                        Text6 = StarProbe.CurrentChip.y

                        Shape_Chip.Top = (fory * StarProbe.DisplayChipSizeY) - 2
                        Shape_Chip.Left = (forx * StarProbe.DisplayChipSizeX) - 2
                    
                        Label_ChipPosition = StarProbe.CurrentChip.x & "/" & StarProbe.CurrentChip.y
                    
                        'Call StarProbe_XY_Moving((xx), (yy - 4))
                        '====================================================================================================================================
                        If CH_SET = 1 Then
                            Call StarProbe_XY_Moving((xx + NEEDLE_POS_X), (yy + (NEEDLE_POS_Y * 1)))                         '[ 2020.01.20 ] : ink needle position
                        ElseIf CH_SET = 2 Then
                            Call StarProbe_XY_Moving((xx + NEEDLE_POS_X), (yy + (NEEDLE_POS_Y * 1)))                         '[ 2020.01.20 ] : ink needle position
                        Else
                            Call StarProbe_XY_Moving((xx + NEEDLE_POS_X), (yy + (NEEDLE_POS_Y * 4)))                         '[ 2020.01.20 ] : ink needle position
                        End If
                        '====================================================================================================================================
                                              
                        bEnd = False
                               
                        Do While (Not bEnd)
                            DoEvents
                            If Not StarProbe_Motor_End_check Then
                                bEnd = True
                            Else
                                MsgBox "Motor not end check !", 16, "STAR PROBE"
                                bEnd = True
                            End If
                            If ErrorStop = True Or bEnd = False Then bEnd = True
                        Loop
                               
                        If StarProbe.Ink_After_LeftPort = 1 Or StarProbe.Ink_After_RightPort = 1 Or StarProbe.Ink_After_CenterPort = 1 Then
                            Sleep 10
                        End If
    
                        If bRight And StarProbe.Ink_After_LeftPort = 1 And InkRun_Left(xx, yy) Then
                            Call StarProbe_Left_Ink_Dot(StarProbe.Ink_LeftPort)
                            Call InkRun_LeftOk((xx), (yy))
                            FindCount = FindCount - 1
                            If DemoMode = 1 Then Call Display_Chip_demo(pZoom, pOriginal, xx, yy)     'È­¸é¿¡ Ç¥½Ã
                        End If
                    
                        If Not bRight And StarProbe.Ink_After_RightPort = 1 And InkRun_Right(xx, yy) Then
                            Call StarProbe_Right_Ink_Dot(StarProbe.Ink_RightPort)
                            Call InkRun_RightOk((xx), (yy))
                            FindCount = FindCount - 1
                            If DemoMode = 1 Then Call Display_Chip_demo(pZoom, pOriginal, xx, yy)     'È­¸é¿¡ Ç¥½Ã
                        End If
                    
                        If StarProbe.Ink_After_CenterPort = 1 And InkRun_Center(xx, yy) Then   ' after center
                            Call StarProbe_Left_Ink_Dot(StarProbe.Ink_LeftPort)
                            Call InkRun_CenterOk((xx), (yy))
                            FindCount = FindCount - 1
                            If DemoMode = 1 Then Call Display_Chip_demo(pZoom, pOriginal, xx, yy)     'È­¸é¿¡ Ç¥½Ã
                        End If
                        StarProbe_WorkDateTime_To = CDate(Date$ & " " & Time$)
                        Call StarProbe_WorkDateTime_HMS(StarProbe_WorkDateTime_Total + DateDiff("S", StarProbe_WorkDateTime_From, StarProbe_WorkDateTime_To))
                        SSPanel_DateTime.Caption = StarProbe_WorkDateTime.D & " Day " & StarProbe_WorkDateTime.h & ":" & StarProbe_WorkDateTime.M & ":" & StarProbe_WorkDateTime.s
                    End If
                
                    If bRight Then
                        forx = forx + 1
                        If forx > (StarProbe.ChipCountX + 1) Then
                            forx = (StarProbe.ChipCountX)
                            bRight = False
                            Exit Do
                        End If
                    Else
                        forx = forx - 1
                        If forx < (StarProbe.ChipCountX \ 2) Then
                            forx = (StarProbe.ChipCountX \ 2)
                            bRight = True
                            Exit Do
                        End If
                    End If
                Loop
            Next

            bRight = False

            ' 2
            forx = StarProbe.ChipCountX \ 2
            For fory = 0 To StarProbe.ChipCountY
                DoEvents
                Do While Not bStop
                    DoEvents
                    If InkRun((forx), (fory)) Then
                        If StarProbe.Ink_After_CenterPort = 0 Then
                            xx = forx - StarProbe.StartChip.x + IIf(bRight, 1, -1)
                        Else
                            xx = forx - StarProbe.StartChip.x
                        End If
                        yy = fory - StarProbe.StartChip.y
    
                        VScroll_Zoom.value = (fory / (Abs(StarProbe.Max.y) + Abs(StarProbe.Min.y) + 1)) * 1000
                        HScroll_Zoom.value = (forx / (Abs(StarProbe.Max.x) + Abs(StarProbe.Min.x) + 1)) * 1000
    
                        StarProbe.CurrentChip.x = xx
                        StarProbe.CurrentChip.y = yy
    
                        Text5 = StarProbe.CurrentChip.x
                        Text6 = StarProbe.CurrentChip.y

                        Shape_Chip.Top = (fory * StarProbe.DisplayChipSizeY) - 2
                        Shape_Chip.Left = (forx * StarProbe.DisplayChipSizeX) - 2
    
                        Label_ChipPosition = StarProbe.CurrentChip.x & "/" & StarProbe.CurrentChip.y
    
                        'Call StarProbe_XY_Moving((xx), (yy - 4))
                        '====================================================================================================================================
                        If CH_SET = 1 Then
                            Call StarProbe_XY_Moving((xx + NEEDLE_POS_X), (yy + (NEEDLE_POS_Y * 1)))                         '[ 2020.01.20 ] : ink needle position
                        ElseIf CH_SET = 2 Then
                            Call StarProbe_XY_Moving((xx + NEEDLE_POS_X), (yy + (NEEDLE_POS_Y * 1)))                         '[ 2020.01.20 ] : ink needle position
                        Else
                            Call StarProbe_XY_Moving((xx + NEEDLE_POS_X), (yy + (NEEDLE_POS_Y * 4)))                         '[ 2020.01.20 ] : ink needle position
                        End If
                        '====================================================================================================================================
    
                        bEnd = False
    
                        Do While (Not bEnd)
                            DoEvents
                            If Not StarProbe_Motor_End_check Then
                                bEnd = True
                            Else
                                MsgBox "Motor not end check !", 16, "STAR PROBE"
                                bEnd = True
                            End If
                            If ErrorStop = True Or bEnd = False Then bEnd = True
                        Loop
                        
                        If StarProbe.Ink_After_LeftPort = 1 Or StarProbe.Ink_After_RightPort = 1 Or StarProbe.Ink_After_CenterPort = 1 Then
                            Sleep 10
                        End If
    
                        If bRight And StarProbe.Ink_After_LeftPort = 1 And InkRun_Left(xx, yy) Then
                            Call StarProbe_Left_Ink_Dot(StarProbe.Ink_LeftPort)
                            Call InkRun_LeftOk((xx), (yy))
                            FindCount = FindCount - 1
                            If DemoMode = 1 Then Call Display_Chip_demo(pZoom, pOriginal, xx, yy)     'È­¸é¿¡ Ç¥½Ã
                        End If
    
                        If Not bRight And StarProbe.Ink_After_RightPort = 1 And InkRun_Right(xx, yy) Then
                            Call StarProbe_Right_Ink_Dot(StarProbe.Ink_RightPort)
                            Call InkRun_RightOk((xx), (yy))
                            FindCount = FindCount - 1
                            If DemoMode = 1 Then Call Display_Chip_demo(pZoom, pOriginal, xx, yy)     'È­¸é¿¡ Ç¥½Ã
                        End If
    
                        If StarProbe.Ink_After_CenterPort = 1 And InkRun_Center(xx, yy) Then
                            Call StarProbe_Left_Ink_Dot(StarProbe.Ink_LeftPort)
                            Call InkRun_CenterOk((xx), (yy))
                            FindCount = FindCount - 1
                            If DemoMode = 1 Then Call Display_Chip_demo(pZoom, pOriginal, xx, yy)     'È­¸é¿¡ Ç¥½Ã
                        End If
                        StarProbe_WorkDateTime_To = CDate(Date$ & " " & Time$)
                        Call StarProbe_WorkDateTime_HMS(StarProbe_WorkDateTime_Total + DateDiff("S", StarProbe_WorkDateTime_From, StarProbe_WorkDateTime_To))
                        SSPanel_DateTime.Caption = StarProbe_WorkDateTime.D & " Day " & StarProbe_WorkDateTime.h & ":" & StarProbe_WorkDateTime.M & ":" & StarProbe_WorkDateTime.s
                    End If
    
                    If bRight Then
                        forx = forx + 1
                        If forx > StarProbe.ChipCountX \ 2 Then
                            forx = (StarProbe.ChipCountX \ 2)
                            bRight = False
                            Exit Do
                        End If
                    Else
                        forx = forx - 1
                        If forx < -1 Then
                            forx = 0
                            bRight = True
                            Exit Do
                        End If
                    End If
                Loop
            Next
        Else
            bRight = False
        
            ' 2
            forx = StarProbe.InkStart.x
            For fory = StarProbe.InkStart.y To StarProbe.ChipCountY
                DoEvents
                Do While Not bStop
                    DoEvents
                    If InkRun((forx), (fory)) Then
                        If StarProbe.Ink_After_CenterPort = 0 Then
                            xx = forx - StarProbe.StartChip.x + IIf(bRight, 1, -1)
                        Else
                            xx = forx - StarProbe.StartChip.x
                        End If
                        yy = fory - StarProbe.StartChip.y
    
                        VScroll_Zoom.value = (fory / (Abs(StarProbe.Max.y) + Abs(StarProbe.Min.y) + 1)) * 1000
                        HScroll_Zoom.value = (forx / (Abs(StarProbe.Max.x) + Abs(StarProbe.Min.x) + 1)) * 1000
    
                        StarProbe.CurrentChip.x = xx
                        StarProbe.CurrentChip.y = yy
    
                        Text5 = StarProbe.CurrentChip.x
                        Text6 = StarProbe.CurrentChip.y

                        Shape_Chip.Top = (fory * StarProbe.DisplayChipSizeY) - 2
                        Shape_Chip.Left = (forx * StarProbe.DisplayChipSizeX) - 2
    
                        Label_ChipPosition = StarProbe.CurrentChip.x & "/" & StarProbe.CurrentChip.y
    
                        'Call StarProbe_XY_Moving((xx), (yy - 4))
                        '====================================================================================================================================
                        If CH_SET = 1 Then
                            Call StarProbe_XY_Moving((xx + NEEDLE_POS_X), (yy + (NEEDLE_POS_Y * 1)))                         '[ 2020.01.20 ] : ink needle position
                        ElseIf CH_SET = 2 Then
                            Call StarProbe_XY_Moving((xx + NEEDLE_POS_X), (yy + (NEEDLE_POS_Y * 1)))                         '[ 2020.01.20 ] : ink needle position
                        Else
                            Call StarProbe_XY_Moving((xx + NEEDLE_POS_X), (yy + (NEEDLE_POS_Y * 4)))                         '[ 2020.01.20 ] : ink needle position
                        End If
                        '====================================================================================================================================
    
                        bEnd = False
    
                        Do While (Not bEnd)
                            DoEvents
                            If Not StarProbe_Motor_End_check Then
                                bEnd = True
                            Else
                                MsgBox "Motor not end check !", 16, "STAR PROBE"
                                bEnd = True
                            End If
                            If ErrorStop = True Or bEnd = False Then bEnd = True
                        Loop
                        
                        If StarProbe.Ink_After_LeftPort = 1 Or StarProbe.Ink_After_RightPort = 1 Or StarProbe.Ink_After_CenterPort = 1 Then
                            Sleep 10
                        End If
    
                        If bRight And StarProbe.Ink_After_LeftPort = 1 And InkRun_Left(xx, yy) Then
                            Call StarProbe_Left_Ink_Dot(StarProbe.Ink_LeftPort)
                            Call InkRun_LeftOk((xx), (yy))
                            FindCount = FindCount - 1
                            If DemoMode = 1 Then Call Display_Chip_demo(pZoom, pOriginal, xx, yy)     'È­¸é¿¡ Ç¥½Ã
                        End If
    
                        If Not bRight And StarProbe.Ink_After_RightPort = 1 And InkRun_Right(xx, yy) Then
                            Call StarProbe_Right_Ink_Dot(StarProbe.Ink_RightPort)
                            Call InkRun_RightOk((xx), (yy))
                            FindCount = FindCount - 1
                            If DemoMode = 1 Then Call Display_Chip_demo(pZoom, pOriginal, xx, yy)     'È­¸é¿¡ Ç¥½Ã
                        End If
    
                        If StarProbe.Ink_After_CenterPort = 1 And InkRun_Center(xx, yy) Then
                            Call StarProbe_Left_Ink_Dot(StarProbe.Ink_LeftPort)
                            Call InkRun_CenterOk((xx), (yy))
                            FindCount = FindCount - 1
                            If DemoMode = 1 Then Call Display_Chip_demo(pZoom, pOriginal, xx, yy)     'È­¸é¿¡ Ç¥½Ã
                        End If
                        StarProbe_WorkDateTime_To = CDate(Date$ & " " & Time$)
                        Call StarProbe_WorkDateTime_HMS(StarProbe_WorkDateTime_Total + DateDiff("S", StarProbe_WorkDateTime_From, StarProbe_WorkDateTime_To))
                        SSPanel_DateTime.Caption = StarProbe_WorkDateTime.D & " Day " & StarProbe_WorkDateTime.h & ":" & StarProbe_WorkDateTime.M & ":" & StarProbe_WorkDateTime.s
                    End If
    
                    If bRight Then
                        forx = forx + 1
                        If forx > StarProbe.ChipCountX \ 2 Then
                            forx = (StarProbe.ChipCountX \ 2)
                            bRight = False
                            Exit Do
                        End If
                    Else
                        forx = forx - 1
                        If forx < -1 Then
                            forx = 0
                            bRight = True
                            Exit Do
                        End If
                    End If
                Loop
            Next
        End If

' 2-Position Skip Die Ink
''''''''''

''''''''''
' Ink

        FindCount = 0
    
        For fory = 0 To StarProbe.ChipCountY
            For forx = 0 To StarProbe.ChipCountX
                If InkRun((forx), (fory)) Then FindCount = FindCount + 1
            Next
        Next
    
        iStepX = StarProbe.CurrentChip.x + StarProbe.StartChip.x
        iStepY = StarProbe.CurrentChip.y + StarProbe.StartChip.y

        FindStep = 1
    
        Do While Not bStop
            DoEvents
            If FindCount <= 0 Then Exit Do
            Exit Do
            
            bFind = False
        
            If InkRun_Center((iStepX), (iStepY)) Then
                FindX = iStepX
                FindY = iStepY
                bFind = True
            End If
        
            FindStep = 1
        
            Do While Not bStop
                DoEvents
                If bFind Then Exit Do
            
                ' 2
                vx = iStepX + FindStep
                vy = iStepY
                If vx > StarProbe.ChipCountX Then vx = StarProbe.ChipCountX
                If vx < 0 Then vx = 0
                If vy > StarProbe.ChipCountY Then vy = StarProbe.ChipCountY
                If vy < 0 Then vy = 0
            
                If Not bFind And InkRun((vx), (vy)) Then
                    FindX = vx: FindY = vy: bFind = True
                End If
            
                ' 3
                vx = iStepX - FindStep
                vy = iStepY
                If vx > StarProbe.ChipCountX Then vx = StarProbe.ChipCountX
                If vx < 0 Then vx = 0
                If vy > StarProbe.ChipCountY Then vy = StarProbe.ChipCountY
                If vy < 0 Then vy = 0
                If Not bFind And InkRun_Center((vx), (vy)) Then
                    FindX = vx: FindY = vy: bFind = True
                End If
            
                ' 1
                vx = iStepX
                vy = iStepY - FindStep
                If vx > StarProbe.ChipCountX Then vx = StarProbe.ChipCountX
                If vx < 0 Then vx = 0
                If vy > StarProbe.ChipCountY Then vy = StarProbe.ChipCountY
                If vy < 0 Then vy = 0
                If Not bFind And InkRun_Center((vx), (vy)) Then
                    FindX = vx: FindY = vy: bFind = True
                End If
                
                ' 4
                vx = iStepX
                vy = iStepY + FindStep
                If vx > StarProbe.ChipCountX Then vx = StarProbe.ChipCountX
                If vx < 0 Then vx = 0
                If vy > StarProbe.ChipCountY Then vy = StarProbe.ChipCountY
                If vy < 0 Then vy = 0
                If Not bFind And InkRun((vx), (vy)) Then
                    FindX = vx: FindY = vy: bFind = True
                End If
                
                ' 5
                vx = iStepX + FindStep
                vy = iStepY - FindStep
                If vx > StarProbe.ChipCountX Then vx = StarProbe.ChipCountX
                If vx < 0 Then vx = 0
                If vy > StarProbe.ChipCountY Then vy = StarProbe.ChipCountY
                If vy < 0 Then vy = 0
                If Not bFind And InkRun((vx), (vy)) Then
                    FindX = vx: FindY = vy: bFind = True
                End If
                
                ' 6
                vx = iStepX - FindStep
                vy = iStepY - FindStep
                If vx > StarProbe.ChipCountX Then vx = StarProbe.ChipCountX
                If vx < 0 Then vx = 0
                If vy > StarProbe.ChipCountY Then vy = StarProbe.ChipCountY
                If vy < 0 Then vy = 0
                If Not bFind And InkRun((vx), (vy)) Then
                    FindX = vx: FindY = vy: bFind = True
                End If
                
                ' 7
                vx = iStepX + FindStep
                vy = iStepY + FindStep
                If vx > StarProbe.ChipCountX Then vx = StarProbe.ChipCountX
                If vx < 0 Then vx = 0
                If vy > StarProbe.ChipCountY Then vy = StarProbe.ChipCountY
                If vy < 0 Then vy = 0
                If Not bFind And InkRun((vx), (vy)) Then
                    FindX = vx: FindY = vy: bFind = True
                End If
                
                ' 8
                vx = iStepX - FindStep
                vy = iStepY + FindStep
                If vx > StarProbe.ChipCountX Then vx = StarProbe.ChipCountX
                If vx < 0 Then vx = 0
                If vy > StarProbe.ChipCountY Then vy = StarProbe.ChipCountY
                If vy < 0 Then vy = 0
                If Not bFind And InkRun((vx), (vy)) Then
                    FindX = vx: FindY = vy: bFind = True
                End If
                
                If Not bFind And FindStep > 1 Then
                    ' 9
                    vy = iStepY - FindStep
                    For FindForX = (iStepX + 1) To ((iStepX + 1) + (FindStep - 2))
                        vx = FindForX
                        If vx > StarProbe.ChipCountX Then vx = StarProbe.ChipCountX
                        If vx < 0 Then vx = 0
                        If vy > StarProbe.ChipCountY Then vy = StarProbe.ChipCountY
                        If vy < 0 Then vy = 0
                        If Not bFind And InkRun((vx), (vy)) Then
                            FindX = vx: FindY = vy: bFind = True
                            Exit For
                        End If
                    Next
            
                    ' 10
                    vy = iStepY - FindStep
                    For FindForX = (iStepX - 1) To ((iStepX - 1) - (FindStep - 2)) Step -1
                        vx = FindForX
                        If vx > StarProbe.ChipCountX Then vx = StarProbe.ChipCountX
                        If vx < 0 Then vx = 0
                        If vy > StarProbe.ChipCountY Then vy = StarProbe.ChipCountY
                        If vy < 0 Then vy = 0
                        If Not bFind And InkRun((vx), (vy)) Then
                            FindX = vx: FindY = vy: bFind = True
                            Exit For
                        End If
                    Next
            
                    ' 11
                    vx = iStepX + FindStep
                    For FindForY = (iStepY - 1) To ((iStepY - 1) - (FindStep - 2)) Step -1
                        vy = FindForY
                        If vx > StarProbe.ChipCountX Then vx = StarProbe.ChipCountX
                        If vx < 0 Then vx = 0
                        If vy > StarProbe.ChipCountY Then vy = StarProbe.ChipCountY
                        If vy < 0 Then vy = 0
                        If Not bFind And InkRun((vx), (vy)) Then
                            FindX = vx: FindY = vy: bFind = True
                            Exit For
                        End If
                    Next
            
                    ' 12
                    vx = iStepX - FindStep
                    For FindForY = (iStepY - 1) To ((iStepY - 1) - (FindStep - 2)) Step -1
                        vy = FindForY
                        If vx > StarProbe.ChipCountX Then vx = StarProbe.ChipCountX
                        If vx < 0 Then vx = 0
                        If vy > StarProbe.ChipCountY Then vy = StarProbe.ChipCountY
                        If vy < 0 Then vy = 0
                        If Not bFind And InkRun((vx), (vy)) Then
                            FindX = vx: FindY = vy: bFind = True
                            Exit For
                        End If
                    Next
            
                    ' 13
                    vy = iStepY + FindStep
                    For FindForX = (iStepX + 1) To ((iStepX + 1) + (FindStep - 2))
                        vx = FindForX
                        If vx > StarProbe.ChipCountX Then vx = StarProbe.ChipCountX
                        If vx < 0 Then vx = 0
                        If vy > StarProbe.ChipCountY Then vy = StarProbe.ChipCountY
                        If vy < 0 Then vy = 0
                        If Not bFind And InkRun((vx), (vy)) Then
                            FindX = vx: FindY = vy: bFind = True
                            Exit For
                        End If
                    Next
            
                    ' 14
                    vy = iStepY + FindStep
                    For FindForX = (iStepX - 1) To ((iStepX - 1) - (FindStep - 2)) Step -1
                        vx = FindForX
                        If vx > StarProbe.ChipCountX Then vx = StarProbe.ChipCountX
                        If vx < 0 Then vx = 0
                        If vy > StarProbe.ChipCountY Then vy = StarProbe.ChipCountY
                        If vy < 0 Then vy = 0
                        If Not bFind And InkRun((vx), (vy)) Then
                            FindX = vx: FindY = vy: bFind = True
                            Exit For
                        End If
                    Next
            
                    ' 15
                    vx = iStepX + FindStep
                    For FindForY = (iStepY + 1) To ((iStepY + 1) - (FindStep - 2))
                        vy = FindForY
                        If vx > StarProbe.ChipCountX Then vx = StarProbe.ChipCountX
                        If vx < 0 Then vx = 0
                        If vy > StarProbe.ChipCountY Then vy = StarProbe.ChipCountY
                        If vy < 0 Then vy = 0
                        If Not bFind And InkRun((vx), (vy)) Then
                            FindX = vx: FindY = vy: bFind = True
                            Exit For
                        End If
                    Next
            
                    ' 16
                    vx = iStepX - FindStep
                    For FindForY = (iStepY + 1) To ((iStepY + 1) - (FindStep - 2))
                        vy = FindForY
                        If vx > StarProbe.ChipCountX Then vx = StarProbe.ChipCountX
                        If vx < 0 Then vx = 0
                        If vy > StarProbe.ChipCountY Then vy = StarProbe.ChipCountY
                        If vy < 0 Then vy = 0
                        If Not bFind And InkRun((vx), (vy)) Then
                            FindX = vx: FindY = vy: bFind = True
                            Exit For
                        End If
                    Next
                End If
            
                If bFind Then
                    Exit Do
                Else
                    FindStep = FindStep + 1
                End If
            Loop
        
            If bFind Then
                If FindX >= (StarProbe.ChipCountX / 2) Then
                    iStepX = FindX + 1
                    iStepY = FindY
            
                    xx = FindX - StarProbe.StartChip.x - 1
                    yy = FindY - StarProbe.StartChip.y
                Else
                    iStepX = FindX - 1
                    iStepY = FindY
                
                    xx = FindX - StarProbe.StartChip.x + 1
                    yy = FindY - StarProbe.StartChip.y
                End If
        
                VScroll_Zoom.value = (FindY / (Abs(StarProbe.Max.y) + Abs(StarProbe.Min.y) + 1)) * 1000
                HScroll_Zoom.value = (FindX / (Abs(StarProbe.Max.x) + Abs(StarProbe.Min.x) + 1)) * 1000
            
                StarProbe.CurrentChip.x = xx
                StarProbe.CurrentChip.y = yy
            
                Text5 = StarProbe.CurrentChip.x
                Text6 = StarProbe.CurrentChip.y
            
                Shape_Chip.Top = (FindY * StarProbe.DisplayChipSizeY) - 2
                Shape_Chip.Left = (FindX * StarProbe.DisplayChipSizeX) - 2
            
                Label_ChipPosition = StarProbe.CurrentChip.x & "/" & StarProbe.CurrentChip.y
            
                'Call StarProbe_XY_Moving((xx), (yy))
                '====================================================================================================================================
                If CH_SET = 1 Then
                    Call StarProbe_XY_Moving((xx + NEEDLE_POS_X), (yy + (NEEDLE_POS_Y * 1)))                         '[ 2020.01.20 ] : ink needle position
                ElseIf CH_SET = 2 Then
                    Call StarProbe_XY_Moving((xx + NEEDLE_POS_X), (yy + (NEEDLE_POS_Y * 1)))                         '[ 2020.01.20 ] : ink needle position
                Else
                    Call StarProbe_XY_Moving((xx + NEEDLE_POS_X), (yy + (NEEDLE_POS_Y * 4)))                         '[ 2020.01.20 ] : ink needle position
                End If
                '====================================================================================================================================
        
                bEnd = False
                       
                Do While (Not bEnd)
                    DoEvents
                    If Not StarProbe_Motor_End_check Then
                        bEnd = True
                    Else
                        MsgBox "Motor not end check !", 16, "STAR PROBE"
                        bEnd = True
                    End If
                    If ErrorStop = True Or bEnd = False Then bEnd = True
                Loop
                       
                If StarProbe.Ink_After_LeftPort = 1 And InkRun_Center(xx, yy) Then
                    Call StarProbe_Left_Ink_Dot(StarProbe.Ink_LeftPort)
                    Call InkRun_CenterOk((xx), (yy))
                    FindCount = FindCount - 1
                End If
            
                If StarProbe.Ink_After_RightPort = 1 And InkRun_Center(xx, yy) Then
                    Call StarProbe_Right_Ink_Dot(StarProbe.Ink_RightPort)
                    Call InkRun_CenterOk((xx), (yy))
                    FindCount = FindCount - 1
                End If
        
                If StarProbe.Ink_After_LeftPort = 1 Or _
                    StarProbe.Ink_After_RightPort = 1 Then
                    Sleep 30
                End If
           
                StarProbe_WorkDateTime_To = CDate(Date$ & " " & Time$)
                Call StarProbe_WorkDateTime_HMS(StarProbe_WorkDateTime_Total + DateDiff("S", StarProbe_WorkDateTime_From, StarProbe_WorkDateTime_To))
                SSPanel_DateTime.Caption = StarProbe_WorkDateTime.D & " Day " & StarProbe_WorkDateTime.h & ":" & StarProbe_WorkDateTime.M & ":" & StarProbe_WorkDateTime.s
            End If
        Loop
    End If
' Ink
''''''''''
    
    If Not bStop Then
        bStarprobe_AfterInk = False
        
        If Right(Load_MAP, 2) = "SP" And (UCase(Load_MAP) <> "TEMP.SP") Then
            'hdd
            If SaveDrive = 0 Then
                BMP_file = "C:\data\" & LOT & "\" & LOT & "_" & SP_CNT & "(INK)" & "PNG"
            Else
                BMP_file = "D:\data\" & LOT & "\" & LOT & "_" & SP_CNT & "(INK)" & "PNG"
            End If
        Else
            'hdd
            If No_Probe = True Then
                If SaveDrive = 0 Then
                    BMP_file = "C:\data\" & LOT & "\" & LOT & "_" & TT_NO + 1 & "(EDGEINK)" & "PNG"
                Else
                    BMP_file = "D:\data\" & LOT & "\" & LOT & "_" & TT_NO + 1 & "(EDGEINK)" & "PNG"
                End If
            Else
                If SaveDrive = 0 Then
                    BMP_file = "C:\data\" & LOT & "\" & LOT & "_" & TT_NO + 1 & "(INK)" & "PNG"
                Else
                    BMP_file = "D:\data\" & LOT & "\" & LOT & "_" & TT_NO + 1 & "(INK)" & "PNG"
                End If
            End If
        End If
        Form_StarProbe_MeasureDataSave.Display_View
        SP_CNT = 0
        
        Do
            DoEvents
            If File_Time_Delay.Elapsed > 1000 Then Exit Do
        Loop
            
        Call Command_Map_Clear_Click
        Call Command_DisplayWafer_Click
            
        File_Time_Delay.Reset
            
        If DemoMode = 0 Then
            If ACCO = 1 Or Model_Select = 1 Then                            'ACCO only
                MSComm1.Output = "PC" & vbCrLf
            Else
                MSComm1.Output = "PC" & vbLf
            End If
        End If
            
        Do
            DoEvents
            If File_Time_Delay.Elapsed > 100 Then Exit Do
        Loop
        
        File_Time_Delay.Reset
            
        If Slot_Max_Count = File_Count Then
            File_Count = 0
            ETS_Count = 0
        End If
        If DemoMode = 0 Then
            If ACCO = 1 Or Model_Select = 1 Then                            'ACCO only
                MSComm1.Output = "PCP1L0S0R0W/-1" & vbCrLf
            Else
                MSComm1.Output = "PCP1L0S0R0W/-1" & vbLf
            End If
        End If
                        
        Old_Lot = New_Lot
                        
        Do
            DoEvents
            If File_Time_Delay.Elapsed > 100 Then Exit Do
        Loop
        File_Time_Delay.Reset
            
        If DemoMode = 0 Then
            If ACCO = 1 Or Model_Select = 1 Then                            'ACCO only
            Else
                MSComm1.Output = ">" & vbLf
            End If
        End If
        File_Time_Delay.Reset
        Do
            DoEvents
            If File_Time_Delay.Elapsed > 100 Then Exit Do
        Loop
               
        Wafer_Start = False
        
        RESET_DATA
  
        Call StarProbe_Unlode_Wafer_New_Wafer
        If AutoAlign_Flag = False Then
            Check1(3).value = 0
            bStop = True
            bStarprobe_AfterInk = False
        End If
        Command7.Enabled = True
         
        File_Time_Delay.Reset
        Do
            DoEvents
            If File_Time_Delay.Elapsed > 30000 Then Exit Do     '[ 2017.03.27 ] : wafer unloadingÈÄ delay ¼öÁ¤ 10000->30000
        Loop
        Set File_Time_Delay = Nothing
        
        '2016.06.14 Å¸ÀÓÅ¬¸®¾îÃß°¡
        StarProbe_WorkDateTime_Total = 0
        StarProbe_WorkDateTime_From = 0
        StarProbe_WorkDateTime_To = 0
        ''''''''''''''''''''''''''''''''''''''''''''''''''''2019.09.22
        If AutoAlign_Flag = False Then
            Check1(3).value = 0
            Check1(3).Enabled = True
            bStop = True
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''
        bStarprobe_AfterInk = False         '2016.06.29
        Check1(1).value = 0                 '2016.09.27 : Ink Start Position Check Box Off
        Ink_Start_Flag = 0                  '2016.09.27 : Ink Start Position Flag Clear
        
        '[ 2020.10.29 ] : ÃøÁ¤ÀÌ ³¡³ª¸é Ä«¼¼Æ®ÀÇ Ã¼Å©¸¦ Ç¥½ÃÇÑ´Ù.
'        Form_Cassette_New.Check1(lblWafer.Caption - 1).value = 1
        NOW_NO(val(lblWafer.Caption) - 1) = False

        If AOI_MODE = 1 Then
            '[ 2020.10.29 ] : AOI°ü·Ã ¼öÁ¤     --> ÆÄÀÏÀÌ¸§¹Ù²Ù±â -> ´ÙÀ½ÆÄÀÏ·Îµù -> ¹æÇâÀüÈ¯ -> first die ¼³Á¤
            If UCase(Right(SSPanel2(0).Caption, 3)) = "AOI" Then
                Name AOI_MAP(lblWafer.Caption) As AOI_MAP(lblWafer.Caption) & "1"               'ÆÄÀÏ ÀÌ¸§À» ¹Ù²ãÁØ´Ù. *.aoi --> *.aoi1
                If AutoAlign_Flag = True Then
                    Dim load_no As Integer
                                
                    For II = 0 To 24
                        If NOW_NO(II) = True Then
                            load_no = II
                            Exit For
                        End If
                    Next II
                    
                    'lblWafer.caption°ú ÀÏÄ¡ÇÏ´Â ***.aoiÆÄÀÏÀ» ºÒ·¯¿Â´Ù.
                    'AOIÆÄÀÏÀÌ¸§ÀÌ wafer no¿Í °°Àº °æ¿ì´Â ´Ù½ÃºÎ¸£Áö ¾Ê°í ³Ñ¾î°£´Ù.
                    If load_no <> 0 Then
                        Command_Map_Clear.Enabled = False
            
                        For xx = 0 To StarProbe.ChipCountX                      'map º¯¼ö ÃÊ±âÈ­
                            For yy = 0 To StarProbe.ChipCountY
                                Wafer(xx, yy).flag = False
                                Wafer(xx, yy).FlagBad = False
                                Wafer(xx, yy).MeasureWait = False
                                Wafer(xx, yy).InkDot = False
                                Wafer(xx, yy).ChipMeasure = False
                                Wafer(xx, yy).BIN = 0
                            Next
                        Next
                        Fail_Loop = False
                        Stop_Measure = False
                        Command_Map_Clear.Enabled = True
                        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        Dim map_first As String                                 'Ã³À½ loadÇÒ ÆÄÀÏ ÀÌ¸§ ÁöÁ¤ º¯¼ö
                        For i = 0 To 24
                            If NOW_NO(i) = True Then
                                map_first = AOI_MAP(i + 1)
                                Exit For
                            End If
                        Next i
                        
                        If map_first = "" Then                              '[ 2021.09.17 ] : AOIÆÄÀÏÀÌ ¾ø´Â °æ¿ì , lot endÀÎ °æ¿ì ºüÁ®³ª°£´Ù.
                            Exit Sub
                        End If
                        
                        '============================================== 2015.11.30 : load map name display
                        For i = Len(map_first) To 1 Step -1
                            If Mid(map_first, i, 1) = "\" Then
                                'map_first = UCase(Mid(map_first, i + 1, Len(map_first) - i + 1))
                                SSPanel2(5).Caption = UCase(Mid(map_first, i + 1, Len(map_first) - i + 1))
                                Exit For
                            End If
                        Next i
                        '==============================================
                        
                        
                        bStarprobe_AfterInk = False                             'INK MODE OFF
                        For xxxx = 0 To 8
                            Shape_Mea(xxxx).Visible = False
                        Next xxxx
                        StarProbe.FileName_Map = Load_MAP
                
                        StarProbeTemp = StarProbe
                        
                        Call StarProbe_FileLoad_ControlMap(map_first)
                                                                                                  
                        StarProbe.CountBadDie = StarProbeTemp.CountBadDie
                        StarProbe.CountGoodDie = StarProbeTemp.CountGoodDie
                        StarProbe.CountSkipDie = StarProbeTemp.CountSkipDie
                        StarProbe.CountTotalChip = StarProbeTemp.CountTotalChip
                
                        '[ 2020.03.27 ] : DB°ü·Ã ¼öÁ¤
                        ''''''''''''''''''''''''''''''''''''''''''''''''
                        Map_Total_Backup = StarProbe.CountTotalChip
                        Map_Good_Backup = StarProbe.CountGoodDie
                        Map_Bad_Backup = StarProbe.CountBadDie
                        Map_Skip_Backup = StarProbe.CountSkipDie
                        '''''''''''''''''''''''''''''''''''''''''''''''''
                
                        'Call Command_Map_Clear_Click                '2016.06.02 : map clearÃß°¡
                
                        Call Command_DisplayWafer_Click
                
                        StarProbe.MeasureStepX = StarProbeTemp.MeasureStepX
                        StarProbe.MeasureStepY = StarProbeTemp.MeasureStepY
                
                        StarProbe.Ink_LeftPort = StarProbeTemp.Ink_LeftPort
                        StarProbe.Ink_RightPort = StarProbeTemp.Ink_RightPort
                        StarProbe.LineOk = StarProbeTemp.LineOk
                        StarProbe.RCount = StarProbeTemp.RCount
                        StarProbe.RCount_Sub = StarProbeTemp.RCount_Sub
                        StarProbe.MeasureSleep = StarProbeTemp.MeasureSleep
                        StarProbe.ReMeasure = StarProbeTemp.ReMeasure
                        StarProbe.Ink_After_LeftPort = StarProbeTemp.Ink_After_LeftPort
                        StarProbe.Ink_After_RightPort = StarProbeTemp.Ink_After_RightPort
                        StarProbe.Ink_After_CenterPort = StarProbeTemp.Ink_After_CenterPort
                        StarProbe.LimitArea = StarProbeTemp.LimitArea
                
                        StarProbe.WaferTest = StarProbeTemp.WaferTest
'                        StarProbe.MeasureAll = StarProbeTemp.MeasureAll
                
                        StarProbe.Ink_After = StarProbeTemp.Ink_After
                
                        Me.MousePointer = 0
                        Command2(1).Enabled = True
                
                        'Call StarProbe_Set_Die_Size(val(StarProbe.ChipSizeX), val(StarProbe.ChipSizeY))
                        
                        Call Command_WaferDirection_Click               '[ 2020.12.17 ] : ¹æÇâ¼³Á¤ Ãß°¡
                        
                        '[ 2021.06.25 ] first die
                        StarProbe.StartChip.x = First_X
                        StarProbe.StartChip.y = First_Y
                        
                        Shape_FirstChip.Top = First_Zoom_TOP
                        Shape_FirstChip.Left = First_Zoom_LEFT

                        Shape_OFirstChip.Top = First_Original_TOP
                        Shape_OFirstChip.Left = First_Original_LEFT
                    End If
                End If
            End If
        End If
    End If
    Exit Sub
End Sub

Private Sub Command_Map_Clear_Click()
    Dim xx As Integer, yy As Integer
    
    Command_Map_Clear.Enabled = False
    
    For xx = 0 To StarProbe.ChipCountX
        For yy = 0 To StarProbe.ChipCountY
            Wafer(xx, yy).flag = False
            Wafer(xx, yy).FlagBad = False
            Wafer(xx, yy).MeasureWait = False
            Wafer(xx, yy).InkDot = False
            Wafer(xx, yy).ChipMeasure = False ' 2005.09.05
            Wafer(xx, yy).BIN = 0             '16.12.02
        Next
    Next
    Stop_Measure = False
    
    YOON_CNT = 0
    RESET_DATA
    Command_Map_Clear.Enabled = True
    TESTING_flag = False
    Needle_Chk_Ok = False                       '[ 2022.07.29 ]
    bStarprobe_AfterInk = False                 '2016.06.29
End Sub

Private Sub Command_MapMove_Click()
    Dim forx As Integer, fory As Integer
    Dim inputx As Integer, inputy As Integer
    
    Command_MapMove.Enabled = False
    
    inputx = val(InputBox("X", "Map Shift", 0))
    inputy = val(InputBox("Y", "Map Shift", 0))
    
    If inputx = 0 And inputy = 0 Then
        Command_MapMove.Enabled = True
        Exit Sub
    End If
    
    Erase WaferTemp
    
    For fory = 0 To StarProbe.ChipCountY
        For forx = 0 To StarProbe.ChipCountX
            If Wafer(forx, fory).Chip Then
                WaferTemp(forx, fory).Chip = True
                If Wafer(forx, fory).ChipPlate Then
                    WaferTemp(forx, fory).ChipPlate = True
                Else
                    WaferTemp(forx, fory).ChipSkipDie = True
                End If
            End If
        Next
    Next
    
    For fory = 0 To StarProbe.ChipCountY
        For forx = 0 To StarProbe.ChipCountX
            If Wafer(forx, fory).Chip And _
               Not Wafer(forx, fory).ChipPlate And _
               Not Wafer(forx, fory).ChipSkipDie Then
                WaferTemp(forx + inputx, fory + inputy) = Wafer(forx, fory)
            End If
        Next
    Next
    
    Erase Wafer
        
    For forx = 0 To StarProbe.ChipCountX
        For fory = 0 To StarProbe.ChipCountY
            Wafer(forx, fory) = WaferTemp(forx, fory)
        Next
    Next
    
    StarProbe.StartChip.x = StarProbe.StartChip.x + inputx
    StarProbe.StartChip.y = StarProbe.StartChip.y + inputy
    
    Call Command_DisplayWafer_Click
    
    Command_MapMove.Enabled = True
End Sub

Private Sub Command_MaskMove_Click()
    Command_MaskMove.Enabled = False
    
    Dim forx As Integer, fory As Integer
    Dim inputx As Integer, inputy As Integer
    
    inputx = val(InputBox("X", "Mask Shift", 0))
    inputy = val(InputBox("Y", "Mask Shift", 0))
    
    If inputx = 0 And inputy = 0 Then
        Command_MaskMove.Enabled = True
        Exit Sub
    End If
    
    Erase WaferTemp
    
'Type tWafer
'    Chip As Boolean
'    ChipMask As Boolean
'    ChipMeasure As Boolean
'    ChipSkipDie As Boolean
'    ChipPlate As Boolean
'    ChipInk As Boolean
'    BIN As Byte
'    flag As Boolean         ' ÃøÁ¤ ¿©ºÎ
'    FlagBad As Boolean      ' ÃøÁ¤ ÈÄ °á°ú ¾çÇ°ÀÌ¸é False, ºÒ·®ÀÌ¸é True
'    MeasureWait As Boolean  ' ÃøÁ¤ ÃßÀû ¾Ë°í¸®Áò ÁÙÀ» ¼¼¿ï¶§ÀÇ ÇÃ·¡±×
'    InkDot As Boolean
'End Type
    
    For fory = 0 To StarProbe.ChipCountY
        For forx = 0 To StarProbe.ChipCountX
            If Wafer(forx, fory).Chip And Wafer(forx, fory).ChipMask Then
                WaferTemp(forx, fory).Chip = True
            Else
                WaferTemp(forx, fory) = Wafer(forx, fory)
            End If
        Next
    Next
    
    For fory = 0 To StarProbe.ChipCountY
        For forx = 0 To StarProbe.ChipCountX
            If Wafer(forx, fory).Chip And Wafer(forx, fory).ChipMask Then
                WaferTemp(forx + inputx, fory + inputy) = Wafer(forx, fory)
            End If
        Next
    Next
    
    Erase Wafer
    
    For forx = 0 To StarProbe.ChipCountX
        For fory = 0 To StarProbe.ChipCountY
            Wafer(forx, fory) = WaferTemp(forx, fory)
        Next
    Next
    Call Command_DisplayWafer_Click
    Command_MaskMove.Enabled = True
End Sub

Private Sub Command_Option_Click()
    '[ 2022.07.29 ] : engineer modeÀÎ °æ¿ì¿¡¸¸ Ä§ÀûÈ®ÀÎ ¹öÆ°ÀÌ º¸ÀÎ´Ù.
'    If Mode_Set = True Then
'        SelectExt.Command6.Visible = True
'    Else
'        SelectExt.Command6.Visible = False
'    End If
    
    Command_Option.Enabled = False
    SelectExt.Show vbModal, Me
    
    If Opt_Select_Flag = True Then
        Command_WaferDraw.Enabled = True
        Command_Map_Clear.Enabled = True
        Command_DisplayWafer.Enabled = True
    End If
    Command_Option.Enabled = True
End Sub

Private Sub Command_SaveAs_Click()
    Dim tmp As String
    Dim Item_Cnt As Integer
    
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    CommonDialog1.FileName = ""
    If Mode_Set = True Then                         'engineer mode
        CommonDialog1.Filter = "StarProbe Map File(*.map)|*.map|StarProbe Map File(*.sp)|*.sp|All file(*.*)|*.*"
    Else                                            'operator mode : spÆÄÀÏ¸¸ save as°¡ °¡´ÉÇÏµµ·Ï ÇÑ´Ù.
        CommonDialog1.Filter = "StarProbe Map File(*.sp)|*.sp"
    End If
    CommonDialog1.FilterIndex = 0
    CommonDialog1.ShowSave
    
    If Not Dir(CommonDialog1.FileName) = "" Then
        x = MsgBox(CommonDialog1.FileName & " This file already exists" & vbCrLf & "Relace existing file?", vbQuestion + vbYesNo, "File save")
        If x = 7 Then Exit Sub
    End If
    
    For i = Len(CommonDialog1.FileName) To 1 Step -1
        If Mid(CommonDialog1.FileName, i, 1) = "\" Then
            tmp = UCase(Mid(CommonDialog1.FileName, i + 1, Len(CommonDialog1.FileName) - i + 1))
            Exit For
        End If
    Next i

    If UCase(Right(CommonDialog1.FileName, 4)) = ".MAP" Then
        StarProbe.FileName_Map = CommonDialog1.FileName
        Call StarProbe_FileSave_ControlMap(CommonDialog1.FileName)
    ElseIf UCase(Right(CommonDialog1.FileName, 3)) = ".SP" Then
        StarProbe.FileName_Data = CommonDialog1.FileName
        Call StarProbe_FileSave_Data(CommonDialog1.FileName)
    Else
        Me.MousePointer = 11
        Me.MousePointer = 0
    End If
    Exit Sub
    
ErrHandler:
    Me.MousePointer = 0
End Sub

Private Sub Command_Skip2Ink_Click()
    Dim forx As Integer, fory As Integer
    
    For fory = 0 To StarProbe.ChipCountY
        For forx = 0 To StarProbe.ChipCountX
            If Wafer(forx, fory).Chip And _
               Wafer(forx, fory).ChipSkipDie Then
                Wafer(forx, fory).ChipInk = True
                Call Display_Chip(pZoom, pOriginal, (forx - StarProbe.StartChip.x), (fory - StarProbe.StartChip.y))
            End If
        Next
    Next
End Sub

Public Sub Command_Stop_Click()
    If bPause_Flag = False Then
        Text11.Enabled = True
        Command7.Enabled = True
        bStarProbeStart = False
        bStop = True
        Check1(3).value = 0
        Check1(3).Enabled = True                '2016.06.14

        If Wafer_Start = True Then
            If DemoMode = 0 Then
                If ACCO = 1 Or Model_Select = 1 Then                            'ACCO only
                    MSComm1.Output = "PA" & vbCrLf
                Else
                    MSComm1.Output = "PA" & vbLf
                End If
            End If
        End If
    Else
        MsgBox "Pause State Not Use!", 16, "STAR PROBE"
    End If
End Sub

Private Sub Command_Wafer_Center_Set_Click()
    bStop = False
    Call StarProbe_First_Die_Set
End Sub

Private Sub Command_WaferDirection_Click()
    Command_WaferDirection.Enabled = False

    Dim angle As Integer
    angle = 0
    
    ' -90µµ Ãß°¡
    If Option_WaferDirection(0).value Then angle = -90
    If Option_WaferDirection(1).value Then angle = 90
    If Option_WaferDirection(2).value Then angle = 180
    If Option_WaferDirection(3).value Then angle = 270
    
    '[ 2021.01.11 ] : angel = 0ÀÌ¸é ±×³É ³ª°£´Ù.
    If angle = -90 Or angle = 90 Or angle = 180 Or angle = 270 Then
        Call WaferDirection(angle)
        Call Command_DisplayWafer_Click
    End If
    Command_WaferDirection.Enabled = True
End Sub

Private Sub Command_WaferDraw_Click()
    Form_StarProbe_WaferDraw.Show vbModal, Me
    Command_Wafer_Center_Set.Enabled = True
End Sub

Private Sub Command1_Click(Index As Integer)
    Command1(Index).Enabled = False
    
    Select Case Index
        Case 0
            Command1(0).Enabled = False
            Me.MousePointer = 11
            
            Call TEST
            
            Me.MousePointer = 0
            Command1(0).Enabled = True
            Command1(0).SetFocus
            
        Case 1
            x = MsgBox("Summary¸¦ ClearÇÏ½Ã°Ú½À´Ï±î?", vbQuestion + vbYesNo, "ClearÈ®ÀÎ")
            
            If x = vbYes Then RESET_DATA
            lblWafer.Caption = ""
    End Select
    Command1(Index).Enabled = True
End Sub

Sub RESET_DATA()
    Dim i As Integer

    Test_Cnt = 0
    Good_Cnt = 0
    Test_Fail_Count1 = 0
    Test_Fail_Count2 = 0
    Test_Fail_Count3 = 0
    Test_Fail_Count4 = 0

    For i = 1 To 24
        Bin_Result(i) = 0
    Next
    For i = 0 To 24
        Bin_Count(i) = 0
        Text_Bin_Count_No(i) = 0
        Text_BinCount(i).Text = 0
    Next
    
    SSPanel2(1).Caption = ""
    SSPanel2(2).Caption = ""
    Text_GoodCount.Text = 0
    Text_BadCount.Text = 0
    Text_TotalCount.Text = 0
    SSPanel2(4).Caption = ""
    SSPanel2(4).BackColor = &H8000000F
End Sub

Private Sub Command10_Click()
    Dim forx As Integer, fory As Integer
    
    For fory = 0 To StarProbe.ChipCountY
        For forx = 0 To StarProbe.ChipCountX
            If Wafer(forx, fory).Chip And _
               Wafer(forx, fory).ChipMask Then
                Wafer(forx, fory).ChipMask = False
                Wafer(forx, fory).ChipSkipDie = True
                Wafer(forx, fory).ChipInk = True
                Call Display_Chip(pZoom, pOriginal, (forx - StarProbe.StartChip.x), (fory - StarProbe.StartChip.y))
            End If
        Next
    Next
End Sub

Private Sub Command11_Click()
    Call StarProbe_Left_Ink_Dot(StarProbe.Ink_LeftPort)
End Sub

Private Sub Command12_Click()
    Dim forx As Integer, fory As Integer
    
    For fory = Stop_yy To StarProbe.ChipCountY
        For forx = 0 To StarProbe.ChipCountX
            If (fory = Stop_yy) And (forx <= Stop_xx) Then
            
            Else
                If Wafer(forx, fory).Chip And _
                   Not Wafer(forx, fory).ChipSkipDie And _
                   Not Wafer(forx, fory).ChipPlate And _
                   Not Wafer(forx, fory).ChipMask And _
                   Not Wafer(forx, fory).flag Then
    
                    Wafer(forx, fory).flag = True
                    Wafer(forx, fory).FlagBad = False
                    Wafer(forx, fory).BIN = GOOD_BIN_NO
                    Wafer(forx, fory).ChipMeasure = True
                        
                    Wafer(forx, fory).ChipInk = True
                    Wafer(forx, fory).InkDot = True
                End If
            End If
        Next
    Next
    Call Command_DisplayWafer_Click
End Sub

Private Sub Command13_Click()
    Dim x, y As Integer    'Ãß°¡

    x = 0
    y = 0

    Call StarProbe_XY_Moving((x), (y))
    If StarProbe_Motor_End_check Then MsgBox "Motor not end check !", 16, "STAR PROBE"
End Sub

Private Sub Command18_Click()
    Dim s As String
    Dim i As Integer
    Dim iSearchCount As Integer
    Dim tmp1 As Variant
    Dim tmp2 As Variant
    
    If ACCO = 1 Or Model_Select = 1 Then                            'ACCO only
        s = Replace(Text_ReciveData, vbCrLf, "")
    Else
        s = Replace(Text_ReciveData, vbLf, "")
    End If
    
    If Left(s, 2) = "ET" Then
        s = Mid(s, 3)
        tmp1 = Split(s, ",")
        tmp2 = Split(s, ",")
        
        iSearchCount = 0
        For icount = 1 To Len(s)
            If Mid(s, icount, 1) = "," Then iSearchCount = iSearchCount + 1
        Next
        For i = 0 To iSearchCount
            tmp2(i) = tmp1(iSearchCount - i)
            If tmp2(i) <> 32 Then
                Text_ChipBIN(iSearchCount - i).Text = tmp2(i)
                Text_ChipRecive(iSearchCount - i).Text = tmp2(i)
            End If
        Next
    End If
End Sub

Private Sub Command2_Click(Index As Integer)
    Dim Status As String
    Dim filenamesaved As String
    
    Needle_check_flag = False       '[ 2022.08.31 ]
    
    Me.MousePointer = 11
    Select Case Index
        Case 1                          ' join program
            Command2(1).Enabled = False
            CommonDialog1.CancelError = True
            
            On Error GoTo ErrHandler
            
            CommonDialog1.FileName = ""
            If AOI_MODE = 0 Then
                '[ 2020.11.02 ] : aoi Ãß°¡
                CommonDialog1.Filter = "StarProbe Map File(*.map)|*.map|StarProbe Map File(*.sp)|*.sp|All file(*.*)|*.*"
            Else
                '[ 2020.11.02 ] : aoi Ãß°¡
                CommonDialog1.Filter = "StarProbe Map File(*.map)|*.map|StarProbe Map File(*.sp)|*.sp|Star Probe Data File(*.aoi)|*.aoi|All file(*.*)|*.*"
            End If
            CommonDialog1.FilterIndex = 0
            CommonDialog1.ShowOpen
                      
            Me.Refresh
            If DemoMode = 0 Then
                Check1(3).Enabled = True
                Command4.Enabled = True
                Command1(0).Enabled = True
                Command1(1).Enabled = True
            End If
            
            If UCase(Right(CommonDialog1.FileName, 4)) = ".MAP" Then
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''2016.03.11
                Dim xx As Integer, yy As Integer

                Command_Map_Clear.Enabled = False
                
                For xx = 0 To StarProbe.ChipCountX
                    For yy = 0 To StarProbe.ChipCountY
                        Wafer(xx, yy).flag = False
                        Wafer(xx, yy).FlagBad = False
                        Wafer(xx, yy).MeasureWait = False
                        Wafer(xx, yy).InkDot = False
                        Wafer(xx, yy).ChipMeasure = False ' 2005.09.05
                        Wafer(xx, yy).BIN = 0             '16.12.02
                    Next
                Next
                Fail_Loop = False
                Stop_Measure = False
                Command_Map_Clear.Enabled = True
                lblWafer.Caption = ""
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                bStarprobe_AfterInk = False         'INK MODE OFF
                StarProbe.FileName_Map = CommonDialog1.FileName
                
                StarProbeTemp = StarProbe
                
                Call StarProbe_FileLoad_ControlMap(CommonDialog1.FileName)
                
                StarProbe.CountBadDie = StarProbeTemp.CountBadDie
                StarProbe.CountGoodDie = StarProbeTemp.CountGoodDie
                StarProbe.CountSkipDie = StarProbeTemp.CountSkipDie
                StarProbe.CountTotalChip = StarProbeTemp.CountTotalChip
                
                Call Command_DisplayWafer_Click
                
                StarProbe.MeasureStepX = StarProbeTemp.MeasureStepX
                StarProbe.MeasureStepY = StarProbeTemp.MeasureStepY
                
                StarProbe.Ink_LeftPort = StarProbeTemp.Ink_LeftPort
                StarProbe.Ink_RightPort = StarProbeTemp.Ink_RightPort
                StarProbe.LineOk = StarProbeTemp.LineOk
                StarProbe.RCount = StarProbeTemp.RCount
                StarProbe.RCount_Sub = StarProbeTemp.RCount_Sub
                StarProbe.MeasureSleep = StarProbeTemp.MeasureSleep
                StarProbe.ReMeasure = StarProbeTemp.ReMeasure
                StarProbe.Ink_After_LeftPort = StarProbeTemp.Ink_After_LeftPort
                StarProbe.Ink_After_RightPort = StarProbeTemp.Ink_After_RightPort
                StarProbe.Ink_After_CenterPort = StarProbeTemp.Ink_After_CenterPort
                StarProbe.LimitArea = StarProbeTemp.LimitArea
                
                StarProbe.WaferTest = StarProbeTemp.WaferTest
                StarProbe.MeasureAll = StarProbeTemp.MeasureAll
                
                StarProbe.Ink_After = StarProbeTemp.Ink_After
                
                Me.MousePointer = 0
                Command2(1).Enabled = True
                
                Text8.Text = StarProbe.ChipSizeX
                Text9.Text = StarProbe.ChipSizeY
                
                Call StarProbe_Set_Die_Size(val(StarProbe.ChipSizeX), val(StarProbe.ChipSizeY))
                
                '[ 2021.01.11 ] : ¹æÇâ ¿É¼Ç ÃÊ±âÈ­
                Option_WaferDirection(0).value = False
                Option_WaferDirection(1).value = False
                Option_WaferDirection(2).value = False
                Option_WaferDirection(3).value = False
                
                Needle_Chk_Ok = False                       '[ 2022.07.29 ]
                
            ElseIf UCase(Right(CommonDialog1.FileName, 4)) = ".AOI" Then
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''2016.03.11
                Command_Map_Clear.Enabled = False
                                   
                For xx = 0 To StarProbe.ChipCountX
                    For yy = 0 To StarProbe.ChipCountY
                        Wafer(xx, yy).flag = False
                        Wafer(xx, yy).FlagBad = False
                        Wafer(xx, yy).MeasureWait = False
                        Wafer(xx, yy).InkDot = False
                        Wafer(xx, yy).ChipMeasure = False '2005.09.05
                        Wafer(xx, yy).BIN = 0             '16.12.02
                    Next
                Next
                Fail_Loop = False
                Stop_Measure = False
                Command_Map_Clear.Enabled = True
                
                RESET_DATA
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                
                '============================================== 2015.11.30 : load map name display
                Load_MAP = CommonDialog1.FileName
                For i = Len(Load_MAP) To 1 Step -1
                    If Mid(Load_MAP, i, 1) = "\" Then
                        Load_MAP = UCase(Mid(Load_MAP, i + 1, Len(Load_MAP) - i + 1))
                        SSPanel2(0).Caption = Load_MAP
                        Exit For
                    End If
                Next i
                
                '[ 2017.03.23 ] : SPÆÄÀÏÀ» ·ÎµùÇÑ °æ¿ì ÆÄÀÏÀÇ ³Ñ¹ö¸¦ ÃßÃâÇÏ´Â ÄÚµå Ãß°¡
                For i = Len(Load_MAP) To 1 Step -1
                    If Mid(Load_MAP, i, 1) = "-" Then
                        ASD = Mid(Load_MAP, i + 1, Len(Load_MAP) - i + 1)
                        SP_CNT = Mid(ASD, 1, Len(ASD) - 4)
                        lblWafer.Caption = SP_CNT
                        AOI_MAP(SP_CNT) = CommonDialog1.FileName
                        AOI_Use = True
                        Exit For
                    End If
                Next i
                '==============================================
                bStarprobe_AfterInk = False         'INK MODE OFF
                For xxxx = 0 To 8
                    Shape_Mea(xxxx).Visible = False
                Next xxxx
                StarProbe.FileName_Map = CommonDialog1.FileName
                
                StarProbeTemp = StarProbe
                
                Call StarProbe_FileLoad_ControlMap(CommonDialog1.FileName)
                                                        
                StarProbe.CountBadDie = StarProbeTemp.CountBadDie
                StarProbe.CountGoodDie = StarProbeTemp.CountGoodDie
                StarProbe.CountSkipDie = StarProbeTemp.CountSkipDie
                StarProbe.CountTotalChip = StarProbeTemp.CountTotalChip
                
                '[ 2020.03.27 ] : DB°ü·Ã ¼öÁ¤
                ''''''''''''''''''''''''''''''''''''''''''''''''
                Map_Total_Backup = StarProbe.CountTotalChip
                Map_Good_Backup = StarProbe.CountGoodDie
                Map_Bad_Backup = StarProbe.CountBadDie
                Map_Skip_Backup = StarProbe.CountSkipDie
                '''''''''''''''''''''''''''''''''''''''''''''''''
                
'                    Call Command_Map_Clear_Click                '2016.06.02 : map clearÃß°¡
                                    
                Call Command_DisplayWafer_Click
                
                StarProbe.MeasureStepX = StarProbeTemp.MeasureStepX
                StarProbe.MeasureStepY = StarProbeTemp.MeasureStepY
                
                StarProbe.Ink_LeftPort = StarProbeTemp.Ink_LeftPort
                StarProbe.Ink_RightPort = StarProbeTemp.Ink_RightPort
                StarProbe.LineOk = StarProbeTemp.LineOk
                StarProbe.RCount = StarProbeTemp.RCount
                StarProbe.RCount_Sub = StarProbeTemp.RCount_Sub
                StarProbe.MeasureSleep = StarProbeTemp.MeasureSleep
                StarProbe.ReMeasure = StarProbeTemp.ReMeasure
                StarProbe.Ink_After_LeftPort = StarProbeTemp.Ink_After_LeftPort
                StarProbe.Ink_After_RightPort = StarProbeTemp.Ink_After_RightPort
                StarProbe.Ink_After_CenterPort = StarProbeTemp.Ink_After_CenterPort
                StarProbe.LimitArea = StarProbeTemp.LimitArea
                
                StarProbe.WaferTest = StarProbeTemp.WaferTest
                StarProbe.MeasureAll = StarProbeTemp.MeasureAll
                
                StarProbe.Ink_After = StarProbeTemp.Ink_After
                
                Me.MousePointer = 0
                Command2(1).Enabled = True
                
                Call StarProbe_Set_Die_Size(val(StarProbe.ChipSizeX), val(StarProbe.ChipSizeY))
                
                '[ 2021.01.11 ] : ¹æÇâ ¿É¼Ç ÃÊ±âÈ­
                Option_WaferDirection(0).value = False
                Option_WaferDirection(1).value = False
                Option_WaferDirection(2).value = False
                Option_WaferDirection(3).value = False
            ElseIf UCase(Right(CommonDialog1.FileName, 3)) = ".SP" Then
                
                StarProbe.FileName_Data = CommonDialog1.FileName
                StarProbeTemp = StarProbe
                
                '============================================== 2015.11.30 : load map name display
                Load_MAP = CommonDialog1.FileName
                For i = Len(Load_MAP) To 1 Step -1
                    If Mid(Load_MAP, i, 1) = "\" Then
                        Load_MAP = UCase(Mid(Load_MAP, i + 1, Len(Load_MAP) - i + 1))
                        SSPanel2(0).Caption = Load_MAP
                        Exit For
                    End If
                Next i
                
                '[ 2017.03.23 ] : SPÆÄÀÏÀ» ·ÎµùÇÑ °æ¿ì ÆÄÀÏÀÇ ³Ñ¹ö¸¦ ÃßÃâÇÏ´Â ÄÚµå Ãß°¡
                For i = Len(Load_MAP) To 1 Step -1
                    If Mid(Load_MAP, i, 1) = "_" Then
                        ASD = Mid(Load_MAP, i + 1, Len(Load_MAP) - i + 1)
                        SP_CNT = Mid(ASD, 1, Len(ASD) - 3)
                        lblWafer.Caption = SP_CNT        '[ 2020.11.13 ] : ³Ñ¹öÇ¥½Ã Ãß°¡.
                        Exit For
                    End If
                Next i
                ''''''''''''
                
                '================================================================2015.12.22
                bStarprobe_AfterInk = IIf(UCase(Right(StarProbe.FileName_Data, 7)) = "TEMP.SP", False, True)
                '================================================================2015.12.22
                bStarprobe_AfterInk = False
                
                '==============================================
                SP_FLAG = True
                If Check_OldSPFile.value = vbChecked Then
                    Call Starprobe_FileLoad_OldData(CommonDialog1.FileName)
                Else
                    Call Starprobe_FileLoad_Data(CommonDialog1.FileName)
                End If
                
                StarProbeTemp.CountBadDie = StarProbe.CountBadDie
                StarProbeTemp.CountGoodDie = StarProbe.CountGoodDie
                StarProbeTemp.CountSkipDie = StarProbe.CountSkipDie
                StarProbeTemp.CountTotalChip = StarProbe.CountTotalChip
                
                Call Command_DisplayWafer_Click
                SP_FLAG = False
                StarProbe.CountBadDie = StarProbeTemp.CountBadDie
                StarProbe.CountGoodDie = StarProbeTemp.CountGoodDie
                StarProbe.CountSkipDie = StarProbeTemp.CountSkipDie
                StarProbe.CountTotalChip = StarProbeTemp.CountTotalChip
                                
                StarProbe.Ink_LeftPort = StarProbeTemp.Ink_LeftPort
                StarProbe.Ink_RightPort = StarProbeTemp.Ink_RightPort
                StarProbe.LineOk = StarProbeTemp.LineOk
                StarProbe.RCount = StarProbeTemp.RCount
                StarProbe.RCount_Sub = StarProbeTemp.RCount_Sub
                StarProbe.MeasureSleep = StarProbeTemp.MeasureSleep
                StarProbe.ReMeasure = StarProbeTemp.ReMeasure
                StarProbe.Ink_After_LeftPort = StarProbeTemp.Ink_After_LeftPort
                StarProbe.Ink_After_RightPort = StarProbeTemp.Ink_After_RightPort
                StarProbe.Ink_After_CenterPort = StarProbeTemp.Ink_After_CenterPort
                StarProbe.LimitArea = StarProbeTemp.LimitArea
                
                StarProbe.WaferTest = StarProbeTemp.WaferTest
                
                StarProbe.Ink_After = StarProbeTemp.Ink_After
                
                SSPanel2(1).Caption = Test_Cnt
                SSPanel2(2).Caption = Good_Cnt
                SSPanel_BadCount.Caption = StarProbe.CountBadDie
                
                Me.MousePointer = 0
                Command2(1).Enabled = True
                
'                Call StarProbe_Set_Die_Size(val(StarProbe.ChipSizeX), val(StarProbe.ChipSizeY))
                
                '[ 2021.01.11 ] : ¹æÇâ ¿É¼Ç ÃÊ±âÈ­
                Option_WaferDirection(0).value = False
                Option_WaferDirection(1).value = False
                Option_WaferDirection(2).value = False
                Option_WaferDirection(3).value = False
                
            ElseIf UCase(Right(CommonDialog1.FileName, 4)) = ".TXT" Then            'only ink mode
                bStarprobe_AfterInk = True
                
                StarProbe.FileName_Data = CommonDialog1.FileName
                                                
                Check1(1).value = 0                 '2016.09.27 : Ink Start Position Check Box Off
                Ink_Start_Flag = 0                  '2016.09.27 : Ink Start Position Flag Clear
                
                StarProbeTemp = StarProbe
                
                If Check_OldSPFile.value = vbChecked Then
                    Call Starprobe_FileLoad_OldData(CommonDialog1.FileName)
                Else
                    Call Starprobe_FileLoad_Data_TXT(CommonDialog1.FileName)
                End If
                
                StarProbeTemp.CountBadDie = StarProbe.CountBadDie
                StarProbeTemp.CountGoodDie = StarProbe.CountGoodDie
                StarProbeTemp.CountSkipDie = StarProbe.CountSkipDie
                StarProbeTemp.CountTotalChip = StarProbe.CountTotalChip
                
                Call Command_DisplayWafer_Click
                                
                StarProbe.CountBadDie = StarProbeTemp.CountBadDie
                StarProbe.CountGoodDie = StarProbeTemp.CountGoodDie
                StarProbe.CountSkipDie = StarProbeTemp.CountSkipDie
                StarProbe.CountTotalChip = StarProbeTemp.CountTotalChip
                                
                StarProbe.Ink_LeftPort = StarProbeTemp.Ink_LeftPort
                StarProbe.Ink_RightPort = StarProbeTemp.Ink_RightPort
                StarProbe.LineOk = StarProbeTemp.LineOk
                StarProbe.RCount = StarProbeTemp.RCount
                StarProbe.RCount_Sub = StarProbeTemp.RCount_Sub
                StarProbe.MeasureSleep = StarProbeTemp.MeasureSleep
                StarProbe.ReMeasure = StarProbeTemp.ReMeasure
                StarProbe.Ink_After_LeftPort = StarProbeTemp.Ink_After_LeftPort
                StarProbe.Ink_After_RightPort = StarProbeTemp.Ink_After_RightPort
                StarProbe.Ink_After_CenterPort = StarProbeTemp.Ink_After_CenterPort
                StarProbe.LimitArea = StarProbeTemp.LimitArea
                
                StarProbe.WaferTest = StarProbeTemp.WaferTest
                
                StarProbe.Ink_After = StarProbeTemp.Ink_After
                
                Me.MousePointer = 0
                
                Command2(1).Enabled = True
                '[ 2021.01.11 ] : ¹æÇâ ¿É¼Ç ÃÊ±âÈ­
                Option_WaferDirection(0).value = False
                Option_WaferDirection(1).value = False
                Option_WaferDirection(2).value = False
                Option_WaferDirection(3).value = False
            End If

            SSTab3.Tab = 0
            
            For i = Len(CommonDialog1.FileName) To 1 Step -1
                If Mid(CommonDialog1.FileName, i, 1) = "\" Then
                    SSPanel2(0).Caption = UCase(Mid(CommonDialog1.FileName, i + 1, Len(CommonDialog1.FileName) - i + 1))
                    Exit For
                End If
            Next i
                        
            Me.MousePointer = 0
            Command2(1).Enabled = True
            Exit Sub
            
ErrHandler:
            Me.MousePointer = 0
            Command2(1).Enabled = True
            Exit Sub
    End Select
    Me.MousePointer = 0
End Sub

'bin clearÈÄ ´Ù½Ã ÃøÁ¤½Ã bin0ÀÌ ³ª¿À´Â Çö»ó ¹ß»ý.
Sub TEST()
    Dim k As Integer
    Dim STBin, s As String
        
    If STT_time = "" Then STT_time = Format(Now, "YYYY.MM.DD hh:mm:ss")           '07.11.21³¯Â¥º¯°æ½Ã ¿À·ù¹ß°ß ¼öÁ¤.
    
    If QueryPerformanceFrequencyAny(frequency) = 0 Then
        MsgBox "This computer doesn't support high-res timers", vbCritical
        Exit Sub
    End If
    QueryPerformanceCounterAny startTime
    
    Set Test_Start_Delay = New ccrpStopWatch
    
    If Tester_Select = 1 Or Model_Select = 2 Then         'AMT-88
        If XAxis = 0 And YAxis = 0 Then
            s = "TFX" & XAxis & "Y" & YAxis & "," & Trim(Text_ChipTest.Text)            'first die test : TF
        Else
            s = "TSX" & XAxis & "Y" & YAxis & "," & Trim(Text_ChipTest.Text)            'normal die test : TS
        End If
    Else                                'amt-88
        If CH_SET = 1 Then
            If XAxis = 0 And YAxis = 0 Then
                s = "TFX" & XAxis & "Y" & YAxis & "," & Trim(Text_ChipTest.Text)
            Else
                s = "TSX" & XAxis & "Y" & YAxis & "," & Trim(Text_ChipTest.Text)
            End If
        ElseIf CH_SET = 2 Then
            If XAxis = 0 And YAxis = 0 Then
                s = "TFX" & XAxis & "Y" & YAxis + 1 & "," & Trim(Text_ChipTest.Text)
            Else
                s = "TSX" & XAxis & "Y" & YAxis + 1 & "," & Trim(Text_ChipTest.Text)
            End If
        Else
            If XAxis = 0 And YAxis = 0 Then
                s = "TFX" & XAxis & "Y" & YAxis + 3 & "," & Trim(Text_ChipTest.Text)
            Else
                s = "TSX" & XAxis & "Y" & YAxis + 3 & "," & Trim(Text_ChipTest.Text)
            End If
        End If
    End If
            
    If DemoMode = 0 Then
        If ACCO = 1 Or Model_Select = 1 Then                            'ACCO only
            MSComm1.Output = s & vbCrLf
        Else
            MSComm1.Output = s & vbLf
        End If
    End If
        
    Test_Bin = 0
    Set Test_End_Delay = New ccrpStopWatch
    Test_End_Delay.Reset
    Test_Start_Delay.Reset
    
    STBin = ""
    s = ""
    
    Do
        Do
            If Test_Start_Delay.Elapsed > 20 Then Exit Do               '[ 2022.05.04 ] : 50->20
        Loop
        
        If DemoMode = 0 Then
            If TESTER_OFF = False Then
                If DemoMode = 0 Then STBin = MSComm1.Input                                       'bin read
            Else
                If ACCO = 1 Or Model_Select = 1 Then                            'ACCO only
                    STBin = "ET1,1,1,1" & vbCrLf            'ÀÓ½Ã
                Else
                    STBin = "ET1,1,1,1" & vbLf            'ÀÓ½Ã
                End If
            End If
        Else
            If ACCO = 1 Or Model_Select = 1 Then                            'ACCO only
                If CH_SET = 1 Then
                    If bad_click = 0 Then
                        STBin = "ET1" & vbCrLf                              'pass
                    Else
                        STBin = "ET10" & vbCrLf                          'fail
                    End If
                ElseIf CH_SET = 2 Then
                    If bad_click = 0 Then
                        STBin = "ET1,1" & vbCrLf                              'pass
                    Else
                        STBin = "ET10,10" & vbCrLf                          'fail
                    End If
                Else
                    If bad_click = 0 Then
                        STBin = "ET1,1,1,1" & vbCrLf                              'pass
                    Else
                        STBin = "ET10,10,1,1" & vbCrLf                          'fail
                    End If
                End If
            Else
                If CH_SET = 1 Then
                    If bad_click = 0 Then
                        STBin = "ET1" & vbLf                              'pass
                    Else
                        STBin = "ET10" & vbLf                          'fail
                    End If
                ElseIf CH_SET = 2 Then
                    If bad_click = 0 Then
                        STBin = "ET1,1" & vbLf                              'pass
                    Else
                        STBin = "ET10,10" & vbLf                          'fail
                    End If
                Else
                    If bad_click = 0 Then
                        STBin = "ET1,1,1,1" & vbLf                              'pass
                    Else
                        STBin = "ET10,10,1,1" & vbLf                          'fail
                    End If
                End If
            End If
            bad_click = 0
        End If
        
        If STBin <> Empty Then
            If ACCO = 1 Or Model_Select = 1 Then                            'ACCO only
                If Right(STBin, 1) = vbCrLf Then
                    STBin = Trim(STBin)
                    Text_ReciveData.Text = STBin
                    Text4.Text = STBin
                    Label6.Caption = STBin
                    Test_Start_Delay.Reset
                    Do
                        If Test_Start_Delay.Elapsed > 5 Then Exit Do
                    Loop
                    
                    Exit Do
                Else
                    STBin = Trim(STBin)
                   
                    Test_Start_Delay.Reset
                    Do
                        If Test_Start_Delay.Elapsed > 10 Then Exit Do
                    Loop
                   
                    If DemoMode = 0 Then
                        s = MSComm1.Input
                    Else
                        s = 1
                    End If
                    STBin = STBin & s
                   
                    If Right(STBin, 1) <> vbCrLf Then
                        Test_Start_Delay.Reset
                        Do
                            If Test_Start_Delay.Elapsed > 20 Then Exit Do
                        Loop
                        If DemoMode = 0 Then
                            s = MSComm1.Input
                        Else
                            s = 1
                        End If
                        STBin = STBin & s
                    End If
                   
                    STBin = Trim(STBin)
                    Text_ReciveData.Text = STBin
                    Text4.Text = STBin
                    Label6.Caption = STBin
                    Test_Start_Delay.Reset
                    Do
                        If Test_Start_Delay.Elapsed > 5 Then Exit Do
                    Loop
                   
                    
                    Exit Do
                End If
            Else
                If Right(STBin, 1) = vbLf Then
                    STBin = Trim(STBin)
                    Text_ReciveData.Text = STBin
                    Text4.Text = STBin
                    Label6.Caption = STBin
                    Test_Start_Delay.Reset
                    Do
                        If Test_Start_Delay.Elapsed > 5 Then Exit Do
                    Loop
                    If DemoMode = 0 Then MSComm1.Output = ">" & vbLf
                    Exit Do
                Else
                    STBin = Trim(STBin)
                   
                    Test_Start_Delay.Reset
                    Do
                        If Test_Start_Delay.Elapsed > 10 Then Exit Do
                    Loop
                   
                    If DemoMode = 0 Then
                        s = MSComm1.Input
                    Else
                        s = 1
                    End If
                    STBin = STBin & s
                   
                    If Right(STBin, 1) <> vbLf Then
                        Test_Start_Delay.Reset
                        Do
                            If Test_Start_Delay.Elapsed > 20 Then Exit Do
                        Loop
                        If DemoMode = 0 Then
                            s = MSComm1.Input
                        Else
                            s = 1
                        End If
                        STBin = STBin & s
                    End If
                   
                    STBin = Trim(STBin)
                    Text_ReciveData.Text = STBin
                    Text4.Text = STBin
                    Label6.Caption = STBin
                    Test_Start_Delay.Reset
                    Do
                        If Test_Start_Delay.Elapsed > 5 Then Exit Do
                    Loop
                   
                    If DemoMode = 0 Then MSComm1.Output = ">" & vbLf
                    Exit Do
                End If
            End If
        End If
        Test_Start_Delay.Reset
        If Test_End_Delay.Elapsed > 3000 Then       '3ÃÊµ¿¾È ÀÀ´äÀÌ ¾øÀ» °æ¿ì ·çÆ¾À» Á¾·áÇÑ´Ù.
            Exit Do
        End If
    Loop
       
    Call Command18_Click
    
    If DemoMode = 0 Then
        STBin = MSComm1.Input
        STBin = MSComm1.Input
    End If
    
    QueryPerformanceCounterAny endTime
    
    If frequency <> 0 Then result = (endTime - startTime) * 1000 / frequency
    result = GetRound(result, 0)
    If result = 0 Then result = 1
        
    If Auto_flag = False Then
        SSPanel2(9).Caption = result & Space(1) & "ms"
        SSPanel2(1).Caption = Test_Cnt
        Text_TotalCount.Text = Test_Cnt
    End If
    END_time = Format(Now, "YYYY.MM.DD hh:mm:ss")           'End TimeÀ» ¼³Á¤ÇØ ÁØ´Ù
    Set Test_End_Delay = Nothing
End Sub

Private Sub Command3_Click()
    Dim forx As Integer, fory As Integer
    
    For fory = Stop_yy To StarProbe.ChipCountY
        For forx = 0 To StarProbe.ChipCountX
            If (fory = Stop_yy) And (forx <= Stop_xx) Then
            Else
                If Wafer(forx, fory).Chip And _
                   Not Wafer(forx, fory).ChipSkipDie And _
                   Not Wafer(forx, fory).ChipPlate And _
                   Not Wafer(forx, fory).ChipMask And _
                   Not Wafer(forx, fory).flag Then
    
                    Wafer(forx, fory).flag = True
                    Wafer(forx, fory).FlagBad = False
                    Wafer(forx, fory).BIN = GOOD_BIN_NO
                    Wafer(forx, fory).ChipMeasure = True
                    Wafer(forx, fory).ChipInk = True
                    Wafer(forx, fory).InkDot = True
                End If
            End If
        Next
    Next
    Call Command_DisplayWafer_Click
End Sub

Private Sub Command4_Click()
'    ID_CHECK = False
    Command4.BackColor = vbRed
    Timer1.Enabled = True
End Sub

'''
'[ 2020.09.14 ] : Æú´õ¿¡¼­ mapÆÄÀÏÀ» Ã£´Â ÇÔ¼ö Ãß°¡
Private Sub Get_Folder(folder As Object)
    
On Error GoTo ErrorSub
    Dim cnt As Integer
    Dim f As Object
    
    cnt = 0
        
    For Each f In folder.Files
        If Array_tmp1 <> "" Then
            If InStr(f, Array_tmp1) <> 0 Then
                If InStr(f, ".map") <> 0 Or InStr(f, ".MAP") <> 0 Then          '°°Àº ÀÌ¸§ ¼ö·®À» Ã¼Å©ÇÑ´Ù.
                    cnt = cnt + 1
                End If
            End If
        End If
    Next
    
    If cnt > 1 Then                                                             '°°Àº ÀÌ¸§ÀÌ ÀÖ´Â °æ¿ì DEV¸¦ ºñ±³ÇÑ´Ù.
        For Each f In folder.Files
            If Array_tmp1 <> "" Then
                If InStr(f, DEV) <> 0 Then
                    If InStr(f, ".map") <> 0 Or InStr(f, ".MAP") <> 0 Then
                        Load_MAP = f
                        Exit Sub
                    End If
                End If
            End If
        Next
    Else
        For Each f In folder.Files
            If Array_tmp1 <> "" Then
                If InStr(f, Array_tmp1) <> 0 Then
                    If InStr(f, ".map") <> 0 Or InStr(f, ".MAP") <> 0 Then          '°°Àº ÀÌ¸§ ¼ö·®À» Ã¼Å©ÇÑ´Ù.
                        Load_MAP = f
                        Exit Sub
                    End If
                End If
            End If
        Next
    End If

ErrorSub:
    Load_MAP = ""
End Sub

Public Sub LOAD_CONTROL()
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''2016.03.11
    Dim fso As Object                                   '[ 2020.09.14 ]
    Dim FolderList As Object                            '[ 2020.09.14 ]
    Dim xx As Integer, yy As Integer
    
    If AOI_Use = True Then
        Command_Map_Clear.Enabled = False

        For xx = 0 To StarProbe.ChipCountX                      'map º¯¼ö ÃÊ±âÈ­
            For yy = 0 To StarProbe.ChipCountY
                Wafer(xx, yy).flag = False
                Wafer(xx, yy).FlagBad = False
                Wafer(xx, yy).MeasureWait = False
                Wafer(xx, yy).InkDot = False
                Wafer(xx, yy).ChipMeasure = False
                Wafer(xx, yy).BIN = 0
            Next
        Next
        Fail_Loop = False
        Stop_Measure = False
        Command_Map_Clear.Enabled = True
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim map_first As String                                 'Ã³À½ loadÇÒ ÆÄÀÏ ÀÌ¸§ ÁöÁ¤ º¯¼ö
        For i = 0 To 24
            If NOW_NO(i) = True Then
                map_first = AOI_MAP(i + 1)
                Exit For
            End If
        Next i
        
        If map_first = "" Then
            MsgBox "MAP File not found!", 16, "File error"
            Exit Sub
        End If
        '============================================== 2015.11.30 : load map name display
        For i = Len(map_first) To 1 Step -1
            If Mid(map_first, i, 1) = "\" Then
                SSPanel2(0).Caption = UCase(Mid(map_first, i + 1, Len(map_first) - i + 1))
                Exit For
            End If
        Next i
        '==============================================
        
        bStarprobe_AfterInk = False                             'INK MODE OFF
        For xxxx = 0 To 8
            Shape_Mea(xxxx).Visible = False
        Next xxxx
        StarProbe.FileName_Map = Load_MAP

        StarProbeTemp = StarProbe
        
        Call StarProbe_FileLoad_ControlMap(map_first)
                                                                              
        StarProbe.CountBadDie = StarProbeTemp.CountBadDie
        StarProbe.CountGoodDie = StarProbeTemp.CountGoodDie
        StarProbe.CountSkipDie = StarProbeTemp.CountSkipDie
        StarProbe.CountTotalChip = StarProbeTemp.CountTotalChip

        '[ 2020.03.27 ] : DB°ü·Ã ¼öÁ¤
        ''''''''''''''''''''''''''''''''''''''''''''''''
        Map_Total_Backup = StarProbe.CountTotalChip
        Map_Good_Backup = StarProbe.CountGoodDie
        Map_Bad_Backup = StarProbe.CountBadDie
        Map_Skip_Backup = StarProbe.CountSkipDie
        '''''''''''''''''''''''''''''''''''''''''''''''''

        'Call Command_Map_Clear_Click                '2016.06.02 : map clearÃß°¡

        Call Command_DisplayWafer_Click

        StarProbe.MeasureStepX = StarProbeTemp.MeasureStepX
        StarProbe.MeasureStepY = StarProbeTemp.MeasureStepY

        StarProbe.Ink_LeftPort = StarProbeTemp.Ink_LeftPort
        StarProbe.Ink_RightPort = StarProbeTemp.Ink_RightPort
        StarProbe.LineOk = StarProbeTemp.LineOk
        StarProbe.RCount = StarProbeTemp.RCount
        StarProbe.RCount_Sub = StarProbeTemp.RCount_Sub
        StarProbe.MeasureSleep = StarProbeTemp.MeasureSleep
        StarProbe.ReMeasure = StarProbeTemp.ReMeasure
        StarProbe.Ink_After_LeftPort = StarProbeTemp.Ink_After_LeftPort
        StarProbe.Ink_After_RightPort = StarProbeTemp.Ink_After_RightPort
        StarProbe.Ink_After_CenterPort = StarProbeTemp.Ink_After_CenterPort
        StarProbe.LimitArea = StarProbeTemp.LimitArea

        StarProbe.WaferTest = StarProbeTemp.WaferTest
        StarProbe.MeasureAll = StarProbeTemp.MeasureAll

        StarProbe.Ink_After = StarProbeTemp.Ink_After

        Me.MousePointer = 0
        Command2(1).Enabled = True
        
        '============================================================ [ 2017.08.22 ]
        For XXX = 0 To 24
            If NOW_NO(XXX) = True Then      'cassette¿¡¼­ ¼³Á¤ÇÑ ¼ø¼­´ë·Î tt_no¸¦ ¼³Á¤ÇÑ´Ù.
                TT_NO = XXX
                Exit For
            End If
        Next XXX
        '============================================================
        lblWafer.Caption = TT_NO + 1        'mainÈ­¸é¿¡ wafer no¸¦ Ç¥½ÃÇÑ´Ù.

        Call StarProbe_Set_Die_Size(val(StarProbe.ChipSizeX), val(StarProbe.ChipSizeY))
    Else
        Command_Map_Clear.Enabled = False
        
        For xx = 0 To StarProbe.ChipCountX
            For yy = 0 To StarProbe.ChipCountY
                Wafer(xx, yy).flag = False
                Wafer(xx, yy).FlagBad = False
                Wafer(xx, yy).MeasureWait = False
                Wafer(xx, yy).InkDot = False
                Wafer(xx, yy).ChipMeasure = False ' 2005.09.05
                Wafer(xx, yy).BIN = 0             '16.12.02
            Next
        Next
        
        '==============================                     '[ 2020.09.14 ] : mapÀÌ¸§ Ã£±â ÇÔ¼ö
        Set fso = CreateObject("Scripting.FileSystemObject")
        If val(Left(LOT, 4)) < 2235 Then                            '[ 2022.09.08 ] : mapÀ» lot no¸¦ ±âÁØÀ¸·Î ±¸ºÐÇÑ´Ù. (lot no°¡ 2235ºÎÅÍ´Â ±âÁ¸°ú °°Àº Æú´õ¸¦ »ç¿ëÇÏ°í 2234±îÁö´Â ±¸ÇüÀº °£ÁÖ (old)Æú´õ¸¦ »ç¿ëÇÑ´Ù.)
            Set FolderList = fso.GetFolder(MAP_path & "(old)")      'old
        Else
            Set FolderList = fso.GetFolder(MAP_path)                'new
        End If
        Call Get_Folder(FolderList)
        '==============================
        
        If Load_MAP = "" Then
            MsgBox "File not found!", 16, "File error"
            Exit Sub
        End If
        If Dir(Load_MAP, vbDirectory) = "" Then             'map ÆÄÀÏÀÌ ¾ø´Â °æ¿ì
            MsgBox "File not found!", 16, "File error"
            Exit Sub
        End If
        
        '2019.01.02 : count clear add
        StarProbe.CountTotalChip = 0
        StarProbe.CountSkipDie = 0
        StarProbe.CountGoodDie = 0
        StarProbe.CountBadDie = 0
        
        RESET_DATA
        
        Fail_Loop = False
        Stop_Measure = False
        Command_Map_Clear.Enabled = True
        lblWafer.Caption = ""
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''   mapÀÌ¸§À» ¾î¶»°Ô ÇÒ °ÍÀÎÁö Á¤ÇØ¾ß ÇÑ´Ù.
        bStarprobe_AfterInk = False         'INK MODE OFF
        StarProbe.FileName_Map = Load_MAP
'        Load_MAP = MAP_path & "\" & DEV & ".map"
        StarProbeTemp = StarProbe
        
        Call StarProbe_FileLoad_ControlMap(Load_MAP)
        
        StarProbe.CountBadDie = StarProbeTemp.CountBadDie
        StarProbe.CountGoodDie = StarProbeTemp.CountGoodDie
        StarProbe.CountSkipDie = StarProbeTemp.CountSkipDie
        StarProbe.CountTotalChip = StarProbeTemp.CountTotalChip
        
        Call Command_DisplayWafer_Click
        
'                '[ 2020.04.06 ]
'                GOOD_COUNT_BACKUP = StarProbe.CountGoodDie
        
        StarProbe.MeasureStepX = StarProbeTemp.MeasureStepX
        StarProbe.MeasureStepY = StarProbeTemp.MeasureStepY
        
        StarProbe.Ink_LeftPort = StarProbeTemp.Ink_LeftPort
        StarProbe.Ink_RightPort = StarProbeTemp.Ink_RightPort
        StarProbe.LineOk = StarProbeTemp.LineOk
        StarProbe.RCount = StarProbeTemp.RCount
        StarProbe.RCount_Sub = StarProbeTemp.RCount_Sub
        StarProbe.MeasureSleep = StarProbeTemp.MeasureSleep
        StarProbe.ReMeasure = StarProbeTemp.ReMeasure
        StarProbe.Ink_After_LeftPort = StarProbeTemp.Ink_After_LeftPort
        StarProbe.Ink_After_RightPort = StarProbeTemp.Ink_After_RightPort
'        StarProbe.Ink_After_CenterPort = StarProbeTemp.Ink_After_CenterPort
        StarProbe.LimitArea = StarProbeTemp.LimitArea
        
        StarProbe.WaferTest = StarProbeTemp.WaferTest
        StarProbe.MeasureAll = StarProbeTemp.MeasureAll
        
        StarProbe.Ink_After = StarProbeTemp.Ink_After
        
        Me.MousePointer = 0
        Command2(1).Enabled = True
        
        SSTab3.Tab = 0
            
        For i = Len(Load_MAP) To 1 Step -1
            If Mid(Load_MAP, i, 1) = "\" Then
                SSPanel2(0).Caption = UCase(Mid(Load_MAP, i + 1, Len(Load_MAP) - i + 1))
                Exit For
            End If
        Next i
    End If
                
    Me.MousePointer = 0
    Command2(1).Enabled = True
    
    '[ 2021.01.11 ] : ¹æÇâ ¿É¼Ç ÃÊ±âÈ­
    Option_WaferDirection(0).value = False
    Option_WaferDirection(1).value = False
    Option_WaferDirection(2).value = False
    Option_WaferDirection(3).value = False
End Sub

Private Sub Command5_Click()
    Form_Cassette_New.Show vbModal
End Sub

Private Sub Command6_Click()
    Form_clean_tip.Show
    'Call StarProbe_tip_clean
    'Sleep 500
'    Call StarProbe_clean_tip_End_check
'    Sleep 500
'    Call Starprobe_Requst_State_tip_clean_error
End Sub

Private Sub Command7_Click()
    Command7.Enabled = False
    Call StarProbe_Z_Down
    Sleep 100
    Call StarProbe_Left_Ink_Dot(1)
    Sleep 100
    Command7.Enabled = True
End Sub

Private Sub Command8_Click()
    Call StarProbe_Set_Die_Size(val(Text8.Text), val(Text9.Text))
End Sub

Private Sub Command9_Click()
    Call Command_DisplayWafer_Click
End Sub

Private Sub Form_Load()
    Me.Move 0, 0
    
    If DemoMode = 0 Then
        MSComm1.RThreshold = 1
        MSComm1.CommPort = 1
        MSComm1.Settings = "9600,n,8,1"
        MSComm1.PortOpen = True
    
        Wafer_Start = False
    End If
        
    No_Probe = False
    Sample_No_Ink = False
    Crack_Wafer = False
    SSTab4.TabIndex = 0
    
    Barcode_Use = True              '[ 2020.09.17 ] : barcode use (default)
    
    MT2000.Caption = "SPC-2001X   " & "Version " & App.Major & "." & App.Minor & "." & App.Revision
    
    Call StarProbe_DefaultValue     '2005.07.05  star probe
    Call Bin_Teach
    LOOP_COUNT = Text11.Text
    TESTING_flag = False            '2016.03.11
    Sample_No_Ink = False           '2016.06.21 : sampling off
    Ink_Start_Flag = 0              '2016.09.27 : Ink Start Position Flag Clear
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Read_Needle_chk                 '[ 2022.07.29 ] : needle check data read
    Needle_Chk_Ok = False           '[ 2022.07.29 ] : needle check flag init
    Needle_check_flag = False       '[ 2022.08.31 ]
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If ACCO = 1 Or Model_Select = 1 Then                            'ACCO only
        GOOD_BIN_NO = 20
    Else
        GOOD_BIN_NO = 1
    End If
    
    If Tester_Select = 1 Or Model_Select = 2 Then
        CH_SET = 4                        '[ 2022.09.08 ] : AMT-88ÀÎ °æ¿ì 4Ã¤³Î °íÁ¤
        SelectExt.Frame11.Visible = False
    End If
End Sub

'[ 2022.07.29 ] : needle check ¼³Á¤ ³»¿ëÀ» ºÒ·¯¿Â´Ù.
Private Sub Read_Needle_chk()
    Dim i As Integer
    
    On Error GoTo errsub

    sfilename = "C:\star probe\needle_chk.dat"
    ifreefile = FreeFile
    
    Open sfilename For Input As ifreefile
        For i = 0 To 24
            Line Input #ifreefile, sLine
            Needle_Chk(i) = sLine
        Next i
    Close ifreefile
    
errsub:
    Resume Next
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer

    x = MsgBox("Do you want to exit system?", vbQuestion + vbYesNo, "System end")
    
    If x = 7 Then
        Cancel = 1
    Else
        If DemoMode = 0 Then
            MSComm1.PortOpen = False
            For i = 0 To 24
                BIN_Command(i) = Text_BinCommand(i).Text
            Next i
            Call StarProbe_FileSave_SystemInfo
        End If
        Unload MT2000
        End
    End If
End Sub

Private Sub HScroll_Zoom_Change()
    pZoom.Left = -(((pZoom.width - 12615) / 1000) * HScroll_Zoom.value)
End Sub

Private Sub HScroll_Zoom_Scroll()
    Call HScroll_Zoom_Change
End Sub

Private Sub MEdit_Click(Index As Integer)
    Select Case Index
        Case 0
            Command2_Click 0
        Case 1
            Command2_Click 2
    End Select
End Sub

Private Sub MT2000_END_Click()
    Unload Me
End Sub

Private Sub open_file_Click()
    Command2_Click 1
End Sub

Private Sub pOriginal_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Check1(3).value = vbChecked Then Exit Sub
    If StarProbe.DisplayOChipSizeX <= 0 Or StarProbe.DisplayOChipSizeY <= 0 Then Exit Sub

    Dim xx As Integer, yy As Integer
    
    xx = x \ StarProbe.DisplayOChipSizeX
    yy = y \ StarProbe.DisplayOChipSizeY
    
    If Wafer(xx, yy).Chip Then
        Shape_OMove.Top = (yy * StarProbe.DisplayOChipSizeY) - 1
        Shape_OMove.Left = (xx * StarProbe.DisplayOChipSizeX) - 1
    End If
End Sub

Private Sub pOriginal_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Check1(3).value = vbChecked Then Exit Sub
    If StarProbe.DisplayOChipSizeX <= 0 Or StarProbe.DisplayOChipSizeY <= 0 Then Exit Sub

    Dim xx As Integer, yy As Integer
    Dim iResult As Integer
    
    Dim ScrollX As Integer, ScrollY As Integer
    Dim KX As Long
    Dim KY As Long
    
    xx = x \ StarProbe.DisplayOChipSizeX
    yy = y \ StarProbe.DisplayOChipSizeY
    
    If Not Wafer(xx, yy).Chip Then Exit Sub
    
    Select Case Button
        Case 1
            KY = (yy / (Abs(StarProbe.Max.y) + Abs(StarProbe.Min.y) + 1)) * 1000
            KX = (xx / (Abs(StarProbe.Max.x) + Abs(StarProbe.Min.x) + 1)) * 1000
                        
            VScroll_Zoom.value = KY
            HScroll_Zoom.value = KX
            
            StarProbe.CurrentChip.x = xx - StarProbe.StartChip.x
            StarProbe.CurrentChip.y = yy - StarProbe.StartChip.y
            
            Text5 = StarProbe.CurrentChip.x
            Text6 = StarProbe.CurrentChip.y
            
            Shape_OChip.Top = (yy * StarProbe.DisplayOChipSizeY) - 1
            Shape_OChip.Left = (xx * StarProbe.DisplayOChipSizeX) - 1
            
            Shape_Chip.Top = (yy * StarProbe.DisplayChipSizeY) - 2
            Shape_Chip.Left = (xx * StarProbe.DisplayChipSizeX) - 2
            
            Label_ChipPosition = StarProbe.CurrentChip.x & "/" & StarProbe.CurrentChip.y
            
            XAxis = Text5
            YAxis = Text6
            
            Call StarProbe_XY_Moving((XAxis), (YAxis))
            
            If StarProbe_Motor_End_check Then MsgBox "Motor not end check !", 16, "STAR PROBE"
        
        Case 2
            If Ink_Start_Flag = 0 Then          'Ink Start PositionÀ» »ç¿ëÇÏÁö ¾Ê´Â °æ¿ì : Normal
                iResult = MsgBox("First Chip Ok ... ?", vbInformation + vbYesNo, "Information")     'first die
                If iResult = vbYes Then
                    Call StarProbe_First_Chip
                                            
                    StarProbe.StartChip.x = xx '- StarProbe.StartChip.X     '186
                    StarProbe.StartChip.y = yy '- StarProbe.StartChip.Y     '15

                    '[ 2021.06.25 ] : first die±â¾ï
                    First_X = StarProbe.StartChip.x
                    First_Y = StarProbe.StartChip.y
                    
                    Shape_FirstChip.Top = (yy * StarProbe.DisplayChipSizeY) - 2
                    First_Zoom_TOP = (yy * StarProbe.DisplayChipSizeY) - 2
                    Shape_FirstChip.Left = (xx * StarProbe.DisplayChipSizeX) - 2
                    First_Zoom_LEFT = (xx * StarProbe.DisplayChipSizeX) - 2
            
                    Shape_OFirstChip.Top = (yy * StarProbe.DisplayOChipSizeY) - 1
                    First_Original_TOP = (yy * StarProbe.DisplayOChipSizeY) - 1
                    Shape_OFirstChip.Left = (xx * StarProbe.DisplayOChipSizeX) - 1
                    First_Original_LEFT = (xx * StarProbe.DisplayOChipSizeX) - 1
                End If
            Else                                'Ink Start PositionÀ» »ç¿ëÇÏ´Â °æ¿ì : 2016.09.27
                iResult = MsgBox("Ink Start Position Chip Ok ... ?", vbInformation + vbYesNo, "Information")     'Ink Start Chip
                If iResult = vbYes Then
                    Label_Ink = (xx - StarProbe.StartChip.x) & "/" & (yy - StarProbe.StartChip.y)
                                                                                                
                    StarProbe.InkStart.x = xx
                    StarProbe.InkStart.y = yy
                    
                    Shape_Ink.Top = (yy * StarProbe.DisplayChipSizeY) - 2     '88
                    Shape_Ink.Left = (xx * StarProbe.DisplayChipSizeX) - 2       '1114
                
                    Shape_OInk.Top = (yy * StarProbe.DisplayOChipSizeY) - 1   '14
                    Shape_OInk.Left = (xx * StarProbe.DisplayOChipSizeX) - 1  '185
                End If
            End If
    End Select
End Sub

Private Sub pZoom_KeyUp(KeyCode As Integer, Shift As Integer)
'    If Mode_Set = False Then Exit Sub                   '[ 2021.12.31 ] : false:Operator, true:Engineer
    If Check1(3).value = vbChecked Then Exit Sub
    If StarProbe.DisplayChipSizeX <= 0 Or StarProbe.DisplayChipSizeY <= 0 Then Exit Sub

    Dim forx As Integer, fory As Integer
    
    Dim xx As Integer, yy As Integer
    Dim b As Boolean
    
    xx = StarProbe.CurrentChip.x + StarProbe.StartChip.x
    yy = StarProbe.CurrentChip.y + StarProbe.StartChip.y

    b = False
    
    Select Case KeyCode
        '[ 2022.08.31 ] : esc ´©¸¦°æ¿ì BIN17·Î º¯°æ
        Case 27                 'esc
            If Wafer(xx, yy).flag Then
                If Wafer(xx, yy).ChipSkipDie Or _
                    Wafer(xx, yy).ChipInk Or _
                    Wafer(xx, yy).ChipPlate Or _
                    Wafer(xx, yy).ChipMask Then
                    StarProbe.CountSkipDie = StarProbe.CountSkipDie - 1
                    StarProbe.CountGoodDie = StarProbe.CountGoodDie + 1
                End If
            
                If Not Wafer(xx, yy).FlagBad Then
                    Wafer(xx, yy).ChipSkipDie = False
                    Wafer(xx, yy).ChipInk = False
                    Wafer(xx, yy).ChipPlate = False
                    Wafer(xx, yy).ChipMask = False
                
                    Wafer(xx, yy).flag = True
                    Wafer(xx, yy).FlagBad = True
                    
                    Text_Bin_Count_No(Wafer(xx, yy).BIN) = Text_Bin_Count_No(Wafer(xx, yy).BIN) - 1
                    Text_Bin_Count_No(17) = Text_Bin_Count_No(17) + 1
                    Wafer(xx, yy).BIN = 17
                    StarProbe.CountGoodDie = StarProbe.CountGoodDie - 1
                    StarProbe.CountBadDie = StarProbe.CountBadDie + 1
                    
                    Good_Cnt = Good_Cnt - 1
                    SSPanel_SkipCount.Caption = StarProbe.CountSkipDie
                    Text_SkipCount.Text = StarProbe.CountSkipDie
                    SSPanel_BadCount.Caption = StarProbe.CountBadDie
                    Text_BadCount.Text = StarProbe.CountBadDie
                    SSPanel_GoodCount.Caption = StarProbe.CountGoodDie
                
                    Call Display_Chip(pZoom, pOriginal, (StarProbe.CurrentChip.x), (StarProbe.CurrentChip.y))
                End If
            End If
            Call Bin_Teach
            SSPanel2(1).Caption = Test_Cnt
            Text_TotalCount.Text = Test_Cnt
            
            If Test_Cnt = 0 Then
                SSPanel2(2).Caption = Good_Cnt & Space(1) & "(" & "0.00" & "%)"
                Text_GoodCount.Text = Good_Cnt
            Else
                SSPanel2(2).Caption = Good_Cnt & Space(1) & "(" & Format(Good_Cnt / Test_Cnt * 100, "0.00") & "%)"
                Text_GoodCount.Text = Good_Cnt
            End If
        Case 49, 83, 50, 77, 51, 80, 52, 90, 53, 88, 48, 66, 32, 77, 78, 73, 79, 219, 221, 67, 71, 72       '[ 2017.03.27 ] : normal chip left, rightÃß°¡(71,72)
            UndoX = xx
            UndoY = yy
        
            Erase UndoWafer
            
            For forx = 0 To 900
                UndoWafer(forx) = Wafer(forx, yy)
            Next
            
            UndoCountGoodDie = StarProbe.CountGoodDie
            UndoCountSkipDie = StarProbe.CountSkipDie
            UndoCountBadDie = StarProbe.CountBadDie
            
        Case 85  ' U
            For forx = 0 To 900
                Wafer(forx, UndoY) = UndoWafer(forx)
                Call Display_Chip(pZoom, pOriginal, (forx - StarProbe.StartChip.x), (UndoY - StarProbe.StartChip.y))
            Next
            
            StarProbe.CountGoodDie = UndoCountGoodDie
            StarProbe.CountSkipDie = UndoCountSkipDie
            StarProbe.CountBadDie = UndoCountBadDie
        
            SSPanel_GoodCount.Caption = StarProbe.CountGoodDie
            SSPanel_SkipCount.Caption = StarProbe.CountSkipDie
            Text_SkipCount.Text = StarProbe.CountSkipDie
         
            SSPanel_BadCount.Caption = StarProbe.CountBadDie
            Text_BadCount.Text = StarProbe.CountBadDie
    End Select
    
    Select Case KeyCode
        Case 49, 83 ' #1, S - Skip Die
            If Wafer(xx, yy).ChipSkipDie Or _
               Wafer(xx, yy).ChipInk Or _
               Wafer(xx, yy).ChipPlate Or _
               Wafer(xx, yy).ChipMask Then
                StarProbe.CountSkipDie = StarProbe.CountSkipDie - 1
                StarProbe.CountGoodDie = StarProbe.CountGoodDie + 1
            End If
            
            If Wafer(xx, yy).FlagBad Then
                StarProbe.CountBadDie = StarProbe.CountBadDie - 1
                StarProbe.CountGoodDie = StarProbe.CountGoodDie + 1
            End If
            
            Wafer(xx, yy).ChipSkipDie = True
            Wafer(xx, yy).ChipInk = False
            Wafer(xx, yy).ChipInk2 = False
            Wafer(xx, yy).ChipPlate = False
            Wafer(xx, yy).ChipMask = False
            
            Wafer(xx, yy).flag = False
            Wafer(xx, yy).FlagBad = False
            
            StarProbe.CountSkipDie = StarProbe.CountSkipDie + 1
            StarProbe.CountGoodDie = StarProbe.CountGoodDie - 1
            
            SSPanel_SkipCount.Caption = StarProbe.CountSkipDie
            Text_SkipCount.Text = StarProbe.CountSkipDie
            SSPanel_BadCount.Caption = StarProbe.CountBadDie
            Text_BadCount.Text = StarProbe.CountBadDie
            SSPanel_GoodCount.Caption = StarProbe.CountGoodDie
            
            Call Display_Chip(pZoom, pOriginal, (StarProbe.CurrentChip.x), (StarProbe.CurrentChip.y))
        
        Case 50 ' #2 - Mask
            If Wafer(xx, yy).ChipSkipDie Or _
               Wafer(xx, yy).ChipInk Or _
               Wafer(xx, yy).ChipPlate Or _
               Wafer(xx, yy).ChipMask Then
                StarProbe.CountSkipDie = StarProbe.CountSkipDie - 1
                StarProbe.CountGoodDie = StarProbe.CountGoodDie + 1
            End If
            
            If Wafer(xx, yy).FlagBad Then
                StarProbe.CountBadDie = StarProbe.CountBadDie - 1
                StarProbe.CountGoodDie = StarProbe.CountGoodDie + 1
            End If
                
            Wafer(xx, yy).ChipSkipDie = False
            Wafer(xx, yy).ChipInk = False
            Wafer(xx, yy).ChipInk2 = False
            Wafer(xx, yy).ChipPlate = False
            Wafer(xx, yy).ChipMask = True
            
            Wafer(xx, yy).flag = False
            Wafer(xx, yy).FlagBad = False
            
            StarProbe.CountSkipDie = StarProbe.CountSkipDie + 1
            StarProbe.CountGoodDie = StarProbe.CountGoodDie - 1
            
            SSPanel_SkipCount.Caption = StarProbe.CountSkipDie
            Text_SkipCount.Text = StarProbe.CountSkipDie
            SSPanel_BadCount.Caption = StarProbe.CountBadDie
            Text_BadCount.Text = StarProbe.CountBadDie
            SSPanel_GoodCount.Caption = StarProbe.CountGoodDie
            
            Call Display_Chip(pZoom, pOriginal, (StarProbe.CurrentChip.x), (StarProbe.CurrentChip.y))
        
        Case 51, 80 ' #3, P - Plate Zone
            If Wafer(xx, yy).ChipSkipDie Or _
               Wafer(xx, yy).ChipInk Or _
               Wafer(xx, yy).ChipPlate Or _
               Wafer(xx, yy).ChipMask Then
                StarProbe.CountSkipDie = StarProbe.CountSkipDie - 1
                StarProbe.CountGoodDie = StarProbe.CountGoodDie + 1
            End If
            
            If Wafer(xx, yy).FlagBad Then
                StarProbe.CountBadDie = StarProbe.CountBadDie - 1
                StarProbe.CountGoodDie = StarProbe.CountGoodDie + 1
            End If
                
            Wafer(xx, yy).ChipSkipDie = False
            Wafer(xx, yy).ChipInk = False
            Wafer(xx, yy).ChipInk2 = False
            Wafer(xx, yy).ChipPlate = True
            Wafer(xx, yy).ChipMask = False
            
            Wafer(xx, yy).flag = False
            Wafer(xx, yy).FlagBad = False
            
            StarProbe.CountSkipDie = StarProbe.CountSkipDie + 1
            StarProbe.CountGoodDie = StarProbe.CountGoodDie - 1
            
            SSPanel_SkipCount.Caption = StarProbe.CountSkipDie
            Text_SkipCount.Text = StarProbe.CountSkipDie
            SSPanel_BadCount.Caption = StarProbe.CountBadDie
            Text_BadCount.Text = StarProbe.CountBadDie
            SSPanel_GoodCount.Caption = StarProbe.CountGoodDie
            
            Call Display_Chip(pZoom, pOriginal, (StarProbe.CurrentChip.x), (StarProbe.CurrentChip.y))
            
        Case 54, 81 ' #4, Q - Plate Die (Left)
            For forx = xx To 0 Step -1
                If Wafer(forx, yy).Chip Then
                    If Wafer(forx, yy).ChipSkipDie Or _
                       Wafer(forx, yy).ChipInk Or _
                       Wafer(forx, yy).ChipPlate Or _
                       Wafer(forx, yy).ChipMask Then
                        StarProbe.CountSkipDie = StarProbe.CountSkipDie - 1
                        StarProbe.CountGoodDie = StarProbe.CountGoodDie + 1
                    End If
                    
                    If Wafer(forx, yy).FlagBad Then
                        StarProbe.CountBadDie = StarProbe.CountBadDie - 1
                        StarProbe.CountGoodDie = StarProbe.CountGoodDie + 1
                    End If
                    
                    Wafer(forx, yy).ChipSkipDie = False
                    Wafer(forx, yy).ChipInk = False
                    Wafer(forx, yy).ChipInk2 = False
                    Wafer(forx, yy).ChipPlate = True
                    Wafer(forx, yy).ChipMask = False
                    
                    Wafer(forx, yy).flag = False
                    Wafer(forx, yy).FlagBad = False
                    Wafer(forx, yy).BIN = 0
                    
                    Call Display_Chip(pZoom, pOriginal, (forx - StarProbe.StartChip.x), (StarProbe.CurrentChip.y))
                    
                    StarProbe.CountGoodDie = StarProbe.CountGoodDie - 1
                    StarProbe.CountSkipDie = StarProbe.CountSkipDie + 1
        
                    SSPanel_SkipCount.Caption = StarProbe.CountSkipDie
                    SSPanel_BadCount.Caption = StarProbe.CountBadDie
                    SSPanel_GoodCount.Caption = StarProbe.CountGoodDie
                End If
            Next
        
        Case 55, 87 ' #5, W - Plate Die (Right)
            For forx = xx To (StarProbe.ChipCountX - 1)
                If Wafer(forx, yy).Chip Then
                    If Wafer(forx, yy).ChipSkipDie Or _
                       Wafer(forx, yy).ChipInk Or _
                       Wafer(forx, yy).ChipPlate Or _
                       Wafer(forx, yy).ChipMask Then
                        StarProbe.CountSkipDie = StarProbe.CountSkipDie - 1
                        StarProbe.CountGoodDie = StarProbe.CountGoodDie + 1
                    End If
                    
                    If Wafer(forx, yy).FlagBad Then
                        StarProbe.CountBadDie = StarProbe.CountBadDie - 1
                        StarProbe.CountGoodDie = StarProbe.CountGoodDie + 1
                    End If
                    
                    Wafer(forx, yy).ChipSkipDie = False
                    Wafer(forx, yy).ChipInk = False
                    Wafer(forx, yy).ChipInk2 = False
                    Wafer(forx, yy).ChipPlate = True
                    Wafer(forx, yy).ChipMask = False
                    
                    Wafer(forx, yy).flag = False
                    Wafer(forx, yy).FlagBad = False
                    Wafer(forx, yy).BIN = 0
                    
                    Call Display_Chip(pZoom, pOriginal, (forx - StarProbe.StartChip.x), (StarProbe.CurrentChip.y))
                    
                    StarProbe.CountGoodDie = StarProbe.CountGoodDie - 1
                    StarProbe.CountSkipDie = StarProbe.CountSkipDie + 1
        
                    SSPanel_SkipCount.Caption = StarProbe.CountSkipDie
                    SSPanel_BadCount.Caption = StarProbe.CountBadDie
                    SSPanel_GoodCount.Caption = StarProbe.CountGoodDie
                End If
            Next
                
        Case 52, 90 ' #4, Z - Skip Die (Left)
            For forx = xx To 0 Step -1
                If Wafer(forx, yy).Chip Then
                    If Wafer(forx, yy).ChipSkipDie Or _
                       Wafer(forx, yy).ChipInk Or _
                       Wafer(forx, yy).ChipPlate Or _
                       Wafer(forx, yy).ChipMask Then
                        StarProbe.CountSkipDie = StarProbe.CountSkipDie - 1
                        StarProbe.CountGoodDie = StarProbe.CountGoodDie + 1
                    End If
                    
                    If Wafer(forx, yy).FlagBad Then
                        StarProbe.CountBadDie = StarProbe.CountBadDie - 1
                        StarProbe.CountGoodDie = StarProbe.CountGoodDie + 1
                    End If
                    
                    Wafer(forx, yy).ChipSkipDie = True
                    Wafer(forx, yy).ChipInk = False
                    Wafer(forx, yy).ChipInk2 = False
                    Wafer(forx, yy).ChipPlate = False
                    Wafer(forx, yy).ChipMask = False
                    
                    Wafer(forx, yy).flag = False
                    Wafer(forx, yy).FlagBad = False
                    Wafer(forx, yy).BIN = 0
                    
                    Call Display_Chip(pZoom, pOriginal, (forx - StarProbe.StartChip.x), (StarProbe.CurrentChip.y))
                    
                    StarProbe.CountGoodDie = StarProbe.CountGoodDie - 1
                    StarProbe.CountSkipDie = StarProbe.CountSkipDie + 1
        
                    SSPanel_SkipCount.Caption = StarProbe.CountSkipDie
                    Text_SkipCount.Text = StarProbe.CountSkipDie
                    SSPanel_BadCount.Caption = StarProbe.CountBadDie
                    Text_BadCount.Text = StarProbe.CountBadDie
                    SSPanel_GoodCount.Caption = StarProbe.CountGoodDie
                End If
            Next
        
        Case 53, 88 ' #5, X - Skip Die (Right)
            For forx = xx To (StarProbe.ChipCountX - 1)
                If Wafer(forx, yy).Chip Then
                    If Wafer(forx, yy).ChipSkipDie Or _
                       Wafer(forx, yy).ChipInk Or _
                       Wafer(forx, yy).ChipPlate Or _
                       Wafer(forx, yy).ChipMask Then
                        StarProbe.CountSkipDie = StarProbe.CountSkipDie - 1
                        StarProbe.CountGoodDie = StarProbe.CountGoodDie + 1
                    End If
                    
                    If Wafer(forx, yy).FlagBad Then
                        StarProbe.CountBadDie = StarProbe.CountBadDie - 1
                        StarProbe.CountGoodDie = StarProbe.CountGoodDie + 1
                    End If
                    
                    Wafer(forx, yy).ChipSkipDie = True
                    Wafer(forx, yy).ChipInk = False
                    Wafer(forx, yy).ChipInk2 = False
                    Wafer(forx, yy).ChipPlate = False
                    Wafer(forx, yy).ChipMask = False
                    
                    Wafer(forx, yy).flag = False
                    Wafer(forx, yy).FlagBad = False
                    Wafer(forx, yy).BIN = 0
                    
                    Call Display_Chip(pZoom, pOriginal, (forx - StarProbe.StartChip.x), (StarProbe.CurrentChip.y))
                    
                    StarProbe.CountGoodDie = StarProbe.CountGoodDie - 1
                    StarProbe.CountSkipDie = StarProbe.CountSkipDie + 1
        
                    SSPanel_SkipCount.Caption = StarProbe.CountSkipDie
                    Text_SkipCount.Text = StarProbe.CountSkipDie
                    SSPanel_BadCount.Caption = StarProbe.CountBadDie
                    Text_BadCount.Text = StarProbe.CountBadDie
                    SSPanel_GoodCount.Caption = StarProbe.CountGoodDie
                End If
            Next
            
        Case 54 ' #6 - Mask (Left)
                    
        Case 55 ' #7 - Mask (Right)
                        
        Case 56 ' #8
        Case 57 ' #9
        
        Case 48, 66 ' #0, B - Bad Flag
            If Wafer(xx, yy).ChipSkipDie Or _
               Wafer(xx, yy).ChipInk Or _
               Wafer(xx, yy).ChipPlate Or _
               Wafer(xx, yy).ChipMask Then
                StarProbe.CountSkipDie = StarProbe.CountSkipDie - 1
                StarProbe.CountGoodDie = StarProbe.CountGoodDie + 1
            End If
            
            If Not Wafer(xx, yy).FlagBad Then
                Wafer(xx, yy).ChipSkipDie = False
                Wafer(xx, yy).ChipInk = False
                Wafer(xx, yy).ChipInk2 = False
                Wafer(xx, yy).ChipPlate = False
                Wafer(xx, yy).ChipMask = False
                
                Wafer(xx, yy).flag = True
                Wafer(xx, yy).FlagBad = True
                Wafer(xx, yy).BIN = 0
                StarProbe.CountGoodDie = StarProbe.CountGoodDie - 1
                StarProbe.CountBadDie = StarProbe.CountBadDie + 1
                
                SSPanel_SkipCount.Caption = StarProbe.CountSkipDie
                Text_SkipCount.Text = StarProbe.CountSkipDie
                SSPanel_BadCount.Caption = StarProbe.CountBadDie
                Text_BadCount.Text = StarProbe.CountBadDie
                SSPanel_GoodCount.Caption = StarProbe.CountGoodDie
                
                Call Display_Chip(pZoom, pOriginal, (StarProbe.CurrentChip.x), (StarProbe.CurrentChip.y))
            End If
        
        'Case 78 ' N - BAD Die (Left)
        Case 188 ' , - BAD Die (Left)
            For forx = xx To 0 Step -1
                If Wafer(forx, yy).Chip Then
                    If Wafer(forx, yy).ChipSkipDie Or _
                       Wafer(forx, yy).ChipInk Or _
                       Wafer(forx, yy).ChipPlate Or _
                       Wafer(forx, yy).ChipMask Then
                        StarProbe.CountSkipDie = StarProbe.CountSkipDie - 1
                        StarProbe.CountGoodDie = StarProbe.CountGoodDie + 1
                    End If
                    
                    If Not Wafer(forx, yy).FlagBad Then
                        StarProbe.CountBadDie = StarProbe.CountBadDie + 1
                        StarProbe.CountGoodDie = StarProbe.CountGoodDie - 1
                                    
                        Wafer(forx, yy).ChipSkipDie = False
                        Wafer(forx, yy).ChipInk = False
                        Wafer(forx, yy).ChipInk2 = False
                        Wafer(forx, yy).ChipPlate = False
                        Wafer(forx, yy).ChipMask = False
                        
                        Wafer(forx, yy).flag = True
                        Wafer(forx, yy).FlagBad = True
                        Wafer(forx, yy).BIN = 0
                        
                        Call Display_Chip(pZoom, pOriginal, (forx - StarProbe.StartChip.x), (StarProbe.CurrentChip.y))
                    End If
                    SSPanel_SkipCount.Caption = StarProbe.CountSkipDie
                    Text_SkipCount.Text = StarProbe.CountSkipDie
                    SSPanel_BadCount.Caption = StarProbe.CountBadDie
                    Text_BadCount.Text = StarProbe.CountBadDie
                    SSPanel_GoodCount.Caption = StarProbe.CountGoodDie
                End If
            Next
            
        'Case 77 ' M - BAD Die (Right)
        Case 190 ' . - BAD Die (Right)
            For forx = xx To (StarProbe.ChipCountX - 1)
                If Wafer(forx, yy).Chip Then
                    If Wafer(forx, yy).ChipSkipDie Or _
                       Wafer(forx, yy).ChipInk Or _
                       Wafer(forx, yy).ChipPlate Or _
                       Wafer(forx, yy).ChipMask Then
                        StarProbe.CountSkipDie = StarProbe.CountSkipDie - 1
                        StarProbe.CountGoodDie = StarProbe.CountGoodDie + 1
                    End If
                    
                    If Not Wafer(forx, yy).FlagBad Then
                        StarProbe.CountBadDie = StarProbe.CountBadDie + 1
                        StarProbe.CountGoodDie = StarProbe.CountGoodDie - 1
                                    
                        Wafer(forx, yy).ChipSkipDie = False
                        Wafer(forx, yy).ChipInk = False
                        Wafer(forx, yy).ChipInk2 = False
                        Wafer(forx, yy).ChipPlate = False
                        Wafer(forx, yy).ChipMask = False
                        
                        Wafer(forx, yy).flag = True
                        Wafer(forx, yy).FlagBad = True
                        Wafer(forx, yy).BIN = 0
                        
                        Call Display_Chip(pZoom, pOriginal, (forx - StarProbe.StartChip.x), (StarProbe.CurrentChip.y))
                    End If
                    SSPanel_SkipCount.Caption = StarProbe.CountSkipDie
                    Text_SkipCount.Text = StarProbe.CountSkipDie
                    SSPanel_BadCount.Caption = StarProbe.CountBadDie
                    Text_BadCount.Text = StarProbe.CountBadDie
                    SSPanel_GoodCount.Caption = StarProbe.CountGoodDie
                End If
            Next
            
        Case 73 ' I - Ink1 Die
            If Wafer(xx, yy).ChipSkipDie Or _
               Wafer(xx, yy).ChipPlate Or _
               Wafer(xx, yy).ChipInk Or _
               Wafer(xx, yy).ChipMask Then
                StarProbe.CountSkipDie = StarProbe.CountSkipDie - 1
                StarProbe.CountGoodDie = StarProbe.CountGoodDie + 1
            End If
            
            If Wafer(xx, yy).FlagBad Then
                StarProbe.CountBadDie = StarProbe.CountBadDie - 1
                StarProbe.CountGoodDie = StarProbe.CountGoodDie + 1
            End If
                        
            If Wafer(xx, yy).flag Then
                Wafer(xx, yy).InkDot = False
            Else
                Wafer(xx, yy).ChipSkipDie = True
                Wafer(xx, yy).ChipInk = True
                Wafer(xx, yy).ChipInk2 = True
                Wafer(xx, yy).ChipInk2 = False
                Wafer(xx, yy).ChipPlate = False
                Wafer(xx, yy).ChipMask = False
            
                Wafer(xx, yy).flag = False
                Wafer(xx, yy).FlagBad = False
            End If
            
            StarProbe.CountSkipDie = StarProbe.CountSkipDie + 1
            StarProbe.CountGoodDie = StarProbe.CountGoodDie - 1
            
            SSPanel_SkipCount.Caption = StarProbe.CountSkipDie
            Text_SkipCount.Text = StarProbe.CountSkipDie
            SSPanel_BadCount.Caption = StarProbe.CountBadDie
            Text_BadCount.Text = StarProbe.CountBadDie
            SSPanel_GoodCount.Caption = StarProbe.CountGoodDie
            
            Call Display_Chip(pZoom, pOriginal, (StarProbe.CurrentChip.x), (StarProbe.CurrentChip.y))
            
        Case 79 ' O - Ink2 Die
            If Wafer(xx, yy).ChipSkipDie Or _
               Wafer(xx, yy).ChipPlate Or _
               Wafer(xx, yy).ChipInk Or _
               Wafer(xx, yy).ChipMask Then
                StarProbe.CountSkipDie = StarProbe.CountSkipDie - 1
                StarProbe.CountGoodDie = StarProbe.CountGoodDie + 1
            End If
            
            If Wafer(xx, yy).FlagBad Then
                StarProbe.CountBadDie = StarProbe.CountBadDie - 1
                StarProbe.CountGoodDie = StarProbe.CountGoodDie + 1
            End If
                        
            If Wafer(xx, yy).flag Then
                Wafer(xx, yy).InkDot = False
            Else
                Wafer(xx, yy).ChipSkipDie = True
                Wafer(xx, yy).ChipInk = True
                Wafer(xx, yy).ChipInk2 = True
                Wafer(xx, yy).ChipPlate = False
                Wafer(xx, yy).ChipMask = False
            
                Wafer(xx, yy).flag = False
                Wafer(xx, yy).FlagBad = False
            End If
            
            StarProbe.CountSkipDie = StarProbe.CountSkipDie + 1
            StarProbe.CountGoodDie = StarProbe.CountGoodDie - 1
            
            SSPanel_SkipCount.Caption = StarProbe.CountSkipDie
            Text_SkipCount.Text = StarProbe.CountSkipDie
            SSPanel_BadCount.Caption = StarProbe.CountBadDie
            Text_BadCount.Text = StarProbe.CountBadDie
            SSPanel_GoodCount.Caption = StarProbe.CountGoodDie
            
            Call Display_Chip(pZoom, pOriginal, (StarProbe.CurrentChip.x), (StarProbe.CurrentChip.y))
            
        Case 219 ' [ - Ink Die (Left)
            For forx = xx To 0 Step -1
                If Wafer(forx, yy).Chip Then
                    If Wafer(forx, yy).ChipSkipDie Or _
                       Wafer(forx, yy).ChipInk Or _
                       Wafer(forx, yy).ChipPlate Or _
                       Wafer(forx, yy).ChipMask Then
                        StarProbe.CountSkipDie = StarProbe.CountSkipDie - 1
                        StarProbe.CountGoodDie = StarProbe.CountGoodDie + 1
                    End If
                    
                    If Wafer(forx, yy).FlagBad Then
                        StarProbe.CountBadDie = StarProbe.CountBadDie - 1
                        StarProbe.CountGoodDie = StarProbe.CountGoodDie + 1
                    End If
                    
                    Wafer(forx, yy).ChipSkipDie = True
                    Wafer(forx, yy).ChipInk = True
                    Wafer(forx, yy).ChipInk2 = False
                    Wafer(forx, yy).ChipPlate = False
                    Wafer(forx, yy).ChipMask = False
                    
                    Wafer(forx, yy).flag = False
                    Wafer(forx, yy).FlagBad = False
                    Wafer(forx, yy).BIN = 0
                    
                    Call Display_Chip(pZoom, pOriginal, (forx - StarProbe.StartChip.x), (StarProbe.CurrentChip.y))
                    
                    StarProbe.CountGoodDie = StarProbe.CountGoodDie - 1
                    StarProbe.CountSkipDie = StarProbe.CountSkipDie + 1
        
                    SSPanel_SkipCount.Caption = StarProbe.CountSkipDie
                    Text_SkipCount.Text = StarProbe.CountSkipDie
                    SSPanel_BadCount.Caption = StarProbe.CountBadDie
                    Text_BadCount.Text = StarProbe.CountBadDie
                    SSPanel_GoodCount.Caption = StarProbe.CountGoodDie
                End If
            Next
        
        Case 221 ' ] - Ink Die (Right)
            For forx = xx To (StarProbe.ChipCountX - 1)
                If Wafer(forx, yy).Chip Then
                    If Wafer(forx, yy).ChipSkipDie Or _
                       Wafer(forx, yy).ChipInk Or _
                       Wafer(forx, yy).ChipPlate Or _
                       Wafer(forx, yy).ChipMask Then
                        StarProbe.CountSkipDie = StarProbe.CountSkipDie - 1
                        StarProbe.CountGoodDie = StarProbe.CountGoodDie + 1
                    End If
                    
                    If Wafer(forx, yy).FlagBad Then
                        StarProbe.CountBadDie = StarProbe.CountBadDie - 1
                        StarProbe.CountGoodDie = StarProbe.CountGoodDie + 1
                    End If
                    
                    Wafer(forx, yy).ChipSkipDie = True
                    Wafer(forx, yy).ChipInk = True
                    Wafer(forx, yy).ChipInk2 = False
                    Wafer(forx, yy).ChipPlate = False
                    Wafer(forx, yy).ChipMask = False
                    
                    Wafer(forx, yy).flag = False
                    Wafer(forx, yy).FlagBad = False
                    Wafer(forx, yy).BIN = 0
                    
                    Call Display_Chip(pZoom, pOriginal, (forx - StarProbe.StartChip.x), (StarProbe.CurrentChip.y))
                    
                    StarProbe.CountGoodDie = StarProbe.CountGoodDie - 1
                    StarProbe.CountSkipDie = StarProbe.CountSkipDie + 1
        
                    SSPanel_SkipCount.Caption = StarProbe.CountSkipDie
                    Text_SkipCount.Text = StarProbe.CountSkipDie
                    SSPanel_BadCount.Caption = StarProbe.CountBadDie
                    Text_BadCount.Text = StarProbe.CountBadDie
                    SSPanel_GoodCount.Caption = StarProbe.CountGoodDie
                End If
            Next
            
        Case 67 ' C - Chip
            If Wafer(xx, yy).Chip Then
                Wafer(xx, yy).Chip = False
                Wafer(xx, yy).ChipMask = False
                Wafer(xx, yy).ChipMeasure = False
                Wafer(xx, yy).ChipSkipDie = False
                Wafer(xx, yy).ChipPlate = False
                Wafer(xx, yy).ChipInk = False
                Wafer(xx, yy).ChipInk2 = False
                Wafer(xx, yy).BIN = False
                Wafer(xx, yy).flag = False
                Wafer(xx, yy).FlagBad = False
                Wafer(xx, yy).MeasureWait = False
                Wafer(xx, yy).InkDot = False
                
                If Wafer(xx, yy).ChipSkipDie Or _
                   Wafer(xx, yy).ChipPlate Or _
                   Wafer(xx, yy).ChipMask Then
                    StarProbe.CountSkipDie = StarProbe.CountSkipDie - 1
                End If
                StarProbe.CountGoodDie = StarProbe.CountGoodDie - 1
            Else
                Wafer(xx, yy).Chip = True
                Wafer(xx, yy).ChipMask = False
                Wafer(xx, yy).ChipMeasure = False
                Wafer(xx, yy).ChipSkipDie = False
                Wafer(xx, yy).ChipPlate = False
                Wafer(xx, yy).ChipInk = False
                Wafer(xx, yy).ChipInk2 = False
                Wafer(xx, yy).BIN = False
                Wafer(xx, yy).flag = False
                Wafer(xx, yy).FlagBad = False
                Wafer(xx, yy).MeasureWait = False
                Wafer(xx, yy).InkDot = False
                StarProbe.CountGoodDie = StarProbe.CountGoodDie + 1
            End If
            
            SSPanel_SkipCount.Caption = StarProbe.CountSkipDie
            Text_SkipCount.Text = StarProbe.CountSkipDie
            SSPanel_BadCount.Caption = StarProbe.CountBadDie
            Text_BadCount.Text = StarProbe.CountBadDie
            SSPanel_GoodCount.Caption = StarProbe.CountGoodDie
    
            Call Display_Chip(pZoom, pOriginal, (StarProbe.CurrentChip.x), (StarProbe.CurrentChip.y))
            
        Case 32 ' Space Bar - Normal
            If Wafer(xx, yy).ChipSkipDie Or _
               Wafer(xx, yy).ChipPlate Or _
               Wafer(xx, yy).ChipMask Then
                StarProbe.CountSkipDie = StarProbe.CountSkipDie - 1
                StarProbe.CountGoodDie = StarProbe.CountGoodDie + 1
            End If
            
            If Wafer(xx, yy).FlagBad Then
                StarProbe.CountBadDie = StarProbe.CountBadDie - 1
                StarProbe.CountGoodDie = StarProbe.CountGoodDie + 1
            End If
            
            SSPanel_SkipCount.Caption = StarProbe.CountSkipDie
            Text_SkipCount.Text = StarProbe.CountSkipDie
            SSPanel_BadCount.Caption = StarProbe.CountBadDie
            Text_BadCount.Text = StarProbe.CountBadDie
            SSPanel_GoodCount.Caption = StarProbe.CountGoodDie
            
            Wafer(xx, yy).ChipSkipDie = False
            Wafer(xx, yy).ChipInk = False
            Wafer(xx, yy).ChipInk2 = False
            Wafer(xx, yy).ChipPlate = False
            Wafer(xx, yy).ChipMask = False
            
            Wafer(xx, yy).flag = False
            Wafer(xx, yy).FlagBad = False
            '[ 2020.11.11 ] : Á¤»óÄ¨À¸·Î ¸¸µé°æ¿ì ±âÁ¸ÀÇ bin Ä«¿îÆ®¸¦ Á¦°ÅÇÑ´Ù.
            Bin_Count(Wafer(xx, yy).BIN) = Bin_Count(Wafer(xx, yy).BIN) - 1
            Wafer(xx, yy).BIN = 0
            
            Call Display_Chip(pZoom, pOriginal, (StarProbe.CurrentChip.x), (StarProbe.CurrentChip.y))
            
        Case 71 ' V - Normal Die (Left)     '[ 2017.03.27 ] : normal chip left, rightÃß°¡(71,72)
            For forx = xx To 0 Step -1
                If Wafer(forx, yy).Chip Then
                    If Wafer(forx, yy).ChipSkipDie Or _
                       Wafer(forx, yy).ChipPlate Or _
                       Wafer(forx, yy).ChipMask Then
                        StarProbe.CountSkipDie = StarProbe.CountSkipDie - 1
                        StarProbe.CountGoodDie = StarProbe.CountGoodDie + 1
                    End If
                    
                    If Wafer(forx, yy).FlagBad Then
                        StarProbe.CountBadDie = StarProbe.CountBadDie - 1
                        StarProbe.CountGoodDie = StarProbe.CountGoodDie + 1
                    End If
                    
                    SSPanel_SkipCount.Caption = StarProbe.CountSkipDie
                    Text_SkipCount.Text = StarProbe.CountSkipDie
                    SSPanel_BadCount.Caption = StarProbe.CountBadDie
                    Text_BadCount.Text = StarProbe.CountBadDie
                    SSPanel_GoodCount.Caption = StarProbe.CountGoodDie
                    
                    Wafer(forx, yy).ChipSkipDie = False
                    Wafer(forx, yy).ChipInk = False
                    Wafer(forx, yy).ChipInk2 = False
                    Wafer(forx, yy).ChipPlate = False
                    Wafer(forx, yy).ChipMask = False
                    
                    Wafer(forx, yy).flag = False
                    Wafer(forx, yy).FlagBad = False
                    Wafer(forx, yy).BIN = 0
                    
                    Call Display_Chip(pZoom, pOriginal, (forx - StarProbe.StartChip.x), (StarProbe.CurrentChip.y))
                End If
            Next
        
        Case 72 ' B - Normal Die (Right)        '[ 2017.03.27 ] : normal chip left, rightÃß°¡(71,72)
            For forx = xx To (StarProbe.ChipCountX - 1)
                If Wafer(forx, yy).Chip Then
                    If Wafer(forx, yy).ChipSkipDie Or _
                       Wafer(forx, yy).ChipPlate Or _
                       Wafer(forx, yy).ChipMask Then
                        StarProbe.CountSkipDie = StarProbe.CountSkipDie - 1
                        StarProbe.CountGoodDie = StarProbe.CountGoodDie + 1
                    End If
                    
                    If Wafer(forx, yy).FlagBad Then
                        StarProbe.CountBadDie = StarProbe.CountBadDie - 1
                        StarProbe.CountGoodDie = StarProbe.CountGoodDie + 1
                    End If
                    
                    SSPanel_SkipCount.Caption = StarProbe.CountSkipDie
                    Text_SkipCount.Text = StarProbe.CountSkipDie
                    SSPanel_BadCount.Caption = StarProbe.CountBadDie
                    Text_BadCount.Text = StarProbe.CountBadDie
                    SSPanel_GoodCount.Caption = StarProbe.CountGoodDie
                    
                    Wafer(forx, yy).ChipSkipDie = False
                    Wafer(forx, yy).ChipInk = False
                    Wafer(forx, yy).ChipInk2 = False
                    Wafer(forx, yy).ChipPlate = False
                    Wafer(forx, yy).ChipMask = False
                    
                    Wafer(forx, yy).flag = False
                    Wafer(forx, yy).FlagBad = False
                    Wafer(forx, yy).BIN = 0
                    Call Display_Chip(pZoom, pOriginal, (forx - StarProbe.StartChip.x), (StarProbe.CurrentChip.y))
                End If
            Next
            
        Case &H42, &H62 ' V - Normal Die (Left)
            For forx = xx To 0 Step -1
                If Wafer(forx, yy).Chip Then
                    If Wafer(xx, yy).ChipSkipDie Or _
                       Wafer(xx, yy).ChipPlate Or _
                       Wafer(xx, yy).ChipMask Then
                        StarProbe.CountSkipDie = StarProbe.CountSkipDie - 1
                        StarProbe.CountGoodDie = StarProbe.CountGoodDie + 1
                    End If
                    
                    If Wafer(xx, yy).FlagBad Then
                        StarProbe.CountBadDie = StarProbe.CountBadDie - 1
                        StarProbe.CountGoodDie = StarProbe.CountGoodDie + 1
                    End If
                    
                    SSPanel_SkipCount.Caption = StarProbe.CountSkipDie
                    Text_SkipCount.Text = StarProbe.CountSkipDie
                    SSPanel_BadCount.Caption = StarProbe.CountBadDie
                    Text_BadCount.Text = StarProbe.CountBadDie
                    SSPanel_GoodCount.Caption = StarProbe.CountGoodDie
                    
                    Wafer(xx, yy).ChipSkipDie = False
                    Wafer(xx, yy).ChipInk = False
                    Wafer(xx, yy).ChipInk2 = False
                    Wafer(xx, yy).ChipPlate = False
                    Wafer(xx, yy).ChipMask = False
                    
                    Wafer(xx, yy).flag = False
                    Wafer(xx, yy).FlagBad = False
                    Wafer(xx, yy).BIN = 0
                    
                    Call Display_Chip(pZoom, pOriginal, (StarProbe.CurrentChip.x), (StarProbe.CurrentChip.y))
                End If
            Next
        
        Case &H56, &H76 ' B - Normal Die (Right)
            For forx = xx To (StarProbe.ChipCountX - 1)
                If Wafer(forx, yy).Chip Then
                    If Wafer(xx, yy).ChipSkipDie Or _
                       Wafer(xx, yy).ChipPlate Or _
                       Wafer(xx, yy).ChipMask Then
                        StarProbe.CountSkipDie = StarProbe.CountSkipDie - 1
                        StarProbe.CountGoodDie = StarProbe.CountGoodDie + 1
                    End If
                    
                    If Wafer(xx, yy).FlagBad Then
                        StarProbe.CountBadDie = StarProbe.CountBadDie - 1
                        StarProbe.CountGoodDie = StarProbe.CountGoodDie + 1
                    End If
                    
                    SSPanel_SkipCount.Caption = StarProbe.CountSkipDie
                    Text_SkipCount.Text = StarProbe.CountSkipDie
                    SSPanel_BadCount.Caption = StarProbe.CountBadDie
                    Text_BadCount.Text = StarProbe.CountBadDie
                    SSPanel_GoodCount.Caption = StarProbe.CountGoodDie
                    
                    Wafer(xx, yy).ChipSkipDie = False
                    Wafer(xx, yy).ChipInk = False
                    Wafer(xx, yy).ChipInk2 = False
                    Wafer(xx, yy).ChipPlate = False
                    Wafer(xx, yy).ChipMask = False
                    
                    Wafer(xx, yy).flag = False
                    Wafer(xx, yy).FlagBad = False
                    Wafer(xx, yy).BIN = 0
                    
                    Call Display_Chip(pZoom, pOriginal, (StarProbe.CurrentChip.x), (StarProbe.CurrentChip.y))
                End If
            Next
        
        Case 37 ' key left
            xx = xx - 1
            If xx <= 0 Then xx = 0
            b = True
            
        Case 38 ' key up
            yy = yy - 1
            If yy <= 0 Then yy = 0
            b = True
            
        Case 39 ' key right
            xx = xx + 1
            If xx >= StarProbe.ChipCountX Then xx = StarProbe.ChipCountX
            b = True
            
        Case 40 ' key down
            yy = yy + 1
            If yy >= StarProbe.ChipCountY Then yy = StarProbe.ChipCountY
            b = True
    End Select
    
    If b Then
    
On Error GoTo jmp1
        VScroll_Zoom.value = (yy / ((Abs(StarProbe.Max.y) + Abs(StarProbe.Min.y) + 1))) * 1000
        HScroll_Zoom.value = (xx / ((Abs(StarProbe.Max.x) + Abs(StarProbe.Min.x) + 1))) * 1000
jmp1:
        
        StarProbe.CurrentChip.x = xx - StarProbe.StartChip.x
        StarProbe.CurrentChip.y = yy - StarProbe.StartChip.y
        
        Text5 = StarProbe.CurrentChip.x
        Text6 = StarProbe.CurrentChip.y
        
        Call StarProbe_XY_Moving((Text5), (Text6))
        
        If StarProbe_Motor_End_check Then MsgBox "Motor not end check !", 16, "STAR PROBE"
   
        Shape_OChip.Top = (yy * StarProbe.DisplayOChipSizeY) - 1
        Shape_OChip.Left = (xx * StarProbe.DisplayOChipSizeX) - 1
        
        Shape_Chip.Top = (yy * StarProbe.DisplayChipSizeY) - 2
        Shape_Chip.Left = (xx * StarProbe.DisplayChipSizeX) - 2
        
        Label_ChipPosition = StarProbe.CurrentChip.x & "/" & StarProbe.CurrentChip.y
    End If
End Sub

Private Sub pZoom_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Check1(3).value = vbChecked Then Exit Sub
    If StarProbe.DisplayChipSizeX <= 0 Or StarProbe.DisplayChipSizeY <= 0 Then Exit Sub

    Dim xx As Integer, yy As Integer
    Dim BinNo As Integer
    Dim i As Integer
    
    xx = x \ StarProbe.DisplayChipSizeX
    yy = y \ StarProbe.DisplayChipSizeY
    
    If Wafer(xx, yy).Chip Then
        Shape_Move.Top = (yy * StarProbe.DisplayChipSizeY) - 2
        Shape_Move.Left = (xx * StarProbe.DisplayChipSizeX) - 2
        
        Label_Move = (xx - StarProbe.StartChip.x) & "/" & (yy - StarProbe.StartChip.y)
        
        BinNo = Wafer(xx, yy).BIN
        
        Label_BIN = ""
        Label_BIN.BackColor = vbGrayText
        Label_BIN.ForeColor = vbBlack
    
        If Wafer(xx, yy).flag And BinNo >= 0 And BinNo <= 24 Then
            Label_BIN = "BIN " & BinNo
            Label_BIN.BackColor = BINColor(BinNo)
        Else
            If Wafer(xx, yy).ChipPlate Then
                Label_BIN = "Plate Die"
                Label_BIN.BackColor = ChipColor(4)
            ElseIf Wafer(xx, yy).ChipMask Then
                If Wafer(xx, yy).ChipInk = True And Wafer(xx, yy).ChipInk2 = False Then
                    Label_BIN = "Ink1 Die"
                    Label_BIN.BackColor = ChipColor(5)
                ElseIf Wafer(xx, yy).ChipInk = True And Wafer(xx, yy).ChipInk2 = True Then
                    Label_BIN = "Ink2 Die"
                    Label_BIN.BackColor = ChipColor(6)
                Else
                    Label_BIN = "Mask Die"
                    Label_BIN.BackColor = ChipColor(1)
                    Label_BIN.ForeColor = vbWhite
                End If
            ElseIf Wafer(xx, yy).ChipSkipDie Then
                If Wafer(xx, yy).ChipInk = True And Wafer(xx, yy).ChipInk2 = False Then
                    Label_BIN = "Ink1 Die"
                    Label_BIN.BackColor = ChipColor(5)
                ElseIf Wafer(xx, yy).ChipInk = True And Wafer(xx, yy).ChipInk2 = True Then
                    Label_BIN = "Ink2 Die"
                    Label_BIN.BackColor = ChipColor(6)
                Else
                    Label_BIN = "Skip Die"
                    Label_BIN.BackColor = ChipColor(3)
                    Label_BIN.ForeColor = vbWhite
                End If
            Else
                Label_BIN.BackColor = ChipColor(0)
            End If
        End If
    End If
End Sub

Private Sub pZoom_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Check1(3).value = vbChecked Then Exit Sub
    If StarProbe.DisplayChipSizeX <= 0 Or StarProbe.DisplayChipSizeY <= 0 Then Exit Sub

    Dim xx As Integer, yy As Integer
    Dim iResult As Integer
    
    Dim ScrollX As Integer, ScrollY As Integer
    Dim KX As Long
    Dim KY As Long
    
    xx = x \ StarProbe.DisplayChipSizeX
    yy = y \ StarProbe.DisplayChipSizeY
    
    If Not Wafer(xx, yy).Chip Then Exit Sub
    
    Select Case Button
        Case 1
            KY = (yy / (Abs(StarProbe.Max.y) + Abs(StarProbe.Min.y) + 1)) * 1000
            KX = (xx / (Abs(StarProbe.Max.x) + Abs(StarProbe.Min.x) + 1)) * 1000
            
            VScroll_Zoom.value = KY
            HScroll_Zoom.value = KX
            
            StarProbe.CurrentChip.x = xx - StarProbe.StartChip.x
            StarProbe.CurrentChip.y = yy - StarProbe.StartChip.y
            
            Text5 = StarProbe.CurrentChip.x
            Text6 = StarProbe.CurrentChip.y
            
            Shape_OChip.Top = (yy * StarProbe.DisplayOChipSizeY) - 1
            Shape_OChip.Left = (xx * StarProbe.DisplayOChipSizeX) - 1
            
            Shape_Chip.Top = (yy * StarProbe.DisplayChipSizeY) - 2
            Shape_Chip.Left = (xx * StarProbe.DisplayChipSizeX) - 2
            
            Label_ChipPosition = StarProbe.CurrentChip.x & "/" & StarProbe.CurrentChip.y
            
            XAxis = Text5
            YAxis = Text6
            
            Call StarProbe_XY_Moving((XAxis), (YAxis))
            Call ChipPosition((xx), (yy))
            If StarProbe_Motor_End_check Then MsgBox "Motor not end check !", 16, "STAR PROBE"
              
        Case 2
            '[ 2022.08.30 ] : ºÒ·® º¯°æ Ä¨ÀÇ ¹üÀ§ ÁöÁ¤.
            If Needle_check_flag = True Then
                If Needle_STT = False Then
                    iResult = MsgBox("Ä§ÀûºÒ·® ½ÃÀÛ½ÃÁ¡À¸·Î ¼³Á¤ÇÏ½Ã°Ú½À´Ï±î?", vbInformation + vbYesNo, "Information")
                    If iResult = vbYes Then
                        xx = StarProbe.CurrentChip.x + StarProbe.StartChip.x
                        yy = StarProbe.CurrentChip.y + StarProbe.StartChip.y
                        Needle_STT = True
                        txt_sttX.Text = xx
                        txt_sttY.Text = yy
                    Else
                        Needle_STT = False
                    End If
                Else
                    iResult = MsgBox("Ä§ÀûºÒ·® ³¡ÁöÁ¡À¸·Î ¼³Á¤ÇÏ½Ã°Ú½À´Ï±î?", vbInformation + vbYesNo, "Information")
                    If iResult = vbYes Then
                        xx = StarProbe.CurrentChip.x + StarProbe.StartChip.x
                        yy = StarProbe.CurrentChip.y + StarProbe.StartChip.y
                        Needle_STT = True
                        txt_endX.Text = xx
                        txt_endY.Text = yy
                        Needle_STT = False
                        
                        Dim i, j As Integer
                        
                        Dim X1 As Integer
                        Dim Y1 As Integer
                        Dim X2 As Integer
                        Dim Y2 As Integer
                        Dim tmp_val As Integer
                        
                        X1 = val(txt_sttX.Text)
                        Y1 = val(txt_sttY.Text)
                        X2 = val(txt_endX.Text)
                        Y2 = val(txt_endY.Text)
                       
                        If X1 > X2 Then
                            tmp_val = X1
                            X1 = X2
                            X2 = X1
                        End If
                        
                        tmp_val = 0
                        If Y1 > Y2 Then
                            tmp_val = Y1
                            Y1 = Y2
                            Y2 = Y1
                        End If
                                                
                        'Á¤»óÄ¨ÀÌ¸é¼­ ÃøÁ¤À» ÇÑ °æ¿ì
                        For j = Y1 To Y2
                            For i = X1 To X2
                                If Wafer(i, j).ChipSkipDie Or _
                                    Wafer(i, j).ChipInk Or _
                                    Wafer(i, j).ChipPlate Or _
                                    Wafer(i, j).ChipMask Then
                                    StarProbe.CountSkipDie = StarProbe.CountSkipDie - 1
                                    StarProbe.CountGoodDie = StarProbe.CountGoodDie + 1
                                End If
                                
                                If Not Wafer(i, j).flag Then                                'ÃøÁ¤ÇÏÁö ¾ÊÀº °æ¿ì
                                    Wafer(i, j).ChipSkipDie = False
                                    Wafer(i, j).ChipInk = False
                                    Wafer(i, j).ChipPlate = False
                                    Wafer(i, j).ChipMask = False
                                
                                    Wafer(i, j).flag = True
                                    Wafer(i, j).FlagBad = True
                                    Wafer(i, j).BIN = 17
                                    Text_Bin_Count_No(17) = Text_Bin_Count_No(17) + 1
                                    Test_Cnt = Test_Cnt + 1
                                                                        
                                    StarProbe.CountBadDie = StarProbe.CountBadDie + 1
                                
                                    SSPanel_SkipCount.Caption = StarProbe.CountSkipDie
                                    Text_SkipCount.Text = StarProbe.CountSkipDie
                                    SSPanel_BadCount.Caption = StarProbe.CountBadDie
                                    Text_BadCount.Text = StarProbe.CountBadDie
                                    SSPanel_GoodCount.Caption = StarProbe.CountGoodDie
                                    
                                    StarProbe.CurrentChip.x = i - StarProbe.StartChip.x
                                    StarProbe.CurrentChip.y = j - StarProbe.StartChip.y
                                
                                    Call Display_Chip(pZoom, pOriginal, (StarProbe.CurrentChip.x), (StarProbe.CurrentChip.y))
                                Else                                                        'ÃøÁ¤ÇÑ °æ¿ì
                                    If Not Wafer(i, j).FlagBad Then                         '¾çÇ°ÀÎ °æ¿ì
                                        Wafer(i, j).ChipSkipDie = False
                                        Wafer(i, j).ChipInk = False
                                        Wafer(i, j).ChipPlate = False
                                        Wafer(i, j).ChipMask = False
                                    
                                        Wafer(i, j).flag = True
                                        Wafer(i, j).FlagBad = True
                                        Text_Bin_Count_No(Wafer(i, j).BIN) = Text_Bin_Count_No(Wafer(i, j).BIN) - 1
                                        Wafer(i, j).BIN = 17
                                        Text_Bin_Count_No(17) = Text_Bin_Count_No(17) + 1
                                        Good_Cnt = Good_Cnt - 1
                                        StarProbe.CountGoodDie = StarProbe.CountGoodDie - 1
                                        StarProbe.CountBadDie = StarProbe.CountBadDie + 1
                                    
                                        SSPanel_SkipCount.Caption = StarProbe.CountSkipDie
                                        Text_SkipCount.Text = StarProbe.CountSkipDie
                                        SSPanel_BadCount.Caption = StarProbe.CountBadDie
                                        Text_BadCount.Text = StarProbe.CountBadDie
                                        SSPanel_GoodCount.Caption = StarProbe.CountGoodDie
                                        
                                        StarProbe.CurrentChip.x = i - StarProbe.StartChip.x
                                        StarProbe.CurrentChip.y = j - StarProbe.StartChip.y
                                    
                                        Call Display_Chip(pZoom, pOriginal, (StarProbe.CurrentChip.x), (StarProbe.CurrentChip.y))
                                    End If
                                End If
                            Next i
                        Next j
                        Call Bin_Teach
                        Needle_check_flag = False
                        SSPanel2(1).Caption = Test_Cnt
                        Text_TotalCount.Text = Test_Cnt
                        
                        If Test_Cnt = 0 Then
                            SSPanel2(2).Caption = Good_Cnt & Space(1) & "(" & "0.00" & "%)"
                            Text_GoodCount.Text = Good_Cnt
                        Else
                            SSPanel2(2).Caption = Good_Cnt & Space(1) & "(" & Format(Good_Cnt / Test_Cnt * 100, "0.00") & "%)"
                            Text_GoodCount.Text = Good_Cnt
                        End If
                    End If
                End If
            Else
                If Ink_Start_Flag = 0 Then          'Ink Start PositionÀ» »ç¿ëÇÏÁö ¾Ê´Â °æ¿ì
                    iResult = MsgBox("First Chip Ok ... ?", vbInformation + vbYesNo, "Information")     'first die
                    If iResult = vbYes Then
                        Call StarProbe_First_Chip
                                                
                        StarProbe.StartChip.x = xx '- StarProbe.StartChip.X     '186
                        StarProbe.StartChip.y = yy '- StarProbe.StartChip.Y     '15
    
                        '[ 2021.06.25 ] : first die±â¾ï
                        First_X = StarProbe.StartChip.x
                        First_Y = StarProbe.StartChip.y
                        
                        Shape_FirstChip.Top = (yy * StarProbe.DisplayChipSizeY) - 2
                        First_Zoom_TOP = (yy * StarProbe.DisplayChipSizeY) - 2
                        Shape_FirstChip.Left = (xx * StarProbe.DisplayChipSizeX) - 2
                        First_Zoom_LEFT = (xx * StarProbe.DisplayChipSizeX) - 2
                
                        Shape_OFirstChip.Top = (yy * StarProbe.DisplayOChipSizeY) - 1
                        First_Original_TOP = (yy * StarProbe.DisplayOChipSizeY) - 1
                        Shape_OFirstChip.Left = (xx * StarProbe.DisplayOChipSizeX) - 1
                        First_Original_LEFT = (xx * StarProbe.DisplayOChipSizeX) - 1
                    End If
                Else                                'Ink Start PositionÀ» »ç¿ëÇÏ´Â °æ¿ì : 2016.09.27
                    iResult = MsgBox("Ink Start Position Chip Ok ... ?", vbInformation + vbYesNo, "Information")     'Ink Start Chip
                    If iResult = vbYes Then
                        Label_Ink = (xx - StarProbe.StartChip.x) & "/" & (yy - StarProbe.StartChip.y)
                                                                
                        StarProbe.InkStart.x = xx
                        StarProbe.InkStart.y = yy
                                                            
                        Shape_Ink.Top = (yy * StarProbe.DisplayChipSizeY) - 2     '88
                        Shape_Ink.Left = (xx * StarProbe.DisplayChipSizeX) - 2       '1114
                    
                        Shape_OInk.Top = (yy * StarProbe.DisplayOChipSizeY) - 1   '14
                        Shape_OInk.Left = (xx * StarProbe.DisplayOChipSizeX) - 1  '185
                    
                        For xx = 0 To StarProbe.ChipCountX          '[ 2017.03.27 ] : ink±â·Ï ÃÊ±âÈ­
                            For yy = 0 To StarProbe.ChipCountY
                                Wafer(xx, yy).InkDot = False
                            Next
                        Next
                    End If
                End If
            End If
    End Select
End Sub

Private Sub Save_Click()
    PROGRAM_EDIT.Save_Click
End Sub

Private Sub SaveAs_Click()
    PROGRAM_EDIT.SaveAs_Click
End Sub

Private Sub SSPanel1_DblClick(Index As Integer)
    If Index > 0 And Index < 6 Then Text1(Index - 1) = ""
End Sub

Private Sub SSPanel25_Click()
    'Form_Login.Show
    bad_click = 1
End Sub

Private Sub SSPanel5_Click(Index As Integer)
    map_command.Show
End Sub

Private Sub SSTab3_Click(PreviousTab As Integer)
    Select Case SSTab3.Caption
        Case "Map"
            'pZoom.SetFocus
        
        Case "Original"
            'pOriginal.SetFocus
    End Select
End Sub

Private Sub SSTab4_Click(PreviousTab As Integer)
    Select Case SSTab4.Caption
        Case "WORK FORM"
           
        Case "BIN FORM"
            Call Bin_Teach
    End Select
End Sub

Private Sub Text1_Change(Index As Integer)
    Dim contest As String

    contest = Text1(Index)
  
    If contest <> "" Then
        For i = 1 To Len(contest)
            If Trim(Mid(contest, i, 1)) = "\" Or Trim(Mid(contest, i, 1)) = "/" Or Trim(Mid(contest, i, 1)) = ":" Or _
                Trim(Mid(contest, i, 1)) = "*" Or Trim(Mid(contest, i, 1)) = "-" Or Trim(Mid(contest, i, 1)) = "?" Or Trim(Mid(contest, i, 1)) = "<" Or Trim(Mid(contest, i, 1)) = ">" Or Trim(Mid(contest, i, 1)) = "|" Or Trim(Mid(contest, i, 1)) = Chr(34) Then
                MsgBox "File Name writed the wrong expression  " & vbCrLf & " W / : * ? < > | -  " & Chr(34), 16, "FILE NAME MISS"
                Text1(Index).Text = ""
                Exit For
            End If
        Next
    End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 And Index >= 0 And Index < 4 Then Text1(Index + 1).SetFocus
End Sub

Sub Bin_Teach()
    Dim i As Integer
    
    For i = 0 To 24
        Command_BINColor(i).BackColor = BINColor(i)
        Text_BinCount(i).Text = Text_Bin_Count_No(i)
        If BIN_Command(i) = Empty Then
            Text_BinCommand(i).Text = Empty
        Else
            Text_BinCommand(i).Text = BIN_Command(i)
        End If
    Next
    
    For i = 0 To 5
        Command_ChipColor(i).BackColor = ChipColor(i)
    Next
End Sub

Sub AllControlEnable()
    For i = 0 To 4
        Command2(i).Enabled = True
    Next i
    
    For i = 0 To 3                        '2005.06.21 star probe 3->4
        If DemoMode = 0 Then
            Check1(i).Enabled = True
        Else
            Check1(i).Enabled = False
        End If
    Next i
    
    If DemoMode = 0 Then
        Command1(0).Enabled = True
        Command1(1).Enabled = True
    Else
        Command1(0).Enabled = False
        Command1(1).Enabled = False
    End If
End Sub

Private Sub Text11_Change()
'    If val(Text11.Text) = 0 Then
'        MsgBox "Min limit value -> 1"
'        Text11.Text = "1"
'    End If
    '[ 2022.05.06 ] : fail step ¼öÁ¤
    If IsNumeric(Text11.Text) = False Then
        MsgBox "Invalid .This blank Inputed to Number !", vbExclamation, "Error"
        Text11.Text = "0"
        LOOP_COUNT = Text11.Text
        Exit Sub
    End If
    LOOP_COUNT = Text11.Text
End Sub

Private Sub Timer1_Timer()
    Dim val, sval As String

    If DemoMode = 0 Then val = MSComm1.Input

    If ACCO = 1 Or Model_Select = 1 Then                            'ACCO only
        sval = Replace(val, vbCrLf, "")
    Else
        sval = Replace(val, vbLf, "")
    End If
    sval = Trim(sval)

    If sval <> Empty And Mid(sval, 1, 2) <> "ID" Then
        If DemoMode = 0 Then
            If ACCO = 1 Or Model_Select = 1 Then                            'ACCO only
                
            Else
                MSComm1.Output = ">" & vbLf
            End If
        End If
    ElseIf Mid(sval, 1, 2) = "ID" Then
        If DemoMode = 0 Then
            If ACCO = 1 Or Model_Select = 1 Then                            'ACCO only
                
                MSComm1.Output = "2001X.CE.249799-001" & vbCrLf
            Else
                MSComm1.Output = ">" & vbLf
                MSComm1.Output = "2001X.CE.249799-001" & vbLf
            End If
        End If
        Command4.BackColor = &HFFFF80
        Timer1.Enabled = False
    End If
End Sub

Private Sub VScroll_Zoom_Change()
    pZoom.Top = -(((pZoom.Height - 12495) / 1000) * VScroll_Zoom.value)
End Sub

Private Sub VScroll_Zoom_Scroll()
    Call VScroll_Zoom_Change
End Sub

Public Function StarProbe_Measure_CH2(xx As Integer, yy As Integer, bRight As Boolean, bBad As Boolean) As Boolean
    Dim bChip As Boolean
    Dim NG As Boolean
    
    Dim x As Integer, y As Integer, i As Long

    Dim forx As Integer, fory As Integer
    Dim xfrom As Integer, xto As Integer
    Dim yfrom As Integer, yto As Integer
    
    Dim linebadcount As Integer
    
    Dim bMeasureStart As Boolean
    Dim END_VALUE As Integer                                'search Á¾·á °ª ÀúÀå
    
    If Stop_Measure Then
        MeasureChipCount = Stop_MeasureChipCount
        MeasureChipCountOk = Stop_MeasureChipCountOk
        bRight = Stop_Right
        
        Erase MeasureSeq
        For i = 0 To 200000
            MeasureSeq(i) = Stop_MeasureSeq(i)
        Next
        
        GoTo MeasureStart
    Else
        x = xx + StarProbe.StartChip.x
        y = yy + StarProbe.StartChip.y
        
        StarProbe.MeasureStartX = x
        StarProbe.MeasureStartY = y
        
        NG = False
        
        Erase MeasureSeq
        MeasureChipCount = 0
        MeasureChipCountOk = 0
        
        XAxis = xx
        YAxis = yy
    End If
    
    Call StarProbe_XY_Moving((XAxis), (YAxis))
    Call ChipPosition((XAxis + StarProbe.StartChip.x), (YAxis + StarProbe.StartChip.y))
    
    If ErrorStop = True Then Exit Function           ' Ãß°¡
   
    SSPanel2(9).Caption = result & Space(1) & "ms"
    SSPanel2(1).Caption = Test_Cnt
    Text_TotalCount.Text = Test_Cnt
    
    If Test_Cnt = 0 Then
        SSPanel2(2).Caption = Good_Cnt & Space(1) & "(" & "0.00" & "%)"
    Else
        SSPanel2(2).Caption = Good_Cnt & Space(1) & "(" & Format(Good_Cnt / Test_Cnt * 100, "0.00") & "%)"
    End If
    Text_GoodCount.Text = Good_Cnt
    
    SSPanel_BadCount.Caption = StarProbe.CountBadDie
    Text_BadCount.Text = StarProbe.CountBadDie
    SSPanel_GoodCount.Caption = StarProbe.CountGoodDie
         
    If Not StarProbe_Motor_End_check Then
       If StarProbe.MeasureAll = 1 Or (XPitch(TT_NO) = 1 And YPitch(TT_NO) = 1) Then
            If StarProbe.Ink_After = 0 Then
                If InkRun_Left((XAxis), (YAxis)) Then
                    Call StarProbe_Left_Ink_Dot(StarProbe.Ink_LeftPort)
                    Call InkRun_LeftOk((XAxis), (YAxis))
                End If
                If InkRun_Right((XAxis), (YAxis)) Then
                    Call StarProbe_Right_Ink_Dot(StarProbe.Ink_RightPort)
                    Call InkRun_RightOk((XAxis), (YAxis))
                End If
            End If
        End If
        
        If Not bBad And (Wafer(x, y).flag = False Or Wafer(x, y + 1).flag = False) Then      '[ 2021.10.27 ] : bin clearÈÄ ¿Àµ¿ÀÛÇÏ´Â ºÎºÐ °ü·Ã ¼öÁ¤.
            If PROD.PGM_CHECK = False Then
                If StarProbe.Tip_Clean = 1 Then
                    StarProbe.Tipclean_Count = StarProbe.Tipclean_Count + 1
                    Text7.Refresh
                    Text7.Text = StarProbe.Tipclean_Count
                End If
                TEST
                '========================================================================================
                If (Test_Cnt Mod val(Text10.Text)) < 2 And Test_Cnt <> 0 Then
                    Call StarProbe_FileSave_Data("c:\Star Probe\Temp.SP")           '1000¹ø¸¶´Ù ÀúÀå
                End If
                '========================================================================================
            End If
            
            If StarProbe.MeasureAll = 1 Or (XPitch(TT_NO) = 1 And YPitch(TT_NO) = 1) Then
                If result < 30 Then Sleep 20
            End If
        End If
        
        StarProbe_WorkDateTime_To = CDate(Date$ & " " & Time$)
          
        Call StarProbe_WorkDateTime_HMS(StarProbe_WorkDateTime_Total + DateDiff("S", StarProbe_WorkDateTime_From, StarProbe_WorkDateTime_To))
                  
        SSPanel_DateTime.Caption = StarProbe_WorkDateTime.D & " Day " & _
                                   StarProbe_WorkDateTime.h & ":" & _
                                   StarProbe_WorkDateTime.M & ":" & _
                                   StarProbe_WorkDateTime.s

        SSPanel_Yield.Caption = Format(((StarProbe.CountGoodDie / (StarProbe.CountGoodDie + StarProbe.CountBadDie)) * 100), "00.00") & "%"
        
        For i = 0 To 1
            If Text_Chip(i).Text = "Chip" Then
                Wafer(x, ((y + 1) - i)).flag = True
                If Text_ChipBIN(i).Text = "" Then
                    Wafer(x, ((y + 1) - i)).BIN = 0
                    Text_Bin_Count_No(Wafer(x, ((y + 1) - i)).BIN) = Text_Bin_Count_No(Wafer(x, ((y + 1) - i)).BIN) + 1
                    Text_BinCount(Wafer(x, ((y + 1) - i)).BIN).Text = Text_Bin_Count_No(Wafer(x, ((y + 1) - i)).BIN)
                Else
                    Wafer(x, ((y + 1) - i)).BIN = Int(val(Text_ChipBIN(i).Text))
                    Text_Bin_Count_No(Wafer(x, ((y + 1) - i)).BIN) = Text_Bin_Count_No(Wafer(x, ((y + 1) - i)).BIN) + 1
                    Text_BinCount(Wafer(x, ((y + 1) - i)).BIN).Text = Text_Bin_Count_No(Wafer(x, ((y + 1) - i)).BIN)
                End If
                Bin_Count(Wafer(x, ((y + 1) - i)).BIN) = Bin_Count(Wafer(x, ((y + 1) - i)).BIN) + 1     '2016.03.11
                Wafer(x, ((y + 1) - i)).ChipMeasure = True
                Wafer(x, ((y + 1) - i)).MeasureWait = False
                Test_Cnt = Test_Cnt + 1
                If Wafer(x, ((y + 1) - i)).BIN = GOOD_BIN_NO Then                 'pass
                    Good_Cnt = Good_Cnt + 1
                    Wafer(x, ((y + 1) - i)).FlagBad = False
                    If i = 0 Then
                        Test_Fail_Count2 = 1
                    ElseIf i = 1 Then
                        Test_Fail_Count1 = 1
                    End If
                    
                    SSPanel2(3).Caption = Wafer(x, ((y + 1) - i)).BIN
                    SSPanel2(4).Caption = "PASS"
                    SSPanel2(4).BackColor = &HFF00&
                Else                                                    'fail
                    Fail_Find = True
                    Wafer(x, ((y + 1) - i)).FlagBad = True
                    StarProbe.CountBadDie = StarProbe.CountBadDie + 1
                    StarProbe.CountGoodDie = StarProbe.CountGoodDie - 1
                    If i = 0 Then
                        Test_Fail_Count2 = Test_Fail_Count2 + 1
                    ElseIf i = 1 Then
                        Test_Fail_Count1 = Test_Fail_Count1 + 1
                    End If
                    SSPanel2(3).Caption = Wafer(x, ((y + 1) - i)).BIN
                    SSPanel2(4).Caption = "FAIL"
                    SSPanel2(4).BackColor = &HFF&
                End If
            End If
        Next i
        
        Text_ReciveData.Text = Empty
        Text_ChipBIN(0).Text = Empty
        Text_ChipBIN(1).Text = Empty
        Text_ChipBIN(2).Text = Empty
        Text_ChipBIN(3).Text = Empty
 
        SSPanel_BadCount.Caption = StarProbe.CountBadDie
        Text_BadCount.Text = StarProbe.CountBadDie
        SSPanel_GoodCount.Caption = StarProbe.CountGoodDie
        
        VScroll_Zoom.value = (y / ((Abs(StarProbe.Max.y) + Abs(StarProbe.Min.y) + 1))) * 1000
        HScroll_Zoom.value = (x / ((Abs(StarProbe.Max.x) + Abs(StarProbe.Min.x) + 1))) * 1000
        
        StarProbe.CurrentChip.x = x - StarProbe.StartChip.x
        StarProbe.CurrentChip.y = y - StarProbe.StartChip.y
        
        Text5 = StarProbe.CurrentChip.x
        Text6 = StarProbe.CurrentChip.y
        
        Shape_Chip.Top = (y * StarProbe.DisplayChipSizeY) - 2
        Shape_Chip.Left = (x * StarProbe.DisplayChipSizeX) - 2
                
        Label_ChipPosition = StarProbe.CurrentChip.x & "/" & StarProbe.CurrentChip.y
                            
        Call Display_Chip(pZoom, pOriginal, xx, yy)
        Call Display_Chip(pZoom, pOriginal, xx, (yy + 1))
'        Call Display_Chip(pZoom, pOriginal, XX, (yy + 2))
'        Call Display_Chip(pZoom, pOriginal, XX, (yy + 3))
        
        '[ 2022.05.19 ] : ¿¬¼ÓºÒ·® ÈÄ clean tipÇÑ´ÙÀ½ µÎ¹øÂ° ÃøÁ¤
        If tip_clean_count_flag_1 = 2 And Test_Fail_Count1 > 1 Then
            tip_clean_count_flag_1 = 0
            MsgBox "¿¬¼ÓºÒ·®ÀÌ ¹ß»ýÇÏ¿´½À´Ï´Ù Tip Å¬¸®´×À» ½Ç½ÃÇØ ÁÖ¼¼¿ä !", 16, "STAR PROBE"
            Test_Fail_Count1 = 1
            bStop = True
            Check1(3).value = 0
            tip_clean_count_flag_1 = 0
            tip_clean_count_flag_2 = 0
            Test_Fail_Count1 = 0
            Test_Fail_Count2 = 0
            Exit Function
        ElseIf tip_clean_count_flag_2 = 2 And Test_Fail_Count2 > 1 Then
            tip_clean_count_flag_2 = 0
            MsgBox "¿¬¼ÓºÒ·®ÀÌ ¹ß»ýÇÏ¿´½À´Ï´Ù Tip Å¬¸®´×À» ½Ç½ÃÇØ ÁÖ¼¼¿ä !", 16, "STAR PROBE"
            Test_Fail_Count2 = 1
            bStop = True
            Check1(3).value = 0
            tip_clean_count_flag_1 = 0
            tip_clean_count_flag_2 = 0
            Test_Fail_Count1 = 0
            Test_Fail_Count2 = 0
            Exit Function
        End If
        
        '[ 2022.05.19 ] : ¿¬¼ÓºÒ·® ÈÄ clean tipÀ» ÇÑ´ÙÀ½ ´Ù½Ã ºÒ·®ÀÎ °æ¿ì
        If tip_clean_count_flag_1 = 1 And Test_Fail_Count1 > 1 Then
            tip_clean_count_flag_1 = 2
            Call StarProbe_tip_clean
            Sleep 500
        ElseIf tip_clean_count_flag_2 = 1 And Test_Fail_Count2 > 1 Then
            tip_clean_count_flag_2 = 2
            Call StarProbe_tip_clean
            Sleep 500
        Else
            tip_clean_count_flag_1 = 0
            tip_clean_count_flag_2 = 0
        End If
        
        
        If StarProbe.Probe_Stop = 1 And Sample_No_Ink = False Then
            If Test_Fail_Count1 > StarProbe.Probe_Stop_Tfail_Count Then
                If tip_clean_count_flag_1 = 0 Then
                    tip_clean_count_flag_1 = 1
                    Call StarProbe_tip_clean
                    Sleep 500
                Else
'                    MsgBox " Test4 Continus Fail  check !", 16, "STAR PROBE"
'                    Test_Fail_Count1 = 1
'                    bStop = True
'                    Check1(3).value = 0
                End If
            ElseIf Test_Fail_Count2 > StarProbe.Probe_Stop_Tfail_Count Then
                If tip_clean_count_flag_2 = 0 Then
                    tip_clean_count_flag_2 = 1
                    Call StarProbe_tip_clean
                    Sleep 500
                Else
'                    MsgBox " Test3 Continus Fail  check !", 16, "STAR PROBE"
'                    Test_Fail_Count2 = 1
'                    bStop = True
'                    Check1(3).value = 0
                End If
            End If
        End If
        
        ' ÇöÀç Ä¨ ÃøÁ¤ °á°ú°¡ ¾çÇ°ÀÌ³ª ºÒ·®ÀÌ ³ª¿Íµµ
        ' ÁÖÀ§¿¡ Ä¨À» ¸ðµÎ ÃøÁ¤ÇÏ¿´´ÂÁö ¾ÈÇÏ¿´´ÂÁö´Â ÀÐ¾î¼­ ´ÙÀ½Ã³¸® ÇÑ´Ù.
       
        If bBad Then
            NG = True
        Else
            If StarProbe.ReMeasure = vbChecked Then
                If Not bChip Then
                    NG = Not StarProbe_MeasureLineOk_CH2(x, y, StarProbe.LineOk)
                Else
                    NG = True
                End If
            Else
                NG = bChip
            End If
        End If
        If StarProbe.MeasureAll = 1 Or Sample_No_Ink = True Or (XPitch(TT_NO) = 1 And YPitch(TT_NO) = 1) Then NG = False
        
        bBad = False
        
        ' 2005.09.01
        MeasureChipCount = 0
        MeasureChipCountOk = 0
       
        ' Å×½ºÆ® ÃøÁ¤ °á°ú°¡ NGÀÌ¸é ÃøÁ¤ ÃßÀû ¾Ë°í¸®ÁòÀ» Àû¿ëÇÏ¶ó.
        M_CNT = 0
        YOON_CNT = 0
        BACK_X = x
        BACK_Y = y
        
        If Sample_No_Ink = False Then
            If (StarProbe.LimitArea = 1 And Fail_Find = True) Then
REMEA:
                MeasureSeq(0).x = x
                MeasureSeq(0).y = y
                MeasureSeq(0).r = True
                    
                If bRight Then
                    Call StarProbe_MeasureLineRight_CH2(x, y, StarProbe.RCount)
                Else
                    Call StarProbe_MeasureLineLeft_CH2(x, y, StarProbe.RCount)
                End If

MeasureStart:
                bMeasureStart = True
                
                'Ãß°¡
                XVAL = XPitch(TT_NO) * 2
                XVAL1 = XVAL - 2
                XVAL2 = XVAL - 6
                          
                Do While bMeasureStart
                    DoEvents
                    '========================================================================================================= [ 2 <= Xpitch <= 30 ]
                    If ((XPitch(TT_NO) > 2) And (XPitch(TT_NO) <= 30)) Then
                        If (((MeasureChipCountOk Mod 2 = 0) And (MeasureChipCountOk > 0 And (MeasureChipCountOk <= XVAL2)) And Fail_Find = True) Or (MeasureChipCountOk = XVAL1)) Then
                            y = y + 1
                            Fail_Find = False
                            X_STT = 1
                            Move_CNT = 1
                            M_CNT = 1
                        ElseIf (MeasureChipCountOk Mod 2 = 0) And (MeasureChipCountOk > 0) And (MeasureChipCountOk < XVAL) And Fail_Find = False Then
                            y = y + 1
                            Fail_Find = False
                            X_STT = 1
                            Move_CNT = 1
                            M_CNT = 1
                            If YOON_CNT > 0 Then
                                If MeasureChipCountOk <= 8 Then
                                    MeasureChipCountOk = MeasureChipCountOk
                                Else
                                    MeasureChipCountOk = XVAL1 + 4
                                End If
                            Else
                                MeasureChipCountOk = MeasureChipCountOk '+ (XVAL1 - MeasureChipCountOk)
                            End If
                            MeasureChipCount = 0
                        ElseIf (MeasureChipCountOk = XVAL) And Fail_Find = True Then
                            Fail_Find = False
                            X_STT = 1
                            Move_CNT = 1
                            
                            If YOON_CNT > 0 Then
                                MeasureChipCountOk = MeasureChipCountOk
                                MeasureChipCount = 0
                            Else
                                MeasureChipCountOk = 0
                                MeasureChipCount = 0
                                GoTo REMEA
                            End If
                        End If
                        If YOON_CNT > 0 Then
                            If (MeasureChipCountOk = XVAL + 2) And Fail_Find = True Then
                                Fail_Find = False
                                X_STT = 1
                                Move_CNT = 1
                                                            
                                MeasureChipCountOk = 0
                                MeasureChipCount = 0
                                GoTo REMEA
                            End If
                        End If
                    ElseIf XPitch(TT_NO) = 2 Then
                        If MeasureChipCountOk = 4 Then
                            Fail_Find = False
                            X_STT = 1
                            Move_CNT = 1
                            M_CNT = 1
                        ElseIf (MeasureChipCountOk = XVAL) And Fail_Find = True Then
                            Fail_Find = False
                            X_STT = 1
                            Move_CNT = 1
                            
                            If YOON_CNT > 0 Then
                                MeasureChipCountOk = MeasureChipCountOk
                                MeasureChipCount = 0
                            Else
                                MeasureChipCountOk = 0
                                MeasureChipCount = 0
                                GoTo REMEA
                            End If
                        End If
                        If YOON_CNT > 0 Then
                            If (MeasureChipCountOk = XVAL + 4) And Fail_Find = True Then
                                Fail_Find = False
                                X_STT = 1
                                Move_CNT = 1
                                                            
                                MeasureChipCountOk = 0
                                MeasureChipCount = 0
                                GoTo REMEA
                            End If
                        End If
                    End If
                    '=========================================================================================================
                      
                    If bStop = True Then
                        Stop_Measure = True
                        Stop_MeasureChipCount = MeasureChipCount
                        Stop_MeasureChipCountOk = MeasureChipCountOk
                        Stop_Right = bRight
                        
                        Erase Stop_MeasureSeq
                        
                        For i = 0 To 200000
                            Stop_MeasureSeq(i) = MeasureSeq(i)
                        Next
                    
                        Exit Do
                    End If
                      
                    MeasureChipCountOk = MeasureChipCountOk + 1
                    
                    If YOON_CNT > 0 Then
                        END_VALUE = XPitch(TT_NO) * 2 + 2
                    Else
                        END_VALUE = XPitch(TT_NO) * 2
                    End If
                    '//////////////////////////////////////////////////////////////////////////////////[6´Ü]
                    DoEvents
                    
                    If MeasureChipCountOk > END_VALUE Then
                        If StarProbe.LimitArea = 1 Then         '[°øÅë]
                            Fail_Loop = False
                            X_STT = 1
                            Fail_Find = False
                        End If
                        Stop_Measure = False
                        MeasureChipCountOk = 0
                        
                        '[ 2022.05.06 ] :Ãß°¡
                        If YOON_CNT = (LOOP_COUNT * 2) Then 'loopÁ¾·á
                            FAIL_COUNT = 0
                            Exit Do
                        End If
                        'Áß°£
                        If YOON_CNT = 0 Then                    '[Áß°£->»ó´Ü1]
                            YOON_CNT = YOON_CNT + 1
                            y = y - 2
                            x = BACK_X
                            FAIL_COUNT = 0
                            GoTo REMEA
                        End If
                                                                        
                        '»ó´Ü½ÃÀÛ
                        If YOON_CNT <= LOOP_COUNT Then           '1 ~ LOOP_COUNT
                            If YOON_CNT = LOOP_COUNT Then       '[»ó´Ü8->ÇÏ´Ü1]
                                'y = y + (4 + (4 * YOON_CNT))       '2016.09.22
                                y = BACK_Y + 2
                                YOON_CNT = YOON_CNT + 1
                                x = BACK_X
                                FAIL_COUNT = 0
                                GoTo REMEA
                            Else
                                If FAIL_COUNT = 0 Then          '¸ðµÎ¾çÇ°ÀÌ¸é ÇÏ´Ü1·Î ÀÌµ¿
                                    'y = y + (4 + (4 * YOON_CNT))   '2016.09.22
                                    y = BACK_Y + 2
                                    YOON_CNT = LOOP_COUNT + 1   'ÇÏ´Ü½ÃÀÛÀ¸·Î ÀÌµ¿
                                Else
                                    y = y - 2
                                    YOON_CNT = YOON_CNT + 1
                                End If
                                x = BACK_X
                                FAIL_COUNT = 0
                                GoTo REMEA
                            End If
                        'ÇÏ´Ü½ÃÀÛ
                        ElseIf YOON_CNT < (LOOP_COUNT * 2) Then 'LOOP_COUNT ~ LOOP_COUNT*2
                            If FAIL_COUNT = 0 Then Exit Do      '¸ðµÎ¾çÇ°ÀÌ¸é loop out
                            YOON_CNT = YOON_CNT + 1
                            y = y + 2
                            x = BACK_X
                            FAIL_COUNT = 0
                            GoTo REMEA
                        ElseIf YOON_CNT = (LOOP_COUNT * 2) Then 'loopÁ¾·á
                            FAIL_COUNT = 0
                            Exit Do
                        End If
                    End If
                    '//////////////////////////////////////////////////////////////////////////////////////
                      
                    ' Tester Measure Start
                    x = MeasureSeq(MeasureChipCountOk).x
                    y = MeasureSeq(MeasureChipCountOk).y
                                        
                    If StarProbe.LimitArea = 1 Then
                        VScroll_Zoom.value = (y / ((Abs(StarProbe.Max.y) + Abs(StarProbe.Min.y) + 1))) * 1000
                    Else
                        ' wafer scroll move
                        VScroll_Zoom.value = (y / ((Abs(StarProbe.Max.y) + Abs(StarProbe.Min.y) + 1))) * 1000
                        HScroll_Zoom.value = (x / ((Abs(StarProbe.Max.x) + Abs(StarProbe.Min.x) + 1))) * 1000
                    End If
                    
                    StarProbe.CurrentChip.x = x - StarProbe.StartChip.x
                    StarProbe.CurrentChip.y = y - StarProbe.StartChip.y
                    
                    Text5 = StarProbe.CurrentChip.x
                    Text6 = StarProbe.CurrentChip.y
                                    
                    XAxis = StarProbe.CurrentChip.x
                    YAxis = StarProbe.CurrentChip.y
                    
                    move_ok = False
                    If Wafer(x, y).ChipSkipDie = True And Wafer(x, y + 1).ChipSkipDie = True Then
                        move_ok = True
                    Else
                        For yyy = 0 To 1
                            If Wafer(x, y + yyy).flag = True Then
                                move_ok = True
                                Exit For
                            End If
                        Next yyy
                        
                        '[ 2022.05.06 ] : chipÀÌ ¾Æ´Ñ °æ¿ì Ã³¸® Ãß°¡
                        If Wafer(x, y).Chip = False And Wafer(x, y + 1).Chip = False Then
                            move_ok = True
                        End If
                    End If
                    
                    If move_ok = False Then
                        Call StarProbe_XY_Moving((XAxis), (YAxis))
                        
                        If Not StarProbe_Motor_End_check Then
                            Label_ChipPosition = StarProbe.CurrentChip.x & "/" & StarProbe.CurrentChip.y
                            
                            Shape_Chip.Top = (y * (StarProbe.DisplayChipSizeY)) - 2
                            Shape_Chip.Left = (x * StarProbe.DisplayChipSizeX) - 2
                                            
                            SSPanel2(9).Caption = result & Space(1) & "ms"
                            SSPanel2(1).Caption = Test_Cnt
                            Text_TotalCount.Text = Test_Cnt
                            If Test_Cnt = 0 Then
                                SSPanel2(2).Caption = Good_Cnt & Space(1) & "(" & "0.00" & "%)"
                            Else
                                SSPanel2(2).Caption = Good_Cnt & Space(1) & "(" & Format(Good_Cnt / Test_Cnt * 100, "0.00") & "%)"
                            End If
                            Text_GoodCount.Text = Good_Cnt
                             
                            SSPanel_BadCount.Caption = StarProbe.CountBadDie
                            Text_BadCount.Text = StarProbe.CountBadDie
                            SSPanel_GoodCount.Caption = StarProbe.CountGoodDie
                            
                            If ErrorStop = True Then
                                Stop_Measure = True
                            
                                Stop_MeasureChipCount = MeasureChipCount - 1
                                Stop_MeasureChipCountOk = MeasureChipCountOk
                                Stop_Right = bRight
                                
                                Erase Stop_MeasureSeq
                                
                                For i = 0 To 200000
                                    Stop_MeasureSeq(i) = MeasureSeq(i)
                                Next
                                Exit Do   'Ãß°¡
                            End If
                            
                            Call ChipPosition((XAxis + StarProbe.StartChip.x), (YAxis + StarProbe.StartChip.y))
                                
                            If StarProbe.Ink_After = 0 Then
                                If InkRun_Left((XAxis), (YAxis)) Then
                                    Call StarProbe_Left_Ink_Dot(StarProbe.Ink_LeftPort)
                                    Call InkRun_LeftOk((XAxis), (YAxis))
                                End If
                                If InkRun_Right((XAxis), (YAxis)) Then
                                    Call StarProbe_Right_Ink_Dot(StarProbe.Ink_RightPort)
                                    Call InkRun_RightOk((XAxis), (YAxis))
                                End If
                            End If
                                  
                            Dim GOGO As Boolean
                            GOGO = False
                            If Text_ChipTest.Text <> 0 Then
                                For XXX = 0 To 1
                                    If Wafer(x, y + XXX).Chip And Not Wafer(x, y + XXX).ChipMask And Not Wafer(x, y + XXX).ChipInk And Not Wafer(x, y + XXX).ChipSkipDie And Not Wafer(x, y + XXX).ChipPlate Then
                                        GOGO = True
                                    Else
                                        'GOGO = True
                                    End If
                                Next XXX
                                If GOGO = True Then
                                'If Wafer(x, y).Chip And Not Wafer(x, y).ChipMask And Not Wafer(x, y).ChipInk And Not Wafer(x, y).ChipSkipDie And Not Wafer(x, y).ChipPlate Then                                                                                                     'Á¤»óÄ¨ÀÎ °æ¿ì
                                    If Not Wafer(x, y).flag Or Not Wafer(x, y + 1).flag Then
                                        If PROD.PGM_CHECK = False Then
                                            TEST
                                            '========================================================================================
                                            If (Test_Cnt Mod val(Text10.Text)) < 2 And Test_Cnt <> 0 Then
                                                Call StarProbe_FileSave_Data("c:\Star Probe\Temp.SP")           '1000¹ø¸¶´Ù ÀúÀå
                                            End If
                                            '========================================================================================
                                        End If
                                            
                                        If result < 30 Then Sleep 20
                                    
                                        StarProbe_WorkDateTime_To = CDate(Date$ & " " & Time$)
                        
                                        Call StarProbe_WorkDateTime_HMS(StarProbe_WorkDateTime_Total + DateDiff("S", StarProbe_WorkDateTime_From, StarProbe_WorkDateTime_To))
                                                
                                        SSPanel_DateTime.Caption = StarProbe_WorkDateTime.D & " Day " & _
                                                                   StarProbe_WorkDateTime.h & ":" & _
                                                                   StarProbe_WorkDateTime.M & ":" & _
                                                                   StarProbe_WorkDateTime.s
                                  
                                        SSPanel_Yield.Caption = Format(((StarProbe.CountGoodDie / (StarProbe.CountGoodDie + StarProbe.CountBadDie)) * 100), "00.00") & "%"
                                        
                                        For i = 0 To 1
                                            If Text_Chip(i).Text = "Chip" Then
                                                Wafer(x, ((y + 1) - i)).flag = True
                                                If Text_ChipBIN(i).Text = "" Then
                                                    Wafer(x, ((y + 1) - i)).BIN = 0
                                                Else
                                                    Wafer(x, ((y + 1) - i)).BIN = Int(val(Text_ChipBIN(i).Text))
                                                End If
                                                Text_Bin_Count_No(Wafer(x, ((y + 1) - i)).BIN) = Text_Bin_Count_No(Wafer(x, ((y + 1) - i)).BIN) + 1
                                                Text_BinCount(Wafer(x, ((y + 1) - i)).BIN).Text = Text_Bin_Count_No(Wafer(x, ((y + 1) - i)).BIN)
                                                    
                                                Wafer(x, ((y + 1) - i)).ChipMeasure = True
                                                Wafer(x, ((y + 1) - i)).MeasureWait = False
                                                Test_Cnt = Test_Cnt + 1
                                                If Wafer(x, ((y + 1) - i)).BIN = GOOD_BIN_NO Then
                                                    Good_Cnt = Good_Cnt + 1
                                                    Wafer(x, ((y + 1) - i)).FlagBad = False
                                                    If i = 0 Then
                                                        Test_Fail_Count2 = 1
                                                    ElseIf i = 1 Then
                                                        Test_Fail_Count1 = 1
                                                    End If
                                                Else
                                                    Fail_Find = True
                                                    FAIL_COUNT = FAIL_COUNT + 1         'Ãß°¡
                                                    Wafer(x, ((y + 1) - i)).FlagBad = True
                                                    StarProbe.CountBadDie = StarProbe.CountBadDie + 1
                                                    StarProbe.CountGoodDie = StarProbe.CountGoodDie - 1
                                                    If i = 0 Then
                                                        Test_Fail_Count2 = Test_Fail_Count2 + 1
                                                    ElseIf i = 1 Then
                                                        Test_Fail_Count1 = Test_Fail_Count1 + 1
                                                    End If
                                                End If
                                            End If
                                        Next i
                                         
                                        Text_ReciveData.Text = Empty
                                        Text_ChipBIN(0).Text = Empty
                                        Text_ChipBIN(1).Text = Empty
                                                                        
                                        SSPanel_BadCount.Caption = StarProbe.CountBadDie
                                        Text_BadCount.Text = StarProbe.CountBadDie
                                        SSPanel_GoodCount.Caption = StarProbe.CountGoodDie
                                       
                                        VScroll_Zoom.value = (y / ((Abs(StarProbe.Max.y) + Abs(StarProbe.Min.y) + 1))) * 1000
                                        HScroll_Zoom.value = (x / ((Abs(StarProbe.Max.x) + Abs(StarProbe.Min.x) + 1))) * 1000
                                       
                                        StarProbe.CurrentChip.x = x - StarProbe.StartChip.x
                                        StarProbe.CurrentChip.y = y - StarProbe.StartChip.y
                                       
                                        Text5 = StarProbe.CurrentChip.x
                                        Text6 = StarProbe.CurrentChip.y
                                       
                                        Shape_Chip.Top = (y * StarProbe.DisplayChipSizeY) - 2
                                        Shape_Chip.Left = (x * StarProbe.DisplayChipSizeX) - 2
                                                                      
                                        Label_ChipPosition = StarProbe.CurrentChip.x & "/" & StarProbe.CurrentChip.y
                                                                                                                
                                        Call Display_Chip(pZoom, pOriginal, Text5, Text6)
                                        Call Display_Chip(pZoom, pOriginal, Text5, (Text6 + 1))
                                                                                
                                        If StarProbe.Probe_Stop = 1 And Sample_No_Ink = False Then
                                            If Test_Fail_Count2 > StarProbe.Probe_Stop_Tfail_Count Then
                                                MsgBox " Test2 Continus Fail  check !", 16, "STAR PROBE"
                                                Test_Fail_Count2 = 1
                                                bStop = True
                                                Check1(3).value = 0
                                            ElseIf Test_Fail_Count1 > StarProbe.Probe_Stop_Tfail_Count Then
                                                MsgBox " Test1 Continus Fail  check !", 16, "STAR PROBE"
                                                Test_Fail_Count1 = 1
                                                bStop = True
                                                Check1(3).value = 0
                                            End If
                                        End If
                                          
                                        If StarProbe.LimitArea = 1 Then
                                        Else
                                            If NG Then
                                                If bRight Then
                                                    Call StarProbe_MeasureLineRight_CH2(x, y, StarProbe.RCount_Sub)
                                                Else
                                                    Call StarProbe_MeasureLineLeft_CH2(x, y, StarProbe.RCount_Sub)
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                    MeasureChipCountOk = MeasureChipCountOk + 1
                Loop
            End If
        End If
    End If
End Function

Public Function StarProbe_Measure_CH1(xx As Integer, yy As Integer, bRight As Boolean, bBad As Boolean) As Boolean
    Dim bChip As Boolean
    Dim NG As Boolean
    
    Dim x As Integer, y As Integer, i As Long

    Dim forx As Integer, fory As Integer
    Dim xfrom As Integer, xto As Integer
    Dim yfrom As Integer, yto As Integer
    
    Dim linebadcount As Integer
    
    Dim bMeasureStart As Boolean
    Dim END_VALUE As Integer                                'search Á¾·á °ª ÀúÀå
    
    If Stop_Measure Then
        MeasureChipCount = Stop_MeasureChipCount
        MeasureChipCountOk = Stop_MeasureChipCountOk
        bRight = Stop_Right
        
        Erase MeasureSeq
        For i = 0 To 200000
            MeasureSeq(i) = Stop_MeasureSeq(i)
        Next
        
        GoTo MeasureStart
    Else
        x = xx + StarProbe.StartChip.x
        y = yy + StarProbe.StartChip.y
        
        StarProbe.MeasureStartX = x
        StarProbe.MeasureStartY = y
        
        NG = False
        
        Erase MeasureSeq
        MeasureChipCount = 0
        MeasureChipCountOk = 0
        
        XAxis = xx
        YAxis = yy
    End If
    
    Call StarProbe_XY_Moving((XAxis), (YAxis))
    Call ChipPosition((XAxis + StarProbe.StartChip.x), (YAxis + StarProbe.StartChip.y))
    
    If ErrorStop = True Then Exit Function           ' Ãß°¡
   
    SSPanel2(9).Caption = result & Space(1) & "ms"
    SSPanel2(1).Caption = Test_Cnt
    Text_TotalCount.Text = Test_Cnt
    
    If Test_Cnt = 0 Then
        SSPanel2(2).Caption = Good_Cnt & Space(1) & "(" & "0.00" & "%)"
    Else
        SSPanel2(2).Caption = Good_Cnt & Space(1) & "(" & Format(Good_Cnt / Test_Cnt * 100, "0.00") & "%)"
    End If
    Text_GoodCount.Text = Good_Cnt
    
    SSPanel_BadCount.Caption = StarProbe.CountBadDie
    Text_BadCount.Text = StarProbe.CountBadDie
    SSPanel_GoodCount.Caption = StarProbe.CountGoodDie
         
    If Not StarProbe_Motor_End_check Then
       If StarProbe.MeasureAll = 1 Or (XPitch(TT_NO) = 1 And YPitch(TT_NO) = 1) Then
            If StarProbe.Ink_After = 0 Then
                If InkRun_Left((XAxis), (YAxis)) Then
                    Call StarProbe_Left_Ink_Dot(StarProbe.Ink_LeftPort)
                    Call InkRun_LeftOk((XAxis), (YAxis))
                End If
                If InkRun_Right((XAxis), (YAxis)) Then
                    Call StarProbe_Right_Ink_Dot(StarProbe.Ink_RightPort)
                    Call InkRun_RightOk((XAxis), (YAxis))
                End If
            End If
        End If
        
        If Not bBad And (Wafer(x, y).flag = False) Then     '[ 2021.10.27 ] : bin clearÈÄ ¿Àµ¿ÀÛÇÏ´Â ºÎºÐ °ü·Ã ¼öÁ¤.
            If PROD.PGM_CHECK = False Then
                If StarProbe.Tip_Clean = 1 Then
                    StarProbe.Tipclean_Count = StarProbe.Tipclean_Count + 1
                    Text7.Refresh
                    Text7.Text = StarProbe.Tipclean_Count
                End If
                TEST
                '========================================================================================
                If (Test_Cnt Mod val(Text10.Text)) < 1 And Test_Cnt <> 0 Then
                    Call StarProbe_FileSave_Data("c:\Star Probe\Temp.SP")           '1000¹ø¸¶´Ù ÀúÀå
                End If
                '========================================================================================
            End If
            
            If StarProbe.MeasureAll = 1 Or (XPitch(TT_NO) = 1 And YPitch(TT_NO) = 1) Then
                If result < 30 Then Sleep 20
            End If
        End If
        
        StarProbe_WorkDateTime_To = CDate(Date$ & " " & Time$)
        Call StarProbe_WorkDateTime_HMS(StarProbe_WorkDateTime_Total + DateDiff("S", StarProbe_WorkDateTime_From, StarProbe_WorkDateTime_To))
        SSPanel_DateTime.Caption = StarProbe_WorkDateTime.D & " Day " & StarProbe_WorkDateTime.h & ":" & StarProbe_WorkDateTime.M & ":" & StarProbe_WorkDateTime.s
        SSPanel_Yield.Caption = Format(((StarProbe.CountGoodDie / (StarProbe.CountGoodDie + StarProbe.CountBadDie)) * 100), "00.00") & "%"
                
        If Text_Chip(0).Text = "Chip" Then
            Wafer(x, y).flag = True
            If Text_ChipBIN(0).Text = "" Then
                Wafer(x, y).BIN = 0
                Text_Bin_Count_No(Wafer(x, y).BIN) = Text_Bin_Count_No(Wafer(x, y).BIN) + 1
                Text_BinCount(Wafer(x, y).BIN).Text = Text_Bin_Count_No(Wafer(x, y).BIN)
            Else
                Wafer(x, y).BIN = Int(val(Text_ChipBIN(i).Text))
                Text_Bin_Count_No(Wafer(x, y).BIN) = Text_Bin_Count_No(Wafer(x, y).BIN) + 1
                Text_BinCount(Wafer(x, y).BIN).Text = Text_Bin_Count_No(Wafer(x, y).BIN)
            End If
            Bin_Count(Wafer(x, y).BIN) = Bin_Count(Wafer(x, y).BIN) + 1     '2016.03.11
            Wafer(x, y).ChipMeasure = True
            Wafer(x, y).MeasureWait = False
            Test_Cnt = Test_Cnt + 1
            If Wafer(x, y).BIN = GOOD_BIN_NO Then                 'pass
                StarprobeBinFlag = False
                Good_Cnt = Good_Cnt + 1
                Wafer(x, y).FlagBad = False
                Test_Fail_Count1 = 1
                
                SSPanel2(3).Caption = Wafer(x, y).BIN
                SSPanel2(4).Caption = "PASS"
                SSPanel2(4).BackColor = &HFF00&
            Else                                                    'fail
                StarprobeBinFlag = True
                Fail_Find = True
                Wafer(x, y).FlagBad = True
                StarProbe.CountBadDie = StarProbe.CountBadDie + 1
                StarProbe.CountGoodDie = StarProbe.CountGoodDie - 1
                Test_Fail_Count1 = Test_Fail_Count1 + 1
                
                SSPanel2(3).Caption = Wafer(x, y).BIN
                SSPanel2(4).Caption = "FAIL"
                SSPanel2(4).BackColor = &HFF&
            End If
        End If
        bChip = StarprobeBinFlag
        
        Text_ReciveData.Text = Empty
        Text_ChipBIN(0).Text = Empty
 
        SSPanel_BadCount.Caption = StarProbe.CountBadDie
        Text_BadCount.Text = StarProbe.CountBadDie
        SSPanel_GoodCount.Caption = StarProbe.CountGoodDie
        
        VScroll_Zoom.value = (y / ((Abs(StarProbe.Max.y) + Abs(StarProbe.Min.y) + 1))) * 1000
        HScroll_Zoom.value = (x / ((Abs(StarProbe.Max.x) + Abs(StarProbe.Min.x) + 1))) * 1000
        
        StarProbe.CurrentChip.x = x - StarProbe.StartChip.x
        StarProbe.CurrentChip.y = y - StarProbe.StartChip.y
        
        Text5 = StarProbe.CurrentChip.x
        Text6 = StarProbe.CurrentChip.y
        
        Shape_Chip.Top = (y * StarProbe.DisplayChipSizeY) - 2
        Shape_Chip.Left = (x * StarProbe.DisplayChipSizeX) - 2
                
        Label_ChipPosition = StarProbe.CurrentChip.x & "/" & StarProbe.CurrentChip.y
                            
        Call Display_Chip(pZoom, pOriginal, xx, yy)

        '[ 2022.05.19 ] : ¿¬¼ÓºÒ·® ÈÄ clean tipÇÑ´ÙÀ½ µÎ¹øÂ° ÃøÁ¤
        If tip_clean_count_flag_1 = 2 And Test_Fail_Count1 > 1 Then
            tip_clean_count_flag_1 = 0
            MsgBox "¿¬¼ÓºÒ·®ÀÌ ¹ß»ýÇÏ¿´½À´Ï´Ù Tip Å¬¸®´×À» ½Ç½ÃÇØ ÁÖ¼¼¿ä !", 16, "STAR PROBE"
            Test_Fail_Count1 = 1
            bStop = True
            Check1(3).value = 0
            tip_clean_count_flag_1 = 0
            Test_Fail_Count1 = 0
            Exit Function
        End If
        
        '[ 2022.05.19 ] : ¿¬¼ÓºÒ·® ÈÄ clean tipÀ» ÇÑ´ÙÀ½ ´Ù½Ã ºÒ·®ÀÎ °æ¿ì
        If tip_clean_count_flag_1 = 1 And Test_Fail_Count1 > 1 Then
            tip_clean_count_flag_1 = 2
            Call StarProbe_tip_clean
            Sleep 500
        Else
            tip_clean_count_flag_1 = 0
            tip_clean_count_flag_2 = 0
            tip_clean_count_flag_3 = 0
            tip_clean_count_flag_4 = 0
        End If
        
        
        If StarProbe.Probe_Stop = 1 And Sample_No_Ink = False Then
            If Test_Fail_Count1 > StarProbe.Probe_Stop_Tfail_Count Then
                If tip_clean_count_flag_1 = 0 Then
                    tip_clean_count_flag_1 = 1
                    Call StarProbe_tip_clean
                    Sleep 500
                Else
'                    MsgBox " Test4 Continus Fail  check !", 16, "STAR PROBE"
'                    Test_Fail_Count1 = 1
'                    bStop = True
'                    Check1(3).value = 0
                End If
            End If
        End If
        
        ' ÇöÀç Ä¨ ÃøÁ¤ °á°ú°¡ ¾çÇ°ÀÌ³ª ºÒ·®ÀÌ ³ª¿Íµµ
        ' ÁÖÀ§¿¡ Ä¨À» ¸ðµÎ ÃøÁ¤ÇÏ¿´´ÂÁö ¾ÈÇÏ¿´´ÂÁö´Â ÀÐ¾î¼­ ´ÙÀ½Ã³¸® ÇÑ´Ù.
       
        If bBad Then
            NG = True
        Else
            If StarProbe.ReMeasure = vbChecked Then
                If Not bChip Then
                    NG = Not StarProbe_MeasureLineOk_CH1(x, y, StarProbe.LineOk)
                Else
                    NG = True
                End If
            Else
                NG = bChip
            End If
        End If
        If StarProbe.MeasureAll = 1 Or Sample_No_Ink = True Or (XPitch(TT_NO) = 1 And YPitch(TT_NO) = 1) Then NG = False
        
        bBad = False
        
        ' 2005.09.01
        If StarProbe.MeasureAll = 0 And (Not bChip) Then
            MeasureChipCount = 0
            MeasureChipCountOk = 0
        
            Select Case XPitch(TT_NO)
                Case 2: xfrom = x:     xto = x + 1
                Case 3: xfrom = x - 1: xto = x + 1
                Case 4: xfrom = x - 1: xto = x + 2
                Case 5: xfrom = x - 2: xto = x + 2
                Case 6: xfrom = x - 2: xto = x + 3
                Case 7: xfrom = x - 3: xto = x + 3
                Case 8: xfrom = x - 3: xto = x + 4
                Case 9: xfrom = x - 4: xto = x + 4
                Case 10: xfrom = x - 4: xto = x + 5
                Case 11: xfrom = x - 5: xto = x + 5
                Case 12: xfrom = x - 5: xto = x + 6
                Case 13: xfrom = x - 6: xto = x + 6
                Case 14: xfrom = x - 6: xto = x + 7
                Case 15: xfrom = x - 7: xto = x + 7
                Case 16: xfrom = x - 7: xto = x + 8
                Case 17: xfrom = x - 8: xto = x + 8
                Case 18: xfrom = x - 8: xto = x + 9
                Case 19: xfrom = x - 9: xto = x + 9
                Case 20: xfrom = x - 9: xto = x + 10
                Case 21: xfrom = x - 10: xto = x + 10
                Case 22: xfrom = x - 10: xto = x + 11
                Case 23: xfrom = x - 11: xto = x + 11
                Case 24: xfrom = x - 11: xto = x + 12
                Case 25: xfrom = x - 12: xto = x + 12
                Case 26: xfrom = x - 12: xto = x + 13
                Case 27: xfrom = x - 13: xto = x + 13
                Case 28: xfrom = x - 13: xto = x + 14
                Case 29: xfrom = x - 14: xto = x + 14
                Case 30: xfrom = x - 14: xto = x + 15
            End Select
        
            Select Case YPitch(TT_NO)
                Case 2: yfrom = y:     yto = y + 1
                Case 3: yfrom = y - 1: yto = y + 1
                Case 4: yfrom = y - 1: yto = y + 2
                Case 5: yfrom = y - 2: yto = y + 2
                Case 6: yfrom = y - 2: yto = y + 3
                Case 7: yfrom = y - 3: yto = y + 3
                Case 8: yfrom = y - 3: yto = y + 4
                Case 9: yfrom = y - 4: yto = y + 4
                Case 10: yfrom = y - 4: yto = y + 5
                Case 11: yfrom = y - 5: yto = y + 5
                Case 12: yfrom = y - 5: yto = y + 6
                Case 13: yfrom = y - 6: yto = y + 6
                Case 14: yfrom = y - 6: yto = y + 7
                Case 15: yfrom = y - 7: yto = y + 7
                Case 16: yfrom = y - 7: yto = y + 8
                Case 17: yfrom = y - 8: yto = y + 8
                Case 18: yfrom = y - 8: yto = y + 9
                Case 19: yfrom = y - 9: yto = y + 9
                Case 20: yfrom = y - 9: yto = y + 10
                Case 21: yfrom = y - 10: yto = y + 10
                Case 22: yfrom = y - 10: yto = y + 11
                Case 23: yfrom = y - 11: yto = y + 11
                Case 24: yfrom = y - 11: yto = y + 12
                Case 25: yfrom = y - 12: yto = y + 12
                Case 26: yfrom = y - 12: yto = y + 13
                Case 27: yfrom = y - 13: yto = y + 13
                Case 28: yfrom = y - 13: yto = y + 14
                Case 29: yfrom = y - 14: yto = y + 14
                Case 30: yfrom = y - 14: yto = y + 15
            End Select
            
            If bRight Then
                If Wafer(x - XPitch(TT_NO), y).Chip And Not Wafer(x - XPitch(TT_NO), y).ChipMask And Not Wafer(x - XPitch(TT_NO), y).ChipSkipDie And Not Wafer(x - XPitch(TT_NO), y).ChipPlate And Wafer(x - XPitch(TT_NO), y).flag And Not Wafer(x - Pitch, y).FlagBad And Not Wafer(x - Pitch, y).MeasureWait Then
                    If Wafer(x - XPitch(TT_NO), y - 1).Chip And Not Wafer(x - XPitch(TT_NO), y - 1).ChipMask And Not Wafer(x - XPitch(TT_NO), y - 1).ChipSkipDie And Not Wafer(x - XPitch(TT_NO), y - 1).ChipPlate And Wafer(x - XPitch(TT_NO), y - 1).flag And Wafer(x - XPitch(TT_NO), y - 1).FlagBad Then
                        If Wafer(x - XPitch(TT_NO) + 1, y).Chip And Not Wafer(x - XPitch(TT_NO) + 1, y).ChipMask And Not Wafer(x - XPitch(TT_NO) + 1, y).ChipSkipDie And Not Wafer(x - XPitch(TT_NO) + 1, y).ChipPlate And Not Wafer(x - XPitch(TT_NO) + 1, y).flag And Not Wafer(x - XPitch(TT_NO) + 1, y).MeasureWait Then
                            MeasureChipCount = MeasureChipCount + 1
                            MeasureSeq(MeasureChipCount).x = x - XPitch(TT_NO) + 1
                            MeasureSeq(MeasureChipCount).y = y
                            MeasureSeq(MeasureChipCount).r = False
                            NG = True
                            bBad = True
                            Wafer(x - XPitch(TT_NO) + 1, y).MeasureWait = True
                         End If
                    End If
                End If
                
                For forx = (xfrom - 1) To xto
                    If Wafer(forx, y - 1).Chip And Wafer(forx, y - 1).flag And Wafer(forx, y - 1).FlagBad And Not Wafer(forx, y - 1).ChipMask And Not Wafer(forx, y - 1).ChipSkipDie And Not Wafer(forx, y - 1).ChipPlate Then
                        If Wafer(forx - 1, y).Chip And Not Wafer(forx - 1, y).ChipMask And Not Wafer(forx - 1, y).ChipSkipDie And Not Wafer(forx - 1, y).ChipPlate And Not Wafer(forx - 1, y).flag And Not Wafer(forx - 1, y).MeasureWait Then
                            MeasureChipCount = MeasureChipCount + 1
                            MeasureSeq(MeasureChipCount).x = forx - 1
                            MeasureSeq(MeasureChipCount).y = y
                            MeasureSeq(MeasureChipCount).r = False
                            NG = True
                            bBad = True
                            Wafer(forx - 1, y).MeasureWait = True
                        End If
    
                        If Wafer(forx, y).Chip And Not Wafer(forx, y).ChipMask And Not Wafer(forx, y).ChipSkipDie And Not Wafer(forx, y).ChipPlate And Not Wafer(forx, y).flag And Not Wafer(forx, y).MeasureWait Then
                            MeasureChipCount = MeasureChipCount + 1
                            MeasureSeq(MeasureChipCount).x = forx
                            MeasureSeq(MeasureChipCount).y = y
                            MeasureSeq(MeasureChipCount).r = False
                            NG = True
                            bBad = True
                            Wafer(forx, y).MeasureWait = True
                        End If
    
                        If Wafer(forx + 1, y).Chip And Not Wafer(forx + 1, y).ChipMask And Not Wafer(forx + 1, y).ChipSkipDie And Not Wafer(forx + 1, y).ChipPlate And Not Wafer(forx + 1, y).flag And Not Wafer(forx + 1, y).MeasureWait Then
                            MeasureChipCount = MeasureChipCount + 1
                            MeasureSeq(MeasureChipCount).x = forx + 1
                            MeasureSeq(MeasureChipCount).y = y
                            MeasureSeq(MeasureChipCount).r = False
                            NG = True
                            bBad = True
                            Wafer(forx + 1, y).MeasureWait = True
                        End If
                    End If
                Next
                
                If Not Wafer(x + XPitch(TT_NO), y).Chip Or (Wafer(x + XPitch(TT_NO), y).Chip And (Wafer(x + XPitch(TT_NO), y).ChipMask Or Wafer(x + XPitch(TT_NO), y).ChipSkipDie Or Wafer(x + XPitch(TT_NO), y).ChipPlate)) Then
                    If Wafer(xto + 1, y - 1).Chip And Not Wafer(xto + 1, y - 1).ChipMask And Not Wafer(xto + 1, y - 1).ChipSkipDie And Not Wafer(xto + 1, y - 1).ChipPlate And Wafer(xto + 1, y - 1).flag And Wafer(xto + 1, y - 1).FlagBad Then
                        If Wafer(xto + 1, y).Chip And Not Wafer(xto + 1, y).ChipMask And Not Wafer(xto + 1, y).ChipSkipDie And Not Wafer(xto + 1, y).ChipPlate And Not Wafer(xto + 1, y).flag And Not Wafer(xto + 1, y).MeasureWait Then
                            MeasureChipCount = MeasureChipCount + 1
                            MeasureSeq(MeasureChipCount).x = xto + 1
                            MeasureSeq(MeasureChipCount).y = y
                            MeasureSeq(MeasureChipCount).r = False
                            NG = True
                            bBad = True
                            Wafer(xto + 1, y).MeasureWait = True
                        End If
                    End If
                End If
    ' bRight = True
    ''''''''''
            Else
    ''''''''''
    ' bRight = False
                If Wafer(x + XPitch(TT_NO), y).Chip And Not Wafer(x + XPitch(TT_NO), y).ChipMask And Not Wafer(x + XPitch(TT_NO), y).ChipSkipDie And Not Wafer(x + XPitch(TT_NO), y).ChipPlate And Wafer(x + XPitch(TT_NO), y).flag And Not Wafer(x + Pitch, y).FlagBad And Not Wafer(x + Pitch, y).MeasureWait Then
                    If Wafer(x + XPitch(TT_NO), y - 1).Chip And Not Wafer(x + XPitch(TT_NO), y - 1).ChipMask And Not Wafer(x + XPitch(TT_NO), y - 1).ChipSkipDie And Not Wafer(x + XPitch(TT_NO), y - 1).ChipPlate And Wafer(x + XPitch(TT_NO), y - 1).flag And Wafer(x + XPitch(TT_NO), y - 1).FlagBad Then
                        If Wafer(x + XPitch(TT_NO) - 1, y).Chip And Not Wafer(x + XPitch(TT_NO) - 1, y).ChipMask And Not Wafer(x + XPitch(TT_NO) - 1, y).ChipSkipDie And Not Wafer(x + XPitch(TT_NO) - 1, y).ChipPlate And Not Wafer(x + XPitch(TT_NO) - 1, y).flag And Not Wafer(x + XPitch(TT_NO) - 1, y).MeasureWait Then
                            MeasureChipCount = MeasureChipCount + 1
                            MeasureSeq(MeasureChipCount).x = x + XPitch(TT_NO) - 1
                            MeasureSeq(MeasureChipCount).y = y
                            MeasureSeq(MeasureChipCount).r = False
                            NG = True
                            bBad = True
                            Wafer(x + XPitch(TT_NO) - 1, y).MeasureWait = True
                         End If
                    End If
                End If
            
                For forx = (xto + 1) To xfrom Step -1
                    If Wafer(forx, y - 1).Chip And Wafer(forx, y - 1).flag And Wafer(forx, y - 1).FlagBad And Not Wafer(forx, y - 1).ChipMask And Not Wafer(forx, y - 1).ChipSkipDie And Not Wafer(forx, y - 1).ChipPlate Then
                        If Wafer(forx + 1, y).Chip And Not Wafer(forx + 1, y).ChipMask And Not Wafer(forx + 1, y).ChipSkipDie And Not Wafer(forx + 1, y).ChipPlate And Not Wafer(forx + 1, y).flag And Not Wafer(forx + 1, y).MeasureWait Then
                            MeasureChipCount = MeasureChipCount + 1
                            MeasureSeq(MeasureChipCount).x = forx + 1
                            MeasureSeq(MeasureChipCount).y = y
                            MeasureSeq(MeasureChipCount).r = False
                            NG = True
                            bBad = True
                            Wafer(forx + 1, y).MeasureWait = True
                        End If
    
                        If Wafer(forx, y).Chip And Not Wafer(forx, y).ChipMask And Not Wafer(forx, y).ChipSkipDie And Not Wafer(forx, y).ChipPlate And Not Wafer(forx, y).flag And Not Wafer(forx, y).MeasureWait Then
                            MeasureChipCount = MeasureChipCount + 1
                            MeasureSeq(MeasureChipCount).x = forx
                            MeasureSeq(MeasureChipCount).y = y
                            MeasureSeq(MeasureChipCount).r = False
                            NG = True
                            bBad = True
                            Wafer(forx, y).MeasureWait = True
                        End If
    
                        If Wafer(forx - 1, y).Chip And Not Wafer(forx - 1, y).ChipMask And Not Wafer(forx - 1, y).ChipSkipDie And Not Wafer(forx - 1, y).ChipPlate And Not Wafer(forx - 1, y).flag And Not Wafer(forx - 1, y).MeasureWait Then
                            MeasureChipCount = MeasureChipCount + 1
                            MeasureSeq(MeasureChipCount).x = forx - 1
                            MeasureSeq(MeasureChipCount).y = y
                            MeasureSeq(MeasureChipCount).r = False
                            NG = True
                            bBad = True
                            Wafer(forx - 1, y).MeasureWait = True
                        End If
                    End If
                Next
                
                If Not Wafer(x - XPitch(TT_NO), y).Chip Or (Wafer(x - XPitch(TT_NO), y).Chip And (Wafer(x - XPitch(TT_NO), y).ChipMask Or Wafer(x - XPitch(TT_NO), y).ChipSkipDie Or Wafer(x - XPitch(TT_NO), y).ChipPlate)) Then
                    If Wafer(xfrom - 1, y - 1).Chip And Not Wafer(xfrom - 1, y - 1).ChipMask And Not Wafer(xfrom - 1, y - 1).ChipSkipDie And Not Wafer(xfrom - 1, y - 1).ChipPlate And Wafer(xfrom - 1, y - 1).flag And Wafer(xfrom - 1, y - 1).FlagBad Then
                        If Wafer(xfrom - 1, y).Chip And Not Wafer(xfrom - 1, y).ChipMask And Not Wafer(xfrom - 1, y).ChipSkipDie And Not Wafer(xfrom - 1, y).ChipPlate And Not Wafer(xfrom - 1, y).flag And Not Wafer(xfrom - 1, y).MeasureWait Then
                            MeasureChipCount = MeasureChipCount + 1
                            MeasureSeq(MeasureChipCount).x = xfrom - 1
                            MeasureSeq(MeasureChipCount).y = y
                            MeasureSeq(MeasureChipCount).r = False
                            NG = True
                            bBad = True
                            Wafer(xfrom - 1, y).MeasureWait = True
                        End If
                    End If
                End If
            End If
        End If
        
        If Sample_No_Ink = False Then
            If NG Then
                MeasureSeq(0).x = x
                MeasureSeq(0).y = y
                MeasureSeq(0).r = True
                
                If Not bBad Then
                    If bRight Then
                        Call StarProbe_MeasureLineRight_CH1(x, y, StarProbe.RCount)
                    Else
                        Call StarProbe_MeasureLineLeft_CH1(x, y, StarProbe.RCount)
                    End If
                End If
                                    
MeasureStart:
                bMeasureStart = True
                                    
                Do While bMeasureStart
                    DoEvents
                      
                    If bStop = True Then
                        Stop_Measure = True
                        Stop_MeasureChipCount = MeasureChipCount
                        Stop_MeasureChipCountOk = MeasureChipCountOk
                        Stop_Right = bRight
                        
                        Erase Stop_MeasureSeq
                        
                        For i = 0 To 200000
                            Stop_MeasureSeq(i) = MeasureSeq(i)
                        Next
                    
                        Exit Do
                    End If
                                        
                    MeasureChipCountOk = MeasureChipCountOk + 1
                    If MeasureChipCountOk > MeasureChipCount Then
                        Stop_Measure = False
                        Exit Do
                    End If
                      
                    ' Tester Measure Start
                    x = MeasureSeq(MeasureChipCountOk).x
                    y = MeasureSeq(MeasureChipCountOk).y
                                        
                    VScroll_Zoom.value = (y / ((Abs(StarProbe.Max.y) + Abs(StarProbe.Min.y) + 1))) * 1000   'scroll y
                    HScroll_Zoom.value = (x / ((Abs(StarProbe.Max.x) + Abs(StarProbe.Min.x) + 1))) * 1000   'scroll x
                    
                    StarProbe.CurrentChip.x = x - StarProbe.StartChip.x
                    StarProbe.CurrentChip.y = y - StarProbe.StartChip.y
                    
                    Text5 = StarProbe.CurrentChip.x
                    Text6 = StarProbe.CurrentChip.y
                                    
                    XAxis = StarProbe.CurrentChip.x
                    YAxis = StarProbe.CurrentChip.y
                                                                                
                    Call StarProbe_XY_Moving((XAxis), (YAxis))
                    
                    If Not StarProbe_Motor_End_check Then
                        Label_ChipPosition = StarProbe.CurrentChip.x & "/" & StarProbe.CurrentChip.y
                        
                        Shape_Chip.Top = (y * (StarProbe.DisplayChipSizeY)) - 2
                        Shape_Chip.Left = (x * StarProbe.DisplayChipSizeX) - 2
                                        
                        SSPanel2(9).Caption = result & Space(1) & "ms"
                        SSPanel2(1).Caption = Test_Cnt
                        Text_TotalCount.Text = Test_Cnt
                        If Test_Cnt = 0 Then
                            SSPanel2(2).Caption = Good_Cnt & Space(1) & "(" & "0.00" & "%)"
                        Else
                            SSPanel2(2).Caption = Good_Cnt & Space(1) & "(" & Format(Good_Cnt / Test_Cnt * 100, "0.00") & "%)"
                        End If
                        Text_GoodCount.Text = Good_Cnt
                         
                        SSPanel_BadCount.Caption = StarProbe.CountBadDie
                        Text_BadCount.Text = StarProbe.CountBadDie
                        SSPanel_GoodCount.Caption = StarProbe.CountGoodDie
                        
                        If ErrorStop = True Then
                            Stop_Measure = True
                        
                            Stop_MeasureChipCount = MeasureChipCount - 1
                            Stop_MeasureChipCountOk = MeasureChipCountOk
                            Stop_Right = bRight
                            
                            Erase Stop_MeasureSeq
                            
                            For i = 0 To 200000
                                Stop_MeasureSeq(i) = MeasureSeq(i)
                            Next
                            Exit Do   'Ãß°¡
                        End If
                        
'                        Call ChipPosition((XAxis + StarProbe.StartChip.x), (YAxis + StarProbe.StartChip.y))
                            
                        If StarProbe.Ink_After = 0 Then
                            If InkRun_Left((XAxis), (YAxis)) Then
                                Call StarProbe_Left_Ink_Dot(StarProbe.Ink_LeftPort)
                                Call InkRun_LeftOk((XAxis), (YAxis))
                            End If
                            If InkRun_Right((XAxis), (YAxis)) Then
                                Call StarProbe_Right_Ink_Dot(StarProbe.Ink_RightPort)
                                Call InkRun_RightOk((XAxis), (YAxis))
                            End If
                        End If
                              
                        If Not Wafer(x, y).flag Then
                            If PROD.PGM_CHECK = False Then
                                TEST
                                '========================================================================================
                                If (Test_Cnt Mod val(Text10.Text)) < 1 And Test_Cnt <> 0 Then
                                    Call StarProbe_FileSave_Data("c:\Star Probe\Temp.SP")           '1000¹ø¸¶´Ù ÀúÀå
                                End If
                                '========================================================================================
                            End If
                                
                            If result < 30 Then Sleep 20
                        
                            StarProbe_WorkDateTime_To = CDate(Date$ & " " & Time$)
            
                            Call StarProbe_WorkDateTime_HMS(StarProbe_WorkDateTime_Total + DateDiff("S", StarProbe_WorkDateTime_From, StarProbe_WorkDateTime_To))
                                    
                            SSPanel_DateTime.Caption = StarProbe_WorkDateTime.D & " Day " & _
                                                       StarProbe_WorkDateTime.h & ":" & _
                                                       StarProbe_WorkDateTime.M & ":" & _
                                                       StarProbe_WorkDateTime.s
                      
                            SSPanel_Yield.Caption = Format(((StarProbe.CountGoodDie / (StarProbe.CountGoodDie + StarProbe.CountBadDie)) * 100), "00.00") & "%"
                                                        
                            If Text_Chip(0).Text = "Chip" Then
                                Wafer(x, y).flag = True
                                If Text_ChipBIN(0).Text = "" Then
                                    Wafer(x, y).BIN = 0
                                Else
                                    Wafer(x, y).BIN = Int(val(Text_ChipBIN(0).Text))
                                End If
                                Text_Bin_Count_No(Wafer(x, y).BIN) = Text_Bin_Count_No(Wafer(x, y).BIN) + 1
                                Text_BinCount(Wafer(x, y).BIN).Text = Text_Bin_Count_No(Wafer(x, y).BIN)
                                    
                                Wafer(x, y).ChipMeasure = True
                                Wafer(x, y).MeasureWait = False
                                Test_Cnt = Test_Cnt + 1
                                If Wafer(x, y).BIN = GOOD_BIN_NO Then
                                    StarprobeBinFlag = False
                                    Fail_Find = False
                                    Good_Cnt = Good_Cnt + 1
                                    Wafer(x, y).FlagBad = False
                                    Test_Fail_Count1 = 1
                                Else
                                    StarprobeBinFlag = True
                                    Fail_Find = True
                                    FAIL_COUNT = FAIL_COUNT + 1         'Ãß°¡
                                    Wafer(x, y).FlagBad = True
                                    StarProbe.CountBadDie = StarProbe.CountBadDie + 1
                                    StarProbe.CountGoodDie = StarProbe.CountGoodDie - 1
                                    Test_Fail_Count1 = Test_Fail_Count1 + 1
                                End If
                            End If
                            bChip = StarprobeBinFlag
                             
                            Text_ReciveData.Text = Empty
                            Text_ChipBIN(0).Text = Empty
                            
                            SSPanel_BadCount.Caption = StarProbe.CountBadDie
                            Text_BadCount.Text = StarProbe.CountBadDie
                            SSPanel_GoodCount.Caption = StarProbe.CountGoodDie
                           
                            VScroll_Zoom.value = (y / ((Abs(StarProbe.Max.y) + Abs(StarProbe.Min.y) + 1))) * 1000
                            HScroll_Zoom.value = (x / ((Abs(StarProbe.Max.x) + Abs(StarProbe.Min.x) + 1))) * 1000
                           
                            StarProbe.CurrentChip.x = x - StarProbe.StartChip.x
                            StarProbe.CurrentChip.y = y - StarProbe.StartChip.y
                           
                            Text5 = StarProbe.CurrentChip.x
                            Text6 = StarProbe.CurrentChip.y
                           
                            Shape_Chip.Top = (y * StarProbe.DisplayChipSizeY) - 2
                            Shape_Chip.Left = (x * StarProbe.DisplayChipSizeX) - 2
                                                          
                            Label_ChipPosition = StarProbe.CurrentChip.x & "/" & StarProbe.CurrentChip.y
                                                                                                    
                            Call Display_Chip(pZoom, pOriginal, Text5, Text6)

                            NG = bChip
                            
                            MeasureSeq(MeasureChipCountOk).r = True
                            
                            If NG Then
                                If bRight Then
                                    Call StarProbe_MeasureLineRight_CH1(x, y, StarProbe.RCount_Sub)
                                Else
                                    Call StarProbe_MeasureLineLeft_CH1(x, y, StarProbe.RCount_Sub)
                                End If
                            End If
                            
                            If StarProbe.Probe_Stop = 1 And Sample_No_Ink = False Then
                                If Test_Fail_Count1 > StarProbe.Probe_Stop_Tfail_Count Then
                                    MsgBox " Test1 Continus Fail  check !", 16, "STAR PROBE"
                                    Test_Fail_Count1 = 1
                                    bStop = True
                                    Check1(3).value = 0
                                End If
                            End If
                        End If
                    End If
                Loop
            End If
        End If
    End If
End Function

Public Function StarProbe_Measure_CH4(xx As Integer, yy As Integer, bRight As Boolean, bBad As Boolean) As Boolean
    Dim bChip As Boolean
    Dim NG As Boolean
    
    Dim x As Integer, y As Integer, i As Long

    Dim forx As Integer, fory As Integer
    Dim xfrom As Integer, xto As Integer
    Dim yfrom As Integer, yto As Integer
    
    Dim linebadcount As Integer
    
    Dim bMeasureStart As Boolean
    Dim END_VALUE As Integer                                'search Á¾·á °ª ÀúÀå
    
    If Stop_Measure Then
        MeasureChipCount = Stop_MeasureChipCount
        MeasureChipCountOk = Stop_MeasureChipCountOk
        bRight = Stop_Right
        
        Erase MeasureSeq
        For i = 0 To 200000
            MeasureSeq(i) = Stop_MeasureSeq(i)
        Next
        
        GoTo MeasureStart
    Else
        x = xx + StarProbe.StartChip.x
        y = yy + StarProbe.StartChip.y
        
        StarProbe.MeasureStartX = x
        StarProbe.MeasureStartY = y
        
        NG = False
        
        Erase MeasureSeq
        MeasureChipCount = 0
        MeasureChipCountOk = 0
        
        XAxis = xx
        YAxis = yy
    End If
    
    Call StarProbe_XY_Moving((XAxis), (YAxis))
    Call ChipPosition((XAxis + StarProbe.StartChip.x), (YAxis + StarProbe.StartChip.y))
    
    If ErrorStop = True Then Exit Function           ' Ãß°¡
   
    SSPanel2(9).Caption = result & Space(1) & "ms"
    SSPanel2(1).Caption = Test_Cnt
    Text_TotalCount.Text = Test_Cnt
    
    If Test_Cnt = 0 Then
        SSPanel2(2).Caption = Good_Cnt & Space(1) & "(" & "0.00" & "%)"
    Else
        SSPanel2(2).Caption = Good_Cnt & Space(1) & "(" & Format(Good_Cnt / Test_Cnt * 100, "0.00") & "%)"
    End If
    Text_GoodCount.Text = Good_Cnt
    
    SSPanel_BadCount.Caption = StarProbe.CountBadDie
    Text_BadCount.Text = StarProbe.CountBadDie
    SSPanel_GoodCount.Caption = StarProbe.CountGoodDie
         
    If Not StarProbe_Motor_End_check Then
       If StarProbe.MeasureAll = 1 Or (XPitch(TT_NO) = 1 And YPitch(TT_NO) = 1) Then
            If StarProbe.Ink_After = 0 Then
                If InkRun_Left((XAxis), (YAxis)) Then
                    Call StarProbe_Left_Ink_Dot(StarProbe.Ink_LeftPort)
                    Call InkRun_LeftOk((XAxis), (YAxis))
                End If
                If InkRun_Right((XAxis), (YAxis)) Then
                    Call StarProbe_Right_Ink_Dot(StarProbe.Ink_RightPort)
                    Call InkRun_RightOk((XAxis), (YAxis))
                End If
            End If
        End If
        
        If Not bBad And (Wafer(x, y).flag = False Or Wafer(x, y + 1).flag = False Or Wafer(x, y + 2).flag = False Or Wafer(x, y + 3).flag = False) Then     '[ 2021.10.27 ] : bin clearÈÄ ¿Àµ¿ÀÛÇÏ´Â ºÎºÐ °ü·Ã ¼öÁ¤.
            If PROD.PGM_CHECK = False Then
                If StarProbe.Tip_Clean = 1 Then
                    StarProbe.Tipclean_Count = StarProbe.Tipclean_Count + 1
                    Text7.Refresh
                    Text7.Text = StarProbe.Tipclean_Count
                End If
                TEST
                '========================================================================================
                If (Test_Cnt Mod val(Text10.Text)) < 4 And Test_Cnt <> 0 Then
                    Call StarProbe_FileSave_Data("c:\Star Probe\Temp.SP")           '1000¹ø¸¶´Ù ÀúÀå
                End If
                '========================================================================================
            End If
            
            If StarProbe.MeasureAll = 1 Or (XPitch(TT_NO) = 1 And YPitch(TT_NO) = 1) Then
                If result < 30 Then Sleep 20
            End If
        End If
        
        StarProbe_WorkDateTime_To = CDate(Date$ & " " & Time$)
          
        Call StarProbe_WorkDateTime_HMS(StarProbe_WorkDateTime_Total + DateDiff("S", StarProbe_WorkDateTime_From, StarProbe_WorkDateTime_To))
                  
        SSPanel_DateTime.Caption = StarProbe_WorkDateTime.D & " Day " & _
                                   StarProbe_WorkDateTime.h & ":" & _
                                   StarProbe_WorkDateTime.M & ":" & _
                                   StarProbe_WorkDateTime.s

        SSPanel_Yield.Caption = Format(((StarProbe.CountGoodDie / (StarProbe.CountGoodDie + StarProbe.CountBadDie)) * 100), "00.00") & "%"
        
        Center_fail = True             '[ 2022.05.06 ] : 4Ã¤³Î Áß Áß°£¿¡¼­ ºÒ·®ÀÌ ¹ß»ýÇÑ°ÇÁö ¿©ºÎ
        For i = 0 To 3
            If Text_Chip(i).Text = "Chip" Then
                Wafer(x, ((y + 3) - i)).flag = True
                If Text_ChipBIN(i).Text = "" Then
                    Wafer(x, ((y + 3) - i)).BIN = 0
                    Text_Bin_Count_No(Wafer(x, ((y + 3) - i)).BIN) = Text_Bin_Count_No(Wafer(x, ((y + 3) - i)).BIN) + 1
                    Text_BinCount(Wafer(x, ((y + 3) - i)).BIN).Text = Text_Bin_Count_No(Wafer(x, ((y + 3) - i)).BIN)
                Else
                    Wafer(x, ((y + 3) - i)).BIN = Int(val(Text_ChipBIN(i).Text))
                    Text_Bin_Count_No(Wafer(x, ((y + 3) - i)).BIN) = Text_Bin_Count_No(Wafer(x, ((y + 3) - i)).BIN) + 1
                    Text_BinCount(Wafer(x, ((y + 3) - i)).BIN).Text = Text_Bin_Count_No(Wafer(x, ((y + 3) - i)).BIN)
                End If
                Bin_Count(Wafer(x, ((y + 3) - i)).BIN) = Bin_Count(Wafer(x, ((y + 3) - i)).BIN) + 1     '2016.03.11
                Wafer(x, ((y + 3) - i)).ChipMeasure = True
                Wafer(x, ((y + 3) - i)).MeasureWait = False
                Test_Cnt = Test_Cnt + 1
                If Wafer(x, ((y + 3) - i)).BIN = GOOD_BIN_NO Then                 'pass
                    
                    Good_Cnt = Good_Cnt + 1
                    Wafer(x, ((y + 3) - i)).FlagBad = False
                    If i = 0 Then
                        Test_Fail_Count4 = 1
                    ElseIf i = 1 Then
                        Test_Fail_Count3 = 1
                    ElseIf i = 2 Then
                        Test_Fail_Count2 = 1
                    ElseIf i = 3 Then
                        Test_Fail_Count1 = 1
                    End If
                    
                    SSPanel2(3).Caption = Wafer(x, ((y + 3) - i)).BIN
                    SSPanel2(4).Caption = "PASS"
                    SSPanel2(4).BackColor = &HFF00&
                Else                                                    'fail
                    If i = 0 Or i = 3 Then                      '[ 2022.05.06 ] : Áß°£¿¡¼­ ºÒ·®ÀÌ ¹ß»ýÇÑ °æ¿ì
                        Center_fail = False
                    End If
                    Fail_Find = True
                    Wafer(x, ((y + 3) - i)).FlagBad = True
                    StarProbe.CountBadDie = StarProbe.CountBadDie + 1
                    StarProbe.CountGoodDie = StarProbe.CountGoodDie - 1
                    If i = 0 Then
                        Test_Fail_Count4 = Test_Fail_Count4 + 1
                    ElseIf i = 1 Then
                        Test_Fail_Count3 = Test_Fail_Count3 + 1
                    ElseIf i = 2 Then
                        Test_Fail_Count2 = Test_Fail_Count2 + 1
                    ElseIf i = 3 Then
                        Test_Fail_Count1 = Test_Fail_Count1 + 1
                    End If
                    
                    SSPanel2(3).Caption = Wafer(x, ((y + 3) - i)).BIN
                    SSPanel2(4).Caption = "FAIL"
                    SSPanel2(4).BackColor = &HFF&
                End If
            End If
        Next i
        
        Text_ReciveData.Text = Empty
        Text_ChipBIN(0).Text = Empty
        Text_ChipBIN(1).Text = Empty
        Text_ChipBIN(2).Text = Empty
        Text_ChipBIN(3).Text = Empty
 
        SSPanel_BadCount.Caption = StarProbe.CountBadDie
        Text_BadCount.Text = StarProbe.CountBadDie
        SSPanel_GoodCount.Caption = StarProbe.CountGoodDie
        
        VScroll_Zoom.value = (y / ((Abs(StarProbe.Max.y) + Abs(StarProbe.Min.y) + 1))) * 1000
        HScroll_Zoom.value = (x / ((Abs(StarProbe.Max.x) + Abs(StarProbe.Min.x) + 1))) * 1000
        
        StarProbe.CurrentChip.x = x - StarProbe.StartChip.x
        StarProbe.CurrentChip.y = y - StarProbe.StartChip.y
        
        Text5 = StarProbe.CurrentChip.x
        Text6 = StarProbe.CurrentChip.y
        
        Shape_Chip.Top = (y * StarProbe.DisplayChipSizeY) - 2
        Shape_Chip.Left = (x * StarProbe.DisplayChipSizeX) - 2
                
        Label_ChipPosition = StarProbe.CurrentChip.x & "/" & StarProbe.CurrentChip.y
                            
        Call Display_Chip(pZoom, pOriginal, xx, yy)
        Call Display_Chip(pZoom, pOriginal, xx, (yy + 1))
        Call Display_Chip(pZoom, pOriginal, xx, (yy + 2))
        Call Display_Chip(pZoom, pOriginal, xx, (yy + 3))
        
        '[ 2022.05.19 ] : ¿¬¼ÓºÒ·® ÈÄ clean tipÇÑ´ÙÀ½ µÎ¹øÂ° ÃøÁ¤
        If tip_clean_count_flag_1 = 2 And Test_Fail_Count1 > 1 Then
            tip_clean_count_flag_1 = 0
            MsgBox "¿¬¼ÓºÒ·®ÀÌ ¹ß»ýÇÏ¿´½À´Ï´Ù Tip Å¬¸®´×À» ½Ç½ÃÇØ ÁÖ¼¼¿ä !", 16, "STAR PROBE"
            Test_Fail_Count1 = 1
            bStop = True
            Check1(3).value = 0
            tip_clean_count_flag_1 = 0
            tip_clean_count_flag_2 = 0
            tip_clean_count_flag_3 = 0
            tip_clean_count_flag_4 = 0
            Test_Fail_Count1 = 0
            Test_Fail_Count2 = 0
            Test_Fail_Count3 = 0
            Test_Fail_Count4 = 0
            Exit Function
        ElseIf tip_clean_count_flag_2 = 2 And Test_Fail_Count2 > 1 Then
            tip_clean_count_flag_2 = 0
            MsgBox "¿¬¼ÓºÒ·®ÀÌ ¹ß»ýÇÏ¿´½À´Ï´Ù Tip Å¬¸®´×À» ½Ç½ÃÇØ ÁÖ¼¼¿ä !", 16, "STAR PROBE"
            Test_Fail_Count2 = 1
            bStop = True
            Check1(3).value = 0
            tip_clean_count_flag_1 = 0
            tip_clean_count_flag_2 = 0
            tip_clean_count_flag_3 = 0
            tip_clean_count_flag_4 = 0
            Test_Fail_Count1 = 0
            Test_Fail_Count2 = 0
            Test_Fail_Count3 = 0
            Test_Fail_Count4 = 0
            Exit Function
        ElseIf tip_clean_count_flag_3 = 2 And Test_Fail_Count3 > 1 Then
            tip_clean_count_flag_3 = 0
            MsgBox "¿¬¼ÓºÒ·®ÀÌ ¹ß»ýÇÏ¿´½À´Ï´Ù Tip Å¬¸®´×À» ½Ç½ÃÇØ ÁÖ¼¼¿ä !", 16, "STAR PROBE"
            Test_Fail_Count3 = 1
            bStop = True
            Check1(3).value = 0
            tip_clean_count_flag_1 = 0
            tip_clean_count_flag_2 = 0
            tip_clean_count_flag_3 = 0
            tip_clean_count_flag_4 = 0
            Test_Fail_Count1 = 0
            Test_Fail_Count2 = 0
            Test_Fail_Count3 = 0
            Test_Fail_Count4 = 0
            Exit Function
        ElseIf tip_clean_count_flag_4 = 2 And Test_Fail_Count4 > 1 Then
            tip_clean_count_flag_4 = 0
            MsgBox "¿¬¼ÓºÒ·®ÀÌ ¹ß»ýÇÏ¿´½À´Ï´Ù Tip Å¬¸®´×À» ½Ç½ÃÇØ ÁÖ¼¼¿ä !", 16, "STAR PROBE"
            Test_Fail_Count4 = 1
            bStop = True
            Check1(3).value = 0
            tip_clean_count_flag_1 = 0
            tip_clean_count_flag_2 = 0
            tip_clean_count_flag_3 = 0
            tip_clean_count_flag_4 = 0
            Test_Fail_Count1 = 0
            Test_Fail_Count2 = 0
            Test_Fail_Count3 = 0
            Test_Fail_Count4 = 0
            Exit Function
'        Else
'            tip_clean_count_flag_1 = 0
'            tip_clean_count_flag_2 = 0
'            tip_clean_count_flag_3 = 0
'            tip_clean_count_flag_4 = 0
        End If
        
        '[ 2022.05.19 ] : ¿¬¼ÓºÒ·® ÈÄ clean tipÀ» ÇÑ´ÙÀ½ ´Ù½Ã ºÒ·®ÀÎ °æ¿ì
        If tip_clean_count_flag_1 = 1 And Test_Fail_Count1 > 1 Then
            tip_clean_count_flag_1 = 2
            Call StarProbe_tip_clean
            Sleep 500
        ElseIf tip_clean_count_flag_2 = 1 And Test_Fail_Count2 > 1 Then
            tip_clean_count_flag_2 = 2
            Call StarProbe_tip_clean
            Sleep 500
        ElseIf tip_clean_count_flag_3 = 1 And Test_Fail_Count3 > 1 Then
            tip_clean_count_flag_3 = 2
            Call StarProbe_tip_clean
            Sleep 500
        ElseIf tip_clean_count_flag_4 = 1 And Test_Fail_Count4 > 1 Then
            tip_clean_count_flag_4 = 2
            Call StarProbe_tip_clean
            Sleep 500
        Else
            tip_clean_count_flag_1 = 0
            tip_clean_count_flag_2 = 0
            tip_clean_count_flag_3 = 0
            tip_clean_count_flag_4 = 0
        End If
        
        
        If StarProbe.Probe_Stop = 1 And Sample_No_Ink = False Then
            If Test_Fail_Count1 > StarProbe.Probe_Stop_Tfail_Count Then
                If tip_clean_count_flag_1 = 0 Then
                    tip_clean_count_flag_1 = 1
                    Call StarProbe_tip_clean
                    Sleep 500
                Else
'                    MsgBox " Test4 Continus Fail  check !", 16, "STAR PROBE"
'                    Test_Fail_Count1 = 1
'                    bStop = True
'                    Check1(3).value = 0
                End If
            ElseIf Test_Fail_Count2 > StarProbe.Probe_Stop_Tfail_Count Then
                If tip_clean_count_flag_2 = 0 Then
                    tip_clean_count_flag_2 = 1
                    Call StarProbe_tip_clean
                    Sleep 500
                Else
'                    MsgBox " Test3 Continus Fail  check !", 16, "STAR PROBE"
'                    Test_Fail_Count2 = 1
'                    bStop = True
'                    Check1(3).value = 0
                End If
            ElseIf Test_Fail_Count3 > StarProbe.Probe_Stop_Tfail_Count Then
                If tip_clean_count_flag_3 = 0 Then
                    tip_clean_count_flag_3 = 1
                    Call StarProbe_tip_clean
                    Sleep 500
                Else
'                    MsgBox " Test2 Continus Fail  check !", 16, "STAR PROBE"
'                    Test_Fail_Count3 = 1
'                    bStop = True
'                    Check1(3).value = 0
                End If
            ElseIf Test_Fail_Count4 > StarProbe.Probe_Stop_Tfail_Count Then
                If tip_clean_count_flag_4 = 0 Then
                    tip_clean_count_flag_4 = 1
                    Call StarProbe_tip_clean
                    Sleep 500
                    Test_Fail_Count1 = 0
                Else
'                    MsgBox " Test1 Continus Fail  check !", 16, "STAR PROBE"
'                    Test_Fail_Count4 = 1
'                    bStop = True
'                    Check1(3).value = 0
                End If
            End If
        End If
        
        ' ÇöÀç Ä¨ ÃøÁ¤ °á°ú°¡ ¾çÇ°ÀÌ³ª ºÒ·®ÀÌ ³ª¿Íµµ
        ' ÁÖÀ§¿¡ Ä¨À» ¸ðµÎ ÃøÁ¤ÇÏ¿´´ÂÁö ¾ÈÇÏ¿´´ÂÁö´Â ÀÐ¾î¼­ ´ÙÀ½Ã³¸® ÇÑ´Ù.
       
        If bBad Then
            NG = True
        Else
            If StarProbe.ReMeasure = vbChecked Then
                If Not bChip Then
                    NG = Not StarProbe_MeasureLineOk_CH4(x, y, StarProbe.LineOk)
                Else
                    NG = True
                End If
            Else
                NG = bChip
            End If
        End If
        If StarProbe.MeasureAll = 1 Or Sample_No_Ink = True Or (XPitch(TT_NO) = 1 And YPitch(TT_NO) = 1) Then NG = False
        
        bBad = False
        
        ' 2005.09.01
        MeasureChipCount = 0
        MeasureChipCountOk = 0
       
        ' Å×½ºÆ® ÃøÁ¤ °á°ú°¡ NGÀÌ¸é ÃøÁ¤ ÃßÀû ¾Ë°í¸®ÁòÀ» Àû¿ëÇÏ¶ó.
        M_CNT = 0
        YOON_CNT = 0
        BACK_X = x
        BACK_Y = y
        
        If Sample_No_Ink = False Then
            If (StarProbe.LimitArea = 1 And Fail_Find = True) Then
REMEA:
                MeasureSeq(0).x = x
                MeasureSeq(0).y = y
                MeasureSeq(0).r = True
                    
                If bRight Then
                    Call StarProbe_MeasureLineRight_CH4(x, y, StarProbe.RCount)
                Else
                    Call StarProbe_MeasureLineLeft_CH4(x, y, StarProbe.RCount)
                End If

MeasureStart:
                bMeasureStart = True
                
                'Ãß°¡
                XVAL = XPitch(TT_NO) * 4
                XVAL1 = XVAL - 4
                XVAL2 = XVAL - 8
                          
                Do While bMeasureStart
                    DoEvents
                    '========================================================================================================= [ 2 <= Xpitch <= 30 ]
                    If ((XPitch(TT_NO) > 2) And (XPitch(TT_NO) <= 30)) Then
                        If (((MeasureChipCountOk Mod 4 = 0) And (MeasureChipCountOk > 0 And (MeasureChipCountOk <= XVAL2)) And Fail_Find = True) Or (MeasureChipCountOk = XVAL1)) Then
                            y = y + 3
                            Fail_Find = False
                            X_STT = 1
                            Move_CNT = 1
                            M_CNT = 1
                        ElseIf (MeasureChipCountOk Mod 4 = 0) And (MeasureChipCountOk > 0) And (MeasureChipCountOk < XVAL) And Fail_Find = False Then
                            y = y + 3
                            Fail_Find = False
                            X_STT = 1
                            Move_CNT = 1
                            M_CNT = 1
                            If YOON_CNT > 0 Then
                                If MeasureChipCountOk <= 8 Then
                                    MeasureChipCountOk = MeasureChipCountOk
                                Else
                                    MeasureChipCountOk = XVAL1 + 4
                                End If
                            Else
                                MeasureChipCountOk = MeasureChipCountOk '+ (XVAL1 - MeasureChipCountOk)
                            End If
                            MeasureChipCount = 0
                        ElseIf (MeasureChipCountOk = XVAL) And Fail_Find = True Then
                            Fail_Find = False
                            X_STT = 1
                            Move_CNT = 1
                            
                            If YOON_CNT > 0 Then
                                MeasureChipCountOk = MeasureChipCountOk
                                MeasureChipCount = 0
                            Else
                                MeasureChipCountOk = 0
                                MeasureChipCount = 0
                                GoTo REMEA
                            End If
                        End If
                        If YOON_CNT > 0 Then
                            If (MeasureChipCountOk = XVAL + 4) And Fail_Find = True Then
                                Fail_Find = False
                                X_STT = 1
                                Move_CNT = 1
                                                            
                                MeasureChipCountOk = 0
                                MeasureChipCount = 0
                                GoTo REMEA
                            End If
                        End If
                    ElseIf XPitch(TT_NO) = 2 Then
                        If MeasureChipCountOk = 4 Then
                            Fail_Find = False
                            X_STT = 1
                            Move_CNT = 1
                            M_CNT = 1
                        ElseIf (MeasureChipCountOk = XVAL) And Fail_Find = True Then
                            Fail_Find = False
                            X_STT = 1
                            Move_CNT = 1
                            
                            If YOON_CNT > 0 Then
                                MeasureChipCountOk = MeasureChipCountOk
                                MeasureChipCount = 0
                            Else
                                MeasureChipCountOk = 0
                                MeasureChipCount = 0
                                GoTo REMEA
                            End If
                        End If
                        If YOON_CNT > 0 Then
                            If (MeasureChipCountOk = XVAL + 4) And Fail_Find = True Then
                                Fail_Find = False
                                X_STT = 1
                                Move_CNT = 1
                                                            
                                MeasureChipCountOk = 0
                                MeasureChipCount = 0
                                GoTo REMEA
                            End If
                        End If
                    End If
                    '=========================================================================================================
                      
                    If bStop = True Then
                        Stop_Measure = True
                        Stop_MeasureChipCount = MeasureChipCount
                        Stop_MeasureChipCountOk = MeasureChipCountOk
                        Stop_Right = bRight
                        
                        Erase Stop_MeasureSeq
                        
                        For i = 0 To 200000
                            Stop_MeasureSeq(i) = MeasureSeq(i)
                        Next
                    
                        Exit Do
                    End If
                      
                    MeasureChipCountOk = MeasureChipCountOk + 1
                    
                    If YOON_CNT > 0 Then
                        END_VALUE = XPitch(TT_NO) * 4 + 4
                    Else
                        END_VALUE = XPitch(TT_NO) * 4
                    End If
                    '//////////////////////////////////////////////////////////////////////////////////[6´Ü]
                    DoEvents
                    
                    If MeasureChipCountOk > END_VALUE Then
                        If StarProbe.LimitArea = 1 Then         '[°øÅë]
                            Fail_Loop = False
                            X_STT = 1
                            Fail_Find = False
                        End If
                        Stop_Measure = False
                        MeasureChipCountOk = 0
                        
                        '[ 2022.05.06 ] :Ãß°¡
                        If YOON_CNT = (LOOP_COUNT * 2) Then 'loopÁ¾·á
                            FAIL_COUNT = 0
                            Exit Do
                        End If
                        'Áß°£
                        If YOON_CNT = 0 Then                    '[Áß°£->»ó´Ü1]
                            If Center_fail = True Then
                                FAIL_COUNT = 0
                                Exit Do
                            Else
                                YOON_CNT = YOON_CNT + 1
                                y = y - 4
                                x = BACK_X
                                FAIL_COUNT = 0
                                GoTo REMEA
                            End If
                        End If
                                                                        
                        '»ó´Ü½ÃÀÛ
                        If YOON_CNT <= LOOP_COUNT Then           '1 ~ LOOP_COUNT
                            If YOON_CNT = LOOP_COUNT Then       '[»ó´Ü8->ÇÏ´Ü1]
                                'y = y + (4 + (4 * YOON_CNT))       '2016.09.22
                                y = BACK_Y + 4
                                YOON_CNT = YOON_CNT + 1
                                x = BACK_X
                                FAIL_COUNT = 0
                                GoTo REMEA
                            Else
                                If FAIL_COUNT = 0 Then          '¸ðµÎ¾çÇ°ÀÌ¸é ÇÏ´Ü1·Î ÀÌµ¿
                                    'y = y + (4 + (4 * YOON_CNT))   '2016.09.22
                                    y = BACK_Y + 4
                                    YOON_CNT = LOOP_COUNT + 1   'ÇÏ´Ü½ÃÀÛÀ¸·Î ÀÌµ¿
                                Else
                                    y = y - 4
                                    YOON_CNT = YOON_CNT + 1
                                End If
                                x = BACK_X
                                FAIL_COUNT = 0
                                GoTo REMEA
                            End If
                        'ÇÏ´Ü½ÃÀÛ
                        ElseIf YOON_CNT < (LOOP_COUNT * 2) Then 'LOOP_COUNT ~ LOOP_COUNT*2
                            If FAIL_COUNT = 0 Then Exit Do      '¸ðµÎ¾çÇ°ÀÌ¸é loop out
                            YOON_CNT = YOON_CNT + 1
                            y = y + 4
                            x = BACK_X
                            FAIL_COUNT = 0
                            GoTo REMEA
                        ElseIf YOON_CNT = (LOOP_COUNT * 2) Then 'loopÁ¾·á
                            FAIL_COUNT = 0
                            Exit Do
                        End If
                    End If
                    '//////////////////////////////////////////////////////////////////////////////////////
                      
                    ' Tester Measure Start
                    x = MeasureSeq(MeasureChipCountOk).x
                    y = MeasureSeq(MeasureChipCountOk).y
                                        
                    If StarProbe.LimitArea = 1 Then
                        VScroll_Zoom.value = (y / ((Abs(StarProbe.Max.y) + Abs(StarProbe.Min.y) + 1))) * 1000
                    Else
                        ' wafer scroll move
                        VScroll_Zoom.value = (y / ((Abs(StarProbe.Max.y) + Abs(StarProbe.Min.y) + 1))) * 1000
                        HScroll_Zoom.value = (x / ((Abs(StarProbe.Max.x) + Abs(StarProbe.Min.x) + 1))) * 1000
                    End If
                    
                    StarProbe.CurrentChip.x = x - StarProbe.StartChip.x
                    StarProbe.CurrentChip.y = y - StarProbe.StartChip.y
                    
                    Text5 = StarProbe.CurrentChip.x
                    Text6 = StarProbe.CurrentChip.y
                                    
                    XAxis = StarProbe.CurrentChip.x
                    YAxis = StarProbe.CurrentChip.y
                    
                    move_ok = False
                    If Wafer(x, y).ChipSkipDie = True And Wafer(x, y + 1).ChipSkipDie = True And Wafer(x, y + 2).ChipSkipDie = True And Wafer(x, y + 3).ChipSkipDie = True Then
                        move_ok = True
                    Else
                        For yyy = 0 To 3
                            If Wafer(x, y + yyy).flag = True Then
                                move_ok = True
                                Exit For
                            End If
                        Next yyy
                        
                        '[ 2022.05.06 ] : chipÀÌ ¾Æ´Ñ °æ¿ì Ã³¸® Ãß°¡
                        If Wafer(x, y).Chip = False And Wafer(x, y + 1).Chip = False And Wafer(x, y + 1).Chip = False And Wafer(x, y + 1).Chip = False Then
                            move_ok = True
                        End If
                    End If
                    
                    If move_ok = False Then
                        Call StarProbe_XY_Moving((XAxis), (YAxis))
                        
                        If Not StarProbe_Motor_End_check Then
                            Label_ChipPosition = StarProbe.CurrentChip.x & "/" & StarProbe.CurrentChip.y
                            
                            Shape_Chip.Top = (y * (StarProbe.DisplayChipSizeY)) - 2
                            Shape_Chip.Left = (x * StarProbe.DisplayChipSizeX) - 2
                                            
                            SSPanel2(9).Caption = result & Space(1) & "ms"
                            SSPanel2(1).Caption = Test_Cnt
                            Text_TotalCount.Text = Test_Cnt
                            If Test_Cnt = 0 Then
                                SSPanel2(2).Caption = Good_Cnt & Space(1) & "(" & "0.00" & "%)"
                            Else
                                SSPanel2(2).Caption = Good_Cnt & Space(1) & "(" & Format(Good_Cnt / Test_Cnt * 100, "0.00") & "%)"
                            End If
                            Text_GoodCount.Text = Good_Cnt
                             
                            SSPanel_BadCount.Caption = StarProbe.CountBadDie
                            Text_BadCount.Text = StarProbe.CountBadDie
                            SSPanel_GoodCount.Caption = StarProbe.CountGoodDie
                            
                            If ErrorStop = True Then
                                Stop_Measure = True
                            
                                Stop_MeasureChipCount = MeasureChipCount - 1
                                Stop_MeasureChipCountOk = MeasureChipCountOk
                                Stop_Right = bRight
                                
                                Erase Stop_MeasureSeq
                                
                                For i = 0 To 200000
                                    Stop_MeasureSeq(i) = MeasureSeq(i)
                                Next
                                Exit Do   'Ãß°¡
                            End If
                            
                            Call ChipPosition((XAxis + StarProbe.StartChip.x), (YAxis + StarProbe.StartChip.y))
                                
                            If StarProbe.Ink_After = 0 Then
                                If InkRun_Left((XAxis), (YAxis)) Then
                                    Call StarProbe_Left_Ink_Dot(StarProbe.Ink_LeftPort)
                                    Call InkRun_LeftOk((XAxis), (YAxis))
                                End If
                                If InkRun_Right((XAxis), (YAxis)) Then
                                    Call StarProbe_Right_Ink_Dot(StarProbe.Ink_RightPort)
                                    Call InkRun_RightOk((XAxis), (YAxis))
                                End If
                            End If
                            Dim GOGO As Boolean
                            GOGO = False
                            If Text_ChipTest.Text <> 0 Then
                                For XXX = 0 To 3
                                    If Wafer(x, y + XXX).Chip And Not Wafer(x, y + XXX).ChipMask And Not Wafer(x, y + XXX).ChipInk And Not Wafer(x, y + XXX).ChipSkipDie And Not Wafer(x, y + XXX).ChipPlate Then
                                        GOGO = True
                                    Else
                                        'GOGO = True
                                    End If
                                Next XXX
                                If GOGO = True Then
                                'If Wafer(x, y).Chip And Not Wafer(x, y).ChipMask And Not Wafer(x, y).ChipInk And Not Wafer(x, y).ChipSkipDie And Not Wafer(x, y).ChipPlate Then                                                                                                     'Á¤»óÄ¨ÀÎ °æ¿ì
                                    If Not Wafer(x, y).flag Or Not Wafer(x, y + 1).flag Or Not Wafer(x, y + 2).flag Or Not Wafer(x, y + 3).flag Then
                                        If PROD.PGM_CHECK = False Then
                                            TEST
                                            '========================================================================================
                                            If (Test_Cnt Mod val(Text10.Text)) < 4 And Test_Cnt <> 0 Then
                                                Call StarProbe_FileSave_Data("c:\Star Probe\Temp.SP")           '1000¹ø¸¶´Ù ÀúÀå
                                            End If
                                            '========================================================================================
                                        End If
                                            
                                        If result < 30 Then Sleep 20
                                    
                                        StarProbe_WorkDateTime_To = CDate(Date$ & " " & Time$)
                        
                                        Call StarProbe_WorkDateTime_HMS(StarProbe_WorkDateTime_Total + DateDiff("S", StarProbe_WorkDateTime_From, StarProbe_WorkDateTime_To))
                                                
                                        SSPanel_DateTime.Caption = StarProbe_WorkDateTime.D & " Day " & _
                                                                   StarProbe_WorkDateTime.h & ":" & _
                                                                   StarProbe_WorkDateTime.M & ":" & _
                                                                   StarProbe_WorkDateTime.s
                                  
                                        SSPanel_Yield.Caption = Format(((StarProbe.CountGoodDie / (StarProbe.CountGoodDie + StarProbe.CountBadDie)) * 100), "00.00") & "%"
                                        
                                        For i = 0 To 3
                                            If Text_Chip(i).Text = "Chip" Then
                                                Wafer(x, ((y + 3) - i)).flag = True
                                                If Text_ChipBIN(i).Text = "" Then
                                                    Wafer(x, ((y + 3) - i)).BIN = 0
                                                Else
                                                    Wafer(x, ((y + 3) - i)).BIN = Int(val(Text_ChipBIN(i).Text))
                                                End If
                                                Text_Bin_Count_No(Wafer(x, ((y + 3) - i)).BIN) = Text_Bin_Count_No(Wafer(x, ((y + 3) - i)).BIN) + 1
                                                Text_BinCount(Wafer(x, ((y + 3) - i)).BIN).Text = Text_Bin_Count_No(Wafer(x, ((y + 3) - i)).BIN)
                                                    
                                                Wafer(x, ((y + 3) - i)).ChipMeasure = True
                                                Wafer(x, ((y + 3) - i)).MeasureWait = False
                                                Test_Cnt = Test_Cnt + 1
                                                If Wafer(x, ((y + 3) - i)).BIN = GOOD_BIN_NO Then
                                                    Good_Cnt = Good_Cnt + 1
                                                    Wafer(x, ((y + 3) - i)).FlagBad = False
                                                    If i = 0 Then
                                                        Test_Fail_Count4 = 1
                                                    ElseIf i = 1 Then
                                                        Test_Fail_Count3 = 1
                                                    ElseIf i = 2 Then
                                                        Test_Fail_Count2 = 1
                                                    ElseIf i = 3 Then
                                                        Test_Fail_Count1 = 1
                                                    End If
                                                    
                                                    SSPanel2(3).Caption = Wafer(x, ((y + 3) - i)).BIN
                                                    SSPanel2(4).Caption = "PASS"
                                                    SSPanel2(4).BackColor = &HFF00&
                                                Else
                                                    Fail_Find = True
                                                    FAIL_COUNT = FAIL_COUNT + 1         'Ãß°¡
                                                    Wafer(x, ((y + 3) - i)).FlagBad = True
                                                    StarProbe.CountBadDie = StarProbe.CountBadDie + 1
                                                    StarProbe.CountGoodDie = StarProbe.CountGoodDie - 1
                                                    If i = 0 Then
                                                        Test_Fail_Count4 = Test_Fail_Count4 + 1
                                                    ElseIf i = 1 Then
                                                        Test_Fail_Count3 = Test_Fail_Count3 + 1
                                                    ElseIf i = 2 Then
                                                        Test_Fail_Count2 = Test_Fail_Count2 + 1
                                                    ElseIf i = 3 Then
                                                        Test_Fail_Count1 = Test_Fail_Count1 + 1
                                                    End If
                                                    
                                                    SSPanel2(3).Caption = Wafer(x, ((y + 3) - i)).BIN
                                                    SSPanel2(4).Caption = "FAIL"
                                                    SSPanel2(4).BackColor = &HFF&
                                                End If
                                            End If
                                        Next i
                                         
                                        Text_ReciveData.Text = Empty
                                        Text_ChipBIN(0).Text = Empty
                                        Text_ChipBIN(1).Text = Empty
                                        Text_ChipBIN(2).Text = Empty
                                        Text_ChipBIN(3).Text = Empty
                                
                                        SSPanel_BadCount.Caption = StarProbe.CountBadDie
                                        Text_BadCount.Text = StarProbe.CountBadDie
                                        SSPanel_GoodCount.Caption = StarProbe.CountGoodDie
                                       
                                        VScroll_Zoom.value = (y / ((Abs(StarProbe.Max.y) + Abs(StarProbe.Min.y) + 1))) * 1000
                                        HScroll_Zoom.value = (x / ((Abs(StarProbe.Max.x) + Abs(StarProbe.Min.x) + 1))) * 1000
                                       
                                        StarProbe.CurrentChip.x = x - StarProbe.StartChip.x
                                        StarProbe.CurrentChip.y = y - StarProbe.StartChip.y
                                       
                                        Text5 = StarProbe.CurrentChip.x
                                        Text6 = StarProbe.CurrentChip.y
                                       
                                        Shape_Chip.Top = (y * StarProbe.DisplayChipSizeY) - 2
                                        Shape_Chip.Left = (x * StarProbe.DisplayChipSizeX) - 2
                                                                      
                                        Label_ChipPosition = StarProbe.CurrentChip.x & "/" & StarProbe.CurrentChip.y
                                                                                                                
                                        Call Display_Chip(pZoom, pOriginal, Text5, Text6)
                                        Call Display_Chip(pZoom, pOriginal, Text5, (Text6 + 1))
                                        Call Display_Chip(pZoom, pOriginal, Text5, (Text6 + 2))
                                        Call Display_Chip(pZoom, pOriginal, Text5, (Text6 + 3))
                                        
                                        If StarProbe.Probe_Stop = 1 And Sample_No_Ink = False Then
                                            If Test_Fail_Count1 > StarProbe.Probe_Stop_Tfail_Count Then
                                                MsgBox " Test4 Continus Fail  check !", 16, "STAR PROBE"
                                                Test_Fail_Count1 = 1
                                                bStop = True
                                                Check1(3).value = 0
                                            ElseIf Test_Fail_Count2 > StarProbe.Probe_Stop_Tfail_Count Then
                                                MsgBox " Test3 Continus Fail  check !", 16, "STAR PROBE"
                                                Test_Fail_Count2 = 1
                                                bStop = True
                                                Check1(3).value = 0
                                            ElseIf Test_Fail_Count3 > StarProbe.Probe_Stop_Tfail_Count Then
                                                MsgBox " Test2 Continus Fail  check !", 16, "STAR PROBE"
                                                Test_Fail_Count3 = 1
                                                bStop = True
                                                Check1(3).value = 0
                                            ElseIf Test_Fail_Count4 > StarProbe.Probe_Stop_Tfail_Count Then
                                                MsgBox " Test1 Continus Fail  check !", 16, "STAR PROBE"
                                                Test_Fail_Count4 = 1
                                                bStop = True
                                                Check1(3).value = 0
                                            End If
                                        End If
                                          
                                        If StarProbe.LimitArea = 1 Then
                                        Else
                                            If NG Then
                                                If bRight Then
                                                    Call StarProbe_MeasureLineRight_CH4(x, y, StarProbe.RCount_Sub)
                                                Else
                                                    Call StarProbe_MeasureLineLeft_CH4(x, y, StarProbe.RCount_Sub)
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                    MeasureChipCountOk = MeasureChipCountOk + 3
                Loop
            End If
        End If
    End If
End Function

Public Function StarProbe_MeasureLineOk_CH4(x As Integer, y As Integer, Optional RCount As Integer) As Boolean
    Dim linecount As Integer
    Dim xx As Integer, yy As Integer
    Dim i As Integer
    
    If RCount = 0 Or RCount = Empty Then RCount = 1

    Erase MeasureLine
    MeasureLine(1).x = x + 1: MeasureLine(1).y = y + 0
    MeasureLine(2).x = x + 1: MeasureLine(2).y = y - 1
    MeasureLine(3).x = x + 0: MeasureLine(3).y = y - 1
    MeasureLine(4).x = x - 1: MeasureLine(4).y = y - 1
    MeasureLine(5).x = x - 1: MeasureLine(5).y = y + 0
    MeasureLine(6).x = x - 1: MeasureLine(6).y = y + 1
    MeasureLine(7).x = x + 0: MeasureLine(7).y = y + 1
    MeasureLine(8).x = x + 1: MeasureLine(8).y = y + 1
    
    If RCount > 1 Then
        i = 9
        For linecount = 1 To RCount
            xx = x + 1 + linecount
            For yy = (y + linecount) To (y - linecount) Step -1
                MeasureLine(i).x = xx: MeasureLine(i).y = yy
                i = i + 1
            Next
            
            yy = y - 1 - linecount
            For xx = (x + 1 + linecount) To (x - 1 - linecount) Step -1
                MeasureLine(i).x = xx: MeasureLine(i).y = yy
                i = i + 1
            Next
            
            xx = x - 1 - linecount
            For yy = (y - linecount) To (y + linecount)
                MeasureLine(i).x = xx: MeasureLine(i).y = yy
                i = i + 1
            Next
            
            yy = y + 1 + linecount
            For xx = (x - 1 - linecount) To (x + 1 + linecount)
                MeasureLine(i).x = xx: MeasureLine(i).y = yy
                i = i + 1
            Next
        Next
    End If
    
    For i = 0 To 72
        Shape_Mea(i).Visible = False
    Next
    
    linecount = 8 + IIf(RCount > 1, 8 * RCount, 0)
    
    Dim iOk As Integer, iNotMeasure As Integer, iBad As Integer
    Dim b As Boolean
    
    iOk = 0
    iNotMeasure = 0
    iBad = 0
    
    For i = 0 To 72
        Shape_Mea(i).Visible = False
    Next
    
    linecount = 8 + IIf(RCount > 1, 8 * RCount, 0)
    
    For i = 1 To linecount
        xx = MeasureLine(i).x
        yy = MeasureLine(i).y
        ' Ä¨ÀÌ ÀÖÀ¸¸é ÃøÁ¤ÇÏ¶ó.
        If Wafer(xx, yy).Chip Then
            If Not Wafer(xx, yy).ChipMask And _
               Not Wafer(xx, yy).ChipSkipDie And _
               Not Wafer(xx, yy).ChipMeasure And _
               Not Wafer(xx, yy).ChipPlate Then
               
                If Wafer(xx, yy).flag Then
                    iOk = iOk + 1
                    If Wafer(xx, yy).FlagBad Then iBad = iBad + 1
                    Shape_Mea(i - 1).Visible = False
                Else
                    iNotMeasure = iNotMeasure + 1
                End If
            Else
                iNotMeasure = iNotMeasure + 1
            End If
        End If
    Next
    
    b = True
    
    If iOk = linecount Or iNotMeasure = linecount Then
        b = True
    ElseIf iBad > 0 Then
        b = False
    End If
    StarProbe_MeasureLineOk = b
End Function

Public Function StarProbe_MeasureLineOk_CH2(x As Integer, y As Integer, Optional RCount As Integer) As Boolean
    Dim linecount As Integer
    Dim xx As Integer, yy As Integer
    Dim i As Integer
    
    If RCount = 0 Or RCount = Empty Then RCount = 1

    Erase MeasureLine
    MeasureLine(1).x = x + 1: MeasureLine(1).y = y + 0
    MeasureLine(2).x = x + 1: MeasureLine(2).y = y - 1
    MeasureLine(3).x = x + 0: MeasureLine(3).y = y - 1
    MeasureLine(4).x = x - 1: MeasureLine(4).y = y - 1
    MeasureLine(5).x = x - 1: MeasureLine(5).y = y + 0
    MeasureLine(6).x = x - 1: MeasureLine(6).y = y + 1
    MeasureLine(7).x = x + 0: MeasureLine(7).y = y + 1
    MeasureLine(8).x = x + 1: MeasureLine(8).y = y + 1
    
    If RCount > 1 Then
        i = 9
        For linecount = 1 To RCount
            xx = x + 1 + linecount
            For yy = (y + linecount) To (y - linecount) Step -1
                MeasureLine(i).x = xx: MeasureLine(i).y = yy
                i = i + 1
            Next
            
            yy = y - 1 - linecount
            For xx = (x + 1 + linecount) To (x - 1 - linecount) Step -1
                MeasureLine(i).x = xx: MeasureLine(i).y = yy
                i = i + 1
            Next
            
            xx = x - 1 - linecount
            For yy = (y - linecount) To (y + linecount)
                MeasureLine(i).x = xx: MeasureLine(i).y = yy
                i = i + 1
            Next
            
            yy = y + 1 + linecount
            For xx = (x - 1 - linecount) To (x + 1 + linecount)
                MeasureLine(i).x = xx: MeasureLine(i).y = yy
                i = i + 1
            Next
        Next
    End If
    
    For i = 0 To 72
        Shape_Mea(i).Visible = False
    Next
    
    linecount = 8 + IIf(RCount > 1, 8 * RCount, 0)
    
    Dim iOk As Integer, iNotMeasure As Integer, iBad As Integer
    Dim b As Boolean
    
    iOk = 0
    iNotMeasure = 0
    iBad = 0
    
    For i = 0 To 72
        Shape_Mea(i).Visible = False
    Next
    
    linecount = 8 + IIf(RCount > 1, 8 * RCount, 0)
    
    For i = 1 To linecount
        xx = MeasureLine(i).x
        yy = MeasureLine(i).y
        ' Ä¨ÀÌ ÀÖÀ¸¸é ÃøÁ¤ÇÏ¶ó.
        If Wafer(xx, yy).Chip Then
            If Not Wafer(xx, yy).ChipMask And _
               Not Wafer(xx, yy).ChipSkipDie And _
               Not Wafer(xx, yy).ChipMeasure And _
               Not Wafer(xx, yy).ChipPlate Then
               
                If Wafer(xx, yy).flag Then
                    iOk = iOk + 1
                    If Wafer(xx, yy).FlagBad Then iBad = iBad + 1
                    Shape_Mea(i - 1).Visible = False
                Else
                    iNotMeasure = iNotMeasure + 1
                End If
                
            Else
                iNotMeasure = iNotMeasure + 1
            End If
        End If
    Next
    
    b = True
    
    If iOk = linecount Or iNotMeasure = linecount Then
        b = True
    ElseIf iBad > 0 Then
        b = False
    End If
    StarProbe_MeasureLineOk = b
End Function

Public Function StarProbe_MeasureLineOk_CH1(x As Integer, y As Integer, Optional RCount As Integer) As Boolean
    If RCount = 0 Or RCount = Empty Then RCount = 1

    Erase MeasureLine

    '
    '  4:-1/-1  3:0/-1   2:+1/-1
    '  5:-1/0   0:0/0    1:+1/0
    '  6:-1/+1  7:0/+1   8:+1/+1
    '
    MeasureLine(1).x = x + 1: MeasureLine(1).y = y + 0
    MeasureLine(2).x = x + 1: MeasureLine(2).y = y - 1
    MeasureLine(3).x = x + 0: MeasureLine(3).y = y - 1
    MeasureLine(4).x = x - 1: MeasureLine(4).y = y - 1
    MeasureLine(5).x = x - 1: MeasureLine(5).y = y + 0
    MeasureLine(6).x = x - 1: MeasureLine(6).y = y + 1
    MeasureLine(7).x = x + 0: MeasureLine(7).y = y + 1
    MeasureLine(8).x = x + 1: MeasureLine(8).y = y + 1
    
    Dim linecount As Integer
    Dim xx As Integer, yy As Integer
    Dim i As Integer
    
    If RCount > 1 Then
        i = 9
        For linecount = 1 To RCount
            xx = x + 1 + linecount
            For yy = (y + linecount) To (y - linecount) Step -1
                MeasureLine(i).x = xx: MeasureLine(i).y = yy
                i = i + 1
            Next
            
            yy = y - 1 - linecount
            For xx = (x + 1 + linecount) To (x - 1 - linecount) Step -1
                MeasureLine(i).x = xx: MeasureLine(i).y = yy
                i = i + 1
            Next
            
            xx = x - 1 - linecount
            For yy = (y - linecount) To (y + linecount)
                MeasureLine(i).x = xx: MeasureLine(i).y = yy
                i = i + 1
            Next
            
            yy = y + 1 + linecount
            For xx = (x - 1 - linecount) To (x + 1 + linecount)
                MeasureLine(i).x = xx: MeasureLine(i).y = yy
                i = i + 1
            Next
        Next
    
    End If
    
    For i = 0 To 127
        Shape_Mea(i).Visible = False
    Next
    
    linecount = 8 + IIf(RCount > 1, 8 * RCount, 0)
    
    Dim iOk As Integer, iNotMeasure As Integer, iBad As Integer
    Dim b As Boolean
    
    iOk = 0
    iNotMeasure = 0
    iBad = 0
    
    For i = 0 To 127
        Shape_Mea(i).Visible = False
    Next
    
    linecount = 8 + IIf(RCount > 1, 8 * RCount, 0)
    
    For i = 1 To linecount
        xx = MeasureLine(i).x
        yy = MeasureLine(i).y
    
        ' Ä¨ÀÌ ÀÖÀ¸¸é ÃøÁ¤ÇÏ¶ó.
        If Wafer(xx, yy).Chip Then
            If Not Wafer(xx, yy).ChipMask And Not Wafer(xx, yy).ChipSkipDie And Not Wafer(xx, yy).ChipMeasure And Not Wafer(xx, yy).ChipPlate Then
                If Wafer(xx, yy).flag Then
                    iOk = iOk + 1
                    If Wafer(xx, yy).FlagBad Then iBad = iBad + 1
                    Shape_Mea(i - 1).Visible = False
                Else
                    iNotMeasure = iNotMeasure + 1
                End If
            Else
                iNotMeasure = iNotMeasure + 1
            End If
        End If
    Next
    
    b = True
    
    If iOk = linecount Or iNotMeasure = linecount Then
        b = True
    ElseIf iBad > 0 Then
        b = False
    End If
    StarProbe_MeasureLineOk = b
End Function


Public Sub StarProbe_MeasureLineRight_CH4(x As Integer, y As Integer, Optional RCount As Integer)
    Dim II, jj, kk As Integer
    
    If RCount = 0 Or RCount = Empty Then RCount = 1

    Erase MeasureLine

  
    If XPitch(TT_NO) >= 2 Then
        If YOON_CNT > 0 Then
            MeasureChipCount = 0
            jj = 1
            For II = 1 To (XPitch(TT_NO) * 4)
                For kk = 0 To 3
                    MeasureLine(II + kk).x = x + jj - 1
                    MeasureLine(II + kk).y = y + kk
                    If kk = 3 Then jj = jj + 1
                Next kk
                II = II + 3
            Next II
            
            For II = ((XPitch(TT_NO) * 4) + 1) To (XPitch(TT_NO) * 4 + 4)
                MeasureLine(II).x = x - 1
                If II = (XPitch(TT_NO) * 4) + 1 Then
                    MeasureLine(II).y = y
                ElseIf II = (XPitch(TT_NO) * 4) + 2 Then
                    MeasureLine(II).y = y + 1
                ElseIf II = (XPitch(TT_NO) * 4) + 3 Then
                    MeasureLine(II).y = y + 2
                ElseIf II = (XPitch(TT_NO) * 4) + 4 Then
                    MeasureLine(II).y = y + 3
                End If
            Next II
        Else
            MeasureChipCount = 0
            jj = 1
            For II = 1 To (XPitch(TT_NO) * 4) - 4
                For kk = 0 To 3
                    MeasureLine(II + kk).x = x + jj
                    MeasureLine(II + kk).y = y + kk
                    If kk = 3 Then jj = jj + 1
                Next kk
                II = II + 3
            Next II
            
            'µÞÂÊ
            For II = ((XPitch(TT_NO) * 4) - 3) To (XPitch(TT_NO) * 4)
                MeasureLine(II).x = x - 1
                If II = (XPitch(TT_NO) * 4) - 3 Then
                    MeasureLine(II).y = y
                ElseIf II = (XPitch(TT_NO) * 4) - 2 Then
                    MeasureLine(II).y = y + 1
                ElseIf II = (XPitch(TT_NO) * 4) - 1 Then
                    MeasureLine(II).y = y + 2
                ElseIf II = (XPitch(TT_NO) * 4) Then
                    MeasureLine(II).y = y + 3
                End If
            Next II
        End If
    Else
        '  4:-1/-1  3:0/-1   2:+1/-1
        '  5:-1/0   0:0/0    1:+1/0
        '  6:-1/+1  7:0/+1   8:+1/+1
        MeasureLine(1).x = x + 1: MeasureLine(1).y = y + 0
        MeasureLine(2).x = x + 1: MeasureLine(2).y = y - 1
        MeasureLine(3).x = x + 0: MeasureLine(3).y = y - 1
        MeasureLine(4).x = x - 1: MeasureLine(4).y = y - 1
        MeasureLine(5).x = x - 1: MeasureLine(5).y = y + 0
        MeasureLine(6).x = x - 1: MeasureLine(6).y = y + 1
        MeasureLine(7).x = x + 0: MeasureLine(7).y = y + 1
        MeasureLine(8).x = x + 1: MeasureLine(8).y = y + 1
    End If
    
    Dim linecount As Integer
    Dim xx As Integer, yy As Integer
    Dim i As Integer
    
    If RCount > 1 Then
        i = 9
        For linecount = 1 To RCount
            xx = x + 1 + linecount
            For yy = (y + linecount) To (y - linecount) Step -1
                MeasureLine(i).x = xx: MeasureLine(i).y = yy
                i = i + 1
            Next
            
            yy = y - 1 - linecount
            For xx = (x + 1 + linecount) To (x - 1 - linecount) Step -1
                MeasureLine(i).x = xx: MeasureLine(i).y = yy
                i = i + 1
            Next
            
            xx = x - 1 - linecount
            For yy = (y - linecount) To (y + linecount)
                MeasureLine(i).x = xx: MeasureLine(i).y = yy
                i = i + 1
            Next
            
            yy = y + 1 + linecount
            For xx = (x - 1 - linecount) To (x + 1 + linecount)
                MeasureLine(i).x = xx: MeasureLine(i).y = yy
                i = i + 1
            Next
        Next
    End If
    
    For i = 0 To 72
        Shape_Mea(i).Visible = False
    Next
    
    linecount = 8 + IIf(RCount > 1, 8 * RCount, 0)
    
    If YOON_CNT > 0 Then
        linecount = XPitch(TT_NO) * 4 + 4
    Else
        linecount = XPitch(TT_NO) * 4
    End If
    
    For i = 1 To linecount
        xx = MeasureLine(i).x
        yy = MeasureLine(i).y
    
        ' Ä¨ÀÌ ÀÖÀ¸¸é ÃøÁ¤ÇÏ¶ó.
        If Wafer(xx, yy).Chip And LimitAreaRight(xx, yy) Then
            If Not Wafer(xx, yy).ChipMask And _
               Not Wafer(xx, yy).ChipSkipDie And _
               Not Wafer(xx, yy).ChipInk And _
               Not Wafer(xx, yy).ChipMeasure And _
               Not Wafer(xx, yy).MeasureWait And _
               Not Wafer(xx, yy).ChipPlate And _
               Not Wafer(xx, yy).flag Then
               
                MeasureChipCount = MeasureChipCount + 1
               
                MeasureSeq(MeasureChipCount).x = xx
                MeasureSeq(MeasureChipCount).y = yy
                MeasureSeq(MeasureChipCount).r = False
               
'                Shape_Mea(i - 1).Visible = True
                
                Shape_Mea(i - 1).Top = (yy * StarProbe.DisplayChipSizeY) - 1
                Shape_Mea(i - 1).Left = (xx * StarProbe.DisplayChipSizeX) - 1
                
                Wafer(xx, yy).MeasureWait = True
            Else
                MeasureChipCount = MeasureChipCount + 1
               
                MeasureSeq(MeasureChipCount).x = xx
                MeasureSeq(MeasureChipCount).y = yy
                MeasureSeq(MeasureChipCount).r = False
               
'                Shape_Mea(i - 1).Visible = True
                
                Shape_Mea(i - 1).Top = (yy * StarProbe.DisplayChipSizeY) - 1
                Shape_Mea(i - 1).Left = (xx * StarProbe.DisplayChipSizeX) - 1
                
                Wafer(xx, yy).MeasureWait = True
            End If
        Else
            MeasureChipCount = MeasureChipCount + 1             '2016.06.14
            Shape_Mea(i - 1).Top = (yy * StarProbe.DisplayChipSizeY) - 1
            Shape_Mea(i - 1).Left = (xx * StarProbe.DisplayChipSizeX) - 1
            Shape_Mea(i - 1).Visible = True
        End If
    Next
End Sub

Public Sub StarProbe_MeasureLineRight_CH2(x As Integer, y As Integer, Optional RCount As Integer)
    Dim II, jj, kk As Integer
    
    If RCount = 0 Or RCount = Empty Then RCount = 1

    Erase MeasureLine

  
    If XPitch(TT_NO) >= 2 Then
        If YOON_CNT > 0 Then
            MeasureChipCount = 0
            jj = 1
            For II = 1 To (XPitch(TT_NO) * 2)
                For kk = 0 To 1
                    MeasureLine(II + kk).x = x + jj - 1
                    MeasureLine(II + kk).y = y + kk
                    If kk = 1 Then jj = jj + 1
                Next kk
                II = II + 1
            Next II
            
            For II = ((XPitch(TT_NO) * 2) + 1) To (XPitch(TT_NO) * 2 + 2)
                MeasureLine(II).x = x - 1
                If II = (XPitch(TT_NO) * 2) + 1 Then
                    MeasureLine(II).y = y
                ElseIf II = (XPitch(TT_NO) * 2) + 2 Then
                    MeasureLine(II).y = y + 1
'                ElseIf ii = (XPitch(TT_NO) * 4) + 3 Then
'                    MeasureLine(ii).y = y + 2
'                ElseIf ii = (XPitch(TT_NO) * 4) + 4 Then
'                    MeasureLine(ii).y = y + 3
                End If
            Next II
        Else
            MeasureChipCount = 0
            jj = 1
            For II = 1 To (XPitch(TT_NO) * 2) - 2
                For kk = 0 To 1
                    MeasureLine(II + kk).x = x + jj
                    MeasureLine(II + kk).y = y + kk
                    If kk = 1 Then jj = jj + 1
                Next kk
                II = II + 1
            Next II
            
            For II = ((XPitch(TT_NO) * 2) - 1) To (XPitch(TT_NO) * 2)
                MeasureLine(II).x = x - 1
                If II = (XPitch(TT_NO) * 2) - 1 Then
                    MeasureLine(II).y = y
                ElseIf II = (XPitch(TT_NO) * 2) - 0 Then
                    MeasureLine(II).y = y + 1
'                ElseIf ii = (XPitch(TT_NO) * 4) - 1 Then
'                    MeasureLine(ii).y = y + 2
'                ElseIf ii = (XPitch(TT_NO) * 4) Then
'                    MeasureLine(ii).y = y + 3
                End If
            Next II
        End If
    Else
        '  4:-1/-1  3:0/-1   2:+1/-1
        '  5:-1/0   0:0/0    1:+1/0
        '  6:-1/+1  7:0/+1   8:+1/+1
        MeasureLine(1).x = x + 1: MeasureLine(1).y = y + 0
        MeasureLine(2).x = x + 1: MeasureLine(2).y = y - 1
        MeasureLine(3).x = x + 0: MeasureLine(3).y = y - 1
        MeasureLine(4).x = x - 1: MeasureLine(4).y = y - 1
        MeasureLine(5).x = x - 1: MeasureLine(5).y = y + 0
        MeasureLine(6).x = x - 1: MeasureLine(6).y = y + 1
        MeasureLine(7).x = x + 0: MeasureLine(7).y = y + 1
        MeasureLine(8).x = x + 1: MeasureLine(8).y = y + 1
    End If
    
    Dim linecount As Integer
    Dim xx As Integer, yy As Integer
    Dim i As Integer
    
    If RCount > 1 Then
        i = 9
        For linecount = 1 To RCount
            xx = x + 1 + linecount
            For yy = (y + linecount) To (y - linecount) Step -1
                MeasureLine(i).x = xx: MeasureLine(i).y = yy
                i = i + 1
            Next
            
            yy = y - 1 - linecount
            For xx = (x + 1 + linecount) To (x - 1 - linecount) Step -1
                MeasureLine(i).x = xx: MeasureLine(i).y = yy
                i = i + 1
            Next
            
            xx = x - 1 - linecount
            For yy = (y - linecount) To (y + linecount)
                MeasureLine(i).x = xx: MeasureLine(i).y = yy
                i = i + 1
            Next
            
            yy = y + 1 + linecount
            For xx = (x - 1 - linecount) To (x + 1 + linecount)
                MeasureLine(i).x = xx: MeasureLine(i).y = yy
                i = i + 1
            Next
        Next
    End If
    
    For i = 0 To 72
        Shape_Mea(i).Visible = False
    Next
    
    linecount = 8 + IIf(RCount > 1, 8 * RCount, 0)
    
    If YOON_CNT > 0 Then
        linecount = XPitch(TT_NO) * 2 + 2
    Else
        linecount = XPitch(TT_NO) * 2
    End If
    
    For i = 1 To linecount
        xx = MeasureLine(i).x
        yy = MeasureLine(i).y
    
        ' Ä¨ÀÌ ÀÖÀ¸¸é ÃøÁ¤ÇÏ¶ó.
        If Wafer(xx, yy).Chip And LimitAreaRight(xx, yy) Then
            If Not Wafer(xx, yy).ChipMask And _
               Not Wafer(xx, yy).ChipSkipDie And _
               Not Wafer(xx, yy).ChipInk And _
               Not Wafer(xx, yy).ChipMeasure And _
               Not Wafer(xx, yy).MeasureWait And _
               Not Wafer(xx, yy).ChipPlate And _
               Not Wafer(xx, yy).flag Then
               
                MeasureChipCount = MeasureChipCount + 1
               
                MeasureSeq(MeasureChipCount).x = xx
                MeasureSeq(MeasureChipCount).y = yy
                MeasureSeq(MeasureChipCount).r = False
               
                Shape_Mea(i - 1).Visible = True
                
                Shape_Mea(i - 1).Top = (yy * StarProbe.DisplayChipSizeY) - 1
                Shape_Mea(i - 1).Left = (xx * StarProbe.DisplayChipSizeX) - 1
                
                Wafer(xx, yy).MeasureWait = True
            Else
                MeasureChipCount = MeasureChipCount + 1
               
                MeasureSeq(MeasureChipCount).x = xx
                MeasureSeq(MeasureChipCount).y = yy
                MeasureSeq(MeasureChipCount).r = False
               
                Shape_Mea(i - 1).Visible = True
                
                Shape_Mea(i - 1).Top = (yy * StarProbe.DisplayChipSizeY) - 1
                Shape_Mea(i - 1).Left = (xx * StarProbe.DisplayChipSizeX) - 1
                
                Wafer(xx, yy).MeasureWait = True
            End If
        Else
            MeasureChipCount = MeasureChipCount + 1             '2016.06.14
            Shape_Mea(i - 1).Top = (yy * StarProbe.DisplayChipSizeY) - 1
            Shape_Mea(i - 1).Left = (xx * StarProbe.DisplayChipSizeX) - 1
            Shape_Mea(i - 1).Visible = True
        End If
    Next
End Sub

Public Sub StarProbe_MeasureLineRight_CH1(x As Integer, y As Integer, Optional RCount As Integer)
    Dim linecount As Integer
    Dim xx As Integer, yy As Integer
    Dim i As Integer
    Dim II, jj, kk As Integer

    If RCount = 0 Or RCount = Empty Then RCount = 1

    Erase MeasureLine
    
    If RCount = 0 Or RCount = Empty Then RCount = 1

    Erase MeasureLine
        
    MeasureLine(1).x = x + 1: MeasureLine(1).y = y + 0
    MeasureLine(2).x = x + 1: MeasureLine(2).y = y - 1
    MeasureLine(3).x = x + 0: MeasureLine(3).y = y - 1
    MeasureLine(4).x = x - 1: MeasureLine(4).y = y - 1
    MeasureLine(5).x = x - 1: MeasureLine(5).y = y + 0
    MeasureLine(6).x = x - 1: MeasureLine(6).y = y + 1
    MeasureLine(7).x = x + 0: MeasureLine(7).y = y + 1
    MeasureLine(8).x = x + 1: MeasureLine(8).y = y + 1
                   
    If RCount > 1 Then
        i = 9
        For linecount = 1 To RCount
            xx = x + 1 + linecount
            For yy = (y + linecount) To (y - linecount) Step -1
                MeasureLine(i).x = xx: MeasureLine(i).y = yy
                i = i + 1
            Next
            
            yy = y - 1 - linecount
            For xx = (x + 1 + linecount) To (x - 1 - linecount) Step -1
                MeasureLine(i).x = xx: MeasureLine(i).y = yy
                i = i + 1
            Next
            
            xx = x - 1 - linecount
            For yy = (y - linecount) To (y + linecount)
                MeasureLine(i).x = xx: MeasureLine(i).y = yy
                i = i + 1
            Next
            
            yy = y + 1 + linecount
            For xx = (x - 1 - linecount) To (x + 1 + linecount)
                MeasureLine(i).x = xx: MeasureLine(i).y = yy
                i = i + 1
            Next
        Next
    End If
    
    For i = 0 To 127
        Shape_Mea(i).Visible = False
    Next
    
    linecount = 8 + IIf(RCount > 1, 8 * RCount, 0)
    
    For i = 1 To linecount
        xx = MeasureLine(i).x
        yy = MeasureLine(i).y
    
        ' Ä¨ÀÌ ÀÖÀ¸¸é ÃøÁ¤ÇÏ¶ó.
        If Wafer(xx, yy).Chip And LimitAreaRight(xx, yy) Then
            If Not Wafer(xx, yy).ChipMask And Not Wafer(xx, yy).ChipSkipDie And Not Wafer(xx, yy).ChipMeasure And Not Wafer(xx, yy).ChipInk And Not Wafer(xx, yy).MeasureWait And Not Wafer(xx, yy).ChipPlate And Not Wafer(xx, yy).flag Then
                MeasureChipCount = MeasureChipCount + 1
                MeasureSeq(MeasureChipCount).x = xx
                MeasureSeq(MeasureChipCount).y = yy
                MeasureSeq(MeasureChipCount).r = False
                Wafer(xx, yy).MeasureWait = True
            End If
        End If
    Next
End Sub

Public Sub StarProbe_MeasureLineLeft_CH4(x As Integer, y As Integer, Optional RCount As Integer)
    Dim II, jj, kk As Integer
    
    If RCount = 0 Or RCount = Empty Then RCount = 1

    Erase MeasureLine
    
    If XPitch(TT_NO) >= 2 Then
        If YOON_CNT > 0 Then
            MeasureChipCount = 0
            jj = 1
            For II = 1 To (XPitch(TT_NO) * 4)
                For kk = 0 To 3
                    MeasureLine(II + kk).x = x - jj + 1
                    MeasureLine(II + kk).y = y + kk
                    If kk = 3 Then jj = jj + 1
                Next kk
                II = II + 3
            Next II
            
            For II = ((XPitch(TT_NO) * 4) + 1) To (XPitch(TT_NO) * 4 + 4)
                MeasureLine(II).x = x + 1
                If II = (XPitch(TT_NO) * 4) + 1 Then
                    MeasureLine(II).y = y
                ElseIf II = (XPitch(TT_NO) * 4) + 2 Then
                    MeasureLine(II).y = y + 1
                ElseIf II = (XPitch(TT_NO) * 4) + 3 Then
                    MeasureLine(II).y = y + 2
                ElseIf II = (XPitch(TT_NO) * 4) + 4 Then
                    MeasureLine(II).y = y + 3
                End If
            Next II
        Else
            MeasureChipCount = 0
            jj = 1
            For II = 1 To (XPitch(TT_NO) * 4) - 4
                For kk = 0 To 3
                    MeasureLine(II + kk).x = x - jj
                    MeasureLine(II + kk).y = y + kk
                    If kk = 3 Then jj = jj + 1
                Next kk
                II = II + 3
            Next II
            
            For II = ((XPitch(TT_NO) * 4) - 3) To (XPitch(TT_NO) * 4)
                MeasureLine(II).x = x + 1
                If II = (XPitch(TT_NO) * 4) - 3 Then
                    MeasureLine(II).y = y
                ElseIf II = (XPitch(TT_NO) * 4) - 2 Then
                    MeasureLine(II).y = y + 1
                ElseIf II = (XPitch(TT_NO) * 4) - 1 Then
                    MeasureLine(II).y = y + 2
                ElseIf II = (XPitch(TT_NO) * 4) Then
                    MeasureLine(II).y = y + 3
                End If
            Next II
        End If
    Else
        '
        '  2:-1/-1  3:0/-1   4:+1/-1
        '  1:-1/0   0:0/0    5:+1/0
        '  8:-1/+1  7:0/+1   6:+1/+1
        '
        MeasureLine(1).x = x - 1: MeasureLine(1).y = y + 0
        MeasureLine(2).x = x - 1: MeasureLine(2).y = y - 1
        MeasureLine(3).x = x + 0: MeasureLine(3).y = y - 1
        MeasureLine(4).x = x + 1: MeasureLine(4).y = y - 1
        MeasureLine(5).x = x + 1: MeasureLine(5).y = y + 0
        MeasureLine(6).x = x + 1: MeasureLine(6).y = y + 1
        MeasureLine(7).x = x + 0: MeasureLine(7).y = y + 1
        MeasureLine(8).x = x - 1: MeasureLine(8).y = y + 1
    End If
    
    Dim linecount As Integer
    Dim xx As Integer, yy As Integer
    Dim i As Integer
    
    If RCount > 1 Then
        i = 9
        For linecount = 1 To RCount
            xx = x - 1 - linecount
            For yy = (y + linecount) To (y - linecount) Step -1
                MeasureLine(i).x = xx: MeasureLine(i).y = yy
                i = i + 1
            Next
            
            yy = y - 1 - linecount
            For xx = (x - 1 - linecount) To (x + 1 + linecount)
                MeasureLine(i).x = xx: MeasureLine(i).y = yy
                i = i + 1
            Next
            
            xx = x + 1 + linecount
            For yy = (y - linecount) To (y + linecount)
                MeasureLine(i).x = xx: MeasureLine(i).y = yy
                i = i + 1
            Next
            
            yy = y + 1 + linecount
            For xx = (x + 1 + linecount) To (x - 1 - linecount) Step -1
                MeasureLine(i).x = xx: MeasureLine(i).y = yy
                i = i + 1
            Next
        Next
    End If
    
    For i = 0 To 72
        Shape_Mea(i).Visible = False
    Next
    
    If YOON_CNT > 0 Then
        linecount = XPitch(TT_NO) * 4 + 4
    Else
        linecount = XPitch(TT_NO) * 4
    End If
    
    For i = 1 To linecount
        xx = MeasureLine(i).x
        yy = MeasureLine(i).y
    
        ' Ä¨ÀÌ ÀÖÀ¸¸é ÃøÁ¤ÇÏ¶ó.
        If Wafer(xx, yy).Chip And LimitAreaLeft(xx, yy) Then
            If Not Wafer(xx, yy).ChipMask And _
               Not Wafer(xx, yy).ChipSkipDie And _
               Not Wafer(xx, yy).ChipInk And _
               Not Wafer(xx, yy).ChipMeasure And _
               Not Wafer(xx, yy).MeasureWait And _
               Not Wafer(xx, yy).ChipPlate And _
               Not Wafer(xx, yy).flag Then
               
                MeasureChipCount = MeasureChipCount + 1
               
                MeasureSeq(MeasureChipCount).x = xx
                MeasureSeq(MeasureChipCount).y = yy
                MeasureSeq(MeasureChipCount).r = False
               
'                Shape_Mea(i - 1).Visible = True
                
                Shape_Mea(i - 1).Top = (yy * StarProbe.DisplayChipSizeY) - 1
                Shape_Mea(i - 1).Left = (xx * StarProbe.DisplayChipSizeX) - 1

                Wafer(xx, yy).MeasureWait = True
            Else
                MeasureChipCount = MeasureChipCount + 1
               
                MeasureSeq(MeasureChipCount).x = xx
                MeasureSeq(MeasureChipCount).y = yy
                MeasureSeq(MeasureChipCount).r = False
               
'                Shape_Mea(i - 1).Visible = True
                
                Shape_Mea(i - 1).Top = (yy * StarProbe.DisplayChipSizeY) - 1
                Shape_Mea(i - 1).Left = (xx * StarProbe.DisplayChipSizeX) - 1

                Wafer(xx, yy).MeasureWait = True
            End If
        Else
            MeasureChipCount = MeasureChipCount + 1             '2016.06.14
            Shape_Mea(i - 1).Visible = True

            Shape_Mea(i - 1).Top = (yy * StarProbe.DisplayChipSizeY) - 1
            Shape_Mea(i - 1).Left = (xx * StarProbe.DisplayChipSizeX) - 1
        End If
    Next
End Sub

Public Sub StarProbe_MeasureLineLeft_CH2(x As Integer, y As Integer, Optional RCount As Integer)
    Dim II, jj, kk As Integer
    
    If RCount = 0 Or RCount = Empty Then RCount = 1

    Erase MeasureLine
    
    If XPitch(TT_NO) >= 2 Then
        If YOON_CNT > 0 Then
            MeasureChipCount = 0
            jj = 1
            For II = 1 To (XPitch(TT_NO) * 2)
                For kk = 0 To 1
                    MeasureLine(II + kk).x = x - jj + 1
                    MeasureLine(II + kk).y = y + kk
                    If kk = 1 Then jj = jj + 1
                Next kk
                II = II + 1
            Next II
            
            For II = ((XPitch(TT_NO) * 2) + 1) To (XPitch(TT_NO) * 2 + 2)
                MeasureLine(II).x = x + 1
                If II = (XPitch(TT_NO) * 2) + 1 Then
                    MeasureLine(II).y = y
                ElseIf II = (XPitch(TT_NO) * 2) + 2 Then
                    MeasureLine(II).y = y + 1
'                ElseIf ii = (XPitch(TT_NO) * 4) + 3 Then
'                    MeasureLine(ii).y = y + 2
'                ElseIf ii = (XPitch(TT_NO) * 4) + 4 Then
'                    MeasureLine(ii).y = y + 3
                End If
            Next II
        Else
            MeasureChipCount = 0
            jj = 1
            For II = 1 To (XPitch(TT_NO) * 2) - 2
                For kk = 0 To 1
                    MeasureLine(II + kk).x = x - jj
                    MeasureLine(II + kk).y = y + kk
                    If kk = 1 Then jj = jj + 1
                Next kk
                II = II + 1
            Next II
            
            For II = ((XPitch(TT_NO) * 2) - 1) To (XPitch(TT_NO) * 2)
                MeasureLine(II).x = x + 1
                If II = (XPitch(TT_NO) * 2) - 1 Then
                    MeasureLine(II).y = y
                ElseIf II = (XPitch(TT_NO) * 2) - 0 Then
                    MeasureLine(II).y = y + 1
'                ElseIf ii = (XPitch(TT_NO) * 4) - 1 Then
'                    MeasureLine(ii).y = y + 2
'                ElseIf ii = (XPitch(TT_NO) * 4) Then
'                    MeasureLine(ii).y = y + 3
                End If
            Next II
        End If
    Else
        '
        '  2:-1/-1  3:0/-1   4:+1/-1
        '  1:-1/0   0:0/0    5:+1/0
        '  8:-1/+1  7:0/+1   6:+1/+1
        '
        MeasureLine(1).x = x - 1: MeasureLine(1).y = y + 0
        MeasureLine(2).x = x - 1: MeasureLine(2).y = y - 1
        MeasureLine(3).x = x + 0: MeasureLine(3).y = y - 1
        MeasureLine(4).x = x + 1: MeasureLine(4).y = y - 1
        MeasureLine(5).x = x + 1: MeasureLine(5).y = y + 0
        MeasureLine(6).x = x + 1: MeasureLine(6).y = y + 1
        MeasureLine(7).x = x + 0: MeasureLine(7).y = y + 1
        MeasureLine(8).x = x - 1: MeasureLine(8).y = y + 1
    End If
    
    Dim linecount As Integer
    Dim xx As Integer, yy As Integer
    Dim i As Integer
    
    If RCount > 1 Then
        i = 9
        For linecount = 1 To RCount
            xx = x - 1 - linecount
            For yy = (y + linecount) To (y - linecount) Step -1
                MeasureLine(i).x = xx: MeasureLine(i).y = yy
                i = i + 1
            Next
            
            yy = y - 1 - linecount
            For xx = (x - 1 - linecount) To (x + 1 + linecount)
                MeasureLine(i).x = xx: MeasureLine(i).y = yy
                i = i + 1
            Next
            
            xx = x + 1 + linecount
            For yy = (y - linecount) To (y + linecount)
                MeasureLine(i).x = xx: MeasureLine(i).y = yy
                i = i + 1
            Next
            
            yy = y + 1 + linecount
            For xx = (x + 1 + linecount) To (x - 1 - linecount) Step -1
                MeasureLine(i).x = xx: MeasureLine(i).y = yy
                i = i + 1
            Next
        Next
    End If
    
    For i = 0 To 72
        Shape_Mea(i).Visible = False
    Next
    
    If YOON_CNT > 0 Then
        linecount = XPitch(TT_NO) * 2 + 2
    Else
        linecount = XPitch(TT_NO) * 2
    End If
    
    For i = 1 To linecount
        xx = MeasureLine(i).x
        yy = MeasureLine(i).y
    
        ' Ä¨ÀÌ ÀÖÀ¸¸é ÃøÁ¤ÇÏ¶ó.
        If Wafer(xx, yy).Chip And LimitAreaLeft(xx, yy) Then
            If Not Wafer(xx, yy).ChipMask And _
               Not Wafer(xx, yy).ChipSkipDie And _
               Not Wafer(xx, yy).ChipInk And _
               Not Wafer(xx, yy).ChipMeasure And _
               Not Wafer(xx, yy).MeasureWait And _
               Not Wafer(xx, yy).ChipPlate And _
               Not Wafer(xx, yy).flag Then
               
                MeasureChipCount = MeasureChipCount + 1
               
                MeasureSeq(MeasureChipCount).x = xx
                MeasureSeq(MeasureChipCount).y = yy
                MeasureSeq(MeasureChipCount).r = False
               
                Shape_Mea(i - 1).Visible = True
                
                Shape_Mea(i - 1).Top = (yy * StarProbe.DisplayChipSizeY) - 1
                Shape_Mea(i - 1).Left = (xx * StarProbe.DisplayChipSizeX) - 1

                Wafer(xx, yy).MeasureWait = True
            Else
                MeasureChipCount = MeasureChipCount + 1
               
                MeasureSeq(MeasureChipCount).x = xx
                MeasureSeq(MeasureChipCount).y = yy
                MeasureSeq(MeasureChipCount).r = False
               
                Shape_Mea(i - 1).Visible = True
                
                Shape_Mea(i - 1).Top = (yy * StarProbe.DisplayChipSizeY) - 1
                Shape_Mea(i - 1).Left = (xx * StarProbe.DisplayChipSizeX) - 1

                Wafer(xx, yy).MeasureWait = True
            End If
        Else
            MeasureChipCount = MeasureChipCount + 1             '2016.06.14
            Shape_Mea(i - 1).Visible = True

            Shape_Mea(i - 1).Top = (yy * StarProbe.DisplayChipSizeY) - 1
            Shape_Mea(i - 1).Left = (xx * StarProbe.DisplayChipSizeX) - 1
        End If
    Next
End Sub

Public Sub StarProbe_MeasureLineLeft_CH1(x As Integer, y As Integer, Optional RCount As Integer)
    Dim linecount As Integer
    Dim xx As Integer, yy As Integer
    Dim i As Integer
    Dim II, jj, kk As Integer

    If RCount = 0 Or RCount = Empty Then RCount = 1

    Erase MeasureLine
            
    If RCount = 0 Or RCount = Empty Then RCount = 1

    Erase MeasureLine

    MeasureLine(1).x = x - 1: MeasureLine(1).y = y + 0
    MeasureLine(2).x = x - 1: MeasureLine(2).y = y - 1
    MeasureLine(3).x = x + 0: MeasureLine(3).y = y - 1
    MeasureLine(4).x = x + 1: MeasureLine(4).y = y - 1
    MeasureLine(5).x = x + 1: MeasureLine(5).y = y + 0
    MeasureLine(6).x = x + 1: MeasureLine(6).y = y + 1
    MeasureLine(7).x = x + 0: MeasureLine(7).y = y + 1
    MeasureLine(8).x = x - 1: MeasureLine(8).y = y + 1
    
    If RCount > 1 Then
        i = 9
        For linecount = 1 To RCount
            xx = x - 1 - linecount
            For yy = (y + linecount) To (y - linecount) Step -1
                MeasureLine(i).x = xx: MeasureLine(i).y = yy
                i = i + 1
            Next
            
            yy = y - 1 - linecount
            For xx = (x - 1 - linecount) To (x + 1 + linecount)
                MeasureLine(i).x = xx: MeasureLine(i).y = yy
                i = i + 1
            Next
            
            xx = x + 1 + linecount
            For yy = (y - linecount) To (y + linecount)
                MeasureLine(i).x = xx: MeasureLine(i).y = yy
                i = i + 1
            Next
            
            yy = y + 1 + linecount
            For xx = (x + 1 + linecount) To (x - 1 - linecount) Step -1
                MeasureLine(i).x = xx: MeasureLine(i).y = yy
                i = i + 1
            Next
        Next
    End If
    
    For i = 0 To 127
        Shape_Mea(i).Visible = False
    Next
    
    linecount = 8 + IIf(RCount > 1, 8 * RCount, 0)
    
    For i = 1 To linecount
        xx = MeasureLine(i).x
        yy = MeasureLine(i).y
    
        ' Ä¨ÀÌ ÀÖÀ¸¸é ÃøÁ¤ÇÏ¶ó.
        If Wafer(xx, yy).Chip And LimitAreaLeft(xx, yy) Then
            If Not Wafer(xx, yy).ChipMask And Not Wafer(xx, yy).ChipSkipDie And Not Wafer(xx, yy).ChipMeasure And Not Wafer(xx, yy).MeasureWait And Not Wafer(xx, yy).ChipInk And Not Wafer(xx, yy).ChipPlate And Not Wafer(xx, yy).flag Then
                MeasureChipCount = MeasureChipCount + 1
                MeasureSeq(MeasureChipCount).x = xx
                MeasureSeq(MeasureChipCount).y = yy
                MeasureSeq(MeasureChipCount).r = False
                Wafer(xx, yy).MeasureWait = True
            End If
        End If
    Next
End Sub

Sub StarProbe_Auto_Test_CH1()
    Dim XYpos As String
    Dim xx As Integer, yy As Integer
    Dim forx As Integer, fory As Integer
    Dim bRight As Boolean
    Dim bCon As Boolean
    
    Dim File_Name, SP_File_Name As String

    Dim bStart As Boolean
    Dim val, sval As String
    Dim sOldFileName As String
    
    Dim FIRST_CHK As Boolean
    
    Dim Y_change As Boolean                 '[ 2022.05.31 ]
  
    bStarProbe_Auto_Start = True
    bRight = True
    ErrorStop = False
      
    xx = StarProbe.StartChip.x
    Start_X = xx
    yy = StarProbe.StartChip.y
      
    Check1(3).Enabled = True
    bStart = True
          
    If Stop_Measure Then
        xx = Stop_xx
        yy = Stop_yy
    End If
    
    FIRST_CHK = True
    Y_change = False                    '[ 2022.05.31 ]
          
    Do While bStarProbe_Auto_Start
        DoEvents
        If bPause_Flag = True Then     'pause & coutinue flag
            Call StarProbe_Pause
            Call StarProbe_Z_Down
            If DemoMode = 0 Then
                If ACCO = 1 Or Model_Select = 1 Then                            'ACCO only
                    MSComm1.Output = "PA" & vbCrLf
                Else
                    MSComm1.Output = "PA" & vbLf
                End If
            End If
            
        '[ 2022.07.29 ] : yÃà±âÁØ 1/2 º¸´Ù Å©¸é¼­ Ä§ÀûÈ®ÀÎ ¼³Á¤ÀÌ µÈ WaferÀÎ °æ¿ì Ä§ÀûÈ®ÀÎ ¸Þ½ÃÁö ³ªÅ¸³½´Ù.
        ElseIf yy > Int(StarProbe.ChipCountY / 2) And Needle_Chk_Ok = False And Needle_Chk(lblWafer.Caption - 1) = True Then
            Command_Stop_Click                          'stop
            Form_Check_List.Show 1                      'check list on
            Needle_Chk_Ok = True
            If Needle_check_flag = False Then           'Ä§Àû ¹þ¾î³² (ºÒ·®Ã³¸®)°¡ ¾Æ´Ñ °æ¿ì
                Check1(3).value = 1                     'auto µ¿ÀÛÀ» ´Ù½Ã ½ÃÀÛÇÑ´Ù.
            End If
        ElseIf StarProbe.Tip_Clean = 1 And StarProbe.Tipclean_Count >= StarProbe.Tipclean_Count_Limit Then
            '[ 2022.07.20 ]
            If LOG_FILE_ON = 1 Then SelectExt.Log_Data_Save (13)
            If DemoMode = 0 Then
                Call StarProbe_Pause
                Call StarProbe_Z_Down
                If ACCO = 1 Or Model_Select = 1 Then                            'ACCO only
                    MSComm1.Output = "PA" & vbCrLf
                Else
                    MSComm1.Output = "PA" & vbLf
                End If
                Sleep 1000

                Call StarProbe_tip_clean

                Sleep 5000
            End If

            StarProbe.Tipclean_Count = 0
            Text7.Refresh
            Text7.Text = StarProbe.Tipclean_Count

            If DemoMode = 0 Then
                Call StarProbe_Pause
                If ACCO = 1 Or Model_Select = 1 Then                            'ACCO only
                    MSComm1.Output = "CO" & vbCrLf
                Else
                    MSComm1.Output = "CO" & vbLf
                End If

'                Call StarProbe_XY_Moving((XAxis), (YAxis))
'
'                Sleep 100
'
'                If StarProbe_Motor_End_check Then
'                    MsgBox "Motor not end check !", 16, "STAR PROBE"
'                End If
                Call StarProbe_Z_UP
            End If
            Sleep 1000
        Else
            If Stop_Measure Then
                bStart = False
                                          
                XYpos = " "
                                     
                If ErrorStop = True Then Exit Do   'Ãß°¡
                If Test_Ready = False And TimerCheck = True Then Exit Do   '2003.09.16 coupling
                Call StarProbe_Measure_CH1(Text5, Text6, bRight, False)
            Else
                If Wafer(xx, yy).Chip And _
                    Not Wafer(xx, yy).ChipMask And _
                    Not Wafer(xx, yy).ChipSkipDie And _
                    Not Wafer(xx, yy).ChipPlate Then
                                           
                    If Not Wafer(xx, yy).flag Then
'                        If StarProbe.MeasureAll = 1 Or (XPitch(TT_NO) = 1 And YPitch(TT_NO) = 1) Then
'                        Else
'                            If Y_change = True Then             '[ 2022.05.31 ] : yÃàÀÌ ¹Ù²ï °æ¿ì ¹ÙÁ®³ª°£´Ù.
'                                Y_change = False
'                                If bRight = True Then
'                                    xx = xx
'                                Else
'                                    xx = xx + 1
'                                End If
'                                'Exit For
'                            End If
'                        End If
                        bStart = False
                                                      
                        XYpos = " "
                                             
                        If ErrorStop = True Then Exit Do   'Ãß°¡
                                    
                        If Test_Ready = False And TimerCheck = True Then Exit Do   '2003.09.16 coupling
                        ''''''''''''''''''''''''''''''''''[2020.03.17] : ½ÇÁ¦ÃøÁ¤ÁÂÇ¥ÀúÀå
                        Wafer(xx, yy).Real_Chk = True
                        ''''''''''''''''''''''''''''''''''
                        VScroll_Zoom.value = (yy / (Abs(StarProbe.Max.y) + Abs(StarProbe.Min.y) + 1)) * 1000
                        HScroll_Zoom.value = (xx / (Abs(StarProbe.Max.x) + Abs(StarProbe.Min.x) + 1)) * 1000
                        
                        StarProbe.CurrentChip.x = xx - StarProbe.StartChip.x
                        StarProbe.CurrentChip.y = yy - StarProbe.StartChip.y
                        
                        Text5 = StarProbe.CurrentChip.x
                        Text6 = StarProbe.CurrentChip.y
                        
                        Shape_Chip.Top = (yy * StarProbe.DisplayChipSizeY) - 2
                        Shape_Chip.Left = (xx * StarProbe.DisplayChipSizeX) - 2
                        Shape_Chip.Height = (StarProbe.DisplayChipSizeY) + 4
                        
                        Label_ChipPosition = StarProbe.CurrentChip.x & "/" & StarProbe.CurrentChip.y
                        Call StarProbe_Measure_CH1(Text5, Text6, bRight, False)
                        Call Display_Chip(pZoom, pOriginal, Text5, Text6)
                        
                        ''
                        SSPanel2(1).Caption = Test_Cnt
                        Text_TotalCount.Text = Test_Cnt
                        
                        If Test_Cnt = 0 Then
                            SSPanel2(2).Caption = Good_Cnt & Space(1) & "(" & "0.00" & "%)"
                        Else
                            SSPanel2(2).Caption = Good_Cnt & Space(1) & "(" & Format(Good_Cnt / Test_Cnt * 100, "0.00") & "%)"
                        End If
                        Text_GoodCount.Text = Good_Cnt
                        ''
                    End If
                End If
            End If
                      
            If Stop_Measure Then
                Stop_xx = xx
                Stop_yy = yy
            End If
            
            If bRight Then
                If StarProbe.MeasureAll = 1 Or (XPitch(TT_NO) = 1 And YPitch(TT_NO) = 1) Then
                    xx = xx + 1
                Else
                    xx = xx + XPitch(TT_NO)
                End If
                
                If xx > StarProbe.ChipCountX Then
                    Y_change = True             '[ 2022.05.31 ]
                    If StarProbe.MeasureAll = 1 Or (XPitch(TT_NO) = 1 And YPitch(TT_NO) = 1) Then
                        yy = yy + 1
                    Else
                        yy = yy + YPitch(TT_NO)
                    End If
                    
                    xx = StarProbe.ChipCountX
                    
                    For forx = StarProbe.ChipCountX To 0 Step -1
                        If Wafer(forx, yy).Chip And Not Wafer(forx, yy).ChipMask And Not Wafer(forx, yy).ChipSkipDie And Not Wafer(forx, yy).ChipPlate Then
                            xx = forx
                            Exit For
                        End If
                    Next
                    bRight = False
                End If
            Else
                If StarProbe.MeasureAll = 1 Or (XPitch(TT_NO) = 1 And YPitch(TT_NO) = 1) Then
                    xx = xx - 1
                Else
                    xx = xx - XPitch(TT_NO)
                End If
                
                If xx < 0 Then
                    Y_change = True             '[ 2022.05.31 ]
                    If StarProbe.MeasureAll = 1 Or (XPitch(TT_NO) = 1 And YPitch(TT_NO) = 1) Then
                        yy = yy + 1
                    Else
                        yy = yy + YPitch(TT_NO)
                    End If
                    
                    xx = 0
                    
                    For forx = 0 To StarProbe.ChipCountX
                        If Wafer(forx, yy).Chip And Not Wafer(forx, yy).ChipMask And Not Wafer(forx, yy).ChipSkipDie And Not Wafer(forx, yy).ChipPlate Then
                            xx = forx
                            Exit For
                        End If
                    Next
                    bRight = True
                End If
            End If

            '[ ÆÄÀÏ ÀúÀåÇÏ´Â ºÎºÐ ]
            If yy > StarProbe.ChipCountY Then
                Needle_Chk_Ok = False                                   '[ 2022.07.29 ]
                ''''''''''''''''''''''[demo]'''''''''''''''''''''
'                Call Command_Map_Clear_Click
'                Call Command_DisplayWafer_Click
'                XX = StarProbe.StartChip.x
'                yy = StarProbe.StartChip.y
'                bRight = True
                ''''''''''''''''''''''[demo]'''''''''''''''''''''
                '=====================================================================================================================
                If (TT_NO + 1) <= 9 Then                                        '[ 2020.02.07 ] : 1~9±îÁö´Â 01~09·Î Ç¥½ÃÇÑ´Ù.
                    W_NO = "0" & TT_NO + 1
                Else                                                            '10ÀÌ»óÀº ±×´ë·Î »ç¿ëÇÏ¸é µÈ´Ù.
                    W_NO = TT_NO + 1
                End If

                If DemoMode = 0 Then
                    If SaveDrive = 0 Then
                        File_Name = "C:\data\" & UCase(Text1(0).Text) & "\" & UCase(Text1(0).Text) & "_" & W_NO        'HDD data save path
                    Else
                        File_Name = "D:\data\" & UCase(Text1(0).Text) & "\" & UCase(Text1(0).Text) & "_" & W_NO        'HDD data save path
                    End If
                Else
                    If SaveDrive = 0 Then
                        File_Name = "C:\data\" & UCase(Text1(0).Text) & "\" & UCase(Text1(0).Text) & "_" & W_NO        'HDD data save path
                    Else
                        File_Name = "D:\data\" & UCase(Text1(0).Text) & "\" & UCase(Text1(0).Text) & "_" & W_NO        'HDD data save path
                    End If
                End If

                If No_Probe = True Then                                         '[ edge ink ]
                Else                                                            '[ normal ]
                    For i = 0 To StarProbe.ChipCountX                           '2015.12.04 : ink ÈçÀûÁ¦°Å(spÆÄÀÏÀ» ºÒ·¯¿Í¼­ ´Ù½Ã ÃøÁ¤ÇÒ ¼ö ÀÖ±âÀ§ÇØ¼­)
                        For j = 0 To StarProbe.ChipCountY
                            Wafer(i, j).InkDot = 0
                        Next j
                    Next i

                    FILE_NAMEING = File_Name
                    Form_StarProbe_MeasureDataSave.Display_View5                'wmd01 ÆÄÀÏ ÀúÀå (server, hdd)
                    Form_StarProbe_MeasureDataSave.Display_View5_1              'map2 [ 2022.09.29 ] Ãß°¡
                    Form_StarProbe_MeasureDataSave.Display_View_Change          'txt ÆÄÀÏ ÀúÀå (server, hdd)
                End If

                Call StarProbe_FileSave_Data(File_Name & ".SP")                 'sp ÆÄÀÏ ÀúÀå (hdd)

                BMP_file = File_Name & ".PNG"                                   '[ 2021.05.11 ] : BMP->PNG
                Call Form_StarProbe_MeasureDataSave.Display_View                '±×¸² ÆÄÀÏ ÀúÀå (server, hdd)

                If AutoAlign_Flag = False Then                                  'auto alignÀÌ ¾Æ´Ñ °æ¿ì
                    Z_HEIGHT.Command2.Visible = IIf(StarProbe.Ink_After = 2, True, False)
                    Z_HEIGHT.Show 1
                End If
                '=====================================================================================================================

                If StarProbe.Ink_After = 2 Then                                 'ink off
                    If AutoAlign_Flag = True Then                               '[ 2021.04.15 ] : autoÀÎ °æ¿ì ink offÇÏ¸é ink¾øÀÌ ¿¬¼Óµ¿ÀÛÇÏµµ·Ï ¼öÁ¤.
                        Call StarProbe_After_Ink_Dot_noink
                    Else                                                        'auto alignÀÌ ¾Æ´Ñ °æ¿ì
                        If INK_OFF_TEST = False Then
                            RESET_DATA
                            Check1(3).value = 0
                            bStop = True
                            bStarprobe_AfterInk = False
                            Call StarProbe_Zero_point       'Ãß°¡
                        Else
                            Call StarProbe_After_Ink_Dot
                        End If
                    End If
                Else                                                            'ink direct or after
                    Call StarProbe_After_Ink_Dot
                End If

                bStarProbe_Auto_Start = False                                   'star probe auto test off
                TESTING_flag = False                '2016.03.11

                'Wafer End½ÅÈ£¸¦ ¹Þ¾ÒÀ¸¹Ç·Î º¯¼ö¸¦ ÃÊ±âÈ­ ÇØÁØ´Ù.
                STT_time = ""
                END_time = ""

                '[ 2021.03.09 ] : lot endÀÎ °æ¿ì º¯¼ö Å¬¸®¾î(Áßº¹ÀúÀåÀ» ¸·±âÀ§ÇØ¼­ Ãß°¡)
                If Slot_Max_Count = Int(W_NO) Then
                    For i = 0 To 24
                        NOW_NO(i) = True
                    Next i
                End If
            End If
        End If
        If bStarProbe_Auto_Start = False Or bStop Then Exit Do
    Loop
    
    SSPanel_BadCount.Caption = StarProbe.CountBadDie
    Text_BadCount.Text = StarProbe.CountBadDie
    SSPanel_GoodCount.Caption = StarProbe.CountGoodDie
    SSPanel2(1).Caption = Test_Cnt
    Text_TotalCount.Text = Test_Cnt
    
    If Test_Cnt = 0 Then
        SSPanel2(2).Caption = Good_Cnt & Space(1) & "(" & "0.00" & "%)"
    Else
        SSPanel2(2).Caption = Good_Cnt & Space(1) & "(" & Format(Good_Cnt / Test_Cnt * 100, "0.00") & "%)"
    End If
    Text_GoodCount.Text = Good_Cnt
End Sub


Sub StarProbe_Auto_Test_CH2()
    Dim XYpos As String
    Dim xx As Integer, yy As Integer
    Dim forx As Integer, fory As Integer
    Dim bRight As Boolean
    Dim bCon As Boolean
    
    Dim File_Name, SP_File_Name As String

    Dim bStart As Boolean
    Dim val, sval As String
    Dim sOldFileName As String
    
    Dim FIRST_CHK As Boolean
    
    Dim Y_change As Boolean             '[ 2022.05.31 ]
  
    bStarProbe_Auto_Start = True
    bRight = True
    ErrorStop = False
      
    xx = StarProbe.StartChip.x
    Start_X = xx
    yy = StarProbe.StartChip.y
      
    Check1(3).Enabled = True
    bStart = True
          
    If Stop_Measure Then
        xx = Stop_xx
        yy = Stop_yy
    End If
    
    FIRST_CHK = True
    Y_change = False                '[ 2022.05.31 ]
          
    Do While bStarProbe_Auto_Start
        DoEvents
        If bPause_Flag = True Then     'pause & coutinue flag
            Call StarProbe_Pause
            Call StarProbe_Z_Down
            If DemoMode = 0 Then
                If ACCO = 1 Or Model_Select = 1 Then                            'ACCO only
                    MSComm1.Output = "PA" & vbCrLf
                Else
                    MSComm1.Output = "PA" & vbLf
                End If
            End If
        '[ 2022.07.29 ] : yÃà±âÁØ 1/2 º¸´Ù Å©¸é¼­ Ä§ÀûÈ®ÀÎ ¼³Á¤ÀÌ µÈ WaferÀÎ °æ¿ì Ä§ÀûÈ®ÀÎ ¸Þ½ÃÁö ³ªÅ¸³½´Ù.
        ElseIf yy > Int(StarProbe.ChipCountY / 2) And Needle_Chk_Ok = False And Needle_Chk(lblWafer.Caption - 1) = True Then
'            MsgBox "Ä§ÀûÀ» È®ÀÎÇÏ¼¼¿ä."
            Command_Stop_Click
            Form_Check_List.Show 1
            Needle_Chk_Ok = True
            If Needle_check_flag = False Then
                Check1(3).value = 1
            End If
        ElseIf StarProbe.Tip_Clean = 1 And StarProbe.Tipclean_Count >= StarProbe.Tipclean_Count_Limit Then
            '[ 2022.07.20 ]
            If LOG_FILE_ON = 1 Then SelectExt.Log_Data_Save (13)
            If DemoMode = 0 Then
                Call StarProbe_Pause
                Call StarProbe_Z_Down
                If ACCO = 1 Or Model_Select = 1 Then                            'ACCO only
                    MSComm1.Output = "PA" & vbCrLf
                Else
                    MSComm1.Output = "PA" & vbLf
                End If
                Sleep 1000

                Call StarProbe_tip_clean

                Sleep 5000
            End If

            StarProbe.Tipclean_Count = 0
            Text7.Refresh
            Text7.Text = StarProbe.Tipclean_Count

            If DemoMode = 0 Then
                Call StarProbe_Pause
                If ACCO = 1 Or Model_Select = 1 Then                            'ACCO only
                    MSComm1.Output = "CO" & vbCrLf
                Else
                    MSComm1.Output = "CO" & vbLf
                End If

'                Call StarProbe_XY_Moving((XAxis), (YAxis))
'
'                Sleep 100
'
'                If StarProbe_Motor_End_check Then
'                    MsgBox "Motor not end check !", 16, "STAR PROBE"
'                End If
                Call StarProbe_Z_UP
            End If
            Sleep 1000
        Else
            If Stop_Measure Then
                bStart = False
                                          
                XYpos = " "
                                     
                If ErrorStop = True Then Exit Do   'Ãß°¡
                If Test_Ready = False And TimerCheck = True Then Exit Do   '2003.09.16 coupling
                Call StarProbe_Measure_CH2(Text5, Text6, bRight, False)
            Else
                For ZZZ = 0 To 1
                    If Wafer(xx, yy + ZZZ).Chip And _
                        Not Wafer(xx, yy + ZZZ).ChipMask And _
                        Not Wafer(xx, yy + ZZZ).ChipSkipDie And _
                        Not Wafer(xx, yy + ZZZ).ChipPlate Then
                                               
                        If Not Wafer(xx, yy + ZZZ).flag Then
'                            If StarProbe.MeasureAll = 1 Or (XPitch(TT_NO) = 1 And YPitch(TT_NO) = 1) Then
'                            Else
'                                If Y_change = True Then             '[ 2022.05.31 ] : yÃàÀÌ ¹Ù²ï °æ¿ì ¹ÙÁ®³ª°£´Ù.
'                                    Y_change = False
'                                    If bRight = True Then
'                                        xx = xx
'                                    Else
'                                        xx = xx + 1
'                                    End If
'                                    'Exit For
'                                End If
'                            End If
                            bStart = False
                                                      
                            XYpos = " "
                                                 
                            If ErrorStop = True Then Exit Do   'Ãß°¡
                                        
                            If Test_Ready = False And TimerCheck = True Then Exit Do   '2003.09.16 coupling
                            ''''''''''''''''''''''''''''''''''[2020.03.17] : ½ÇÁ¦ÃøÁ¤ÁÂÇ¥ÀúÀå
                            Wafer(xx, yy + ZZZ).Real_Chk = True
                            ''''''''''''''''''''''''''''''''''
                            VScroll_Zoom.value = (yy / (Abs(StarProbe.Max.y) + Abs(StarProbe.Min.y) + 1)) * 1000
                            HScroll_Zoom.value = (xx / (Abs(StarProbe.Max.x) + Abs(StarProbe.Min.x) + 1)) * 1000
                            
                            StarProbe.CurrentChip.x = xx - StarProbe.StartChip.x
                            StarProbe.CurrentChip.y = yy - StarProbe.StartChip.y
                            
                            Text5 = StarProbe.CurrentChip.x
                            Text6 = StarProbe.CurrentChip.y
                            
                            Shape_Chip.Top = (yy * StarProbe.DisplayChipSizeY) - 2
                            Shape_Chip.Left = (xx * StarProbe.DisplayChipSizeX) - 2
                            Shape_Chip.Height = (StarProbe.DisplayChipSizeY * 2) + 4
                            
                            Label_ChipPosition = StarProbe.CurrentChip.x & "/" & StarProbe.CurrentChip.y
                            Call StarProbe_Measure_CH2(Text5, Text6, bRight, False)
                            Call Display_Chip(pZoom, pOriginal, Text5, Text6)
                            
                            ''
                            SSPanel2(1).Caption = Test_Cnt
                            Text_TotalCount.Text = Test_Cnt
                            
                            If Test_Cnt = 0 Then
                                SSPanel2(2).Caption = Good_Cnt & Space(1) & "(" & "0.00" & "%)"
                            Else
                                SSPanel2(2).Caption = Good_Cnt & Space(1) & "(" & Format(Good_Cnt / Test_Cnt * 100, "0.00") & "%)"
                            End If
                            Text_GoodCount.Text = Good_Cnt
                            ''
                        End If
                    End If
                Next ZZZ
            End If
                      
            If Stop_Measure Then
                Stop_xx = xx
                Stop_yy = yy
            End If
            
            If bRight Then
                If StarProbe.MeasureAll = 1 Or (XPitch(TT_NO) = 1 And YPitch(TT_NO) = 1) Then         'Àü¼ö°Ë»ç x+1
                    xx = xx + 1
                Else
                    xx = xx + (XPitch(TT_NO))
                End If
                
                If StarProbe.MeasureAll = 1 Or (XPitch(TT_NO) = 1 And YPitch(TT_NO) = 1) Then         'Àü¼ö°Ë»ç x+1
                    If xx > StarProbe.ChipCountX And FIRST_CHK = True Then     'ÇöÀç xÁÂÇ¥°¡ ¿þ¾îÆÛÀÇ xÃà ÁÂÇ¥º¸´Ù Å«°æ¿ì(¿ìÃø ³¡±îÁö ÀÌµ¿ÇÑ°æ¿ì)
                        FIRST_CHK = False
                        xx = StarProbe.ChipCountX           '¿ìÃø ³¡À¸·Î ÀÌµ¿
                        
                        For forx = StarProbe.ChipCountX To 0 Step -1
                            If Wafer(forx, yy).Chip And _
                               Not Wafer(forx, yy).ChipMask And _
                               Not Wafer(forx, yy).ChipSkipDie And _
                               Not Wafer(forx, yy).ChipPlate Then           'Á¤»ó Ä¨ÀÎ °æ¿ì
                                xx = StarProbe.ChipCountX                   'xx¿¡ °ªÁöÁ¤
                                Exit For
                            End If
                        Next
                        bRight = False                      'ÁÂÃøÀÌµ¿ ¼³Á¤
                    End If
                Else                                        'Àü¼ö°Ë»ç + step mode
                    If xx > StarProbe.ChipCountX And FIRST_CHK = True Then      'ÇöÀç xÁÂÇ¥°¡ ¿þ¾îÆÛÀÇ xÃà ÁÂÇ¥º¸´Ù Å«°æ¿ì(¿ìÃø ³¡±îÁö ÀÌµ¿ÇÑ°æ¿ì)
                        FIRST_CHK = False
                        xx = StarProbe.ChipCountX           '¿ìÃø ³¡À¸·Î ÀÌµ¿
                        
                        For forx = StarProbe.ChipCountX To 0 Step -1
                            For ZZZ = 0 To 1
'                                If Wafer(forx, yy + ZZZ).Chip And _
'                                   Not Wafer(forx, yy + ZZZ).ChipMask And _
'                                   Not Wafer(forx, yy + ZZZ).ChipSkipDie And _
'                                   Not Wafer(forx, yy + ZZZ).ChipPlate Then         'Á¤»ó Ä¨ÀÎ °æ¿ì
'                                    XX = Start_X                   'xx ½ÃÀÛÁöÁ¡À¸·Î ÀÌµ¿
'    '                                Exit For
'                                Else
                                    If Wafer(forx, yy + ZZZ).flag = True Then       '¿ìÃø¿¡¼­ ÁÂÃøÀ¸·Î ÀÌµ¿ÇÏ¸é¼­ Ã³À½ ÃøÁ¤ÇÑ ÁöÁ¡À» Ã£´Â ÄÚµå
                                        xx = Start_X
                                    End If
'                                End If
                            Next ZZZ
                        Next
                        bRight = False                      'ÁÂÃøÀÌµ¿ ¼³Á¤
                    End If
                End If
                
                If xx > StarProbe.ChipCountX Then      'ÇöÀç xÁÂÇ¥°¡ ¿þ¾îÆÛÀÇ xÃà ÁÂÇ¥º¸´Ù Å«°æ¿ì(¿ìÃø ³¡±îÁö ÀÌµ¿ÇÑ°æ¿ì)
'                    yy = yy + 4                     '2015.09.14
                    yy = yy + (YPitch(TT_NO) * 2)      '[ 2018.01.29 ]
                    Y_change = True                     '[ 2022.05.31 ]
                    If yy < StarProbe.ChipCountY Then   '2015.09.14
                        VScroll_Zoom.value = (yy / (Abs(StarProbe.Max.y) + Abs(StarProbe.Min.y) + 1)) * 1000
                    End If
                    xx = StarProbe.ChipCountX           '¿ìÃø ³¡À¸·Î ÀÌµ¿
                    
                    For forx = StarProbe.ChipCountX To 0 Step -1
                        If Wafer(forx, yy).Chip And _
                           Not Wafer(forx, yy).ChipMask And _
                           Not Wafer(forx, yy).ChipSkipDie And _
                           Not Wafer(forx, yy).ChipPlate Then           'Á¤»ó Ä¨ÀÎ °æ¿ì
                            'xx = forx                   'xx¿¡ °ªÁöÁ¤
                            xx = StarProbe.ChipCountX                   'xx¿¡ °ªÁöÁ¤
                            Exit For
                        End If
                    Next
                    bRight = False                      'ÁÂÃøÀÌµ¿ ¼³Á¤
                End If
            Else
                If StarProbe.MeasureAll = 1 Or (XPitch(TT_NO) = 1 And YPitch(TT_NO) = 1) Then         'Àü¼ö°Ë»ç x+1
                    xx = xx - 1
                Else                                    'step mode °Ë»ç
                    xx = xx - (XPitch(TT_NO))
                End If
                If xx < 0 Then
'                    yy = yy + 4                     '2015.09.14
                    yy = yy + (YPitch(TT_NO) * 2)      '[ 2018.01.29 ]
                    Y_change = True                         '[ 2022.05.31 ]
                    If yy < StarProbe.ChipCountY Then   '2015.09.14
                        VScroll_Zoom.value = (yy / (Abs(StarProbe.Max.y) + Abs(StarProbe.Min.y) + 1)) * 1000
                    End If

                    xx = 0
                    
                    For forx = 0 To StarProbe.ChipCountX
                        If Wafer(forx, yy).Chip And _
                            Not Wafer(forx, yy).ChipMask And _
                            Not Wafer(forx, yy).ChipSkipDie And _
                            Not Wafer(forx, yy).ChipPlate Then   'Á¤»ó Ä¨ÀÎ °æ¿ì
                            xx = 0                       '2015.09.14
                            Exit For
                        End If
                    Next
                    bRight = True                       '¿ìÃøÀÌµ¿ ¼³Á¤
                End If
            End If

            '[ ÆÄÀÏ ÀúÀåÇÏ´Â ºÎºÐ ]
            If yy > StarProbe.ChipCountY Then
                Needle_Chk_Ok = False
                ''''''''''''''''''''''[demo]'''''''''''''''''''''
'                Call Command_Map_Clear_Click
'                Call Command_DisplayWafer_Click
'                XX = StarProbe.StartChip.x
'                yy = StarProbe.StartChip.y
'                bRight = True
                ''''''''''''''''''''''[demo]'''''''''''''''''''''
                '=====================================================================================================================
                If (TT_NO + 1) <= 9 Then                                        '[ 2020.02.07 ] : 1~9±îÁö´Â 01~09·Î Ç¥½ÃÇÑ´Ù.
                    W_NO = "0" & TT_NO + 1
                Else                                                            '10ÀÌ»óÀº ±×´ë·Î »ç¿ëÇÏ¸é µÈ´Ù.
                    W_NO = TT_NO + 1
                End If

                If DemoMode = 0 Then
                    If SaveDrive = 0 Then
                        File_Name = "C:\data\" & UCase(Text1(0).Text) & "\" & UCase(Text1(0).Text) & "_" & W_NO        'HDD data save path
                    Else
                        File_Name = "D:\data\" & UCase(Text1(0).Text) & "\" & UCase(Text1(0).Text) & "_" & W_NO        'HDD data save path
                    End If
                Else
                    If SaveDrive = 0 Then
                        File_Name = "C:\data\" & UCase(Text1(0).Text) & "\" & UCase(Text1(0).Text) & "_" & W_NO        'HDD data save path
                    Else
                        File_Name = "D:\data\" & UCase(Text1(0).Text) & "\" & UCase(Text1(0).Text) & "_" & W_NO        'HDD data save path
                    End If
                End If

                If No_Probe = True Then                                         '[ edge ink ]
                Else                                                            '[ normal ]
                    For i = 0 To StarProbe.ChipCountX                           '2015.12.04 : ink ÈçÀûÁ¦°Å(spÆÄÀÏÀ» ºÒ·¯¿Í¼­ ´Ù½Ã ÃøÁ¤ÇÒ ¼ö ÀÖ±âÀ§ÇØ¼­)
                        For j = 0 To StarProbe.ChipCountY
                            Wafer(i, j).InkDot = 0
                        Next j
                    Next i

                    FILE_NAMEING = File_Name
                    Form_StarProbe_MeasureDataSave.Display_View5                'wmd01 ÆÄÀÏ ÀúÀå (server, hdd)
                    Form_StarProbe_MeasureDataSave.Display_View5_1              'map2 [ 2022.09.29 ] Ãß°¡
                    Form_StarProbe_MeasureDataSave.Display_View_Change          'txt ÆÄÀÏ ÀúÀå (server, hdd)
                End If

                Call StarProbe_FileSave_Data(File_Name & ".SP")                 'sp ÆÄÀÏ ÀúÀå (hdd)

                BMP_file = File_Name & ".PNG"                                   '[ 2021.05.11 ] : BMP->PNG
                Call Form_StarProbe_MeasureDataSave.Display_View                '±×¸² ÆÄÀÏ ÀúÀå (server, hdd)

                If AutoAlign_Flag = False Then                                  'auto alignÀÌ ¾Æ´Ñ °æ¿ì
                    Z_HEIGHT.Command2.Visible = IIf(StarProbe.Ink_After = 2, True, False)
                    Z_HEIGHT.Show 1
                End If
                '=====================================================================================================================

                If StarProbe.Ink_After = 2 Then                                 'ink off
                    If AutoAlign_Flag = True Then                               '[ 2021.04.15 ] : autoÀÎ °æ¿ì ink offÇÏ¸é ink¾øÀÌ ¿¬¼Óµ¿ÀÛÇÏµµ·Ï ¼öÁ¤.
                        Call StarProbe_After_Ink_Dot_noink
                    Else                                                        'auto alignÀÌ ¾Æ´Ñ °æ¿ì
                        If INK_OFF_TEST = False Then
                            RESET_DATA
                            Check1(3).value = 0
                            bStop = True
                            bStarprobe_AfterInk = False
                            Call StarProbe_Zero_point       'Ãß°¡
                        Else
                            Call StarProbe_After_Ink_Dot
                        End If
                    End If
                Else                                                            'ink direct or after
                    Call StarProbe_After_Ink_Dot
                End If

                bStarProbe_Auto_Start = False                                   'star probe auto test off
                TESTING_flag = False                '2016.03.11

                'Wafer End½ÅÈ£¸¦ ¹Þ¾ÒÀ¸¹Ç·Î º¯¼ö¸¦ ÃÊ±âÈ­ ÇØÁØ´Ù.
                STT_time = ""
                END_time = ""

                '[ 2021.03.09 ] : lot endÀÎ °æ¿ì º¯¼ö Å¬¸®¾î(Áßº¹ÀúÀåÀ» ¸·±âÀ§ÇØ¼­ Ãß°¡)
                If Slot_Max_Count = Int(W_NO) Then
                    For i = 0 To 24
                        NOW_NO(i) = True
                    Next i
                End If
            End If
        End If
        If bStarProbe_Auto_Start = False Or bStop Then Exit Do
    Loop
    
    SSPanel_BadCount.Caption = StarProbe.CountBadDie
    Text_BadCount.Text = StarProbe.CountBadDie
    SSPanel_GoodCount.Caption = StarProbe.CountGoodDie
    SSPanel2(1).Caption = Test_Cnt
    Text_TotalCount.Text = Test_Cnt
    
    If Test_Cnt = 0 Then
        SSPanel2(2).Caption = Good_Cnt & Space(1) & "(" & "0.00" & "%)"
    Else
        SSPanel2(2).Caption = Good_Cnt & Space(1) & "(" & Format(Good_Cnt / Test_Cnt * 100, "0.00") & "%)"
    End If
    Text_GoodCount.Text = Good_Cnt
End Sub



Sub StarProbe_Auto_Test_CH4()
    Dim XYpos As String
    Dim xx As Integer, yy As Integer
    Dim forx As Integer, fory As Integer
    Dim bRight As Boolean
    Dim bCon As Boolean
    
    Dim File_Name, SP_File_Name As String

    Dim bStart As Boolean
    Dim val, sval As String
    Dim sOldFileName As String
    
    Dim FIRST_CHK As Boolean
    
    Dim Y_change As Boolean             '[ 2022.05.31 ]
  
    bStarProbe_Auto_Start = True
    bRight = True
    ErrorStop = False
      
    xx = StarProbe.StartChip.x
    Start_X = xx
    yy = StarProbe.StartChip.y
      
    Check1(3).Enabled = True
    bStart = True
          
    If Stop_Measure Then
        xx = Stop_xx
        yy = Stop_yy
    End If
    
    FIRST_CHK = True
    Y_change = False                    '[ 2022.05.31 ]
          
    Do While bStarProbe_Auto_Start
        DoEvents
        If bPause_Flag = True Then     'pause & coutinue flag
            Call StarProbe_Pause
            Call StarProbe_Z_Down
            If DemoMode = 0 Then
                If ACCO = 1 Or Model_Select = 1 Then                            'ACCO only
                    MSComm1.Output = "PA" & vbCrLf
                Else
                    MSComm1.Output = "PA" & vbLf
                End If
            End If
        '[ 2022.07.29 ] : yÃà±âÁØ 1/2 º¸´Ù Å©¸é¼­ Ä§ÀûÈ®ÀÎ ¼³Á¤ÀÌ µÈ WaferÀÎ °æ¿ì Ä§ÀûÈ®ÀÎ ¸Þ½ÃÁö ³ªÅ¸³½´Ù.
        ElseIf yy > Int(StarProbe.ChipCountY / 2) And Needle_Chk_Ok = False And Needle_Chk(lblWafer.Caption - 1) = True Then
            Command_Stop_Click
            Form_Check_List.Show 1
            Needle_Chk_Ok = True
            If Needle_check_flag = False Then
                Check1(3).value = 1
            End If
        ElseIf StarProbe.Tip_Clean = 1 And StarProbe.Tipclean_Count >= StarProbe.Tipclean_Count_Limit Then
            '[ 2022.07.20 ]
            If LOG_FILE_ON = 1 Then SelectExt.Log_Data_Save (13)
            If DemoMode = 0 Then
                Call StarProbe_Pause
                Call StarProbe_Z_Down
                If ACCO = 1 Or Model_Select = 1 Then                            'ACCO only
                    MSComm1.Output = "PA" & vbCrLf
                Else
                    MSComm1.Output = "PA" & vbLf
                End If
                Sleep 1000

                Call StarProbe_tip_clean

                Sleep 5000
            End If

            StarProbe.Tipclean_Count = 0
            Text7.Refresh
            Text7.Text = StarProbe.Tipclean_Count

            If DemoMode = 0 Then
                Call StarProbe_Pause
                If ACCO = 1 Or Model_Select = 1 Then                            'ACCO only
                    MSComm1.Output = "CO" & vbCrLf
                Else
                    MSComm1.Output = "CO" & vbLf
                End If

'                Call StarProbe_XY_Moving((XAxis), (YAxis))
'
'                Sleep 100
'
'                If StarProbe_Motor_End_check Then
'                    MsgBox "Motor not end check !", 16, "STAR PROBE"
'                End If
                Call StarProbe_Z_UP
            End If
            Sleep 1000
        Else
            If Stop_Measure Then
                bStart = False
                                          
                XYpos = " "
                                     
                If ErrorStop = True Then Exit Do   'Ãß°¡
                If Test_Ready = False And TimerCheck = True Then Exit Do   '2003.09.16 coupling
                Call StarProbe_Measure_CH4(Text5, Text6, bRight, False)
            Else
                For ZZZ = 0 To 3
                    If Wafer(xx, yy + ZZZ).Chip And _
                        Not Wafer(xx, yy + ZZZ).ChipMask And _
                        Not Wafer(xx, yy + ZZZ).ChipSkipDie And _
                        Not Wafer(xx, yy + ZZZ).ChipPlate Then
                                               
                        If Not Wafer(xx, yy + ZZZ).flag Then
'                            If StarProbe.MeasureAll = 1 Or (XPitch(TT_NO) = 1 And YPitch(TT_NO) = 1) Then
'                            Else
'                                If Y_change = True Then             '[ 2022.05.31 ] : yÃàÀÌ ¹Ù²ï °æ¿ì ¹ÙÁ®³ª°£´Ù.
'                                    Y_change = False
'                                    If bRight = True Then
'                                        xx = xx
'                                    Else
'                                        xx = xx + 1
'                                    End If
'                                    'Exit For
'                                End If
'                            End If
                            
                            bStart = False
                                                      
                            XYpos = " "
                                                 
                            If ErrorStop = True Then Exit Do   'Ãß°¡
                                        
                            If Test_Ready = False And TimerCheck = True Then Exit Do   '2003.09.16 coupling
                            ''''''''''''''''''''''''''''''''''[2020.03.17] : ½ÇÁ¦ÃøÁ¤ÁÂÇ¥ÀúÀå
                            Wafer(xx, yy + ZZZ).Real_Chk = True
                            ''''''''''''''''''''''''''''''''''
                            VScroll_Zoom.value = (yy / (Abs(StarProbe.Max.y) + Abs(StarProbe.Min.y) + 1)) * 1000
                            HScroll_Zoom.value = (xx / (Abs(StarProbe.Max.x) + Abs(StarProbe.Min.x) + 1)) * 1000
                            
                            StarProbe.CurrentChip.x = xx - StarProbe.StartChip.x
                            StarProbe.CurrentChip.y = yy - StarProbe.StartChip.y
                            
                            Text5 = StarProbe.CurrentChip.x
                            Text6 = StarProbe.CurrentChip.y
                            
                            Shape_Chip.Top = (yy * StarProbe.DisplayChipSizeY) - 2
                            Shape_Chip.Left = (xx * StarProbe.DisplayChipSizeX) - 2
                            Shape_Chip.Height = (StarProbe.DisplayChipSizeY * 4) + 4
                            
                            Label_ChipPosition = StarProbe.CurrentChip.x & "/" & StarProbe.CurrentChip.y
                            Call StarProbe_Measure_CH4(Text5, Text6, bRight, False)
                            Call Display_Chip(pZoom, pOriginal, Text5, Text6)
                            
                            ''
                            SSPanel2(1).Caption = Test_Cnt
                            Text_TotalCount.Text = Test_Cnt
                            
                            If Test_Cnt = 0 Then
                                SSPanel2(2).Caption = Good_Cnt & Space(1) & "(" & "0.00" & "%)"
                            Else
                                SSPanel2(2).Caption = Good_Cnt & Space(1) & "(" & Format(Good_Cnt / Test_Cnt * 100, "0.00") & "%)"
                            End If
                            Text_GoodCount.Text = Good_Cnt
                            ''
                        End If
                    End If
                Next ZZZ
            End If
                      
            If Stop_Measure Then
                Stop_xx = xx
                Stop_yy = yy
            End If
            
            If bRight Then
                If StarProbe.MeasureAll = 1 Or (XPitch(TT_NO) = 1 And YPitch(TT_NO) = 1) Then         'Àü¼ö°Ë»ç x+1
                    xx = xx + 1
                Else
                    xx = xx + (XPitch(TT_NO))
                End If
                
                If StarProbe.MeasureAll = 1 Or (XPitch(TT_NO) = 1 And YPitch(TT_NO) = 1) Then         'Àü¼ö°Ë»ç x+1
                    If xx > StarProbe.ChipCountX And FIRST_CHK = True Then     'ÇöÀç xÁÂÇ¥°¡ ¿þ¾îÆÛÀÇ xÃà ÁÂÇ¥º¸´Ù Å«°æ¿ì(¿ìÃø ³¡±îÁö ÀÌµ¿ÇÑ°æ¿ì)
                        FIRST_CHK = False
                        xx = StarProbe.ChipCountX           '¿ìÃø ³¡À¸·Î ÀÌµ¿
                        
                        For forx = StarProbe.ChipCountX To 0 Step -1
                            If Wafer(forx, yy).Chip And _
                               Not Wafer(forx, yy).ChipMask And _
                               Not Wafer(forx, yy).ChipSkipDie And _
                               Not Wafer(forx, yy).ChipPlate Then           'Á¤»ó Ä¨ÀÎ °æ¿ì
                                xx = StarProbe.ChipCountX                   'xx¿¡ °ªÁöÁ¤
                                Exit For
                            End If
                        Next
                        bRight = False                      'ÁÂÃøÀÌµ¿ ¼³Á¤
                    End If
                Else                                        'Àü¼ö°Ë»ç + step mode
                    If xx > StarProbe.ChipCountX And FIRST_CHK = True Then      'ÇöÀç xÁÂÇ¥°¡ ¿þ¾îÆÛÀÇ xÃà ÁÂÇ¥º¸´Ù Å«°æ¿ì(¿ìÃø ³¡±îÁö ÀÌµ¿ÇÑ°æ¿ì)
                        FIRST_CHK = False
                        xx = StarProbe.ChipCountX           '¿ìÃø ³¡À¸·Î ÀÌµ¿
                        
                        For forx = StarProbe.ChipCountX To 0 Step -1
                            For ZZZ = 0 To 3
'                                If Wafer(forx, yy + ZZZ).Chip And _
'                                   Not Wafer(forx, yy + ZZZ).ChipMask And _
'                                   Not Wafer(forx, yy + ZZZ).ChipSkipDie And _
'                                   Not Wafer(forx, yy + ZZZ).ChipPlate Then         'Á¤»ó Ä¨ÀÎ °æ¿ì
'                                    XX = Start_X                   'xx ½ÃÀÛÁöÁ¡À¸·Î ÀÌµ¿
'    '                                Exit For
'                                Else
                                    If Wafer(forx, yy + ZZZ).flag = True Then       '¿ìÃø¿¡¼­ ÁÂÃøÀ¸·Î ÀÌµ¿ÇÏ¸é¼­ Ã³À½ ÃøÁ¤ÇÑ ÁöÁ¡À» Ã£´Â ÄÚµå
                                        xx = Start_X
                                    End If
'                                End If
                            Next ZZZ
                        Next
                        bRight = False                      'ÁÂÃøÀÌµ¿ ¼³Á¤
                    End If
                End If
                
                If xx > StarProbe.ChipCountX Then      'ÇöÀç xÁÂÇ¥°¡ ¿þ¾îÆÛÀÇ xÃà ÁÂÇ¥º¸´Ù Å«°æ¿ì(¿ìÃø ³¡±îÁö ÀÌµ¿ÇÑ°æ¿ì)
'                    yy = yy + 4                     '2015.09.14
                    yy = yy + (YPitch(TT_NO) * 4)      '[ 2018.01.29 ]
                    
                    Y_change = True                     '[ 2022.05.31 ]
                    If yy < StarProbe.ChipCountY Then   '2015.09.14
'                        VScroll_Zoom.value = (yy / (Abs(StarProbe.Max.y) + Abs(StarProbe.Min.y) + 1)) * 1000
                    End If
                    xx = StarProbe.ChipCountX           '¿ìÃø ³¡À¸·Î ÀÌµ¿
                    
                    For forx = StarProbe.ChipCountX To 0 Step -1
                        If Wafer(forx, yy).Chip And _
                           Not Wafer(forx, yy).ChipMask And _
                           Not Wafer(forx, yy).ChipSkipDie And _
                           Not Wafer(forx, yy).ChipPlate Then           'Á¤»ó Ä¨ÀÎ °æ¿ì
                            'xx = forx                   'xx¿¡ °ªÁöÁ¤
                            xx = StarProbe.ChipCountX                   'xx¿¡ °ªÁöÁ¤
                            Exit For
                        End If
                    Next
                    bRight = False                      'ÁÂÃøÀÌµ¿ ¼³Á¤
                End If
            Else
                If StarProbe.MeasureAll = 1 Or (XPitch(TT_NO) = 1 And YPitch(TT_NO) = 1) Then         'Àü¼ö°Ë»ç x+1
                    xx = xx - 1
                Else                                    'step mode °Ë»ç
                    xx = xx - (XPitch(TT_NO))
                End If
                If xx < 0 Then
'                    yy = yy + 4                     '2015.09.14
                    yy = yy + (YPitch(TT_NO) * 4)      '[ 2018.01.29 ]
                    Y_change = True                 '[ 2022.05.31 ]
                    If yy < StarProbe.ChipCountY Then   '2015.09.14
'                        VScroll_Zoom.value = (yy / (Abs(StarProbe.Max.y) + Abs(StarProbe.Min.y) + 1)) * 1000
                    End If

                    xx = 0
                    
                    For forx = 0 To StarProbe.ChipCountX
                        If Wafer(forx, yy).Chip And _
                            Not Wafer(forx, yy).ChipMask And _
                            Not Wafer(forx, yy).ChipSkipDie And _
                            Not Wafer(forx, yy).ChipPlate Then   'Á¤»ó Ä¨ÀÎ °æ¿ì
                            xx = 0                       '2015.09.14
                            Exit For
                        End If
                    Next
                    bRight = True                       '¿ìÃøÀÌµ¿ ¼³Á¤
                End If
            End If

            '[ ÆÄÀÏ ÀúÀåÇÏ´Â ºÎºÐ ]
            If yy > StarProbe.ChipCountY Then
                Needle_Chk_Ok = False
                ''''''''''''''''''''''[demo]'''''''''''''''''''''
                If TESTER_OFF = True Then                               '[ 2022.08.10 ] : tester offÀÎ °æ¿ì ink¾øÀÌ ¿¬¼Ó ÀÛ¾÷
                    Call Command_Map_Clear_Click
                    Call Command_DisplayWafer_Click
                    xx = StarProbe.StartChip.x
                    yy = StarProbe.StartChip.y
                    bRight = True
                Else
                    '=====================================================================================================================
                    If (TT_NO + 1) <= 9 Then                                        '[ 2020.02.07 ] : 1~9±îÁö´Â 01~09·Î Ç¥½ÃÇÑ´Ù.
                        W_NO = "0" & TT_NO + 1
                    Else                                                            '10ÀÌ»óÀº ±×´ë·Î »ç¿ëÇÏ¸é µÈ´Ù.
                        W_NO = TT_NO + 1
                    End If
    
                    If SaveDrive = 0 Then
                        File_Name = "C:\data\" & UCase(Text1(0).Text) & "\" & UCase(Text1(0).Text) & "_" & W_NO        'HDD data save path
                    Else
                        File_Name = "D:\data\" & UCase(Text1(0).Text) & "\" & UCase(Text1(0).Text) & "_" & W_NO        'HDD data save path
                    End If
    
                    If No_Probe = True Then                                         '[ edge ink ]
                    Else                                                            '[ normal ]
                        For i = 0 To StarProbe.ChipCountX                           '2015.12.04 : ink ÈçÀûÁ¦°Å(spÆÄÀÏÀ» ºÒ·¯¿Í¼­ ´Ù½Ã ÃøÁ¤ÇÒ ¼ö ÀÖ±âÀ§ÇØ¼­)
                            For j = 0 To StarProbe.ChipCountY
                                Wafer(i, j).InkDot = 0
                            Next j
                        Next i
    
                        FILE_NAMEING = File_Name
                        Form_StarProbe_MeasureDataSave.Display_View5                'wmd01 ÆÄÀÏ ÀúÀå (server, hdd)
                        Form_StarProbe_MeasureDataSave.Display_View5_1              'map2 [ 2022.09.29 ] Ãß°¡
                        Form_StarProbe_MeasureDataSave.Display_View_Change          'txt ÆÄÀÏ ÀúÀå (server, hdd)
                    End If
    
                    Call StarProbe_FileSave_Data(File_Name & ".SP")                 'sp ÆÄÀÏ ÀúÀå (hdd)
    
                    BMP_file = File_Name & ".PNG"                                   '[ 2021.05.11 ] : BMP->PNG
                    Call Form_StarProbe_MeasureDataSave.Display_View                '±×¸² ÆÄÀÏ ÀúÀå (server, hdd)
    
                    If AutoAlign_Flag = False Then                                  'auto alignÀÌ ¾Æ´Ñ °æ¿ì
                        Z_HEIGHT.Command2.Visible = IIf(StarProbe.Ink_After = 2, True, False)
                        Z_HEIGHT.Show 1
                    End If
                    '=====================================================================================================================
    
                    If StarProbe.Ink_After = 2 Then                                 'ink off
                        If AutoAlign_Flag = True Then                               '[ 2021.04.15 ] : autoÀÎ °æ¿ì ink offÇÏ¸é ink¾øÀÌ ¿¬¼Óµ¿ÀÛÇÏµµ·Ï ¼öÁ¤.
                            Call StarProbe_After_Ink_Dot_noink
                        Else                                                        'auto alignÀÌ ¾Æ´Ñ °æ¿ì
                            If INK_OFF_TEST = False Then
                                RESET_DATA
                                Check1(3).value = 0
                                bStop = True
                                bStarprobe_AfterInk = False
                                Call StarProbe_Zero_point       'Ãß°¡
                            Else
                                Call StarProbe_After_Ink_Dot
                            End If
                        End If
                    Else                                                            'ink direct or after
                        Call StarProbe_After_Ink_Dot
                    End If
    
                    bStarProbe_Auto_Start = False                                   'star probe auto test off
                    TESTING_flag = False                '2016.03.11
    
                    'Wafer End½ÅÈ£¸¦ ¹Þ¾ÒÀ¸¹Ç·Î º¯¼ö¸¦ ÃÊ±âÈ­ ÇØÁØ´Ù.
                    STT_time = ""
                    END_time = ""
    
                    '[ 2021.03.09 ] : lot endÀÎ °æ¿ì º¯¼ö Å¬¸®¾î(Áßº¹ÀúÀåÀ» ¸·±âÀ§ÇØ¼­ Ãß°¡)
                    If Slot_Max_Count = Int(W_NO) Then
                        For i = 0 To 24
                            NOW_NO(i) = True
                        Next i
                    End If
                End If
            End If
        End If
        If bStarProbe_Auto_Start = False Or bStop Then Exit Do
    Loop
    
    SSPanel_BadCount.Caption = StarProbe.CountBadDie
    Text_BadCount.Text = StarProbe.CountBadDie
    SSPanel_GoodCount.Caption = StarProbe.CountGoodDie
    SSPanel2(1).Caption = Test_Cnt
    Text_TotalCount.Text = Test_Cnt
    
    If Test_Cnt = 0 Then
        SSPanel2(2).Caption = Good_Cnt & Space(1) & "(" & "0.00" & "%)"
    Else
        SSPanel2(2).Caption = Good_Cnt & Space(1) & "(" & Format(Good_Cnt / Test_Cnt * 100, "0.00") & "%)"
    End If
    Text_GoodCount.Text = Good_Cnt
End Sub

Sub StarProbe_Clear_Test(Index As Integer)
    Dim xx As Integer, yy As Integer
    Dim forx As Integer, fory As Integer
    Dim bRight As Boolean
                  
    xx = StarProbe.StartChip.x
    yy = StarProbe.StartChip.y
    bRight = True
    Do
        DoEvents
        If Wafer(xx, yy).Chip And _
            Not Wafer(xx, yy).ChipMask And _
            Not Wafer(xx, yy).ChipSkipDie And _
            Not Wafer(xx, yy).ChipPlate Then
            
            If Wafer(xx, yy).flag Then
                If Wafer(xx, yy).BIN = Index Then
                    '[ 2021.10.27 ] : bin1(ok bin)À» Áö¿î °æ¿ì Ã³¸®.
                    If Index = 1 Then Good_Cnt = 0              'bin1 = good bin
                    SSPanel2(2).Caption = Good_Cnt & Space(1) & "(" & Format(Good_Cnt / Test_Cnt * 100, "0.00") & "%)"
                    
                    Bin_Result(Index) = 0 'Bin_Result(Index) - 1
                    Text_BinCount(Index).Text = 0  'Text_BinCount(Index) - 1
                    Text_Bin_Count_No(Index) = 0 'Text_Bin_Count_No(Index) - 1
                    
                    Test_Cnt = Test_Cnt - 1
                    SSPanel2(1).Caption = Test_Cnt
                    
                    Text_TotalCount.Text = SSPanel_TotalCount.Caption
                    Text_GoodCount.Text = SSPanel_GoodCount.Caption
                    Text_BadCount.Text = SSPanel_BadCount.Caption
                                        
                    SSPanel_GoodCount.Caption = val(SSPanel_GoodCount.Caption) + 1
                    SSPanel_BadCount.Caption = val(SSPanel_BadCount.Caption) - 1
                    
                    '
                    If Wafer(xx, yy).ChipSkipDie Or _
                        Wafer(xx, yy).ChipPlate Or _
                        Wafer(xx, yy).ChipMask Then
                        
                        StarProbe.CountSkipDie = StarProbe.CountSkipDie - 1
                        StarProbe.CountGoodDie = StarProbe.CountGoodDie + 1
                    End If
                
                    If Wafer(xx, yy).FlagBad Then
                        StarProbe.CountBadDie = StarProbe.CountBadDie - 1
                        StarProbe.CountGoodDie = StarProbe.CountGoodDie + 1
                    End If
                
                    SSPanel_SkipCount.Caption = StarProbe.CountSkipDie
                    Text_SkipCount.Text = StarProbe.CountSkipDie
                    SSPanel_BadCount.Caption = StarProbe.CountBadDie
                    Text_BadCount.Text = StarProbe.CountBadDie
                    SSPanel_GoodCount.Caption = StarProbe.CountGoodDie
                
                    Wafer(xx, yy).ChipSkipDie = False
                    Wafer(xx, yy).ChipInk = False
                    Wafer(xx, yy).ChipPlate = False
                    Wafer(xx, yy).ChipMask = False
                
                    Wafer(xx, yy).flag = False
                    Wafer(xx, yy).FlagBad = False
                    Wafer(xx, yy).BIN = 0
                                                                            
                    StarProbe.CurrentChip.x = xx - StarProbe.StartChip.x
                    StarProbe.CurrentChip.y = yy - StarProbe.StartChip.y
                    
                    Text5.Text = StarProbe.CurrentChip.x
                    Text6.Text = StarProbe.CurrentChip.y
                    Call Display_Chip(pZoom, pOriginal, Text5.Text, Text6.Text)
                End If
            End If
        End If
                  
        If bRight Then
            xx = xx + 1
            If xx > StarProbe.ChipCountX Then
                yy = yy + 1
                xx = StarProbe.ChipCountX
                For forx = StarProbe.ChipCountX To 0 Step -1
                    If Wafer(forx, yy).Chip And _
                        Not Wafer(forx, yy).ChipMask And _
                        Not Wafer(forx, yy).ChipSkipDie And _
                        Not Wafer(forx, yy).ChipPlate Then
                        
                        xx = forx
                        Exit For
                    End If
                Next
                bRight = False
            End If
        Else
            xx = xx - 1
            If xx < 0 Then
                yy = yy + 1
                xx = 0
                For forx = 0 To StarProbe.ChipCountX
                    If Wafer(forx, yy).Chip And _
                       Not Wafer(forx, yy).ChipMask And _
                       Not Wafer(forx, yy).ChipSkipDie And _
                       Not Wafer(forx, yy).ChipPlate Then
                       xx = forx
                       Exit For
                    End If
                Next
                bRight = True
            End If
        End If
        If yy > StarProbe.ChipCountY Then
            Exit Do
        End If
    Loop
End Sub


Sub StarProbe_Auto_Probing()
'[2021.09.02] : trimÃß°¡
    Dim Wafer_Status As String
    Dim bpoint As Boolean
    Dim ErrorCount As Integer
    Dim Profile_Check, Align_check As String
    Set WaitTime_Delay = New ccrpStopWatch
    Starprobe_Auto_probing_Flag = True
        
    Wafer_Status = " "
    Wafer_Status = Trim(Starprobe_Requst_State_Wafer_On_Chuck)
    
    If Wafer_Status = "0" Or Wafer_Status = " " Or Wafer_Status = "" Then
        MsgBox "Wafer is not on the Chuck..", 16, "Error"
        Starprobe_Auto_probing_Flag = False
        '[ 2021.03.09 ] : lot endÀÎ °æ¿ì º¯¼ö Å¬¸®¾î(Áßº¹ÀúÀåÀ» ¸·±âÀ§ÇØ¼­ Ãß°¡)
        If Slot_Max_Count = Int(W_NO) Then
            For i = 0 To 24
                NOW_NO(i) = True
            Next i
        End If
        Exit Sub
    End If
    
    Profile_Check = " "
    Profile_Check = Trim(Starprobe_Requst_State_Wafer_Profile)
    
    Align_check = " "
    Align_check = Trim(Starprobe_Requst_State_Wafer_Auto_Aligned)
    
    If Int(Profile_Check) = 1 And Int(Align_check) = 1 Then
        Exit Sub
    End If
  
    bpoint = True
    ErrorCount = 0
    Wafer_Status = " "
  
    Call StarProbe_Profile_Wafer_Thickness
  
    WaitTime_Delay.Reset
    
    Do
        DoEvents
        If WaitTime_Delay.Elapsed > 20000 Then Exit Do
    Loop
   
    Do While bpoint
        DoEvents
        Wafer_Status = Trim(Starprobe_Requst_State_Wafer_Profile)
        If Wafer_Status = "" Or Wafer_Status = " " Then                           '[ 2021.06.28 ] : ¾ó¶óÀÎÀÌ ±æ¾îÁú°æ¿ì °ø¹é¹®ÀÚ°¡ ¿À´Â °æ¿ì Ã³¸®
            ErrorCount = ErrorCount + 1
        ElseIf Int(Wafer_Status) = 1 Then
            bpoint = False
        ElseIf Int(Wafer_Status) = 0 Then
            ErrorCount = ErrorCount + 1
        End If
     
        If ErrorCount > 5 Then
            MsgBox "GPIB Communication is Error.", 16, "GPIB ERROR"
            Starprobe_Auto_probing_Flag = False
            bpoint = False
        End If
        If bpoint = False Or bStop = True Then Exit Do
    Loop
  
    If Starprobe_Auto_probing_Flag = False Then Exit Sub
    
    bpoint = True
    ErrorCount = 0
    Wafer_Status = " "
  
    Call StarProbe_Auto_Align_Wafer(1)
    
    WaitTime_Delay.Reset
    
    Do
        DoEvents
        If WaitTime_Delay.Elapsed > 20000 Then Exit Do
    Loop
  
    Do While bpoint
        DoEvents
        Wafer_Status = Trim(Starprobe_Requst_State_Wafer_Auto_Aligned)
        If Wafer_Status = "" Or Wafer_Status = " " Then                            '[ 2021.06.28 ] : ¾ó¶óÀÎÀÌ ±æ¾îÁú°æ¿ì °ø¹é¹®ÀÚ°¡ ¿À´Â °æ¿ì Ã³¸®
            ErrorCount = ErrorCount + 1
        ElseIf Int(Wafer_Status) = 1 Then
            bpoint = False
        ElseIf Int(Wafer_Status) = 0 Then
            ErrorCount = ErrorCount + 1
        End If
         
        If ErrorCount > 5 Then
            MsgBox "GPIB Communication is Error.", 16, "GPIB ERROR"
            Starprobe_Auto_probing_Flag = False
            bpoint = False
        End If
        If bpoint = False Or bStop = True Then Exit Do   'Ãß°¡
    Loop
    Sleep 7000
End Sub

Public Sub ChipPosition(x As Integer, y As Integer)
    Dim icount As Integer
    Dim forx As Integer, fory As Integer
    Dim b As Boolean
    Dim s As Integer
    
    If CH_SET = 1 Then
        Text_AreaY = 1
    ElseIf CH_SET = 2 Then
        Text_AreaY = 2
    Else
        Text_AreaY = 4
    End If
    
    For icount = 0 To 15
        Text_Chip(icount).Text = ""
        Text_ChipX(icount).Text = ""
        Text_ChipY(icount).Text = ""
        Text_ChipBIN(icount).Text = ""
    Next
    
    icount = 0
    
    Text_StartX.Text = x - StarProbe.StartChip.x
    Text_StartY.Text = y - StarProbe.StartChip.y

    For fory = (val(Text_AreaY) - 1) To 0 Step -1
        For forx = (val(Text_AreaX) - 1) To 0 Step -1
            b = IIf(Wafer(x + forx, y + fory).Chip And _
                    Not Wafer(x + forx, y + fory).ChipMask And _
                    Not Wafer(x + forx, y + fory).ChipSkipDie And _
                    Not Wafer(x + forx, y + fory).ChipPlate, True, False)

            If b Then
                Text_Chip(icount).Text = "Chip"
            Else
                Text_Chip(icount).Text = ""
            End If
            Text_ChipX(icount).Text = x + forx - StarProbe.StartChip.x
            Text_ChipY(icount).Text = y + fory - StarProbe.StartChip.y
            icount = icount + 1
        Next
    Next
    
    icount = 0
    s = 0

    For fory = val(Text_AreaY.Text) - 1 To 0 Step -1
        For forx = val(Text_AreaX.Text) - 1 To 0 Step -1
            If Text_Chip(icount).Text = "Chip" Then
                s = s + (2 ^ icount)
            End If
            icount = icount + 1
        Next
    Next
    Text_ChipTest.Text = s
End Sub
