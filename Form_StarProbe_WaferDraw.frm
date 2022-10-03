VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form_StarProbe_WaferDraw 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "Wafer Draw"
   ClientHeight    =   9255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13575
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form_StarProbe_WaferDraw.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   617
   ScaleMode       =   3  '픽셀
   ScaleWidth      =   905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   8820
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9135
      Left            =   9300
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   60
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   16113
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Information"
      TabPicture(0)   =   "Form_StarProbe_WaferDraw.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "SSPanel19(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "SSPanel19(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "SSPanel13"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Command_ChipOk"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "SSPanel1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "SSPanel2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "SSPanel5"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "SSPanel9"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      Begin Threed.SSPanel SSPanel9 
         Height          =   975
         Left            =   120
         TabIndex        =   20
         Top             =   4560
         Width           =   3975
         _Version        =   65536
         _ExtentX        =   7011
         _ExtentY        =   1720
         _StockProps     =   15
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
         BevelOuter      =   1
         Begin VB.TextBox Text_ChipCountX 
            Alignment       =   2  '가운데 맞춤
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   405
            Left            =   2220
            Locked          =   -1  'True
            TabIndex        =   5
            Text            =   "0.32"
            Top             =   60
            Width           =   1155
         End
         Begin VB.TextBox Text_ChipCountY 
            Alignment       =   2  '가운데 맞춤
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   405
            Left            =   2220
            Locked          =   -1  'True
            TabIndex        =   6
            Text            =   "0.32"
            Top             =   480
            Width           =   1155
         End
         Begin Threed.SSPanel SSPanel10 
            Height          =   375
            Left            =   1560
            TabIndex        =   21
            Top             =   60
            Width           =   570
            _Version        =   65536
            _ExtentX        =   1005
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "X"
            BackColor       =   12632256
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
         Begin Threed.SSPanel SSPanel11 
            Height          =   795
            Left            =   60
            TabIndex        =   22
            Top             =   60
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   1402
            _StockProps     =   15
            Caption         =   "Chip Count"
            BackColor       =   12632256
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
         Begin Threed.SSPanel SSPanel12 
            Height          =   375
            Left            =   1560
            TabIndex        =   23
            Top             =   480
            Width           =   570
            _Version        =   65536
            _ExtentX        =   1005
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "Y"
            BackColor       =   12632256
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
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "EA"
            Height          =   240
            Left            =   3480
            TabIndex        =   25
            Top             =   120
            Width           =   270
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "EA"
            Height          =   240
            Left            =   3480
            TabIndex        =   24
            Top             =   540
            Width           =   270
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   975
         Left            =   120
         TabIndex        =   14
         Top             =   3600
         Width           =   3975
         _Version        =   65536
         _ExtentX        =   7011
         _ExtentY        =   1720
         _StockProps     =   15
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
         BevelOuter      =   1
         Begin VB.TextBox Text_ChipSizeY 
            Alignment       =   2  '가운데 맞춤
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   405
            Left            =   2220
            TabIndex        =   4
            Text            =   "0.32"
            Top             =   480
            Width           =   1155
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   375
            Left            =   1560
            TabIndex        =   17
            Top             =   60
            Width           =   570
            _Version        =   65536
            _ExtentX        =   1005
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "X"
            BackColor       =   12632256
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
         Begin VB.TextBox Text_ChipSizeX 
            Alignment       =   2  '가운데 맞춤
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   405
            Left            =   2220
            TabIndex        =   3
            Text            =   "0.32"
            Top             =   60
            Width           =   1155
         End
         Begin Threed.SSPanel SSPanel6 
            Height          =   795
            Left            =   60
            TabIndex        =   15
            Top             =   60
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   1402
            _StockProps     =   15
            Caption         =   "Chip Size"
            BackColor       =   12632256
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
         Begin Threed.SSPanel SSPanel8 
            Height          =   375
            Left            =   1560
            TabIndex        =   19
            Top             =   480
            Width           =   570
            _Version        =   65536
            _ExtentX        =   1005
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "Y"
            BackColor       =   12632256
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
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "mm"
            Height          =   240
            Left            =   3480
            TabIndex        =   18
            Top             =   540
            Width           =   330
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "mm"
            Height          =   240
            Left            =   3480
            TabIndex        =   16
            Top             =   120
            Width           =   330
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   975
         Left            =   120
         TabIndex        =   10
         Top             =   2520
         Width           =   3975
         _Version        =   65536
         _ExtentX        =   7011
         _ExtentY        =   1720
         _StockProps     =   15
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
         BevelOuter      =   1
         Begin VB.OptionButton Option_WaferSize 
            BackColor       =   &H00C0FFFF&
            Caption         =   "mm"
            Height          =   255
            Index           =   1
            Left            =   1200
            TabIndex        =   35
            Top             =   540
            Width           =   855
         End
         Begin VB.OptionButton Option_WaferSize 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Inch"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   34
            Top             =   540
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.TextBox Text_WaferSizemm 
            Alignment       =   2  '가운데 맞춤
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
            ForeColor       =   &H000000FF&
            Height          =   405
            Left            =   2220
            TabIndex        =   32
            Text            =   "125"
            Top             =   480
            Width           =   1155
         End
         Begin Threed.SSPanel SSPanel3 
            Height          =   375
            Left            =   60
            TabIndex        =   11
            Top             =   60
            Width           =   2055
            _Version        =   65536
            _ExtentX        =   3625
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "Wafer Size"
            BackColor       =   12632256
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
         Begin VB.TextBox Text_WaferSize 
            Alignment       =   2  '가운데 맞춤
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   405
            Left            =   2220
            TabIndex        =   2
            Text            =   "6"
            Top             =   60
            Width           =   1155
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "mm"
            Height          =   240
            Left            =   3480
            TabIndex        =   33
            Top             =   540
            Width           =   330
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "inch"
            Height          =   240
            Left            =   3480
            TabIndex        =   12
            Top             =   120
            Width           =   360
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   2040
         Width           =   3975
         _Version        =   65536
         _ExtentX        =   7011
         _ExtentY        =   873
         _StockProps     =   15
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
         BevelOuter      =   1
         Begin VB.TextBox Text_Inch 
            Alignment       =   2  '가운데 맞춤
            BackColor       =   &H00C0C0FF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   405
            Left            =   2220
            TabIndex        =   1
            Text            =   "2.539"
            Top             =   60
            Width           =   1155
         End
         Begin Threed.SSPanel SSPanel4 
            Height          =   375
            Left            =   60
            TabIndex        =   13
            Top             =   60
            Width           =   2055
            _Version        =   65536
            _ExtentX        =   3625
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "1 Inch Unit"
            BackColor       =   12632256
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
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "mm"
            Height          =   240
            Left            =   3480
            TabIndex        =   9
            Top             =   135
            Width           =   330
         End
      End
      Begin VB.CommandButton Command_ChipOk 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Chip Ok"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         Style           =   1  '그래픽
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   8520
         Width           =   3975
      End
      Begin Threed.SSPanel SSPanel13 
         Height          =   1455
         Left            =   120
         TabIndex        =   26
         Top             =   480
         Width           =   3975
         _Version        =   65536
         _ExtentX        =   7011
         _ExtentY        =   2566
         _StockProps     =   15
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
         BevelOuter      =   1
         Begin VB.CommandButton Command_DeviceSave 
            Caption         =   "Save (Update)"
            Height          =   375
            Left            =   2040
            TabIndex        =   30
            Top             =   960
            Width           =   1875
         End
         Begin VB.CommandButton Command_DeviceLoad 
            Caption         =   "Load"
            Height          =   375
            Left            =   60
            TabIndex        =   29
            Top             =   960
            Width           =   1875
         End
         Begin VB.TextBox Text_DeviceName 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   435
            Left            =   60
            TabIndex        =   27
            Top             =   480
            Width           =   3855
         End
         Begin Threed.SSPanel SSPanel14 
            Height          =   375
            Left            =   60
            TabIndex        =   28
            Top             =   60
            Width           =   3855
            _Version        =   65536
            _ExtentX        =   6800
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "Device"
            BackColor       =   16761024
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
      End
      Begin Threed.SSPanel SSPanel19 
         Height          =   630
         Index           =   0
         Left            =   120
         TabIndex        =   36
         Top             =   5760
         Width           =   3975
         _Version        =   65536
         _ExtentX        =   7011
         _ExtentY        =   1111
         _StockProps     =   15
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
         BevelOuter      =   1
         Begin VB.TextBox Text_EdgeChipmm 
            Alignment       =   2  '가운데 맞춤
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   435
            Left            =   2280
            TabIndex        =   37
            Text            =   "3"
            Top             =   120
            Width           =   1095
         End
         Begin Threed.SSPanel SSPanel21 
            Height          =   435
            Index           =   0
            Left            =   120
            TabIndex        =   39
            Top             =   120
            Width           =   2055
            _Version        =   65536
            _ExtentX        =   3625
            _ExtentY        =   767
            _StockProps     =   15
            Caption         =   "Skip Die"
            BackColor       =   12632256
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
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "mm"
            Height          =   240
            Index           =   0
            Left            =   3480
            TabIndex        =   38
            Top             =   210
            Width           =   330
         End
      End
      Begin Threed.SSPanel SSPanel19 
         Height          =   630
         Index           =   1
         Left            =   120
         TabIndex        =   40
         Top             =   6480
         Width           =   3975
         _Version        =   65536
         _ExtentX        =   7011
         _ExtentY        =   1111
         _StockProps     =   15
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
         BevelOuter      =   1
         Begin VB.TextBox Text_PlateZone 
            Alignment       =   2  '가운데 맞춤
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   435
            Left            =   2280
            TabIndex        =   41
            Text            =   "45"
            Top             =   120
            Width           =   1095
         End
         Begin Threed.SSPanel SSPanel21 
            Height          =   435
            Index           =   1
            Left            =   120
            TabIndex        =   42
            Top             =   120
            Width           =   2055
            _Version        =   65536
            _ExtentX        =   3625
            _ExtentY        =   767
            _StockProps     =   15
            Caption         =   "Plate Zone"
            BackColor       =   12632256
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
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "mm"
            Height          =   240
            Index           =   1
            Left            =   3480
            TabIndex        =   43
            Top             =   210
            Width           =   330
         End
      End
   End
   Begin VB.PictureBox pWaferDraw 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   9135
      Left            =   60
      ScaleHeight     =   605
      ScaleMode       =   3  '픽셀
      ScaleWidth      =   605
      TabIndex        =   0
      Top             =   60
      Width           =   9135
      Begin VB.Shape Shape_Pattern 
         BorderColor     =   &H00C0C0C0&
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  '단색
         Height          =   375
         Index           =   4
         Left            =   5160
         Top             =   3300
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Shape Shape_Pattern 
         BorderColor     =   &H00C0C0C0&
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  '단색
         Height          =   375
         Index           =   3
         Left            =   3060
         Top             =   3360
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Shape Shape_Pattern 
         BorderColor     =   &H00C0C0C0&
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  '단색
         Height          =   375
         Index           =   2
         Left            =   4140
         Top             =   4380
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Shape Shape_Pattern 
         BorderColor     =   &H00C0C0C0&
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  '단색
         Height          =   375
         Index           =   1
         Left            =   4140
         Top             =   2100
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Shape Shape_Pattern 
         BorderColor     =   &H00C0C0C0&
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  '단색
         Height          =   375
         Index           =   0
         Left            =   4200
         Top             =   3360
         Visible         =   0   'False
         Width           =   435
      End
   End
End
Attribute VB_Name = "Form_StarProbe_WaferDraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command_ChipOk_Click()
    Command_ChipOk.Enabled = False
    Call WaferInformationChange
    
    If Option_WaferSize(0).value Then
        StarProbe.Unit = 0
    Else
        StarProbe.Unit = 1
    End If

    StarProbe.InchUnit = Text_Inch
    StarProbe.WaferSize = Text_WaferSize
    StarProbe.WaferSizemm = Text_WaferSizemm
    StarProbe.ChipSizeX = Text_ChipSizeX
    StarProbe.ChipSizeY = Text_ChipSizeY
    StarProbe.ChipCountX = Text_ChipCountX
    StarProbe.ChipCountY = Text_ChipCountY
    StarProbe.EdgeChipmm = Text_EdgeChipmm
    StarProbe.PlateZone = Text_PlateZone
    StarProbe_DeviceName = Text_DeviceName
    
    Call ChipOk
    
    StarProbeTemp = StarProbe
    Command_ChipOk.Enabled = True
End Sub

Private Sub Command_DeviceLoad_Click()
    CommonDialog1.CancelError = True

    On Error GoTo ErrorSub
    
    Dim sfilename As String, iReturn As Integer

    CommonDialog1.DialogTitle = "Star Probe Device File Load"
    CommonDialog1.Filter = "Star Probe Device Files(*.DEV)|*.DEV"
    CommonDialog1.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
    CommonDialog1.FileName = IIf(Trim(Text_DeviceName) = Empty, "c:\Star Probe\Device\*.DEV", Text_DeviceName)
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName <> "" Then
        sfilename = CommonDialog1.FileName
        iReturn = StarProbe_FileDeviceLoad(sfilename)
        Select Case iReturn
            Case 0
                Call MsgBox(sfilename & vbCrLf & "File Not Found ...", vbCritical + vbOKOnly, "Error")
            Case 2
                Call MsgBox(sfilename & vbCrLf & "Star Probe Not Device File ...", vbCritical + vbOKOnly, "Error")
        End Select
        
        Option_WaferSize(StarProbe.Unit).value = True
    
        Text_Inch = StarProbe.InchUnit
        Text_WaferSize = StarProbe.WaferSize
        Text_WaferSizemm = StarProbe.WaferSizemm
        Text_ChipSizeX = StarProbe.ChipSizeX
        Text_ChipSizeY = StarProbe.ChipSizeY
        Text_ChipCountX = StarProbe.ChipCountX
        Text_ChipCountY = StarProbe.ChipCountY
        Text_EdgeChipmm = StarProbe.EdgeChipmm
        
        Text_DeviceName = sfilename
        Call Command_ChipOk_Click
    End If
    
ErrorSub:
    CommonDialog1.CancelError = False
End Sub

Private Sub Command_DeviceSave_Click()
    CommonDialog1.CancelError = True
    On Error GoTo ErrorSub
    Dim sfilename As String

    CommonDialog1.DialogTitle = "Star Probe Device File Save (Update)"
    CommonDialog1.Filter = "Star Probe Device Files(*.DEV)|*.DEV"
    CommonDialog1.Flags = cdlOFNFileMustExist + cdlOFNOverwritePrompt + cdlOFNHideReadOnly
    CommonDialog1.FileName = IIf(Trim(Text_DeviceName) = Empty, "c:\Star Probe\Device\*.DEV", Text_DeviceName)
    CommonDialog1.ShowSave
    
    If CommonDialog1.FileName <> "" Then
        sfilename = CommonDialog1.FileName
        If Option_WaferSize(0).value Then
            StarProbe.Unit = 0
        Else
            StarProbe.Unit = 1
        End If
        StarProbe.InchUnit = Text_Inch
        StarProbe.WaferSize = Text_WaferSize
        StarProbe.WaferSizemm = Text_WaferSizemm
        StarProbe.ChipSizeX = Text_ChipSizeX
        StarProbe.ChipSizeY = Text_ChipSizeY
        StarProbe.ChipCountX = Text_ChipCountX
        StarProbe.ChipCountY = Text_ChipCountY
        StarProbe.EdgeChipmm = Text_EdgeChipmm
        Call StarProbe_FileDeviceSave(sfilename)
        Text_DeviceName = sfilename
    End If
    
ErrorSub:
    CommonDialog1.CancelError = False
End Sub

Private Sub Form_Load()
    StarProbeTemp = StarProbe
    
    Text_EdgeChipmm = StarProbe.EdgeChipmm
    Text_DeviceName = StarProbe_DeviceName

    Option_WaferSize(StarProbe.Unit).value = True

    Text_Inch = Format(StarProbe.InchUnit, "0.00000")
    Text_WaferSize = StarProbe.WaferSize
    Text_WaferSizemm = StarProbe.WaferSizemm
    Text_ChipSizeX = Format(StarProbe.ChipSizeX, "0.00000")
    Text_ChipSizeY = Format(StarProbe.ChipSizeY, "0.00000")
    Text_ChipCountX = StarProbe.ChipCountX
    Text_ChipCountY = StarProbe.ChipCountY
    Text_PlateZone = StarProbe.PlateZone
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Option_WaferSize(0).value Then
        StarProbe.Unit = 0
    Else
        StarProbe.Unit = 1
    End If
    StarProbe.InchUnit = Text_Inch
    StarProbe.WaferSize = Text_WaferSize
    StarProbe.WaferSizemm = Text_WaferSizemm
    StarProbe.ChipSizeX = Text_ChipSizeX
    StarProbe.ChipSizeY = Text_ChipSizeY
    StarProbe.ChipCountX = Text_ChipCountX
    StarProbe.ChipCountY = Text_ChipCountY
    StarProbe.EdgeChipmm = Text_EdgeChipmm
    StarProbe.PlateZone = Text_PlateZone
    StarProbe_DeviceName = Text_DeviceName
    StarProbe = StarProbeTemp
End Sub

Private Sub Option_WaferSize_Click(Index As Integer)
    If Option_WaferSize(0).value Then
        Text_WaferSize.Enabled = True
        Text_WaferSizemm.Enabled = False
        StarProbe.Unit = 0
    Else
        Text_WaferSize.Enabled = False
        Text_WaferSizemm.Enabled = True
        StarProbe.Unit = 1
    End If
    Call WaferInformationChange
End Sub

Private Sub Text_ChipSizeX_Change()
    Call WaferInformationChange
End Sub

Private Sub Text_ChipSizeY_Change()
    Call WaferInformationChange
End Sub

Private Sub Text_EdgeChipmm_Change()
    Call WaferInformationChange
End Sub

Private Sub Text_Inch_Change()
    Call WaferInformationChange
End Sub

Private Sub Text_PlateZone_Change()
    StarProbe.PlateZone = val(Text_PlateZone)
End Sub

Private Sub Text_WaferSize_Change()
    Call WaferInformationChange
End Sub

Public Sub WaferInformationChange()
    Command_ChipOk.Enabled = False
    
    Text_WaferSize.BackColor = vbWhite
    Text_WaferSizemm.BackColor = vbWhite
    Text_Inch.BackColor = vbWhite
    Text_ChipSizeX.BackColor = vbWhite
    Text_ChipSizeY.BackColor = vbWhite
    Text_ChipCountX.BackColor = vbWhite
    Text_ChipCountY.BackColor = vbWhite
    Text_EdgeChipmm.BackColor = vbWhite
    
    If val(Text_WaferSize) < 0.1 Then Text_WaferSize.BackColor = &HC0C0FF
    If val(Text_WaferSizemm) < 0.1 Then Text_WaferSizemm.BackColor = &HC0C0FF
    If val(Text_Inch) < 0.1 Then Text_Inch.BackColor = &HC0C0FF
    If val(Text_ChipSizeX) < 0.1 Then Text_ChipSizeX.BackColor = &HC0C0FF
    If val(Text_ChipSizeY) < 0.1 Then Text_ChipSizeY.BackColor = &HC0C0FF
    If val(Text_EdgeChipmm) < 0 Then Text_ChipSizeY.BackColor = &HC0C0FF
    
    If val(Text_WaferSize) < 0.1 Or _
       val(Text_WaferSizemm) < 0.1 Or _
       val(Text_Inch) < 0.1 Or _
       val(Text_ChipSizeX) < 0.1 Or _
       val(Text_ChipSizeY) < 0.1 Or _
       val(Text_EdgeChipmm) < 0 Then
       
        Text_ChipCountX = 0
        Text_ChipCountY = 0
    
        Text_ChipCountX.BackColor = &HC0C0FF
        Text_ChipCountY.BackColor = &HC0C0FF
       
        Exit Sub
    End If
    
    StarProbe.EdgeChipmm = val(Text_EdgeChipmm)
    
    Dim Unit As Integer
    Dim WaferSize As Double, WaferSizemm As Double
    Dim InchUnit As Double
    Dim ChipSizeX As Double, ChipSizeY As Double
    Dim Range As Double, RangeHalf As Double
    Dim RangeHalfX As Double, RangeHalfY As Double
    Dim HalfX As Integer, HalfY As Integer
    Dim ChipCountX As Integer, ChipCountY As Integer
    Dim ChipHalfCountX As Integer, ChipHalfCountY As Integer
    
    If Option_WaferSize(0).value Then
        Unit = 0
    Else
        Unit = 1
    End If
    WaferSize = val(Text_WaferSize)
    WaferSizemm = val(Text_WaferSizemm)
    InchUnit = val(Text_Inch)
    ChipSizeX = val(Text_ChipSizeX)
    ChipSizeY = val(Text_ChipSizeY)
    
    If Unit = 0 Then
        Range = (WaferSize * InchUnit) * 10
    Else
        Range = WaferSizemm
    End If
    RangeHalf = Range / 2
    RangeHalfX = RangeHalf / ChipSizeX
    RangeHalfY = RangeHalf / ChipSizeY
      
    ChipHalfCountX = ChipCount(RangeHalfX)
    ChipCountX = ChipHalfCountX * 2
    
    ChipHalfCountY = ChipCount(RangeHalfY)
    ChipCountY = ChipHalfCountY * 2
    
    Text_ChipCountX = ChipCountX - 1
    Text_ChipCountY = ChipCountY - 1
    
    Command_ChipOk.Enabled = True
End Sub

Public Sub ChipOk()
    Dim Unit As Integer
    Dim WaferSize As Double, WaferSizemm As Double
    Dim InchUnit As Double
    Dim ChipSizeX As Double, ChipSizeY As Double
    Dim ChipSizeHalfX As Double, ChipSizeHalfY As Double
    Dim CenterChipX As Integer, CenterChipY As Integer
    Dim Range As Double, RangeHalf As Double
    Dim ChipCountX As Integer, ChipCountY As Integer
    Dim ChipHalfCountX As Integer, ChipHalfCountY As Integer
    Dim D As Double
    
    If Option_WaferSize(0).value Then
        Unit = 0
    Else
        Unit = 1
    End If
    WaferSize = val(Text_WaferSize)
    WaferSizemm = val(Text_WaferSizemm)
    InchUnit = val(Text_Inch)
    
    ChipSizeX = val(Text_ChipSizeX)
    ChipSizeY = val(Text_ChipSizeY)
    ChipSizeHalfX = ChipSizeX / 2
    ChipSizeHalfY = ChipSizeY / 2
    
    If Unit = 0 Then
        Range = (WaferSize * InchUnit) * 10
    Else
        Range = WaferSizemm
    End If
    RangeHalf = Range / 2
    
    ChipCountX = ChipCount(Range / ChipSizeX)
    ChipCountY = ChipCount(Range / ChipSizeY)
    
    Text_ChipCountX = ChipCountX
    Text_ChipCountY = ChipCountY
    
    CenterChipX = ChipCount(((Range / 2) / ChipSizeX))
    CenterChipY = ChipCount(((Range / 2) / ChipSizeY))
    
    Erase Wafer
    Erase WaferTest
    
    Dim forx As Double, fory As Double
    Dim i As Integer, j As Integer
    Dim LineRange As Double, LineRangeChip As Double
    Dim LineChipCount As Double
    
    i = 0
    
    For fory = (RangeHalf - ChipSizeHalfY) To ChipSizeHalfY Step (ChipSizeY * -1)
        LineRange = Sqr((RangeHalf ^ 2) - (fory ^ 2))
        j = 0
        For forx = ChipSizeHalfX To LineRange Step ChipSizeX
            j = j + 1

            Wafer((CenterChipX - j) - 1, i).Chip = True
            Wafer((CenterChipX + j) - 1, i).Chip = True
            Wafer((CenterChipX - j) - 1, (CenterChipY + (CenterChipY - i - 2))).Chip = True
            Wafer((CenterChipX + j) - 1, (CenterChipY + (CenterChipY - i - 2))).Chip = True
            
            WaferTest((CenterChipX - j) - 1, i) = IIf(Int(Rnd(100) * 10) = 0, False, True)
            WaferTest((CenterChipX + j) - 1, i) = IIf(Int(Rnd(100) * 10) = 0, False, True)
            WaferTest((CenterChipX - j) - 1, (CenterChipY + (CenterChipY - i - 2))) = IIf(Int(Rnd(100) * 10) = 0, False, True)
            WaferTest((CenterChipX + j) - 1, (CenterChipY + (CenterChipY - i - 2))) = IIf(Int(Rnd(100) * 10) = 0, False, True)
            
            If StarProbe.EdgeChipmm > 0 Then
                Wafer((CenterChipX - j) - 1, i).ChipSkipDie = True
                Wafer((CenterChipX + j) - 1, i).ChipSkipDie = True
                Wafer((CenterChipX - j) - 1, (CenterChipY + (CenterChipY - i - 2))).ChipSkipDie = True
                Wafer((CenterChipX + j) - 1, (CenterChipY + (CenterChipY - i - 2))).ChipSkipDie = True
            End If
        Next
        i = i + 1
    Next

    i = 0
    
    For forx = (RangeHalf - ChipSizeHalfX) To ChipSizeHalfX Step (ChipSizeX * -1)
        Wafer(i, CenterChipY - 1).Chip = True
        WaferTest(i, CenterChipY - 1) = IIf(Int(Rnd(100) * 10) = 0, False, True)
        Wafer(CenterChipX + (CenterChipX - i - 2), CenterChipY - 1).Chip = True
        WaferTest(CenterChipX + (CenterChipX - i - 2), CenterChipY - 1) = IIf(Int(Rnd(100) * 10) = 0, False, True)
        
        If (RangeHalf - StarProbe.EdgeChipmm) <= forx Then
            Wafer(i, CenterChipY - 1).ChipSkipDie = True
            Wafer(CenterChipX + (CenterChipX - i - 2), CenterChipY - 1).ChipSkipDie = True
        Else
            Wafer(i, CenterChipY - 1).ChipSkipDie = False
            Wafer(CenterChipX + (CenterChipX - i - 2), CenterChipY - 1).ChipSkipDie = False
        End If
        i = i + 1
    Next

    j = 0
    For fory = (RangeHalf - ChipSizeHalfY) To ChipSizeHalfY Step (ChipSizeY * -1)
        Wafer(CenterChipX - 1, j).Chip = True
        WaferTest(CenterChipX - 1, j) = IIf(Int(Rnd(100) * 10) = 0, False, True)
        Wafer(CenterChipX - 1, CenterChipY + (CenterChipY - j - 2)).Chip = True
        WaferTest(CenterChipX - 1, CenterChipY + (CenterChipY - j - 2)) = IIf(Int(Rnd(100) * 10) = 0, False, True)
        
        If (RangeHalf - StarProbe.EdgeChipmm) <= fory Then
            Wafer(CenterChipX - 1, j).ChipSkipDie = True
            Wafer(CenterChipX - 1, CenterChipY + (CenterChipY - j - 2)).ChipSkipDie = True
        Else
            Wafer(CenterChipX - 1, j).ChipSkipDie = False
            Wafer(CenterChipX - 1, CenterChipY + (CenterChipY - j - 2)).ChipSkipDie = False
        End If
        j = j + 1
    Next

    Wafer(CenterChipX - 1, CenterChipY - 1).Chip = True
    WaferTest(CenterChipX - 1, CenterChipY - 1) = IIf(Int(Rnd(100) * 10) = 0, False, True)
    Wafer(CenterChipX - 1, CenterChipY - 1).ChipSkipDie = False
    
    If StarProbe.EdgeChipmm > 0 Then
        i = 0
        For fory = ChipSizeHalfY To (RangeHalf - StarProbe.EdgeChipmm) Step ChipSizeY
            LineRange = Sqr(((RangeHalf - StarProbe.EdgeChipmm) ^ 2) - (fory ^ 2))
            i = i + 1
            j = 0
            For forx = ChipSizeHalfX To LineRange Step ChipSizeX
                j = j + 1
                Wafer((CenterChipX - j) - 1, (CenterChipY - i) - 1).Chip = True
                Wafer((CenterChipX + j) - 1, (CenterChipY - i) - 1).Chip = True
                Wafer((CenterChipX - j) - 1, (CenterChipY + i) - 1).Chip = True
                Wafer((CenterChipX + j) - 1, (CenterChipY + i) - 1).Chip = True
    
                Wafer((CenterChipX - j) - 1, (CenterChipY - i) - 1).ChipSkipDie = False
                Wafer((CenterChipX + j) - 1, (CenterChipY - i) - 1).ChipSkipDie = False
                Wafer((CenterChipX - j) - 1, (CenterChipY + i) - 1).ChipSkipDie = False
                Wafer((CenterChipX + j) - 1, (CenterChipY + i) - 1).ChipSkipDie = False
            Next
        Next
    End If
    
    If StarProbe.PlateZone > 0 Then
        D = Sqr((RangeHalf ^ 2) - ((StarProbe.PlateZone / 2) ^ 2))
        D = RangeHalf - D
        D = D / ChipSizeY
        D = ChipCount(D)
        
        For fory = (ChipCountY - D) To ChipCountY
            For forx = 0 To ChipCountX
                If Wafer(forx, fory).Chip Then
                    Wafer(forx, fory).ChipPlate = True
                End If
            Next
        Next
    End If
    
    Dim bStartChip As Boolean
    Dim DisplayChipColor As Long

    pWaferDraw.Cls

    StarProbe.StartChip.x = 0
    StarProbe.StartChip.y = 0

    StarProbe.Max.x = 0
    StarProbe.Max.y = 0

    StarProbe.Min.x = 0
    StarProbe.Min.y = 0

    StarProbe.ChipCountX = ChipCountX - 1
    StarProbe.ChipCountY = ChipCountY - 1

    bStartChip = True

    For fory = 0 To ChipCountY
        For forx = 0 To ChipCountX
            If Wafer(forx, fory).Chip Then
                If bStartChip Then
                    StarProbe.StartChip.x = forx
                    StarProbe.StartChip.y = fory
                    bStartChip = False
                End If

                If Wafer(forx, fory).ChipMask Then
                    DisplayChipColor = ChipColor(1)
                ElseIf Wafer(forx, fory).ChipPlate Then
                    DisplayChipColor = ChipColor(4)
                ElseIf Wafer(forx, fory).ChipSkipDie Then
                    If Wafer(forx, fory).ChipInk = True And Wafer(forx, fory).ChipInk2 = False Then
                        DisplayChipColor = ChipColor(5)
                    ElseIf Wafer(forx, fory).ChipInk = True And Wafer(forx, fory).ChipInk2 = True Then
                        DisplayChipColor = ChipColor(6)
                    Else
                        DisplayChipColor = ChipColor(3)
                    End If
                Else
                    DisplayChipColor = ChipColor(0)
                End If
                pWaferDraw.Line (forx, fory)-(forx, fory), DisplayChipColor, BF
            End If
        Next
    Next

    StarProbe.Max.x = StarProbe.ChipCountX - StarProbe.StartChip.x
    StarProbe.Max.y = StarProbe.ChipCountY - StarProbe.StartChip.y

    StarProbe.Min.x = StarProbe.Max.x - StarProbe.ChipCountX
    StarProbe.Min.y = StarProbe.Max.y - StarProbe.ChipCountY
    
    StarProbe.CenterChipX = CenterChipX
    StarProbe.CenterChipY = CenterChipY
    
    StarProbe.StartChip.x = CenterChipX - 1
    StarProbe.StartChip.y = CenterChipY - 1
End Sub

Private Sub Text_WaferSizemm_Change()
    Call WaferInformationChange
End Sub
