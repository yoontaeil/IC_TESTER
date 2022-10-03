VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form Form_clean_tip 
   BorderStyle     =   1  '단일 고정
   Caption         =   "Clean Tip"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3855
   Icon            =   "Form_clean_tip.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   3855
   StartUpPosition =   1  '소유자 가운데
   Begin VB.Frame Frame9 
      Height          =   2175
      Left            =   120
      TabIndex        =   44
      Top             =   120
      Width           =   3615
      Begin VB.TextBox txt_Z 
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
         Left            =   600
         TabIndex        =   51
         Text            =   "1"
         Top             =   1695
         Width           =   975
      End
      Begin VB.TextBox txt_Y 
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
         Left            =   600
         TabIndex        =   50
         Text            =   "1"
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txt_X 
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
         Left            =   600
         TabIndex        =   49
         Text            =   "1"
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Tip Clean진행"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   48
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "위치값 적용"
         Height          =   375
         Left            =   1680
         TabIndex        =   47
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "X,Y값 지정"
         Height          =   375
         Left            =   1680
         TabIndex        =   46
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton Command23 
         Caption         =   "Move (X,Y)"
         Height          =   375
         Left            =   1680
         TabIndex        =   45
         Top             =   2760
         Width           =   1695
      End
      Begin Threed.SSPanel SSPanel7 
         Height          =   375
         Left            =   120
         TabIndex        =   52
         Top             =   240
         Width           =   3255
         _Version        =   65536
         _ExtentX        =   5741
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "CLEAN TIP POSITION"
         ForeColor       =   4210752
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FloodColor      =   12582912
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
         Index           =   4
         Left            =   240
         TabIndex        =   55
         Top             =   1290
         Width           =   495
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
         Index           =   0
         Left            =   240
         TabIndex        =   54
         Top             =   825
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
         Index           =   10
         Left            =   240
         TabIndex        =   53
         Top             =   1800
         Width           =   495
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "SAVE "
      Height          =   495
      Left            =   120
      TabIndex        =   43
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "CLOSE"
      Height          =   495
      Left            =   2040
      TabIndex        =   42
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "JOG POSITION"
      Height          =   3495
      Left            =   5520
      TabIndex        =   14
      Top             =   4200
      Width           =   8655
      Begin VB.CommandButton Command6 
         Caption         =   "X-"
         Height          =   375
         Left            =   120
         TabIndex        =   37
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton Command7 
         Caption         =   "X+"
         Height          =   375
         Left            =   1560
         TabIndex        =   36
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Y-"
         Height          =   375
         Left            =   840
         TabIndex        =   35
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Y+"
         Height          =   375
         Left            =   840
         TabIndex        =   34
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton Command13 
         Caption         =   "POS"
         Height          =   375
         Left            =   600
         TabIndex        =   33
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1320
         TabIndex        =   32
         Top             =   2160
         Width           =   855
      End
      Begin VB.CommandButton Command12 
         Caption         =   "MOVE"
         Height          =   375
         Left            =   600
         TabIndex        =   31
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1320
         TabIndex        =   30
         Top             =   2520
         Width           =   855
      End
      Begin VB.CommandButton Command14 
         Caption         =   "XY GO"
         Height          =   495
         Left            =   240
         TabIndex        =   29
         Top             =   1440
         Width           =   1935
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Move wafer center"
         Height          =   375
         Left            =   2400
         TabIndex        =   28
         Top             =   360
         Width           =   2055
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Move home position"
         Height          =   375
         Left            =   2400
         TabIndex        =   27
         Top             =   840
         Width           =   2055
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Z down"
         Height          =   375
         Left            =   2400
         TabIndex        =   26
         Top             =   1320
         Width           =   2055
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Z up"
         Height          =   375
         Left            =   2400
         TabIndex        =   25
         Top             =   1800
         Width           =   2055
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   6360
         TabIndex        =   24
         Text            =   "4000"
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Set"
         Height          =   375
         Left            =   6960
         TabIndex        =   23
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   6360
         TabIndex        =   22
         Text            =   "2000"
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Set"
         Height          =   375
         Left            =   6960
         TabIndex        =   21
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   6360
         TabIndex        =   20
         Text            =   "2000"
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton Command19 
         Caption         =   "Set"
         Height          =   375
         Left            =   6960
         TabIndex        =   19
         Top             =   1320
         Width           =   735
      End
      Begin VB.CommandButton Command20 
         Caption         =   "Request Miscellaneous Position Data"
         Height          =   375
         Left            =   4680
         TabIndex        =   18
         Top             =   1800
         Width           =   3855
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   4680
         TabIndex        =   17
         Top             =   2160
         Width           =   3855
      End
      Begin VB.CommandButton Command21 
         Caption         =   "Request Various Wafer Position Data"
         Height          =   375
         Left            =   4680
         TabIndex        =   16
         Top             =   2640
         Width           =   3855
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   4680
         TabIndex        =   15
         Top             =   3000
         Width           =   3855
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
         Index           =   3
         Left            =   240
         TabIndex        =   41
         Top             =   2280
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Z Up limit : "
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
         Left            =   4680
         TabIndex        =   40
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Z Down limit : "
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
         Left            =   4680
         TabIndex        =   39
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Z undertravel :"
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
         Index           =   8
         Left            =   4680
         TabIndex        =   38
         Top             =   1440
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2175
      Left            =   120
      TabIndex        =   4
      Top             =   4320
      Width           =   3615
      Begin VB.CommandButton Command22 
         Caption         =   "Get position"
         Height          =   375
         Left            =   1680
         TabIndex        =   9
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txt_CamX 
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
         Left            =   600
         TabIndex        =   8
         Text            =   "1"
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txt_CamY 
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
         Left            =   600
         TabIndex        =   7
         Text            =   "1"
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txt_CamZ 
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
         Left            =   600
         TabIndex        =   6
         Text            =   "1"
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton Command24 
         Caption         =   "Move Pos"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   5
         Top             =   1200
         Width           =   1695
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   3255
         _Version        =   65536
         _ExtentX        =   5741
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "CAMERA POSITION"
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
         Index           =   1
         Left            =   240
         TabIndex        =   13
         Top             =   825
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
         Index           =   2
         Left            =   240
         TabIndex        =   12
         Top             =   1290
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
         Index           =   9
         Left            =   240
         TabIndex        =   11
         Top             =   1770
         Width           =   495
      End
   End
   Begin VB.TextBox Text8 
      Enabled         =   0   'False
      Height          =   270
      Left            =   3960
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox Text9 
      Enabled         =   0   'False
      Height          =   270
      Left            =   3960
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text10 
      Enabled         =   0   'False
      Height          =   270
      Left            =   3960
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command25 
      Caption         =   "Read Position"
      Height          =   1695
      Left            =   5280
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Current X : "
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
      Left            =   3960
      TabIndex        =   58
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Current Y : "
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
      Index           =   11
      Left            =   3960
      TabIndex        =   57
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Current Z : "
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
      Index           =   12
      Left            =   3960
      TabIndex        =   56
      Top             =   1440
      Width           =   1215
   End
End
Attribute VB_Name = "Form_clean_tip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===================================================================
'[ 2022.06.28 ] : tip clean관련 수정.
'===================================================================
' - 참고사항 -
'설정 카운트마다 tip clean을 실시한다.
'tip clean기능 on/off가능하도록 한다.
'설정 x,y,z값을 저장한다. shutdown시 적용한다.
' - 주요 명령어 -
' ?H, MM 명령은 motor 좌표를 나타낸다. - MM좌표 명령으로 어떻게 동작하는지 확인 이 필요하다.
' ?P, MO 명령은 first die기준의 좌표를 나타낸다.
' Z축관련 ZM와 SP9명령의 차이를 시험해본다.

'최초 설정순서
'1) wafer profile
'2) first die set
'3) move clean tip position
'4) clean tip Z set
'5) Starprobe PC에서 get position
'6) 2001X에서 tipclean 내용 저장

'2001X reboot시 설정
'1) wafer profile
'2) first die set
'3) set position
'===================================================================

'[ Clean tip : Get position ]
Private Sub Command1_Click()
    If Move_select = False Then
        '===============================================================
        '1) Probe card 좌표            ?P, MOX0Y0 : die step을 읽어온다.
        '===============================================================
        StarProbe.Tipclean_X = StarProbe_Step_X_Value
        txt_X.Text = StarProbe.Tipclean_X       'X
    
        StarProbe.Tipclean_Y = Starprobe_Step_Y_Value
        txt_Y.Text = StarProbe.Tipclean_Y       'Y
    Else
        '===============================================================
        '2) Motor 좌표                 ?H MMX0Y0 : motor위치를 읽어온다.
        '===============================================================
        StarProbe.Tipclean_X = StarProbe_Motor_X_Value
        txt_X.Text = StarProbe.Tipclean_X       'X
    
        StarProbe.Tipclean_Y = Starprobe_motor_Y_Value
        txt_Y.Text = StarProbe.Tipclean_Y       'Y
    End If
'    StarProbe.Tipclean_Z = StarProbe_Motor_Z_Value
'    txt_Z.Text = StarProbe.Tipclean_Z       'Z
End Sub

'[ Clean tip : Set position ]
Private Sub Command2_Click()
    Dim Z_backup As String
    StarProbe_Z_UP
    Z_backup = StarProbe_Motor_Z_Value                                                              '1.현재의 Z값을 가져와서 임시로 저장한다.
    
    StarProbe_Z_Down                                                                                '2.Z down
    
    If Move_select = False Then
        '1) Probe card위치  MOX0Y0
        Call StarProbe_XY_Moving(val(StarProbe.Tipclean_X), val(StarProbe.Tipclean_Y))              '3. XY move : Die step만큼 좌표를 이동한다.
    Else
        '2) Motor 위치 MMX0Y0
        x = StarProbe_Motor_X_Value
        y = Starprobe_motor_Y_Value
        Call StarProbe_Pluse_Move(val(StarProbe.Tipclean_X - x), val(StarProbe.Tipclean_Y - y))     '3. XY move : 설정좌표-현재좌표해서 구해진 만큼 이동한다.
    End If
    If Not StarProbe_Motor_End_check Then   'Move OK
        StarProbe_Z_UP                                                                              '4. Z up
        StarProbe_Move_Z (txt_Z.Text)                                                               '5. Tip clean Z move
        StarProbe_Scrub_Position                                                                    '6. Clean position set (PN)
        StarProbe_Move_Z (Z_backup)                                                                '7. Z축 원래값으로 복구
    End If
    StarProbe_Z_Down                                                                                '8. Z down
    Call StarProbe_Pluse_Move(val(StarProbe.Tipclean_X - x) * -1, val(StarProbe.Tipclean_Y - y) * -1) '3. XY move : 설정좌표-현재좌표해서 구해진 만큼 이동한다.
    If Not StarProbe_Motor_End_check Then   'Move OK
    End If
End Sub

Private Sub Command25_Click()
    Text8.Text = StarProbe_Motor_X_Value
    Text9.Text = Starprobe_motor_Y_Value
    Text10.Text = StarProbe_Motor_Z_Value
End Sub

'[ Run clean tip ]
Private Sub Command3_Click()
    Call StarProbe_tip_clean                                                            'CP명령어
'    StarProbe_LogSave (2)                       '[ 2022.07.26 ] : log
    Sleep 500
End Sub

'[ SAVE ]
Private Sub Command4_Click()
    StarProbe.Tipclean_X = txt_X
    StarProbe.Tipclean_Y = txt_Y
    StarProbe.Tipclean_Z = txt_Z
    
    StarProbe.Cam_X = txt_CamX
    StarProbe.Cam_Y = txt_CamY
    StarProbe.Cam_Z = txt_CamZ
    
    Call StarProbe_FileSave_SystemInfo
End Sub

'[ CLOSE ]
Private Sub Command5_Click()
    Unload Me
End Sub






'기타동작
'[ X,Y Move ]
Private Sub Command23_Click()
    Dim X_val As Integer
    Dim y_val As Integer
    
    StarProbe_Z_Down                                                                    ' 1.Z down
    If Move_select = False Then
        '1) Probe card위치  MOX0Y0
        Call StarProbe_XY_Moving(val(StarProbe.Tipclean_X), val(StarProbe.Tipclean_Y))
    Else
        X_val = StarProbe_Motor_X_Value                 '현재 X좌표
        y_val = Starprobe_motor_Y_Value                 '현재 Y좌표
        Text10.Text = StarProbe_Motor_Z_Value           '현재 Z좌표
        
        '2) Motor 위치      MMX0Y0 : MM명령은 현재 좌표를기준으로 motor setp을 움직이는 명령이다.
        Call StarProbe_Pluse_Move(val(StarProbe.Tipclean_X - X_val), val(StarProbe.Tipclean_Y - y_val))     'clean position - 현재 좌표
    End If
    If Not StarProbe_Motor_End_check Then
    End If
End Sub


'[ Z down ]
Private Sub Command15_Click()
    StarProbe_Z_Down
End Sub

'[ Z up ]
Private Sub Command16_Click()
    StarProbe_Z_UP
End Sub

'[ Z uplimit ]
Private Sub Command17_Click()
    Call StarProbe_Z_Up_Limit(Text3.Text)
End Sub



Private Sub Command20_Click()
    Text6.Text = StarProbe_Miscellaneous_data
End Sub

Private Sub Command21_Click()
    Text7.Text = StarProbe_Various_Position
End Sub

'[ Camera position : Get position ]
Private Sub Command22_Click()
    StarProbe.Cam_X = StarProbe_Step_X_Value
    txt_CamX.Text = StarProbe.Cam_X        'X
        
    StarProbe.Cam_Y = Starprobe_Step_Y_Value
    txt_CamY.Text = StarProbe.Cam_Y        'Y
        
    StarProbe.Cam_Z = StarProbe_Motor_Z_Value
    txt_CamZ.Text = StarProbe.Cam_Z        'Z
End Sub


'[ Camera position : Move position ]
Private Sub Command24_Click()
    Call StarProbe_Pluse_Move(val(StarProbe.Cam_X), val(StarProbe.Cam_Y))
    If Not StarProbe_Motor_End_check Then
    End If
End Sub

'[ X- ]
Private Sub Command6_Click()
    txt_X.Text = txt_X - 1
    Call StarProbe_Pluse_Move(txt_X.Text, txt_Y.Text)
    If Not StarProbe_Motor_End_check Then
    End If
End Sub

'[ X+ ]
Private Sub Command7_Click()
    txt_X.Text = txt_X + 1
    Call StarProbe_Pluse_Move(txt_X.Text, txt_Y.Text)
    If Not StarProbe_Motor_End_check Then
    End If
End Sub

'[ Y- ]
Private Sub Command8_Click()
    txt_Y.Text = txt_Y - 1
    Call StarProbe_Pluse_Move(txt_X.Text, txt_Y.Text)
    If Not StarProbe_Motor_End_check Then
    End If
End Sub

'[ Y+ ]
Private Sub Command9_Click()
    txt_Y.Text = txt_Y + 1
    Call StarProbe_Pluse_Move(txt_X.Text, txt_Y.Text)
    If Not StarProbe_Motor_End_check Then
    End If
End Sub

'[ Move wafer center ]
Private Sub Command10_Click()
    StarProbe_Tip_center
End Sub

'[ Move home position ]
Private Sub Command11_Click()
    StarProbe_Motor_Home
End Sub

'[ ZM ]
Private Sub Command12_Click()
    StarProbe_Move_Z (Text1.Text)
End Sub

'[ ?Z ]
Private Sub Command13_Click()
    Text2.Text = StarProbe_Motor_Z_Value
End Sub

'[ XY POS MOVE ]
Private Sub Command14_Click()
    Call StarProbe_Pluse_Move(val(StarProbe.Tipclean_X), val(StarProbe.Tipclean_Y))
    If Not StarProbe_Motor_End_check Then
    End If
End Sub

'[ X,Y,Z값을 가져온다. ]
Private Sub Form_Load()
    txt_X.Text = StarProbe.Tipclean_X
    txt_Y.Text = StarProbe.Tipclean_Y
    txt_Z.Text = StarProbe.Tipclean_Z
    
    txt_CamX.Text = StarProbe.Cam_X
    txt_CamY.Text = StarProbe.Cam_Y
    txt_CamZ.Text = StarProbe.Cam_Z
    
    Move_select = True
End Sub

Private Sub txt_CamX_Change()
    If IsNumeric(txt_CamX.Text) = False Then   'X입력값이 숫자인지 여부를 판단
        MsgBox "Invalid. This blank Inputed to Number !", vbExclamation, "Error"
        txt_CamX.Text = "0"
        Exit Sub
    End If
End Sub

Private Sub txt_CamY_Change()
    If IsNumeric(txt_CamY.Text) = False Then   'Y입력값이 숫자인지 여부를 판단
        MsgBox "Invalid. This blank Inputed to Number !", vbExclamation, "Error"
        txt_CamY.Text = "0"
        Exit Sub
    End If
End Sub

Private Sub txt_CamZ_Change()                  'Z입력값이 숫자인지 여부를 판단
    If IsNumeric(txt_CamZ.Text) = False Then
        MsgBox "Invalid. This blank Inputed to Number !", vbExclamation, "Error"
        txt_CamZ.Text = "200"
        Exit Sub
    End If

End Sub

Private Sub txt_X_Change()
    If IsNumeric(txt_X.Text) = False Then   'X입력값이 숫자인지 여부를 판단
        MsgBox "Invalid. This blank Inputed to Number !", vbExclamation, "Error"
        txt_X.Text = "0"
        Exit Sub
    End If
End Sub

Private Sub txt_Y_Change()
    If IsNumeric(txt_Y.Text) = False Then   'Y입력값이 숫자인지 여부를 판단
        MsgBox "Invalid. This blank Inputed to Number !", vbExclamation, "Error"
        txt_Y.Text = "0"
        Exit Sub
    End If
End Sub

Private Sub txt_Z_Change()                  'Z입력값이 숫자인지 여부를 판단
    If IsNumeric(txt_Z.Text) = False Then
        MsgBox "Invalid. This blank Inputed to Number !", vbExclamation, "Error"
        txt_Z.Text = "200"
        Exit Sub
    End If
End Sub

