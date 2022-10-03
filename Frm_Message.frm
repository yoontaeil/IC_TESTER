VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form Frm_Message 
   BorderStyle     =   1  '단일 고정
   Caption         =   "Check list"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7965
   Icon            =   "Frm_Message.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   7965
   StartUpPosition =   1  '소유자 가운데
   Begin VB.CommandButton Command1 
      Caption         =   "확인"
      Height          =   495
      Left            =   2880
      TabIndex        =   8
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox txt_X1 
      Enabled         =   0   'False
      Height          =   270
      Left            =   4920
      TabIndex        =   6
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox txt_Y1 
      Enabled         =   0   'False
      Height          =   270
      Left            =   4920
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txt_Z1 
      Enabled         =   0   'False
      Height          =   270
      Left            =   4920
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txt_X2 
      Enabled         =   0   'False
      Height          =   270
      Left            =   6600
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox txt_Y2 
      Enabled         =   0   'False
      Height          =   270
      Left            =   6600
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txt_Z2 
      Enabled         =   0   'False
      Height          =   270
      Left            =   6600
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   480
      Top             =   2400
   End
   Begin VB.CommandButton Command2 
      Caption         =   "이전으로"
      Height          =   495
      Left            =   4080
      TabIndex        =   0
      Top             =   2880
      Width           =   975
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   4335
      _Version        =   65536
      _ExtentX        =   7646
      _ExtentY        =   873
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Width           =   4335
      _Version        =   65536
      _ExtentX        =   7646
      _ExtentY        =   873
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
   End
   Begin Threed.SSPanel SSPanel3 
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   4335
      _Version        =   65536
      _ExtentX        =   7646
      _ExtentY        =   873
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
   End
   Begin Threed.SSPanel SSPanel4 
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   1560
      Width           =   4335
      _Version        =   65536
      _ExtentX        =   7646
      _ExtentY        =   873
      _StockProps     =   15
      BackColor       =   15790320
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
   End
   Begin Threed.SSPanel SSPanel50 
      Height          =   495
      Left            =   4800
      TabIndex        =   12
      Top             =   1560
      Width           =   3015
      _Version        =   65536
      _ExtentX        =   5318
      _ExtentY        =   873
      _StockProps     =   15
      Caption         =   "주의 : 2001X의 Main Menu상태에서 확인을 누르세요"
      BackColor       =   15790320
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
   End
   Begin Threed.SSPanel SSPanel5 
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   2160
      Width           =   4335
      _Version        =   65536
      _ExtentX        =   7646
      _ExtentY        =   873
      _StockProps     =   15
      BackColor       =   15790320
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
   End
   Begin VB.Label Label1 
      Caption         =   "X1"
      Height          =   255
      Index           =   0
      Left            =   4680
      TabIndex        =   19
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Y1"
      Height          =   255
      Index           =   1
      Left            =   4680
      TabIndex        =   18
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Z1"
      Height          =   255
      Index           =   2
      Left            =   4680
      TabIndex        =   17
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "X2"
      Height          =   255
      Index           =   3
      Left            =   6240
      TabIndex        =   16
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Y2"
      Height          =   255
      Index           =   4
      Left            =   6240
      TabIndex        =   15
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Z2"
      Height          =   255
      Index           =   5
      Left            =   6240
      TabIndex        =   14
      Top             =   960
      Width           =   375
   End
End
Attribute VB_Name = "Frm_Message"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'[ 2022.08.12 ] : 선택한 메시지를 확인하는 창
Dim flag As Boolean

Private Sub Command1_Click()
On Error GoTo err

    Dim xx As Integer
    Dim yy As Integer
    Dim zz As Integer
    
    '
    txt_X2.Text = StarProbe_Motor_X_Value
    txt_Y2.Text = Starprobe_motor_Y_Value
    txt_Z2.Text = StarProbe_Motor_Z_Value
    
    Needle_check_flag = False              '[ 2022.08.31 ]
    
    If SSPanel1.Caption = "침적 위치 변화를 확인하세요." And SSPanel2.Caption = "" Then       'x,y check
        If txt_X1.Text <> txt_X2.Text Or txt_Y1.Text <> txt_Y2.Text Then
            flag = True
        Else
            flag = False
        End If
    ElseIf SSPanel1.Caption = "침적 위치 변화를 확인하세요." And SSPanel2.Caption <> "" Then       'x,y,z check
        If (txt_X1.Text <> txt_X2.Text Or txt_Y1.Text <> txt_Y2.Text) And txt_Z1.Text <> txt_Z2.Text Then
            flag = True
        Else
            flag = False
        End If
    ElseIf SSPanel1.Caption = "침적 벗어남(불량처리)." Then
        flag = True
        Needle_check_flag = True
        Unload Me
        Unload Form_Check_List
    Else                                                            'z check
        If txt_Z1.Text <> txt_Z2.Text Then
            flag = True
        Else
            flag = False
        End If
    End If
    '
    
    xx = txt_X1.Text - txt_X2.Text
    yy = txt_Y1.Text - txt_Y2.Text
    zz = txt_Z1.Text - txt_Z2.Text
    
    If flag = False Then
        Exit Sub
    End If
    '[ 2022.07.20 ]
    MSG_DATA = "이동내역 : " & "X:" & xx & "." & "Y:" & yy & "." & "Z:" & zz
                               
    If LOG_FILE_ON = 1 Then SelectExt.Log_Data_Save (18)
    Unload Me
    Unload Form_Check_List
    
err:
    Resume Next
End Sub

Private Sub Command2_Click()
    Form_Check_List.Check1(0).value = 0
    Form_Check_List.Check1(1).value = 0
    Form_Check_List.Check1(2).value = 0
    Form_Check_List.Check1(3).value = 0
    Form_Check_List.Check1(5).value = 0
    
    Unload Me
End Sub

Private Sub Form_Load()
    MSG_DATA = ""
    RemoveCancelMenuItem Me '종료"X"버튼을 사용하지 못하게 한다.
    flag = False
        
    txt_X1.Text = StarProbe_Motor_X_Value
    txt_Y1.Text = Starprobe_motor_Y_Value
    txt_Z1.Text = StarProbe_Motor_Z_Value
End Sub

Private Sub Timer1_Timer()
    If SSPanel50.ForeColor = vbBlack Then
        SSPanel50.ForeColor = vbRed
    Else
        SSPanel50.ForeColor = vbBlack
    End If

    DoEvents
End Sub

