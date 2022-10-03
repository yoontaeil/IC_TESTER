VERSION 5.00
Begin VB.Form Form_Check_List 
   BorderStyle     =   1  '단일 고정
   Caption         =   "Check List"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4800
   Icon            =   "Form_Check_List.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   4800
   StartUpPosition =   2  '화면 가운데
   Begin VB.CheckBox Check1 
      Caption         =   "5. 침적 벗어남(불량 처리)"
      Height          =   375
      Index           =   5
      Left            =   240
      TabIndex        =   13
      Top             =   2040
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "확인"
      Height          =   495
      Left            =   3600
      TabIndex        =   12
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox txt_Z2 
      Height          =   270
      Left            =   6240
      TabIndex        =   11
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox txt_Z1 
      Height          =   270
      Left            =   6240
      TabIndex        =   10
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox txt_Y2 
      Height          =   270
      Left            =   5400
      TabIndex        =   9
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox txt_X2 
      Height          =   270
      Left            =   5040
      TabIndex        =   8
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox txt_Y1 
      Height          =   270
      Left            =   5400
      TabIndex        =   7
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox txt_X1 
      Height          =   270
      Left            =   5040
      TabIndex        =   6
      Top             =   480
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3960
      Top             =   4920
   End
   Begin VB.CommandButton Command1 
      Caption         =   "확인"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   5040
      Width           =   1815
   End
   Begin VB.CheckBox Check1 
      Caption         =   "6.작업진행"
      Height          =   375
      Index           =   4
      Left            =   240
      TabIndex        =   4
      Top             =   2400
      Width           =   3255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "4. 편마모"
      Height          =   375
      Index           =   3
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   3255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "3. Probe Card마모(침적 크기 (大))"
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   3255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "2. 침적 세기(Chuck 높이)"
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   3255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "1. 침적 위치 변화(X,Y)"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Wafer 외관불량 방지를 위해 침적을 확인하세요."
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "Form_Check_List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'[ 2022.08.30 ] : 침적확인 방법
'===============================================================================================================================
'[ 확인순서 및 내용 ]
'===============================================================================================================================
'1.침적 위치변화            : x,y 좌표의 이동이 발생하면 된다.
'2.침적세기(chuck높이)      : z측의 높이가 변경이 되면 발생한다.
'3.침적세기(probe card)     : z측의 높이가 변경이 되면 발생한다.
'4.편마모                   : 확인
'5.작업진행                 : 작업자가 체크를 하면 현재 창이 종료 되면서 auto작업을 이어간다.
'===============================================================================================================================

Private Sub Check1_Click(Index As Integer)
    Select Case Index
        Case 4:                                                                                     '[ 6.작업 진행 ]
            If CHK_CANCEL = False Then
                x = MsgBox("체크한 항목의 검사를 실시하엿습니까?", vbQuestion + vbYesNo, "검사확인")
                If x = 7 Then
                    CHK_CANCEL = True
                    Check1(Index).value = 0
                    Exit Sub
                End If
                CHK_CANCEL = False
                Unload Me
            Else
                CHK_CANCEL = False
            End If
    End Select
End Sub

Private Sub Command2_Click()
    MSG_DATA = ""
    RemoveCancelMenuItem Me '종료"X"버튼을 사용하지 못하게 한다.
    flag = False
    
    If Form_Check_List.Check1(0).value = 1 Then
        Frm_Message.SSPanel1.Caption = "침적 위치 변화를 확인하세요."
        '[ 2022.07.20 ]
        If LOG_FILE_ON = 1 Then SelectExt.Log_Data_Save (14)
    End If
    If Form_Check_List.Check1(1).value = 1 Then
        If Frm_Message.SSPanel1.Caption = "" Then
            Frm_Message.SSPanel1.Caption = "침적 세기(Chuck 높이)를 확인하세요."
        Else
            Frm_Message.SSPanel2.Caption = "침적 세기(Chuck 높이)를 확인하세요."
        End If
        '[ 2022.07.20 ]
        If LOG_FILE_ON = 1 Then SelectExt.Log_Data_Save (15)
    End If
    If Form_Check_List.Check1(2).value = 1 Then
        If Frm_Message.SSPanel1.Caption = "" Then
            Frm_Message.SSPanel1.Caption = "침적 세기(Probe Card)를 확인하세요."
        ElseIf Frm_Message.SSPanel2.Caption = "" Then
            Frm_Message.SSPanel2.Caption = "침적 세기(Probe Card)를 확인하세요."
        Else
            Frm_Message.SSPanel3.Caption = "침적 세기(Probe Card)를 확인하세요."
        End If
        '[ 2022.07.20 ]
        If LOG_FILE_ON = 1 Then SelectExt.Log_Data_Save (16)
    End If
    If Form_Check_List.Check1(3).value = 1 Then
        If Frm_Message.SSPanel1.Caption = "" Then
            Frm_Message.SSPanel1.Caption = "편마모가 발생 되었으니, Probe Card를 교체 하세요."
        ElseIf Frm_Message.SSPanel2.Caption = "" Then
            Frm_Message.SSPanel2.Caption = "편마모가 발생 되었으니, Probe Card를 교체 하세요."
        ElseIf Frm_Message.SSPanel3.Caption = "" Then
            Frm_Message.SSPanel3.Caption = "편마모가 발생 되었으니, Probe Card를 교체 하세요."
        ElseIf Frm_Message.SSPanel4.Caption = "" Then
            Frm_Message.SSPanel4.Caption = "편마모가 발생 되었으니, Probe Card를 교체 하세요."
        End If
        '[ 2022.07.20 ]
        If LOG_FILE_ON = 1 Then SelectExt.Log_Data_Save (17)
    End If
    
    If Form_Check_List.Check1(5).value = 1 Then
        If Frm_Message.SSPanel1.Caption = "" Then
            Frm_Message.SSPanel1.Caption = "침적 벗어남(불량처리)."
        ElseIf Frm_Message.SSPanel2.Caption = "" Then
            Frm_Message.SSPanel2.Caption = "침적 벗어남(불량처리)."
        ElseIf Frm_Message.SSPanel3.Caption = "" Then
            Frm_Message.SSPanel3.Caption = "침적 벗어남(불량처리)."
        ElseIf Frm_Message.SSPanel4.Caption = "" Then
            Frm_Message.SSPanel4.Caption = "침적 벗어남(불량처리)."
        ElseIf Frm_Message.SSPanel5.Caption = "" Then
            Frm_Message.SSPanel5.Caption = "침적 벗어남(불량처리)."
        End If
        '[ 2022.07.20 ]
        If LOG_FILE_ON = 1 Then SelectExt.Log_Data_Save (20)
    End If
    Frm_Message.Show 1
End Sub

Private Sub Form_Load()
    CHK_CANCEL = False
    
    RemoveCancelMenuItem Me '종료"X"버튼을 사용하지 못하게 한다.
    
    '[ 2022.07.20 ]
    If LOG_FILE_ON = 1 Then SelectExt.Log_Data_Save (6)
End Sub



