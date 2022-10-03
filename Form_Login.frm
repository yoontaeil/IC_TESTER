VERSION 5.00
Begin VB.Form Form_Login 
   Caption         =   "Log In"
   ClientHeight    =   2295
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   2445
   Icon            =   "Form_Login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   2445
   StartUpPosition =   2  '화면 가운데
   Begin VB.Frame Frame1 
      Caption         =   "Select"
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2175
      Begin VB.TextBox Text1 
         Alignment       =   2  '가운데 맞춤
         Height          =   270
         IMEMode         =   3  '사용 못함
         Left            =   240
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1200
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Operator Mode"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Engineer Mode"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   1815
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   2175
   End
End
Attribute VB_Name = "Form_Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim model As String
    
    Select Case Model_Select
        Case 0:                                                 'not use
            If Tester_Select = 0 Then
                model = "SPC-2001X(EAGLE-2001X)"
            Else
                model = "SPC-2001X(AMT-88)"
            End If
        Case 1:                                                 'ACCO
            model = "SPC-2001X(ACCO)"
        Case 2:                                                 'AMT-88
            model = "SPC-2001X(AMT-88)"
        Case 3:                                                 'EAGLE(2001X)
            model = "SPC-2001X(EAGLE-2001X)"
        Case 4:                                                 'EAGLE(E4090)
            model = "SPC-2001X(EAGLE-E4090)"
    End Select

    If Option1(0).value = True Then
        MT2000.Caption = model & " : OPERATOR MODE"
        Mode_Set = False                'operator mode
        '[ 2022.07.20 ]
        If LOG_FILE_ON = 1 Then SelectExt.Log_Data_Save (2)
        MT2000.Show
        Unload Me
        Exit Sub
    Else
        If Text1 <> "starprobe1!" Then
            MT2000.Caption = model & " : OPERATOR MODE"
            Mode_Set = False            'operator mode
            MsgBox "Wrong Master password!!", 16, "Not match"
            Text1.SetFocus
        Else
            Mode_Set = True             'engineer mode
            MT2000.Caption = model & " : ENGINEER MODE"
            '[ 2022.07.20 ]
            If LOG_FILE_ON = 1 Then SelectExt.Log_Data_Save (1)
            MT2000.Show
            Unload Me
            Exit Sub
        End If
    End If
End Sub

Private Sub Option1_Click(Index As Integer)
    If Option1(0).value = True Then
        Text1.Visible = False
    Else
        Text1.Visible = True
        Text1.SetFocus
    End If
End Sub
