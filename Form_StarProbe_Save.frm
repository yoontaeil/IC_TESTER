VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form_StarProbe_Save 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "Map Image & Summary Report"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8550
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   8550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.CommandButton Command_Save 
      BackColor       =   &H00C0C0C0&
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
      Height          =   375
      Index           =   2
      Left            =   8040
      Picture         =   "Form_StarProbe_Save.frx":0000
      Style           =   1  '그래픽
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command_Save 
      BackColor       =   &H00C0C0C0&
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
      Height          =   375
      Index           =   1
      Left            =   8040
      Picture         =   "Form_StarProbe_Save.frx":0386
      Style           =   1  '그래픽
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   600
      Width           =   375
   End
   Begin VB.CommandButton Command_Save 
      BackColor       =   &H00C0C0C0&
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
      Height          =   375
      Index           =   0
      Left            =   8040
      Picture         =   "Form_StarProbe_Save.frx":070C
      Style           =   1  '그래픽
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox Text_FileName 
      Enabled         =   0   'False
      Height          =   375
      Index           =   2
      Left            =   2040
      TabIndex        =   7
      Top             =   1080
      Width           =   5895
   End
   Begin VB.TextBox Text_FileName 
      Enabled         =   0   'False
      Height          =   375
      Index           =   1
      Left            =   2040
      TabIndex        =   6
      Top             =   600
      Width           =   5895
   End
   Begin VB.TextBox Text_FileName 
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   2040
      TabIndex        =   5
      Top             =   120
      Width           =   5895
   End
   Begin VB.CheckBox Check_Save 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Summary Report"
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CheckBox Check_Save 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Map (Original)"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1935
   End
   Begin VB.CheckBox Check_Save 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Map Image"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton Command_Cancel 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Cancel"
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
      Left            =   4320
      Style           =   1  '그래픽
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command_Ok 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Ok"
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
      Left            =   3000
      Style           =   1  '그래픽
      TabIndex        =   0
      Top             =   1680
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Form_StarProbe_Save"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check_Save_Click(Index As Integer)
    Text_FileName(Index).Enabled = Check_Save(Index).value
    Command_Save(Index).Enabled = Check_Save(Index).value
End Sub

Private Sub Command_Cancel_Click()
    Unload Me
End Sub

Private Sub Command_Ok_Click()
    On Error GoTo ErrorSub

    If Check_Save(0).value = vbChecked And Trim(Text_FileName(0)) <> Empty Then
        StarProbe.FileName_MapZoom = Text_FileName(0)
        SavePicture MT2000.pZoom.Image, Text_FileName(0)
    End If
    
    If Check_Save(1).value = vbChecked And Trim(Text_FileName(1)) <> Empty Then
        StarProbe.FileName_MapOriginal = Text_FileName(1)
        SavePicture MT2000.pOriginal.Image, Text_FileName(1)
    End If
    
    If Check_Save(2).value = vbChecked And Trim(Text_FileName(2)) <> Empty Then
        StarProbe.FIleName_MeasureResult = Text_FileName(2)
        Form_StarProbe_MeasureDataSave.Show vbModal, Me
    End If
    Unload Me
    Exit Sub
    
ErrorSub:
    Call MsgBox("MAP Save Error" & vbCrLf & "(Error No." & Err.Number & "-" & Err.Description & ")", vbCritical + vbOKOnly, "ERROR")
End Sub

Private Sub Command_Save_Click(Index As Integer)
    CommonDialog1.CancelError = True

    On Error GoTo ErrorSub
    
    Dim sfilename As String
    
    If Trim(Text_FileName(Index)) = Empty Then
        sfilename = "c:\Star Probe\Image\*.bmp"
    Else
        sfilename = Text_FileName(Index)
    End If
    
    If InStr(1, sfilename, ".") = 0 Then sfilename = sfilename & ".BMP"
    sfilename = Mid(sfilename, 1, InStr(1, sfilename, ".") - 1) & ".BMP"

    CommonDialog1.DialogTitle = "Wafer Image File BMP Save"
    CommonDialog1.Filter = "Wafer Image Files(*.BMP)|*.BMP"
    CommonDialog1.Flags = cdlOFNFileMustExist + cdlOFNOverwritePrompt + cdlOFNHideReadOnly
    CommonDialog1.FileName = IIf(Trim(sfilename) = Empty, "c:\Star Probe\Image\*.bmp", sfilename)
    
    CommonDialog1.ShowSave
    
    If CommonDialog1.FileName <> "" Then
        sfilename = CommonDialog1.FileName
        Text_FileName(Index) = sfilename
    End If
    
ErrorSub:
    CommonDialog1.CancelError = False
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    Text_FileName(0) = StarProbe.FileName_MapZoom
    Text_FileName(1) = StarProbe.FileName_MapOriginal
    Text_FileName(2) = StarProbe.FIleName_MeasureResult
    
    For i = 0 To 2
        If Trim(Text_FileName(i)) <> Empty Then Check_Save(i).value = vbChecked
    Next
End Sub
