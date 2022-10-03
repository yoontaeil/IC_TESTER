Attribute VB_Name = "Module_StarProbe"
Option Explicit

Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hsrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long

Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long

Type tXY
    x As Integer
    y As Integer
    Mx As Double
    My As Double
End Type

Type tXYR
    x As Integer
    y As Integer
    r As Boolean
End Type

Type tWafer
    Chip As Boolean         ' CHIP 여부
    ChipMask As Boolean
    ChipMask_Backup As Boolean          '2018.06.01
    ChipMeasure As Boolean  ' CHIP 측정 여부
    ChipSkipDie As Boolean  ' SKIP DIE 여부
    ChipSkipDie_Backup As Boolean       '2018.06.01
    ChipPlate As Boolean    ' Wafer Plate 여부
    ChipInk As Boolean
    ChipInk2 As Boolean
    ChipInk_Backup As Boolean           '2018.06.01
    BIN As Byte
    flag As Boolean         ' 측정 여부
    FlagBad As Boolean      ' 측정 후 결과 양품이면 False, 불량이면 True
    MeasureWait As Boolean  ' 측정 추적 알고리즘 줄을 세울때의 플래그
    InkDot As Boolean
    Real_Chk As Boolean     '[2020.03.17] : 실제 측정한 좌표 저장
End Type

Type tStarProbe
    Min As tXY
    Max As tXY
    
    StartChip As tXY
    InkStart As tXY                 '2016.09.27
    CurrentChip As tXY
    CenterChipX As Integer
    CenterChipY As Integer
    
    Unit As Integer
    InchUnit As Double
    WaferSize As Double
    WaferSizemm As Double
    ChipSizeX As Double
    ChipSizeY As Double
    ChipCountX As Integer
    ChipCountY As Integer
    
    EdgeChipmm As Double
    PlateZone As Double
    
    DisplayChipSizeX As Integer
    DisplayChipSizeY As Integer
    
    DisplayOChipSizeX As Integer
    DisplayOChipSizeY As Integer
    
    Pattern_PositionCenter As tXY
    Pattern_PositionTop As tXY
    Pattern_PositionBottom As tXY
    Pattern_PositionLeft As tXY
    Pattern_PositionRight As tXY
    
    Pattern_SizeCenter As tXY
    Pattern_SizeTop As tXY
    Pattern_SizeBottom As tXY
    Pattern_SizeLeft As tXY
    Pattern_SizeRight As tXY
    
    ReMeasure As Integer
    LineOk As Integer
    RCount As Integer
    RCount_Sub As Integer
    MeasureSleep As Integer
    
    Tip_Clean As Integer
    Tipclean_Count As Long
    Tipclean_Count_Limit As Long
    
    '[ 2022.06.24 ] : tip clean 좌표 기억용
    Tipclean_X As String
    Tipclean_Y As String
    Tipclean_Z As String
    
    '[ 2022.06.28 ] : camera 좌표 기억용
    Cam_X As String
    Cam_Y As String
    Cam_Z As String
    
    Ink_After As Integer
    Ink_After_LeftPort As Integer
    Ink_After_RightPort As Integer
    Ink_After_CenterPort As Integer   '이반장
    Ink_LeftPort As Integer
    Ink_RightPort As Integer
    
    MeasureStepX As Integer
    MeasureStepY As Integer
    MeasureStartX As Integer
    MeasureStartY As Integer
    MeasureAll As Integer
    
    LimitArea As Integer
    WaferTest As Integer
    
    CountTotalChip As Double
    CountGoodDie As Double
    CountBadDie As Double
    CountSkipDie As Double
    
  
    FileName_Map As String
    FileName_MapZoom As String
    FileName_MapOriginal As String
    FIleName_MeasureResult As String
    FileName_Data As String
    
    WorkLastChipX As Integer
    WorkLastChipY As Integer
    
    WaferDivision As Integer
    
    Probe_Stop As Integer
    Probe_Stop_Tfail_Count As Long
End Type

Type tDHMS
    D As Integer
    h As Integer
    M As Integer
    s As Integer
End Type

Public StarProbe_WorkDateTime_From As Date
Public StarProbe_WorkDateTime_To As Date
Public StarProbe_WorkDateTime_Total As Double
Public StarProbe_WorkDateTime As tDHMS

Public Wafer(-100 To 900, -100 To 900) As tWafer
Public WaferTest(-100 To 900, -100 To 900) As Boolean

Public WaferTemp(-100 To 900, -100 To 900) As tWafer

''''''''''
' UNDO
Public UndoWafer(-100 To 900) As tWafer
Public UndoX As Integer, UndoY As Integer
Public UndoCountSkipDie As Long
Public UndoCountGoodDie As Long
Public UndoCountBadDie As Long
' UNDO
''''''''''

Public MeasureSeq(0 To 200000) As tXYR  ' Measure Sequence
Public MeasureChipCount As Long
Public MeasureChipCountOk As Long

''''''''''
' STOP
Public Stop_Measure As Boolean
Public Stop_MeasureSeq(0 To 200000) As tXYR
Public Stop_MeasureChipCount As Long
Public Stop_MeasureChipCountOk As Long
Public Stop_Right As Boolean
Public Stop_xx As Integer
Public Stop_yy As Integer
' STOP
''''''''''
'DOOR
Public Door1 As Integer
'DOOR
''''''''''


Public MeasureLine(1 To 124) As tXY

Public BINColor(0 To 26) As Long
Public ChipColor(0 To 6) As Long

Public StarProbe As tStarProbe
Public StarProbeTemp As tStarProbe

Public bStarProbeStart As Boolean
Public bStarprobe_AfterInk As Boolean
Public ErrorStop As Boolean
Public ProbeioMsg(1 To 13) As Integer

Type tStarProbeMessage
    Title As String
    Message As String
End Type

Public StarProbeMessage As tStarProbeMessage

Public StarProbe_DeviceName As String

Public Sub StarProbe_DefaultValue()
On Error GoTo err
    Dim i As Integer
    
    BINColor(0) = RGB(0, 50, 50)
    
    For i = 1 To 26
        BINColor(i) = RGB(30 + (8 * i), 0, 0)
    Next
    
    ChipColor(0) = &HFFFFC0
    ChipColor(1) = vbGrayText
    ChipColor(2) = vbBlue
    ChipColor(3) = vbBlack
    ChipColor(4) = vbGrayText
    ChipColor(5) = vbGreen
    
    With StarProbe
    
        .Unit = 0
        .InchUnit = 2.53999
        .WaferSize = 6
        .WaferSizemm = 125
        .ChipSizeX = 0.32
        .ChipSizeY = 0.32
        
        .EdgeChipmm = 3
        .PlateZone = 45
    
        .Pattern_PositionCenter.x = 0
        .Pattern_PositionCenter.y = 0
        .Pattern_PositionTop.x = 0
        .Pattern_PositionTop.y = -10
        .Pattern_PositionBottom.x = 0
        .Pattern_PositionBottom.y = 10
        .Pattern_PositionLeft.x = -10
        .Pattern_PositionLeft.y = 0
        .Pattern_PositionRight.x = 10
        .Pattern_PositionRight.y = 0
    
        .Pattern_SizeCenter.x = 3
        .Pattern_SizeCenter.y = 3
        .Pattern_SizeTop.x = 3
        .Pattern_SizeTop.y = 3
        .Pattern_SizeBottom.x = 3
        .Pattern_SizeBottom.y = 3
        .Pattern_SizeLeft.x = 3
        .Pattern_SizeLeft.y = 3
        .Pattern_SizeRight.x = 3
        .Pattern_SizeRight.y = 3
        
        .LineOk = 1
        .RCount = 1
        .RCount_Sub = 1
        .MeasureSleep = 100
        
        .MeasureStepX = 3
        .MeasureStepY = 4
        
        .LimitArea = 0
        .WaferTest = 0
    
    End With

    If Dir("c:\Star Probe", vbDirectory) = "" Then MkDir ("c:\Star Probe")
    If Dir("c:\Star Probe\Device", vbDirectory) = "" Then MkDir ("c:\Star Probe\Device")
    If Dir("c:\Star Probe\Data", vbDirectory) = "" Then MkDir ("c:\Star Probe\Data")
    If Dir("c:\Star Probe\Image", vbDirectory) = "" Then MkDir ("c:\Star Probe\Image")
    If Dir("c:\Star Probe\Map", vbDirectory) = "" Then MkDir ("c:\Star Probe\Map")
    
    ' 2005.08.11
    Call StarProbe_FileLoad_SystemInfo
    
err:
Resume Next
End Sub

Public Sub StarProbe_FileDeviceSave(sfilename As String)
    Dim ifreefile As Integer
    
    If Dir(sfilename, vbNormal) <> "" Then
        Kill (sfilename)
        Sleep 100
    End If

    ifreefile = FreeFile
    
    Open sfilename For Output As ifreefile
        Print #ifreefile, "Star Probe v1"
        
        '                  1234567890123456789012345678901234567890
        '                  12 1234567890123456789012345678901234567890
        Print #ifreefile, "27,Unit                                    ," & StarProbe.Unit
        Print #ifreefile, "01,Inch Unit                               ," & StarProbe.InchUnit
        Print #ifreefile, "02,Wafer Size                              ," & StarProbe.WaferSize
        Print #ifreefile, "28,Wafer Size (mm)                         ," & StarProbe.WaferSizemm
        Print #ifreefile, "03,Chip Size (X)                           ," & StarProbe.ChipSizeX
        Print #ifreefile, "04,Chip Size (Y)                           ," & StarProbe.ChipSizeY
        Print #ifreefile, "05,Chip Count (X)                          ," & StarProbe.ChipCountX
        Print #ifreefile, "06,Chip Count (Y)                          ," & StarProbe.ChipCountY
        Print #ifreefile, "29,Edge Chip (mm)                          ," & StarProbe.EdgeChipmm
        Print #ifreefile, "30,Wafer Center (X)                        ," & StarProbe.CenterChipX
        Print #ifreefile, "31,Wafer Center (Y)                        ," & StarProbe.CenterChipY
        '                  12 1234567890123456789012345678901234567890
        Print #ifreefile, "07,Pattern Position Center (X)             ," & StarProbe.Pattern_PositionCenter.x
        Print #ifreefile, "08,Pattern Position Center (Y)             ," & StarProbe.Pattern_PositionCenter.y
        Print #ifreefile, "09,Pattern Position Top (X)                ," & StarProbe.Pattern_PositionTop.x
        Print #ifreefile, "10,Pattern Position Top (Y)                ," & StarProbe.Pattern_PositionTop.y
        Print #ifreefile, "11,Pattern Position Bottom (X)             ," & StarProbe.Pattern_PositionBottom.x
        Print #ifreefile, "12,Pattern Position Bottom (Y)             ," & StarProbe.Pattern_PositionBottom.y
        Print #ifreefile, "13,Pattern Position Left (X)               ," & StarProbe.Pattern_PositionLeft.x
        Print #ifreefile, "14,Pattern Position Left (Y)               ," & StarProbe.Pattern_PositionLeft.y
        Print #ifreefile, "15,Pattern Position Right (X)              ," & StarProbe.Pattern_PositionRight.x
        Print #ifreefile, "16,Pattern Position Right (Y)              ," & StarProbe.Pattern_PositionRight.y
        '                  12 1234567890123456789012345678901234567890
        Print #ifreefile, "17,Pattern Size Center (X)                 ," & StarProbe.Pattern_SizeCenter.x
        Print #ifreefile, "18,Pattern Size Center (Y)                 ," & StarProbe.Pattern_SizeCenter.y
        Print #ifreefile, "19,Pattern Size Top (X)                    ," & StarProbe.Pattern_SizeTop.x
        Print #ifreefile, "20,Pattern Size Top (Y)                    ," & StarProbe.Pattern_SizeTop.y
        Print #ifreefile, "21,Pattern Size Bottom (X)                 ," & StarProbe.Pattern_SizeBottom.x
        Print #ifreefile, "22,Pattern Size Bottom (Y)                 ," & StarProbe.Pattern_SizeBottom.y
        Print #ifreefile, "23,Pattern Size Left (X)                   ," & StarProbe.Pattern_SizeLeft.x
        Print #ifreefile, "24,Pattern Size Left (Y)                   ," & StarProbe.Pattern_SizeLeft.y
        Print #ifreefile, "25,Pattern Size Right (X)                  ," & StarProbe.Pattern_SizeRight.x
        Print #ifreefile, "26,Pattern Size Right (Y)                  ," & StarProbe.Pattern_SizeRight.y
    Close ifreefile
End Sub

Public Function StarProbe_FileDeviceLoad(sfilename As String) As Integer
    Dim ifreefile As Integer
    Dim sLine As String
    Dim iCommand As Integer

    If Dir(sfilename, vbNormal) = "" Then
        StarProbe_FileDeviceLoad = 0  ' File Not Found
        Exit Function
    End If
    
    ifreefile = FreeFile
    
    Open sfilename For Input As ifreefile
        Line Input #ifreefile, sLine
        If Left(sLine, 10) <> "Star Probe" Then
            StarProbe_FileDeviceLoad = 2  ' Error
            Exit Function
        End If
        
        Do While Not EOF(ifreefile)
            Line Input #ifreefile, sLine
            iCommand = val(Left(sLine, 2))
            
            Select Case iCommand
                Case 1:  StarProbe.InchUnit = val(Mid(sLine, 45))
                Case 2:  StarProbe.WaferSize = val(Mid(sLine, 45))
                Case 3:  StarProbe.ChipSizeX = val(Mid(sLine, 45))
                Case 4:  StarProbe.ChipSizeY = val(Mid(sLine, 45))
                Case 5:  StarProbe.ChipCountX = val(Mid(sLine, 45))
                Case 6:  StarProbe.ChipCountY = val(Mid(sLine, 45))
                
                Case 7:  StarProbe.Pattern_PositionCenter.x = val(Mid(sLine, 45))
                Case 8:  StarProbe.Pattern_PositionCenter.y = val(Mid(sLine, 45))
                Case 9:  StarProbe.Pattern_PositionTop.x = val(Mid(sLine, 45))
                Case 10: StarProbe.Pattern_PositionTop.y = val(Mid(sLine, 45))
                Case 11: StarProbe.Pattern_PositionBottom.x = val(Mid(sLine, 45))
                Case 12: StarProbe.Pattern_PositionBottom.y = val(Mid(sLine, 45))
                Case 13: StarProbe.Pattern_PositionLeft.x = val(Mid(sLine, 45))
                Case 14: StarProbe.Pattern_PositionLeft.y = val(Mid(sLine, 45))
                Case 15: StarProbe.Pattern_PositionRight.x = val(Mid(sLine, 45))
                Case 16: StarProbe.Pattern_PositionRight.y = val(Mid(sLine, 45))
                
                Case 17: StarProbe.Pattern_SizeCenter.x = val(Mid(sLine, 45))
                Case 18: StarProbe.Pattern_SizeCenter.y = val(Mid(sLine, 45))
                Case 19: StarProbe.Pattern_SizeTop.x = val(Mid(sLine, 45))
                Case 20: StarProbe.Pattern_SizeTop.y = val(Mid(sLine, 45))
                Case 21: StarProbe.Pattern_SizeBottom.x = val(Mid(sLine, 45))
                Case 22: StarProbe.Pattern_SizeBottom.y = val(Mid(sLine, 45))
                Case 23: StarProbe.Pattern_SizeLeft.x = val(Mid(sLine, 45))
                Case 24: StarProbe.Pattern_SizeLeft.y = val(Mid(sLine, 45))
                Case 25: StarProbe.Pattern_SizeRight.x = val(Mid(sLine, 45))
                Case 26: StarProbe.Pattern_SizeRight.y = val(Mid(sLine, 45))
                
                Case 27:  StarProbe.Unit = val(Mid(sLine, 45))
                Case 28:  StarProbe.WaferSizemm = val(Mid(sLine, 45))
                Case 29:  StarProbe.EdgeChipmm = val(Mid(sLine, 45))
                Case 30:  StarProbe.CenterChipX = val(Mid(sLine, 45))
                Case 31:  StarProbe.CenterChipY = val(Mid(sLine, 45))
            End Select
        Loop
    Close ifreefile
    StarProbe_FileDeviceLoad = 1  ' Ok
End Function

Public Sub StarProbe_FileSave_ControlMap(sfilename As String)

    If Dir(sfilename, vbNormal) <> "" Then
        Kill (sfilename)
        Sleep 100
    End If
    
    Dim ifreefile As Integer
    
    Dim x As Integer, y As Integer, i As Integer
    Dim s As String

    ifreefile = FreeFile
    
    Open sfilename For Output As ifreefile
    
    Print #ifreefile, "Star Probe v1 - Control Map"
    
    With StarProbe
    
        Print #ifreefile, .Min.x
        Print #ifreefile, .Min.y
        
        Print #ifreefile, .Max.x
        Print #ifreefile, .Max.y
        
        Print #ifreefile, .StartChip.x
        Print #ifreefile, .StartChip.y
        
        Print #ifreefile, .CurrentChip.x
        Print #ifreefile, .CurrentChip.y
        
        Print #ifreefile, .CenterChipX
        Print #ifreefile, .CenterChipY
    
        Print #ifreefile, .Unit
        Print #ifreefile, .InchUnit
        Print #ifreefile, .WaferSize
        Print #ifreefile, .WaferSizemm
        Print #ifreefile, .ChipSizeX
        Print #ifreefile, .ChipSizeY
        Print #ifreefile, .ChipCountX
        Print #ifreefile, .ChipCountY
    
        Print #ifreefile, .EdgeChipmm
        Print #ifreefile, .PlateZone
    
        Print #ifreefile, .DisplayChipSizeX
        Print #ifreefile, .DisplayChipSizeY
    
        Print #ifreefile, .Pattern_PositionCenter.x
        Print #ifreefile, .Pattern_PositionCenter.y
        Print #ifreefile, .Pattern_PositionTop.x
        Print #ifreefile, .Pattern_PositionTop.y
        Print #ifreefile, .Pattern_PositionBottom.x
        Print #ifreefile, .Pattern_PositionBottom.y
        Print #ifreefile, .Pattern_PositionLeft.x
        Print #ifreefile, .Pattern_PositionLeft.y
        Print #ifreefile, .Pattern_PositionRight.x
        Print #ifreefile, .Pattern_PositionRight.y
    
        Print #ifreefile, .Pattern_SizeCenter.x
        Print #ifreefile, .Pattern_SizeCenter.y
        Print #ifreefile, .Pattern_SizeTop.x
        Print #ifreefile, .Pattern_SizeTop.y
        Print #ifreefile, .Pattern_SizeBottom.x
        Print #ifreefile, .Pattern_SizeBottom.y
        Print #ifreefile, .Pattern_SizeLeft.x
        Print #ifreefile, .Pattern_SizeLeft.y
        Print #ifreefile, .Pattern_SizeRight.x
        Print #ifreefile, .Pattern_SizeRight.y
    
        Print #ifreefile, .ReMeasure
        Print #ifreefile, .LineOk
        Print #ifreefile, .RCount
        Print #ifreefile, .RCount_Sub
        Print #ifreefile, .MeasureSleep
    
        Print #ifreefile, .Ink_After
        Print #ifreefile, .Ink_After_LeftPort
        Print #ifreefile, .Ink_After_RightPort
        Print #ifreefile, .Ink_LeftPort
        Print #ifreefile, .Ink_RightPort
    
        Print #ifreefile, .MeasureStepX
        Print #ifreefile, .MeasureStepY
        Print #ifreefile, .MeasureStartX
        Print #ifreefile, .MeasureStartY
    
        Print #ifreefile, .LimitArea
        Print #ifreefile, .WaferTest
    
    End With
    
    Print #ifreefile, ""
    Print #ifreefile, "########## Control Map ##########"
    
    For y = 0 To 900
        For x = 0 To 900
            If Wafer(x, y).Chip Then
                s = "@" & STR_FIX(Str(x), 3) & STR_FIX(Str(y), 3)
                s = s & IIf(Wafer(x, y).Chip, 1, 0)
                s = s & IIf(Wafer(x, y).ChipMask, 1, 0)
                s = s & IIf(Wafer(x, y).ChipMeasure, 1, 0)
                s = s & IIf(Wafer(x, y).ChipSkipDie, 1, 0)
                s = s & IIf(Wafer(x, y).ChipPlate, 1, 0)
                s = s & IIf(Wafer(x, y).ChipInk, 1, 0) ' 2005.09.09
                's = s & IIf(Wafer(x, y).ChipInk2, 1, 0) ' 2005.09.09
                s = s & IIf(Wafer(x, y).flag, 1, 0)
                s = s & IIf(Wafer(x, y).FlagBad, 1, 0)
                s = s & IIf(Wafer(x, y).MeasureWait, 1, 0)
                s = s & IIf(Wafer(x, y).InkDot, 1, 0)
                s = s & Wafer(x, y).BIN
                Print #ifreefile, s
            End If
        Next
    Next
    
    Print #ifreefile, "########## ETC ##########"
    
                      '1234567890123456789012345678901
    Print #ifreefile, "#Wafer Divion                :" & StarProbe.WaferDivision  ' 2005.09.12
    
    Print #ifreefile, "########## BIN Color ##########"
    
    For i = 0 To 26
                          '1234567890123456789012345678901
        Print #ifreefile, "#BIN Color                   :" & BINColor(i)
    Next
    
    Print #ifreefile, "########## Chip Color ##########"
    
    For i = 0 To 6
        Print #ifreefile, "#Chip Color                  :" & ChipColor(i)
    Next
    
    Print #ifreefile, "########## End Of File ##########"
    
    Close ifreefile

End Sub

Public Sub StarProbe_FileLoad_ControlMap(sfilename As String)
'On Error GoTo err

    Dim ink2_flag As Boolean
    Dim first_find As Boolean
    
    first_find = False
    ink2_flag = False
    
    If Dir(sfilename, vbNormal) = "" Then
        Exit Sub
    End If

    Dim ifreefile As Integer
    Dim sLine As String
    
    Dim x As Integer, y As Integer, i As Integer
    
    Dim waferx As Integer, wafery As Integer, b As Boolean
    
    Dim bETC As Boolean
    bETC = False
    
    ifreefile = FreeFile
    
    Open sfilename For Input As ifreefile
    
    Line Input #ifreefile, sLine  ' Version Information
     
    With StarProbe
    
        Line Input #ifreefile, sLine
        .Min.x = val(sLine)
        Line Input #ifreefile, sLine
        .Min.y = val(sLine)
        
        Line Input #ifreefile, sLine
        .Max.x = val(sLine)
        Line Input #ifreefile, sLine
        .Max.y = val(sLine)
        
        If UCase(Right(MT2000.SSPanel2(0).Caption, 3)) = "AOI" Then
            Line Input #ifreefile, sLine
            Line Input #ifreefile, sLine
        Else
            Line Input #ifreefile, sLine
            .StartChip.x = val(sLine)
            Line Input #ifreefile, sLine
            .StartChip.y = val(sLine)
        End If
        
        Line Input #ifreefile, sLine
        .CurrentChip.x = val(sLine)
        Line Input #ifreefile, sLine
        .CurrentChip.y = val(sLine)
        
        Line Input #ifreefile, sLine
        .CenterChipX = val(sLine)
        Line Input #ifreefile, sLine
        .CenterChipY = val(sLine)
    
        Line Input #ifreefile, sLine
        .Unit = val(sLine)
        Line Input #ifreefile, sLine
        .InchUnit = val(sLine)
        Line Input #ifreefile, sLine
        .WaferSize = val(sLine)
        Line Input #ifreefile, sLine
        .WaferSizemm = val(sLine)
        Line Input #ifreefile, sLine
        .ChipSizeX = val(sLine)
        Line Input #ifreefile, sLine
        .ChipSizeY = val(sLine)
        Line Input #ifreefile, sLine
        .ChipCountX = val(sLine)
        Line Input #ifreefile, sLine
        .ChipCountY = val(sLine)
    
        Line Input #ifreefile, sLine
        .EdgeChipmm = val(sLine)
        Line Input #ifreefile, sLine
        .PlateZone = val(sLine)
    
        Line Input #ifreefile, sLine
        .DisplayChipSizeX = val(sLine)
        Line Input #ifreefile, sLine
        .DisplayChipSizeY = val(sLine)
    
        Line Input #ifreefile, sLine
        .Pattern_PositionCenter.x = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_PositionCenter.y = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_PositionTop.x = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_PositionTop.y = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_PositionBottom.x = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_PositionBottom.y = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_PositionLeft.x = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_PositionLeft.y = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_PositionRight.x = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_PositionRight.y = val(sLine)
    
        Line Input #ifreefile, sLine
        .Pattern_SizeCenter.x = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_SizeCenter.y = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_SizeTop.x = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_SizeTop.y = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_SizeBottom.x = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_SizeBottom.y = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_SizeLeft.x = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_SizeLeft.y = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_SizeRight.x = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_SizeRight.y = val(sLine)
    
        Line Input #ifreefile, sLine
'        .ReMeasure = Val(sLine)
        Line Input #ifreefile, sLine
'        .LineOk = Val(sLine)
        Line Input #ifreefile, sLine
'        .RCount = Val(sLine)
        Line Input #ifreefile, sLine
'        .RCount_Sub = Val(sLine)
        Line Input #ifreefile, sLine
'        .MeasureSleep = Val(sLine)
    
        Line Input #ifreefile, sLine
'        .Ink_After = Val(sLine)
        Line Input #ifreefile, sLine
'        .Ink_After_LeftPort = Val(sLine)
        Line Input #ifreefile, sLine
'        .Ink_After_RightPort = Val(sLine)
        Line Input #ifreefile, sLine
'        .Ink_LeftPort = Val(sLine)
        Line Input #ifreefile, sLine
'        .Ink_RightPort = Val(sLine)
    
        Line Input #ifreefile, sLine
'        .MeasureStepX = Val(sLine)
        Line Input #ifreefile, sLine
'        .MeasureStepY = Val(sLine)
        Line Input #ifreefile, sLine
'        .MeasureStartX = Val(sLine)
        Line Input #ifreefile, sLine
'        .MeasureStartY = Val(sLine)
    
        Line Input #ifreefile, sLine
        .LimitArea = val(sLine)
        Line Input #ifreefile, sLine
        .WaferTest = val(sLine)
    
    End With
    
    Line Input #ifreefile, sLine
    Line Input #ifreefile, sLine
    
    Erase Wafer
    
'    For y = 0 To 900
'
'        Line Input #ifreefile, sLine
'
'        For x = 0 To 900
'
'     '       i = Val(Mid(sLine, x + 1, 1)) - 64
'             i = Val(Asc(Mid(sLine, x + 1, 1))) - 64
'
'            Wafer(x, y).Chip = IIf(i And 1, True, False)
'            Wafer(x, y).ChipMask = IIf(i And 2, True, False)
'            Wafer(x, y).ChipMeasure = IIf(i And 4, True, False)
'            Wafer(x, y).ChipSkipDie = IIf(i And 8, True, False)
'            Wafer(x, y).ChipPlate = IIf(i And 16, True, False)
'
'        Next
'
'    Next

    Do While Not EOF(ifreefile)
    
        Line Input #ifreefile, sLine
        
        If InStr(UCase(sfilename), "AOI") <> 0 Then
            If Left(sLine, 1) = "@" Then
                waferx = val(Mid(sLine, 2, 3))
                wafery = val(Mid(sLine, 5, 3))
                
                b = IIf(val(Mid(sLine, 8, 1)) = 1, True, False):  Wafer(waferx, wafery).Chip = b
                b = IIf(val(Mid(sLine, 9, 1)) = 1, True, False):  Wafer(waferx, wafery).ChipMask = b
                b = IIf(val(Mid(sLine, 10, 1)) = 1, True, False): Wafer(waferx, wafery).ChipMeasure = b
                b = IIf(val(Mid(sLine, 11, 1)) = 1, True, False): Wafer(waferx, wafery).ChipSkipDie = b
                b = IIf(val(Mid(sLine, 12, 1)) = 1, True, False): Wafer(waferx, wafery).ChipPlate = b
                b = IIf(val(Mid(sLine, 13, 1)) = 1, True, False): Wafer(waferx, wafery).ChipInk = b
                b = IIf(val(Mid(sLine, 14, 1)) = 1, True, False): Wafer(waferx, wafery).flag = b
                b = IIf(val(Mid(sLine, 15, 1)) = 1, True, False): Wafer(waferx, wafery).FlagBad = b
                b = IIf(val(Mid(sLine, 16, 1)) = 1, True, False): Wafer(waferx, wafery).MeasureWait = b
                b = IIf(val(Mid(sLine, 17, 1)) = 1, True, False): Wafer(waferx, wafery).InkDot = b
                i = val(Mid(sLine, 18, 3)):                       Wafer(waferx, wafery).BIN = i
                '[ 2020.11.02 ]
                If Wafer(waferx, wafery).BIN = AOI_BIN Then                  'aoi ng number = 17
                    AOI_FAIL_COUNT = AOI_FAIL_COUNT + 1
                End If
            End If
        Else
            If Left(sLine, 1) = "@" Then
                If first_find = False Then
                    first_find = True
                    If ink2_flag = False Then
                        If Len(sLine) = 19 Then
                            ink2_flag = True
                        End If
                    End If
                End If
                
                waferx = val(Mid(sLine, 2, 3))
                wafery = val(Mid(sLine, 5, 3))
                
                If ink2_flag = True Then
                    b = IIf(val(Mid(sLine, 8, 1)) = 1, True, False):  Wafer(waferx, wafery).Chip = b
                    b = IIf(val(Mid(sLine, 9, 1)) = 1, True, False):  Wafer(waferx, wafery).ChipMask = b
                    If Wafer(waferx, wafery).ChipMask = True Then               '2018.06.01
                        Wafer(waferx, wafery).ChipMask_Backup = True
                    Else
                        Wafer(waferx, wafery).ChipMask_Backup = False
                    End If
                    b = IIf(val(Mid(sLine, 10, 1)) = 1, True, False): Wafer(waferx, wafery).ChipMeasure = b
                    b = IIf(val(Mid(sLine, 11, 1)) = 1, True, False): Wafer(waferx, wafery).ChipSkipDie = b
                    If Wafer(waferx, wafery).ChipSkipDie = True Then               '2018.06.01
                        Wafer(waferx, wafery).ChipSkipDie_Backup = True
                    Else
                        Wafer(waferx, wafery).ChipSkipDie_Backup = False
                    End If
                    b = IIf(val(Mid(sLine, 12, 1)) = 1, True, False): Wafer(waferx, wafery).ChipPlate = b
                    b = IIf(val(Mid(sLine, 13, 1)) = 1, True, False): Wafer(waferx, wafery).ChipInk = b
                    If Wafer(waferx, wafery).ChipInk = True Then               '2018.06.01
                        Wafer(waferx, wafery).ChipInk_Backup = True
                    Else
                        Wafer(waferx, wafery).ChipInk_Backup = False
                    End If
                    b = IIf(val(Mid(sLine, 14, 1)) = 1, True, False): Wafer(waferx, wafery).ChipInk2 = b
                    
                    b = IIf(val(Mid(sLine, 15, 1)) = 1, True, False): Wafer(waferx, wafery).flag = False 'b
                    b = IIf(val(Mid(sLine, 16, 1)) = 1, True, False): Wafer(waferx, wafery).FlagBad = False 'b
                    b = IIf(val(Mid(sLine, 17, 1)) = 1, True, False): Wafer(waferx, wafery).MeasureWait = False 'b
                    b = IIf(val(Mid(sLine, 18, 1)) = 1, True, False): Wafer(waferx, wafery).InkDot = False 'b
                    i = val(Mid(sLine, 19, 3)):                       Wafer(waferx, wafery).BIN = i
                Else
                    b = IIf(val(Mid(sLine, 8, 1)) = 1, True, False):  Wafer(waferx, wafery).Chip = b
                    b = IIf(val(Mid(sLine, 9, 1)) = 1, True, False):  Wafer(waferx, wafery).ChipMask = b
                    If Wafer(waferx, wafery).ChipMask = True Then               '2018.06.01
                        Wafer(waferx, wafery).ChipMask_Backup = True
                    Else
                        Wafer(waferx, wafery).ChipMask_Backup = False
                    End If
                    b = IIf(val(Mid(sLine, 10, 1)) = 1, True, False): Wafer(waferx, wafery).ChipMeasure = b
                    b = IIf(val(Mid(sLine, 11, 1)) = 1, True, False): Wafer(waferx, wafery).ChipSkipDie = b
                    If Wafer(waferx, wafery).ChipSkipDie = True Then               '2018.06.01
                        Wafer(waferx, wafery).ChipSkipDie_Backup = True
                    Else
                        Wafer(waferx, wafery).ChipSkipDie_Backup = False
                    End If
                    b = IIf(val(Mid(sLine, 12, 1)) = 1, True, False): Wafer(waferx, wafery).ChipPlate = b
                    b = IIf(val(Mid(sLine, 13, 1)) = 1, True, False): Wafer(waferx, wafery).ChipInk = b
                    If Wafer(waferx, wafery).ChipInk = True Then               '2018.06.01
                        Wafer(waferx, wafery).ChipInk_Backup = True
                    Else
                        Wafer(waferx, wafery).ChipInk_Backup = False
                    End If
                    b = IIf(val(Mid(sLine, 14, 1)) = 1, True, False): Wafer(waferx, wafery).flag = False 'b
                    b = IIf(val(Mid(sLine, 15, 1)) = 1, True, False): Wafer(waferx, wafery).FlagBad = False 'b
                    b = IIf(val(Mid(sLine, 16, 1)) = 1, True, False): Wafer(waferx, wafery).MeasureWait = False 'b
                    b = IIf(val(Mid(sLine, 17, 1)) = 1, True, False): Wafer(waferx, wafery).InkDot = False 'b
                    i = val(Mid(sLine, 18, 3)):                       Wafer(waferx, wafery).BIN = i
                End If
            End If
        End If
        
        If sLine = "########## ETC ##########" Then
            bETC = True
            Exit Do
        End If
        
    Loop
    
    If bETC Then
        
        Line Input #ifreefile, sLine
        If Trim(Left(sLine, 30)) = "#Wafer Divion                :" Then
            StarProbe.WaferDivision = val(Mid(sLine, 31))
        End If
    
        Line Input #ifreefile, sLine
        For i = 0 To 26
            Line Input #ifreefile, sLine
            BINColor(i) = val(Mid(sLine, 31))
        Next
    
        Line Input #ifreefile, sLine
        For i = 0 To 6
            Line Input #ifreefile, sLine
            ChipColor(i) = val(Mid(sLine, 31))
        Next
    
    End If
    
    Close ifreefile

    Stop_Measure = False
    
'err:
'    Resume Next
End Sub

Public Sub StarProbe_FileSave_SystemInfo()

    Dim sfilename As String
    
    sfilename = "c:\Star Probe\StarProbe.IFO"

    If Dir(sfilename, vbNormal) <> "" Then
        Kill (sfilename)
        Sleep 100
    End If
    
    Dim ifreefile As Integer
    
    Dim i As Integer

    ifreefile = FreeFile
    
    Open sfilename For Output As ifreefile
    
    Print #ifreefile, "Star Probe v1 - System Information"
    
    Print #ifreefile, "########## System ##########"
    
    With StarProbe
    
        Print #ifreefile, .Min.x
        Print #ifreefile, .Min.y
        
        Print #ifreefile, .Max.x
        Print #ifreefile, .Max.y
        
        Print #ifreefile, .StartChip.x
        Print #ifreefile, .StartChip.y
        
        Print #ifreefile, .CurrentChip.x
        Print #ifreefile, .CurrentChip.y
        
        Print #ifreefile, .CenterChipX
        Print #ifreefile, .CenterChipY
    
        Print #ifreefile, .Unit
        Print #ifreefile, .InchUnit
        Print #ifreefile, .WaferSize
        Print #ifreefile, .WaferSizemm
        Print #ifreefile, .ChipSizeX
        Print #ifreefile, .ChipSizeY
        Print #ifreefile, .ChipCountX
        Print #ifreefile, .ChipCountY
    
        Print #ifreefile, .EdgeChipmm
        Print #ifreefile, .PlateZone
    
        Print #ifreefile, .DisplayChipSizeX
        Print #ifreefile, .DisplayChipSizeY
    
        Print #ifreefile, .Pattern_PositionCenter.x
        Print #ifreefile, .Pattern_PositionCenter.y
        Print #ifreefile, .Pattern_PositionTop.x
        Print #ifreefile, .Pattern_PositionTop.y
        Print #ifreefile, .Pattern_PositionBottom.x
        Print #ifreefile, .Pattern_PositionBottom.y
        Print #ifreefile, .Pattern_PositionLeft.x
        Print #ifreefile, .Pattern_PositionLeft.y
        Print #ifreefile, .Pattern_PositionRight.x
        Print #ifreefile, .Pattern_PositionRight.y
    
        Print #ifreefile, .Pattern_SizeCenter.x
        Print #ifreefile, .Pattern_SizeCenter.y
        Print #ifreefile, .Pattern_SizeTop.x
        Print #ifreefile, .Pattern_SizeTop.y
        Print #ifreefile, .Pattern_SizeBottom.x
        Print #ifreefile, .Pattern_SizeBottom.y
        Print #ifreefile, .Pattern_SizeLeft.x
        Print #ifreefile, .Pattern_SizeLeft.y
        Print #ifreefile, .Pattern_SizeRight.x
        Print #ifreefile, .Pattern_SizeRight.y
    
        Print #ifreefile, .ReMeasure
        Print #ifreefile, .LineOk
        Print #ifreefile, .RCount
        Print #ifreefile, .RCount_Sub
        Print #ifreefile, .MeasureSleep
        
        Print #ifreefile, .Tip_Clean
        Print #ifreefile, .Tipclean_Count
        Print #ifreefile, .Tipclean_Count_Limit
    
        Print #ifreefile, .Ink_After
        Print #ifreefile, .Ink_After_LeftPort
        Print #ifreefile, .Ink_After_RightPort
        Print #ifreefile, .Ink_After_CenterPort
        Print #ifreefile, .Ink_LeftPort
        Print #ifreefile, .Ink_RightPort
    
        Print #ifreefile, .MeasureStepX
        Print #ifreefile, .MeasureStepY
        Print #ifreefile, .MeasureStartX
        Print #ifreefile, .MeasureStartY
        Print #ifreefile, .MeasureAll
    
        Print #ifreefile, .LimitArea
        Print #ifreefile, .WaferTest
        
        Print #ifreefile, .Probe_Stop                     '추가
        Print #ifreefile, .Probe_Stop_Tfail_Count
        
        '[ 2022.06.24 ]
        Print #ifreefile, .Tipclean_X
        Print #ifreefile, .Tipclean_Y
        Print #ifreefile, .Tipclean_Z
    
    End With
    
    
    For i = 1 To 13
        Print #ifreefile, ProbeioMsg(i)
    Next
    
    
    Print #ifreefile, "########## BIN Color ##########"
    
    For i = 0 To 26
        Print #ifreefile, BINColor(i)
    Next
    
    Print #ifreefile, "########## Chip Color ##########"
    
    For i = 0 To 5
        Print #ifreefile, ChipColor(i)
    Next
    
    Print #ifreefile, "######## BIN COMMAND ##########"
    
    For i = 0 To 24
        Print #ifreefile, BIN_Command(i)
    Next
    
    '' 2019.12.17 : server path add
    Print #ifreefile, SelectExt.Text4.Text
    ''
    '' 2020.09.07 : map path add
    Print #ifreefile, SelectExt.Text5.Text
    ''
    If AOI_MODE = 1 Then
        '' 2020.09.07 : aoi path add
        Print #ifreefile, SelectExt.Text10.Text
    End If
    ''
    
    '[ 2021.12.16 ] : channel select
    If SelectExt.Option3(0).value = True Then
        Print #ifreefile, "1"
        CH_SET = 1
    ElseIf SelectExt.Option3(1).value = True Then
        Print #ifreefile, "2"
        CH_SET = 2
    Else
        Print #ifreefile, "4"
        CH_SET = 4
    End If
    
    Print #ifreefile, "########## End Of File ##########"
    
    Close ifreefile

End Sub

Public Sub StarProbe_FileLoad_SystemInfo()
On Error GoTo err

    Dim sfilename As String
    
    sfilename = "c:\Star Probe\StarProbe.IFO"

    If Dir(sfilename, vbNormal) = "" Then
        Exit Sub
    End If

    Dim ifreefile As Integer
    Dim sLine As String
    
    Dim i As Integer
    
    ifreefile = FreeFile
    
    Open sfilename For Input As ifreefile
    
    Line Input #ifreefile, sLine  ' Version Information
     
    Line Input #ifreefile, sLine
     
    With StarProbe
    
        Line Input #ifreefile, sLine
        .Min.x = val(sLine)
        Line Input #ifreefile, sLine
        .Min.y = val(sLine)
        
        Line Input #ifreefile, sLine
        .Max.x = val(sLine)
        Line Input #ifreefile, sLine
        .Max.y = val(sLine)
        
        Line Input #ifreefile, sLine
        .StartChip.x = val(sLine)
        Line Input #ifreefile, sLine
        .StartChip.y = val(sLine)
        
        Line Input #ifreefile, sLine
        .CurrentChip.x = val(sLine)
        Line Input #ifreefile, sLine
        .CurrentChip.y = val(sLine)
        
        Line Input #ifreefile, sLine
        .CenterChipX = val(sLine)
        Line Input #ifreefile, sLine
        .CenterChipY = val(sLine)
    
        Line Input #ifreefile, sLine
        .Unit = val(sLine)
        Line Input #ifreefile, sLine
        .InchUnit = val(sLine)
        Line Input #ifreefile, sLine
        .WaferSize = val(sLine)
        Line Input #ifreefile, sLine
        .WaferSizemm = val(sLine)
        Line Input #ifreefile, sLine
        .ChipSizeX = val(sLine)
        Line Input #ifreefile, sLine
        .ChipSizeY = val(sLine)
        Line Input #ifreefile, sLine
        .ChipCountX = val(sLine)
        Line Input #ifreefile, sLine
        .ChipCountY = val(sLine)
    
        Line Input #ifreefile, sLine
        .EdgeChipmm = val(sLine)
        Line Input #ifreefile, sLine
        .PlateZone = val(sLine)
    
        Line Input #ifreefile, sLine
        .DisplayChipSizeX = val(sLine)
        Line Input #ifreefile, sLine
        .DisplayChipSizeY = val(sLine)
    
        Line Input #ifreefile, sLine
        .Pattern_PositionCenter.x = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_PositionCenter.y = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_PositionTop.x = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_PositionTop.y = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_PositionBottom.x = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_PositionBottom.y = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_PositionLeft.x = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_PositionLeft.y = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_PositionRight.x = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_PositionRight.y = val(sLine)
    
        Line Input #ifreefile, sLine
        .Pattern_SizeCenter.x = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_SizeCenter.y = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_SizeTop.x = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_SizeTop.y = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_SizeBottom.x = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_SizeBottom.y = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_SizeLeft.x = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_SizeLeft.y = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_SizeRight.x = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_SizeRight.y = val(sLine)
    
        Line Input #ifreefile, sLine
        .ReMeasure = val(sLine)
        Line Input #ifreefile, sLine
        .LineOk = val(sLine)
        Line Input #ifreefile, sLine
        .RCount = val(sLine)
        Line Input #ifreefile, sLine
        .RCount_Sub = val(sLine)
        Line Input #ifreefile, sLine
        .MeasureSleep = val(sLine)
        
        Line Input #ifreefile, sLine
        .Tip_Clean = val(sLine)
        Line Input #ifreefile, sLine
        .Tipclean_Count = val(sLine)
        Line Input #ifreefile, sLine
        .Tipclean_Count_Limit = val(sLine)
    
        Line Input #ifreefile, sLine
        .Ink_After = val(sLine)
        Line Input #ifreefile, sLine
        .Ink_After_LeftPort = val(sLine)
        Line Input #ifreefile, sLine
        .Ink_After_RightPort = val(sLine)
        Line Input #ifreefile, sLine
        .Ink_After_CenterPort = val(sLine)
        Line Input #ifreefile, sLine
        .Ink_LeftPort = val(sLine)
        Line Input #ifreefile, sLine
        .Ink_RightPort = val(sLine)
    
        Line Input #ifreefile, sLine
        .MeasureStepX = val(sLine)
        Line Input #ifreefile, sLine
        .MeasureStepY = val(sLine)
        Line Input #ifreefile, sLine
        .MeasureStartX = val(sLine)
        Line Input #ifreefile, sLine
        .MeasureStartY = val(sLine)
    
        Line Input #ifreefile, sLine
        .MeasureAll = val(sLine)
        
        Line Input #ifreefile, sLine
        .LimitArea = val(sLine)
        Line Input #ifreefile, sLine
        .WaferTest = val(sLine)
        
        Line Input #ifreefile, sLine     ' 추가
        .Probe_Stop = val(sLine)
        Line Input #ifreefile, sLine
        .Probe_Stop_Tfail_Count = val(sLine)
        
        '[ 2022.06.24 ]
        Line Input #ifreefile, sLine
        .Tipclean_X = val(sLine)
        Line Input #ifreefile, sLine
        .Tipclean_Y = val(sLine)
        Line Input #ifreefile, sLine
        .Tipclean_Z = val(sLine)
    
    End With
    
    '[ 2017.08.19 ] : x,y step을 배열에 적용
    For i = 0 To 49
        XPitch(i) = StarProbe.MeasureStepX
        YPitch(i) = StarProbe.MeasureStepY
    Next i
    
    For i = 1 To 13
        Line Input #ifreefile, sLine
        ProbeioMsg(i) = val(sLine)
    Next
    
    Line Input #ifreefile, sLine
    
    For i = 0 To 26
        Line Input #ifreefile, sLine
        BINColor(i) = val(sLine)
    Next
    
    Line Input #ifreefile, sLine
    
    For i = 0 To 5
        Line Input #ifreefile, sLine
        ChipColor(i) = val(sLine)
    Next
    
    Line Input #ifreefile, sLine
    
    For i = 0 To 24
       Line Input #ifreefile, sLine
        BIN_Command(i) = sLine
    Next
    Line Input #ifreefile, sLine
    Server_path = sLine
    SelectExt.Text4.Text = Server_path
    
    Line Input #ifreefile, sLine                '2020.09.07 : barcode map path
    MAP_path = sLine
    SelectExt.Text5.Text = MAP_path
    
    If AOI_MODE = 1 Then
        Line Input #ifreefile, sLine                '2020.09.07 : barcode map path
        AOI_path = sLine
        SelectExt.Text10.Text = AOI_path
    End If
    
    
    '[ 2021.12.16 ] : channel select
    Line Input #ifreefile, sLine
    If sLine = "1" Then
        SelectExt.Option3(0).value = True
        CH_SET = 1
    ElseIf sLine = "2" Then
        SelectExt.Option3(1).value = True
        CH_SET = 2
    Else
        SelectExt.Option3(2).value = True
        CH_SET = 4
    End If
    MT2000.Label15.Caption = CH_SET & "CH"
    
    Close ifreefile

err:
    Resume Next
End Sub

Public Sub Display_Wafer(pZoom As PictureBox, pOriginal As PictureBox, _
                         Shape_Chip As Shape, Shape_FirstChip As Shape, Shape_Ink As Shape, Shape_Move As Shape, _
                         Shape_OChip As Shape, Shape_OFirstChip As Shape, Shape_OInk As Shape, Shape_OMove As Shape, _
                         VScroll_Zoom As VScrollBar, HScroll_Zoom As HScrollBar)
    
    ' 오류가 발생하면 ErrorHandler 루틴으로 가서 처리
    On Error GoTo ErrorHandler

    Dim x As Integer, y As Integer
    Dim xx As Integer, yy As Integer
    Dim SizeX As Long, SizeY As Long
    Dim DisplayChipColor As Long
    Dim YOON As Integer
    
    If MT2000.Option3(0).value = True Then
        YOON = 1
    ElseIf MT2000.Option3(1).value = True Then
        YOON = 2
    ElseIf MT2000.Option3(2).value = True Then
        YOON = 3
    ElseIf MT2000.Option3(3).value = True Then
        YOON = 4
    End If
    
    StarProbe.DisplayChipSizeX = Round(StarProbe.ChipSizeX, 1) * 10 * YOON
    StarProbe.DisplayChipSizeY = Round(StarProbe.ChipSizeY, 1) * 10 * YOON

    If StarProbe.DisplayChipSizeX <= 4 Or StarProbe.DisplayChipSizeY <= 4 Then
        StarProbe.DisplayChipSizeX = 3
        StarProbe.DisplayChipSizeY = 3
    End If

Redraw:

    SizeX = StarProbe.ChipCountX
    SizeX = SizeX + 3
    SizeX = SizeX * StarProbe.DisplayChipSizeX
    SizeX = SizeX * 15
    
    SizeY = StarProbe.ChipCountY
    SizeY = SizeY + 3
    SizeY = SizeY * StarProbe.DisplayChipSizeY
    SizeY = SizeY * 15
    
    Shape_Chip.width = StarProbe.DisplayChipSizeX + 4
    Shape_Chip.Height = StarProbe.DisplayChipSizeY + 4
    
    Shape_FirstChip.width = StarProbe.DisplayChipSizeX + 4
    Shape_FirstChip.Height = StarProbe.DisplayChipSizeY + 4
    
    Shape_Ink.width = StarProbe.DisplayChipSizeX + 4
    Shape_Ink.Height = StarProbe.DisplayChipSizeY + 4
    
    Shape_Move.width = StarProbe.DisplayChipSizeX + 4
    Shape_Move.Height = StarProbe.DisplayChipSizeY + 4
    
    pZoom.width = SizeX
    pZoom.Height = SizeY
    
    StarProbe.DisplayOChipSizeX = 1
    StarProbe.DisplayOChipSizeY = 1
    
'    Select Case StarProbe.ChipCountX
'    Case Is >= 300
'        StarProbe.DisplayOChipSizeX = 1
'    Case Is >= 200
'        StarProbe.DisplayOChipSizeX = 2
'    Case Is >= 100
'        StarProbe.DisplayOChipSizeX = 3
'    Case Is < 100
'        StarProbe.DisplayOChipSizeX = 6
'    End Select
    Select Case StarProbe.ChipCountX
    Case Is >= 300
        StarProbe.DisplayOChipSizeX = 1
    Case 200 To 299
        StarProbe.DisplayOChipSizeX = 2
    Case 100 To 199
        StarProbe.DisplayOChipSizeX = 3
    Case 90 To 99
        StarProbe.DisplayOChipSizeX = 6
    Case 80 To 89
        StarProbe.DisplayOChipSizeX = 7
    Case 70 To 79
        StarProbe.DisplayOChipSizeX = 8
    Case 60 To 69
        StarProbe.DisplayOChipSizeX = 9
    Case 50 To 59
        StarProbe.DisplayOChipSizeX = 10
    Case 40 To 49
        StarProbe.DisplayOChipSizeX = 12
    Case 30 To 39
        StarProbe.DisplayOChipSizeX = 16
    Case 20 To 29
        StarProbe.DisplayOChipSizeX = 20
    Case 10 To 19
        StarProbe.DisplayOChipSizeX = 30
    Case Is < 10
        StarProbe.DisplayOChipSizeX = 60
    End Select

'    Select Case StarProbe.ChipCountY
'    Case Is >= 300
'        StarProbe.DisplayOChipSizeY = 1
'    Case Is >= 200
'        StarProbe.DisplayOChipSizeY = 2
'    Case Is >= 100
'        StarProbe.DisplayOChipSizeY = 3
'    Case Is < 100
'        StarProbe.DisplayOChipSizeY = 6
'    End Select
    Select Case StarProbe.ChipCountY
    Case Is >= 300
        StarProbe.DisplayOChipSizeY = 1
    Case 200 To 299
        StarProbe.DisplayOChipSizeY = 2
    Case 100 To 199
        StarProbe.DisplayOChipSizeY = 3
    Case 90 To 99
        StarProbe.DisplayOChipSizeY = 6
    Case 80 To 89
        StarProbe.DisplayOChipSizeY = 7
    Case 70 To 79
        StarProbe.DisplayOChipSizeY = 8
    Case 60 To 69
        StarProbe.DisplayOChipSizeY = 9
    Case 50 To 59
        StarProbe.DisplayOChipSizeY = 10
    Case 40 To 49
        StarProbe.DisplayOChipSizeY = 12
    Case 30 To 39
        StarProbe.DisplayOChipSizeY = 16
    Case 20 To 29
        StarProbe.DisplayOChipSizeY = 20
    Case 10 To 19
        StarProbe.DisplayOChipSizeY = 30
    Case Is < 10
        StarProbe.DisplayOChipSizeY = 60
    End Select
    
    Shape_OChip.width = StarProbe.DisplayOChipSizeX + 3
    Shape_OChip.Height = StarProbe.DisplayOChipSizeY + 3
    
    Shape_OFirstChip.width = StarProbe.DisplayOChipSizeX + 3
    Shape_OFirstChip.Height = StarProbe.DisplayOChipSizeY + 3
    
    Shape_OInk.width = StarProbe.DisplayOChipSizeX + 3
    Shape_OInk.Height = StarProbe.DisplayOChipSizeY + 3
    
    Shape_OMove.width = StarProbe.DisplayOChipSizeX + 3
    Shape_OMove.Height = StarProbe.DisplayOChipSizeY + 3
    
'    pOriginal.width = StarProbe.ChipCountX * StarProbe.DisplayOChipSizeX * 15
'    pOriginal.Height = StarProbe.ChipCountY * StarProbe.DisplayOChipSizeY * 15
'    pOriginal.Refresh

    pZoom.Cls
    pOriginal.Cls
    
    StarProbe.CountTotalChip = 0
    StarProbe.CountSkipDie = 0
    StarProbe.CountGoodDie = 0
    StarProbe.CountBadDie = 0
    Dim bin_idx As Integer
    For xx = 0 To StarProbe.ChipCountX
    
        x = xx
        
        For yy = 0 To StarProbe.ChipCountY
        
            y = yy
            
            If Wafer(x, y).Chip Then
            
                If Wafer(x, y).ChipMask Then         'skip die
                    DisplayChipColor = ChipColor(1)
                    StarProbe.CountSkipDie = StarProbe.CountSkipDie + 1
                ElseIf Wafer(x, y).ChipPlate Then    'skip die
                    Wafer(x, y).ChipSkipDie = False
                    Wafer(x, y).ChipInk = False
                    Wafer(x, y).ChipInk2 = False
                    DisplayChipColor = ChipColor(4)
                    StarProbe.CountSkipDie = StarProbe.CountSkipDie + 1
                ElseIf Not Wafer(x, y).ChipPlate And Wafer(x, y).ChipSkipDie Then   'skip die
                    If Wafer(x, y).ChipInk = True And Wafer(x, y).ChipInk2 = False Then
                        DisplayChipColor = ChipColor(5)
                    ElseIf Wafer(x, y).ChipInk = True And Wafer(x, y).ChipInk2 = True Then
                        DisplayChipColor = ChipColor(6)
                    Else
                        DisplayChipColor = ChipColor(3)
                    End If
                    StarProbe.CountSkipDie = StarProbe.CountSkipDie + 1
'                ElseIf Wafer(x, y).ChipMeasure Then
'                    DisplayChipColor = ChipColor(2)
                ElseIf Wafer(x, y).flag Then
                    DisplayChipColor = BINColor(Wafer(x, y).BIN)
                    
                    '[ 2020.12.16 ] : bincount
                    bin_idx = Wafer(x, y).BIN
                    Text_Bin_Count_No(bin_idx) = Text_Bin_Count_No(bin_idx) + 1
                Else
                    DisplayChipColor = ChipColor(0)
                    StarProbe.CountGoodDie = StarProbe.CountGoodDie + 1
                End If
                
                StarProbe.CountTotalChip = StarProbe.CountTotalChip + 1
                
                If StarProbe.WaferTest = 1 And Not WaferTest(x, y) Then DisplayChipColor = vbBlue
                
                pZoom.Line ((x * StarProbe.DisplayChipSizeX), (y * StarProbe.DisplayChipSizeY))- _
                           ((x * StarProbe.DisplayChipSizeX) + StarProbe.DisplayChipSizeX - 2, _
                            (y * StarProbe.DisplayChipSizeY) + StarProbe.DisplayChipSizeY - 2), DisplayChipColor, BF
                pOriginal.Line ((x * StarProbe.DisplayOChipSizeX), (y * StarProbe.DisplayOChipSizeY))- _
                               ((x * StarProbe.DisplayOChipSizeX) + StarProbe.DisplayOChipSizeX - 1, _
                                (y * StarProbe.DisplayOChipSizeY) + StarProbe.DisplayOChipSizeY - 1), DisplayChipColor, BF
            End If
            
        Next
        
    Next

    Shape_FirstChip.Top = (StarProbe.StartChip.y * StarProbe.DisplayChipSizeY) - 2
    Shape_FirstChip.Left = (StarProbe.StartChip.x * StarProbe.DisplayChipSizeX) - 2
    
    Shape_OFirstChip.Top = (StarProbe.StartChip.y * StarProbe.DisplayOChipSizeY) - 1
    Shape_OFirstChip.Left = (StarProbe.StartChip.x * StarProbe.DisplayOChipSizeX) - 1
    
    VScroll_Zoom.value = 0
    HScroll_Zoom.value = 0
    
    Exit Sub
    
ErrorHandler:  ' 오류 처리 루틴
    
    Select Case err.Number
    End Select
    
'    Resume  ' 오류를 유발한 줄과 동일한 줄에서 실행을 재개합니다.

    err.Clear

    Exit Sub

End Sub

Public Sub Display_Chip(pZoom As PictureBox, pOriginal As PictureBox, x As Integer, y As Integer)

    Dim ChipX As Integer, ChipY As Integer
    Dim DisplayChipColor As Long
                
    ChipX = x + StarProbe.StartChip.x
    ChipY = y + StarProbe.StartChip.y
    
'    If Wafer(ChipX, ChipY).Chip Then 'And Wafer(ChipX, ChipY).Flag Then
        
        If Wafer(ChipX, ChipY).ChipMask Then
            If Wafer(ChipX, ChipY).ChipInk = True And Wafer(ChipX, ChipY).ChipInk2 = False Then
                DisplayChipColor = ChipColor(5)
            ElseIf Wafer(ChipX, ChipY).ChipInk = True And Wafer(ChipX, ChipY).ChipInk2 = True Then
                DisplayChipColor = ChipColor(6)
            Else
                DisplayChipColor = ChipColor(1)
            End If
            'DisplayChipColor = ChipColor(1)
            
        ElseIf Wafer(ChipX, ChipY).ChipPlate Then
            DisplayChipColor = ChipColor(4)
        
        ElseIf Wafer(ChipX, ChipY).ChipSkipDie Then
            If Wafer(ChipX, ChipY).ChipInk = True And Wafer(ChipX, ChipY).ChipInk2 = False Then
                DisplayChipColor = ChipColor(5)
            ElseIf Wafer(ChipX, ChipY).ChipInk = True And Wafer(ChipX, ChipY).ChipInk2 = True Then
                DisplayChipColor = ChipColor(6)
            Else
                DisplayChipColor = ChipColor(3)
            End If
        
'        ElseIf Wafer(ChipX, ChipY).ChipMeasure Then
'            DisplayChipColor = ChipColor(2)
        
        ElseIf Wafer(ChipX, ChipY).flag Then
        If Wafer(ChipX, ChipY).BIN <= 24 Then DisplayChipColor = BINColor(Wafer(ChipX, ChipY).BIN)
        
        Else
        
            If Wafer(ChipX, ChipY).Chip Then
                DisplayChipColor = ChipColor(0)
            Else
                DisplayChipColor = vbWhite
            End If
            
        End If
        
        'DisplayChipColor = IIf(Wafer(ChipX, ChipY).FlagBad, BinColor(0), BinColor(Wafer(ChipX, ChipY).Bin))
        'DisplayChipColor = BinColor(Wafer(ChipX, ChipY).Bin)
        'DisplayChipColor = IIf(WaferTest(ChipX, ChipY), vbBlue, BinColor(0))
        
        pZoom.Line ((ChipX * StarProbe.DisplayChipSizeX), (ChipY * StarProbe.DisplayChipSizeY))- _
                   ((ChipX * StarProbe.DisplayChipSizeX) + StarProbe.DisplayChipSizeX - 2, _
                    (ChipY * StarProbe.DisplayChipSizeY) + StarProbe.DisplayChipSizeY - 2), DisplayChipColor, BF

        pOriginal.Line ((ChipX * StarProbe.DisplayOChipSizeX), (ChipY * StarProbe.DisplayOChipSizeY))- _
                       ((ChipX * StarProbe.DisplayOChipSizeX) + StarProbe.DisplayOChipSizeX, _
                        (ChipY * StarProbe.DisplayOChipSizeY) + StarProbe.DisplayOChipSizeY), DisplayChipColor, BF
'    End If

End Sub

'[ 2017.03.23 ] : DEMO Mode인 경우 INK시 black으로 표시해주는 부분
Public Sub Display_Chip_demo(pZoom As PictureBox, pOriginal As PictureBox, x As Integer, y As Integer)

    Dim ChipX As Integer, ChipY As Integer
    Dim DisplayChipColor As Long
                
    ChipX = x + StarProbe.StartChip.x
    ChipY = y + StarProbe.StartChip.y
    
    If Wafer(ChipX, ChipY).Chip Then 'And Wafer(ChipX, ChipY).Flag Then
        
        If Wafer(ChipX, ChipY).ChipMask Then
            DisplayChipColor = ChipColor(1)
            
        ElseIf Wafer(ChipX, ChipY).ChipPlate Then
            DisplayChipColor = ChipColor(4)
        
        ElseIf Wafer(ChipX, ChipY).ChipSkipDie Then
            If Wafer(ChipX, ChipY).ChipInk Or Wafer(ChipX, ChipY).ChipInk2 Then
                DisplayChipColor = vbBlack
            Else
                DisplayChipColor = ChipColor(3)
            End If

        
'        ElseIf Wafer(ChipX, ChipY).ChipMeasure Then
'            DisplayChipColor = ChipColor(2)
        
        ElseIf Wafer(ChipX, ChipY).flag Then
            If Wafer(ChipX, ChipY).BIN <= 24 Then DisplayChipColor = vbBlack
        Else
            DisplayChipColor = ChipColor(0)
            
        End If
        
        'DisplayChipColor = IIf(Wafer(ChipX, ChipY).FlagBad, BinColor(0), BinColor(Wafer(ChipX, ChipY).Bin))
        'DisplayChipColor = BinColor(Wafer(ChipX, ChipY).Bin)
        'DisplayChipColor = IIf(WaferTest(ChipX, ChipY), vbBlue, BinColor(0))
        
        pZoom.Line ((ChipX * StarProbe.DisplayChipSizeX), (ChipY * StarProbe.DisplayChipSizeY))- _
                   ((ChipX * StarProbe.DisplayChipSizeX) + StarProbe.DisplayChipSizeX - 2, _
                    (ChipY * StarProbe.DisplayChipSizeY) + StarProbe.DisplayChipSizeY - 2), DisplayChipColor, BF

        pOriginal.Line ((ChipX * StarProbe.DisplayOChipSizeX), (ChipY * StarProbe.DisplayOChipSizeY))- _
                       ((ChipX * StarProbe.DisplayOChipSizeX) + StarProbe.DisplayOChipSizeX, _
                        (ChipY * StarProbe.DisplayOChipSizeY) + StarProbe.DisplayOChipSizeY), DisplayChipColor, BF
    End If

End Sub

Public Sub Display_WaferTestChip(pZoom As PictureBox, pOriginal As PictureBox, x As Integer, y As Integer)

    Dim ChipX As Integer, ChipY As Integer
    Dim DisplayChipColor As Long
                
    ChipX = x + StarProbe.StartChip.x
    ChipY = y + StarProbe.StartChip.y
    
    If Wafer(ChipX, ChipY).Chip Then 'And Wafer(ChipX, ChipY).Flag Then
        
        DisplayChipColor = IIf(WaferTest(ChipX, ChipY), vbBlue, BINColor(0))
        
        pZoom.Line ((ChipX * StarProbe.DisplayChipSizeX), (ChipY * StarProbe.DisplayChipSizeY))- _
                   ((ChipX * StarProbe.DisplayChipSizeX) + StarProbe.DisplayChipSizeX - 2, _
                    (ChipY * StarProbe.DisplayChipSizeY) + StarProbe.DisplayChipSizeY - 2), DisplayChipColor, BF

        pOriginal.Line ((ChipX * StarProbe.DisplayOChipSizeX), (ChipY * StarProbe.DisplayOChipSizeY))- _
                       ((ChipX * StarProbe.DisplayOChipSizeX) + StarProbe.DisplayOChipSizeX, _
                        (ChipY * StarProbe.DisplayOChipSizeY) + StarProbe.DisplayOChipSizeY), DisplayChipColor, BF
    
    End If

End Sub

Public Function WaferImageSave(CommonDialog1 As CommonDialog, pZoom As PictureBox) As Boolean

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
End Function

Function ChipCount(D As Double) As Integer
    Dim s As String, s1 As String, s2 As String
    Dim i As Integer
    Dim r As Double
    
    s = Str(D)
    i = InStr(1, s, ".")
        
    If i > 0 Then
        s1 = Left(s, i - 1)
        s2 = Mid(s, i + 1)
        r = val(s1)
        If val(s2) > 0 Then r = r + 1
    Else
        r = D
    End If
    ChipCount = r
End Function

Function ChipMod(D As Double) As Double
    Dim s As String, s1 As String, s2 As String
    Dim i As Integer
    Dim r As Double
    
    s = Str(D)
    i = InStr(1, s, ".")
        
    s1 = Left(s, i - 1)
    s2 = Mid(s, i + 1)
    
    r = D - val(s1)
    
    ChipMod = r

End Function

''''''''''''''''''''
' Ink Dot Run
Public Function InkRun(xx As Integer, yy As Integer) As Boolean

'    Dim x As Integer, y As Integer
'    Dim b As Boolean
'
'    x = xx
'    y = yy
'
'    b = False
'
''    If Wafer(x, y).Chip And _
''       ((Wafer(x, y).flag And _
''         Wafer(x, y).FlagBad) Or _
''        (Not Wafer(x, y).ChipPlate And _
''         (Wafer(x, y).ChipMask Or _
''          (Wafer(x, y).ChipSkipDie)))) Then
''    If (Wafer(x, y).Chip And Wafer(x, y).flag And Wafer(x, y).FlagBad) Or _
''       (Not Wafer(x, y).ChipPlate And Wafer(x, y).ChipSkipDie And Wafer(x, y).ChipInk) Then
''        If Wafer(x, y).InkDot Then
''            b = False
''        Else
''            b = True
''        End If
'        'b = Not Wafer(x, y).InkDot
'    If (Wafer(x, y).Chip And Wafer(x, y).FlagBad) Or _
'       (Not Wafer(x, y).ChipPlate And Wafer(x, y).ChipSkipDie And Wafer(x, y).ChipInk) Then
'        b = True            '16.12.13
'    End If
'
'    InkRun = b

'[ 2017.03.23 ] : ink시 두가지 옵션으로 설정하도록 수정.
    Dim x As Integer, y As Integer
    Dim b As Boolean
    
    x = xx
    y = yy
    
    b = False
    
    If Ink_Start_Flag = 0 Then
        If Wafer(x, y).Chip And _
           ((Wafer(x, y).flag And Wafer(x, y).FlagBad) Or _
            (Not Wafer(x, y).ChipPlate And (Wafer(x, y).ChipMask Or _
              (Wafer(x, y).ChipSkipDie And Wafer(x, y).ChipInk)))) Then

            b = Not Wafer(x, y).InkDot
        End If
    Else
'        If (Wafer(x, y).Chip And Wafer(x, y).FlagBad) Or _
'            (Not Wafer(x, y).ChipPlate And Wafer(x, y).ChipSkipDie And Wafer(x, y).ChipInk) Then
        If Wafer(x, y).Chip And _
           ((Wafer(x, y).flag And Wafer(x, y).FlagBad) Or _
            (Not Wafer(x, y).ChipPlate And (Wafer(x, y).ChipMask Or _
              (Wafer(x, y).ChipSkipDie And Wafer(x, y).ChipInk)))) Then
'            b = Not Wafer(x, y).InkDot            '16.12.13
            b = Not Wafer(x, y).InkDot
        End If
    End If
    
    InkRun = b
End Function



Public Function InkRun_Left(xx As Integer, yy As Integer) As Boolean

'    Dim x As Integer, y As Integer
'    Dim b As Boolean
'
'    x = xx + StarProbe.StartChip.x - 1
'    y = yy + StarProbe.StartChip.y
'
'    b = False
'
''    If Wafer(x, y).Chip And _
''       ((Wafer(x, y).flag And _
''         Wafer(x, y).FlagBad) Or _
''        (Not Wafer(x, y).ChipPlate And _
''         (Wafer(x, y).ChipMask Or _
''          (Wafer(x, y).ChipSkipDie)))) Then
''        If Wafer(x, y).InkDot Then
''            b = False
''        Else
''            b = True
''        End If
'        'b = Not Wafer(x, y).InkDot
'    If (Wafer(x, y).Chip And Wafer(x, y).FlagBad) Or _
'       (Not Wafer(x, y).ChipPlate And Wafer(x, y).ChipSkipDie And Wafer(x, y).ChipInk) Then
'        b = True            '16.12.13
'    End If
'
'    InkRun_Left = b

'[ 2017.03.23 ] : ink시 두가지 옵션으로 설정하도록 수정.
    Dim x As Integer, y As Integer
    Dim b As Boolean
    
    x = xx + StarProbe.StartChip.x - 1
    y = yy + StarProbe.StartChip.y
    
    b = False
    
    If Ink_Start_Flag = 0 Then
        If Wafer(x, y).Chip And _
           ((Wafer(x, y).flag And Wafer(x, y).FlagBad) Or _
            (Not Wafer(x, y).ChipPlate And (Wafer(x, y).ChipMask Or _
              (Wafer(x, y).ChipSkipDie And Wafer(x, y).ChipInk)))) Then

            b = Not Wafer(x, y).InkDot
        End If
    Else
'        If (Wafer(x, y).Chip And Wafer(x, y).FlagBad) Or _
'            (Not Wafer(x, y).ChipPlate And Wafer(x, y).ChipSkipDie And Wafer(x, y).ChipInk) Then
        If Wafer(x, y).Chip And _
           ((Wafer(x, y).flag And Wafer(x, y).FlagBad) Or _
            (Not Wafer(x, y).ChipPlate And (Wafer(x, y).ChipMask Or _
              (Wafer(x, y).ChipSkipDie And Wafer(x, y).ChipInk)))) Then
            b = Not Wafer(x, y).InkDot
        End If
    End If
    
    InkRun_Left = b
End Function



Public Function InkRun_Center(xx As Integer, yy As Integer) As Boolean

'    Dim x As Integer, y As Integer
'    Dim b As Boolean
'
'    x = xx + StarProbe.StartChip.x
'    y = yy + StarProbe.StartChip.y
'
'    b = False
'
''    If Wafer(x, y).Chip And _
''       ((Wafer(x, y).flag And _
''         Wafer(x, y).FlagBad) Or _
''        (Not Wafer(x, y).ChipPlate And _
''         (Wafer(x, y).ChipMask Or _
''          (Wafer(x, y).ChipSkipDie)))) Then
''    If (Wafer(x, y).Chip And Wafer(x, y).flag And Wafer(x, y).FlagBad) Or _
''       (Not Wafer(x, y).ChipPlate And Wafer(x, y).ChipSkipDie And Wafer(x, y).ChipInk) Then
''        If Wafer(x, y).InkDot Then
''            b = False
''        Else
''            b = True
''        End If
'        'b = Not Wafer(x, y).InkDot
'    If (Wafer(x, y).Chip And Wafer(x, y).FlagBad) Or _
'       (Not Wafer(x, y).ChipPlate And Wafer(x, y).ChipSkipDie And Wafer(x, y).ChipInk) Then
'        b = True            '16.12.13
'    End If
'
'    InkRun_Center = b

'[ 2017.03.23 ] : ink시 두가지 옵션으로 설정하도록 수정.
    Dim x As Integer, y As Integer
    Dim b As Boolean
    
    x = xx + StarProbe.StartChip.x
    y = yy + StarProbe.StartChip.y
    
    b = False
    
    If Ink_Start_Flag = 0 Then
        If Wafer(x, y).Chip And _
           ((Wafer(x, y).flag And Wafer(x, y).FlagBad) Or _
            (Not Wafer(x, y).ChipPlate And (Wafer(x, y).ChipMask Or _
              (Wafer(x, y).ChipSkipDie And Wafer(x, y).ChipInk)))) Then

            b = Not Wafer(x, y).InkDot
        End If
    Else
'        If (Wafer(x, y).Chip And Wafer(x, y).FlagBad) Or _
'            (Not Wafer(x, y).ChipPlate And Wafer(x, y).ChipSkipDie And Wafer(x, y).ChipInk) Then
        If Wafer(x, y).Chip And _
           ((Wafer(x, y).flag And Wafer(x, y).FlagBad) Or _
            (Not Wafer(x, y).ChipPlate And (Wafer(x, y).ChipMask Or _
              (Wafer(x, y).ChipSkipDie And Wafer(x, y).ChipInk)))) Then

            b = Not Wafer(x, y).InkDot
        End If
    End If
    
    InkRun_Center = b
End Function


Public Function InkRun_Right(xx As Integer, yy As Integer) As Boolean
'    Dim x As Integer, y As Integer
'    Dim b As Boolean
'
'    x = xx + StarProbe.StartChip.x + 1
'    y = yy + StarProbe.StartChip.y
'
'    b = False
'
''    If Wafer(x, y).Chip And _
''       ((Wafer(x, y).flag And _
''         Wafer(x, y).FlagBad) Or _
''        (Not Wafer(x, y).ChipPlate And _
''         Wafer(x, y).ChipMask And _
''         Wafer(x, y).ChipSkipDie)) Then
''If (Wafer(x, y).Chip And Wafer(x, y).flag And Wafer(x, y).FlagBad) Or _
''       (Not Wafer(x, y).ChipPlate And Wafer(x, y).ChipSkipDie And Wafer(x, y).ChipInk) Then
''        If Wafer(x, y).InkDot Then
''            b = False
''        Else
''            b = True
''        End If
'        'b = Not Wafer(x, y).InkDot
'    If (Wafer(x, y).Chip And Wafer(x, y).FlagBad) Or _
'       (Not Wafer(x, y).ChipPlate And Wafer(x, y).ChipSkipDie And Wafer(x, y).ChipInk) Then
'        b = True            '16.12.13
'    End If
'
'    InkRun_Right = b

'[ 2017.03.23 ] : ink시 두가지 옵션으로 설정하도록 수정.
    Dim x As Integer, y As Integer
    Dim b As Boolean
    
    x = xx + StarProbe.StartChip.x + 1
    y = yy + StarProbe.StartChip.y
    
    b = False
    
    If Ink_Start_Flag = 0 Then
        If Wafer(x, y).Chip And _
           ((Wafer(x, y).flag And _
             Wafer(x, y).FlagBad) Or _
            (Not Wafer(x, y).ChipPlate And _
             Wafer(x, y).ChipMask And _
             Wafer(x, y).ChipSkipDie And _
             Wafer(x, y).ChipInk)) Then

            b = Not Wafer(x, y).InkDot
        End If
    Else
        If (Wafer(x, y).Chip And Wafer(x, y).FlagBad) Or _
            (Not Wafer(x, y).ChipPlate And Wafer(x, y).ChipSkipDie And Wafer(x, y).ChipInk) Then
            b = Not Wafer(x, y).InkDot
        End If
    End If
    
    InkRun_Right = b
End Function
' Ink Dot Run
''''''''''''''''''''



''''''''''''''''''''
' Ink Dot Run Ok
Public Sub InkRun_LeftOk(xx As Integer, yy As Integer)

    xx = xx + StarProbe.StartChip.x - 1
    yy = yy + StarProbe.StartChip.y
    
    Wafer(xx, yy).InkDot = True
    
End Sub

Public Sub InkRun_CenterOk(xx As Integer, yy As Integer)

    xx = xx + StarProbe.StartChip.x
    yy = yy + StarProbe.StartChip.y
    
    Wafer(xx, yy).InkDot = True

End Sub

Public Sub InkRun_RightOk(xx As Integer, yy As Integer)

    xx = xx + StarProbe.StartChip.x + 1
    yy = yy + StarProbe.StartChip.y
    
    Wafer(xx, yy).InkDot = True

End Sub
' Ink Dot Run Ok
''''''''''''''''''''

''''''''''''''''''''
' 작업 금지 구역
Public Function LimitAreaRight(x As Integer, y As Integer) As Boolean

    If StarProbe.LimitArea = 0 Then
        LimitAreaRight = True
        Exit Function
    End If

    Dim xx As Integer, yy As Integer
    Dim bX As Boolean, bY As Boolean
    Dim bReturn As Boolean
    
    bReturn = True
    
    xx = (StarProbe.MeasureStartX - x) * -1
    yy = (StarProbe.MeasureStartY - y) * -1
    
'    StarProbe.MeasureStepY = 1         '2018.02.06
    
    If yy >= StarProbe.MeasureStepY Then
        bReturn = False
'    ElseIf xx < 0 And yy < 0 Then
'        bReturn = True
'    ElseIf xx < StarProbe.MeasureStepX And yy < StarProbe.MeasureStepY Then
'        bReturn = True
'    ElseIf xx > 0 And yy < 0 Then
'        bReturn = True
 '   Else
 '       bReturn = False
    End If
    
    LimitAreaRight = True

End Function

Public Function LimitAreaLeft(x As Integer, y As Integer) As Boolean

    If StarProbe.LimitArea = 0 Then
        LimitAreaLeft = True
        Exit Function
    End If

    Dim xx As Integer, yy As Integer
    Dim bX As Boolean, bY As Boolean
    Dim bReturn As Boolean
    
    bReturn = True
    
    xx = (StarProbe.MeasureStartX - x)
    yy = (StarProbe.MeasureStartY - y) * -1
    
'    StarProbe.MeasureStepY = 1             '2018.02.06
    
    If yy >= StarProbe.MeasureStepY Then
        bReturn = False
'    ElseIf xx < 0 And yy < 0 Then
'        bReturn = True
'    ElseIf xx < StarProbe.MeasureStepX And yy < StarProbe.MeasureStepY Then
'        bReturn = True
'    ElseIf xx > 0 And yy < 0 Then
'        bReturn = True
'    Else
'       bReturn = False
    End If
    
    LimitAreaLeft = True

End Function
' 작업 금지 구역
''''''''''''''''''''

''''''''''''''''''''
' Wafer Direction
Public Sub WaferDirection(angle As Integer)

    Dim forx As Integer, fory As Integer
    Dim x As Integer, y As Integer
    
    Dim StarProbeTemp As tStarProbe
    
    Dim oldx As Integer, oldy As Integer
    
    Erase WaferTemp
    
    StarProbeTemp = StarProbe
    
    Select Case angle
    Case -90
    
        x = 0
        y = 0
        
        For fory = 0 To StarProbe.ChipCountY
        
            x = fory
    
            For forx = 0 To StarProbe.ChipCountX
        
                y = StarProbe.ChipCountX - forx
            
                WaferTemp(x, y) = Wafer(forx, fory)
            
                WaferTemp(x, y).Chip = Wafer(forx, fory).Chip
                WaferTemp(x, y).ChipMask = Wafer(forx, fory).ChipMask
                WaferTemp(x, y).ChipMeasure = Wafer(forx, fory).ChipMeasure
                WaferTemp(x, y).ChipSkipDie = Wafer(forx, fory).ChipSkipDie
                WaferTemp(x, y).ChipPlate = Wafer(forx, fory).ChipPlate
                WaferTemp(x, y).ChipInk = Wafer(forx, fory).ChipInk
                WaferTemp(x, y).ChipInk2 = Wafer(forx, fory).ChipInk2
                WaferTemp(x, y).BIN = Wafer(forx, fory).BIN
                WaferTemp(x, y).flag = Wafer(forx, fory).flag
                WaferTemp(x, y).FlagBad = Wafer(forx, fory).FlagBad
                WaferTemp(x, y).MeasureWait = Wafer(forx, fory).MeasureWait
                WaferTemp(x, y).InkDot = Wafer(forx, fory).InkDot
            
            Next
        
        Next
    
        StarProbe.CenterChipX = StarProbeTemp.CenterChipY
        StarProbe.CenterChipY = StarProbeTemp.CenterChipX
        
        StarProbe.ChipCountX = StarProbeTemp.ChipCountY
        StarProbe.ChipCountY = StarProbeTemp.ChipCountX
        
        StarProbe.ChipSizeX = StarProbeTemp.ChipSizeY
        StarProbe.ChipSizeY = StarProbeTemp.ChipSizeX
        
        StarProbe.DisplayChipSizeX = StarProbeTemp.DisplayChipSizeY
        StarProbe.DisplayChipSizeY = StarProbeTemp.DisplayChipSizeX
        
        StarProbe.StartChip.x = StarProbeTemp.StartChip.y
        StarProbe.StartChip.y = StarProbeTemp.StartChip.x
        
        StarProbe.Max.x = StarProbeTemp.Max.y
        StarProbe.Max.y = StarProbeTemp.Max.x
        
        StarProbe.Min.x = StarProbeTemp.Min.y
        StarProbe.Min.y = StarProbeTemp.Min.x
        
    Case 90
    
        y = 0
        
        'For forx = 0 To (StarProbe.ChipCountX - 1)
        For forx = 0 To StarProbe.ChipCountX
            
            'For fory = (StarProbe.ChipCountY - 1) To 0 Step -1
            For fory = StarProbe.ChipCountY To 0 Step -1
            
                'x = (StarProbe.ChipCountY - 1) - fory
                x = StarProbe.ChipCountY - fory
                
                'WaferTemp(x, forx) = Wafer(forx, fory)
                
                WaferTemp(x, forx).Chip = Wafer(forx, fory).Chip
                WaferTemp(x, forx).ChipMask = Wafer(forx, fory).ChipMask
                WaferTemp(x, forx).ChipMeasure = Wafer(forx, fory).ChipMeasure
                WaferTemp(x, forx).ChipSkipDie = Wafer(forx, fory).ChipSkipDie
                WaferTemp(x, forx).ChipPlate = Wafer(forx, fory).ChipPlate
                WaferTemp(x, forx).ChipInk = Wafer(forx, fory).ChipInk
                WaferTemp(x, forx).ChipInk2 = Wafer(forx, fory).ChipInk2
                WaferTemp(x, forx).BIN = Wafer(forx, fory).BIN
                WaferTemp(x, forx).flag = Wafer(forx, fory).flag
                WaferTemp(x, forx).FlagBad = Wafer(forx, fory).FlagBad
                WaferTemp(x, forx).MeasureWait = Wafer(forx, fory).MeasureWait
                WaferTemp(x, forx).InkDot = Wafer(forx, fory).InkDot
                
            Next
            
            y = y + 1
            
        Next
    
        StarProbe.CenterChipX = StarProbeTemp.CenterChipY
        StarProbe.CenterChipY = StarProbeTemp.CenterChipX
        
        StarProbe.ChipCountX = StarProbeTemp.ChipCountY
        StarProbe.ChipCountY = StarProbeTemp.ChipCountX
        
        StarProbe.ChipSizeX = StarProbeTemp.ChipSizeY
        StarProbe.ChipSizeY = StarProbeTemp.ChipSizeX
        
        StarProbe.DisplayChipSizeX = StarProbeTemp.DisplayChipSizeY
        StarProbe.DisplayChipSizeY = StarProbeTemp.DisplayChipSizeX
        
        StarProbe.StartChip.x = StarProbeTemp.StartChip.y
        StarProbe.StartChip.y = StarProbeTemp.StartChip.x
        
        StarProbe.Max.x = StarProbeTemp.Max.y
        StarProbe.Max.y = StarProbeTemp.Max.x
        
        StarProbe.Min.x = StarProbeTemp.Min.y
        StarProbe.Min.y = StarProbeTemp.Min.x
        
    Case 180
    
        'For fory = (StarProbe.ChipCountY - 1) To 0 Step -1
        For fory = StarProbe.ChipCountY To 0 Step -1
        
            'y = (StarProbe.ChipCountY - 1) - fory
            y = StarProbe.ChipCountY - fory
            
            'For forx = 0 To (StarProbe.ChipCountX - 1)
            For forx = 0 To StarProbe.ChipCountX - 1
            
                'x = (StarProbe.ChipCountX - 1) - forx
                x = StarProbe.ChipCountX - forx
                WaferTemp(x, y) = Wafer(forx, fory)
            Next
        Next
        
    Case 270
    
        'For forx = 0 To (StarProbe.ChipCountX - 1)
        For forx = 0 To StarProbe.ChipCountX
            'y = (StarProbe.ChipCountX - 1) - forx
            y = StarProbe.ChipCountX - forx
            'For fory = (StarProbe.ChipCountY - 1) To 0 Step -1
            For fory = StarProbe.ChipCountY To 0 Step -1
                x = fory
                WaferTemp(x, y) = Wafer(forx, fory)
            Next
        Next
        
        StarProbe.CenterChipX = StarProbeTemp.CenterChipY
        StarProbe.CenterChipY = StarProbeTemp.CenterChipX
        
        StarProbe.ChipCountX = StarProbeTemp.ChipCountY
        StarProbe.ChipCountY = StarProbeTemp.ChipCountX
        
        StarProbe.ChipSizeX = StarProbeTemp.ChipSizeY
        StarProbe.ChipSizeY = StarProbeTemp.ChipSizeX
        
        StarProbe.DisplayChipSizeX = StarProbeTemp.DisplayChipSizeY
        StarProbe.DisplayChipSizeY = StarProbeTemp.DisplayChipSizeX
        
        StarProbe.StartChip.x = StarProbeTemp.StartChip.y
        StarProbe.StartChip.y = StarProbeTemp.StartChip.x
        
        StarProbe.Max.x = StarProbeTemp.Max.y
        StarProbe.Max.y = StarProbeTemp.Max.x
        
        StarProbe.Min.x = StarProbeTemp.Min.y
        StarProbe.Min.y = StarProbeTemp.Min.x
        
    End Select
    
    Erase Wafer
    
    For fory = 0 To 900
        For forx = 0 To 900
            'Wafer(forx, fory) = WaferTemp(forx, fory)
            Wafer(forx, fory).Chip = WaferTemp(forx, fory).Chip
            Wafer(forx, fory).ChipMask = WaferTemp(forx, fory).ChipMask
            Wafer(forx, fory).ChipMeasure = WaferTemp(forx, fory).ChipMeasure
            Wafer(forx, fory).ChipSkipDie = WaferTemp(forx, fory).ChipSkipDie
            Wafer(forx, fory).ChipPlate = WaferTemp(forx, fory).ChipPlate
            Wafer(forx, fory).ChipInk = WaferTemp(forx, fory).ChipInk
            Wafer(forx, fory).ChipInk2 = WaferTemp(forx, fory).ChipInk2
            Wafer(forx, fory).BIN = WaferTemp(forx, fory).BIN
            Wafer(forx, fory).flag = WaferTemp(forx, fory).flag
            Wafer(forx, fory).FlagBad = WaferTemp(forx, fory).FlagBad
            Wafer(forx, fory).MeasureWait = WaferTemp(forx, fory).MeasureWait
            Wafer(forx, fory).InkDot = WaferTemp(forx, fory).InkDot
        Next
    Next
    
End Sub
' Wafer Direction
''''''''''''''''''''

Public Sub StarProbe_WorkDateTime_HMS(totaltime As Double)

    Dim T As Double
    Dim D As Double
    Dim h As Double
    Dim M As Double
    Dim s As Double
    
    T = totaltime
    
    h = T \ 3600
    T = T - (h * 3600)
    
    M = T \ 60
    T = T - (M * 60)
    
    s = T
    
    D = h \ 24
    h = h - (D * 24)
    
    StarProbe_WorkDateTime.D = D
    StarProbe_WorkDateTime.h = h
    StarProbe_WorkDateTime.M = M
    StarProbe_WorkDateTime.s = s
End Sub

Public Sub StarProbe_FileSave_Data(sfilename As String)
'On Error GoTo err

    If Dir(sfilename, vbNormal) <> "" Then
        Kill (sfilename)
        Sleep 100
    End If
    
    Dim ifreefile As Integer

    Dim x As Integer, y As Integer, i As Integer
    Dim s As String

    ifreefile = FreeFile
    
    Open sfilename For Output As ifreefile
    
    Print #ifreefile, "Star Probe v1 - Wafer Data"
    
    With StarProbe
    
        Print #ifreefile, .Min.x
        Print #ifreefile, .Min.y
        
        Print #ifreefile, .Max.x
        Print #ifreefile, .Max.y
        
        Print #ifreefile, .StartChip.x
        Print #ifreefile, .StartChip.y
        
        Print #ifreefile, .CurrentChip.x
        Print #ifreefile, .CurrentChip.y
        
        Print #ifreefile, .CenterChipX
        Print #ifreefile, .CenterChipY
    
        Print #ifreefile, .Unit
        Print #ifreefile, .InchUnit
        Print #ifreefile, .WaferSize
        Print #ifreefile, .WaferSizemm
        Print #ifreefile, .ChipSizeX
        Print #ifreefile, .ChipSizeY
        Print #ifreefile, .ChipCountX
        Print #ifreefile, .ChipCountY
    
        Print #ifreefile, .EdgeChipmm
        Print #ifreefile, .PlateZone
    
        Print #ifreefile, .DisplayChipSizeX
        Print #ifreefile, .DisplayChipSizeY
    
        Print #ifreefile, .Pattern_PositionCenter.x
        Print #ifreefile, .Pattern_PositionCenter.y
        Print #ifreefile, .Pattern_PositionTop.x
        Print #ifreefile, .Pattern_PositionTop.y
        Print #ifreefile, .Pattern_PositionBottom.x
        Print #ifreefile, .Pattern_PositionBottom.y
        Print #ifreefile, .Pattern_PositionLeft.x
        Print #ifreefile, .Pattern_PositionLeft.y
        Print #ifreefile, .Pattern_PositionRight.x
        Print #ifreefile, .Pattern_PositionRight.y
    
        Print #ifreefile, .Pattern_SizeCenter.x
        Print #ifreefile, .Pattern_SizeCenter.y
        Print #ifreefile, .Pattern_SizeTop.x
        Print #ifreefile, .Pattern_SizeTop.y
        Print #ifreefile, .Pattern_SizeBottom.x
        Print #ifreefile, .Pattern_SizeBottom.y
        Print #ifreefile, .Pattern_SizeLeft.x
        Print #ifreefile, .Pattern_SizeLeft.y
        Print #ifreefile, .Pattern_SizeRight.x
        Print #ifreefile, .Pattern_SizeRight.y
    
        Print #ifreefile, .ReMeasure
        Print #ifreefile, .LineOk
        Print #ifreefile, .RCount
        Print #ifreefile, .RCount_Sub
        Print #ifreefile, .MeasureSleep
    
        Print #ifreefile, .Ink_After
        Print #ifreefile, .Ink_After_LeftPort
        Print #ifreefile, .Ink_After_RightPort
        Print #ifreefile, .Ink_After_CenterPort
        Print #ifreefile, .Ink_LeftPort
        Print #ifreefile, .Ink_RightPort
    
        Print #ifreefile, .MeasureStepX
        Print #ifreefile, .MeasureStepY
        Print #ifreefile, .MeasureStartX
        Print #ifreefile, .MeasureStartY
    
        Print #ifreefile, .LimitArea
        Print #ifreefile, .WaferTest
        
        Print #ifreefile, .CountBadDie
        Print #ifreefile, .CountGoodDie
        Print #ifreefile, .CountSkipDie
        Print #ifreefile, .CountTotalChip
        
        Print #ifreefile, .FileName_Data
        Print #ifreefile, .FileName_Map
        Print #ifreefile, .FileName_MapOriginal
        Print #ifreefile, .FileName_MapZoom
        Print #ifreefile, .FIleName_MeasureResult
        
        Print #ifreefile, .WorkLastChipX
        Print #ifreefile, .WorkLastChipY
    
    End With
    
    Print #ifreefile, StarProbe_WorkDateTime_Total + DateDiff("S", StarProbe_WorkDateTime_From, StarProbe_WorkDateTime_To)
    Print #ifreefile, IIf(StarProbe.MeasureAll = 1 Or (XPitch(TT_NO) = 1 And YPitch(TT_NO) = 1), 0, StarProbe.MeasureStepX)
    Print #ifreefile, IIf(StarProbe.MeasureAll = 1 Or (XPitch(TT_NO) = 1 And YPitch(TT_NO) = 1), 0, StarProbe.MeasureStepY)
    
    ' StarProbe Information
    Print #ifreefile, ""
    Print #ifreefile, "########## StarProbe SORT SUMMARY ##########"
    Print #ifreefile, PROD.Test_PGM   ' Program
    Print #ifreefile, MT2000.Text1(0) ' Lot no
    Print #ifreefile, MT2000.Text1(1) ' item
    Print #ifreefile, MT2000.Text1(2) ' type
    Print #ifreefile, MT2000.Text1(3) ' machine no
    Print #ifreefile, MT2000.Text1(4) ' operator no
    
    Print #ifreefile, Test_Cnt
    Print #ifreefile, Good_Cnt
    

    Print #ifreefile, ""
    Print #ifreefile, "########## Wafer Data ##########"
    
    For y = 0 To 900
        For x = 0 To 900
            If Wafer(x, y).Chip Then
                s = "@" & STR_FIX(Str(x), 3) & STR_FIX(Str(y), 3)
                s = s & IIf(Wafer(x, y).Chip, 1, 0)
                s = s & IIf(Wafer(x, y).ChipMask, 1, 0)
                s = s & IIf(Wafer(x, y).ChipMeasure, 1, 0)
                s = s & IIf(Wafer(x, y).ChipSkipDie, 1, 0)
                s = s & IIf(Wafer(x, y).ChipPlate, 1, 0)
                s = s & IIf(Wafer(x, y).ChipInk, 1, 0)
                s = s & IIf(Wafer(x, y).ChipInk2, 1, 0)
                s = s & IIf(Wafer(x, y).flag, 1, 0)
                s = s & IIf(Wafer(x, y).FlagBad, 1, 0)
                s = s & IIf(Wafer(x, y).MeasureWait, 1, 0)
            '    s = s & IIf(Wafer(x, y).InkDot, 1, 0)          'multi probe
                s = s & IIf(Wafer(x, y).InkDot, 0, 0)
                s = s & Wafer(x, y).BIN
                Print #ifreefile, s
            End If
        Next
    Next
    
    Print #ifreefile, "########## ETC ##########"
    
                      '1234567890123456789012345678901
    Print #ifreefile, "#Wafer Divion                :" & StarProbe.WaferDivision  ' 2005.09.12
    
    Print #ifreefile, "########## BIN Color ##########"
    
    For i = 0 To 26
        Print #ifreefile, "#BIN Color                   :" & BINColor(i)
    Next
    
    Print #ifreefile, "########## Chip Color ##########"
    
    For i = 0 To 6
        Print #ifreefile, "#Chip Color                  :" & ChipColor(i)
    Next
    
    Print #ifreefile, "########## End Of File ##########"
    
    Close ifreefile
    Exit Sub
    
'err:
'    MsgBox "SP파일 저장시 에러가 발생하였습니다."
End Sub

Public Sub Starprobe_FileLoad_Data(sfilename As String)
'On Error GoTo err

    Dim ink2_flag As Boolean
    Dim first_find As Boolean
    
    first_find = False
    ink2_flag = False
    
    If Dir(sfilename, vbNormal) = "" Or UCase(Right(sfilename, 3)) <> ".SP" Then
        Exit Sub
    End If

    Dim ifreefile As Integer
    Dim sLine As String
    Dim x As Integer, y As Integer, i As Integer
    Dim waferx As Integer, wafery As Integer, b As Boolean
    Dim bETC As Boolean
    
    bETC = False
    
    ifreefile = FreeFile
    
    Open sfilename For Input As ifreefile
    
    Line Input #ifreefile, sLine  ' Version Information
     
    With StarProbe
    
        Line Input #ifreefile, sLine
        .Min.x = val(sLine)
        Line Input #ifreefile, sLine
        .Min.y = val(sLine)
        
        Line Input #ifreefile, sLine
        .Max.x = val(sLine)
        Line Input #ifreefile, sLine
        .Max.y = val(sLine)
        
        Line Input #ifreefile, sLine
        .StartChip.x = val(sLine)
        Line Input #ifreefile, sLine
        .StartChip.y = val(sLine)
        
        Line Input #ifreefile, sLine
        .CurrentChip.x = val(sLine)
        Line Input #ifreefile, sLine
        .CurrentChip.y = val(sLine)
        
        Line Input #ifreefile, sLine
        .CenterChipX = val(sLine)
        Line Input #ifreefile, sLine
        .CenterChipY = val(sLine)
    
        Line Input #ifreefile, sLine
        .Unit = val(sLine)
        Line Input #ifreefile, sLine
        .InchUnit = val(sLine)
        Line Input #ifreefile, sLine
        .WaferSize = val(sLine)
        Line Input #ifreefile, sLine
        .WaferSizemm = val(sLine)
        Line Input #ifreefile, sLine
        .ChipSizeX = val(sLine)
        Line Input #ifreefile, sLine
        .ChipSizeY = val(sLine)
        Line Input #ifreefile, sLine
        .ChipCountX = val(sLine)
        Line Input #ifreefile, sLine
        .ChipCountY = val(sLine)
    
        Line Input #ifreefile, sLine
        .EdgeChipmm = val(sLine)
        Line Input #ifreefile, sLine
        .PlateZone = val(sLine)
    
        Line Input #ifreefile, sLine
        .DisplayChipSizeX = val(sLine)
        Line Input #ifreefile, sLine
        .DisplayChipSizeY = val(sLine)
    
        Line Input #ifreefile, sLine
        .Pattern_PositionCenter.x = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_PositionCenter.y = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_PositionTop.x = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_PositionTop.y = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_PositionBottom.x = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_PositionBottom.y = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_PositionLeft.x = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_PositionLeft.y = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_PositionRight.x = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_PositionRight.y = val(sLine)
    
        Line Input #ifreefile, sLine
        .Pattern_SizeCenter.x = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_SizeCenter.y = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_SizeTop.x = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_SizeTop.y = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_SizeBottom.x = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_SizeBottom.y = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_SizeLeft.x = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_SizeLeft.y = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_SizeRight.x = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_SizeRight.y = val(sLine)
    
        Line Input #ifreefile, sLine
'        .ReMeasure = Val(sLine)
        Line Input #ifreefile, sLine
'        .LineOk = Val(sLine)
        Line Input #ifreefile, sLine
'        .RCount = Val(sLine)
        Line Input #ifreefile, sLine
'        .RCount_Sub = Val(sLine)
        Line Input #ifreefile, sLine
'        .MeasureSleep = Val(sLine)
    
       Line Input #ifreefile, sLine
'        .Ink_After = Val(sLine)
        Line Input #ifreefile, sLine
'        .Ink_After_LeftPort = Val(sLine)
        Line Input #ifreefile, sLine
'        .Ink_After_RightPort = Val(sLine)
        Line Input #ifreefile, sLine
'        .Ink_After_CenterPort = Val(sLine)
        Line Input #ifreefile, sLine
'        .Ink_LeftPort = Val(sLine)
        Line Input #ifreefile, sLine
'        .Ink_RightPort = Val(sLine)
    
        Line Input #ifreefile, sLine
'        .MeasureStepX = Val(sLine)
        Line Input #ifreefile, sLine
'        .MeasureStepY = Val(sLine)
        Line Input #ifreefile, sLine
'        .MeasureStartX = Val(sLine)
        Line Input #ifreefile, sLine
'        .MeasureStartY = Val(sLine)
    
        Line Input #ifreefile, sLine
        .LimitArea = val(sLine)
        Line Input #ifreefile, sLine
        .WaferTest = val(sLine)
        
        Line Input #ifreefile, sLine: .CountBadDie = val(sLine)
        Line Input #ifreefile, sLine: .CountGoodDie = val(sLine)
        Line Input #ifreefile, sLine: .CountSkipDie = val(sLine)
        Line Input #ifreefile, sLine: .CountTotalChip = val(sLine)
        
        Line Input #ifreefile, sLine: .FileName_Data = sLine
        Line Input #ifreefile, sLine: .FileName_Map = sLine
        Line Input #ifreefile, sLine: .FileName_MapOriginal = sLine
        Line Input #ifreefile, sLine: .FileName_MapZoom = sLine
        Line Input #ifreefile, sLine: .FIleName_MeasureResult = sLine
        
        Line Input #ifreefile, sLine: .WorkLastChipX = val(sLine)
        Line Input #ifreefile, sLine: .WorkLastChipY = val(sLine)
    
    End With
    
    Line Input #ifreefile, sLine: StarProbe_WorkDateTime_Total = val(sLine)
    
    Call StarProbe_WorkDateTime_HMS(StarProbe_WorkDateTime_Total)
    
    '[ 2017.03.23 ] : SP파일 로드시 기존의 카운트를 화면에 표시
    MT2000.SSPanel_GoodCount.Caption = StarProbe.CountGoodDie
    MT2000.SSPanel_BadCount.Caption = StarProbe.CountBadDie
    MT2000.SSPanel_SkipCount.Caption = StarProbe.CountSkipDie
    MT2000.SSPanel_TotalCount.Caption = StarProbe.CountTotalChip
    ''''''''''''''''
                  
    MT2000.SSPanel_DateTime.Caption = StarProbe_WorkDateTime.D & " Day " & _
                               StarProbe_WorkDateTime.h & ":" & _
                               StarProbe_WorkDateTime.M & ":" & _
                               StarProbe_WorkDateTime.s
                                   
    Line Input #ifreefile, sLine: StarProbe.MeasureStepX = val(sLine)
    Line Input #ifreefile, sLine: StarProbe.MeasureStepY = val(sLine)
    
    If StarProbe.MeasureStepX = 0 Or StarProbe.MeasureStepY = 0 Then
        StarProbe.MeasureAll = 1
        StarProbe.MeasureStepX = 1
        StarProbe.MeasureStepY = 1
    Else
        StarProbe.MeasureAll = 0
    End If
    
    XPitch(TT_NO) = StarProbe.MeasureStepX
    YPitch(TT_NO) = StarProbe.MeasureStepY
                                   
    Line Input #ifreefile, sLine
    Line Input #ifreefile, sLine
    
    Line Input #ifreefile, sLine: PROD.Test_PGM = sLine   ' Program
    Line Input #ifreefile, sLine: MT2000.Text1(0) = sLine ' Lot no
    Line Input #ifreefile, sLine: MT2000.Text1(1) = sLine ' item
    Line Input #ifreefile, sLine: MT2000.Text1(2) = sLine ' type
    Line Input #ifreefile, sLine: MT2000.Text1(3) = sLine ' machine no
    Line Input #ifreefile, sLine: MT2000.Text1(4) = sLine ' operator no
    
    Line Input #ifreefile, sLine: Test_Cnt = val(sLine)
    Line Input #ifreefile, sLine: Good_Cnt = val(sLine)
    
    
    ' Wafer data
    Line Input #ifreefile, sLine
    Line Input #ifreefile, sLine
    
    Erase Wafer
        
    Do While Not EOF(ifreefile)
        Line Input #ifreefile, sLine
        If Left(sLine, 1) = "@" Then
            If first_find = False Then
                first_find = True
                If ink2_flag = False Then
                    If Len(sLine) = 19 Then
                        ink2_flag = True
                    End If
                End If
            End If
            '         1         2
            '12345678901234567890
            '@61 0  10010000000
                        
            waferx = val(Mid(sLine, 2, 3))
            wafery = val(Mid(sLine, 5, 3))
            
            If ink2_flag = True Then
                b = IIf(val(Mid(sLine, 8, 1)) = 1, True, False):  Wafer(waferx, wafery).Chip = b
                b = IIf(val(Mid(sLine, 9, 1)) = 1, True, False):  Wafer(waferx, wafery).ChipMask = b
                If Wafer(waferx, wafery).ChipMask = True Then               '2018.06.01
                    Wafer(waferx, wafery).ChipMask_Backup = True
                Else
                    Wafer(waferx, wafery).ChipMask_Backup = False
                End If
                b = IIf(val(Mid(sLine, 10, 1)) = 1, True, False): Wafer(waferx, wafery).ChipMeasure = b
                b = IIf(val(Mid(sLine, 11, 1)) = 1, True, False): Wafer(waferx, wafery).ChipSkipDie = b
                If Wafer(waferx, wafery).ChipSkipDie = True Then               '2018.06.01
                    Wafer(waferx, wafery).ChipSkipDie_Backup = True
                Else
                    Wafer(waferx, wafery).ChipSkipDie_Backup = False
                End If
                b = IIf(val(Mid(sLine, 12, 1)) = 1, True, False): Wafer(waferx, wafery).ChipPlate = b
                b = IIf(val(Mid(sLine, 13, 1)) = 1, True, False): Wafer(waferx, wafery).ChipInk = b
                If Wafer(waferx, wafery).ChipInk = True Then               '2018.06.01
                    Wafer(waferx, wafery).ChipInk_Backup = True
                Else
                    Wafer(waferx, wafery).ChipInk_Backup = False
                End If
                b = IIf(val(Mid(sLine, 14, 1)) = 1, True, False): Wafer(waferx, wafery).ChipInk2 = b
                b = IIf(val(Mid(sLine, 15, 1)) = 1, True, False): Wafer(waferx, wafery).flag = b
                b = IIf(val(Mid(sLine, 16, 1)) = 1, True, False): Wafer(waferx, wafery).FlagBad = b
                b = IIf(val(Mid(sLine, 17, 1)) = 1, True, False): Wafer(waferx, wafery).MeasureWait = b
                b = IIf(val(Mid(sLine, 18, 1)) = 1, True, False): Wafer(waferx, wafery).InkDot = b
                i = val(Mid(sLine, 19, 3)):                       Wafer(waferx, wafery).BIN = i
            Else
                b = IIf(val(Mid(sLine, 8, 1)) = 1, True, False):  Wafer(waferx, wafery).Chip = b
                b = IIf(val(Mid(sLine, 9, 1)) = 1, True, False):  Wafer(waferx, wafery).ChipMask = b
                If Wafer(waferx, wafery).ChipMask = True Then               '2018.06.01
                    Wafer(waferx, wafery).ChipMask_Backup = True
                Else
                    Wafer(waferx, wafery).ChipMask_Backup = False
                End If
                b = IIf(val(Mid(sLine, 10, 1)) = 1, True, False): Wafer(waferx, wafery).ChipMeasure = b
                b = IIf(val(Mid(sLine, 11, 1)) = 1, True, False): Wafer(waferx, wafery).ChipSkipDie = b
                If Wafer(waferx, wafery).ChipSkipDie = True Then               '2018.06.01
                    Wafer(waferx, wafery).ChipSkipDie_Backup = True
                Else
                    Wafer(waferx, wafery).ChipSkipDie_Backup = False
                End If
                b = IIf(val(Mid(sLine, 12, 1)) = 1, True, False): Wafer(waferx, wafery).ChipPlate = b
                b = IIf(val(Mid(sLine, 13, 1)) = 1, True, False): Wafer(waferx, wafery).ChipInk = b
                If Wafer(waferx, wafery).ChipInk = True Then               '2018.06.01
                    Wafer(waferx, wafery).ChipInk_Backup = True
                Else
                    Wafer(waferx, wafery).ChipInk_Backup = False
                End If
                b = IIf(val(Mid(sLine, 14, 1)) = 1, True, False): Wafer(waferx, wafery).flag = b
                b = IIf(val(Mid(sLine, 15, 1)) = 1, True, False): Wafer(waferx, wafery).FlagBad = b
                b = IIf(val(Mid(sLine, 16, 1)) = 1, True, False): Wafer(waferx, wafery).MeasureWait = b
                b = IIf(val(Mid(sLine, 17, 1)) = 1, True, False): Wafer(waferx, wafery).InkDot = b
                i = val(Mid(sLine, 18, 3)):                       Wafer(waferx, wafery).BIN = i
            End If
        End If
        
        If sLine = "########## ETC ##########" Then
            bETC = True
            Exit Do
        End If
    Loop
    
    If bETC Then
        Line Input #ifreefile, sLine
        If Trim(Left(sLine, 30)) = "#Wafer Divion                :" Then
            StarProbe.WaferDivision = val(Mid(sLine, 31))
        End If
        
        Line Input #ifreefile, sLine
        For i = 0 To 26
            Line Input #ifreefile, sLine
            BINColor(i) = val(Mid(sLine, 31))
        Next
    
        Line Input #ifreefile, sLine
        For i = 0 To 6
            Line Input #ifreefile, sLine
            ChipColor(i) = val(Mid(sLine, 31))
        Next
    End If
    Close ifreefile
    Stop_Measure = False
    Exit Sub
    
'err:
'    Resume Next
End Sub

Public Sub Starprobe_FileLoad_Data_TXT(sfilename As String)
    If Dir(sfilename, vbNormal) = "" Or UCase(Right(sfilename, 4)) <> ".TXT" Then
        Exit Sub
    End If

    Dim ifreefile As Integer
    Dim sLine As String
    Dim x As Integer, y As Integer, i As Integer
    Dim waferx As Integer, wafery As Integer, b As Boolean
    Dim bETC As Boolean
    
    ''''''''''''''''''''''''
    Dim xcnt As Integer
    Dim ycnt As Integer
    ''''''''''''''''''''''''
    
    bETC = False
    
    ifreefile = FreeFile
    
    Open sfilename For Input As ifreefile
    
    Line Input #ifreefile, sLine  ' Version Information
    
    MT2000.Text1(0) = Mid(sLine, InStr(sLine, "Lot =>") + 7, InStr(sLine, "Wafer =>") - InStr(sLine, "Lot =>") - 7) 'lot
    MT2000.Text1(1) = Mid(sLine, InStr(sLine, "Product =>") + 11, InStr(sLine, "X =>") - InStr(sLine, "Product =>") - 11) 'lot
    With StarProbe
        .ChipCountX = val(Mid(sLine, InStr(sLine, "X =>") + 4, 3))
        .ChipCountY = val(Mid(sLine, InStr(sLine, "Y =>") + 4, 3))
        .Min.x = -200
        .Min.y = 0
        .Max.x = StarProbe.ChipCountX
        .Max.y = StarProbe.ChipCountY
    End With
    
    XPitch(TT_NO) = 1
    YPitch(TT_NO) = 1
    
    Erase Wafer
    
    Dim cx As Integer
    Dim cy As Integer
    Dim ic As Integer
            
    cx = 0
    cy = 1
       
    For ic = 1 To (StarProbe.ChipCountX * (StarProbe.ChipCountY + 1))
        If (Mid(sLine, ic, 1) = "1") Then           'pass
            Wafer(cx, cy - 1).Chip = True
            Wafer(cx, cy - 1).ChipMask = False
            Wafer(cx, cy - 1).ChipMeasure = False
            Wafer(cx, cy - 1).ChipSkipDie = False
            Wafer(cx, cy - 1).ChipPlate = False
            Wafer(cx, cy - 1).ChipInk = False
            Wafer(cx, cy - 1).ChipInk2 = False
            Wafer(cx, cy - 1).flag = True
            Wafer(cx, cy - 1).FlagBad = False
            Wafer(cx, cy - 1).MeasureWait = False
            Wafer(cx, cy - 1).InkDot = False
            Wafer(cx, cy - 1).BIN = 12
            cx = cx + 1
        ElseIf (Mid(sLine, ic, 1) = "-") Then
            Wafer(cx, cy - 1).Chip = False
            Wafer(cx, cy - 1).ChipMask = False
            Wafer(cx, cy - 1).ChipMeasure = False
            Wafer(cx, cy - 1).ChipSkipDie = False
            Wafer(cx, cy - 1).ChipPlate = False
            Wafer(cx, cy - 1).ChipInk = False
            Wafer(cx, cy - 1).ChipInk2 = False
            Wafer(cx, cy - 1).flag = False
            Wafer(cx, cy - 1).FlagBad = False
            Wafer(cx, cy - 1).MeasureWait = False
            Wafer(cx, cy - 1).InkDot = False
            Wafer(cx, cy - 1).BIN = 0
            cx = cx + 1
        ElseIf Mid(sLine, ic, 1) = "2" Then                                                       'fail
            Wafer(cx, cy - 1).Chip = True
            Wafer(cx, cy - 1).ChipMask = False
            Wafer(cx, cy - 1).ChipMeasure = False
            Wafer(cx, cy - 1).ChipSkipDie = False
            Wafer(cx, cy - 1).ChipPlate = False
            Wafer(cx, cy - 1).ChipInk = False
            Wafer(cx, cy - 1).ChipInk2 = False
            Wafer(cx, cy - 1).flag = True
            Wafer(cx, cy - 1).FlagBad = True
            Wafer(cx, cy - 1).MeasureWait = False
            Wafer(cx, cy - 1).InkDot = False
            Wafer(cx, cy - 1).BIN = 2
            cx = cx + 1
        ElseIf Mid(sLine, ic, 1) = "3" Then                                                       'fail
            Wafer(cx, cy - 1).Chip = True
            Wafer(cx, cy - 1).ChipMask = False
            Wafer(cx, cy - 1).ChipMeasure = False
            Wafer(cx, cy - 1).ChipSkipDie = False
            Wafer(cx, cy - 1).ChipPlate = False
            Wafer(cx, cy - 1).ChipInk = False
            Wafer(cx, cy - 1).ChipInk2 = False
            Wafer(cx, cy - 1).flag = True
            Wafer(cx, cy - 1).FlagBad = True
            Wafer(cx, cy - 1).MeasureWait = False
            Wafer(cx, cy - 1).InkDot = False
            Wafer(cx, cy - 1).BIN = 3
            cx = cx + 1
        ElseIf Mid(sLine, ic, 1) = "4" Then                                                       'fail
            Wafer(cx, cy - 1).Chip = True
            Wafer(cx, cy - 1).ChipMask = False
            Wafer(cx, cy - 1).ChipMeasure = False
            Wafer(cx, cy - 1).ChipSkipDie = False
            Wafer(cx, cy - 1).ChipPlate = False
            Wafer(cx, cy - 1).ChipInk = False
            Wafer(cx, cy - 1).ChipInk2 = False
            Wafer(cx, cy - 1).flag = True
            Wafer(cx, cy - 1).FlagBad = True
            Wafer(cx, cy - 1).MeasureWait = False
            Wafer(cx, cy - 1).InkDot = False
            Wafer(cx, cy - 1).BIN = 4
            cx = cx + 1
        ElseIf Mid(sLine, ic, 1) = "5" Then                                                       'fail
            Wafer(cx, cy - 1).Chip = True
            Wafer(cx, cy - 1).ChipMask = False
            Wafer(cx, cy - 1).ChipMeasure = False
            Wafer(cx, cy - 1).ChipSkipDie = False
            Wafer(cx, cy - 1).ChipPlate = False
            Wafer(cx, cy - 1).ChipInk = False
            Wafer(cx, cy - 1).ChipInk2 = False
            Wafer(cx, cy - 1).flag = True
            Wafer(cx, cy - 1).FlagBad = True
            Wafer(cx, cy - 1).MeasureWait = False
            Wafer(cx, cy - 1).InkDot = False
            Wafer(cx, cy - 1).BIN = 5
            cx = cx + 1
        ElseIf Mid(sLine, ic, 1) = "6" Then                                                       'fail
            Wafer(cx, cy - 1).Chip = True
            Wafer(cx, cy - 1).ChipMask = False
            Wafer(cx, cy - 1).ChipMeasure = False
            Wafer(cx, cy - 1).ChipSkipDie = False
            Wafer(cx, cy - 1).ChipPlate = False
            Wafer(cx, cy - 1).ChipInk = False
            Wafer(cx, cy - 1).ChipInk2 = False
            Wafer(cx, cy - 1).flag = True
            Wafer(cx, cy - 1).FlagBad = True
            Wafer(cx, cy - 1).MeasureWait = False
            Wafer(cx, cy - 1).InkDot = False
            Wafer(cx, cy - 1).BIN = 6
            cx = cx + 1
        ElseIf Mid(sLine, ic, 1) = "7" Then                                                       'fail
            Wafer(cx, cy - 1).Chip = True
            Wafer(cx, cy - 1).ChipMask = False
            Wafer(cx, cy - 1).ChipMeasure = False
            Wafer(cx, cy - 1).ChipSkipDie = False
            Wafer(cx, cy - 1).ChipPlate = False
            Wafer(cx, cy - 1).ChipInk = False
            Wafer(cx, cy - 1).ChipInk2 = False
            Wafer(cx, cy - 1).flag = True
            Wafer(cx, cy - 1).FlagBad = True
            Wafer(cx, cy - 1).MeasureWait = False
            Wafer(cx, cy - 1).InkDot = False
            Wafer(cx, cy - 1).BIN = 7
            cx = cx + 1
        ElseIf Mid(sLine, ic, 1) = "8" Then                                                       'fail
            Wafer(cx, cy - 1).Chip = True
            Wafer(cx, cy - 1).ChipMask = False
            Wafer(cx, cy - 1).ChipMeasure = False
            Wafer(cx, cy - 1).ChipSkipDie = False
            Wafer(cx, cy - 1).ChipPlate = False
            Wafer(cx, cy - 1).ChipInk = False
            Wafer(cx, cy - 1).ChipInk2 = False
            Wafer(cx, cy - 1).flag = True
            Wafer(cx, cy - 1).FlagBad = True
            Wafer(cx, cy - 1).MeasureWait = False
            Wafer(cx, cy - 1).InkDot = False
            Wafer(cx, cy - 1).BIN = 8
            cx = cx + 1
        ElseIf Mid(sLine, ic, 1) = "9" Then                                                       'fail
            Wafer(cx, cy - 1).Chip = True
            Wafer(cx, cy - 1).ChipMask = False
            Wafer(cx, cy - 1).ChipMeasure = False
            Wafer(cx, cy - 1).ChipSkipDie = False
            Wafer(cx, cy - 1).ChipPlate = False
            Wafer(cx, cy - 1).ChipInk = False
            Wafer(cx, cy - 1).ChipInk2 = False
            Wafer(cx, cy - 1).flag = True
            Wafer(cx, cy - 1).FlagBad = True
            Wafer(cx, cy - 1).MeasureWait = False
            Wafer(cx, cy - 1).InkDot = False
            Wafer(cx, cy - 1).BIN = 9
            cx = cx + 1
        ElseIf Mid(sLine, ic, 1) = "P" Then                                                       'fail
            Wafer(cx, cy - 1).Chip = True
            Wafer(cx, cy - 1).ChipMask = False
            Wafer(cx, cy - 1).ChipMeasure = False
            Wafer(cx, cy - 1).ChipSkipDie = False
            Wafer(cx, cy - 1).ChipPlate = False
            Wafer(cx, cy - 1).ChipInk = False
            Wafer(cx, cy - 1).ChipInk2 = False
            Wafer(cx, cy - 1).flag = True
            Wafer(cx, cy - 1).FlagBad = True
            Wafer(cx, cy - 1).MeasureWait = False
            Wafer(cx, cy - 1).InkDot = False
            Wafer(cx, cy - 1).BIN = 10
            cx = cx + 1
        ElseIf Mid(sLine, ic, 1) = "Q" Then                                                       'fail
            Wafer(cx, cy - 1).Chip = True
            Wafer(cx, cy - 1).ChipMask = False
            Wafer(cx, cy - 1).ChipMeasure = False
            Wafer(cx, cy - 1).ChipSkipDie = False
            Wafer(cx, cy - 1).ChipPlate = False
            Wafer(cx, cy - 1).ChipInk = False
            Wafer(cx, cy - 1).ChipInk2 = False
            Wafer(cx, cy - 1).flag = True
            Wafer(cx, cy - 1).FlagBad = True
            Wafer(cx, cy - 1).MeasureWait = False
            Wafer(cx, cy - 1).InkDot = False
            Wafer(cx, cy - 1).BIN = 11
            cx = cx + 1
        End If
        If ic Mod (StarProbe.ChipCountX + 1) = 0 Then
            cx = 0
            cy = cy + 1
        End If
    Next ic
        
    BINColor(0) = 4194432
    BINColor(1) = 16711808
    BINColor(2) = 33023
    BINColor(3) = 255
    BINColor(4) = 16744703
    BINColor(5) = 65536
    BINColor(6) = 4930692
    BINColor(7) = 16711680
    BINColor(8) = 8388736
    BINColor(9) = 32896
    BINColor(10) = 16576
    BINColor(11) = 6927536
    BINColor(12) = 65280
    BINColor(13) = 8454016
    BINColor(14) = 8712369
    BINColor(15) = 8454016
    BINColor(16) = 4194432
    
    Close ifreefile
    Stop_Measure = False
End Sub


Public Sub Starprobe_FileLoad_OldData(sfilename As String)
    If Dir(sfilename, vbNormal) = "" Or UCase(Right(sfilename, 3)) <> ".SP" Then
        Exit Sub
    End If

    Dim ifreefile As Integer
    Dim sLine As String
    
    Dim x As Integer, y As Integer, i As Integer
    
    Dim waferx As Integer, wafery As Integer, b As Boolean
    
    ifreefile = FreeFile
    
    Open sfilename For Input As ifreefile
    
    Line Input #ifreefile, sLine  ' Version Information
     
    With StarProbe
    
        Line Input #ifreefile, sLine
        .Min.x = val(sLine)
        Line Input #ifreefile, sLine
        .Min.y = val(sLine)
        
        Line Input #ifreefile, sLine
        .Max.x = val(sLine)
        Line Input #ifreefile, sLine
        .Max.y = val(sLine)
        
        Line Input #ifreefile, sLine
        .StartChip.x = val(sLine)
        Line Input #ifreefile, sLine
        .StartChip.y = val(sLine)
        
        Line Input #ifreefile, sLine
        .CurrentChip.x = val(sLine)
        Line Input #ifreefile, sLine
        .CurrentChip.y = val(sLine)
        
        Line Input #ifreefile, sLine
        .CenterChipX = val(sLine)
        Line Input #ifreefile, sLine
        .CenterChipY = val(sLine)
    
        Line Input #ifreefile, sLine
        .Unit = val(sLine)
        Line Input #ifreefile, sLine
        .InchUnit = val(sLine)
        Line Input #ifreefile, sLine
        .WaferSize = val(sLine)
        Line Input #ifreefile, sLine
        .WaferSizemm = val(sLine)
        Line Input #ifreefile, sLine
        .ChipSizeX = val(sLine)
        Line Input #ifreefile, sLine
        .ChipSizeY = val(sLine)
        Line Input #ifreefile, sLine
        .ChipCountX = val(sLine)
        Line Input #ifreefile, sLine
        .ChipCountY = val(sLine)
    
        Line Input #ifreefile, sLine
        .EdgeChipmm = val(sLine)
        Line Input #ifreefile, sLine
        .PlateZone = val(sLine)
    
        Line Input #ifreefile, sLine
        .DisplayChipSizeX = val(sLine)
        Line Input #ifreefile, sLine
        .DisplayChipSizeY = val(sLine)
    
        Line Input #ifreefile, sLine
        .Pattern_PositionCenter.x = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_PositionCenter.y = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_PositionTop.x = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_PositionTop.y = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_PositionBottom.x = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_PositionBottom.y = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_PositionLeft.x = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_PositionLeft.y = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_PositionRight.x = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_PositionRight.y = val(sLine)
    
        Line Input #ifreefile, sLine
        .Pattern_SizeCenter.x = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_SizeCenter.y = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_SizeTop.x = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_SizeTop.y = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_SizeBottom.x = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_SizeBottom.y = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_SizeLeft.x = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_SizeLeft.y = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_SizeRight.x = val(sLine)
        Line Input #ifreefile, sLine
        .Pattern_SizeRight.y = val(sLine)
    
        Line Input #ifreefile, sLine
'        .ReMeasure = Val(sLine)
        Line Input #ifreefile, sLine
'        .LineOk = Val(sLine)
        Line Input #ifreefile, sLine
'        .RCount = Val(sLine)
        Line Input #ifreefile, sLine
'        .RCount_Sub = Val(sLine)
        Line Input #ifreefile, sLine
'        .MeasureSleep = Val(sLine)
    
        Line Input #ifreefile, sLine
'        .Ink_After = Val(sLine)
        Line Input #ifreefile, sLine
'        .Ink_After_LeftPort = Val(sLine)
        Line Input #ifreefile, sLine
'        .Ink_After_RightPort = Val(sLine)
        Line Input #ifreefile, sLine
'        .Ink_LeftPort = Val(sLine)
        Line Input #ifreefile, sLine
'        .Ink_RightPort = Val(sLine)
    
        Line Input #ifreefile, sLine
'        .MeasureStepX = Val(sLine)
        Line Input #ifreefile, sLine
'        .MeasureStepY = Val(sLine)
        Line Input #ifreefile, sLine
'        .MeasureStartX = Val(sLine)
        Line Input #ifreefile, sLine
'        .MeasureStartY = Val(sLine)
    
        Line Input #ifreefile, sLine
        .LimitArea = val(sLine)
        Line Input #ifreefile, sLine
        .WaferTest = val(sLine)
        
        Line Input #ifreefile, sLine: .CountBadDie = val(sLine)
        Line Input #ifreefile, sLine: .CountGoodDie = val(sLine)
        Line Input #ifreefile, sLine: .CountSkipDie = val(sLine)
        Line Input #ifreefile, sLine: .CountTotalChip = val(sLine)
        
        Line Input #ifreefile, sLine: .FileName_Data = sLine
        Line Input #ifreefile, sLine: .FileName_Map = sLine
        Line Input #ifreefile, sLine: .FileName_MapOriginal = sLine
        Line Input #ifreefile, sLine: .FileName_MapZoom = sLine
        Line Input #ifreefile, sLine: .FIleName_MeasureResult = sLine
        
        Line Input #ifreefile, sLine: .WorkLastChipX = val(sLine)
        Line Input #ifreefile, sLine: .WorkLastChipY = val(sLine)
    
    End With
    
    Line Input #ifreefile, sLine: StarProbe_WorkDateTime_Total = val(sLine)
    
    Call StarProbe_WorkDateTime_HMS(StarProbe_WorkDateTime_Total)
                  
    MT2000.SSPanel_DateTime.Caption = StarProbe_WorkDateTime.D & " Day " & _
                               StarProbe_WorkDateTime.h & ":" & _
                               StarProbe_WorkDateTime.M & ":" & _
                               StarProbe_WorkDateTime.s
                                   
    Line Input #ifreefile, sLine: StarProbe.MeasureStepX = val(sLine)
    Line Input #ifreefile, sLine: StarProbe.MeasureStepY = val(sLine)
    
    If StarProbe.MeasureStepX = 0 Or StarProbe.MeasureStepY = 0 Then
        StarProbe.MeasureAll = 1
        StarProbe.MeasureStepX = 1
        StarProbe.MeasureStepY = 1
    Else
        StarProbe.MeasureAll = 0
    End If
    
    XPitch(TT_NO) = StarProbe.MeasureStepX
    YPitch(TT_NO) = StarProbe.MeasureStepY
                                   
    Line Input #ifreefile, sLine
    Line Input #ifreefile, sLine
    
    Line Input #ifreefile, sLine: PROD.Test_PGM = sLine   ' Program
    Line Input #ifreefile, sLine: MT2000.Text1(0) = sLine ' Lot no
    Line Input #ifreefile, sLine: MT2000.Text1(1) = sLine ' item
    Line Input #ifreefile, sLine: MT2000.Text1(2) = sLine ' type
    Line Input #ifreefile, sLine: MT2000.Text1(3) = sLine ' machine no
    Line Input #ifreefile, sLine: MT2000.Text1(4) = sLine ' operator no
    
    Line Input #ifreefile, sLine: Test_Cnt = val(sLine)
    Line Input #ifreefile, sLine: Good_Cnt = val(sLine)
    
    For i = 0 To 24
 '       Line Input #ifreefile, sline: Bin_Count(i) = Val(sline)
    Next
    
    ' Wafer data
    Line Input #ifreefile, sLine
    Line Input #ifreefile, sLine
    
    Erase Wafer
    
'    Type tWafer
'        Chip As Boolean
'        ChipMask As Boolean
'        ChipMeasure As Boolean
'        ChipSkipDie As Boolean
'        ChipPlate As Boolean
'        BIN As Byte
'        flag As Boolean         ' 측정 여부
'        FlagBad As Boolean      ' 측정 후 결과 양품이면 False, 불량이면 True
'        MeasureWait As Boolean  ' 측정 추적 알고리즘 줄을 세울때의 플래그
'        InkDot As Boolean
'    End Type
    
    Do While Not EOF(ifreefile)
    
        Line Input #ifreefile, sLine
        
        If Left(sLine, 1) = "@" Then
            '         1         2
            '12345678901234567890
            '@61 0  10010000000
            
            waferx = val(Mid(sLine, 2, 3))
            wafery = val(Mid(sLine, 5, 3))
            
            b = IIf(val(Mid(sLine, 8, 1)) = 1, True, False):  Wafer(waferx, wafery).Chip = b
            b = IIf(val(Mid(sLine, 9, 1)) = 1, True, False):  Wafer(waferx, wafery).ChipMask = b
            b = IIf(val(Mid(sLine, 10, 1)) = 1, True, False): Wafer(waferx, wafery).ChipMeasure = b
            b = IIf(val(Mid(sLine, 11, 1)) = 1, True, False): Wafer(waferx, wafery).ChipSkipDie = b
            b = IIf(val(Mid(sLine, 12, 1)) = 1, True, False): Wafer(waferx, wafery).ChipPlate = b
            b = IIf(val(Mid(sLine, 13, 1)) = 1, True, False): Wafer(waferx, wafery).flag = b
            b = IIf(val(Mid(sLine, 14, 1)) = 1, True, False): Wafer(waferx, wafery).FlagBad = b
            b = IIf(val(Mid(sLine, 15, 1)) = 1, True, False): Wafer(waferx, wafery).MeasureWait = b
            b = IIf(val(Mid(sLine, 16, 1)) = 1, True, False): Wafer(waferx, wafery).InkDot = b
            i = val(Mid(sLine, 17, 3)):                       Wafer(waferx, wafery).BIN = i
        End If
    Loop
    Close ifreefile
End Sub
