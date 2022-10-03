Attribute VB_Name = "Module_StarProbe_Move"
Option Explicit

'==============================================
' COMMAND
'==============================================

'[ Z-THETA CONTROL ]
' MT   : Move theta
' SM5  : Select Z travel mode
'       -> SM5En : n = Z travel mode
'           0 = LIMITS mode (limit to limit, Zup and Zdown)
'           1 = edge sensor
'           2 = Profiler
        
' SP6  : Set Z clearance
' SP7  : Set Z up limit
' SP8  : Set Z down limit
' SP9  : Set Z align height
'       ex)SP9Z2850 = 285.0mils
' SP10 : Set Z undertravel
' SP12 : Set Z scale factor
' ZD   : Move Z stage to down position
' ZM   : Move Z stage to specified height
' ZU   : Move Z stage to up position



'[ PROBE CLEANING/CONTINUITY TESTING ]
' CP   : Clean probe tips
' CT   : Perform continuity test

' SM12 : Set probe clean count
'   ex)SM12C5W1 = wafer5개마다 클린을 실시한다.

' SM50 : Enable probe tip scrub
'   ex)SM50R0 or SM50R1
'       스크럽 기능 활성화/비활성화: 프로브 어레이가 clean 패드에서 clean될 때마다 1-mil의 팔각형 모션이 발생합니다.
'       따라서 각 프로브 팁은 8개의 다른 방향에서 문질러집니다. 각 스크럽은 마모를 최소화하기 위해 패드의 다른 위치에서 수행됩니다.

' SM94 : CLN/CT Z 높이 설정을 활성화/비활성화합니다. 사용자가 프로브 팁 청소 높이 및 연속성 테스트 높이 설정을 비활성화할 수 있습니다.
' SM95 : CLN/CT Z 높이의 자동 조정을 활성화/비활성화합니다. 새 프로브 팁 접촉 높이가 설정될 때마다 사용자가 프로브 팁 깨끗한 높이와 연속성 테스트 높이를 자동으로 조정할 수 있습니다.













'[ AUTO ALIGN CONTROL ]
' AA   : Auto align wafer
' SM19 : Enable stop if auto align fails
' SM31 : Enable fine auto align (sub-pixel)
' SM53 : Set cognex Q value range



'[ BINNING/INKING CONTROL ]
' EI   : Enable/disable inkers
' IK   : Ink device
' SM10 : Egde die inking enable/disable
' SM23 : Assign 16 good die bins
' SM24 : Assign 16 logical ink codes
' SM29 : Enable skipdie inking
' SM52 : Select edge and skipdie inkers
' SP11 : Select inker delay
' SP20 : Set inker pulse width in mesc
' SP21 : Set inker counter limit
' SP22 : Set inker 1 counter
' SP23 : Set inker 2 counter
' SP24 : Set inker 3 counter
' SP25 : Set inker 4 counter
' SP26 : Reset inker time limit
' SP27 : Reset inker(s)

'[ BLOCK UP/DOWN LOAD ]
' DF   : Download micro list
' DP   : Download PRU data
' DQ   : Download learn list
' DR   : Download row list
' DS   : Download setup data
' UF   : Upload micro list
' UP   : Upload PRU data
' UQ   : Upload learn list
' UR   : Upload row list
' US   : Upload setup data

'[ DATA LOG CONTROL ]
' CB   : Clear printer data buffers
' LP   : Printer wafer log or cassette log
' RD   : Define "Device Type" string
' RL   : Define "Lot Number" string
' SM7  : Print error message enable/disable
' SM8  : Print wafer log enable/disable
' SM9  : Print cassette log enable/disable
' SM17 : Reset wafer number
' SM28 : Set printer format

'[ ID READER CONTROL ]
' BS   : Spin prealigner twise
' SM21 : Lot ID not identical
' SM25 : Define number of wafer ID read attempts
' SM26 : ID position angle
' SM27 : Bar code size angle
' SM37 : Set ID reader type
' SM38 : Set ID reader fail recovery

'[ LIST MANIPULATION ]
' AD   : Add point to learn list
' DE   : Delete point from  learn list
' FA   : Add micro site to micro list
' RC   : Add row to row list
'        Add column to column list
' RF   : Clear the micro list
' RR   : Clear the row/column list
' RS   : Reset learn list
' SM6  : Skipdie enable/disable
       
'[ MISCELLANEOUS ]
' CE   : Clear error buzzer and code
' DA   : Set the date
' EW   : Genarate E.O.W. Pulse on tester interface
' LA   : Lamp on/off
' ME   : Display a message to the operator
' SM15 : Enable 9 response message groups
' SM30 : Enable 30 Mil drop at load position
' SM40 : Enable screen/lamp saver
' SO   : Set option
' SP16 : Set align scan speed
' SP17 : Set AC line frequency(X)
' TI   : Set run time display clock
' TS   : Generate start test pulse on tester interface
' VA   : Chuck vacuum off/on



'[ PROBER OPERATION CONTROL ]
' AP   : Abort probing
' BA   : Begin autoprobing
' FC   : Microdie test complete
' PA   : Pause/Continue probing
' PR   : Probe one wafer
' TC   : Test complete and bin device
' WM   : Wafer mapping on/off

'[ PROBING PARAMETERS ]
' SM2  : Set initial probing direction
' SM4  : Select probe mode
' SM11 : Set cpprdinate quadrant
' SM14 : Enable continue at last die
' SM22 : Enable ignore vacuum
' SM32 : Set count pulse width
' SM43 : Enable dual die probing
' SM44 : Set dual die direction
' SP3  : Set matrix probe size
' SP13 : Set turnaround count
' SP14 : Set reprobe count
' SP15 : Set maximum row count
' SP19 : Set touchdown counter
' SP28 : Reset probing UP time

'[ PROFILER CONTROL ]
' PH   : Set probe height for profiler
' PZ   : Use profiler to find wafer thickness
' SM13 : Enable auto diameter measurement
' SM20 : Enable profiler and find center
' SM41 : Enhanced profiler
' SM42 : Profiling retries
' SP18 : Set air sensor X-Y position

'[ QUERY COMMAND ]
' ?A   : Request hot chuck information
' ?A0  : Requests temperature
' ?A1  : Requests setpoint
' ?A2  : Requests delay
' ?A3  : Reauest H.C. model

' ?C   : Request handler status and wafer ID
' ?E   : Request error code
' ?F   : Request micro coordinates
' ?H   : Request absolute motor position
' ?I   : Request first die position, wafer center position, and wafer diameter
' ?L   : Request multiprobe location code
' ?O   : Request current options settings
' ?P   : Request current XY position
' ?R   : Request state variables
' ?S   : Request status
' ?T   : Request theta position
' ?U   : reauest total probing up time
' ?W   : Request wafer ID
' ?Y   : Request yield data
' ?Z   : Request current Z height
' ID   : Request prober ID and S/W revision

'[ TEMPERATURE CONTROL ]
' SM33 : Set hot chuck temperature
' SM34 : enable early hot chuck recovery
' SM51 : Set hot chuck delay - SM51Rn(Rn : 0 ~ 16)
' SM54 : Set hoy chuck model type - SM54Rn(Rn : 1(EG hot chuck), 2(non-EG hot chuck))
' SX1  : Enable temperature compensation
' SX2  : Platen X axis coefficient of expansion
' SX3  : Platen Y axis coefficient of expansion
' SX4  : Wafer X axis coefficient of expansion
' SX5  : Wafer Y axis coefficient of expansion
' SX6  : Platen delta temperature
' SX7  : Wafer delta temperature

'[ WAFER DESCRIPTION ]
' SM1  : Set english or Metric motion

' SP1  : Set die size
'       ex)SP1X1234Y5678 = X(123.4mil or 1.234mm), Y(567.8mil or 5.678mm)

' SP29 : Set six digit die size
'   ex)SP29X345.25Y301.75 = X(345.25mil), Y(301.75mil)

' SP2  : Define preset (first) die coordinate

' SP4  : Set wafer diameter
'       ex)SP4D125 = 125mm

'[ WAFER HANDLING CONTROL ]
' HW   : Handle wafer
' LO   : Unload wafer and load new wafer
' RP   : Retry a failed prealign
' SM3  : Set flat orientation
' SM18 : Enable notch select
' SM39 : Enable wait before unload
' UL   : Unload wafer

'[ X-Y CONTROL ]
' FM   : Microdie move relative to micro origin
' GF   : Go to a micro site
' HO   : Move to home position
' MD   : Move relative in die steps
' MF   : Move to first (preset) position

' MM   : Move relative in machine steps
'       - 척은 이동 전에 내리고 이동 후에 올라갑니다.
'       - 양수와 음수 모두 사용이 가능합니다. (-999999 ~ 999999)
'       - 호스트가 각 축에서 0.1mil/2.5미크론의 기계 좌표에서 현재 위치를 기준으로 지정된 거리를 이동할 수 있습니다.

' MO   : Move absolute in die steps


'[ MESSAGE-UNSOLICITED ]
' AT   : Attention
' CO   : Continue
' MC   : Command coimpleted
' MF   : Command Failed
' PA   : Pause
' PC   : Pattern complete
' SP   : Start probing
' TC   : Signal test complete
' TS   : Test Start
'==============================================


Global GpibTmp As Long
Global GpibAdd As Integer       '2005.06.21 star probe
Global XYAxis As String         '2005.06.21 star probe
Global XAxis As Long            '2005.06.21 star probe
Global YAxis As Long            '2005.06.21 star probe
Global TimerCheck As Boolean    '2005.06.21 star probe

Global Opt_Select_Flag As Boolean
Global Stress As String * 80
Global XY_Demetion As String
Global Left_MM, Left_POS, LX_MM As String
Global Right_MM, Right_POS, RX_MM, Tatal_MM As String
Global bStop As Boolean
Global bStarProbe_Auto_Start As Boolean
Global bPause_Flag As Boolean
Global XPitch(50) As Integer
Global YPitch(50) As Integer
Global XPitch_MAIN As Integer               '2017.08.01 : main xpitch
Global YPitch_MAIN As Integer               '2017.08.01 : main ypitch
Global Pass_Yield As Double
Global Good_Die As Double
Global Bad_Die As Double
Global Test_End_Delay As ccrpStopWatch
Global Test_Start_Delay As ccrpStopWatch
Global WaitTime_Delay As ccrpStopWatch
Global File_Time_Delay As ccrpStopWatch
Global Input_Timer_Delay As ccrpStopWatch
Global RS232Time_Delay As ccrpStopWatch
Global Test_time_delay As ccrpStopWatch

Public Function Auto_Probing_Mode() As Boolean
    If DemoMode = 1 Then Exit Function
End Function

Public Function StarProbe_RS232_Output(x As String)
    Dim SO, SI As String

    SO = ""
    SI = ""
    SI = MT2000.MSComm1.Input
    SI = MT2000.MSComm1.Input
    SO = Trim(x)
    If ACCO = 1 Or Model_Select = 1 Then                            'ACCO only
        MT2000.MSComm1.Output = SO & vbCrLf
    Else
        MT2000.MSComm1.Output = SO & vbLf
    End If
End Function

Public Function StarProbe_RS232_Input() As String
    Dim SI, Str, s, P As String
    Dim Lencount As Double
 
    Set Input_Timer_Delay = New ccrpStopWatch
    Set RS232Time_Delay = New ccrpStopWatch
 
    RS232Time_Delay.Reset
 ' Text2.Text = ""
    Do
        If RS232Time_Delay.Elapsed > 150 Then Exit Do
    Loop
    Lencount = 0
    SI = ""
    Input_Timer_Delay.Reset
    
    If ACCO = 1 Or Model_Select = 1 Then                            'ACCO only
        Do
            P = MT2000.MSComm1.InBufferCount
                
            Str = MT2000.MSComm1.Input
            SI = SI & Str
            Lencount = Lencount + Trim(val(P))
            If SI <> Empty Then
                If ((Right(SI, 1) = vbCrLf) And Lencount > 2) Or InStr(1, SI, "MC") > 0 Then
                    SI = Trim(SI)
                    Exit Do
                Else
                    SI = Trim(SI)
                    RS232Time_Delay.Reset
                    Do
                        If RS232Time_Delay.Elapsed > 50 Then Exit Do
                    Loop
            
                    s = MT2000.MSComm1.Input
            
                    SI = SI & s
                    SI = Trim(SI)
                    If Right(SI, 1) = vbCrLf Or InStr(1, SI, "MC") > 0 Then
                        Exit Do
                    End If
                End If
            End If
            If Input_Timer_Delay.Elapsed > 6000 Then
                Exit Do
            End If
        Loop
        
        SI = Replace(SI, vbCrLf, "")
    Else
        Do
            P = MT2000.MSComm1.InBufferCount
            
            Str = MT2000.MSComm1.Input
            SI = SI & Str
            Lencount = Lencount + Trim(val(P))
            If SI <> Empty Then
                If ((Right(SI, 1) = vbLf) And Lencount > 2) Or InStr(1, SI, "MC") > 0 Then
                    SI = Trim(SI)
                    Exit Do
                Else
                    SI = Trim(SI)
                    RS232Time_Delay.Reset
                    Do
                        If RS232Time_Delay.Elapsed > 50 Then Exit Do
                    Loop
            
                    s = MT2000.MSComm1.Input
            
                    SI = SI & s
                    SI = Trim(SI)
                    If Right(SI, 1) = vbLf Or InStr(1, SI, "MC") > 0 Then
                        Exit Do
                    End If
                End If
            End If
            If Input_Timer_Delay.Elapsed > 6000 Then
                Exit Do
            End If
        Loop
        
        SI = Replace(SI, vbLf, "")
    End If
    
    SI = Replace(SI, ">", "")
    SI = Trim(SI)
 '   Text2.Text = SI
    StarProbe_RS232_Input = SI
    
    SI = MT2000.MSComm1.Input
    SI = MT2000.MSComm1.Input
    
    Set Input_Timer_Delay = Nothing
    Set RS232Time_Delay = Nothing
End Function

Public Function StarProbe_Step_X_Value() As String
    Dim SMot As String

    If DemoMode = 1 Then Exit Function

    Stress = " "

    ivprintf GpibAdd, "?P" + Chr$(10)                   '?P:first die를 잡은 후의 좌표를 리턴한다.
    iread GpibAdd, Stress, 100, 0&, GpibTmp
    Stress = Replace(Stress, vbCrLf, "")

    If InStr(Stress, "X") = 0 Then
        SMot = " "
    Else
        SMot = Trim(Mid(Stress, InStr(Stress, "X") + 1, InStr(Stress, "Y") - 2))
    End If

    StarProbe_Step_X_Value = SMot
End Function

Public Function Starprobe_Step_Y_Value() As String
    Dim SMot As String

    If DemoMode = 1 Then Exit Function

    Stress = " "

    ivprintf GpibAdd, "?P" + Chr$(10)                   '?P:first die를 잡은 후의 좌표를 리턴한다.
    iread GpibAdd, Stress, 100, 0&, GpibTmp
    Stress = Replace(Stress, vbCrLf, "")

    If InStr(Stress, "Y") = 0 Then
        SMot = ""
    Else
        SMot = Trim(Mid(Stress, InStr(Stress, "Y") + 1))
    End If

    Starprobe_Step_Y_Value = SMot
End Function

Public Function StarProbe_XY_Moving(X_Value As String, Y_Value As String)
    If DemoMode = 1 Then Exit Function

    Dim XYAxis As String
    Dim SO, SI As String

    If IO_2001X = 0 Then
        XYAxis = "MOX" & X_Value & "Y" & Y_Value
        ivprintf GpibAdd, XYAxis + Chr$(10)
    Else
        XYAxis = "MOX" & X_Value & "Y" & Y_Value
        SO = ""
        SI = ""
        SI = MT2000.MSComm1.Input
        SI = MT2000.MSComm1.Input
        SO = Trim(XYAxis)
        If ACCO = 1 Or Model_Select = 1 Then                            'ACCO only
            MT2000.MSComm1.Output = SO & vbCrLf
        Else
            MT2000.MSComm1.Output = SO & vbLf
        End If
    End If
End Function

Public Function StarProbe_XY_Position() As String
    Dim s As String
    If DemoMode = 1 Then Exit Function
   
    Stress = " "
    If IO_2001X = 0 Then
        ivprintf GpibAdd, "?P" + Chr$(10)
        iread GpibAdd, Stress, 100, 0&, GpibTmp
         
        StarProbe_XY_Position = Stress
    Else
        s = "?P"
        Call StarProbe_RS232_Output(s)
        Stress = StarProbe_RS232_Input
        StarProbe_XY_Position = Stress
    End If
End Function

'Clean/Continue Z 높이의 자동 조정을 활성화 한다.(manual 7.2.4참조)
Public Function StarProbe_Auto_ZHeight_Set()
    Dim Opt As String
   If DemoMode = 1 Then Exit Function

   Opt = "SM95B1" & Trim(Opt)                           '새 프로브 팁 접촉 높이가 설정될 때마다 사용자가 프로브 팁 Claen높이와 Continuity 테스트 높이를 자동으로 조정할 수 있습니다.
   ivprintf GpibAdd, Opt + Chr$(10)
End Function

'Z move
Public Function StarProbe_Move_Z(Z_Value As String)
    Dim XYAxis As String
    
    If DemoMode = 1 Then Exit Function

    XYAxis = "ZM" & Z_Value
    ivprintf GpibAdd, XYAxis + Chr$(10)
End Function

'현재위치를 스크럽 위치로 지정 X,Y,Z
Public Function StarProbe_Scrub_Position()
    If DemoMode = 1 Then Exit Function
    ivprintf GpibAdd, "PN" + Chr$(10)
End Function

'[현재위치를 Z profile height로 지정]

'높이가 키보드에서 수동으로 전송된 경우 재설정되므로 연속성 또는 청소 패드 높이를 재설정하지 않습니다.

'현재 Z 스테이지 높이를 프로브 팁 높이로 설정 및 저장(Z 이동 모드가 PROFILED로 설정되고 프로파일러가 활성화된 경우에만 유용)

'작업자가 이를 수행하도록 요구하는 대신 테스터가 프로브/패드 접점을 결정할 수 있습니다.
'테스터는 먼저 PZ 명령으로 웨이퍼를 프로파일링한 다음 ZM 명령을 사용하여 Z 모터 해상도에 따라 한 번에 1/2 또는 1/4 mil씩 스테이지를 이동하여 매번 연속성 테스트를 수행해야 합니다.
'테스터가 접촉이 이루어졌다고 판단하면 PH 명령을 발행하여 이 높이를 프로브 팁 높이로 설정해야 합니다.
'그런 다음 이 높이는 프로빙 중에 웨이퍼 두께와 Z 초과 이동을 보상하기 위한 "기본" 높이로 사용됩니다.
Public Function StarProbe_Profile_Height_Set()
    If DemoMode = 1 Then Exit Function
    ivprintf GpibAdd, "PH" + Chr$(10)
End Function

'Set Z align height
Public Function StarProbe_Align_Z(Z_Value As String)
    Dim XYAxis As String
    
    If DemoMode = 1 Then Exit Function

    XYAxis = "SP9" & Z_Value
    ivprintf GpibAdd, XYAxis + Chr$(10)
End Function

'Z up limit set
Public Function StarProbe_Z_Up_Limit(Z_Value As String)
    Dim XYAxis As String
    
    If DemoMode = 1 Then Exit Function

    XYAxis = "SP7" & Z_Value                    'SP7:Z up limit set
    ivprintf GpibAdd, XYAxis + Chr$(10)
End Function


'Z값을 읽어온다.
Public Function StarProbe_Motor_Z_Value() As String
    Dim SMot As String

    If DemoMode = 1 Then Exit Function

    Stress = " "

    ivprintf GpibAdd, "?Z" + Chr$(10)
    iread GpibAdd, Stress, 100, 0&, GpibTmp
    Stress = Replace(Stress, vbCrLf, "")

    If InStr(Stress, "Z") = 0 Then
        SMot = " "
    Else
        SMot = Trim(Mid(Stress, InStr(Stress, "Z") + 1))
    End If

    StarProbe_Motor_Z_Value = SMot
End Function

Public Function StarProbe_Z_Position() As String
    If DemoMode = 1 Then Exit Function

    Dim Zstr, s As String
    Dim SO, SI As String

    Stress = " "
    Zstr = " "
    
    If IO_2001X = 0 Then
        ivprintf GpibAdd, "?S" + Chr$(10)
        iread GpibAdd, Stress, 100, 0&, GpibTmp
    
        Stress = Replace(Stress, vbCrLf, "")
    
        If InStr(Stress, "Z") = 0 Then
            Zstr = " "
        Else
            Zstr = Trim(Mid(Stress, InStr(Stress, "Z") + 1, InStr(Stress, "W") - 3))
        End If
        StarProbe_Z_Position = Zstr
    Else
        Stress = " "
        Zstr = " "
        s = "?S"

        SO = ""
        SI = ""
        SI = MT2000.MSComm1.Input
        SI = MT2000.MSComm1.Input
        SO = Trim(s)
        If ACCO = 1 Or Model_Select = 1 Then                            'ACCO only
            MT2000.MSComm1.Output = SO & vbCrLf
        Else
            MT2000.MSComm1.Output = SO & vbLf
        End If
        Stress = StarProbe_RS232_Input
        If InStr(Stress, "Z") = 0 Then
            Zstr = " "
        Else
            Zstr = Trim(Mid(Stress, InStr(Stress, "Z") + 1, InStr(Stress, "W") - 3))
        End If
        StarProbe_Z_Position = Zstr
    End If
End Function

Public Function Starprobe_Edge_Check() As Integer
    Dim Estr  As Integer
    Dim s As String
    
    If DemoMode = 1 Then Exit Function

    Stress = " "
    If IO_2001X = 0 Then
        
        Call iclear(GpibAdd)
    
        ivprintf GpibAdd, "?S" + Chr$(10)
        iread GpibAdd, Stress, 100, 0&, GpibTmp
    
        Stress = Replace(Stress, vbCrLf, "")
        If InStr(Stress, "C") = 0 Then
            Estr = 3
        Else
            Estr = val(Trim(Mid(Stress, InStr(Stress, "C") + 1)))
        End If
        Starprobe_Edge_Check = Estr
    Else
        s = "?S"
        Call StarProbe_RS232_Output(s)
        Stress = StarProbe_RS232_Input
        If InStr(Stress, "C") = 0 Then
            Estr = 3
        Else
            Estr = val(Trim(Mid(Stress, InStr(Stress, "C") + 1)))
        End If
        Starprobe_Edge_Check = Estr
    End If
End Function

Public Function StarProbe_Address() As Integer
    Dim Add As String

    If DemoMode = 1 Then Exit Function
    
    If IO_2001X = 0 Then
        'Add = iopen("lan[LCRY0301J11646]:inst0")
        Add = iopen("gpib0,1")
        
        If Add = -7 Then
            Add = iopen("gpib0,1")
        End If
        
        If Add = -7 Then
              MsgBox "Cannot find HPIB Card !" & vbCrLf & _
                     "To use 2001X " & vbCrLf & _
                     "Connect HPIB Card and Restart this Program!!", 16        '수정
              Exit Function
        End If
        Call itimeout(Add, 100000)
                   
        StarProbe_Address = Add
    Else
        If Add = -7 Then
            MsgBox "Cannot find HPIB Card !" & vbCrLf & _
                    "To use 2001X " & vbCrLf & _
                    "Connect HPIB Card and Restart this Program!!", 16        '수정
            Exit Function
        End If
        StarProbe_Address = Add
    End If
End Function

Public Function StarProbe_Option_Set(Opt As String)
    If DemoMode = 1 Then Exit Function
    
    If IO_2001X = 0 Then
        Opt = "SM15M" & Trim(Opt)
        ivprintf GpibAdd, Opt + Chr$(10)
    Else
        Opt = "SM15M" & Trim(Opt)
        Call StarProbe_RS232_Output(Opt)
    End If
End Function

Public Function StarProbe_Wafer_Size() As Long
    Dim Wstr As Long
    Dim s As String

    If DemoMode = 1 Then Exit Function
    
    Stress = " "
    If IO_2001X = 0 Then
        ivprintf GpibAdd, "?I" + Chr$(10)
        iread GpibAdd, Stress, 100, 0&, GpibTmp
        Stress = Replace(Stress, vbCrLf, "")
        
        If InStr(Stress, "D") = 0 Then
            Wstr = 0
        Else
            Wstr = Trim(Mid(Stress, InStr(Stress, "D") + 1))
        End If
        StarProbe_Wafer_Size = Wstr
    Else
        s = "I"
        Call StarProbe_RS232_Output(s)
        Stress = StarProbe_RS232_Input
        If InStr(Stress, "D") = 0 Then
            Wstr = 0
        Else
            Wstr = Trim(Mid(Stress, InStr(Stress, "D") + 1))
        End If
        StarProbe_Wafer_Size = Wstr
    End If
End Function

Public Function StarProbe_Chuck_X_Center() As String
    Dim Sstr, s As String

    If DemoMode = 1 Then Exit Function
    
    Stress = " "
    Sstr = " "
    If IO_2001X = 0 Then
        ivprintf GpibAdd, "?I" + Chr$(10)
        iread GpibAdd, Stress, 100, 0&, GpibTmp
        
        Stress = Replace(Stress, vbCrLf, "")
        If InStr(Stress, "X") = 0 Or InStr(Stress, "Y") = 0 Then
            Sstr = " "
        Else
            Sstr = Mid(Stress, InStr(Stress, "Y") + 1)
            Sstr = Mid(Stress, 1, InStr(Stress, "Y") - 1)
            Sstr = Trim(Mid(Stress, InStr(Stress, "X") + 1))
        End If
        StarProbe_Chuck_X_Center = Sstr
    Else
        s = "?I"
        Call StarProbe_RS232_Output(s)
        Stress = StarProbe_RS232_Input
        If InStr(Stress, "X") = 0 Or InStr(Stress, "Y") = 0 Then
            Sstr = " "
        Else
            Sstr = Mid(Stress, InStr(Stress, "Y") + 1)
            Sstr = Mid(Stress, 1, InStr(Stress, "Y") - 1)
            Sstr = Trim(Mid(Stress, InStr(Stress, "X") + 1))
        End If
        StarProbe_Chuck_X_Center = Sstr
    End If
End Function

Public Function StarProbe_First_X() As String
    Dim Sstr As String
    Dim s As String
    
    Stress = " "
    Sstr = " "

    If DemoMode = 1 Then Exit Function

    If IO_2001X = 0 Then
        Call iclear(GpibAdd)
     
        ivprintf GpibAdd, "?I" + Chr$(10)
        ' Sleep 20
        iread GpibAdd, Stress, 1000, 0&, GpibTmp
     
        Stress = Replace(Stress, vbCrLf, "")
        If InStr(Stress, "X") = 0 Then
            Sstr = " "
        Else
            Sstr = Mid(Stress, InStr(Stress, "X") + 1)
            Sstr = Mid(Sstr, 1, InStr(Sstr, "X") - 1)
            Sstr = Mid(Sstr, 1, InStr(Sstr, "Y") - 1)
            Sstr = Replace(Sstr, vbCrLf, "")
        End If
        StarProbe_First_X = Sstr
    Else
        s = "?I"
  
        Call StarProbe_RS232_Output(s)
   
        Stress = StarProbe_RS232_Input
  
        If InStr(Stress, "X") = 0 Then
            Sstr = " "
        Else
            Sstr = Mid(Stress, InStr(Stress, "X") + 1)
            Sstr = Mid(Sstr, 1, InStr(Sstr, "X") - 1)
            Sstr = Mid(Sstr, 1, InStr(Sstr, "Y") - 1)
            Sstr = Replace(Sstr, vbCrLf, "")
        End If
        StarProbe_First_X = Sstr
    End If
End Function

Public Function StarProbe_First_Y() As String
    Dim Sstr, s As String

    Stress = " "
    Sstr = " "
    
    If DemoMode = 1 Then Exit Function

    If IO_2001X = 0 Then
        Stress = " "
        Sstr = " "
        Call iclear(GpibAdd)
        
        ivprintf GpibAdd, "?I" + Chr$(10)
        
        iread GpibAdd, Stress, 100, 0&, GpibTmp
        
        Stress = Replace(Stress, vbCrLf, "")
        
        If InStr(Stress, "Y") = 0 Then
            Sstr = " "
        Else
            Sstr = Mid(Stress, InStr(Stress, "Y") + 1)
            Sstr = Mid(Sstr, 1, InStr(Sstr, "X") - 1)
            Sstr = Replace(Sstr, vbCrLf, "")
        End If
        
        StarProbe_First_Y = Sstr
    Else
        s = "?I"
        Call StarProbe_RS232_Output(s)
        Stress = StarProbe_RS232_Input
    
        If InStr(Stress, "Y") = 0 Then
            Sstr = " "
        Else
            Sstr = Mid(Stress, InStr(Stress, "Y") + 1)
            Sstr = Mid(Sstr, 1, InStr(Sstr, "X") - 1)
            Sstr = Replace(Sstr, vbCrLf, "")
        End If
        StarProbe_First_Y = Sstr
    End If
End Function

Public Function StarProbe_chuck_Y_Center() As String
    Dim Sstr, s As String

    If DemoMode = 1 Then Exit Function

    Stress = " "
    Sstr = " "
    
    If IO_2001X = 0 Then
        Call iclear(GpibAdd)
        
        ivprintf GpibAdd, "?I" + Chr$(10)
        iread GpibAdd, Stress, 100, 0&, GpibTmp
        Stress = Replace(Stress, vbCrLf, "")
        
        If InStr(Stress, "Y") = 0 Then
            Sstr = " "
        Else
            Sstr = Mid(Stress, InStr(Stress, "Y") + 1)
            Sstr = Mid(Stress, 1, InStr(Stress, "D") - 1)
            Sstr = Trim(Mid(Sstr, InStr(Sstr, "Y") + 1))
        End If
        
        StarProbe_chuck_Y_Center = Sstr
    Else
        s = "?I"
        Call StarProbe_RS232_Output(s)
        Stress = StarProbe_RS232_Input
    
        If InStr(Stress, "Y") = 0 Then
            Sstr = " "
        Else
            Sstr = Mid(Stress, InStr(Stress, "Y") + 1)
            Sstr = Mid(Stress, 1, InStr(Stress, "D") - 1)
            Sstr = Trim(Mid(Sstr, InStr(Sstr, "Y") + 1))
        End If
        StarProbe_chuck_Y_Center = Sstr
    End If
End Function

Public Function StarProbe_Z_UP()
    If DemoMode = 1 Then Exit Function
    ivprintf GpibAdd, "ZU" + Chr$(10)
End Function

Public Function StarProbe_Z_Down()
    Dim bpoint As Boolean
    Dim s As String

    If DemoMode = 1 Then Exit Function
    
    If IO_2001X = 0 Then
        ivprintf GpibAdd, "ZD" + Chr$(10)
    Else
        s = "ZU"
        Call StarProbe_RS232_Output(s)
    End If
End Function

'lamp off
Public Function StarProbe_LAMP_OFF()  ' kks laser
    Dim s As String
    
    If DemoMode = 1 Then Exit Function
    
    If IO_2001X = 0 Then
        ivprintf GpibAdd, "LAL0" + Chr$(10)
    Else
        s = "LAL0"
        Call StarProbe_RS232_Output(s)
    End If
End Function

Public Function StarProbe_FistDie()
    Dim s As String
    If DemoMode = 1 Then Exit Function
    
    If IO_2001X = 0 Then
        ivprintf GpibAdd, "MF" + Chr$(10)
    Else
        s = "MF"
        Call StarProbe_RS232_Output(s)
    End If
End Function

Public Function StarProbe_Pause()
    Dim s As String
    
    If DemoMode = 1 Then Exit Function
    
    If IO_2001X = 0 Then
        ivprintf GpibAdd, "PA" + Chr$(10)
    Else
        s = "PA"
        Call StarProbe_RS232_Output(s)
    End If
End Function

Public Function StarProbe_Continue()
    Dim s As String
    
    If DemoMode = 1 Then Exit Function
    If IO_2001X = 0 Then
        ivprintf GpibAdd, "CO" + Chr$(10)
    Else
        s = "CO"
        Call StarProbe_RS232_Output(s)
    End If
End Function

Public Function StarProbe_First_Chip()
    Dim s As String
    
    If DemoMode = 1 Then Exit Function
    If IO_2001X = 0 Then
        ivprintf GpibAdd, "FD" + Chr$(10)
    Else
        s = "FD"
        Call StarProbe_RS232_Output(s)
    End If
End Function

Public Function StarProbe_Tip_center()
    Dim s As String
    
    If DemoMode = 1 Then Exit Function
    If IO_2001X = 0 Then
        ivprintf GpibAdd, "MP" + Chr$(10)
    Else
        s = "MP"
        Call StarProbe_RS232_Output(s)
    End If
End Function

Public Function StarProbe_tip_clean()
    Dim s As String
    
    If DemoMode = 1 Then Exit Function
    If IO_2001X = 0 Then
        ivprintf GpibAdd, "CP" + Chr$(10)
    Else
        s = "CP"
        Call StarProbe_RS232_Output(s)
    End If
End Function

Public Function StarProbe_Motor_Home()
    Dim s As String
    
    If DemoMode = 1 Then Exit Function
    If IO_2001X = 0 Then
        ivprintf GpibAdd, "HO" + Chr$(10)
    Else
        s = "HO"
        Call StarProbe_RS232_Output(s)
    End If
End Function

Public Function StarProbe_Unload_Wafer()
    Dim s As String
    
    If DemoMode = 1 Then Exit Function
    If IO_2001X = 0 Then
        ivprintf GpibAdd, "UL" + Chr$(10)
    Else
        s = "UL"
        Call StarProbe_RS232_Output(s)
    End If
End Function

Public Function StarProbe_Unlode_Wafer_New_Wafer()  ' kks laser
    Dim s As String
    
    If DemoMode = 1 Then Exit Function
    If IO_2001X = 0 Then
        ivprintf GpibAdd, "LO" + Chr$(10)
    Else
        s = "LO"
        Call StarProbe_RS232_Output(s)
    End If
End Function

Public Function StarProbe_Profile_Wafer_Thickness()   'kks laser
    Dim s As String
    
    If DemoMode = 1 Then Exit Function
    If IO_2001X = 0 Then
        ivprintf GpibAdd, "PZ" + Chr$(10)
    Else
        s = "PZ"
        Call StarProbe_RS232_Output(s)
    End If
End Function

Public Function StarProbe_Auto_Align_Wafer(D As Integer)  ' kks laser
    Dim Str As String
 
    Str = ""
    Str = "AA" & D
    If DemoMode = 1 Then Exit Function
    
    If IO_2001X = 0 Then
        ivprintf GpibAdd, Str + Chr$(10)
    Else
        Call StarProbe_RS232_Output(Str)
    End If
End Function

Public Function StarProbe_Set_Die_Size(x As Double, y As Double)
    Dim lX, lY As Long
    Dim SXY As String

    If DemoMode = 1 Then Exit Function

    SXY = "SP29X" & Trim(Str(x)) & "Y" & Trim(Str(y))
    ivprintf GpibAdd, SXY + Chr$(10)
End Function

Public Function StarProbe_Set_Wafer_Diameter(D As Long)
    Dim lD As Long
    Dim SD As String

    If DemoMode = 1 Then Exit Function

    lD = D
    SD = "SP4D" & Trim(Str(lD))
    If IO_2001X = 0 Then
        ivprintf GpibAdd, SD + Chr$(10)
    Else
        Call StarProbe_RS232_Output(SD)
    End If
End Function

Public Function StarProbe_Ink_Dot(Number As Integer)
    Dim Snum As String

    If DemoMode = 1 Then Exit Function

    Snum = Trim(Str(Number))
    Snum = "IK" & Snum
    If IO_2001X = 0 Then
        ivprintf GpibAdd, Snum + Chr$(10)
    Else
        Call StarProbe_RS232_Output(Snum)
    End If
End Function

Public Function StarProbe_Left_Ink_Dot(Number As Integer)
    Dim SLeft As String
    Dim SO, SI As String
  
    If DemoMode = 1 Then Exit Function
  
    If IO_2001X = 0 Then
        If Number <> 0 Then
            SLeft = Number
            SLeft = "IK" & SLeft
            ivprintf GpibAdd, SLeft + Chr$(10)
        End If
    Else
        If Number <> 0 Then
            SLeft = Number
            SLeft = "IK" & SLeft
            SO = ""
            SI = ""
            SI = MT2000.MSComm1.Input
            SI = MT2000.MSComm1.Input
            SO = Trim(SLeft)
            If ACCO = 1 Or Model_Select = 1 Then                            'ACCO only
                MT2000.MSComm1.Output = SO & vbCrLf
            Else
                MT2000.MSComm1.Output = SO & vbLf
            End If
    
            Set Test_time_delay = New ccrpStopWatch
            Test_time_delay.Reset
    
            Do
                SI = MT2000.MSComm1.Input
                If Left(SI, 1) = ">" Then Exit Do
                If Test_time_delay.Elapsed > 100 Then Exit Do
            Loop
            Set Test_time_delay = Nothing
        End If
    End If
End Function

Public Function StarProbe_Right_Ink_Dot(Number As Integer)
    Dim SRight As String
    Dim SO, SI As String

    If DemoMode = 1 Then Exit Function

    If IO_2001X = 0 Then
        If Number <> 0 Then
            SRight = Number
            SRight = "IK" & SRight
            ivprintf GpibAdd, SRight + Chr$(10)
        End If
    Else
        If Number <> 0 Then
            SRight = Number
            SRight = "IK" & SRight
            SO = ""
            SI = ""
            SI = MT2000.MSComm1.Input
            SI = MT2000.MSComm1.Input
            SO = Trim(SRight)
            If ACCO = 1 Or Model_Select = 1 Then                            'ACCO only
                MT2000.MSComm1.Output = SO & vbCrLf
            Else
                MT2000.MSComm1.Output = SO & vbLf
            End If
    
            Set Test_time_delay = New ccrpStopWatch
            Test_time_delay.Reset
    
            Do
                SI = MT2000.MSComm1.Input
                If Left(SI, 1) = ">" Then Exit Do
                If Test_time_delay.Elapsed > 100 Then Exit Do
            Loop
            Set Test_time_delay = Nothing
        End If
    End If
End Function

Public Function StarProbe_Ink_Enable()
    Dim s As String
    
    If DemoMode = 1 Then Exit Function
    
    If IO_2001X = 0 Then
        ivprintf GpibAdd, "EI" + Chr$(10)
    Else
        s = "EI"
        Call StarProbe_RS232_Output(s)
    End If
End Function

Public Function StarProbe_Ink_Reset_Timer_Limit(Number As Double)
    Dim Snum As String

    Snum = Number
    Snum = "SP26R" & Trim(Str(Snum))
    If DemoMode = 1 Then Exit Function

    If IO_2001X = 0 Then
        ivprintf GpibAdd, Snum + Chr$(10)
    Else
        Call StarProbe_RS232_Output(Snum)
    End If
End Function

Public Function StarProbe_Ink_Pulse_Width(Number As Integer)
    Dim Snum As String

    If DemoMode = 1 Then Exit Function

    Snum = Number
    Snum = "SP20P" & Trim(Str(Snum))
    If IO_2001X = 0 Then
        ivprintf GpibAdd, Snum + Chr$(10)
    Else
        Call StarProbe_RS232_Output(Snum)
    End If
End Function

Public Function StarProbe_Ink_Count_Limit(Number As Long)
    Dim Snum As String

    If DemoMode = 1 Then Exit Function

    Snum = Number
    Snum = "SP21" & Trim(Str(Snum))
    If IO_2001X = 0 Then
        ivprintf GpibAdd, Snum + Chr$(10)
    Else
        Call StarProbe_RS232_Output(Snum)
    End If
End Function

Public Function Starprobe_Requst_State_tip_clean_error() As String   'kks laser
    Dim Str, s As String

    If DemoMode = 1 Then Exit Function
    Str = " "
    If IO_2001X = 0 Then
        ivprintf GpibAdd, "?E" + Chr$(10)
        iread GpibAdd, Stress, 100, 0&, GpibTmp
        Stress = Replace(Stress, vbCrLf, "")
    
        MT2000.Text3.Refresh
        MT2000.Text3.Text = Stress
        
        Starprobe_Requst_State_tip_clean_error = Str
    Else
        s = "?E"
        Call StarProbe_RS232_Output(s)
        Stress = StarProbe_RS232_Input
        Stress = Replace(Stress, vbCrLf, "")
    
        MT2000.Text3.Refresh
        MT2000.Text3.Text = Stress
        
        Starprobe_Requst_State_tip_clean_error = Str
    End If
End Function

Public Function Starprobe_Requst_State_Wafer_On_Chuck() As String   'kks laser
    Dim Str, s As String

    If DemoMode = 1 Then Exit Function
    Str = " "
    If IO_2001X = 0 Then
        ivprintf GpibAdd, "?R" + Chr$(10)
        iread GpibAdd, Stress, 100, 0&, GpibTmp
        Stress = Replace(Stress, vbCrLf, "")
    
        If InStr(Stress, "W") = 0 Then
            Str = " "
        Else
            Str = Trim(Mid(Stress, InStr(Stress, "W") + 1, 1))
        End If
        Starprobe_Requst_State_Wafer_On_Chuck = Str
    Else
        s = "?R"
        Call StarProbe_RS232_Output(s)
        Sleep 4000
        Stress = StarProbe_RS232_Input

        If InStr(Stress, "W") = 0 Then
            Str = " "
        Else
            Str = Trim(Mid(Stress, InStr(Stress, "W") + 1, 1))
        End If
        Starprobe_Requst_State_Wafer_On_Chuck = Str
    End If
End Function

Public Function Starprobe_Requst_State_Wafer_Profile() As String   'kks laser
    Dim Str, s As String

    If DemoMode = 1 Then Exit Function
    Str = " "
    If IO_2001X = 0 Then
        ivprintf GpibAdd, "?R" + Chr$(10)
        iread GpibAdd, Stress, 100, 0&, GpibTmp
        Stress = Replace(Stress, vbCrLf, "")
    
        If InStr(Stress, "P") = 0 Then
            Str = " "
        Else
            Str = Trim(Mid(Stress, InStr(Stress, "P") + 1, 1))
        End If
        Starprobe_Requst_State_Wafer_Profile = Str
    Else
        s = "?R"
        Call StarProbe_RS232_Output(s)
        Stress = StarProbe_RS232_Input
    
        If InStr(Stress, "P") = 0 Then
            Str = " "
        Else
            Str = Trim(Mid(Stress, InStr(Stress, "P") + 1, 1))
        End If
        Starprobe_Requst_State_Wafer_Profile = Str
    End If
End Function

Public Function Starprobe_Requst_State_Wafer_Auto_Aligned() As String   'kks laser
    Dim Str, s As String

    If DemoMode = 1 Then Exit Function
    Str = " "
    If IO_2001X = 0 Then
        ivprintf GpibAdd, "?R" + Chr$(10)
        Sleep 20
        iread GpibAdd, Stress, 100, 0&, GpibTmp
        Stress = Replace(Stress, vbCrLf, "")
    
        If InStr(Stress, "A") = 0 Then
            Str = " "
        Else
            Str = Trim(Mid(Stress, InStr(Stress, "A") + 1, 1))
        End If
        Starprobe_Requst_State_Wafer_Auto_Aligned = Str
    Else
        s = "?R"
        Call StarProbe_RS232_Output(s)
        Stress = StarProbe_RS232_Input

        If InStr(Stress, "A") = 0 Then
            Str = " "
        Else
            Str = Trim(Mid(Stress, InStr(Stress, "A") + 1, 1))
        End If
        Starprobe_Requst_State_Wafer_Auto_Aligned = Str
    End If
End Function

Public Function StarProbe_Motor_X_Value() As String
    Dim SMot, s As String

    If DemoMode = 1 Then Exit Function
    Stress = " "
    If IO_2001X = 0 Then
        ivprintf GpibAdd, "?H" + Chr$(10)
        iread GpibAdd, Stress, 100, 0&, GpibTmp
        Stress = Replace(Stress, vbCrLf, "")
    
        If InStr(Stress, "X") = 0 Then
            SMot = " "
        Else
            SMot = Trim(Mid(Stress, InStr(Stress, "X") + 1, InStr(Stress, "Y") - 3))
        End If
        StarProbe_Motor_X_Value = SMot
    Else
        s = "?H"
        Call StarProbe_RS232_Output(s)
        Stress = StarProbe_RS232_Input

        If InStr(Stress, "X") = 0 Then
            SMot = " "
        Else
            SMot = Trim(Mid(Stress, InStr(Stress, "X") + 1, InStr(Stress, "Y") - 3))
        End If
        StarProbe_Motor_X_Value = SMot
    End If
End Function

Public Function Starprobe_motor_Y_Value() As String
    Dim SMot, s As String

    If DemoMode = 1 Then Exit Function

    If IO_2001X = 0 Then
        Stress = " "
    
        ivprintf GpibAdd, "?H" + Chr$(10)
        iread GpibAdd, Stress, 100, 0&, GpibTmp
        Stress = Replace(Stress, vbCrLf, "")
    
        If InStr(Stress, "Y") = 0 Then
            SMot = ""
        Else
            SMot = Trim(Mid(Stress, InStr(Stress, "Y") + 1))
        End If
        Starprobe_motor_Y_Value = SMot
    Else
        Stress = " "
        s = "?H"
        Call StarProbe_RS232_Output(s)
        Stress = StarProbe_RS232_Input

        If InStr(Stress, "Y") = 0 Then
            SMot = ""
        Else
            SMot = Trim(Mid(Stress, InStr(Stress, "Y") + 1))
        End If
        Starprobe_motor_Y_Value = SMot
    End If
End Function

Public Function StarProbe_Zero_point()
    Dim x, y, XYVal, s As String
    Dim bpoint As Boolean
    Dim ErrorCount As Integer
    
    If DemoMode = 1 Then Exit Function
    
    If IO_2001X = 0 Then
        bpoint = True
        ErrorCount = 0
        
        Do While bpoint
            DoEvents   '추가
            
            Stress = " "
            ivprintf GpibAdd, "?H" + Chr$(10)
            iread GpibAdd, Stress, 100, 0&, GpibTmp
              
            If InStr(Stress, "X") = 0 Then
         '      x = 0
         '      y = 0
                ErrorCount = ErrorCount + 1
            Else
                x = Trim(Mid(Stress, InStr(Stress, "X") + 1, InStr(Stress, "Y") - 3))
                y = Trim(Mid(Stress, InStr(Stress, "Y") + 1))
               
                If Left(x, 1) = "0" And Left(y, 1) = "0" Then
                    bpoint = False
                End If
    
                If Left(x, 1) = "-" Then
                    x = x
                Else
                    x = "-" & x
                End If
                If Left(y, 1) = "-" Then
                    y = y
                Else
                    y = "-" & y
                End If
                
                y = Trim(Replace(y, vbCrLf, ""))
                Call StarProbe_Pluse_Move((x), (y))
                Sleep 30
                
                If StarProbe_Motor_End_check Then
                    MsgBox "Motor not end check !", 16, "STAR PROBE"
                    ErrorStop = True  '추가
                Else
                    bpoint = False
                End If
            End If
                 
            If ErrorCount > 5 Then
                MsgBox "GPIB Communication is Error.", 16, "GPIB ERROR"
                x = 0
                y = 0
                bpoint = False
                ErrorStop = True  '추가
            End If
            If bpoint = False Or ErrorStop = True Or bStop = True Then Exit Do      '추가
        Loop
        
        '******************
        
        Call Wafer_End(True)
    
        Sleep 2
        
        Call Wafer_End(False)
        Call Tester_Clear
        
        '******************
        StarProbe.FIleName_MeasureResult = "c:\Star Probe\Image\Temp.BMP"
        Form_StarProbe_MeasureDataSave.Show vbModal
        bStarprobe_AfterInk = False
    Else
        bpoint = True
        ErrorCount = 0
    
        Do While bpoint
            DoEvents   '추가
            Stress = " "
            s = "?H"
            Call StarProbe_RS232_Output(s)
            Stress = StarProbe_RS232_Input
            If InStr(Stress, "X") = 0 Then
                ErrorCount = ErrorCount + 1
            Else
                x = Trim(Mid(Stress, InStr(Stress, "X") + 1, InStr(Stress, "Y") - 3))
                y = Trim(Mid(Stress, InStr(Stress, "Y") + 1))
           
                If Left(x, 1) = "0" And Left(y, 1) = "0" Then
                End If
                bpoint = False

                If Left(x, 1) = "-" Then
                    x = x
                Else
                    x = "-" & x
                End If
                If Left(y, 1) = "-" Then
                    y = y
                Else
                    y = "-" & y
                End If
                y = Trim(Replace(y, vbCrLf, ""))
                Call StarProbe_Pluse_Move((x), (y))
                Sleep 30
            End If
             
            If ErrorCount > 5 Then
                MsgBox "GPIB Communication is Error.", 16, "GPIB ERROR"
                x = 0
                y = 0
                bpoint = False
                ErrorStop = True  '추가
            End If
            If bpoint = False Or ErrorStop = True Or bStop = True Then Exit Do      '추가
        Loop
        StarProbe.FIleName_MeasureResult = "c:\Star Probe\Image\Temp.BMP"
        Form_StarProbe_MeasureDataSave.Show vbModal
        bStarprobe_AfterInk = False
    End If
End Function

Public Function StarProbe_Pluse_Move(x As String, y As String)
    Dim XYVal As String
 
    If DemoMode = 1 Then Exit Function
    XYVal = "MMX" & x & "Y" & y
    If IO_2001X = 0 Then
        ivprintf GpibAdd, XYVal + Chr$(10)
    Else
        Call StarProbe_RS232_Output(XYVal)
    End If
End Function

Public Function StarProbe_Top_Edge()
    Dim Edge As Integer
    Dim i As Double
    Dim bpoint As Boolean
    Dim ErrorCount As Integer

    If DemoMode = 1 Then Exit Function
    bpoint = True
    ErrorCount = 0
    If IO_2001X = 0 Then
        Do While bpoint
            DoEvents                      '추가
        
            Edge = Starprobe_Edge_Check
    
            If Edge = 1 Then
                bpoint = False
            ElseIf Edge = 3 Then
                ErrorCount = ErrorCount + 1
            Else
                i = 10
                Call StarProbe_Pluse_Move((0), (i))
                Sleep 10
           
                If StarProbe_Motor_End_check Then
                    MsgBox "Motor not end check !", 16, "STAR PROBE"
                    bpoint = False
                End If
            End If
        
            If ErrorCount > 5 Then
                MsgBox "GPIB Communication is Error.", 16, "GPIB ERROR"
                bpoint = False
            End If
            If bpoint = False Or bStop = True Then Exit Do  '추가
        Loop
    Else
        Do While bpoint
            DoEvents                      '추가
            Edge = Starprobe_Edge_Check
            If Edge = 1 Then
                bpoint = False
            ElseIf Edge = 3 Then
                ErrorCount = ErrorCount + 1
            Else
                i = 10
                Call StarProbe_Pluse_Move((0), (i))
                Sleep 10
            End If
    
            If ErrorCount > 5 Then
                MsgBox "GPIB Communication is Error.", 16, "GPIB ERROR"
                bpoint = False
            End If
            If bpoint = False Or bStop = True Then Exit Do  '추가
        Loop
    End If
End Function

Public Function StarProbe_Bottom_edge()
    Dim Edge As Integer
    Dim i As Double
    Dim bpoint As Boolean
    Dim ErrorCount As Integer

    If DemoMode = 1 Then Exit Function

    If IO_2001X = 0 Then
        bpoint = True
        ErrorCount = 0
    
        Do While bpoint
            DoEvents           '추가
            Edge = Starprobe_Edge_Check
            If Edge = 1 Then
                bpoint = False
            ElseIf Edge = 3 Then
                ErrorCount = ErrorCount + 1
            Else
                i = -10
                Call StarProbe_Pluse_Move((0), (i))
                Sleep 10
             
                If StarProbe_Motor_End_check Then
                    MsgBox "Motor not end check !", 16, "STAR PROBE"
                    bpoint = False
                End If
            End If
       
            If ErrorCount > 10 Then
                MsgBox "GPIB Communication is Error.", 16, "GPIB ERROR"
                bpoint = False
            End If
            If bpoint = False Or bStop = True Then Exit Do   '추가
        Loop
    Else
        bpoint = True
        ErrorCount = 0
  
        Do While bpoint
            DoEvents           '추가
            Edge = Starprobe_Edge_Check
            If Edge = 1 Then
                bpoint = False
            ElseIf Edge = 3 Then
                ErrorCount = ErrorCount + 1
            Else
                i = -10
                Call StarProbe_Pluse_Move((0), (i))
                Sleep 10
            End If
     
            If ErrorCount > 10 Then
                MsgBox "GPIB Communication is Error.", 16, "GPIB ERROR"
                bpoint = False
            End If
            If bpoint = False Or bStop = True Then Exit Do   '추가
        Loop
    End If
End Function

Public Function StarProbe_Top_Die_Step_Edge(y As String)
    Dim Edge As Integer
    Dim i As Double
    Dim bpoint As Boolean
    Dim ErrorCount As Integer

    If DemoMode = 1 Then Exit Function
    If IO_2001X = 0 Then
        bpoint = True
        ErrorCount = 0
      
        i = y
        Do While bpoint
            DoEvents                      '추가
            Edge = Starprobe_Edge_Check
    
            If Edge = 1 Then
                bpoint = False
            ElseIf Edge = 3 Then
                ErrorCount = ErrorCount + 1
            Else
                i = i + 1
                Call StarProbe_XY_Moving((0), (i))
                Sleep 10
            End If
        
            If ErrorCount > 10 Then
                MsgBox "GPIB Communication is Error.", 16, "GPIB ERROR"
                bpoint = False
            End If
        
            If bpoint = False Or bStop = True Then Exit Do  '추가
        Loop
    Else
        bpoint = True
        ErrorCount = 0
  
        i = y
        Do While bpoint
            DoEvents                      '추가
            Edge = Starprobe_Edge_Check

            If Edge = 1 Then
                bpoint = False
            ElseIf Edge = 3 Then
                ErrorCount = ErrorCount + 1
            Else
                i = i + 1
                Sleep 10
            End If
    
            If ErrorCount > 10 Then
                MsgBox "GPIB Communication is Error.", 16, "GPIB ERROR"
                bpoint = False
            End If
    
            If bpoint = False Or bStop = True Then Exit Do  '추가
        Loop
    End If
End Function

Public Function Starprobe_Left_Edge()
    Dim Edge As Integer
    Dim i  As Double
    Dim bpoint As Boolean
    Dim ErrorCount As Integer

    If DemoMode = 1 Then Exit Function

    If IO_2001X = 0 Then
        bpoint = True
        ErrorCount = 0
    
        Do While bpoint
            DoEvents                     '추가
            Edge = Starprobe_Edge_Check
            If Edge = 1 Then
                bpoint = False
            ElseIf Edge = 3 Then
                ErrorCount = ErrorCount + 1
            Else
                i = 10
                Call StarProbe_Pluse_Move((i), (0))
                Sleep 5
           
                If StarProbe_Motor_End_check Then
                    MsgBox "Motor not end check !", 16, "STAR PROBE"
                    bpoint = False
                End If
            End If
        
            If ErrorCount > 10 Then
                MsgBox "GPIB Communication is Error.", 16, "GPIB ERROR"
                bpoint = False
            End If
        
            If bpoint = False Or bStop = True Then Exit Do      '추가
        Loop
    Else
        bpoint = True
        ErrorCount = 0
    
        Do While bpoint
            DoEvents                     '추가
            Edge = Starprobe_Edge_Check
            If Edge = 1 Then
                bpoint = False
            ElseIf Edge = 3 Then
                ErrorCount = ErrorCount + 1
            Else
                i = 10
                Call StarProbe_Pluse_Move((i), (0))
                Sleep 5
            End If
        
            If ErrorCount > 10 Then
                MsgBox "GPIB Communication is Error.", 16, "GPIB ERROR"
                bpoint = False
            End If
            If bpoint = False Or bStop = True Then Exit Do      '추가
        Loop
    End If
End Function

Public Function Starprobe_Left_Die_Step_Edge(x As String)
    Dim Edge As Integer
    Dim i  As Double
    Dim bpoint As Boolean
    Dim ErrorCount As Integer

    If DemoMode = 1 Then Exit Function
    bpoint = True
    ErrorCount = 0
    i = x
    Do While bpoint
        DoEvents                     '추가
        Edge = Starprobe_Edge_Check
        If Edge = 1 Then
            bpoint = False
        ElseIf Edge = 3 Then
            ErrorCount = ErrorCount + 1
        Else
            i = i + 1
            Call StarProbe_XY_Moving((i), (0))
            Sleep 10
        End If

        If ErrorCount > 10 Then
            MsgBox "GPIB Communication is Error.", 16, "GPIB ERROR"
            bpoint = False
        End If
        If bpoint = False Or bStop = True Then Exit Do '추가
    Loop
End Function

Public Function StarProbe_Right_Edge()
    Dim Edge As Integer
    Dim i As Double
    Dim bpoint As Boolean
    Dim ErrorCount As Integer

    If DemoMode = 1 Then Exit Function

    If IO_2001X = 0 Then
        bpoint = True
        ErrorCount = 0
    
        Do While bpoint
            DoEvents                       '추가
            Edge = Starprobe_Edge_Check
            If Edge = 1 Then
                bpoint = False
            ElseIf Edge = 3 Then
                ErrorCount = ErrorCount + 1
            Else
                i = -10
                Call StarProbe_Pluse_Move((i), (0))
                Sleep 5
           
                If StarProbe_Motor_End_check Then
                    MsgBox "Motor not end check !", 16, "STAR PROBE"
                    bpoint = False
                End If
            End If
        
            If ErrorCount > 10 Then
                MsgBox "GPIB Communication is Error.", 16, "GPIB ERROR"
                bpoint = False
            End If
            If bpoint = False Or bStop = True Then Exit Do
        Loop
    Else
        bpoint = True
        ErrorCount = 0

        Do While bpoint
            DoEvents                       '추가
            Edge = Starprobe_Edge_Check
            If Edge = 1 Then
                bpoint = False
            ElseIf Edge = 3 Then
                ErrorCount = ErrorCount + 1
            Else
                i = -10
                Call StarProbe_Pluse_Move((i), (0))
                Sleep 5
            End If
            If ErrorCount > 10 Then
                MsgBox "GPIB Communication is Error.", 16, "GPIB ERROR"
                bpoint = False
            End If
            If bpoint = False Or bStop = True Then Exit Do
        Loop
    End If
End Function

Public Function StarProbe_Right_Die_Step_Edge(x As String)
    Dim Edge As Integer
    Dim i As Double
    Dim bpoint As Boolean
    Dim ErrorCount As Integer

    If DemoMode = 1 Then Exit Function
    
    bpoint = True
    ErrorCount = 0

    i = x

    Do While bpoint
        DoEvents                    '추가
        Edge = Starprobe_Edge_Check
        If Edge = 1 Then
            bpoint = False
        ElseIf Edge = 3 Then
            ErrorCount = ErrorCount + 1
        Else
            i = i - 1
            Call StarProbe_XY_Moving((i), (0))
            Sleep 30
        End If
        
        If ErrorCount > 10 Then
            MsgBox "GPIB Communication is Error.", 16, "GPIB ERROR"
            bpoint = False
        End If
    
        If bpoint = False Or bStop = True Then Exit Do   '추가
    Loop
End Function

Public Function StarProbe_Motor_End_check() As Boolean
    Dim XYpos, x, y, Z As String
    Dim bpoint As Boolean
    Dim Error_count As Integer
    Dim SI, Str, s, P As String
    Dim Lencount As Double
    Dim Input_Timer_Delay As ccrpStopWatch
    Dim RS232Time_Delay As ccrpStopWatch

    bpoint = True
    Error_count = 0

    If DemoMode = 1 Then Exit Function
    
    If IO_2001X = 0 Then
        Do While bpoint
            Stress = "  "
            Call iread(GpibAdd, Stress, 10, 0&, GpibTmp)
            
            Stress = Replace(Stress, vbCrLf, "")
           
            If InStr(1, Stress, "MC") > 0 Then
                bpoint = False
            Else
                Error_count = Error_count + 1
                Sleep 40
            End If
          
            If Error_count > 2 Then
                bStop = True
                Exit Do
            End If
            If bpoint = False Or bStop = True Then Exit Do
        Loop
        StarProbe_Motor_End_check = bpoint
    Else
        bpoint = False
        Error_count = 0
 
        Set Input_Timer_Delay = New ccrpStopWatch
        Set RS232Time_Delay = New ccrpStopWatch

        RS232Time_Delay.Reset
 
        Do
            If RS232Time_Delay.Elapsed > 150 Then Exit Do
        Loop
  
        Lencount = 0
        SI = ""
        Input_Timer_Delay.Reset
        
        Do
            DoEvents
            Str = MT2000.MSComm1.Input
            SI = SI & Str
            If SI <> Empty Then
                If InStr(1, SI, "MC") > 0 Then
                    SI = Trim(SI)
                    Exit Do
                Else
                    SI = Trim(SI)
                    RS232Time_Delay.Reset
                    Do
                        DoEvents
                        If RS232Time_Delay.Elapsed > 50 Then Exit Do
                    Loop
                    s = MT2000.MSComm1.Input
                    SI = SI & s
                    SI = Trim(SI)
                    If ACCO = 1 Or Model_Select = 1 Then                            'ACCO only
                        If Right(SI, 1) = vbCrLf And InStr(1, SI, "MC") > 0 Then
                            Exit Do
                        End If
                    Else
                        If Right(SI, 1) = vbLf And InStr(1, SI, "MC") > 0 Then
                            Exit Do
                        End If
                    End If
                End If
            End If

            If Input_Timer_Delay.Elapsed > 3000 Then
                Exit Do
            End If
        Loop

        If ACCO = 1 Or Model_Select = 1 Then                            'ACCO only
            SI = Replace(SI, vbCrLf, "")
        Else
            SI = Replace(SI, vbLf, "")
        End If
        SI = Replace(SI, ">", "")
        SI = Trim(SI)

        SI = MT2000.MSComm1.Input
        SI = MT2000.MSComm1.Input
        SI = MT2000.MSComm1.Input

        Set Input_Timer_Delay = Nothing
        Set RS232Time_Delay = Nothing
    
        StarProbe_Motor_End_check = bpoint
    End If
End Function

Public Function StarProbe_Z_End_check() As Boolean
    Dim XYpos, x, y, Z As String
    Dim bpoint As Boolean

    If DemoMode = 1 Then Exit Function

    If IO_2001X = 0 Then
        bpoint = True
    
        Stress = ""
        Do While bpoint
            
            Call iread(GpibAdd, Stress, 10, 0&, GpibTmp)
            
            Stress = Replace(Stress, vbCrLf, "")
      
            If InStr(1, Stress, "MC") > 0 Then
                bpoint = False
                Sleep 100
            End If
         
            If bpoint = False Or bStop = True Then Exit Do
        Loop
        StarProbe_Z_End_check = bpoint
    Else
        bpoint = True
        Stress = ""
        Do While bpoint
            Stress = StarProbe_RS232_Input
            If InStr(1, Stress, "MC") > 0 Then
                bpoint = False
                Sleep 100
            End If
            If bpoint = False Or bStop = True Then Exit Do
        Loop
        StarProbe_Z_End_check = bpoint
    End If
End Function

Public Function StarProbe_clean_tip_End_check() As Boolean
    Dim XYpos, x, y, Z As String
    Dim bpoint As Boolean

    If DemoMode = 1 Then Exit Function

    bpoint = True
    Stress = ""
    If IO_2001X = 0 Then
        Do While bpoint
            Call iread(GpibAdd, Stress, 10, 0&, GpibTmp)
            
            Stress = Replace(Stress, vbCrLf, "")
            
            MT2000.Text2.Refresh
            MT2000.Text2.Text = Stress
      
            If InStr(1, Stress, "MC") > 0 Then
                bpoint = False
                Sleep 100
            End If
            If bpoint = False Or bStop = True Then Exit Do
        Loop
        StarProbe_clean_tip_End_check = bpoint
    Else
        Do While bpoint
            Stress = StarProbe_RS232_Input
            
            MT2000.Text2.Refresh
            MT2000.Text2.Text = Stress
            
            If InStr(1, Stress, "MC") > 0 Then
                bpoint = False
                Sleep 100
            End If
            If bpoint = False Or bStop = True Then Exit Do
        Loop
        StarProbe_clean_tip_End_check = bpoint
    End If
End Function

Public Function StarProbe_First_Die_Set()
    Dim x, y As String
    Dim waferhalf As Double
    Dim DemY As Double
    Dim Xcenter, Ycenter As String
    Dim DemX As Double
    Dim FirstX, FirstY As String
  
    If DemoMode = 1 Then Exit Function
    waferhalf = val(StarProbe.WaferSizemm) / 2
    x = 0
    y = Trim(((waferhalf / StarProbe.ChipSizeY) + 1) * (-1))

    Call StarProbe_Motor_Home
    Call StarProbe_Tip_center
  
    x = waferhalf
     
    DemX = Round(((x + 1) / 0.0025) * (-1))
    
    Call StarProbe_Pluse_Move((DemX), (0))
   
    If StarProbe_Motor_End_check Then
        MsgBox "Motor not end check !", 16, "STAR PROBE"
        Exit Function
    End If
  
    Sleep 100
 
    Call Starprobe_Left_Edge
      
    Left_POS = StarProbe_XY_Position
  
    Left_MM = StarProbe_Motor_X_Value
        
    x = waferhalf
    DemX = Round(((x * 2) / 0.0025))
    DemX = DemX + (StarProbe.ChipSizeX * 5)
      
    Call StarProbe_Pluse_Move((DemX), (0))
   
    If StarProbe_Motor_End_check Then
        MsgBox "Motor not end check !", 16, "STAR PROBE"
        Exit Function
    End If

    Call StarProbe_Right_Edge

    Right_POS = StarProbe_XY_Position
    Right_MM = StarProbe_Motor_X_Value
   
    Tatal_MM = Abs(Left_MM) - Abs(Right_MM)
    Tatal_MM = Abs(Tatal_MM)
    Tatal_MM = val(Tatal_MM) * 0.0025
    
    DemX = val(StarProbe.WaferSizemm) / 2
    DemX = ((DemX) / 0.0025) * (-1)
    DemY = Round((StarProbe.ChipSizeY / 0.0025) * 10)
    DemY = DemX - DemY
    Call StarProbe_Pluse_Move((DemX), (DemY))
   
    If StarProbe_Motor_End_check Then
        MsgBox "Motor not end check !", 16, "STAR PROBE"
        Exit Function
    End If

    Call StarProbe_Top_Edge
   
    DemX = DemX * (-1)

    Call StarProbe_Pluse_Move((0), (DemX))
   
    If StarProbe_Motor_End_check Then
        MsgBox "Motor not end check !", 16, "STAR PROBE"
        Exit Function
    End If

    FirstX = StarProbe_XY_Position()
    FirstY = Mid(FirstX, InStr(FirstX, "Y") + 1)
    FirstX = Mid(FirstX, 2, InStr(FirstX, "Y") - 2)
'
    Call StarProbe_First_Chip
 
    x = waferhalf
    DemX = Round(((x + 1) / 0.0025) * (-1))
    Call StarProbe_Pluse_Move((DemX), (0))
    If StarProbe_Motor_End_check Then
        MsgBox "Motor not end check !", 16, "STAR PROBE"
        Exit Function
    End If

    Call Starprobe_Left_Edge

    Left_POS = StarProbe_XY_Position
  
    Left_POS = Mid(Left_POS, 2, InStr(Left_POS, "Y") - 2)
    Left_MM = StarProbe_Motor_X_Value

    x = waferhalf
    DemX = Round(((x * 2) / 0.0025))
    DemX = DemX + (StarProbe.ChipSizeX * 10)

    Call StarProbe_Pluse_Move((DemX), (0))

    If StarProbe_Motor_End_check Then
        MsgBox "Motor not end check !", 16, "STAR PROBE"
        Exit Function
    End If

    Call StarProbe_Right_Edge

    Right_POS = StarProbe_XY_Position
    Right_POS = Mid(Right_POS, 2, InStr(Right_POS, "Y") - 2)
    Right_MM = StarProbe_Motor_X_Value
   
    FirstX = Abs(val(Left_POS)) + Abs(val(Right_POS)) + 1
   
    FirstX = FirstX / 2
   
    If (FirstX Mod 2) = 0 Then
        FirstX = FirstX - 1
    Else
        FirstX = FirstX
    End If
   
    FirstX = ChipCount((FirstX)) + val(Left_POS)
   
   Call StarProbe_FistDie
    
    If StarProbe_Motor_End_check Then
        MsgBox "Motor not end check !", 16, "STAR PROBE"
        Exit Function
    End If
   
    FirstX = (FirstX * StarProbe.ChipSizeX) / 0.0025
    Call StarProbe_Pluse_Move((FirstX), (0))
   
    If StarProbe_Motor_End_check Then
        MsgBox "Motor not end check !", 16, "STAR PROBE"
        Exit Function
    End If
    Call StarProbe_First_Chip
    Sleep 200
   
    x = waferhalf
    DemX = Round(((x) / 0.0025) * (-1))
    
    Call StarProbe_Pluse_Move((DemX), (0))

    
    If StarProbe_Motor_End_check Then
        MsgBox "Motor not end check !", 16, "STAR PROBE"
        Exit Function
    End If

    Call Starprobe_Left_Edge

    Left_POS = StarProbe_XY_Position
    Left_POS = Mid(Left_POS, 2, InStr(Left_POS, "Y") - 2)
    Left_MM = StarProbe_Motor_X_Value

'   Exit Function
    x = waferhalf
    DemX = Round(((x * 2) / 0.0025))
    DemX = DemX + (StarProbe.ChipSizeX)

    Call StarProbe_Pluse_Move((DemX), (0))
'   Sleep 100
   If StarProbe_Motor_End_check Then
     MsgBox "Motor not end check !", 16, "STAR PROBE"
     Exit Function
   End If

   Call StarProbe_Right_Edge

   Right_POS = StarProbe_XY_Position
   Right_POS = Mid(Right_POS, 2, InStr(Right_POS, "Y") - 2)
   Right_MM = StarProbe_Motor_X_Value
   
   Call StarProbe_FistDie
   
   If StarProbe_Motor_End_check Then
     MsgBox "Motor not end check !", 16, "STAR PROBE"
     Exit Function
   End If
   
   Tatal_MM = Abs(Left_MM) - Abs(Right_MM)
   Tatal_MM = Abs(Tatal_MM)
   Tatal_MM = val(Tatal_MM) * 0.0025

     
End Function

Function GetRound(ByVal vNumber As Double, ByVal vDigit As Integer) As Double
    
    Dim vRemain As Double
    
    vRemain = vNumber - Fix(vNumber)
    vRemain = vRemain * 10 ^ vDigit
    vRemain = Fix(vRemain) / 10 ^ vDigit
    GetRound = Fix(vNumber) + vRemain
    
End Function

Public Function StarProbe_First_Chip_Motor_Set()
Dim x, y, MXstr, MYstr As String
Dim bpoint As Boolean
Dim ErrorCount As Integer
Dim Mx, My, ValX, ValY As Double
Dim First_Chip_Error As Boolean

If DemoMode = 1 Then Exit Function

    First_Chip_Error = False
    bpoint = True
    ErrorCount = 0
    MXstr = " "
    MYstr = " "
    
    Mx = StarProbe.CurrentChip.Mx
    My = StarProbe.CurrentChip.My
    
    Do While bpoint
    
        DoEvents   '추가
        
        Stress = " "
        ivprintf GpibAdd, "?H" + Chr$(10)
        iread GpibAdd, Stress, 100, 0&, GpibTmp
          
        If InStr(Stress, "X") = 0 Then
            ErrorCount = ErrorCount + 1
        Else
        
           x = Trim(Mid(Stress, InStr(Stress, "X") + 1, InStr(Stress, "Y") - 3))
           y = Trim(Mid(Stress, InStr(Stress, "Y") + 1))
           
'           If Left(x, 1) = "0" And Left(y, 1) = "0" Then
'              bpoint = False
'           End If

           ValX = Mx - val(x)
           ValY = My - val(y)
          
           If (ValX = 0) And (ValY = 0) Then Exit Do
           
          
'           If Left(x, 1) = "-" Then
'               x = x
'           Else
'               x = "-" & x
'           End If
'           If Left(y, 1) = "-" Then
'               y = y
'           Else
'               y = "-" & y
'           End If

           MXstr = Trim(Replace(Str(ValX), vbCrLf, ""))
           MYstr = Trim(Replace(Str(ValY), vbCrLf, ""))
           
           Call StarProbe_Pluse_Move((MXstr), (MYstr))
            Sleep 30
            
            If StarProbe_Motor_End_check Then
              MsgBox "Motor not end check !", 16, "STAR PROBE"
              First_Chip_Error = True  '추가
            Else
            '  bpoint = False
            End If
            
        End If
             
        If ErrorCount > 5 Then
           MsgBox "GPIB Communication is Error.", 16, "GPIB ERROR"
           x = 0
           y = 0
           bpoint = False
           First_Chip_Error = True  '추가
        End If
        
        If bpoint = False Or First_Chip_Error = True Or bStop = True Then Exit Do      '추가
        
    Loop
    
    If First_Chip_Error = True Then
        MsgBox "First Chip Not Setting.", 16, "GPIB ERROR"
        bStop = True
  '  Else
        
    End If
    
  
End Function

Function STR_FIX(ST As String, N As Integer) As String

    Dim l As Integer
    Dim D As Integer
    
    ST = Trim(ST)
    l = Len(ST)
    
    If N < 0 Then
         D = -1
         N = N * -1
    Else
         D = 1
    End If
    If (l >= N) Then
         STR_FIX = Mid(ST, 1, N)
    Else
         If D = 1 Then
              STR_FIX = ST + Space(N - l)
         Else
              STR_FIX = Space(N - l) + ST
         End If
    End If
    
End Function
