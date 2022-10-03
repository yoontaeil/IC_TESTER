Attribute VB_Name = "CODE"
Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long


Public Declare Function QueryPerformanceCounter Lib "Kernel32" _
                                 (x As Currency) As Boolean
Public Declare Function QueryPerformanceFrequency Lib "Kernel32" _
                                 (x As Currency) As Boolean
Public Declare Function QueryPerformanceFrequencyAny Lib "Kernel32" Alias _
    "QueryPerformanceFrequency" (lpFrequency As Any) As Long

Public Declare Function QueryPerformanceCounterAny Lib "Kernel32" Alias _
    "QueryPerformanceCounter" (lpPerformanceCount As Any) As Long

Rem 레지스트리

Global Const KEY_SUB = "Software\MT2000"
Global Const KEY_ALL_ACCESS = &H2003F
Global Const HKEY_LOCAL_MACHINE = &H80000002

Type test_info
    PGM_CHECK As Boolean
    Test_PGM As String
    Bin_Count As Integer
End Type
'TIME
Public STT_time As String                           'Wafer의 Test를 시작한 시간을 기억
Public END_time As String                           'Wafer의 Test가 끝난 시간을 기억

'[ 2021.06.25 ]
Global First_X As Integer
Global First_Y As Integer
Global First_Zoom_TOP As Integer
Global First_Zoom_LEFT As Integer
Global First_Original_TOP As Integer
Global First_Original_LEFT As Integer

Global Test_Bin As Integer
Global Test_Cnt As Long
Global Good_Cnt As Long
Global Test_Ready As Boolean
Global Test_Fail_Count1 As Long
Global Test_Fail_Count2 As Long
Global Test_Fail_Count3 As Long
Global Test_Fail_Count4 As Long

'Global No_Probe(50) As Boolean
Global No_Probe As Boolean
Global Sample_No_Ink As Boolean
Global Sample_Yes_Ink As Boolean

Global Slot_No(1 To 25) As Boolean
Global Slot_Max_Count As Integer
Global ETS_Count As Integer
Global File_Count As Integer
Global Tower_Lamp_Add As Integer
Global R_File_Name As String
Global Bin_Count(30) As Long
Global New_Lot, Old_Lot As String
Global Wafer_Start As Boolean

Global PROD As test_info
Global Const DemoMode = 0            ' O:System  1: ONLY PC
Global Const Tester_Select = 0      '0:EAGLE, 1:AMT88
Global Const ACCO = 0                ' O:EAGLE   1: ACCO
Global Const AOI_MODE = 1           '0:normal,1:AOI
Global Const IO_2001X = 0           '0:GPIB, 1:RS232
Global Const SaveDrive = 1              '0:C:\, 1:D:\
Global Const LOG_FILE_ON = 1        '0:log file off, 1:log file on                  '[ 2022.07.20 ]
Global WINDIR As String               ' Window 98,xp
Global GOOD_COUNT_BACKUP As Long        '[ 2020.04.06 ] map load시 초기 good카운트 저장
Global BMP_file As String               '[ 2020.03.11 ] BMP저장경로
Global BIN_Command(30) As String

Global Text_Bin_Count_No(30) As Long
Global SAVECNT As Long                       '자동저장 카운트
Global LOOP_COUNT As Long                    'fail count up/down
Global FILE_NAMEING As String
Global O_X As Integer
Global O_Y As Integer

Global Fail_Find As Boolean             '4개중 불량이 하나라도 발생하면 true
Global INK_OFF_TEST As Boolean      '2015.11.30 : ink off시 측정후 ink를 할것인지 물어보는 flag


Global AutoAlign_Flag As Boolean    '2016.01.18

Global TESTING_flag As Boolean      '2016.03.11

Global YOON_CNT As Integer          '2016.06.15 : 불량시 위아래 측정하기위한 카운트

Global DEMO_FAIL As Boolean

Global SAMPLING As Boolean          '2016.06.22 : sampling mode set

Global FAIL_COUNT As Integer        '2016.06.23 : NG loop시 해당 y축에서 발생한 불량수

Global Ink_Start_Flag As Integer    '2016.09.27 : Ink Start Position Flag
Global SP_FLAG As Boolean           'sp load유무

Global SP_CNT As Integer            '[ 2017.03.23 ] : SP로드시 해당파일의 숫자 저장.

Global bad_click As Integer         '0:pass,1:fail

Global Load_MAP As String           '2015.11.30 : load map name

Global NOW_NO(25) As Boolean        '2017.08.01 : 측정해야할 wafer의 배열
Global TT_NO As Integer             '2017.08.01 : 현재 측정중인 wafer 넘버
Global W_NO As String              '2017.08.01 : 현재 측정중인 wafer 넘버

Global EQU As String
Global OPE As String
Global DEV As String
Global LOT As String
Global WAF As String
Global PRO As String

Global Server_path As String                '2019.12.17 : server 저장 경로
Global MAP_path As String                   '2020.09.07 : map 저장 경로
Global AOI_path As String                               'aoi path       [ 2020.10.29 ]
Global Barcode_Use As Boolean                           '바코드사용유무
Global Barcode_Name As String                           '바코드로 찍은 이름 저장.
Global Path_Check As Integer
'Global BMP_file As String               '[ 2020.03.11 ] BMP저장경로
'저장경로 선택 1:data, 2:PGM, 3:MAP

Global Array_str() As String
Global Array_tmp1 As String
Global Array_tmp2 As String
Global Array_tmp3 As String

'[ 2020.10.29 ] : AOI 관련 수정
Global AOI_MAP(51) As String        'AOI 검사후의 파일 이름을 저장하는 배열
Global AOI_Use As Boolean           'true : AOI use, false : AOI not use
Global AOI_FAIL_COUNT As Double    'map load시 aoi불량 숫자를 저장하는 변수
Global Const AOI_BIN = 19           'AOI 불량 BIN지정


Global CH_SET As Integer            '[ 2021.12.13 ] : channel select

'MODE select
Public Mode_Set As Boolean              '[ 2021.12.31 ] : false:Operator, true:Engineer

Public Center_fail As Boolean           '[ 2022.05.06 ] : 4채널시 중간에서 불량이 발생한 경우

'[ 2022.05.19 ]
Global tip_clean_count_flag_1 As Integer
Global tip_clean_count_flag_2 As Integer
Global tip_clean_count_flag_3 As Integer
Global tip_clean_count_flag_4 As Integer

Global Needle_Chk(25) As Boolean        '[ 2022.07.29 ] : 침적확인 유무 설정 변수
Global Needle_Chk_Ok As Boolean         '[ 2022.07.29 ] : 현재 웨이퍼의 니들체크 메시지 확인.
Global CHK_CANCEL As Boolean            '[ 2022.08.08 ] : check list를 확인했는지 유무

Global TESTER_OFF As Boolean            '[ 2022.08.10 ] : true : tester off, false tester on

Global Move_select As Boolean

'Global Const LOG_FILE_ON = 1        '0:log file off, 1:log file on                  '[ 2022.07.20 ]
Global MSG_DATA As String               '[ 2022.08.17 ] : x,y,z이동값 저장

Global Needle_check_flag As Boolean         '[ 2022.08.31 ] : 침적 벗어남 확인 flag
Global Needle_STT As Boolean                '[ 2022.08.31 ]

'[ 2022.09.30 ]
Global Model_Select As Integer         '1:ACCO, 2:AMT-88, 3:EAGLE(2001X), 4:EAGLE(E4090)
Global GOOD_BIN_NO As Integer

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 프로그램 수정내용
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'[16.12.01] : - Wafer(xx, yy).BIN = 0             '16.12.02
'             - StarProbe_Measure_1ch에서도 Sampling기능을 사용하도록 수정.
'[16.12.13] : - 테스트 중인 경우 Stop후 Auto시 first Die로 이동 하지 않도록 수정.
'             - INK시 ink변수의 결과에 상관 없이 다시 ink할 수 있도록 수정.

'[ 2017.03.23 ] : 주석처리
'[ 2017.03.23 ] : SP파일 로딩이 아닌 경우에만 적용하도록 수정.
'[ 2017.03.23 ] : SP파일인 경우의 처리를 추가
'[ 2017.03.23 ] : SP파일을 로딩한 경우 파일의 넘버를 추출하는 코드 추가
'[ 2017.03.23 ] : ink시 두가지 옵션으로 설정하도록 수정.
'[ 2017.03.23 ] : SP파일 로드시 기존의 카운트를 화면에 표시
'[ 2017.03.23 ] : SP로드시 해당파일의 숫자 저장.
'[ 2017.03.23 ] : DEMO Mode인 경우 INK시 black으로 표시해주는 부분
'[ 2017.09.18 ] : fail count up down 추가, MT2000의 text11추가
'[ 2018.01.29 ] : 4channel test시 y축으로도 jump구간 설정하는 기능 추가.
'[ 2021.10.27 ] : bin1(ok bin)을 지운 경우 처리.
'                 bin clear후 오동작하는 부분 관련 수정.
'[ 2022.05.06 ] : fail step 0입력가능하도록 수정.
'                 chip이 아닌 경우 처리 추가
'[ 2022.05.19 ] : tip clean
'[ 2022.07.20 ] : log file 저장되는 부분 수정. (c:\star probe\starprobe_log.dat)
'[ 2022.07.29 ] : 침적확인 관련 수정. (c:\star probe\needle_chk.dat)
'[ 2022.08.30 ] : 침적확인 방법
'[ 2022.08.31 ] : esc 누를경우 BIN17로 변경
'[ 2022.09.08 ] : map을 lot no를 기준으로 구분한다.
'[ 2022.09.30 ] : 체크를 하나도 안한 경우 경고 메시지 출력
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

