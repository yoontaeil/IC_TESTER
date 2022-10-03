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

Rem ������Ʈ��

Global Const KEY_SUB = "Software\MT2000"
Global Const KEY_ALL_ACCESS = &H2003F
Global Const HKEY_LOCAL_MACHINE = &H80000002

Type test_info
    PGM_CHECK As Boolean
    Test_PGM As String
    Bin_Count As Integer
End Type
'TIME
Public STT_time As String                           'Wafer�� Test�� ������ �ð��� ���
Public END_time As String                           'Wafer�� Test�� ���� �ð��� ���

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
Global GOOD_COUNT_BACKUP As Long        '[ 2020.04.06 ] map load�� �ʱ� goodī��Ʈ ����
Global BMP_file As String               '[ 2020.03.11 ] BMP������
Global BIN_Command(30) As String

Global Text_Bin_Count_No(30) As Long
Global SAVECNT As Long                       '�ڵ����� ī��Ʈ
Global LOOP_COUNT As Long                    'fail count up/down
Global FILE_NAMEING As String
Global O_X As Integer
Global O_Y As Integer

Global Fail_Find As Boolean             '4���� �ҷ��� �ϳ��� �߻��ϸ� true
Global INK_OFF_TEST As Boolean      '2015.11.30 : ink off�� ������ ink�� �Ұ����� ����� flag


Global AutoAlign_Flag As Boolean    '2016.01.18

Global TESTING_flag As Boolean      '2016.03.11

Global YOON_CNT As Integer          '2016.06.15 : �ҷ��� ���Ʒ� �����ϱ����� ī��Ʈ

Global DEMO_FAIL As Boolean

Global SAMPLING As Boolean          '2016.06.22 : sampling mode set

Global FAIL_COUNT As Integer        '2016.06.23 : NG loop�� �ش� y�࿡�� �߻��� �ҷ���

Global Ink_Start_Flag As Integer    '2016.09.27 : Ink Start Position Flag
Global SP_FLAG As Boolean           'sp load����

Global SP_CNT As Integer            '[ 2017.03.23 ] : SP�ε�� �ش������� ���� ����.

Global bad_click As Integer         '0:pass,1:fail

Global Load_MAP As String           '2015.11.30 : load map name

Global NOW_NO(25) As Boolean        '2017.08.01 : �����ؾ��� wafer�� �迭
Global TT_NO As Integer             '2017.08.01 : ���� �������� wafer �ѹ�
Global W_NO As String              '2017.08.01 : ���� �������� wafer �ѹ�

Global EQU As String
Global OPE As String
Global DEV As String
Global LOT As String
Global WAF As String
Global PRO As String

Global Server_path As String                '2019.12.17 : server ���� ���
Global MAP_path As String                   '2020.09.07 : map ���� ���
Global AOI_path As String                               'aoi path       [ 2020.10.29 ]
Global Barcode_Use As Boolean                           '���ڵ�������
Global Barcode_Name As String                           '���ڵ�� ���� �̸� ����.
Global Path_Check As Integer
'Global BMP_file As String               '[ 2020.03.11 ] BMP������
'������ ���� 1:data, 2:PGM, 3:MAP

Global Array_str() As String
Global Array_tmp1 As String
Global Array_tmp2 As String
Global Array_tmp3 As String

'[ 2020.10.29 ] : AOI ���� ����
Global AOI_MAP(51) As String        'AOI �˻����� ���� �̸��� �����ϴ� �迭
Global AOI_Use As Boolean           'true : AOI use, false : AOI not use
Global AOI_FAIL_COUNT As Double    'map load�� aoi�ҷ� ���ڸ� �����ϴ� ����
Global Const AOI_BIN = 19           'AOI �ҷ� BIN����


Global CH_SET As Integer            '[ 2021.12.13 ] : channel select

'MODE select
Public Mode_Set As Boolean              '[ 2021.12.31 ] : false:Operator, true:Engineer

Public Center_fail As Boolean           '[ 2022.05.06 ] : 4ä�ν� �߰����� �ҷ��� �߻��� ���

'[ 2022.05.19 ]
Global tip_clean_count_flag_1 As Integer
Global tip_clean_count_flag_2 As Integer
Global tip_clean_count_flag_3 As Integer
Global tip_clean_count_flag_4 As Integer

Global Needle_Chk(25) As Boolean        '[ 2022.07.29 ] : ħ��Ȯ�� ���� ���� ����
Global Needle_Chk_Ok As Boolean         '[ 2022.07.29 ] : ���� �������� �ϵ�üũ �޽��� Ȯ��.
Global CHK_CANCEL As Boolean            '[ 2022.08.08 ] : check list�� Ȯ���ߴ��� ����

Global TESTER_OFF As Boolean            '[ 2022.08.10 ] : true : tester off, false tester on

Global Move_select As Boolean

'Global Const LOG_FILE_ON = 1        '0:log file off, 1:log file on                  '[ 2022.07.20 ]
Global MSG_DATA As String               '[ 2022.08.17 ] : x,y,z�̵��� ����

Global Needle_check_flag As Boolean         '[ 2022.08.31 ] : ħ�� ��� Ȯ�� flag
Global Needle_STT As Boolean                '[ 2022.08.31 ]

'[ 2022.09.30 ]
Global Model_Select As Integer         '1:ACCO, 2:AMT-88, 3:EAGLE(2001X), 4:EAGLE(E4090)
Global GOOD_BIN_NO As Integer

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ���α׷� ��������
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'[16.12.01] : - Wafer(xx, yy).BIN = 0             '16.12.02
'             - StarProbe_Measure_1ch������ Sampling����� ����ϵ��� ����.
'[16.12.13] : - �׽�Ʈ ���� ��� Stop�� Auto�� first Die�� �̵� ���� �ʵ��� ����.
'             - INK�� ink������ ����� ��� ���� �ٽ� ink�� �� �ֵ��� ����.

'[ 2017.03.23 ] : �ּ�ó��
'[ 2017.03.23 ] : SP���� �ε��� �ƴ� ��쿡�� �����ϵ��� ����.
'[ 2017.03.23 ] : SP������ ����� ó���� �߰�
'[ 2017.03.23 ] : SP������ �ε��� ��� ������ �ѹ��� �����ϴ� �ڵ� �߰�
'[ 2017.03.23 ] : ink�� �ΰ��� �ɼ����� �����ϵ��� ����.
'[ 2017.03.23 ] : SP���� �ε�� ������ ī��Ʈ�� ȭ�鿡 ǥ��
'[ 2017.03.23 ] : SP�ε�� �ش������� ���� ����.
'[ 2017.03.23 ] : DEMO Mode�� ��� INK�� black���� ǥ�����ִ� �κ�
'[ 2017.09.18 ] : fail count up down �߰�, MT2000�� text11�߰�
'[ 2018.01.29 ] : 4channel test�� y�����ε� jump���� �����ϴ� ��� �߰�.
'[ 2021.10.27 ] : bin1(ok bin)�� ���� ��� ó��.
'                 bin clear�� �������ϴ� �κ� ���� ����.
'[ 2022.05.06 ] : fail step 0�Է°����ϵ��� ����.
'                 chip�� �ƴ� ��� ó�� �߰�
'[ 2022.05.19 ] : tip clean
'[ 2022.07.20 ] : log file ����Ǵ� �κ� ����. (c:\star probe\starprobe_log.dat)
'[ 2022.07.29 ] : ħ��Ȯ�� ���� ����. (c:\star probe\needle_chk.dat)
'[ 2022.08.30 ] : ħ��Ȯ�� ���
'[ 2022.08.31 ] : esc ������� BIN17�� ����
'[ 2022.09.08 ] : map�� lot no�� �������� �����Ѵ�.
'[ 2022.09.30 ] : üũ�� �ϳ��� ���� ��� ��� �޽��� ���
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

