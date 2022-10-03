Attribute VB_Name = "Module_StarProbe9052"
Option Explicit

Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function PciStart Lib "EL9050.DLL" () As Byte
Public Declare Function PciEnd Lib "EL9050.DLL" () As Byte
Public Declare Function PciCardExist Lib "EL9050.DLL" (ByVal card As Byte) As Byte
Public Declare Function PciIrqSet Lib "EL9050.DLL" () As Byte
Public Declare Function PciIrqClr Lib "EL9050.DLL" () As Byte
Public Declare Function PciGetInt Lib "EL9050.DLL" (ByVal card As Byte) As Long
Public Declare Function PciIntStatus Lib "EL9050.DLL" (ByVal card As Byte) As Byte
Public Declare Function PciGetBaseAddr Lib "EL9050.DLL" (ByVal card As Byte, ByVal cs As Byte) As Long
Public Declare Function PciGetBaseLength Lib "EL9050.DLL" (ByVal card As Byte, ByVal cs As Byte) As Long
Public Declare Function PciWrite Lib "EL9050.DLL" (ByVal card As Byte, ByVal address As Long, ByVal data As Long) As Byte
Public Declare Function PciWriteW Lib "EL9050.DLL" (ByVal card As Byte, ByVal address As Long, ByVal data As Integer) As Byte
Public Declare Function PciWriteB Lib "EL9050.DLL" (ByVal card As Byte, ByVal address As Long, ByVal data As Byte) As Byte
Public Declare Function PciRead Lib "EL9050.DLL" (ByVal card As Byte, ByVal address As Long) As Long
Public Declare Function PciReadW Lib "EL9050.DLL" (ByVal card As Byte, ByVal address As Long) As Integer
Public Declare Function PciReadB Lib "EL9050.DLL" (ByVal card As Byte, ByVal address As Long) As Byte
Public Declare Function PciReadReg Lib "EL9050.DLL" (ByVal card As Byte, ByVal offset As Long) As Long
Public Declare Function PciWriteReg Lib "EL9050.DLL" (ByVal card As Byte, ByVal offset As Long, ByVal data As Long) As Byte
Public Declare Function EepromReadDword Lib "EL9050.DLL" (ByVal card As Byte, ByVal reg As Long) As Long
Public Declare Function EepromWriteDword Lib "EL9050.DLL" (ByVal card As Byte, ByVal reg As Long, ByVal val As Long) As Byte
Public Declare Function EepromReadWord Lib "EL9050.DLL" (ByVal card As Byte, ByVal reg As Long) As Integer
Public Declare Function EepromWriteWord Lib "EL9050.DLL" (ByVal card As Byte, ByVal reg As Long, ByVal val As Long) As Byte


'#define  EL9050_IN_ADDR       0x8000  //  ~ 0x8003
'#define  EL9050_OUT_ADDR      0x8000  //  ~ 0x8003

Public Const PCI_IO_IN_ADDR  As Long = 32768  '&H8000
Public Const PCI_IO_OUT_ADDR  As Long = 32768 '&H8000

Public oRet As Boolean
Public Const CardNo As Byte = 0
Public OUT_W As Byte
Public bData As Byte
Public CardExist(0 To 5) As Boolean
Public DipSwAddr(0 To 5) As Long
Public Seg7Addr(0 To 5) As Long
Public Bin_Result(0 To 24) As Integer

' Pin Number 1 :  end
'            9    bin 1
'            10   bin 2
'            11   bin 3
'            12   bin 4
'            13   bin 5
'            14   bin 6
'            15   bin 7
'            16   bin 8
'            17   bin 9
'            18   bin 10
'            19   bin 11
'            20   bin 12
'            21   bin 13
'            22   bin 14
'            23   bin 15
'            33   start
'            34   wafer end
'            66   +24V
'            67   GND
'            68   GND

Public Sub Tester_Start(ByVal cmd As Boolean)
    If cmd Then
        OUT_W = OUT_W Or (2 ^ 0)
    Else
        OUT_W = OUT_W And (Not (2 ^ 0))
    End If
    PciWriteB CardNo, DipSwAddr(CardNo), OUT_W
End Sub

Public Sub Wafer_End(ByVal cmd As Boolean)
    If cmd Then
        OUT_W = OUT_W Or (2 ^ 1)
    Else
        OUT_W = OUT_W And (Not (2 ^ 1))
    End If
    PciWriteB CardNo, DipSwAddr(CardNo), OUT_W
End Sub

Public Sub Tester_Clear()
    PciWriteB CardNo, DipSwAddr(CardNo) + 3, &H80
    Sleep 3
    PciWriteB CardNo, DipSwAddr(CardNo) + 3, &H0
End Sub

Public Function Tester_End() As Boolean
    bData = PciReadB(CardNo, DipSwAddr(CardNo))
    oRet = IIf((bData And 2 ^ 0) <> 0, True, False)
    If oRet Then
        Tester_End = True
    Else
        Tester_End = False
    End If
End Function

Public Function Bin_NO() As Integer
    Dim i As Integer

    bData = PciReadB(CardNo, DipSwAddr(CardNo) + 1)
    
    For i = 0 To 7
        If ((bData And 2 ^ i) <> 0) Then
            Bin_Result(i + 1) = 1
        Else
            Bin_Result(i + 1) = 0
        End If
    Next
    
    bData = PciReadB(CardNo, DipSwAddr(CardNo) + 2)
    
    For i = 0 To 7
        If ((bData And 2 ^ i) <> 0) Then
            Bin_Result(i + 9) = 1
        Else
            Bin_Result(i + 9) = 0
        End If
    Next
    
    bData = PciReadB(CardNo, DipSwAddr(CardNo) + 3)
    
    For i = 0 To 7
        If ((bData And 2 ^ i) <> 0) Then
            Bin_Result(i + 17) = 1
        Else
            Bin_Result(i + 17) = 0
        End If
    Next

    For i = 1 To 24
        If Bin_Result(i) = 1 Then
            Bin_Count(i) = Bin_Count(i) + 1
            Bin_NO = i
            Exit For
        Else
            Bin_NO = 0
        End If
    Next
End Function
