Attribute VB_Name = "Module_PortIO"
' 원 소스에서 이부분 주석처리
'Declare Sub vbOut Lib "PORT_IO.DLL" (ByVal nPort As Integer, ByVal nData As Integer)
'Declare Function vbInp Lib "PORT_IO.DLL" (ByVal nPort As Integer) As Integer

Public Declare Function DlPortReadPortUchar Lib "ibm_io.dll" (ByVal Port As Long) As Byte
Public Declare Function DlPortReadPortUshort Lib "ibm_io.dll" (ByVal Port As Long) As Integer
Public Declare Function DlPortReadPortUlong Lib "ibm_io.dll" (ByVal Port As Long) As Long

Public Declare Sub DlPortWritePortUchar Lib "ibm_io.dll" (ByVal Port As Long, ByVal Value As Byte)
Public Declare Sub DlPortWritePortUshort Lib "ibm_io.dll" (ByVal Port As Long, ByVal Value As Integer)
Public Declare Sub DlPortWritePortUlong Lib "ibm_io.dll" (ByVal Port As Long, ByVal Value As Long)

Public Sub vbOut(ByVal nPort As Long, ByVal nData As Integer)
    
    ' demomode가 0 이면 처리하지 마라.
    If DemoMode = 1 Then Exit Sub
    If nData < 0 Then nData = nData + 256
    Call DlPortWritePortUchar(nPort, nData)
    
End Sub

Public Function vbInp(ByVal nPort As Long) As Byte

    If DemoMode = 1 Then Exit Function
    
    vbInp = DlPortReadPortUchar(nPort)

End Function


Public Function InPut8(ByVal Address As Long) As Byte

    InPut8 = DlPortReadPortUchar(Address)

End Function

Public Function InPut16u(ByVal Address As Long) As Long
    
    Dim dL As Byte
    Dim dH As Byte
    
    dL = DlPortReadPortUchar(Address)
    dH = DlPortReadPortUchar(Address + 1)
    InPut16u = (dH * &H100) + dL

End Function

Public Function InPut16(ByVal Address As Long) As Integer
    
    InPut16 = DlPortReadPortUshort(Address)
    
End Function

Public Function InPut32(ByVal Address As Long) As Long
    
    InPut32 = DlPortReadPortUlong(Address)
    
End Function

Public Sub OutPut8(ByVal Address As Long, ByVal Value As Byte)
    
    Call DlPortWritePortUchar(Address, Value)
    
End Sub

Public Sub OutPut16u(ByVal Address As Long, ByVal Value As Long)
    
    Dim i As Long
    Dim dL As Byte
    Dim dH As Byte
    
    i = (Value And &HFF)
    dL = CByte(i)
    i = (Value And &HFF00) \ 256
    dH = CByte(i)
    
    Call DlPortWritePortUchar(Address, dL)
    Call DlPortWritePortUchar(Address + 1, dH)

End Sub

Public Sub OutPut16(ByVal Address As Long, ByVal Value As Integer)
    
    Call DlPortWritePortUshort(Address, Value)

End Sub

Public Sub OutPut32(ByVal Address As Long, ByVal Value As Long)

    Call DlPortWritePortUlong(Address, Value)

End Sub


