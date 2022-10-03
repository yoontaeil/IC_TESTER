Attribute VB_Name = "Module_DLL"
Option Explicit

'I/O Win98,NT40 Input-8bit, Output-8bit
Public Declare Function DlPortReadPortUchar Lib "dlportio.dll" (ByVal Port As Long) As Byte
Public Declare Function DlPortReadPortUshort Lib "dlportio.dll" (ByVal Port As Long) As Integer
Public Declare Function DlPortReadPortUlong Lib "dlportio.dll" (ByVal Port As Long) As Long

Public Declare Sub DlPortWritePortUchar Lib "dlportio.dll" (ByVal Port As Long, ByVal value As Integer)
Public Declare Sub DlPortWritePortUshort Lib "dlportio.dll" (ByVal Port As Long, ByVal value As Integer)
Public Declare Sub DlPortWritePortUlong Lib "dlportio.dll" (ByVal Port As Long, ByVal value As Long)


'[ 2022.07.29 ] : form의 "X"버튼 관련
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hmenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Const MF_BYPOSITION = &H400&



'Public Function InPut8(ByVal address As Long) As Byte
'
'    InPut8 = DlPortReadPortUchar(address)
'
'End Function

Public Function vbInp(ByVal address As Long) As Byte
    If DemoMode = 1 Then Exit Function              ' 2004.03.15
    vbInp = DlPortReadPortUchar(address)
End Function

Public Function InPut16u(ByVal address As Long) As Long
    Dim dL As Byte
    Dim dH As Byte

    dL = DlPortReadPortUchar(address)
    dH = DlPortReadPortUchar(address + 1)
    InPut16u = (dH * &H100) + dL
End Function

Public Function InPut16(ByVal address As Long) As Integer
    InPut16 = DlPortReadPortUshort(address)
End Function

Public Function InPut32(ByVal address As Long) As Long
    InPut32 = DlPortReadPortUlong(address)
End Function

'Public Sub OutPut8(ByVal address As Long, ByVal Value As Byte)
'
'    Call DlPortWritePortUchar(address, Value)
'
'End Sub
Public Sub vbOut(ByVal address As Long, ByVal value As Integer)
    If DemoMode = 1 Then Exit Sub                '2004.03.15
'     If value < 0 Then value = value + 256
    Call DlPortWritePortUchar(address, value)
End Sub

Public Sub OutPut16u(ByVal address As Long, ByVal value As Long)
    Dim i As Long
    Dim dL As Byte
    Dim dH As Byte

    i = (value And &HFF)
    dL = CByte(i)
    i = (value And &HFF00) \ 256
    dH = CByte(i)

    Call DlPortWritePortUchar(address, dL)
    Call DlPortWritePortUchar(address + 1, dH)
End Sub

Public Sub OutPut16(ByVal address As Long, ByVal value As Integer)
    Call DlPortWritePortUshort(address, value)
End Sub

Public Sub OutPut32(ByVal address As Long, ByVal value As Long)
    Call DlPortWritePortUlong(address, value)
End Sub

Public Function Input8b(ByVal address As Long) As Byte
    Input8b = DlPortReadPortUchar(address)
End Function

Public Function Input8w(ByVal address As Long) As Integer
    Input8w = DlPortReadPortUshort(address)
End Function

Public Function Input8d(ByVal address As Long) As Long
    Input8d = DlPortReadPortUlong(address)
End Function

Public Sub Output8b(ByVal address As Long, ByVal value As Byte)
    Call DlPortWritePortUchar(address, value)
End Sub

Public Sub Output8w(ByVal address As Long, ByVal value As Integer)
    Call DlPortWritePortUshort(address, value)
End Sub

Public Sub Output8d(ByVal address As Long, ByVal value As Long)
    Call DlPortWritePortUlong(address, value)
End Sub




'[ 2022.07.29 ] : X버튼 제거
Public Sub RemoveCancelMenuItem(frm As Form)
    Dim hSysMenu As Long
    
    hSysMenu = GetSystemMenu(frm.hwnd, 0)
    Call RemoveMenu(hSysMenu, 6, MF_BYPOSITION)
    Call RemoveMenu(hSysMenu, 5, MF_BYPOSITION)
End Sub
