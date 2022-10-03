Attribute VB_Name = "port_io"
Declare Sub vbOut Lib "e:\MT2000\PORT_IO.DLL" (ByVal nPort As Integer, ByVal nData As Integer)
Declare Function vbInp Lib "e:\MT2000\PORT_IO.DLL" (ByVal nPort As Integer) As Integer
