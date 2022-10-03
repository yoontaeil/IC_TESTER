Attribute VB_Name = "Registe"
Option Explicit
    Const KEY_ALL_ACCESS = &H2003F
    Const HKEY_CLASSES_ROOT = &H80000000
    Const HKEY_CURRENT_CONFIG = &H80000005
    Const HKEY_CURRENT_USER = &H80000001
    Const HKEY_DYN_DATA = &H80000006
    Const HKEY_LOCAL_MACHINE = &H80000002
    Const HKEY_USERS = &H80000003
    '------------------------------------------------------------
    '������Ʈ�� �Լ��� ��ȯ ���� �ѹ�
    '------------------------------------------------------------
    Const ERROR_SUCCESS = 0
    Const ERROR_FILE_NOT_FOUND = 2
    Const ERROR_ACCESS_DENIED = 5
    Const ERROR_INVALID_HANDLE = 6
    Const ERROR_BAD_NETPATH = 53
    Const ERROR_INVALID_PARAMETER = 87
    Const ERROR_CALL_NOT_IMPLEMENTED = 120
    Const ERROR_INSUFFICIENT_BUFFER = 122
    Const ERROR_BAD_PATHNAME = 161
    Const ERROR_LOCK_FAILED = 167
    Const ERROR_TRANSFER_TOO_LONG = 222
    Const ERROR_MORE_DATA = 234
    Const ERROR_NO_MORE_ITEMS = 259
    Const ERROR_BADDB = 1009
    Const ERROR_BADKEY = 1010
    Const ERROR_CANTOKEY = 1011
    Const ERROR_CANREAD = 1012
    Const ERROR_CANWRITE = 1013
    Const ERROR_REGISTRY_RECOVERED = 1014
    Const ERROR_REGISTRY_CORRUPT = 1015
    Const ERROR_REGISTRY_IO_FALED = 1016
    Const ERROR_NOT_REGISTRY_FILE = 1017
    Const ERROR_KEY_DELETED = 1018

    Const REG_SZ = 1 ' nul terminated string
    Const REG_BINARY = 3 ' Binery data
    Const REG_DWORD = 4 ' Double Word Number

    Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" _
    (ByVal hKey As Long, ByVal lpSubKey As String)

    Public Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" _
    (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
    ByVal samDesired As Long, ByRef phkResult As Long) As Long

    Public Declare Function RegSetValueEx& Lib "advapi32" Alias _
    "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal _
    Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long)

    Public Declare Function RegQueryValueEx& Lib "advapi32.dll" Alias _
    "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal _
    lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long)

    Public Declare Function RegCloseKey& Lib "advapi32" (ByVal hKey As Long)

    Public Declare Function RegCreateKey& Lib "advapi32" Alias "RegCreateKeyA" _
    (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long)

    Private Declare Function RegEnumValue& Lib "advapi32.dll" Alias _
    "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal _
    lpValueName As String, lpcbValueName As Long, lpReserved As Long, _
    lpType As Long, lpData As Byte, lpcbData As Long)

    Private Declare Function RegQueryInfoKey& Lib "advapi32.dll" Alias _
    "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, lpcbClass _
    As Long, lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, _
    lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, _
    lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, _
    lpftLastWriteTime As Long)

    '****************************************************************************
    '���� : ������Ʈ������ ���ο� Ű�� �����ϴ� �Լ�
    'Arguments :
    ' RootKeyName : ������Ʈ�� ROOTŰ �̸�
    ' SubKeyName : ����Ű �̸�
    'Returns : �Լ��� ���������� ����Ǹ� ERROR_SUCCESS(0)��ȯ�ϰ� �׷��� ������
    ' �����ѹ��� ��ȯ�Ѵ�.
    '*****************************************************************************
Public Function CreateRegKey(ByVal RootKeyName As Long, ByVal SubKeyName _
    As String) As Long
    Dim lngRet As Long, lngHKey As Long
    lngRet = RegCreateKey&(HKEY_LOCAL_MACHINE, SubKeyName, lngHKey)
    If (lngRet <> ERROR_SUCCESS) Then GoTo errHandle ' ���� �߻���
        CreateRegKey = ERROR_SUCCESS
        RegCloseKey lngHKey
    Exit Function
errHandle:
    CreateRegKey = lngRet
    RegCloseKey lngHKey
End Function

    '*****************************************************************************
    '���� : ������Ʈ�� Ư��Ű�� ��� ���̸��� ���� �������� �Լ�
    'Arguments :
    ' RootKeyName : ������Ʈ�� ROOTŰ �̸�
    ' SubKeyName : ����Ű �̸�
    ' strEnvNames() : �� �̸��� ����� ��Ʈ�� �迭(Call by reference�� ���� �Ѱ�
    ' �޴´�.)
    'strEnvValues() : ���� ����� ��Ʈ�� ���� �迭(Call by reference�� ����
    ' �Ѱ� �޴´�.)
    '*�������
    '1. �о�� ������Ʈ�� ���� ���̴� 1024�� ���� �ʵ��� �Ѵ�.
    '2. �о�� ����Ÿ Ÿ���� ����������̰ų� ���������̸� ��簪�� ��Ʈ��������
    ' ��ȯ�Ѵ�.
    'Returns : �Լ��� ���������� ����Ǹ� ERROR_SUCCESS(0)��ȯ�ϰ� �׷��� ������
    ' �����ѹ��� ��ȯ�Ѵ�.
    ' �����ѹ��� -1 �̸� Ư��Ű�� ����Ÿ�� ����
    '*****************************************************************************

    '*****************************************************************************
    '���� : ������Ʈ�� Ư��Ű�� Ư�� �̸����� ���� �������� �Լ�
    'Arguments :
    ' RootKeyName : ������Ʈ�� ROOTŰ �̸�
    ' SubKeyName : ����Ű�̸�
    ' ValueName : ���̸�
    ' KeyValue : ��������� ������ ��,��Ʈ�� ����(Call by reference�� ���� �Ѱ�
    ' �޴´�.)
    ' *�������
    ' 1. �о�� ������Ʈ�� ���� ���̴� 1024�� ���� �ʵ��� �Ѵ�.
    ' 2. �о�� ����Ÿ Ÿ���� ����������̰ų� ���������̸� ��簪�� ��Ʈ��������
    'Returns : �Լ��� ���������� ����Ǹ� ERROR_SUCCESS(0)��ȯ�ϰ� �׷��� ������
    ' �����ѹ��� ��ȯ�Ѵ�.
    '*****************************************************************************
Public Function GetRegValue(ByVal RootKeyName As Long, ByVal SubKeyName _
    As String, ByVal ValueName As String, KeyValue As String) As Long
    Dim i As Long ' ���� ī����
    Dim lngRet As Long ' API�Լ� ���ϰ�
    Dim lngHKey As Long ' Open�� ������Ʈ�� �ڵ鰪
    Dim lngKeyValType As Long ' �о�� ����Ÿ Ÿ��
    Dim bytTmpVal(1024) As Byte ' �о�� ��������� �ӽ����
    Dim strTmp As String ' �о�� ��������� �ӽ����
    Dim lngKeyValSize As Long ' �о�� ����Ÿ�� ũ��
    strTmp = String(1024, 0)
    '------------------------------------------------------------
    '������Ʈ�� Ű�� ����
    '------------------------------------------------------------
    lngRet = RegOpenKeyEx(RootKeyName, SubKeyName, 0, _
    KEY_ALL_ACCESS, lngHKey)
    If (lngRet <> ERROR_SUCCESS) Then GoTo errHandle ' ���� �߻���
    lngKeyValSize = 1024 ' �о�� ����Ʈ ��
    '------------------------------------------------------------
    ' ������Ʈ�� ���� �о��
    '------------------------------------------------------------
    lngRet = RegQueryValueEx(lngHKey, ValueName, 0, lngKeyValType, _
    bytTmpVal(0), lngKeyValSize)
    Select Case lngKeyValType
        Case REG_SZ ' ������Ʈ���� �� Ÿ���� ��Ʈ�����̸�
            lngRet = RegQueryValueEx(lngHKey, ValueName, 0, lngKeyValType, _
                ByVal strTmp, lngKeyValSize)
        Case REG_DWORD, REG_BINARY ' ������Ʈ���� �� Ÿ���� Double Word���̸�
            lngRet = RegQueryValueEx(lngHKey, ValueName, 0, lngKeyValType, _
                bytTmpVal(0), lngKeyValSize)
    End Select
    If (lngRet <> ERROR_SUCCESS) Then GoTo errHandle ' ���� �߻���

    If (bytTmpVal(lngKeyValSize - 1) = 0) Then ' Win95������ ���ڿ����� Null
    ' Terminated String����
    lngKeyValSize = lngKeyValSize - 1 ' WinNT������ ���ڿ����� Null
    ' Terminate String���Ծ���
    End If
    '------------------------------------------------------------
    ' ����Ÿ Ÿ�Կ� ���� ����Ÿ Conversion
    '------------------------------------------------------------
    Select Case lngKeyValType
        Case REG_SZ ' ������Ʈ���� �� Ÿ���� ��Ʈ�����̸�
            'For i = lngKeyValSize To 1 Step -1
            '    If Asc(Mid(strTmp, i, 1)) = 0 Then
            '        KeyValue = Left(strTmp, i - 1)
            '    Else
            '        Exit For
            '    End If
            'Next
            KeyValue = strTmp
        Case REG_DWORD, REG_BINARY ' ������Ʈ���� �� Ÿ���� DoubleWord���̸�
            For i = lngKeyValSize - 1 To 0 Step -1 ' �Ǵ� Bynary���̸�
                If bytTmpVal(i) > 16 Then
                    KeyValue = KeyValue + Hex(bytTmpVal(i))
                Else
                    KeyValue = KeyValue + "0" + Hex(bytTmpVal(i))
                End If
            Next
    End Select
    
    GetRegValue = ERROR_SUCCESS ' Return Success
    lngRet = RegCloseKey(lngHKey) ' Close Registry Key
    Exit Function ' Exit
errHandle:
    KeyValue = ""
    GetRegValue = lngRet
    RegCloseKey lngHKey
    End Function

    '*****************************************************************************
    '���� : ������Ʈ���� Ư��Ű�� Ư�� �̸��� ���� �����ϴ� �Լ�
    'Arguments :
    ' RootKeyName : ������Ʈ�� ROOTŰ �̸�
    ' SubKeyName : ����Ű �̸�
    ' ValueName : ���̸�
    ' KeyValue : ������ ��
    ' DataType : ������ ����Ÿ Ÿ��
    ' 1. ��Ʈ��Ÿ�� : 1(����)
    ' 2. ���̳ʸ�Ÿ�� : 3(����)
    ' 3. �������Ÿ�� : 4(����)
    '*�������
    '1. ������ ����Ÿ Ÿ���� ��Ʈ��Ÿ���� String������ �������Ÿ�԰�
    ' ���̳ʸ�Ÿ���� LONG ������(�����������������·�) KeyVlaue Argument��
    ' �Ѱ��ش�.
    '2. ������ ����Ÿ�� ���̴� ��Ʈ��Ÿ���� 1024byte�� ���� �ʵ���(�������)
    ' �ϰ� ������峪 ���̳ʸ�Ÿ���� ���� ������ Long���� ������ �����ʵ����Ѵ�
    'Returns : �Լ��� ���������� ����Ǹ� ERROR_SUCCESS(0)��ȯ�ϰ� �׷��� ������
    ' �����ѹ��� ��ȯ�Ѵ�.
    '*****************************************************************************
    Public Function SaveRegValue(ByVal RootKeyName As Long, ByVal SubKeyName As _
    String, ByVal ValueName As String, ByVal KeyValue, _
    ByVal DataType As Integer) As Long
    Dim lngRet As Long ' API�Լ� ���ϰ�
    Dim lngHKey As Long ' Open�� ������Ʈ�� �ڵ鰪
    Dim lngLen As Long
    '------------------------------------------------------------
    '������Ʈ�� Ű�� ����
    '------------------------------------------------------------
    lngRet = RegOpenKeyEx&(HKEY_LOCAL_MACHINE, SubKeyName, 0, _
    KEY_ALL_ACCESS, lngHKey)
    If lngRet <> ERROR_SUCCESS Then GoTo errHandle ' ���� �߻���

    Select Case DataType
        Case REG_SZ ' �����ҵ���Ÿ�� Ÿ���� ��Ʈ�����̸�
            Dim strTmp As String
            strTmp = CStr(KeyValue)
            lngLen = Len(strTmp)
            If lngLen = 0 Then
            lngLen = 0
            strTmp = ""
            End If
            lngRet = RegSetValueEx(lngHKey, ValueName, 0, DataType, _
            ByVal strTmp, lngLen)
        Case REG_BINARY, REG_DWORD ' �����ҵ���Ÿ�� Ÿ���� Binary���̰ų�
            Dim lngTmp As Long ' Double Word���̸�
            lngTmp = CLng(KeyValue)
            lngRet = RegSetValueEx(lngHKey, ValueName, 0, DataType, _
            lngTmp, Len(lngTmp))
    End Select

    If lngRet <> ERROR_SUCCESS Then GoTo errHandle
    lngRet = RegCloseKey(lngHKey)
    SaveRegValue = ERROR_SUCCESS
    Exit Function
errHandle:
    SaveRegValue = lngRet
    RegCloseKey (lngHKey)
    End Function
