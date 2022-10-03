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
    '레지스트리 함수의 반환 에러 넘버
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
    '목적 : 레지스트리에서 새로운 키를 생성하는 함수
    'Arguments :
    ' RootKeyName : 레지스트리 ROOT키 이름
    ' SubKeyName : 서브키 이름
    'Returns : 함수가 성공적으로 수행되면 ERROR_SUCCESS(0)반환하고 그렇지 않으면
    ' 에러넘버를 반환한다.
    '*****************************************************************************
Public Function CreateRegKey(ByVal RootKeyName As Long, ByVal SubKeyName _
    As String) As Long
    Dim lngRet As Long, lngHKey As Long
    lngRet = RegCreateKey&(HKEY_LOCAL_MACHINE, SubKeyName, lngHKey)
    If (lngRet <> ERROR_SUCCESS) Then GoTo errHandle ' 에러 발생시
        CreateRegKey = ERROR_SUCCESS
        RegCloseKey lngHKey
    Exit Function
errHandle:
    CreateRegKey = lngRet
    RegCloseKey lngHKey
End Function

    '*****************************************************************************
    '목적 : 레지스트리 특정키의 모든 값이름과 값을 가져오는 함수
    'Arguments :
    ' RootKeyName : 레지스트리 ROOT키 이름
    ' SubKeyName : 서브키 이름
    ' strEnvNames() : 값 이름이 저장될 스트링 배열(Call by reference로 값을 넘겨
    ' 받는다.)
    'strEnvValues() : 값이 저장될 스트링 변수 배열(Call by reference로 값을
    ' 넘겨 받는다.)
    '*참고사항
    '1. 읽어올 레지스트리 값의 길이는 1024가 넘지 않도록 한다.
    '2. 읽어올 데이타 타입이 더블워드형이거나 이진형태이면 헥사값을 스트링형으로
    ' 반환한다.
    'Returns : 함수가 성공적으로 수행되면 ERROR_SUCCESS(0)반환하고 그렇지 않으면
    ' 에러넘버를 반환한다.
    ' 에러넘버가 -1 이면 특정키에 데이타가 없음
    '*****************************************************************************

    '*****************************************************************************
    '목적 : 레지스트리 특정키의 특정 이름에서 값을 가져오는 함수
    'Arguments :
    ' RootKeyName : 레지스트리 ROOT키 이름
    ' SubKeyName : 서브키이름
    ' ValueName : 값이름
    ' KeyValue : 값이저장될 포인터 즉,스트링 변수(Call by reference로 값을 넘겨
    ' 받는다.)
    ' *참고사항
    ' 1. 읽어올 레지스트리 값의 길이는 1024가 넘지 않도록 한다.
    ' 2. 읽어올 데이타 타입이 더블워드형이거나 이진형태이면 헥사값을 스트링형으로
    'Returns : 함수가 성공적으로 수행되면 ERROR_SUCCESS(0)반환하고 그렇지 않으면
    ' 에러넘버를 반환한다.
    '*****************************************************************************
Public Function GetRegValue(ByVal RootKeyName As Long, ByVal SubKeyName _
    As String, ByVal ValueName As String, KeyValue As String) As Long
    Dim i As Long ' 루프 카운터
    Dim lngRet As Long ' API함수 리턴값
    Dim lngHKey As Long ' Open된 레지스트리 핸들값
    Dim lngKeyValType As Long ' 읽어올 데이타 타입
    Dim bytTmpVal(1024) As Byte ' 읽어온 값이저장될 임시장소
    Dim strTmp As String ' 읽어온 값이저장될 임시장소
    Dim lngKeyValSize As Long ' 읽어온 데이타의 크기
    strTmp = String(1024, 0)
    '------------------------------------------------------------
    '레지스트리 키값 오픈
    '------------------------------------------------------------
    lngRet = RegOpenKeyEx(RootKeyName, SubKeyName, 0, _
    KEY_ALL_ACCESS, lngHKey)
    If (lngRet <> ERROR_SUCCESS) Then GoTo errHandle ' 에러 발생시
    lngKeyValSize = 1024 ' 읽어올 바이트 수
    '------------------------------------------------------------
    ' 레지스트리 값을 읽어옴
    '------------------------------------------------------------
    lngRet = RegQueryValueEx(lngHKey, ValueName, 0, lngKeyValType, _
    bytTmpVal(0), lngKeyValSize)
    Select Case lngKeyValType
        Case REG_SZ ' 레지스트리의 값 타입이 스트링형이면
            lngRet = RegQueryValueEx(lngHKey, ValueName, 0, lngKeyValType, _
                ByVal strTmp, lngKeyValSize)
        Case REG_DWORD, REG_BINARY ' 레지스트리의 값 타입이 Double Word형이면
            lngRet = RegQueryValueEx(lngHKey, ValueName, 0, lngKeyValType, _
                bytTmpVal(0), lngKeyValSize)
    End Select
    If (lngRet <> ERROR_SUCCESS) Then GoTo errHandle ' 에러 발생시

    If (bytTmpVal(lngKeyValSize - 1) = 0) Then ' Win95에서는 문자열끝에 Null
    ' Terminated String포함
    lngKeyValSize = lngKeyValSize - 1 ' WinNT에서는 문자열끝어 Null
    ' Terminate String포함안함
    End If
    '------------------------------------------------------------
    ' 데이타 타입에 따라서 데이타 Conversion
    '------------------------------------------------------------
    Select Case lngKeyValType
        Case REG_SZ ' 레지스트리의 값 타입이 스트링형이면
            'For i = lngKeyValSize To 1 Step -1
            '    If Asc(Mid(strTmp, i, 1)) = 0 Then
            '        KeyValue = Left(strTmp, i - 1)
            '    Else
            '        Exit For
            '    End If
            'Next
            KeyValue = strTmp
        Case REG_DWORD, REG_BINARY ' 레지스트리의 값 타입이 DoubleWord형이면
            For i = lngKeyValSize - 1 To 0 Step -1 ' 또는 Bynary형이면
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
    '목적 : 레지스트리의 특정키의 특정 이름의 값을 저장하는 함수
    'Arguments :
    ' RootKeyName : 레지스트리 ROOT키 이름
    ' SubKeyName : 서브키 이름
    ' ValueName : 값이름
    ' KeyValue : 저장할 값
    ' DataType : 저장할 데이타 타입
    ' 1. 스트링타입 : 1(정수)
    ' 2. 바이너리타입 : 3(정수)
    ' 3. 더블워드타입 : 4(정수)
    '*참고사항
    '1. 저장할 데이타 타입이 스트링타입은 String형으로 더블워드타입과
    ' 바이너리타입은 LONG 형으로(십진수양의정수형태로) KeyVlaue Argument에
    ' 넘겨준다.
    '2. 저장할 데이타의 길이는 스트링타입은 1024byte가 넘지 않도록(권장사항)
    ' 하고 더블워드나 바이너리타입은 값의 범위가 Long형의 범위를 넘지않도록한다
    'Returns : 함수가 성공적으로 수행되면 ERROR_SUCCESS(0)반환하고 그렇지 않으면
    ' 에러넘버를 반환한다.
    '*****************************************************************************
    Public Function SaveRegValue(ByVal RootKeyName As Long, ByVal SubKeyName As _
    String, ByVal ValueName As String, ByVal KeyValue, _
    ByVal DataType As Integer) As Long
    Dim lngRet As Long ' API함수 리턴값
    Dim lngHKey As Long ' Open된 레지스트리 핸들값
    Dim lngLen As Long
    '------------------------------------------------------------
    '레지스트리 키값 오픈
    '------------------------------------------------------------
    lngRet = RegOpenKeyEx&(HKEY_LOCAL_MACHINE, SubKeyName, 0, _
    KEY_ALL_ACCESS, lngHKey)
    If lngRet <> ERROR_SUCCESS Then GoTo errHandle ' 에러 발생시

    Select Case DataType
        Case REG_SZ ' 저장할데이타의 타입이 스트링형이면
            Dim strTmp As String
            strTmp = CStr(KeyValue)
            lngLen = Len(strTmp)
            If lngLen = 0 Then
            lngLen = 0
            strTmp = ""
            End If
            lngRet = RegSetValueEx(lngHKey, ValueName, 0, DataType, _
            ByVal strTmp, lngLen)
        Case REG_BINARY, REG_DWORD ' 저장할데이타의 타입이 Binary형이거나
            Dim lngTmp As Long ' Double Word형이면
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
