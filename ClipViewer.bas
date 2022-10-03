Attribute VB_Name = "basClipViewer"
'Option Explicit
Public Const CSIDL_DESKTOPDIRECTORY = &H10 '바탕화면의 폴더(jma-2004.09.02)
Public Type SHITEMID
    cb As Long
    abID As Byte
End Type
   
Public Type ITEMIDLIST
    mkid As SHITEMID
End Type
   
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Public Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long


Public Const MAX_ITEMS = 100

Public Const CF_TEXT = 1
Public Const GMEM_SHARE = &H2000&
Public Const GMEM_MOVEABLE = &H2
Public Const GMEM_ZEROINIT = &H40
Public Const FOR_CLIPBOARD = GMEM_MOVEABLE Or GMEM_SHARE Or GMEM_ZEROINIT

Public Const GWL_WNDPROC = -4

Public Const WM_DRAWCLIPBOARD = &H308
Public Const WM_CHANGECBCHAIN = &H30D

Public sItem(0 To MAX_ITEMS - 1) As String
Public cItems As Long

Public hPrevWndProc As Long
Public bIsSubclassed As Boolean
Public bAddToList As Boolean

Public hNextViewer As Long

' User Declare
Public hwndForm As Long

' For subclass
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
   ByVal hwnd As Long, _
   ByVal nIndex As Long, _
   ByVal dwNewLong As Long _
) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" ( _
   ByVal lpPrevWndFunc As Long, _
   ByVal hwnd As Long, _
   ByVal Msg As Long, _
   ByVal wParam As Long, _
   ByVal lParam As Long _
) As Long

' Clipboard API
Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
   ByVal hwnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any _
) As Long

Declare Function EmptyClipboard Lib "user32" () As Long

Declare Function GlobalSize Lib "Kernel32" ( _
   ByVal hMem As Long _
) As Long

Declare Function IsClipboardFormatAvailable Lib "user32" ( _
   ByVal wFormat As Long _
) As Long

Declare Function OpenClipboard Lib "user32" ( _
   ByVal hwnd As Long _
) As Long

Declare Function CloseClipboard Lib "user32" () As Long

Declare Function GetClipboardData Lib "user32" ( _
   ByVal wFormat As Long _
) As Long

Declare Function GlobalLock Lib "Kernel32" ( _
   ByVal hMem As Long _
) As Long

Declare Function GlobalUnlock Lib "Kernel32" ( _
   ByVal hMem As Long _
) As Long

Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" ( _
   Destination As Any, _
   Source As Any, _
   ByVal length As Long)
   
Declare Function SetClipboardData Lib "user32" ( _
   ByVal wFormat As Long, _
   ByVal hMem As Long _
) As Long

Declare Function GlobalAlloc Lib "Kernel32" ( _
   ByVal wFlags As Long, _
   ByVal dwBytes As Long _
) As Long
   
' For viewer
Declare Function ChangeClipboardChain Lib "user32" ( _
   ByVal hwnd As Long, _
   ByVal hWndNext As Long _
) As Long

Declare Function SetClipboardViewer Lib "user32" ( _
   ByVal hwnd As Long _
) As Long

Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
   ByVal lpClassName As String, _
   ByVal lpWindowName As String _
) As Long

Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Sub CopyTextToClipboard(sText As String)
Dim hMem As Long, pMem As Long
 
    If StationNo = 1 Then
       sText = "S1=" & sText
    End If
    
    hMem = GlobalAlloc(FOR_CLIPBOARD, LenB(sText))
    pMem = GlobalLock(hMem)
    CopyMemory ByVal pMem, ByVal sText, LenB(sText)
    GlobalUnlock hMem
    
    If OpenClipboard(hwndForm) <> 0 Then
       EmptyClipboard
       SetClipboardData CF_TEXT, hMem
       CloseClipboard
    End If

End Sub

Public Function PasteTextFromClipboard() As String

Dim hMem As Long, pMem As Long
Dim lMemSize As Long
Dim sText As String

PasteTextFromClipboard = ""

' Check for text on clipboard
If IsClipboardFormatAvailable(CF_TEXT) = 0 Then
   Exit Function
End If

' Open clipboard
If OpenClipboard(hwndForm) <> 0 Then
   hMem = GetClipboardData(CF_TEXT)
   ' If no text, close clipboard and exit
   If hMem = 0 Then
      CloseClipboard
      Exit Function
   Else
      ' Get memory pointer
      pMem = GlobalLock(hMem)
      ' Get size of memory
      lMemSize = GlobalSize(hMem)
      ' Allocate local string
      sText = String$(lMemSize, 0)
      ' Copy clipboard text
      CopyMemory ByVal sText, ByVal pMem, lMemSize
      ' Unlock clipboard memory
      GlobalUnlock hMem
      ' Close clipboard
      CloseClipboard
      ' Return text
      PasteTextFromClipboard = Trim0(sText)
   End If
End If

End Function

Public Function WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Dim sItem As String


Select Case iMsg
   
   Case WM_DRAWCLIPBOARD
      
      ' Get clipboard text and put it into text box
        
        
        If StationNo = 2 Then
           sItem = PasteTextFromClipboard
             If sItem <> "" Then
                 If Left(sItem, 3) = "S1=" Then
                   MT2000.Text_R = sItem
      '            DataConversion = sItem
       '            Dataitem = sItem
      '           sItem = ""
                 End If
             End If
        Else
            sItem = PasteTextFromClipboard
             If sItem <> "" Then
                 If Left(sItem, 3) = "S2=" Then
                    MT2000.Text_R = sItem
      '             DataConversion = sItem
      '              Dataitem = sItem
      '           sItem = ""
                 End If
             End If
        End If
        
        
      ' Send message to next clipboard viewer
      
      If hNextViewer <> 0 Then
         SendMessage hNextViewer, WM_DRAWCLIPBOARD, wParam, lParam
      End If
      
      Exit Function
   
   Case WM_CHANGECBCHAIN
   
      ' Check to see which viewer is being removed.
      ' Is it the next one in line?
      ' wParam contains handle of viewer being removed.
      If wParam = hNextViewer Then
         ' Remove this viewer from chain
         hNextViewer = lParam
      Else
         ' Not the very next viewer so pass request on
         SendMessage hNextViewer, WM_CHANGECBCHAIN, wParam, lParam
      End If
      
      Exit Function

End Select
' Call original window procedure
WindowProc = CallWindowProc(hPrevWndProc, hwnd, iMsg, wParam, lParam)

End Function

Public Function Trim0(sName As String) As String

' Keep left portion of string sName up to first 0. Useful with Win API null terminated strings.

Dim x As Integer
x = InStr(sName, Chr$(0))
If x > 0 Then Trim0 = Left$(sName, x - 1) Else Trim0 = sName

End Function

' jma(2004.09.02)
Public Function GetSpecialPath(CSIDL As Long) As String
    Dim r As Long
    Dim IDL As ITEMIDLIST
    
    r = SHGetSpecialFolderLocation(100, CSIDL, IDL)
'    If r = NOERROR Then
        Path$ = Space$(512)
        r = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal Path$)
        GetSpecialPath = Left$(Path, InStr(Path, Chr$(0)) - 1)
        Exit Function
'    End If
    GetSpecialPath = ""
End Function

