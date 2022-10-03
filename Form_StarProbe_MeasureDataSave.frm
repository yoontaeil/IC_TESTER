VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form_StarProbe_MeasureDataSave 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "Measure Data Save ..."
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12150
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form_StarProbe_MeasureDataSave.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   543
   ScaleMode       =   3  '픽셀
   ScaleWidth      =   810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin TabDlg.SSTab SSTab1 
      Height          =   7935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   13996
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "View"
      TabPicture(0)   =   "Form_StarProbe_MeasureDataSave.frx":08CA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "VScroll_Zoom"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "HScroll_Zoom"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "SSPanel1"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Measure"
      TabPicture(1)   =   "Form_StarProbe_MeasureDataSave.frx":08E6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "SSPanel2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "HScroll_Measure"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "VScroll_Measure"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.VScrollBar VScroll_Measure 
         Height          =   7095
         LargeChange     =   100
         Left            =   11520
         Max             =   1000
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   480
         Width           =   255
      End
      Begin VB.HScrollBar HScroll_Measure 
         Height          =   255
         LargeChange     =   100
         Left            =   120
         Max             =   1000
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   7560
         Width           =   11415
      End
      Begin VB.VScrollBar VScroll_Zoom 
         Height          =   7095
         LargeChange     =   100
         Left            =   -63480
         Max             =   1000
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   480
         Width           =   255
      End
      Begin VB.HScrollBar HScroll_Zoom 
         Height          =   255
         LargeChange     =   100
         Left            =   -74880
         Max             =   1000
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   7560
         Width           =   11415
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   7095
         Left            =   -74880
         TabIndex        =   3
         Top             =   480
         Width           =   11415
         _Version        =   65536
         _ExtentX        =   20135
         _ExtentY        =   12515
         _StockProps     =   15
         BackColor       =   16777152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         Begin VB.PictureBox pView 
            Appearance      =   0  '평면
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   13095
            Left            =   0
            ScaleHeight     =   873
            ScaleMode       =   3  '픽셀
            ScaleWidth      =   1125
            TabIndex        =   4
            Top             =   0
            Width           =   16875
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   7095
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   11415
         _Version        =   65536
         _ExtentX        =   20135
         _ExtentY        =   12515
         _StockProps     =   15
         BackColor       =   16777152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         Begin VB.PictureBox pMeasure 
            Appearance      =   0  '평면
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   13095
            Left            =   0
            ScaleHeight     =   873
            ScaleMode       =   3  '픽셀
            ScaleWidth      =   1125
            TabIndex        =   8
            Top             =   0
            Width           =   16875
         End
      End
   End
End
Attribute VB_Name = "Form_StarProbe_MeasureDataSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub Display_View5()             'skip die 포함
    Dim i As Integer, j As Integer, k As Integer
    Dim x As Integer, y As Integer
    Dim dx As Integer, dy As Integer
    Dim forx As Integer, fory As Integer
    Dim xfrom As Integer, xto As Integer
    Dim yfrom As Integer, yto As Integer
    Dim ifreefile As Integer
    Dim tempbinno As Byte
    Dim Count_Total As Long, Count_Good As Long
    Dim Count_Bad As Long, Count_Skip As Long
    Dim BinCount(0 To 26) As Long
    Dim bexit As Boolean
    Dim sfilename As String
    Dim buf As String
    Dim TEMP_BIN As String
    
    On Error GoTo err
    
    Erase WaferTemp
    
    For y = 0 To StarProbe.ChipCountY
        For x = 0 To StarProbe.ChipCountX
            WaferTemp(x, y) = Wafer(x, y)
        Next
    Next
    
    Count_Total = 0
    Count_Good = 0
    Count_Bad = 0
    Count_Skip = 0
    
    For i = 0 To 26
        BinCount(i) = 0
    Next
        
    For y = 0 To StarProbe.ChipCountY
        For x = 0 To StarProbe.ChipCountX
            Select Case XPitch(TT_NO)
                Case 2: xfrom = x:     xto = x + 1
                Case 3: xfrom = x - 1: xto = x + 1
                Case 4: xfrom = x - 1: xto = x + 2
                Case 5: xfrom = x - 2: xto = x + 2
                Case 6: xfrom = x - 2: xto = x + 3
                Case 7: xfrom = x - 3: xto = x + 3
                Case 8: xfrom = x - 3: xto = x + 4
                Case 9: xfrom = x - 4: xto = x + 4
                Case 10: xfrom = x - 4: xto = x + 5
                Case 11: xfrom = x - 5: xto = x + 5
                Case 12: xfrom = x - 5: xto = x + 6
                Case 13: xfrom = x - 6: xto = x + 6
                Case 14: xfrom = x - 6: xto = x + 7
                Case 15: xfrom = x - 7: xto = x + 7
                Case 16: xfrom = x - 7: xto = x + 8
                Case 17: xfrom = x - 8: xto = x + 8
                Case 18: xfrom = x - 8: xto = x + 9
                Case 19: xfrom = x - 9: xto = x + 9
                Case 20: xfrom = x - 9: xto = x + 10
                Case 21: xfrom = x - 10: xto = x + 10
                Case 22: xfrom = x - 10: xto = x + 11
                Case 23: xfrom = x - 11: xto = x + 11
                Case 24: xfrom = x - 11: xto = x + 12
                Case 25: xfrom = x - 12: xto = x + 12
                Case 26: xfrom = x - 12: xto = x + 13
                Case 27: xfrom = x - 13: xto = x + 13
                Case 28: xfrom = x - 13: xto = x + 14
                Case 29: xfrom = x - 14: xto = x + 14
                Case 30: xfrom = x - 14: xto = x + 15
            End Select
                                                 '4channel fix
            'yfrom = y - 1: yto = y + 2
            Select Case YPitch(TT_NO)
                Case 2: yfrom = y:     yto = y + 1
                Case 3: yfrom = y - 1: yto = y + 1
                Case 4: yfrom = y - 1: yto = y + 2
                Case 5: yfrom = y - 2: yto = y + 2
                Case 6: yfrom = y - 2: yto = y + 3
                Case 7: yfrom = y - 3: yto = y + 3
                Case 8: yfrom = y - 3: yto = y + 4
                Case 9: yfrom = y - 4: yto = y + 4
                Case 10: yfrom = y - 4: yto = y + 5
                Case 11: yfrom = y - 5: yto = y + 5
                Case 12: yfrom = y - 5: yto = y + 6
                Case 13: yfrom = y - 6: yto = y + 6
                Case 14: yfrom = y - 6: yto = y + 7
                Case 15: yfrom = y - 7: yto = y + 7
                Case 16: yfrom = y - 7: yto = y + 8
                Case 17: yfrom = y - 8: yto = y + 8
                Case 18: yfrom = y - 8: yto = y + 9
                Case 19: yfrom = y - 9: yto = y + 9
                Case 20: yfrom = y - 9: yto = y + 10
                Case 21: yfrom = y - 10: yto = y + 10
                Case 22: yfrom = y - 10: yto = y + 11
                Case 23: yfrom = y - 11: yto = y + 11
                Case 24: yfrom = y - 11: yto = y + 12
                Case 25: yfrom = y - 12: yto = y + 12
                Case 26: yfrom = y - 12: yto = y + 13
                Case 27: yfrom = y - 13: yto = y + 13
                Case 28: yfrom = y - 13: yto = y + 14
                Case 29: yfrom = y - 14: yto = y + 14
                Case 30: yfrom = y - 14: yto = y + 15
            End Select

            If WaferTemp(x, y).Chip And Not WaferTemp(x, y).ChipMask And Not WaferTemp(x, y).ChipSkipDie And Not WaferTemp(x, y).ChipPlate And WaferTemp(x, y).flag And WaferTemp(x, y).ChipMeasure Then
                If WaferTemp(x, y).FlagBad Then
                    bexit = False
                    For fory = y To yfrom Step -1
                        For forx = x To xfrom Step -1
                            If WaferTemp(forx, fory).Chip And Not WaferTemp(forx, fory).ChipMask And Not WaferTemp(forx, fory).ChipSkipDie And Not WaferTemp(forx, fory).ChipPlate And WaferTemp(forx, fory).flag And Not WaferTemp(forx, fory).FlagBad Then
                                tempbinno = WaferTemp(forx, fory).BIN
                                bexit = True
                                Exit For
                            End If
                        Next

                        If bexit Then Exit For

                        For forx = x To xto
                            If WaferTemp(forx, fory).Chip And Not WaferTemp(forx, fory).ChipMask And Not WaferTemp(forx, fory).ChipSkipDie And Not WaferTemp(forx, fory).ChipPlate And WaferTemp(forx, fory).flag And Not WaferTemp(forx, fory).FlagBad Then
                                tempbinno = WaferTemp(forx, fory).BIN
                                bexit = True
                                Exit For
                            End If
                        Next
                        If bexit Then Exit For
                    Next

                    For fory = y To yto
                        If bexit Then Exit For
                        For forx = x To xfrom Step -1
                            If WaferTemp(forx, fory).Chip And Not WaferTemp(forx, fory).ChipMask And Not WaferTemp(forx, fory).ChipSkipDie And Not WaferTemp(forx, fory).ChipPlate And WaferTemp(forx, fory).flag And Not WaferTemp(forx, fory).FlagBad Then
                                tempbinno = WaferTemp(forx, fory).BIN
                                bexit = True
                                Exit For
                            End If
                        Next

                        If bexit Then Exit For

                        For forx = x To xto
                            If WaferTemp(forx, fory).Chip And Not WaferTemp(forx, fory).ChipMask And Not WaferTemp(forx, fory).ChipSkipDie And Not WaferTemp(forx, fory).ChipPlate And WaferTemp(forx, fory).flag And Not WaferTemp(forx, fory).FlagBad Then
                                tempbinno = WaferTemp(forx, fory).BIN
                                bexit = True
                                Exit For
                            End If
                        Next
                        If bexit Then Exit For
                    Next
                Else
                    tempbinno = WaferTemp(x, y).BIN
                End If

                For forx = xfrom To xto
                    For fory = yfrom To yto
                        If WaferTemp(forx, fory).Chip And Not WaferTemp(forx, fory).ChipMask And Not WaferTemp(forx, fory).ChipSkipDie And Not WaferTemp(forx, fory).ChipMeasure And Not WaferTemp(forx, fory).MeasureWait And Not WaferTemp(forx, fory).ChipPlate And Not WaferTemp(forx, fory).flag And Not WaferTemp(forx, fory).FlagBad And Not ((x = forx) And (y = fory)) Then
                            WaferTemp(forx, fory).flag = True
                            WaferTemp(forx, fory).BIN = tempbinno
                        End If
                    Next
                Next
            End If
        Next
    Next

    For y = 0 To StarProbe.ChipCountY
        For x = 0 To StarProbe.ChipCountX
            If WaferTemp(x, y).Chip And Not WaferTemp(x, y).ChipMask And Not WaferTemp(x, y).ChipSkipDie And Not WaferTemp(x, y).MeasureWait And Not WaferTemp(x, y).ChipPlate And Not WaferTemp(x, y).flag And Not WaferTemp(x, y).FlagBad Then
                If WaferTemp(x, y - 1).flag And Not WaferTemp(x, y - 1).FlagBad Then
                    tempbinno = WaferTemp(x, y - 1).BIN
                ElseIf WaferTemp(x + 1, y).flag And Not WaferTemp(x + 1, y).FlagBad Then
                    tempbinno = WaferTemp(x + 1, y).BIN
                ElseIf WaferTemp(x, y + 1).flag And Not WaferTemp(x, y + 1).FlagBad Then
                    tempbinno = WaferTemp(x, y + 1).BIN
                ElseIf WaferTemp(x - 1, y).flag And Not WaferTemp(x - 1, y).FlagBad Then
                    tempbinno = WaferTemp(x - 1, y).BIN
                End If
                WaferTemp(x, y).flag = True
                WaferTemp(x, y).BIN = tempbinno
            End If
        Next
    Next

    For y = 0 To StarProbe.ChipCountY
        For x = 0 To StarProbe.ChipCountX
            If WaferTemp(x, y).Chip Then
                Count_Total = Count_Total + 1
                If WaferTemp(x, y).ChipMask Or WaferTemp(x, y).ChipSkipDie Or WaferTemp(x, y).ChipPlate Then
                    Count_Skip = Count_Skip + 1
                Else 'If WaferTemp(x, y).flag Then
                    If WaferTemp(x, y).FlagBad Then
                        Count_Bad = Count_Bad + 1
                    Else
                        Count_Good = Count_Good + 1
'                        WaferTemp(x, y).BIN = GOOD_BIN_NO
                    End If
                    BinCount(WaferTemp(x, y).BIN) = BinCount(WaferTemp(x, y).BIN) + 1
                End If
            End If
        Next
    Next
        
        
    'map
    If UCase(Right(MT2000.SSPanel2(0), 3)) = ".RP" Then      '[ 2020.06.23 ] : double check
        sfilename = FILE_NAMEING & ".wmd01"
    Else
        sfilename = FILE_NAMEING & ".wmd01"
    End If
    ifreefile = FreeFile
            
    Open sfilename For Output As ifreefile
        '파일 내용을 XML형식으로 저장한다.
        Print #ifreefile, "" 'Format(Now, "AMPM hh:mm YYYY-MM-DD")           '07.11.21날짜변경시 오류발견 수정.
        Print #ifreefile, "  <?xml version=" & """1.0""" & " encoding=" & """utf-8""" & "?>"
        Print #ifreefile, "  <Maps>"
        Print #ifreefile, "  <Map xmlns=" & """http://www.semi.org""" & _
                        " SubstrateId=" & """"; LOT & "_" & MT2000.lblWafer; """" & _
                        " SubstrateType=" & """Wafer""" & _
                        " FormatRevision=" & """SEMI G85-0703""" & ">"
        Print #ifreefile, "    <Device Rows=" & """"; Trim(StarProbe.ChipCountY); """" & _
                             " LotId=" & """"; LOT; """" & _
                             " BinType=" & """" & "ASCII" & """" & _
                             " Columns=" & """"; Trim(StarProbe.ChipCountX); """" & _
                             " MapType=" & """" & "Array" & """" & _
                             " NullBin=" & """" & "0" & """" & _
                             " ProductId=" & """"; Trim(DEV); """" & _
                             " WaferSize=" & """"; Trim(StarProbe.WaferSizemm); """" & _
                             " CreateDate=" & """"; Format(Now, "YYYYMMDDhhmmssms") & """" & _
                             " DeviceSizeX=" & """"; Format(Trim(StarProbe.ChipSizeX * 1000), "0000.00"); """" & _
                             " DeviceSizeY=" & """"; Format(Trim(StarProbe.ChipSizeY * 1000), "0000.00"); """" & _
                             " LotNo=" & """"; Trim(LOT); """" & _
                             " Orientation=" & """" & "0" & """" & _
                             " SupplierName=" & """" & "" & """" & _
                             " Originlocation=" & """" & "2" & """" & ">"
        Print #ifreefile, "    <ReferenceDevice ReferenceDeviceX=" & """" & O_X & """" & " " & "Referencedevice=" & """" & O_Y & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "0" & """" & " BinCount=" & """"; Trim(BinCount(0)); """" & " BinQuality=" & """" & "BIN0" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "1" & """" & " BinCount=" & """"; Trim(BinCount(1)); """" & " BinQuality=" & """" & "BIN1" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "2" & """" & " BinCount=" & """"; Trim(BinCount(2)); """" & " BinQuality=" & """" & "BIN2" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "3" & """" & " BinCount=" & """"; Trim(BinCount(3)); """" & " BinQuality=" & """" & "BIN3" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "4" & """" & " BinCount=" & """"; Trim(BinCount(4)); """" & " BinQuality=" & """" & "BIN4" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "5" & """" & " BinCount=" & """"; Trim(BinCount(5)); """" & " BinQuality=" & """" & "BIN5" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "6" & """" & " BinCount=" & """"; Trim(BinCount(6)); """" & " BinQuality=" & """" & "BIN6" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "7" & """" & " BinCount=" & """"; Trim(BinCount(7)); """" & " BinQuality=" & """" & "BIN7" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "8" & """" & " BinCount=" & """"; Trim(BinCount(8)); """" & " BinQuality=" & """" & "BIN8" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "9" & """" & " BinCount=" & """"; Trim(BinCount(9)); """" & " BinQuality=" & """" & "BIN9" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "A" & """" & " BinCount=" & """"; Trim(BinCount(10)); """" & " BinQuality=" & """" & "BIN10" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "B" & """" & " BinCount=" & """"; Trim(BinCount(11)); """" & " BinQuality=" & """" & "BIN11" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "C" & """" & " BinCount=" & """"; Trim(BinCount(12)); """" & " BinQuality=" & """" & "BIN12" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "D" & """" & " BinCount=" & """"; Trim(BinCount(13)); """" & " BinQuality=" & """" & "BIN13" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "E" & """" & " BinCount=" & """"; Trim(BinCount(14)); """" & " BinQuality=" & """" & "BIN14" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "F" & """" & " BinCount=" & """"; Trim(BinCount(15)); """" & " BinQuality=" & """" & "BIN15" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "G" & """" & " BinCount=" & """"; Trim(BinCount(16)); """" & " BinQuality=" & """" & "BIN16" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "H" & """" & " BinCount=" & """"; Trim(BinCount(17)); """" & " BinQuality=" & """" & "BIN17" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "I" & """" & " BinCount=" & """"; Trim(BinCount(18)); """" & " BinQuality=" & """" & "BIN18" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "J" & """" & " BinCount=" & """"; Trim(BinCount(19)); """" & " BinQuality=" & """" & "BIN19" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "K" & """" & " BinCount=" & """"; Trim(BinCount(20)); """" & " BinQuality=" & """" & "BIN20" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "L" & """" & " BinCount=" & """"; Trim(BinCount(21)); """" & " BinQuality=" & """" & "BIN21" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "M" & """" & " BinCount=" & """"; Trim(BinCount(22)); """" & " BinQuality=" & """" & "BIN22" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "N" & """" & " BinCount=" & """"; Trim(BinCount(23)); """" & " BinQuality=" & """" & "BIN23" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "O" & """" & " BinCount=" & """"; Trim(BinCount(24)); """" & " BinQuality=" & """" & "BIN24" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "P" & """" & " BinCount=" & """"; "0"; """" & " BinQuality=" & """" & "BIN25" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "Q" & """" & " BinCount=" & """"; "0"; """" & " BinQuality=" & """" & "BIN26" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "R" & """" & " BinCount=" & """"; "0"; """" & " BinQuality=" & """" & "BIN27" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "S" & """" & " BinCount=" & """"; "0"; """" & " BinQuality=" & """" & "BIN28" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "T" & """" & " BinCount=" & """"; "0"; """" & " BinQuality=" & """" & "BIN29" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "U" & """" & " BinCount=" & """"; "0"; """" & " BinQuality=" & """" & "BIN30" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "V" & """" & " BinCount=" & """"; "0"; """" & " BinQuality=" & """" & "BIN31" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "X" & """" & " BinCount=" & """"; "0"; """" & " BinQuality=" & """" & "Over the BIN31" & """" & "/>"
        Print #ifreefile, "    <Data>"
        
'        GOOD_COUNT_BACKUP = BinCount(20)            'pass bin fix
       
        '''''''''''''''''''''''''''''''''''''
        For y = 0 To StarProbe.ChipCountY
            For x = 0 To StarProbe.ChipCountX
                If WaferTemp(x, y).Chip Then
                    If WaferTemp(x, y).ChipMask Or WaferTemp(x, y).ChipPlate Or WaferTemp(x, y).ChipSkipDie Then        'ugly die
                        buf = buf & "-"
                    Else                                                                                                'normal die
                        Select Case WaferTemp(x, y).BIN
                            Case 0 To 9
                                TEMP_BIN = WaferTemp(x, y).BIN
                            Case 10
                                TEMP_BIN = "A"
                            Case 11
                                TEMP_BIN = "B"
                            Case 12
                                TEMP_BIN = "C"
                            Case 13
                                TEMP_BIN = "D"
                            Case 14
                                TEMP_BIN = "E"
                            Case 15
                                TEMP_BIN = "F"
                            Case 16
                                TEMP_BIN = "G"
                            Case 17
                                TEMP_BIN = "H"
                            Case 18
                                TEMP_BIN = "I"
                            Case 19
                                TEMP_BIN = "J"
                            Case 20
                                TEMP_BIN = "K"
                            Case 21
                                TEMP_BIN = "L"
                            Case 22
                                TEMP_BIN = "M"
                            Case 23
                                TEMP_BIN = "N"
                            Case 24
                                TEMP_BIN = "O"
                        End Select
                        buf = buf & TEMP_BIN             '불량 (bin no를 사용하도록 수정해야 한다.)
                    End If
                Else
                    buf = buf & "-"
                End If
            Next
            Print #ifreefile, "     <Row><![CDATA[" & buf & "]]></Row>"
            buf = ""
        Next
        Print #ifreefile, "    </Data>"
        Print #ifreefile, "    </Device>"
        Print #ifreefile, "  </Map>"
        Print #ifreefile, "  </Maps>"
    Close ifreefile
    
    FileCopy sfilename, R_File_Name & "ST1_" & UCase(MT2000.Text1(0)) & "_" & W_NO & ".wmd01"
    
    Exit Sub
    
err:
    MsgBox "wmd01파일 저장시 에러가 발생하였습니다."
End Sub

Sub Display_View5_1()             'skip die 포함
    Dim i As Integer, j As Integer, k As Integer
    Dim x As Integer, y As Integer
    Dim dx As Integer, dy As Integer
    Dim forx As Integer, fory As Integer
    Dim xfrom As Integer, xto As Integer
    Dim yfrom As Integer, yto As Integer
    Dim ifreefile As Integer
    Dim tempbinno As Byte
    Dim Count_Total As Long, Count_Good As Long
    Dim Count_Bad As Long, Count_Skip As Long
    Dim BinCount(0 To 26) As Long
    Dim bexit As Boolean
    Dim sfilename As String
    Dim buf As String
    Dim TEMP_BIN As String
    
    On Error GoTo err
    
    Erase WaferTemp
    
    For y = 0 To StarProbe.ChipCountY
        For x = 0 To StarProbe.ChipCountX
            WaferTemp(x, y) = Wafer(x, y)
        Next
    Next
    
    Count_Total = 0
    Count_Good = 0
    Count_Bad = 0
    Count_Skip = 0
    
    For i = 0 To 26
        BinCount(i) = 0
    Next
        
    For y = 0 To StarProbe.ChipCountY
        For x = 0 To StarProbe.ChipCountX
            Select Case XPitch(TT_NO)
                Case 2: xfrom = x:     xto = x + 1
                Case 3: xfrom = x - 1: xto = x + 1
                Case 4: xfrom = x - 1: xto = x + 2
                Case 5: xfrom = x - 2: xto = x + 2
                Case 6: xfrom = x - 2: xto = x + 3
                Case 7: xfrom = x - 3: xto = x + 3
                Case 8: xfrom = x - 3: xto = x + 4
                Case 9: xfrom = x - 4: xto = x + 4
                Case 10: xfrom = x - 4: xto = x + 5
                Case 11: xfrom = x - 5: xto = x + 5
                Case 12: xfrom = x - 5: xto = x + 6
                Case 13: xfrom = x - 6: xto = x + 6
                Case 14: xfrom = x - 6: xto = x + 7
                Case 15: xfrom = x - 7: xto = x + 7
                Case 16: xfrom = x - 7: xto = x + 8
                Case 17: xfrom = x - 8: xto = x + 8
                Case 18: xfrom = x - 8: xto = x + 9
                Case 19: xfrom = x - 9: xto = x + 9
                Case 20: xfrom = x - 9: xto = x + 10
                Case 21: xfrom = x - 10: xto = x + 10
                Case 22: xfrom = x - 10: xto = x + 11
                Case 23: xfrom = x - 11: xto = x + 11
                Case 24: xfrom = x - 11: xto = x + 12
                Case 25: xfrom = x - 12: xto = x + 12
                Case 26: xfrom = x - 12: xto = x + 13
                Case 27: xfrom = x - 13: xto = x + 13
                Case 28: xfrom = x - 13: xto = x + 14
                Case 29: xfrom = x - 14: xto = x + 14
                Case 30: xfrom = x - 14: xto = x + 15
            End Select
                                                 '4channel fix
            'yfrom = y - 1: yto = y + 2
            Select Case YPitch(TT_NO)
                Case 2: yfrom = y:     yto = y + 1
                Case 3: yfrom = y - 1: yto = y + 1
                Case 4: yfrom = y - 1: yto = y + 2
                Case 5: yfrom = y - 2: yto = y + 2
                Case 6: yfrom = y - 2: yto = y + 3
                Case 7: yfrom = y - 3: yto = y + 3
                Case 8: yfrom = y - 3: yto = y + 4
                Case 9: yfrom = y - 4: yto = y + 4
                Case 10: yfrom = y - 4: yto = y + 5
                Case 11: yfrom = y - 5: yto = y + 5
                Case 12: yfrom = y - 5: yto = y + 6
                Case 13: yfrom = y - 6: yto = y + 6
                Case 14: yfrom = y - 6: yto = y + 7
                Case 15: yfrom = y - 7: yto = y + 7
                Case 16: yfrom = y - 7: yto = y + 8
                Case 17: yfrom = y - 8: yto = y + 8
                Case 18: yfrom = y - 8: yto = y + 9
                Case 19: yfrom = y - 9: yto = y + 9
                Case 20: yfrom = y - 9: yto = y + 10
                Case 21: yfrom = y - 10: yto = y + 10
                Case 22: yfrom = y - 10: yto = y + 11
                Case 23: yfrom = y - 11: yto = y + 11
                Case 24: yfrom = y - 11: yto = y + 12
                Case 25: yfrom = y - 12: yto = y + 12
                Case 26: yfrom = y - 12: yto = y + 13
                Case 27: yfrom = y - 13: yto = y + 13
                Case 28: yfrom = y - 13: yto = y + 14
                Case 29: yfrom = y - 14: yto = y + 14
                Case 30: yfrom = y - 14: yto = y + 15
            End Select

            If WaferTemp(x, y).Chip And Not WaferTemp(x, y).ChipMask And Not WaferTemp(x, y).ChipSkipDie And Not WaferTemp(x, y).ChipPlate And WaferTemp(x, y).flag And WaferTemp(x, y).ChipMeasure Then
                If WaferTemp(x, y).FlagBad Then
                    bexit = False
                    For fory = y To yfrom Step -1
                        For forx = x To xfrom Step -1
                            If WaferTemp(forx, fory).Chip And Not WaferTemp(forx, fory).ChipMask And Not WaferTemp(forx, fory).ChipSkipDie And Not WaferTemp(forx, fory).ChipPlate And WaferTemp(forx, fory).flag And Not WaferTemp(forx, fory).FlagBad Then
                                tempbinno = WaferTemp(forx, fory).BIN
                                bexit = True
                                Exit For
                            End If
                        Next

                        If bexit Then Exit For

                        For forx = x To xto
                            If WaferTemp(forx, fory).Chip And Not WaferTemp(forx, fory).ChipMask And Not WaferTemp(forx, fory).ChipSkipDie And Not WaferTemp(forx, fory).ChipPlate And WaferTemp(forx, fory).flag And Not WaferTemp(forx, fory).FlagBad Then
                                tempbinno = WaferTemp(forx, fory).BIN
                                bexit = True
                                Exit For
                            End If
                        Next
                        If bexit Then Exit For
                    Next

                    For fory = y To yto
                        If bexit Then Exit For
                        For forx = x To xfrom Step -1
                            If WaferTemp(forx, fory).Chip And Not WaferTemp(forx, fory).ChipMask And Not WaferTemp(forx, fory).ChipSkipDie And Not WaferTemp(forx, fory).ChipPlate And WaferTemp(forx, fory).flag And Not WaferTemp(forx, fory).FlagBad Then
                                tempbinno = WaferTemp(forx, fory).BIN
                                bexit = True
                                Exit For
                            End If
                        Next

                        If bexit Then Exit For

                        For forx = x To xto
                            If WaferTemp(forx, fory).Chip And Not WaferTemp(forx, fory).ChipMask And Not WaferTemp(forx, fory).ChipSkipDie And Not WaferTemp(forx, fory).ChipPlate And WaferTemp(forx, fory).flag And Not WaferTemp(forx, fory).FlagBad Then
                                tempbinno = WaferTemp(forx, fory).BIN
                                bexit = True
                                Exit For
                            End If
                        Next
                        If bexit Then Exit For
                    Next
                Else
                    tempbinno = WaferTemp(x, y).BIN
                End If

                For forx = xfrom To xto
                    For fory = yfrom To yto
                        If WaferTemp(forx, fory).Chip And Not WaferTemp(forx, fory).ChipMask And Not WaferTemp(forx, fory).ChipSkipDie And Not WaferTemp(forx, fory).ChipMeasure And Not WaferTemp(forx, fory).MeasureWait And Not WaferTemp(forx, fory).ChipPlate And Not WaferTemp(forx, fory).flag And Not WaferTemp(forx, fory).FlagBad And Not ((x = forx) And (y = fory)) Then
                            WaferTemp(forx, fory).flag = True
                            WaferTemp(forx, fory).BIN = tempbinno
                        End If
                    Next
                Next
            End If
        Next
    Next

    For y = 0 To StarProbe.ChipCountY
        For x = 0 To StarProbe.ChipCountX
            If WaferTemp(x, y).Chip And Not WaferTemp(x, y).ChipMask And Not WaferTemp(x, y).ChipSkipDie And Not WaferTemp(x, y).MeasureWait And Not WaferTemp(x, y).ChipPlate And Not WaferTemp(x, y).flag And Not WaferTemp(x, y).FlagBad Then
                If WaferTemp(x, y - 1).flag And Not WaferTemp(x, y - 1).FlagBad Then
                    tempbinno = WaferTemp(x, y - 1).BIN
                ElseIf WaferTemp(x + 1, y).flag And Not WaferTemp(x + 1, y).FlagBad Then
                    tempbinno = WaferTemp(x + 1, y).BIN
                ElseIf WaferTemp(x, y + 1).flag And Not WaferTemp(x, y + 1).FlagBad Then
                    tempbinno = WaferTemp(x, y + 1).BIN
                ElseIf WaferTemp(x - 1, y).flag And Not WaferTemp(x - 1, y).FlagBad Then
                    tempbinno = WaferTemp(x - 1, y).BIN
                End If
                WaferTemp(x, y).flag = True
                WaferTemp(x, y).BIN = tempbinno
            End If
        Next
    Next

    For y = 0 To StarProbe.ChipCountY
        For x = 0 To StarProbe.ChipCountX
            If WaferTemp(x, y).Chip Then
                Count_Total = Count_Total + 1
                If WaferTemp(x, y).ChipMask Or WaferTemp(x, y).ChipSkipDie Or WaferTemp(x, y).ChipPlate Then
                    Count_Skip = Count_Skip + 1
                Else 'If WaferTemp(x, y).flag Then
                    If WaferTemp(x, y).FlagBad Then
                        Count_Bad = Count_Bad + 1
                    Else
                        Count_Good = Count_Good + 1
'                        WaferTemp(x, y).BIN = GOOD_BIN_NO
                    End If
                    BinCount(WaferTemp(x, y).BIN) = BinCount(WaferTemp(x, y).BIN) + 1
                End If
            End If
        Next
    Next
        
        
    'map
    If UCase(Right(MT2000.SSPanel2(0), 3)) = ".RP" Then      '[ 2020.06.23 ] : double check
        sfilename = FILE_NAMEING & "(MAP2).xml"
    Else
        sfilename = FILE_NAMEING & "(MAP2).xml"
    End If
    ifreefile = FreeFile
            
    Open sfilename For Output As ifreefile
        '파일 내용을 XML형식으로 저장한다.
        Print #ifreefile, "" 'Format(Now, "AMPM hh:mm YYYY-MM-DD")           '07.11.21날짜변경시 오류발견 수정.
        Print #ifreefile, "  <?xml version=" & """1.0""" & " encoding=" & """utf-8""" & "?>"
        Print #ifreefile, "  <Maps>"
        Print #ifreefile, "  <Map xmlns=" & """http://www.semi.org""" & _
                        " SubstrateId=" & """"; LOT & "_" & MT2000.lblWafer; """" & _
                        " SubstrateType=" & """Wafer""" & _
                        " FormatRevision=" & """SEMI G85-0703""" & ">"
        Print #ifreefile, "    <Device Rows=" & """"; Trim(StarProbe.ChipCountY); """" & _
                             " LotId=" & """"; LOT; """" & _
                             " BinType=" & """" & "ASCII" & """" & _
                             " Columns=" & """"; Trim(StarProbe.ChipCountX); """" & _
                             " MapType=" & """" & "Array" & """" & _
                             " NullBin=" & """" & "0" & """" & _
                             " ProductId=" & """"; Trim(DEV); """" & _
                             " WaferSize=" & """"; Trim(StarProbe.WaferSizemm); """" & _
                             " CreateDate=" & """"; Format(Now, "YYYYMMDDhhmmssms") & """" & _
                             " DeviceSizeX=" & """"; Format(Trim(StarProbe.ChipSizeX * 1000), "0000.00"); """" & _
                             " DeviceSizeY=" & """"; Format(Trim(StarProbe.ChipSizeY * 1000), "0000.00"); """" & _
                             " LotNo=" & """"; Trim(LOT); """" & _
                             " Orientation=" & """" & "0" & """" & _
                             " SupplierName=" & """" & "" & """" & _
                             " Originlocation=" & """" & "2" & """" & ">"
        Print #ifreefile, "    <ReferenceDevice ReferenceDeviceX=" & """" & O_X & """" & " " & "Referencedevice=" & """" & O_Y & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "0" & """" & " BinCount=" & """"; Trim(BinCount(0)); """" & " BinQuality=" & """" & "BIN0" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "1" & """" & " BinCount=" & """"; Trim(BinCount(1)); """" & " BinQuality=" & """" & "BIN1" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "2" & """" & " BinCount=" & """"; Trim(BinCount(2)); """" & " BinQuality=" & """" & "BIN2" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "3" & """" & " BinCount=" & """"; Trim(BinCount(3)); """" & " BinQuality=" & """" & "BIN3" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "4" & """" & " BinCount=" & """"; Trim(BinCount(4)); """" & " BinQuality=" & """" & "BIN4" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "5" & """" & " BinCount=" & """"; Trim(BinCount(5)); """" & " BinQuality=" & """" & "BIN5" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "6" & """" & " BinCount=" & """"; Trim(BinCount(6)); """" & " BinQuality=" & """" & "BIN6" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "7" & """" & " BinCount=" & """"; Trim(BinCount(7)); """" & " BinQuality=" & """" & "BIN7" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "8" & """" & " BinCount=" & """"; Trim(BinCount(8)); """" & " BinQuality=" & """" & "BIN8" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "9" & """" & " BinCount=" & """"; Trim(BinCount(9)); """" & " BinQuality=" & """" & "BIN9" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "A" & """" & " BinCount=" & """"; Trim(BinCount(10)); """" & " BinQuality=" & """" & "BIN10" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "B" & """" & " BinCount=" & """"; Trim(BinCount(11)); """" & " BinQuality=" & """" & "BIN11" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "C" & """" & " BinCount=" & """"; Trim(BinCount(12)); """" & " BinQuality=" & """" & "BIN12" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "D" & """" & " BinCount=" & """"; Trim(BinCount(13)); """" & " BinQuality=" & """" & "BIN13" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "E" & """" & " BinCount=" & """"; Trim(BinCount(14)); """" & " BinQuality=" & """" & "BIN14" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "F" & """" & " BinCount=" & """"; Trim(BinCount(15)); """" & " BinQuality=" & """" & "BIN15" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "G" & """" & " BinCount=" & """"; Trim(BinCount(16)); """" & " BinQuality=" & """" & "BIN16" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "H" & """" & " BinCount=" & """"; Trim(BinCount(17)); """" & " BinQuality=" & """" & "BIN17" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "I" & """" & " BinCount=" & """"; Trim(BinCount(18)); """" & " BinQuality=" & """" & "BIN18" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "J" & """" & " BinCount=" & """"; Trim(BinCount(19)); """" & " BinQuality=" & """" & "BIN19" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "K" & """" & " BinCount=" & """"; Trim(BinCount(20)); """" & " BinQuality=" & """" & "BIN20" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "L" & """" & " BinCount=" & """"; Trim(BinCount(21)); """" & " BinQuality=" & """" & "BIN21" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "M" & """" & " BinCount=" & """"; Trim(BinCount(22)); """" & " BinQuality=" & """" & "BIN22" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "N" & """" & " BinCount=" & """"; Trim(BinCount(23)); """" & " BinQuality=" & """" & "BIN23" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "O" & """" & " BinCount=" & """"; Trim(BinCount(24)); """" & " BinQuality=" & """" & "BIN24" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "P" & """" & " BinCount=" & """"; "0"; """" & " BinQuality=" & """" & "BIN25" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "Q" & """" & " BinCount=" & """"; "0"; """" & " BinQuality=" & """" & "BIN26" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "R" & """" & " BinCount=" & """"; "0"; """" & " BinQuality=" & """" & "BIN27" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "S" & """" & " BinCount=" & """"; "0"; """" & " BinQuality=" & """" & "BIN28" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "T" & """" & " BinCount=" & """"; "0"; """" & " BinQuality=" & """" & "BIN29" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "U" & """" & " BinCount=" & """"; "0"; """" & " BinQuality=" & """" & "BIN30" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "V" & """" & " BinCount=" & """"; "0"; """" & " BinQuality=" & """" & "BIN31" & """" & "/>"
        Print #ifreefile, "    <Bin BinCode=" & """" & "X" & """" & " BinCount=" & """"; "0"; """" & " BinQuality=" & """" & "Over the BIN31" & """" & "/>"
        Print #ifreefile, "    <Data>"
        
'        GOOD_COUNT_BACKUP = BinCount(20)            'pass bin fix
       
        '''''''''''''''''''''''''''''''''''''
        '측정한 값의 Y값을 표시해주는 부분
        For j = 0 To StarProbe.ChipCountY
            For k = 0 To StarProbe.ChipCountX
                If WaferTemp(k, j).Chip Then
                    If WaferTemp(k, j).ChipMask Or WaferTemp(k, j).ChipPlate Or WaferTemp(k, j).ChipSkipDie Then        'ugly die
                        TEMP_BIN = "."
                    Else
                        '=================================BIN NO 10이상인 경우는 영어로 처리하여 BIT수를 일치시킨다.
                        Select Case WaferTemp(k, j).BIN
                            Case 0 To 9
                                'TEMP_BIN = Wafer(k, j).BIN
                                TEMP_BIN = WaferTemp(k, j).BIN
                            Case 10
                                TEMP_BIN = "A"
                            Case 11
                                TEMP_BIN = "B"
                            Case 12
                                TEMP_BIN = "C"
                            Case 13
                                TEMP_BIN = "D"
                            Case 14
                                TEMP_BIN = "E"
                            Case 15
                                TEMP_BIN = "F"
                            Case 16
                                TEMP_BIN = "G"
                            Case 17
                                TEMP_BIN = "H"
                            Case 18
                                TEMP_BIN = "I"
                            Case 19
                                TEMP_BIN = "J"
                            Case 20
                                TEMP_BIN = "K"
                            Case 21
                                TEMP_BIN = "L"
                            Case 22
                                TEMP_BIN = "M"
                            Case 23
                                TEMP_BIN = "N"
                            Case 24
                                TEMP_BIN = "O"
                            Case Else
                                TEMP_BIN = "."
                        End Select
                    End If
                Else
                    TEMP_BIN = "."
                End If
'                If k = 0 Then
'                    If j >= 100 Then          '100이상
'                        buf = j & Space(1) & buf & TEMP_BIN
'                    ElseIf j >= 10 Then      '10이상
'                        buf = j & Space(2) & buf & TEMP_BIN
'                    Else
'                        buf = j & Space(3) & buf & TEMP_BIN
'                    End If
'                Else
                    buf = buf & TEMP_BIN
'                End If
            Next k
            Print #1, "<Row><![CDATA[" & buf & "]]></Row>"
            buf = ""
        Next j
        Print #ifreefile, "    </Data>"
        Print #ifreefile, "    </Device>"
        Print #ifreefile, "  </Map>"
        Print #ifreefile, "  </Maps>"
    Close ifreefile
    
    FileCopy sfilename, R_File_Name & LOT & "_" & W_NO & "(map2).xml"
    
    Exit Sub
    
err:
    MsgBox "wmd01파일 저장시 에러가 발생하였습니다."
End Sub


Sub Display_View_Change()             'skip die 포함
    Dim i As Integer, j As Integer, k As Integer
    Dim x As Integer, y As Integer
    Dim dx As Integer, dy As Integer
    Dim forx As Integer, fory As Integer
    Dim xfrom As Integer, xto As Integer
    Dim yfrom As Integer, yto As Integer
    Dim ifreefile As Integer
    Dim tempbinno As Byte
    Dim Count_Total As Long, Count_Good As Long
    Dim Count_Bad As Long, Count_Skip As Long
    Dim BinCount(0 To 26) As Long
    Dim bexit As Boolean
    Dim sfilename As String
    Dim buf As String
    Dim TEMP_BIN As String
    
    On Error GoTo err
    
    Erase WaferTemp
    
    For y = 0 To StarProbe.ChipCountY
        For x = 0 To StarProbe.ChipCountX
            WaferTemp(x, y) = Wafer(x, y)
        Next
    Next
    
    Count_Total = 0
    Count_Good = 0
    Count_Bad = 0
    Count_Skip = 0
    
    For i = 0 To 26
        BinCount(i) = 0
    Next
        
    For y = 0 To StarProbe.ChipCountY
        For x = 0 To StarProbe.ChipCountX
            Select Case XPitch(TT_NO)
                Case 2: xfrom = x:     xto = x + 1
                Case 3: xfrom = x - 1: xto = x + 1
                Case 4: xfrom = x - 1: xto = x + 2
                Case 5: xfrom = x - 2: xto = x + 2
                Case 6: xfrom = x - 2: xto = x + 3
                Case 7: xfrom = x - 3: xto = x + 3
                Case 8: xfrom = x - 3: xto = x + 4
                Case 9: xfrom = x - 4: xto = x + 4
                Case 10: xfrom = x - 4: xto = x + 5
                Case 11: xfrom = x - 5: xto = x + 5
                Case 12: xfrom = x - 5: xto = x + 6
                Case 13: xfrom = x - 6: xto = x + 6
                Case 14: xfrom = x - 6: xto = x + 7
                Case 15: xfrom = x - 7: xto = x + 7
                Case 16: xfrom = x - 7: xto = x + 8
                Case 17: xfrom = x - 8: xto = x + 8
                Case 18: xfrom = x - 8: xto = x + 9
                Case 19: xfrom = x - 9: xto = x + 9
                Case 20: xfrom = x - 9: xto = x + 10
                Case 21: xfrom = x - 10: xto = x + 10
                Case 22: xfrom = x - 10: xto = x + 11
                Case 23: xfrom = x - 11: xto = x + 11
                Case 24: xfrom = x - 11: xto = x + 12
                Case 25: xfrom = x - 12: xto = x + 12
                Case 26: xfrom = x - 12: xto = x + 13
                Case 27: xfrom = x - 13: xto = x + 13
                Case 28: xfrom = x - 13: xto = x + 14
                Case 29: xfrom = x - 14: xto = x + 14
                Case 30: xfrom = x - 14: xto = x + 15
            End Select
                                                 '4channel fix
            'yfrom = y - 1: yto = y + 2
            Select Case YPitch(TT_NO)
                Case 2: yfrom = y:     yto = y + 1
                Case 3: yfrom = y - 1: yto = y + 1
                Case 4: yfrom = y - 1: yto = y + 2
                Case 5: yfrom = y - 2: yto = y + 2
                Case 6: yfrom = y - 2: yto = y + 3
                Case 7: yfrom = y - 3: yto = y + 3
                Case 8: yfrom = y - 3: yto = y + 4
                Case 9: yfrom = y - 4: yto = y + 4
                Case 10: yfrom = y - 4: yto = y + 5
                Case 11: yfrom = y - 5: yto = y + 5
                Case 12: yfrom = y - 5: yto = y + 6
                Case 13: yfrom = y - 6: yto = y + 6
                Case 14: yfrom = y - 6: yto = y + 7
                Case 15: yfrom = y - 7: yto = y + 7
                Case 16: yfrom = y - 7: yto = y + 8
                Case 17: yfrom = y - 8: yto = y + 8
                Case 18: yfrom = y - 8: yto = y + 9
                Case 19: yfrom = y - 9: yto = y + 9
                Case 20: yfrom = y - 9: yto = y + 10
                Case 21: yfrom = y - 10: yto = y + 10
                Case 22: yfrom = y - 10: yto = y + 11
                Case 23: yfrom = y - 11: yto = y + 11
                Case 24: yfrom = y - 11: yto = y + 12
                Case 25: yfrom = y - 12: yto = y + 12
                Case 26: yfrom = y - 12: yto = y + 13
                Case 27: yfrom = y - 13: yto = y + 13
                Case 28: yfrom = y - 13: yto = y + 14
                Case 29: yfrom = y - 14: yto = y + 14
                Case 30: yfrom = y - 14: yto = y + 15
            End Select

            '정상칩 중에서 검색
            If WaferTemp(x, y).Chip And Not WaferTemp(x, y).ChipMask And Not WaferTemp(x, y).ChipSkipDie And Not WaferTemp(x, y).ChipPlate And WaferTemp(x, y).flag And WaferTemp(x, y).ChipMeasure Then
                If WaferTemp(x, y).FlagBad Then
                    bexit = False
                    For fory = y To yfrom Step -1
                        For forx = x To xfrom Step -1
                            If WaferTemp(forx, fory).Chip And Not WaferTemp(forx, fory).ChipMask And Not WaferTemp(forx, fory).ChipSkipDie And Not WaferTemp(forx, fory).ChipPlate And WaferTemp(forx, fory).flag And Not WaferTemp(forx, fory).FlagBad Then
                                tempbinno = WaferTemp(forx, fory).BIN
                                bexit = True
                                Exit For
                            End If
                        Next

                        If bexit Then Exit For

                        For forx = x To xto
                            If WaferTemp(forx, fory).Chip And Not WaferTemp(forx, fory).ChipMask And Not WaferTemp(forx, fory).ChipSkipDie And Not WaferTemp(forx, fory).ChipPlate And WaferTemp(forx, fory).flag And Not WaferTemp(forx, fory).FlagBad Then
                                tempbinno = WaferTemp(forx, fory).BIN
                                bexit = True
                                Exit For
                            End If
                        Next
                        If bexit Then Exit For
                    Next

                    For fory = y To yto
                        If bexit Then Exit For
                        For forx = x To xfrom Step -1
                            If WaferTemp(forx, fory).Chip And Not WaferTemp(forx, fory).ChipMask And Not WaferTemp(forx, fory).ChipSkipDie And Not WaferTemp(forx, fory).ChipPlate And WaferTemp(forx, fory).flag And Not WaferTemp(forx, fory).FlagBad Then
                                tempbinno = WaferTemp(forx, fory).BIN
                                bexit = True
                                Exit For
                            End If
                        Next

                        If bexit Then Exit For

                        For forx = x To xto
                            If WaferTemp(forx, fory).Chip And Not WaferTemp(forx, fory).ChipMask And Not WaferTemp(forx, fory).ChipSkipDie And Not WaferTemp(forx, fory).ChipPlate And WaferTemp(forx, fory).flag And Not WaferTemp(forx, fory).FlagBad Then
                                tempbinno = WaferTemp(forx, fory).BIN
                                bexit = True
                                Exit For
                            End If
                        Next
                        If bexit Then Exit For
                    Next
                Else
                    tempbinno = WaferTemp(x, y).BIN
                End If

                For forx = xfrom To xto
                    For fory = yfrom To yto
                        If WaferTemp(forx, fory).Chip And Not WaferTemp(forx, fory).ChipMask And Not WaferTemp(forx, fory).ChipSkipDie And Not WaferTemp(forx, fory).ChipMeasure And Not WaferTemp(forx, fory).MeasureWait And Not WaferTemp(forx, fory).ChipPlate And Not WaferTemp(forx, fory).flag And Not WaferTemp(forx, fory).FlagBad And Not ((x = forx) And (y = fory)) Then
                            WaferTemp(forx, fory).flag = True
                            WaferTemp(forx, fory).BIN = tempbinno
                        End If
                    Next
                Next
            End If
        Next
    Next            '양품 bin을 구해서 불량 이외의 부분을 양품으로 바꾼다.

    For y = 0 To StarProbe.ChipCountY
        For x = 0 To StarProbe.ChipCountX
            If WaferTemp(x, y).Chip And Not WaferTemp(x, y).ChipMask And Not WaferTemp(x, y).ChipSkipDie And Not WaferTemp(x, y).MeasureWait And Not WaferTemp(x, y).ChipPlate And Not WaferTemp(x, y).flag And Not WaferTemp(x, y).FlagBad Then
                If WaferTemp(x, y - 1).flag And Not WaferTemp(x, y - 1).FlagBad Then
                    tempbinno = WaferTemp(x, y - 1).BIN
                ElseIf WaferTemp(x + 1, y).flag And Not WaferTemp(x + 1, y).FlagBad Then
                    tempbinno = WaferTemp(x + 1, y).BIN
                ElseIf WaferTemp(x, y + 1).flag And Not WaferTemp(x, y + 1).FlagBad Then
                    tempbinno = WaferTemp(x, y + 1).BIN
                ElseIf WaferTemp(x - 1, y).flag And Not WaferTemp(x - 1, y).FlagBad Then
                    tempbinno = WaferTemp(x - 1, y).BIN
                End If
                WaferTemp(x, y).flag = True
                WaferTemp(x, y).BIN = tempbinno
            End If
        Next
    Next

    For y = 0 To StarProbe.ChipCountY
        For x = 0 To StarProbe.ChipCountX
            If WaferTemp(x, y).Chip Then
                Count_Total = Count_Total + 1
                If WaferTemp(x, y).ChipMask Or WaferTemp(x, y).ChipSkipDie Or WaferTemp(x, y).ChipPlate Then
                    Count_Skip = Count_Skip + 1
                Else 'If WaferTemp(x, y).flag Then
                    If WaferTemp(x, y).FlagBad Then
                        Count_Bad = Count_Bad + 1
                    Else
                        Count_Good = Count_Good + 1
'                        WaferTemp(x, y).BIN = GOOD_BIN_NO
                    End If
                    BinCount(WaferTemp(x, y).BIN) = BinCount(WaferTemp(x, y).BIN) + 1
                End If
            End If
        Next
    Next
                
    'map
    sfilename = FILE_NAMEING & ".txt"
    ifreefile = FreeFile
            
    '[ 2021.05.26 ] : txt 파일 저장시 form 수정
    Dim strtmp_X As String                              'X값 임시 저장
    Dim strtmp_Y As String                              'Y값 임시 저장
    Dim GAP As Integer
    
    GAP = 39                                            '표시 영역 자릿수
    
    Open sfilename For Output As ifreefile
                                  '1234567890123456
        Print #ifreefile, STR_FIX("START_TIME    : " & STT_time, GAP) & _
                          STR_FIX("END_TIME      : " & END_time, GAP)
        Print #ifreefile, STR_FIX("TEST_TIME     : " & StarProbe_WorkDateTime.h & ":" & StarProbe_WorkDateTime.M & ":" & StarProbe_WorkDateTime.s, GAP) & _
                          STR_FIX("WAFER_SIZE    : " & StarProbe.WaferSizemm & " mm", GAP)
        Print #ifreefile, STR_FIX("X_SIZE        : " & StarProbe.ChipSizeX, GAP) & _
                          STR_FIX("Y_SIZE        : " & StarProbe.ChipSizeY, GAP)
        Print #ifreefile, STR_FIX("LOT NO        : " & LOT, GAP) & _
                          STR_FIX("WAFER NO      : " & W_NO, GAP)
        Print #ifreefile, STR_FIX("WAFER ID      : " & LOT & W_NO, GAP) & _
                          STR_FIX("TOTAL_DIE     : " & StarProbe.CountGoodDie + StarProbe.CountBadDie, GAP)
        Print #ifreefile, STR_FIX("Bin die(PASS) : " & StarProbe.CountGoodDie, GAP) & _
                          STR_FIX("Fail die      : " & StarProbe.CountBadDie, GAP)
        'Flat zone 방향
        If MT2000.Option2(1).value = True Then             '상
            Print #ifreefile, STR_FIX("FLAT ZONE     : " & "UP", GAP)
        ElseIf MT2000.Option2(2).value = True Then         '하
            Print #ifreefile, STR_FIX("FLAT ZONE     : " & "DOWN", GAP)
        ElseIf MT2000.Option2(3).value = True Then         '좌
            Print #ifreefile, STR_FIX("FLAT ZONE     : " & "Left", GAP)
        Else                                            '우
            Print #ifreefile, STR_FIX("FLAT ZONE     : " & "Right", GAP)
        End If
                
        If StarProbe.ChipCountX < 100 Then
            strtmp_X = "0" & StarProbe.ChipCountX           'chip count가 100보다 작은 경우 099로 표시한다.
        Else
            strtmp_X = StarProbe.ChipCountX                 'chip count가 100보다 작은 경우 099로 표시한다.
        End If
        
        If StarProbe.ChipCountY < 100 Then
            strtmp_Y = "0" & StarProbe.ChipCountY           'chip count가 100보다 작은 경우 099로 표시한다.
        Else
            strtmp_Y = StarProbe.ChipCountY                 'chip count가 100보다 작은 경우 099로 표시한다.
        End If
        Print #ifreefile, STR_FIX("[X:000~" & strtmp_X & "]" & "   [Y:000~" & strtmp_Y & "]", GAP)
                
        'GOOD_COUNT_BACKUP = BinCount(20)                    'pass bin fix
       
        '''''''''''''''''''''''''''''''''''''
        'chip이 아닌 경우 "."표시하고 mask,skip,plate도 "."표시, Start chip은 "R"로 표시.
        For y = 0 To StarProbe.ChipCountY
            For x = 0 To StarProbe.ChipCountX
                If WaferTemp(x, y).Chip Then
                    If WaferTemp(x, y).ChipMask Or WaferTemp(x, y).ChipPlate Or WaferTemp(x, y).ChipSkipDie Then        'ugly die
                        If x = StarProbe.StartChip.x And y = StarProbe.StartChip.y Then
                            buf = buf & "R"             'first die를 reference die로 표시한다.
                        Else
                            buf = buf & "."
                        End If
                    Else                                                                                                'normal die
                        If x = StarProbe.StartChip.x And y = StarProbe.StartChip.y Then
                            TEMP_BIN = "R"              'first die를 reference die로 표시한다.
                        'ElseIf WaferTemp(x, y).BIN >= 11 And WaferTemp(x, y).BIN <> 18 Then                             '[ 2021.06.08 ] : bin이 11보다 크면서 bin18이 아닌 경우 양품
                        ElseIf WaferTemp(x, y).BIN = GOOD_BIN_NO Then                             '[ 2021.06.08 ] : bin이 11보다 크면서 bin18이 아닌 경우 양품
                            TEMP_BIN = "1"              '양품은 bin1로 변경 한다.
                        Else
                            TEMP_BIN = "X"              '불량은 X로 변경 한다.
                        End If
                        buf = buf & TEMP_BIN
                    End If
                Else
                    buf = buf & "."
                End If
            Next
            Print #ifreefile, buf
            buf = ""
        Next
    Close ifreefile
    
    FileCopy sfilename, R_File_Name & UCase(MT2000.Text1(0)) & "_" & W_NO & ".txt"
    Exit Sub
    
err:
    MsgBox "txt파일 저장시 에러가 발생하였습니다."
End Sub

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

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub HScroll_Measure_Change()
    pMeasure.Left = -(((pMeasure.width - 11390) / 1000) * HScroll_Measure.value)
    HScroll_Zoom.value = HScroll_Measure.value
End Sub

Private Sub HScroll_Measure_Scroll()
    Call HScroll_Measure_Change
End Sub

Private Sub HScroll_Zoom_Change()
    pView.Left = -(((pView.width - 11390) / 1000) * HScroll_Zoom.value)
    HScroll_Measure.value = HScroll_Zoom.value
End Sub

Private Sub HScroll_Zoom_Scroll()
    Call HScroll_Zoom_Change
End Sub

Private Sub VScroll_Measure_Change()
    pMeasure.Top = -(((pMeasure.Height - 7095) / 1000) * VScroll_Measure.value)
    VScroll_Zoom.value = VScroll_Measure.value
End Sub

Private Sub VScroll_Measure_Scroll()
    Call VScroll_Measure_Change
End Sub

Private Sub VScroll_Zoom_Change()
    pView.Top = -(((pView.Height - 7095) / 1000) * VScroll_Zoom.value)
    VScroll_Measure.value = VScroll_Zoom.value
End Sub

Private Sub VScroll_Zoom_Scroll()
    Call VScroll_Zoom_Change
End Sub

Sub stop_control()
    Unload Me
End Sub

Sub Display_View()
'    On Error GoTo ErrorSub
On Error GoTo err

    Dim i As Integer, j As Integer
    Dim x As Integer, y As Integer
    Dim dx As Integer, dy As Integer
    Dim lColor As Long
    Dim DisplayX As Integer, DisplayY As Integer
    Dim forx As Integer, fory As Integer
    Dim tempbinno As Byte
    Dim xfrom As Integer, xto As Integer
    Dim yfrom As Integer, yto As Integer
    Dim Count_Total As Long, Count_Good As Long
    Dim Count_Bad As Long, Count_Skip As Long
    Dim Yield1 As Double, Yield2 As Double
    Dim BinCount(0 To 26) As Long
    Dim bexit As Boolean
    Dim ifreefile As Integer
    Dim sLine As String
    Dim sfilename As String
    Dim shift_value As Integer              '2016.05.27(그림크기관련 수정)
    
    shift_value = 50                       '2016.05.27
    
    ''''''''''''''''''''
    pView.ScaleWidth = 1125
    pView.ScaleHeight = 873

    If StarProbe.ChipSizeX < 0.18 Then
        pView.ScaleWidth = 1125 + shift_value
        pView.ScaleHeight = 873 + shift_value
    End If
    ''''''''''''''''''''
    
    sfilename = UCase(BMP_file)
    sfilename = Mid(sfilename, 1, Len(sfilename) - 4) & ".CSV"
    ifreefile = FreeFile
    
    Open sfilename For Output As ifreefile
    
    Erase WaferTemp
    
    For y = 0 To StarProbe.ChipCountY
        For x = 0 To StarProbe.ChipCountX
            WaferTemp(x, y) = Wafer(x, y)
        Next
    Next
    
    Count_Total = 0
    Count_Good = 0
    Count_Bad = 0
    Count_Skip = 0
    
    For i = 0 To 26
        BinCount(i) = 0
    Next
    
    For y = 0 To StarProbe.ChipCountY
        For x = 0 To StarProbe.ChipCountX
            Select Case XPitch(TT_NO)
                Case 2: xfrom = x:     xto = x + 1
                Case 3: xfrom = x - 1: xto = x + 1
                Case 4: xfrom = x - 1: xto = x + 2
                Case 5: xfrom = x - 2: xto = x + 2
                Case 6: xfrom = x - 2: xto = x + 3
                Case 7: xfrom = x - 3: xto = x + 3
                Case 8: xfrom = x - 3: xto = x + 4
                Case 9: xfrom = x - 4: xto = x + 4
                Case 10: xfrom = x - 4: xto = x + 5
                Case 11: xfrom = x - 5: xto = x + 5
                Case 12: xfrom = x - 5: xto = x + 6
                Case 13: xfrom = x - 6: xto = x + 6
                Case 14: xfrom = x - 6: xto = x + 7
                Case 15: xfrom = x - 7: xto = x + 7
                Case 16: xfrom = x - 7: xto = x + 8
                Case 17: xfrom = x - 8: xto = x + 8
                Case 18: xfrom = x - 8: xto = x + 9
                Case 19: xfrom = x - 9: xto = x + 9
                Case 20: xfrom = x - 9: xto = x + 10
                Case 21: xfrom = x - 10: xto = x + 10
                Case 22: xfrom = x - 10: xto = x + 11
                Case 23: xfrom = x - 11: xto = x + 11
                Case 24: xfrom = x - 11: xto = x + 12
                Case 25: xfrom = x - 12: xto = x + 12
                Case 26: xfrom = x - 12: xto = x + 13
                Case 27: xfrom = x - 13: xto = x + 13
                Case 28: xfrom = x - 13: xto = x + 14
                Case 29: xfrom = x - 14: xto = x + 14
                Case 30: xfrom = x - 14: xto = x + 15
            End Select

'            yfrom = y - 1: yto = y + 2

            Select Case YPitch(TT_NO)
                Case 2: yfrom = y:     yto = y + 1
                Case 3: yfrom = y - 1: yto = y + 1
                Case 4: yfrom = y - 1: yto = y + 2
                Case 5: yfrom = y - 2: yto = y + 2
                Case 6: yfrom = y - 2: yto = y + 3
                Case 7: yfrom = y - 3: yto = y + 3
                Case 8: yfrom = y - 3: yto = y + 4
                Case 9: yfrom = y - 4: yto = y + 4
                Case 10: yfrom = y - 4: yto = y + 5
                Case 11: yfrom = y - 5: yto = y + 5
                Case 12: yfrom = y - 5: yto = y + 6
                Case 13: yfrom = y - 6: yto = y + 6
                Case 14: yfrom = y - 6: yto = y + 7
                Case 15: yfrom = y - 7: yto = y + 7
                Case 16: yfrom = y - 7: yto = y + 8
                Case 17: yfrom = y - 8: yto = y + 8
                Case 18: yfrom = y - 8: yto = y + 9
                Case 19: yfrom = y - 9: yto = y + 9
                Case 20: yfrom = y - 9: yto = y + 10
                Case 21: yfrom = y - 10: yto = y + 10
                Case 22: yfrom = y - 10: yto = y + 11
                Case 23: yfrom = y - 11: yto = y + 11
                Case 24: yfrom = y - 11: yto = y + 12
                Case 25: yfrom = y - 12: yto = y + 12
                Case 26: yfrom = y - 12: yto = y + 13
                Case 27: yfrom = y - 13: yto = y + 13
                Case 28: yfrom = y - 13: yto = y + 14
                Case 29: yfrom = y - 14: yto = y + 14
                Case 30: yfrom = y - 14: yto = y + 15
            End Select

            If WaferTemp(x, y).Chip And Not WaferTemp(x, y).ChipMask And Not WaferTemp(x, y).ChipSkipDie And Not WaferTemp(x, y).ChipPlate And WaferTemp(x, y).flag And WaferTemp(x, y).ChipMeasure Then
                If WaferTemp(x, y).FlagBad Then         'NG인 경우
                    bexit = False

                    For fory = y To yfrom Step -1
                        For forx = x To xfrom Step -1
                            If WaferTemp(forx, fory).Chip And Not WaferTemp(forx, fory).ChipMask And Not WaferTemp(forx, fory).ChipSkipDie And Not WaferTemp(forx, fory).ChipPlate And WaferTemp(forx, fory).flag And Not WaferTemp(forx, fory).FlagBad Then
                                tempbinno = WaferTemp(forx, fory).BIN
                                bexit = True
                                Exit For
                            End If
                        Next
                        If bexit Then Exit For

                        For forx = x To xto
                            If WaferTemp(forx, fory).Chip And Not WaferTemp(forx, fory).ChipMask And Not WaferTemp(forx, fory).ChipSkipDie And Not WaferTemp(forx, fory).ChipPlate And WaferTemp(forx, fory).flag And Not WaferTemp(forx, fory).FlagBad Then
                                tempbinno = WaferTemp(forx, fory).BIN
                                bexit = True
                                Exit For
                            End If
                        Next
                        If bexit Then Exit For
                    Next

                    For fory = y To yto
                        If bexit Then Exit For
                        For forx = x To xfrom Step -1
                            If WaferTemp(forx, fory).Chip And Not WaferTemp(forx, fory).ChipMask And Not WaferTemp(forx, fory).ChipSkipDie And Not WaferTemp(forx, fory).ChipPlate And WaferTemp(forx, fory).flag And Not WaferTemp(forx, fory).FlagBad Then
                                tempbinno = WaferTemp(forx, fory).BIN
                                bexit = True
                                Exit For
                            End If
                        Next
                        If bexit Then Exit For

                        For forx = x To xto
                            If WaferTemp(forx, fory).Chip And Not WaferTemp(forx, fory).ChipMask And Not WaferTemp(forx, fory).ChipSkipDie And Not WaferTemp(forx, fory).ChipPlate And WaferTemp(forx, fory).flag And Not WaferTemp(forx, fory).FlagBad Then
                                tempbinno = WaferTemp(forx, fory).BIN
                                bexit = True
                                Exit For
                            End If
                        Next
                        If bexit Then Exit For
                    Next
                Else
                    tempbinno = WaferTemp(x, y).BIN
                End If

                For forx = xfrom To xto
                    For fory = yfrom To yto
                        If WaferTemp(forx, fory).Chip And _
                           Not WaferTemp(forx, fory).ChipMask And _
                           Not WaferTemp(forx, fory).ChipSkipDie And _
                           Not WaferTemp(forx, fory).ChipMeasure And _
                           Not WaferTemp(forx, fory).MeasureWait And _
                           Not WaferTemp(forx, fory).ChipPlate And _
                           Not WaferTemp(forx, fory).flag And _
                           Not WaferTemp(forx, fory).FlagBad And _
                           Not ((x = forx) And (y = fory)) Then

                            WaferTemp(forx, fory).flag = True
                            WaferTemp(forx, fory).BIN = tempbinno
                        End If
                    Next
                Next
            End If
        Next
    Next

    For y = 0 To StarProbe.ChipCountY
        For x = 0 To StarProbe.ChipCountX
            If WaferTemp(x, y).Chip And _
               Not WaferTemp(x, y).ChipMask And _
               Not WaferTemp(x, y).ChipSkipDie And _
               Not WaferTemp(x, y).MeasureWait And _
               Not WaferTemp(x, y).ChipPlate And _
               Not WaferTemp(x, y).flag And _
               Not WaferTemp(x, y).FlagBad Then

                If WaferTemp(x, y - 1).flag And Not WaferTemp(x, y - 1).FlagBad Then
                    tempbinno = WaferTemp(x, y - 1).BIN
                ElseIf WaferTemp(x + 1, y).flag And Not WaferTemp(x + 1, y).FlagBad Then
                    tempbinno = WaferTemp(x + 1, y).BIN
                ElseIf WaferTemp(x, y + 1).flag And Not WaferTemp(x, y + 1).FlagBad Then
                    tempbinno = WaferTemp(x, y + 1).BIN
                ElseIf WaferTemp(x - 1, y).flag And Not WaferTemp(x - 1, y).FlagBad Then
                    tempbinno = WaferTemp(x - 1, y).BIN
                End If
                WaferTemp(x, y).flag = True
                WaferTemp(x, y).BIN = tempbinno
            End If
        Next
    Next

    For y = 0 To StarProbe.ChipCountY
        For x = 0 To StarProbe.ChipCountX
            If WaferTemp(x, y).Chip Then
                Count_Total = Count_Total + 1
                If WaferTemp(x, y).ChipMask Or _
                   WaferTemp(x, y).ChipSkipDie Or _
                   WaferTemp(x, y).ChipPlate Then

                    Count_Skip = Count_Skip + 1
                Else 'If WaferTemp(x, y).flag Then
                    If WaferTemp(x, y).FlagBad Then
                        Count_Bad = Count_Bad + 1
                    Else
                        Count_Good = Count_Good + 1
'                        WaferTemp(x, y).BIN = GOOD_BIN_NO
                    End If
                    BinCount(WaferTemp(x, y).BIN) = BinCount(WaferTemp(x, y).BIN) + 1
                End If
            End If
        Next
    Next
        
    pView.Cls

    pView.Font = "Arial"

    Print #ifreefile, "Star Probe Wafer Summary Report"
    pView.CurrentX = 75
    pView.CurrentY = 3
    pView.FontSize = 14
    pView.FontBold = True
    pView.ForeColor = vbRed
    pView.Print "Star Probe Wafer Summary Report"
    pView.FontBold = False
        
    pView.CurrentX = 720
    If StarProbe.ChipSizeX < 0.18 Then pView.CurrentX = 720 + shift_value
    pView.CurrentY = 13
    pView.FontSize = 8
    pView.FontBold = True
    pView.ForeColor = vbBlue
    pView.Print "Copyright (c) 2015.10 by KEC CO., LTD. All rights reserved."
    pView.FontBold = False
    
    pView.CurrentX = 846
    If StarProbe.ChipSizeX < 0.18 Then pView.CurrentX = 846 + shift_value
    pView.CurrentY = 13
    pView.FontSize = 8
    pView.FontBold = True
    pView.ForeColor = vbBlack

    pView.FontBold = False
    
    If StarProbe.ChipSizeX < 0.18 Then
        pView.Line (0, 30)-(1124 + shift_value, 31), vbBlack, BF
    Else
        pView.Line (0, 30)-(1124, 31), vbBlack, BF
    End If

    Print #ifreefile, ""
    Print #ifreefile, "Work Information"
    pView.CurrentX = 10
    pView.CurrentY = 35
    pView.FontSize = 12
    pView.FontBold = True
    pView.ForeColor = vbBlack
    pView.Print "Work Information"
    pView.FontBold = False

    pView.FontSize = 10
    pView.FontBold = False
    pView.ForeColor = vbBlack
    pView.CurrentX = 10: pView.CurrentY = 60:  pView.Print "LOT NO."
    pView.CurrentX = 10: pView.CurrentY = 75:  pView.Print "ITEM"
    pView.CurrentX = 10: pView.CurrentY = 90:  pView.Print "TYPE"
    pView.CurrentX = 10: pView.CurrentY = 105: pView.Print "MACHINE NO."
    pView.CurrentX = 10: pView.CurrentY = 120: pView.Print "OPERATER NO."
    pView.FontBold = False

    pView.FontSize = 10
    pView.FontBold = False
    pView.ForeColor = vbBlack
    pView.CurrentX = 110: pView.CurrentY = 60:  pView.Print ":"
    pView.CurrentX = 110: pView.CurrentY = 75:  pView.Print ":"
    pView.CurrentX = 110: pView.CurrentY = 90:  pView.Print ":"
    pView.CurrentX = 110: pView.CurrentY = 105: pView.Print ":"
    pView.CurrentX = 110: pView.CurrentY = 120: pView.Print ":"
    pView.FontBold = False

    pView.FontSize = 10
    pView.FontBold = True
    pView.ForeColor = vbBlack

    pView.CurrentX = 120: pView.CurrentY = 60:  pView.Print MT2000.Text1(0)
    pView.CurrentX = 120: pView.CurrentY = 75:  pView.Print MT2000.Text1(1)
    pView.CurrentX = 120: pView.CurrentY = 90:  pView.Print MT2000.Text1(2)
    pView.CurrentX = 120: pView.CurrentY = 105: pView.Print MT2000.Text1(3)
    pView.CurrentX = 120: pView.CurrentY = 120: pView.Print MT2000.Text1(4)
    
    Print #ifreefile, "LOT NO.," & MT2000.Text1(0)
    Print #ifreefile, "ITEM," & MT2000.Text1(1)
    Print #ifreefile, "TYPE," & MT2000.Text1(2)
    Print #ifreefile, "MACHINE NO.," & MT2000.Text1(3)
    Print #ifreefile, "OPERATER NO.," & MT2000.Text1(4)
    Print #ifreefile, "DATE," & Format(Date$, "YYYY.MM.DD ")
    Print #ifreefile, "TIME," & Format(Time$, "HH:MM:SS")

    pView.FontBold = False

    Print #ifreefile, ""
    Print #ifreefile, "Wafer Information"
    
    pView.CurrentX = 350
    pView.CurrentY = 35
    pView.FontSize = 12
    pView.FontBold = True
    pView.ForeColor = vbBlack
    pView.Print "Wafer Information"
    pView.FontBold = False

    pView.FontSize = 10
    pView.FontBold = False
    pView.ForeColor = vbBlack
    pView.CurrentX = 350: pView.CurrentY = 60:  pView.Print "Wafer Size"
    pView.CurrentX = 350: pView.CurrentY = 75:  pView.Print "Chip Size (W x H)"
    pView.CurrentX = 350: pView.CurrentY = 90:  pView.Print "Chip Count (X x Y)"
    pView.CurrentX = 350: pView.CurrentY = 105: pView.Print "Pitch (X x Y)"
    pView.CurrentX = 350: pView.CurrentY = 120: pView.Print "MAP"
    pView.FontBold = False

    pView.FontSize = 10
    pView.FontBold = False
    pView.ForeColor = vbBlack
    pView.CurrentX = 460: pView.CurrentY = 60:  pView.Print ":"
    pView.CurrentX = 460: pView.CurrentY = 75:  pView.Print ":"
    pView.CurrentX = 460: pView.CurrentY = 90:  pView.Print ":"
    pView.CurrentX = 460: pView.CurrentY = 105: pView.Print ":"
    pView.CurrentX = 460: pView.CurrentY = 120: pView.Print ":"
    pView.FontBold = False

    pView.FontSize = 10
    pView.FontBold = True
    pView.ForeColor = vbBlack

    pView.CurrentX = 470: pView.CurrentY = 60:  pView.Print StarProbe.WaferSizemm & " mm"
    pView.CurrentX = 470: pView.CurrentY = 75:  pView.Print StarProbe.ChipSizeX & " mm X " & StarProbe.ChipSizeY & " mm"
    pView.CurrentX = 470: pView.CurrentY = 90:  pView.Print StarProbe.ChipCountX & " X " & StarProbe.ChipCountY
    pView.CurrentX = 470: pView.CurrentY = 105: pView.Print IIf(StarProbe.MeasureAll = 1 Or (XPitch(TT_NO) = 1 And YPitch(TT_NO) = 1), "FULL", XPitch(TT_NO) & " x " & YPitch(TT_NO))
    'pView.CurrentX = 470: pView.CurrentY = 120: pView.Print Dir(StarProbe.FileName_Map, vbNormal)
    pView.CurrentX = 470: pView.CurrentY = 120: pView.Print MT2000.SSPanel2(0).Caption

    Print #ifreefile, "Wafer Size," & StarProbe.WaferSizemm & " mm"
    Print #ifreefile, "Chip Size (W x H)," & StarProbe.ChipSizeX & " mm X " & StarProbe.ChipSizeY & " mm"
    Print #ifreefile, "Chip Count (X x Y)," & StarProbe.ChipCountX & " X " & StarProbe.ChipCountY
    If StarProbe.MeasureAll = 1 Or (XPitch(TT_NO) = 1 And YPitch(TT_NO) = 1) Then
        Print #ifreefile, "Pitch (X x Y),FULL"
    Else
        Print #ifreefile, "Pitch (X x Y)," & XPitch(TT_NO) & " x " & YPitch(TT_NO)
    End If
    Print #ifreefile, "MAP," & StarProbe.FileName_Map
    
    pView.FontBold = False

    Print #ifreefile, ""
    Print #ifreefile, "Measure Information"
        
    pView.CurrentX = 720
    If StarProbe.ChipSizeX < 0.18 Then pView.CurrentX = 720 + shift_value
    pView.CurrentY = 35
    pView.FontSize = 12
    pView.FontBold = True
    pView.ForeColor = vbBlack
    pView.Print "Measure Information"
    pView.FontBold = False

    pView.FontSize = 10
    pView.FontBold = False
    pView.ForeColor = vbBlack
    
    If StarProbe.ChipSizeX < 0.18 Then
        pView.CurrentX = 720 + shift_value: pView.CurrentY = 60: pView.Print "PROGRAM"
        pView.CurrentX = 720 + shift_value: pView.CurrentY = 75: pView.Print "COUNT (TOTAL)"
        pView.CurrentX = 720 + shift_value: pView.CurrentY = 90: pView.Print "COUNT (GOOD)"
        pView.CurrentX = 720 + shift_value: pView.CurrentY = 105: pView.Print "COUNT (BAD)"
        pView.CurrentX = 720 + shift_value: pView.CurrentY = 120: pView.Print "COUNT (SKIP)"
        pView.CurrentX = 720 + shift_value: pView.CurrentY = 135: pView.Print "YIELD"
        pView.CurrentX = 720 + shift_value: pView.CurrentY = 150: pView.Print "WORK TIME"
    Else
        pView.CurrentX = 720: pView.CurrentY = 60:  pView.Print "PROGRAM"
        pView.CurrentX = 720: pView.CurrentY = 75:  pView.Print "COUNT (TOTAL)"
        pView.CurrentX = 720: pView.CurrentY = 90:  pView.Print "COUNT (GOOD)"
        pView.CurrentX = 720: pView.CurrentY = 105: pView.Print "COUNT (BAD)"
        pView.CurrentX = 720: pView.CurrentY = 120: pView.Print "COUNT (SKIP)"
        pView.CurrentX = 720: pView.CurrentY = 135: pView.Print "YIELD"
        pView.CurrentX = 720: pView.CurrentY = 150: pView.Print "WORK TIME"
    End If
    pView.FontBold = False

    pView.FontSize = 10
    pView.FontBold = False
    pView.ForeColor = vbBlack
    
    If StarProbe.ChipSizeX < 0.18 Then
        pView.CurrentX = 830 + shift_value: pView.CurrentY = 60: pView.Print ":"
        pView.CurrentX = 830 + shift_value: pView.CurrentY = 75: pView.Print ":"
        pView.CurrentX = 830 + shift_value: pView.CurrentY = 90: pView.Print ":"
        pView.CurrentX = 830 + shift_value: pView.CurrentY = 105: pView.Print ":"
        pView.CurrentX = 830 + shift_value: pView.CurrentY = 120: pView.Print ":"
        pView.CurrentX = 830 + shift_value: pView.CurrentY = 135: pView.Print ":"
        pView.CurrentX = 830 + shift_value: pView.CurrentY = 150: pView.Print ":"
    Else
        pView.CurrentX = 830: pView.CurrentY = 60:  pView.Print ":"
        pView.CurrentX = 830: pView.CurrentY = 75:  pView.Print ":"
        pView.CurrentX = 830: pView.CurrentY = 90:  pView.Print ":"
        pView.CurrentX = 830: pView.CurrentY = 105: pView.Print ":"
        pView.CurrentX = 830: pView.CurrentY = 120: pView.Print ":"
        pView.CurrentX = 830: pView.CurrentY = 135: pView.Print ":"
        pView.CurrentX = 830: pView.CurrentY = 150: pView.Print ":"
    End If
    pView.FontBold = False

    pView.FontSize = 10
    pView.FontBold = True
    pView.ForeColor = vbBlack
    
    '2016.05.27
    Count_Good = val(MT2000.SSPanel_GoodCount.Caption)
    
    If StarProbe.ChipSizeX < 0.18 Then
        pView.CurrentX = 840 + shift_value: pView.CurrentY = 60: pView.Print PROD.Test_PGM
        pView.CurrentX = 840 + shift_value: pView.CurrentY = 75: pView.Print Format(Count_Total, "###,##0") & " (" & Format(Test_Cnt, "###,##0") & ")"
        pView.CurrentX = 840 + shift_value: pView.CurrentY = 90: pView.Print Format(Count_Good, "###,##0") & " (" & Format(Good_Cnt, "###,##0") & ")"
        pView.CurrentX = 840 + shift_value: pView.CurrentY = 105: pView.Print Format(Count_Bad, "###,##0") & " (" & Format(Test_Cnt - Good_Cnt, "###,##0") & ")"
        pView.CurrentX = 840 + shift_value: pView.CurrentY = 120: pView.Print Format(Count_Skip, "###,##0")
    Else
        pView.CurrentX = 840: pView.CurrentY = 60:  pView.Print PROD.Test_PGM
        pView.CurrentX = 840: pView.CurrentY = 75:  pView.Print Format(Count_Total, "###,##0") & " (" & Format(Test_Cnt, "###,##0") & ")"
        pView.CurrentX = 840: pView.CurrentY = 90:  pView.Print Format(Count_Good, "###,##0") & " (" & Format(Good_Cnt, "###,##0") & ")"
        pView.CurrentX = 840: pView.CurrentY = 105: pView.Print Format(Count_Bad, "###,##0") & " (" & Format(Test_Cnt - Good_Cnt, "###,##0") & ")"
        pView.CurrentX = 840: pView.CurrentY = 120: pView.Print Format(Count_Skip, "###,##0")
    End If
    
    If (Count_Good + Count_Bad) > 0 Then
        Yield1 = (Count_Good / (Count_Good + Count_Bad)) * 100
    Else
        Yield1 = 0
    End If
    If Test_Cnt > 0 Then
        Yield2 = (Good_Cnt / Test_Cnt) * 100
    Else
        Yield2 = 0
    End If
    
    If StarProbe.ChipSizeX < 0.18 Then
        pView.CurrentX = 840 + shift_value: pView.CurrentY = 135: pView.Print Format(Yield1, "##0.00") & "% (" & Format(Yield2, "##0.00") & "%)"
        pView.CurrentX = 840 + shift_value: pView.CurrentY = 150: pView.Print MT2000.SSPanel_DateTime
    Else
        pView.CurrentX = 840: pView.CurrentY = 135: pView.Print Format(Yield1, "##0.00") & "% (" & Format(Yield2, "##0.00") & "%)"
        pView.CurrentX = 840: pView.CurrentY = 150: pView.Print MT2000.SSPanel_DateTime
    End If
    
    Print #ifreefile, "PROGRAM," & PROD.Test_PGM
    Print #ifreefile, ",STD.,Test"
    Print #ifreefile, "COUNT (TOTAL)," & Count_Total & "," & Test_Cnt
    Print #ifreefile, "COUNT (GOOD)," & Count_Good & "," & Good_Cnt
    Print #ifreefile, "COUNT (BAD)," & Count_Bad & "," & Test_Cnt - Good_Cnt
    Print #ifreefile, "COUNT (SKIP)," & Count_Skip
    Print #ifreefile, "YIELD," & Format(Yield1, "##0.00") & "%," & Format(Yield2, "##0.00") & "%"
    Print #ifreefile, "WORK TIME," & MT2000.SSPanel_DateTime
    
    Print #ifreefile, ""
    Print #ifreefile, "BIN Information"
    
    Print #ifreefile, "BIN No,BIN Comment,STD. Count,Test Count, STD. Yield, Test Yield"
    
    pView.FontBold = False
    
    If StarProbe.ChipSizeX < 0.18 Then
        pView.CurrentX = 720 + shift_value
    Else
        pView.CurrentX = 720
    End If
    pView.CurrentY = 180
    pView.FontSize = 12
    pView.FontBold = True
    pView.ForeColor = vbBlack
    pView.Print "BIN Information"
    pView.FontBold = False
    '''''''''''''''''''''''
'    BinCount(0) = 0      '16.12.02 : 강제로 BIN0을 '0'으로 만들어준다.
    '''''''''''''''''''''''
    For i = 0 To 24
        If StarProbe.ChipSizeX < 0.18 Then
            pView.Line (720 + shift_value, (215 + (i * 15)))-(750 + shift_value, (228 + (i * 15))), BINColor(i), BF
        Else
            pView.Line (720, (215 + (i * 15)))-(750, (228 + (i * 15))), BINColor(i), BF
        End If
        
        pView.FontSize = 10
        pView.FontBold = False
        pView.ForeColor = vbBlack
                        
        If StarProbe.ChipSizeX < 0.18 Then
            pView.CurrentX = 755 + shift_value: pView.CurrentY = 214 + (i * 15): pView.Print "BIN #" & i
        Else
            pView.CurrentX = 755: pView.CurrentY = 214 + (i * 15): pView.Print "BIN #" & i
        End If
        
        pView.FontBold = False
        
        sLine = "BIN #" & i & ","
        
        '2016.05.27
        BinCount(tempbinno) = val(MT2000.SSPanel_GoodCount.Caption)
        
        pView.FontSize = 10
        pView.FontBold = True
        pView.ForeColor = vbBlack
               
        If StarProbe.ChipSizeX < 0.18 Then
            pView.CurrentX = 915 + shift_value: pView.CurrentY = 214 + (i * 15): pView.Print Format(BinCount(i), "###,##0") & " (" & Format(val(MT2000.Text_BinCount(i)), "###,##0") & ")"
        Else
            pView.CurrentX = 915: pView.CurrentY = 214 + (i * 15): pView.Print Format(BinCount(i), "###,##0") & " (" & Format(val(MT2000.Text_BinCount(i)), "###,##0") & ")"
        End If
        
        pView.FontBold = False
                
        '2016.05.27
        sLine = sLine & "," & BinCount(i) & "," & val(MT2000.Text_BinCount(i)) & ","
     
        pView.FontSize = 10
        pView.FontBold = True
        pView.ForeColor = vbBlack
        
        If Test_Cnt > 0 Then
            If (Count_Good + Count_Bad) > 0 Then
                Yield1 = (BinCount(i) / (Count_Good + Count_Bad))
            Else
                Yield1 = 0
            End If
            
            '2016.05.27
            Yield2 = (val(MT2000.Text_BinCount(i)) / Test_Cnt)
            
            If StarProbe.ChipSizeX < 0.18 Then
                pView.CurrentX = 1020 + shift_value: pView.CurrentY = 214 + (i * 15): pView.Print Format(Yield1, "##0.00%") & " (" & Format(Yield2, "##0.00%") & ")"
            Else
                pView.CurrentX = 1020: pView.CurrentY = 214 + (i * 15): pView.Print Format(Yield1, "##0.00%") & " (" & Format(Yield2, "##0.00%") & ")"
            End If
            sLine = sLine & Format(Yield1, "##0.00%") & "," & Format(Yield2, "##0.00%")
        End If
        pView.FontBold = False
        Print #ifreefile, sLine
    Next

    If StarProbe.ChipSizeX < 0.18 Then
        pView.Line (0, 32)-(1, 851 + shift_value), vbBlack, BF      '세로1
    Else
        pView.Line (0, 32)-(1, 851), vbBlack, BF        '세로1
    End If
    pView.Line (340, 32)-(340, 140), vbBlack, BF    '세로2
    
    If StarProbe.ChipSizeX < 0.18 Then
        pView.Line (709 + shift_value, 31)-(1124 + shift_value, 851 + shift_value), vbBlack, B '우측큰사각형
        pView.Line (710 + shift_value, 30)-(1123 + shift_value, 850 + shift_value), vbBlack, B '우측큰사각형
    Else
        pView.Line (709, 31)-(1124, 851), vbBlack, B    '우측큰사각형
        pView.Line (710, 30)-(1123, 850), vbBlack, B    '우측큰사각형
    End If
        
    If StarProbe.ChipSizeX < 0.18 Then
        pView.Line (0, 140)-(710 + shift_value, 141), vbBlack, BF
        pView.Line (0, 850 + shift_value)-(1125 + shift_value, 851 + shift_value), vbBlack, BF
    Else
        pView.Line (0, 140)-(710, 141), vbBlack, BF
        pView.Line (0, 850)-(1125, 851), vbBlack, BF
    End If

    pView.CurrentX = 5
    If StarProbe.ChipSizeX < 0.18 Then
        pView.CurrentY = 855 + shift_value
    Else
        pView.CurrentY = 855
    End If
    pView.FontSize = 10
    pView.FontBold = True
    pView.ForeColor = vbBlack
    pView.Print "Star Probe v1.0.0"
    pView.FontBold = False
    
    pView.CurrentX = 995
    If StarProbe.ChipSizeX < 0.18 Then
        pView.CurrentY = 855 + shift_value
    Else
        pView.CurrentY = 855
    End If
    
    pView.FontSize = 10
    pView.FontBold = True
    pView.ForeColor = vbBlack
    pView.Print Format(Date$, "YYYY.MM.DD ") & " " & Format(Time$, "HH:MM:SS")
    pView.FontBold = False
    
    Print #ifreefile, ""
    
    Me.Refresh
    
    ' Wafer 측정 데이터를 화면에 뿌려라.
    
    dx = 5
    dy = 145
    
    If StarProbe.ChipCountX >= 350 Or StarProbe.ChipCountY >= 350 Then
        DisplayX = 1 ' StarProbe.DisplayOChipSizeX
        DisplayY = 1 ' StarProbe.DisplayOChipSizeY
    ElseIf StarProbe.ChipCountX >= 200 Or StarProbe.ChipCountY >= 200 Then
        DisplayX = 2
        DisplayY = 2
    ElseIf StarProbe.ChipCountX >= 100 Or StarProbe.ChipCountY >= 100 Then
        DisplayX = 3
        DisplayY = 3
    Else
        DisplayX = 6
        DisplayY = 6
    End If
    
    For y = 0 To StarProbe.ChipCountY
        dx = 5
        For x = 0 To StarProbe.ChipCountX
            If WaferTemp(x, y).Chip Then
                If WaferTemp(x, y).ChipMask Then
                    lColor = ChipColor(1)
                ElseIf WaferTemp(x, y).ChipPlate Then
                    lColor = ChipColor(4)
                ElseIf WaferTemp(x, y).ChipSkipDie Then
                    If WaferTemp(x, y).ChipInk = True And WaferTemp(x, y).ChipInk2 = False Then
                        lColor = ChipColor(5)
                    ElseIf WaferTemp(x, y).ChipInk = True And WaferTemp(x, y).ChipInk2 = True Then
                        lColor = ChipColor(6)
                    Else
                        lColor = ChipColor(3)
                    End If
                ElseIf WaferTemp(x, y).flag Then
                    lColor = BINColor(WaferTemp(x, y).BIN)
                Else
                    lColor = BINColor(tempbinno)
                    'lColor = ChipColor(0)
                End If
                pView.Line (dx, dy)-(dx + IIf(DisplayX = 1, 0, DisplayX - 1), dy + IIf(DisplayY = 1, 0, DisplayY - 1)), lColor, BF
            End If
            dx = dx + DisplayX + IIf(DisplayX = 2, 0, 0)
        Next
        dy = dy + DisplayY + IIf(DisplayY = 2, 0, 0)
    Next
    
    pView.Refresh
    Me.Refresh
    
    Print #ifreefile, Chr(34) & "Copyright (c) 2015.10 by KEC CO., LTD. All rights reserved." & Chr(34)
    
    Close ifreefile

    'SavePicture pView.Image, StarProbe.FIleName_MeasureResult 'Form_StarProbe_Save.Text_FileName(2)
    PicSave.SavePicture pView.Image, BMP_file, fmtPNG, 70                                              '[ 2021.05.11 ] : BMP->PNG
    Kill (sfilename)        '2016.01.18
    Exit Sub
    
err:
    MsgBox "BMP파일 저장시 에러가 발생하였습니다."
'ErrorSub:
'    Call MsgBox("Summary Report Save Error" & vbCrLf & "(Error No." & Err.Number & "-" & Err.Description & ")", vbCritical + vbOKOnly, "ERROR")
End Sub


