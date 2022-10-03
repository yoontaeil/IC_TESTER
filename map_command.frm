VERSION 5.00
Begin VB.Form map_command 
   BorderStyle     =   1  '단일 고정
   Caption         =   "Map Command"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3030
   Icon            =   "map_command.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   3030
   StartUpPosition =   3  'Windows 기본값
   Begin VB.Frame Frame5 
      Caption         =   "ETC"
      Height          =   2175
      Left            =   120
      TabIndex        =   17
      Top             =   6720
      Width           =   2775
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "- H : Normal Right"
         Height          =   180
         Index           =   2
         Left            =   240
         TabIndex        =   22
         Top             =   1800
         Width           =   1545
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "- G : Normal Left"
         Height          =   180
         Index           =   1
         Left            =   240
         TabIndex        =   21
         Top             =   1440
         Width           =   1440
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "- Space Bar : Normal"
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   20
         Top             =   1080
         Width           =   1830
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "- C : Chip"
         Height          =   180
         Left            =   240
         TabIndex        =   19
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "- 2 : Mask"
         Height          =   180
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   885
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Plate Zone"
      Height          =   1455
      Left            =   120
      TabIndex        =   13
      Top             =   5160
      Width           =   2775
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "- 7 or W : Plate Zone Right"
         Height          =   180
         Left            =   240
         TabIndex        =   16
         Top             =   1080
         Width           =   2250
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "- 6 or Q : Plate Zone Left"
         Height          =   180
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Width           =   2115
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "- 3 or P : Plate Zone"
         Height          =   180
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   1740
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Skip Die"
      Height          =   1455
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Width           =   2775
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "- 5 or X : Skip Die Right"
         Height          =   180
         Left            =   240
         TabIndex        =   12
         Top             =   1080
         Width           =   1995
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "- 4 or Z : Skip Die Left"
         Height          =   180
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   1875
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "- 1 or S : Skip Die"
         Height          =   180
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   1515
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ink Die"
      Height          =   1815
      Left            =   120
      TabIndex        =   4
      Top             =   3240
      Width           =   2775
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "- O : Ink2 Die"
         Height          =   180
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   1125
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "- ] : Ink Die Right"
         Height          =   180
         Left            =   240
         TabIndex        =   7
         Top             =   1440
         Width           =   1470
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "- [ : Ink Die Left"
         Height          =   180
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   1350
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "- I : Ink1 Die"
         Height          =   180
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1035
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Bad Flag"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "- . : Bad Die Right"
         Height          =   180
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   1530
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "- , : Bad Die Left"
         Height          =   180
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   1410
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "- 0 : Bad Flag"
         Height          =   180
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1170
      End
   End
End
Attribute VB_Name = "map_command"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.Move 0, 0
End Sub
