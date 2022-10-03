VERSION 5.00
Begin VB.Form Form_StarProbe_DeviceInfo 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  '크기 고정 대화 상자
   ClientHeight    =   9195
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   9135
   ControlBox      =   0   'False
   Icon            =   "Form_StarProbe_DeviceInfo.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   613
   ScaleMode       =   3  '픽셀
   ScaleWidth      =   609
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows 기본값
   Begin VB.PictureBox pDevice 
      BackColor       =   &H00FFFFFF&
      Height          =   9195
      Left            =   0
      ScaleHeight     =   609
      ScaleMode       =   3  '픽셀
      ScaleWidth      =   605
      TabIndex        =   0
      Top             =   0
      Width           =   9135
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Start Chip "
         Height          =   795
         Index           =   5
         Left            =   2580
         TabIndex        =   6
         Top             =   420
         Width           =   1215
         Begin VB.Shape Shape1 
            BorderColor     =   &H000000FF&
            BorderWidth     =   2
            Height          =   315
            Index           =   29
            Left            =   180
            Top             =   300
            Width           =   315
         End
         Begin VB.Shape Shape1 
            FillStyle       =   0  '단색
            Height          =   195
            Index           =   25
            Left            =   240
            Top             =   360
            Width           =   195
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Right "
         Height          =   1035
         Index           =   4
         Left            =   6240
         TabIndex        =   5
         Top             =   4080
         Width           =   975
         Begin VB.Shape Shape1 
            BorderColor     =   &H000000FF&
            BorderWidth     =   2
            Height          =   315
            Index           =   24
            Left            =   180
            Top             =   300
            Width           =   315
         End
         Begin VB.Shape Shape1 
            FillStyle       =   0  '단색
            Height          =   195
            Index           =   23
            Left            =   540
            Top             =   660
            Width           =   195
         End
         Begin VB.Shape Shape1 
            FillStyle       =   0  '단색
            Height          =   195
            Index           =   22
            Left            =   240
            Top             =   660
            Width           =   195
         End
         Begin VB.Shape Shape1 
            FillStyle       =   0  '단색
            Height          =   195
            Index           =   21
            Left            =   540
            Top             =   360
            Width           =   195
         End
         Begin VB.Shape Shape1 
            FillStyle       =   0  '단색
            Height          =   195
            Index           =   20
            Left            =   240
            Top             =   360
            Width           =   195
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Left "
         Height          =   1035
         Index           =   3
         Left            =   1920
         TabIndex        =   4
         Top             =   4080
         Width           =   975
         Begin VB.Shape Shape1 
            BorderColor     =   &H000000FF&
            BorderWidth     =   2
            Height          =   315
            Index           =   19
            Left            =   180
            Top             =   300
            Width           =   315
         End
         Begin VB.Shape Shape1 
            FillStyle       =   0  '단색
            Height          =   195
            Index           =   18
            Left            =   540
            Top             =   660
            Width           =   195
         End
         Begin VB.Shape Shape1 
            FillStyle       =   0  '단색
            Height          =   195
            Index           =   17
            Left            =   240
            Top             =   660
            Width           =   195
         End
         Begin VB.Shape Shape1 
            FillStyle       =   0  '단색
            Height          =   195
            Index           =   16
            Left            =   540
            Top             =   360
            Width           =   195
         End
         Begin VB.Shape Shape1 
            FillStyle       =   0  '단색
            Height          =   195
            Index           =   15
            Left            =   240
            Top             =   360
            Width           =   195
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Bottom "
         Height          =   1035
         Index           =   2
         Left            =   4080
         TabIndex        =   3
         Top             =   6240
         Width           =   975
         Begin VB.Shape Shape1 
            BorderColor     =   &H000000FF&
            BorderWidth     =   2
            Height          =   315
            Index           =   14
            Left            =   180
            Top             =   300
            Width           =   315
         End
         Begin VB.Shape Shape1 
            FillStyle       =   0  '단색
            Height          =   195
            Index           =   13
            Left            =   540
            Top             =   660
            Width           =   195
         End
         Begin VB.Shape Shape1 
            FillStyle       =   0  '단색
            Height          =   195
            Index           =   12
            Left            =   240
            Top             =   660
            Width           =   195
         End
         Begin VB.Shape Shape1 
            FillStyle       =   0  '단색
            Height          =   195
            Index           =   11
            Left            =   540
            Top             =   360
            Width           =   195
         End
         Begin VB.Shape Shape1 
            FillStyle       =   0  '단색
            Height          =   195
            Index           =   10
            Left            =   240
            Top             =   360
            Width           =   195
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Top "
         Height          =   1035
         Index           =   1
         Left            =   4080
         TabIndex        =   2
         Top             =   1920
         Width           =   975
         Begin VB.Shape Shape1 
            FillStyle       =   0  '단색
            Height          =   195
            Index           =   9
            Left            =   240
            Top             =   360
            Width           =   195
         End
         Begin VB.Shape Shape1 
            FillStyle       =   0  '단색
            Height          =   195
            Index           =   8
            Left            =   540
            Top             =   360
            Width           =   195
         End
         Begin VB.Shape Shape1 
            FillStyle       =   0  '단색
            Height          =   195
            Index           =   7
            Left            =   240
            Top             =   660
            Width           =   195
         End
         Begin VB.Shape Shape1 
            FillStyle       =   0  '단색
            Height          =   195
            Index           =   6
            Left            =   540
            Top             =   660
            Width           =   195
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H000000FF&
            BorderWidth     =   2
            Height          =   315
            Index           =   5
            Left            =   180
            Top             =   300
            Width           =   315
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Center "
         Height          =   1035
         Index           =   0
         Left            =   4080
         TabIndex        =   1
         Top             =   4080
         Width           =   975
         Begin VB.Shape Shape1 
            BorderColor     =   &H000000FF&
            BorderWidth     =   2
            Height          =   315
            Index           =   4
            Left            =   180
            Top             =   300
            Width           =   315
         End
         Begin VB.Shape Shape1 
            FillStyle       =   0  '단색
            Height          =   195
            Index           =   3
            Left            =   540
            Top             =   660
            Width           =   195
         End
         Begin VB.Shape Shape1 
            FillStyle       =   0  '단색
            Height          =   195
            Index           =   2
            Left            =   240
            Top             =   660
            Width           =   195
         End
         Begin VB.Shape Shape1 
            FillStyle       =   0  '단색
            Height          =   195
            Index           =   1
            Left            =   540
            Top             =   360
            Width           =   195
         End
         Begin VB.Shape Shape1 
            FillStyle       =   0  '단색
            Height          =   195
            Index           =   0
            Left            =   240
            Top             =   360
            Width           =   195
         End
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   4
         X1              =   204
         X2              =   276
         Y1              =   88
         Y2              =   264
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   3
         X1              =   304
         X2              =   304
         Y1              =   348
         Y2              =   408
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   2
         X1              =   348
         X2              =   404
         Y1              =   308
         Y2              =   308
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   1
         X1              =   304
         X2              =   304
         Y1              =   204
         Y2              =   264
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   0
         X1              =   204
         X2              =   260
         Y1              =   308
         Y2              =   308
      End
   End
End
Attribute VB_Name = "Form_StarProbe_DeviceInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    Me.Top = 630
    Me.Left = 300

End Sub
