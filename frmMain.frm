VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Path Finder"
   ClientHeight    =   8775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   585
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   731
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkShowWork 
      Caption         =   "Animate progress on this screen"
      Height          =   195
      Left            =   7305
      TabIndex        =   20
      Top             =   8550
      Width           =   3555
   End
   Begin VB.ComboBox cboPct 
      Height          =   315
      ItemData        =   "frmMain.frx":0000
      Left            =   7290
      List            =   "frmMain.frx":0025
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   390
      Width           =   2610
   End
   Begin VB.CheckBox chkBotSize 
      Caption         =   "Restrict paths to yellow bot size"
      Height          =   240
      Left            =   7305
      TabIndex        =   11
      Top             =   750
      Width           =   2745
   End
   Begin VB.CheckBox chkOvals 
      Caption         =   "Allow Oval Shapes"
      Height          =   315
      Left            =   2625
      TabIndex        =   9
      Top             =   405
      Width           =   1635
   End
   Begin VB.CheckBox chkDiagonals 
      Caption         =   "Allow Diagonals"
      Height          =   315
      Left            =   2625
      TabIndex        =   8
      Top             =   705
      Width           =   1635
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00800080&
      Caption         =   "Show Path"
      Height          =   405
      Left            =   60
      TabIndex        =   5
      Top             =   90
      Width           =   2355
   End
   Begin VB.CheckBox chkView 
      Caption         =   "Cyborg View"
      Height          =   255
      Left            =   135
      TabIndex        =   3
      Top             =   735
      Width           =   2085
   End
   Begin VB.PictureBox picMap 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   7425
      Left            =   105
      ScaleHeight     =   495
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   720
      TabIndex        =   0
      Top             =   1065
      Width           =   10800
      Begin VB.CommandButton cmdAni 
         BackColor       =   &H0000FFFF&
         Enabled         =   0   'False
         Height          =   195
         Index           =   1
         Left            =   255
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   4200
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.CommandButton cmdAni 
         BackColor       =   &H000000C0&
         Enabled         =   0   'False
         Height          =   195
         Index           =   0
         Left            =   255
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   3930
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.CommandButton cmdStop 
         BackColor       =   &H000000C0&
         Caption         =   "Z"
         Height          =   240
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1260
         Width           =   240
      End
      Begin VB.CommandButton cmdStart 
         BackColor       =   &H0000FFFF&
         Caption         =   "A"
         Height          =   240
         Left            =   990
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   2085
         Width           =   240
      End
      Begin VB.Line AntiLines 
         BorderColor     =   &H000000C0&
         BorderWidth     =   3
         Index           =   4
         Visible         =   0   'False
         X1              =   143
         X2              =   175
         Y1              =   203
         Y2              =   166
      End
      Begin VB.Line AntiLines 
         BorderColor     =   &H000000C0&
         BorderWidth     =   3
         Index           =   3
         Visible         =   0   'False
         X1              =   532
         X2              =   371
         Y1              =   336
         Y2              =   375
      End
      Begin VB.Line AntiLines 
         BorderColor     =   &H000000C0&
         BorderWidth     =   3
         Index           =   2
         Visible         =   0   'False
         X1              =   454
         X2              =   333
         Y1              =   394
         Y2              =   367
      End
      Begin VB.Line AntiLines 
         BorderColor     =   &H000000C0&
         BorderWidth     =   3
         Index           =   1
         Visible         =   0   'False
         X1              =   484
         X2              =   403
         Y1              =   307
         Y2              =   423
      End
      Begin VB.Line AntiLines 
         BorderColor     =   &H000000C0&
         BorderWidth     =   3
         Index           =   0
         Visible         =   0   'False
         X1              =   190
         X2              =   143
         Y1              =   196
         Y2              =   149
      End
      Begin VB.Shape ShapeR 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   165
         Index           =   32
         Left            =   735
         Top             =   1365
         Width           =   915
      End
      Begin VB.Shape ShapeR 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   1785
         Index           =   31
         Left            =   1305
         Top             =   1815
         Width           =   240
      End
      Begin VB.Shape ShapeR 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   1170
         Index           =   30
         Left            =   735
         Top             =   1455
         Width           =   165
      End
      Begin VB.Shape ShapeR 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   165
         Index           =   29
         Left            =   735
         Top             =   2535
         Width           =   1695
      End
      Begin VB.Shape ShapeR 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   135
         Index           =   28
         Left            =   750
         Top             =   975
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "The starting and ending nodes (A && Z) can be dragged && dropped during run-time. They cannot intersect obstacles else path fails."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   435
         Index           =   1
         Left            =   3960
         TabIndex        =   16
         Top             =   6840
         Visible         =   0   'False
         Width           =   6015
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":008B
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   435
         Index           =   0
         Left            =   2985
         TabIndex        =   15
         Top             =   2085
         Visible         =   0   'False
         Width           =   6015
      End
      Begin VB.Shape ShapeR 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   135
         Index           =   27
         Left            =   5850
         Top             =   1560
         Width           =   1530
      End
      Begin VB.Shape ShapeR 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   1005
         Index           =   26
         Left            =   7290
         Top             =   1080
         Width           =   135
      End
      Begin VB.Shape ShapeR 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   630
         Index           =   25
         Left            =   9120
         Top             =   2625
         Width           =   720
      End
      Begin VB.Shape ShapeR 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   855
         Index           =   24
         Left            =   8685
         Top             =   4800
         Width           =   150
      End
      Begin VB.Shape ShapeR 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   1230
         Index           =   23
         Left            =   9645
         Top             =   4395
         Width           =   195
      End
      Begin VB.Shape ShapeR 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   735
         Index           =   22
         Left            =   1005
         Top             =   5190
         Width           =   1035
      End
      Begin VB.Shape ShapeR 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   330
         Index           =   21
         Left            =   3135
         Top             =   6450
         Width           =   360
      End
      Begin VB.Shape ShapeR 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   1350
         Index           =   20
         Left            =   5415
         Top             =   1500
         Width           =   270
      End
      Begin VB.Shape ShapeR 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   330
         Index           =   19
         Left            =   3060
         Top             =   3810
         Width           =   360
      End
      Begin VB.Shape ShapeR 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   195
         Index           =   18
         Left            =   4890
         Top             =   4110
         Width           =   1530
      End
      Begin VB.Shape ShapeR 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   150
         Index           =   17
         Left            =   7890
         Top             =   4275
         Width           =   1590
      End
      Begin VB.Shape ShapeR 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   135
         Index           =   16
         Left            =   1335
         Top             =   3495
         Width           =   1845
      End
      Begin VB.Shape ShapeR 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   330
         Index           =   15
         Left            =   2325
         Top             =   4470
         Width           =   360
      End
      Begin VB.Shape ShapeR 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   120
         Index           =   14
         Left            =   4500
         Top             =   1080
         Width           =   2790
      End
      Begin VB.Shape ShapeR 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   1230
         Index           =   13
         Left            =   7785
         Top             =   465
         Width           =   165
      End
      Begin VB.Shape ShapeR 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   330
         Index           =   12
         Left            =   6360
         Top             =   435
         Width           =   1590
      End
      Begin VB.Shape ShapeR 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   1005
         Index           =   11
         Left            =   3885
         Top             =   5250
         Width           =   360
      End
      Begin VB.Shape ShapeR 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   195
         Index           =   10
         Left            =   6030
         Top             =   3690
         Width           =   2955
      End
      Begin VB.Shape ShapeR 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   1905
         Index           =   9
         Left            =   4560
         Top             =   3390
         Width           =   135
      End
      Begin VB.Shape ShapeR 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   330
         Index           =   8
         Left            =   8685
         Top             =   630
         Width           =   360
      End
      Begin VB.Shape ShapeR 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   1230
         Index           =   7
         Left            =   3510
         Top             =   435
         Width           =   225
      End
      Begin VB.Shape ShapeR 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   3630
         Index           =   6
         Left            =   3420
         Top             =   3375
         Width           =   150
      End
      Begin VB.Shape ShapeR 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   330
         Index           =   5
         Left            =   2985
         Top             =   7125
         Width           =   840
      End
      Begin VB.Shape ShapeR 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   1530
         Index           =   4
         Left            =   2640
         Top             =   1110
         Width           =   165
      End
      Begin VB.Shape ShapeR 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   195
         Index           =   3
         Left            =   6075
         Top             =   2775
         Width           =   2430
      End
      Begin VB.Shape ShapeR 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   120
         Index           =   2
         Left            =   3435
         Top             =   3360
         Width           =   3420
      End
      Begin VB.Shape ShapeR 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   165
         Index           =   1
         Left            =   3555
         Top             =   1500
         Width           =   2130
      End
      Begin VB.Shape ShapeR 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   330
         Index           =   0
         Left            =   3540
         Top             =   435
         Width           =   2475
      End
      Begin VB.Line Lines 
         BorderColor     =   &H00404040&
         BorderWidth     =   3
         Index           =   3
         Visible         =   0   'False
         X1              =   103
         X2              =   153
         Y1              =   99
         Y2              =   69
      End
      Begin VB.Line Lines 
         BorderColor     =   &H00404040&
         BorderWidth     =   3
         Index           =   2
         Visible         =   0   'False
         X1              =   632
         X2              =   522
         Y1              =   413
         Y2              =   338
      End
      Begin VB.Shape ShapeO 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   900
         Index           =   3
         Left            =   5910
         Shape           =   2  'Oval
         Top             =   5445
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.Shape ShapeO 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   450
         Index           =   4
         Left            =   6195
         Shape           =   2  'Oval
         Top             =   5295
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Shape ShapeO 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   870
         Index           =   1
         Left            =   5595
         Shape           =   2  'Oval
         Top             =   4650
         Visible         =   0   'False
         Width           =   2205
      End
      Begin VB.Shape ShapeO 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   1125
         Index           =   2
         Left            =   5085
         Shape           =   2  'Oval
         Top             =   5145
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Shape ShapeO 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   780
         Index           =   0
         Left            =   2055
         Shape           =   2  'Oval
         Top             =   2205
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Line Lines 
         BorderColor     =   &H00404040&
         BorderWidth     =   3
         Index           =   1
         Visible         =   0   'False
         X1              =   585
         X2              =   634
         Y1              =   166
         Y2              =   112
      End
      Begin VB.Line Lines 
         BorderColor     =   &H00404040&
         BorderWidth     =   3
         Index           =   0
         Visible         =   0   'False
         X1              =   95
         X2              =   228
         Y1              =   495
         Y2              =   410
      End
   End
   Begin VB.PictureBox picWork 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      FillColor       =   &H8000000F&
      ForeColor       =   &H00C0C000&
      Height          =   7425
      Left            =   90
      ScaleHeight     =   495
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   720
      TabIndex        =   4
      Top             =   1095
      Visible         =   0   'False
      Width           =   10800
      Begin VB.CommandButton cmdMakerA 
         BackColor       =   &H0000FFFF&
         Caption         =   "A"
         Height          =   240
         Left            =   1020
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   945
         Width           =   240
      End
      Begin VB.CommandButton cmdMarkerZ 
         BackColor       =   &H000000C0&
         Caption         =   "Z"
         Height          =   240
         Left            =   8040
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3045
         Width           =   240
      End
   End
   Begin VB.Label lblLength 
      Caption         =   "Path Length:  0 units"
      ForeColor       =   &H00800080&
      Height          =   255
      Left            =   75
      TabIndex        =   19
      Top             =   495
      Width           =   2340
   End
   Begin VB.Label Label1 
      Caption         =   "Shortest path accuracy affects both speed && accuracy on complex maps"
      ForeColor       =   &H00FF0000&
      Height          =   465
      Index           =   2
      Left            =   7290
      TabIndex        =   18
      Top             =   -15
      Width           =   2580
   End
   Begin VB.Label Label1 
      Caption         =   "decreases tme, but recreates map >>"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   1
      Left            =   4365
      TabIndex        =   14
      Top             =   780
      Width           =   2820
   End
   Begin VB.Label Label1 
      Caption         =   "<< increases calculation time"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   0
      Left            =   4395
      TabIndex        =   13
      Top             =   465
      Width           =   2580
   End
   Begin VB.Label lblStat 
      Caption         =   "Ready"
      Height          =   255
      Left            =   90
      TabIndex        =   12
      Top             =   8535
      Width           =   8460
   End
   Begin VB.Label lblTime 
      Caption         =   "Path Calculation Time"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   2625
      TabIndex        =   10
      Top             =   150
      Width           =   3165
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' This form is just to display what the class returns

' This form is not commented because if it's lack of value.

' The class is heavily commented and is a work in progress.


Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function FillRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function CreateRectRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetRegionData Lib "gdi32" (ByVal hRgn As Long, ByVal dwCount As Long, lpRgnData As Any) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetRectEmpty Lib "user32" (lpRect As RECT) As Long
Private Declare Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Const RGN_AND As Long = 1
Private Const RGN_COPY As Long = 5
Private Const RGN_DIFF As Long = 4
Private Const RGN_OR As Long = 2
Private Const RGN_XOR As Long = 3

Private rgnMap As Long
Private cMap As clsPathFinder
Private bDirty As Boolean

Private Sub CreateRegions()
    If rgnMap Then DeleteObject rgnMap
    rgnMap = CreateRectRgn(0, 0, picMap.Width, picMap.Height)
End Sub

Private Function ExtractRectangles(hRgn As Long, bAdd2Class As Boolean) As Long

Dim rSize As Long, X As Long, rgnLooper As Long, vRgnData() As Byte, hRect As RECT

'For rgnLooper = 1 To 1
    ' 1st get the buffer size needed to return rectangles info from this region
    rSize = GetRegionData(hRgn, ByVal 0&, ByVal 0&)
    If rSize > 0 Then   ' success
        ' create the buffer & call function again to fill the buffer
        ReDim vRgnData(0 To rSize - 1) As Byte
        If rSize = GetRegionData(hRgn, rSize, vRgnData(0)) Then     ' success
        
            ' Here are some tips of the structure returned
            ' Bytes 8-11 are the number of rectangles in the region
            ' Bytes 12-15 is structure size information -- not important for what we need
            ' Bytes 16-31 are the bounding rectangle's dimensions
            ' Bytes 32 to end of structure are the individual rectangle's dimensions
            ' The rectangle structure (RECT) is 16 bytes or LenB(RECT)
        
            ' Let's retrieve the number of rectangles in the structure (b:8-11)
            CopyMemory rSize, vRgnData(8), ByVal 4&
            'If bAdd2Class Then cMap.NumberNodes = rSize
        
            For X = 0 To rSize - 1
                CopyMemory hRect, vRgnData(X * 16 + 32), ByVal 16&
                If bAdd2Class Then
                    'cMap.AddNode hRect.Left, hRect.Top, hRect.Right, hRect.Bottom
                Else
                
                    Rectangle picWork.hdc, hRect.Left, hRect.Top, hRect.Right, hRect.Bottom
                End If
            Next
        Else
            Stop
        End If
    Else
        Stop
    End If
'Next
Erase vRgnData
ExtractRectangles = rSize
End Function

Private Sub CreateWorkArea()

If bDirty Then

    If Not cMap Is Nothing Then Set cMap = Nothing
    Set cMap = New clsPathFinder
    
End If

    lblStat = "Drawing the map"
    lblStat.Refresh



picWork.Cls
picMap.Cls

Dim I As Integer, tRgn As Long, tRect As RECT, cRgn As Long, tBrush As Long

If bDirty Then

    CreateRegions
    
    For I = 0 To ShapeR.UBound
        With ShapeR(I)
            SetRect tRect, .Left, .Top, .Width + .Left, .Top + .Height
        End With
        tRgn = CreateRectRgnIndirect(tRect)
        CombineRgn rgnMap, rgnMap, tRgn, RGN_DIFF
        DeleteObject tRgn
    Next
    
End If
    
    If chkOvals Then
        tBrush = CreateSolidBrush(&H404040)
        For I = 0 To ShapeO.UBound
            With ShapeO(I)
                SetRect tRect, .Left, .Top, .Width + .Left, .Top + .Height
            End With
            tRgn = CreateEllipticRgn(tRect.Left, tRect.Top, tRect.Right, tRect.Bottom)
            If bDirty Then
                CombineRgn rgnMap, rgnMap, tRgn, RGN_DIFF
            End If
            FillRgn picMap.hdc, tRgn, tBrush
            DeleteObject tRgn
        Next
        DeleteObject tBrush
    End If


If chkDiagonals Then
    Dim tPoints(0 To 3) As POINTAPI, Looper As Integer
    Dim myLine As Line
    
    For Looper = 1 To 2
        tBrush = CreateSolidBrush(Choose(Looper, &H404040, 0&))
        
        For I = 0 To Choose(Looper, Lines.UBound, AntiLines.UBound)
            If Looper = 1 Then
                Set myLine = Lines(I)
            Else
                Set myLine = AntiLines(I)
            End If
            With myLine
                If .X1 > .X2 Then
                    tPoints(0).X = .X1
                    tPoints(1).X = .X1 + .BorderWidth
                    tPoints(2).X = .X2 + .BorderWidth
                    tPoints(3).X = .X2
                Else
                    tPoints(0).X = .X1 + .BorderWidth
                    tPoints(1).X = .X1
                    tPoints(2).X = .X2
                    tPoints(3).X = .X2 + .BorderWidth
                End If
                If .Y1 < .Y2 Then
                    tPoints(0).Y = .Y1 - .BorderWidth
                    tPoints(1).Y = .Y1
                    tPoints(2).Y = .Y2
                    tPoints(3).Y = .Y2 - .BorderWidth
                Else
                    tPoints(0).Y = .Y1
                    tPoints(1).Y = .Y1 - .BorderWidth
                    tPoints(2).Y = .Y2 - .BorderWidth
                    tPoints(3).Y = .Y2
                End If
            End With
            tRgn = CreatePolygonRgn(tPoints(0), UBound(tPoints) + 1, 2&)
            If bDirty Then
                CombineRgn rgnMap, rgnMap, tRgn, Choose(Looper, RGN_DIFF, RGN_OR)
            End If
            FillRgn picMap.hdc, tRgn, tBrush
            DeleteObject tRgn
        Next
        DeleteObject tBrush
    Next
    Set myLine = Nothing
    Erase tPoints
End If

picWork.ForeColor = &H808080
ExtractRectangles rgnMap, False

If bDirty Then

    lblStat = "Creating the map"
    lblStat.Refresh
    
    ' the second parameter determines the minimum distance from the obstacle
    ' walls. A value of zero will ride alongside any obstacle, if applicable.
    ' Here, I'm supplying a value of 1 if the the check box isn't checked
    ' so there is a one pixel separation between edges & the path
    cMap.SetMapRegion rgnMap, (Abs(chkBotSize.Value) * (cmdStart.Width - 1)) + 1

End If
bDirty = False
End Sub


Private Sub chkBotSize_Click()
bDirty = True
DisplayWorkArea
End Sub

Private Sub chkDiagonals_Click()
bDirty = True
DisplayWorkArea
End Sub

Private Sub chkOvals_Click()
bDirty = True
DisplayWorkArea
End Sub

Private Sub cmdMakerA_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Or Button = 5 Then
    MoveObject cmdMakerA.hwnd
    ReleaseCapture
    cmdStart.Move cmdMakerA.Left, cmdMakerA.Top
End If
End Sub

Private Sub cmdMarkerZ_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Or Button = 5 Then
    MoveObject cmdMarkerZ.hwnd
    ReleaseCapture
    cmdStop.Move cmdMarkerZ.Left, cmdMarkerZ.Top
End If
End Sub


Private Sub cmdStart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Or Button = 5 Then
    MoveObject cmdStart.hwnd
    ReleaseCapture
    cmdMakerA.Move cmdStart.Left, cmdStart.Top
End If
End Sub

Private Sub cmdStop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Or Button = 5 Then
    MoveObject cmdStop.hwnd
    ReleaseCapture
    cmdMarkerZ.Move cmdStop.Left, cmdStop.Top
End If
End Sub

Private Sub Command1_Click()

DisplayWorkArea

lblStat = "Generating path"
lblStat.Refresh

Dim sX As Long, sY As Long, eX As Long, eY As Long
Dim PathPts() As Long

With cmdStart
    sX = .Left + .Width \ 2
    sY = .Top + .Height \ 2
End With
With cmdStop
    eX = .Left + .Width \ 2
    eY = .Top + .Height \ 2
End With

If chkShowWork.Value = 1 And chkShowWork.Enabled Then
    cmdAni(0).Visible = True
    cmdAni(1).Visible = True
    Refresh
End If

Dim X As Long, pathLen As Long
X = GetTickCount()
DoEvents
pathLen = cMap.CreatePaths(sX, sY, eX, eY, PathPts(), cboPct.ListIndex * 10)
If pathLen Then
    lblTime.Caption = GetTickCount - X & " milliseconds to find path"
    DoEvents
    DisplayPath PathPts()
    lblLength.Caption = "Path Length: " & pathLen & " units"
Else
    lblTime = "failed to find path"
    MsgBox "Failed to find path. Ensure red/yellow bots are not overlapping any obstacles."
End If
If chkShowWork.Value = 1 And chkShowWork.Enabled Then
    cmdAni(0).Visible = False
    cmdAni(1).Visible = False
End If

picWork.ForeColor = &H808080

lblStat = "Ready"
lblStat.Refresh

End Sub

Private Sub DisplayPath(vPaths() As Long)
Dim X As Long, Y As Long, X1 As Long, Y1 As Long, Pts As Long
'On Error GoTo ExitRoutine
picWork.ForeColor = vbWhite
picMap.ForeColor = vbGreen

X = LoWord(vPaths(Pts))
Y = HiWord(vPaths(Pts))

For Pts = 1 To UBound(vPaths)
    X1 = LoWord(vPaths(Pts))
    Y1 = HiWord(vPaths(Pts))
    picWork.Line (X, Y)-(X1, Y1)
    picMap.Line (X, Y)-(X1, Y1)
    X = X1
    Y = Y1
Next

ExitRoutine:
picWork.Refresh
picMap.Refresh

End Sub

Private Sub Form_Load()
cmdMakerA.Move cmdStart.Left, cmdStart.Top
cmdMarkerZ.Move cmdStop.Left, cmdStop.Top
cboPct.ListIndex = 10
bDirty = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not cMap Is Nothing Then Set cMap = Nothing
If rgnMap Then DeleteObject rgnMap
End Sub

Private Sub chkView_Click()
Dim bToggle As Boolean
If chkView Then
    picWork.ZOrder
    bToggle = True
Else
    picMap.ZOrder
End If
picMap.Enabled = Not bToggle
picWork.Enabled = bToggle
picMap.Visible = Not bToggle
picWork.Visible = bToggle
If bToggle = True And rgnMap = 0 Then DisplayWorkArea
chkShowWork.Enabled = picMap.Enabled
End Sub

Private Sub DisplayWorkArea()

CreateWorkArea


Dim tBrush As Long, tRgn As Long
tBrush = CreateSolidBrush(&H800080)
tRgn = CreateRectRgn(0, 0, picWork.Width, picWork.Height)
CombineRgn tRgn, tRgn, rgnMap, RGN_DIFF
FillRgn picWork.hdc, tRgn, tBrush
DeleteObject tBrush
DeleteObject tRgn
bDirty = False
picWork.Refresh
picMap.Refresh

lblStat = "Ready"

End Sub

