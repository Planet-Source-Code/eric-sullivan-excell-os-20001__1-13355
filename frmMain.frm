VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "MacShell 1.0.0"
   ClientHeight    =   12960
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   17280
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmMain.frx":08CA
   MousePointer    =   1  'Arrow
   ScaleHeight     =   12960
   ScaleWidth      =   17280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   855
      Left            =   1560
      TabIndex        =   36
      Top             =   480
      Visible         =   0   'False
      Width           =   1815
      Begin VB.Label MnuAlign 
         BackColor       =   &H00FFFFFF&
         Caption         =   "To right side"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   45
         TabIndex        =   39
         Top             =   450
         Width           =   1605
      End
      Begin VB.Line Line24 
         BorderColor     =   &H00404040&
         X1              =   1680
         X2              =   1680
         Y1              =   0
         Y2              =   720
      End
      Begin VB.Line Line23 
         BorderColor     =   &H00404040&
         X1              =   1680
         X2              =   0
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line22 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   0
         Y1              =   720
         Y2              =   0
      End
      Begin VB.Line Line21 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   1680
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label MnuAlign 
         BackColor       =   &H00FFFFFF&
         Caption         =   "To left side"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   45
         TabIndex        =   38
         Top             =   240
         Width           =   1605
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         Caption         =   "Align..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   215
         Left            =   0
         TabIndex        =   37
         Top             =   0
         Width           =   1695
      End
      Begin VB.Shape Shape19 
         BackColor       =   &H00808080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         Height          =   735
         Left            =   1680
         Top             =   120
         Width           =   135
      End
      Begin VB.Shape Shape18 
         BackColor       =   &H00808080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         Height          =   135
         Left            =   120
         Top             =   720
         Width           =   1695
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   855
      Left            =   1560
      TabIndex        =   32
      Top             =   960
      Visible         =   0   'False
      Width           =   1815
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         Caption         =   "New..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   215
         Left            =   0
         TabIndex        =   35
         Top             =   0
         Width           =   1695
      End
      Begin VB.Label MnuNew 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Folder"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   45
         TabIndex        =   34
         Top             =   240
         Width           =   1605
      End
      Begin VB.Line Line20 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   1680
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line19 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   0
         Y1              =   720
         Y2              =   0
      End
      Begin VB.Line Line18 
         BorderColor     =   &H00404040&
         X1              =   1680
         X2              =   0
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line17 
         BorderColor     =   &H00404040&
         X1              =   1680
         X2              =   1680
         Y1              =   0
         Y2              =   720
      End
      Begin VB.Label MnuNew 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Text document"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   45
         TabIndex        =   33
         Top             =   480
         Width           =   1605
      End
      Begin VB.Shape Shape16 
         BackColor       =   &H00808080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         Height          =   135
         Left            =   120
         Top             =   720
         Width           =   1695
      End
      Begin VB.Shape Shape17 
         BackColor       =   &H00808080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         Height          =   735
         Left            =   1680
         Top             =   120
         Width           =   135
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   1095
      Left            =   1560
      TabIndex        =   27
      Top             =   720
      Visible         =   0   'False
      Width           =   1815
      Begin VB.Label MnuClock 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   45
         TabIndex        =   31
         Top             =   675
         Width           =   1605
      End
      Begin VB.Label MnuClock 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Time && date"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   45
         TabIndex        =   29
         Top             =   450
         Width           =   1605
      End
      Begin VB.Line Line16 
         BorderColor     =   &H00404040&
         X1              =   1680
         X2              =   1680
         Y1              =   0
         Y2              =   960
      End
      Begin VB.Line Line15 
         BorderColor     =   &H00404040&
         X1              =   1680
         X2              =   0
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Line14 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   0
         Y1              =   960
         Y2              =   0
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   1680
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label MnuClock 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   45
         TabIndex        =   28
         Top             =   240
         Width           =   1605
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         Caption         =   "View..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   215
         Left            =   0
         TabIndex        =   30
         Top             =   0
         Width           =   1695
      End
      Begin VB.Shape Shape14 
         BackColor       =   &H00808080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         Height          =   135
         Left            =   120
         Top             =   960
         Width           =   1695
      End
      Begin VB.Shape Shape15 
         BackColor       =   &H00808080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         Height          =   975
         Left            =   1680
         Top             =   120
         Width           =   135
      End
   End
   Begin VB.Timer TimerTime 
      Interval        =   1000
      Left            =   3000
      Top             =   11880
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   3480
      Top             =   11880
   End
   Begin VB.TextBox ChangeLabelCap 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4440
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   12120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3960
      Top             =   11880
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   12480
      Width           =   17055
      Begin VB.OptionButton Option1 
         Caption         =   "Desktop"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   0
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   18
      Text            =   "frmMain.frx":1194
      Top             =   5280
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   1455
      Left            =   240
      TabIndex        =   21
      Top             =   400
      Visible         =   0   'False
      Width           =   1455
      Begin VB.Label Mnu1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   5
         Left            =   45
         TabIndex        =   26
         Top             =   1080
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00808080&
         Visible         =   0   'False
         X1              =   0
         X2              =   1320
         Y1              =   1010
         Y2              =   1010
      End
      Begin VB.Label Mnu1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Options"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   4
         Left            =   45
         TabIndex        =   25
         Top             =   750
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00808080&
         Visible         =   0   'False
         X1              =   0
         X2              =   1320
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Mnu1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "New            >"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   3
         Left            =   45
         TabIndex        =   24
         Top             =   480
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label Mnu1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Clock          >"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   2
         Left            =   45
         TabIndex        =   23
         Top             =   280
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Line Line13 
         BorderColor     =   &H00808080&
         Visible         =   0   'False
         X1              =   0
         X2              =   1320
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Label Mnu1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Align           >"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   1
         Left            =   45
         TabIndex        =   22
         Top             =   10
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Line Line11 
         BorderColor     =   &H00404040&
         Visible         =   0   'False
         X1              =   0
         X2              =   1320
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00808080&
         Visible         =   0   'False
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   1320
      End
      Begin VB.Line Line12 
         BorderColor     =   &H00404040&
         Visible         =   0   'False
         X1              =   1320
         X2              =   1320
         Y1              =   0
         Y2              =   1320
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00808080&
         Visible         =   0   'False
         X1              =   0
         X2              =   1320
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Shape Shape12 
         BackColor       =   &H00808080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         Height          =   135
         Left            =   120
         Top             =   1320
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Shape Shape11 
         BackColor       =   &H00808080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         Height          =   1335
         Left            =   1320
         Top             =   120
         Visible         =   0   'False
         Width           =   135
      End
   End
   Begin VB.Shape Shape13 
      Height          =   12135
      Left            =   16320
      Top             =   360
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Label Mnu1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&System"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   20
      Top             =   120
      Width           =   855
   End
   Begin VB.Label MnuHelp 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Help..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   19
      Top             =   120
      Width           =   615
   End
   Begin VB.Shape Shape6 
      Height          =   12135
      Left            =   240
      Top             =   360
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Shape Bubble 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   2
      Left            =   4920
      Shape           =   3  'Circle
      Top             =   4800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape Bubble 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      Height          =   135
      Index           =   0
      Left            =   4440
      Shape           =   3  'Circle
      Top             =   4560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape Bubble 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   1
      Left            =   4680
      Shape           =   3  'Circle
      Top             =   4680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label6 
      Caption         =   "0"
      Height          =   255
      Left            =   1680
      TabIndex        =   17
      Top             =   12120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "0"
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   12120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label LblDoc 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "New text document #1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   8280
      TabIndex        =   15
      Top             =   1320
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Image ImgFolder 
      Height          =   480
      Index           =   0
      Left            =   7200
      Picture         =   "frmMain.frx":123E
      Top             =   840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label LblFolder 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "New Folder #1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   6720
      TabIndex        =   12
      Top             =   1320
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Image ImgDoc 
      Height          =   480
      Index           =   0
      Left            =   8640
      Picture         =   "frmMain.frx":1B08
      Top             =   840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image DesktopImage 
      Height          =   480
      Index           =   4
      Left            =   240
      Picture         =   "frmMain.frx":23D2
      Top             =   4800
      Width           =   480
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Rename"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   7680
      MouseIcon       =   "frmMain.frx":2C9C
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Move..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   6720
      MouseIcon       =   "frmMain.frx":2FA6
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   8760
      MouseIcon       =   "frmMain.frx":32B0
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   480
      Width           =   615
   End
   Begin VB.Image DesktopImage 
      Height          =   480
      Index           =   5
      Left            =   240
      Picture         =   "frmMain.frx":35BA
      Top             =   6000
      Width           =   480
   End
   Begin VB.Label DesktopLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Games"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   5
      Left            =   240
      TabIndex        =   7
      Top             =   6480
      Width           =   720
   End
   Begin VB.Label lblDeskTrash 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Trash Can"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   7560
      Width           =   1005
   End
   Begin VB.Image ImgTrash 
      Height          =   480
      Index           =   0
      Left            =   240
      Picture         =   "frmMain.frx":3E84
      Top             =   7080
      Width           =   480
   End
   Begin VB.Image ImgTrash 
      Height          =   480
      Index           =   1
      Left            =   240
      Picture         =   "frmMain.frx":474E
      Top             =   7080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label DesktopLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Internet  Browser"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   4
      Left            =   240
      TabIndex        =   4
      Top             =   5280
      Width           =   795
   End
   Begin VB.Label DesktopLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Media Player"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   3
      Left            =   240
      TabIndex        =   3
      Top             =   3960
      Width           =   615
   End
   Begin VB.Image DesktopImage 
      Height          =   480
      Index           =   3
      Left            =   240
      Picture         =   "frmMain.frx":5018
      Top             =   3480
      Width           =   480
   End
   Begin VB.Label DesktopLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   2160
      Width           =   450
   End
   Begin VB.Image DesktopImage 
      Height          =   480
      Index           =   1
      Left            =   240
      Picture         =   "frmMain.frx":58E2
      Top             =   1680
      Width           =   480
   End
   Begin VB.Label DesktopLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Explorer  "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   1560
      TabIndex        =   0
      Top             =   3000
      Width           =   855
   End
   Begin VB.Image DesktopImage 
      Height          =   480
      Index           =   0
      Left            =   1560
      Picture         =   "frmMain.frx":61AC
      Top             =   2520
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Time and Date will go here...!"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   11280
      TabIndex        =   6
      Top             =   120
      Width           =   5820
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Height          =   12855
      Left            =   60
      Top             =   60
      Width           =   17175
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      Height          =   12915
      Left            =   30
      Top             =   30
      Width           =   17235
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      Height          =   12975
      Left            =   0
      Top             =   0
      Width           =   17295
   End
   Begin VB.Label DesktopLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Find"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   3000
      Width           =   435
   End
   Begin VB.Image DesktopImage 
      Height          =   480
      Index           =   2
      Left            =   240
      Picture         =   "frmMain.frx":6A76
      Top             =   2520
      Width           =   480
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00E0E0E0&
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   2
      Height          =   12735
      Left            =   120
      Top             =   120
      Width           =   17055
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Height          =   12795
      Left            =   90
      Top             =   90
      Width           =   17115
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      Height          =   855
      Index           =   2
      Left            =   7560
      Shape           =   4  'Rounded Rectangle
      Top             =   -120
      Width           =   975
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      Height          =   855
      Index           =   1
      Left            =   6600
      Shape           =   4  'Rounded Rectangle
      Top             =   -120
      Width           =   855
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      Height          =   855
      Index           =   0
      Left            =   8640
      Shape           =   4  'Rounded Rectangle
      Top             =   -120
      Width           =   855
   End
   Begin VB.Shape Bubble 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   3
      Left            =   5160
      Shape           =   4  'Rounded Rectangle
      Top             =   5040
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      Height          =   285
      Left            =   120
      Top             =   120
      Width           =   17055
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OldX As Long, OldY As Long, IsMoving As Boolean
Dim Selected As Integer, Stuffin As Boolean
Dim ChangingCaption As Boolean

Private Sub ChangeLabelCap_Change()
    ChangingCaption = True
    ChrNum = Len(ChangeLabelCap)
    Select Case ChrNum
        Case 13: ChangeLabelCap.Height = 525: Label1.Height = 525
        Case 26: ChangeLabelCap.Height = 765: Label1.Height = 765
        Case 39: ChangeLabelCap.Height = 1005: Label1.Height = 1005
    End Select
End Sub

Private Sub DesktopImage_Click(Index As Integer)
    Select Case Index
        Case 0: Call ChangeStyle(0): Text1.Text = "used to explore through Your Windows(R) OS, currently..."
        Case 1: Call ChangeStyle(1): Text1.Text = "This option currently does not work"
        Case 2: Call ChangeStyle(2): Text1.Text = "This option currently does not work"
        Case 3: Call ChangeStyle(3): Text1.Text = "used to play your favourite CD's"
        Case 4: Call ChangeStyle(4): Text1.Text = "Used to browse the Internet. Like voteing for this program!"
        Case 5: Call ChangeStyle(5): Text1.Text = "A folder to store your games"
    End Select
    
    Selected = DesktopLabel(Index).Index
    
    If MnuHelp.BorderStyle = 1 Then
        Bubble(0).Top = DesktopImage(Selected).Top + 120
        Bubble(0).Left = DesktopImage(Selected).Left + 360
        Bubble(1).Top = Bubble(0).Top + 120
        Bubble(1).Left = Bubble(0).Left + 240
        Bubble(2).Top = Bubble(1).Top + 120
        Bubble(2).Left = Bubble(1).Left + 240
        Bubble(3).Top = Bubble(2).Top + 120
        Bubble(3).Left = Bubble(2).Left + 240
        Text1.Left = Bubble(3).Left + 120
        Text1.Top = Bubble(3).Top + 240
        For i = 0 To 3
            Bubble(i).Visible = True
        Next i
        Text1.Visible = True
    ElseIf MnuHelp.BorderStyle = 0 Then
        For i = 0 To 3
            Bubble(i).Visible = False
        Next i
        Text1.Visible = False
    End If
End Sub

Private Sub DesktopImage_DblClick(Index As Integer)
    Select Case Index
        Case 0: FrmExplorer.Visible = True
        Case 1:
        Case 2:
        Case 3: frmMed.Visible = True
        Case 4: frmInet.Visible = True
        Case 5
    End Select
    
    Static i As Integer
    i = i + 1
    Load Option1(i)
    
    Option1(i).Left = Option1(i - 1).Left + 1500
    Option1(i).Top = Option1(i - 1).Top
    Option1(i).Caption = DesktopLabel(Selected)
    Option1(i).Visible = True
    TaskbarComponent = DesktopLabel(Selected)
    
    TaskbarComponent = Option1(i).Caption
End Sub

Private Sub DesktopImage_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
        Case 0: OldX = X: OldY = Y: IsMoving = True
        Case 1: OldX = X: OldY = Y: IsMoving = True
        Case 2: OldX = X: OldY = Y: IsMoving = True
        Case 3: OldX = X: OldY = Y: IsMoving = True
        Case 4: OldX = X: OldY = Y: IsMoving = True
        Case 5: OldX = X: OldY = Y: IsMoving = True
    End Select
End Sub

Private Sub DesktopImage_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
        Case 0
            If IsMoving Then
                DesktopImage(Num).Top = DesktopImage(Num).Top - (OldY - Y)
                DesktopImage(Num).Left = DesktopImage(Num).Left - (OldX - X)
        
                DesktopLabel(Num).Top = DesktopLabel(Num).Top - (OldY - Y)
                DesktopLabel(Num).Left = DesktopLabel(Num).Left - (OldX - X)
            End If
            
        Case 1
            If IsMoving Then
                DesktopImage(1).Top = DesktopImage(1).Top - (OldY - Y)
                DesktopImage(1).Left = DesktopImage(1).Left - (OldX - X)
        
                DesktopLabel(1).Top = DesktopLabel(1).Top - (OldY - Y)
                DesktopLabel(1).Left = DesktopLabel(1).Left - (OldX - X)
            End If
        
        Case 2
            If IsMoving Then
                DesktopImage(2).Top = DesktopImage(2).Top - (OldY - Y)
                DesktopImage(2).Left = DesktopImage(2).Left - (OldX - X)
        
                DesktopLabel(2).Top = DesktopLabel(2).Top - (OldY - Y)
                DesktopLabel(2).Left = DesktopLabel(2).Left - (OldX - X)
            End If
        
        Case 3
            If IsMoving Then
                DesktopImage(3).Top = DesktopImage(3).Top - (OldY - Y)
                DesktopImage(3).Left = DesktopImage(3).Left - (OldX - X)
        
                DesktopLabel(3).Top = DesktopLabel(3).Top - (OldY - Y)
                DesktopLabel(3).Left = DesktopLabel(3).Left - (OldX - X)
            End If
            
        Case 4
            If IsMoving Then
                DesktopImage(4).Top = DesktopImage(4).Top - (OldY - Y)
                DesktopImage(4).Left = DesktopImage(4).Left - (OldX - X)
        
                DesktopLabel(4).Top = DesktopLabel(4).Top - (OldY - Y)
                DesktopLabel(4).Left = DesktopLabel(4).Left - (OldX - X)
            End If
        
        Case 5
            If IsMoving Then
                DesktopImage(5).Top = DesktopImage(5).Top - (OldY - Y)
                DesktopImage(5).Left = DesktopImage(5).Left - (OldX - X)
        
                DesktopLabel(5).Top = DesktopLabel(5).Top - (OldY - Y)
                DesktopLabel(5).Left = DesktopLabel(5).Left - (OldX - X)
            End If
        
    End Select
End Sub

Private Sub DesktopImage_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
        Case 0
            IsMoving = False
            SaveSetting "ExcellOS", "Icons", "ImgExplorerPosL", DesktopImage(0).Left
            SaveSetting "ExcellOS", "Icons", "ImgExplorerPosT", DesktopImage(0).Top
            SaveSetting "ExcellOS", "Icons", "LblExplorerPosL", DesktopLabel(0).Left
            SaveSetting "ExcellOS", "Icons", "LblExplorerPosT", DesktopLabel(0).Top
        Case 1
            IsMoving = False
            SaveSetting "ExcellOS", "Icons", "ImgRunPosL", DesktopImage(1).Left
            SaveSetting "ExcellOS", "Icons", "ImgRunPosT", DesktopImage(1).Top
            SaveSetting "ExcellOS", "Icons", "LblRunPosL", DesktopLabel(1).Left
            SaveSetting "ExcellOS", "Icons", "LblRunPosT", DesktopLabel(1).Top
        Case 2
            IsMoving = False
            SaveSetting "ExcellOS", "Icons", "ImgFindPosL", DesktopImage(2).Left
            SaveSetting "ExcellOS", "Icons", "ImgFindPosT", DesktopImage(2).Top
            SaveSetting "ExcellOS", "Icons", "LblFindPosL", DesktopLabel(2).Left
            SaveSetting "ExcellOS", "Icons", "LblFindPosT", DesktopLabel(2).Top
        Case 3
            IsMoving = False
            SaveSetting "ExcellOS", "Icons", "ImgMediaPosL", DesktopImage(3).Left
            SaveSetting "ExcellOS", "Icons", "ImgMediaPosT", DesktopImage(3).Top
            SaveSetting "ExcellOS", "Icons", "LblMediaPosL", DesktopLabel(3).Left
            SaveSetting "ExcellOS", "Icons", "LblMediaPosT", DesktopLabel(3).Top
        Case 4
            IsMoving = False
            SaveSetting "ExcellOS", "Icons", "ImgWebPosL", DesktopImage(4).Left
            SaveSetting "ExcellOS", "Icons", "ImgWebPosT", DesktopImage(4).Top
            SaveSetting "ExcellOS", "Icons", "LblWebPosL", DesktopLabel(4).Left
            SaveSetting "ExcellOS", "Icons", "LblWebPosT", DesktopLabel(4).Top
        Case 5
            IsMoving = False
            SaveSetting "ExcellOS", "Icons", "ImgGamesPosL", DesktopImage(5).Left
            SaveSetting "ExcellOS", "Icons", "ImgGamesPosT", DesktopImage(5).Top
            SaveSetting "ExcellOS", "Icons", "LblGamesPosL", DesktopLabel(5).Left
            SaveSetting "ExcellOS", "Icons", "LblGamesPosT", DesktopLabel(5).Top
    End Select
End Sub

Private Sub DesktopLabel_Click(Index As Integer)
'    ChangeLabelCap.Left = DesktopLabel(Selected).Left
'    ChangeLabelCap.Top = DesktopLabel(Selected).Top
'    ChangeLabelCap.Visible = True
'    DesktopLabel(Selected).Visible = False
'    ChangeLabelCap.SetFocus
End Sub

Private Sub Form_Click()
    For Num = 0 To 5
        DesktopLabel(Num).FontBold = False
    Next Num
    
    For i = 0 To 2
        Label3(i).Visible = False
        Shape9(i).Visible = False
    Next i
    
    If ChangingCaption = True Then
        DesktopLabel(Selected).Caption = ChangeLabelCap.Text
        ChangeLabelCap.Visible = False
        DesktopLabel(Selected).Visible = True
    Else
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandle

    Label1.Caption = Date & " " & Time
    
    For i = 0 To 2
        Label3(i).Visible = False
        Shape9(i).Visible = False
    Next i

    DesktopImage(0).Left = GetSetting("ExcellOS", "Icons", "ImgExplorerPosL")
    DesktopImage(0).Top = GetSetting("ExcellOS", "Icons", "ImgExplorerPosT")
    DesktopLabel(0).Left = GetSetting("ExcellOS", "Icons", "LblExplorerPosL")
    DesktopLabel(0).Top = GetSetting("ExcellOS", "Icons", "LblExplorerPosT")
    
    DesktopImage(1).Left = GetSetting("ExcellOS", "Icons", "ImgRunPosL")
    DesktopImage(1).Top = GetSetting("ExcellOS", "Icons", "ImgRunPosT")
    DesktopLabel(1).Left = GetSetting("ExcellOS", "Icons", "LblRunPosL")
    DesktopLabel(1).Top = GetSetting("ExcellOS", "Icons", "LblRunPosT")
    
    DesktopImage(2).Left = GetSetting("ExcellOS", "Icons", "ImgFindPosL")
    DesktopImage(2).Top = GetSetting("ExcellOS", "Icons", "ImgFindPosT")
    DesktopLabel(2).Left = GetSetting("ExcellOS", "Icons", "LblFindPosL")
    DesktopLabel(2).Top = GetSetting("ExcellOS", "Icons", "LblFindPosT")
    
    DesktopImage(3).Left = GetSetting("ExcellOS", "Icons", "ImgMediaPosL")
    DesktopImage(3).Top = GetSetting("ExcellOS", "Icons", "ImgMediaPosT")
    DesktopLabel(3).Left = GetSetting("ExcellOS", "Icons", "LblMediaPosL")
    DesktopLabel(3).Top = GetSetting("ExcellOS", "Icons", "LblMediaPosT")
    
    DesktopImage(4).Left = GetSetting("ExcellOS", "Icons", "ImgWebPosL")
    DesktopImage(4).Top = GetSetting("ExcellOS", "Icons", "ImgWebPosT")
    DesktopLabel(4).Left = GetSetting("ExcellOS", "Icons", "LblWebPosL")
    DesktopLabel(4).Top = GetSetting("ExcellOS", "Icons", "LblWebPosT")
    
    DesktopImage(5).Left = GetSetting("ExcellOS", "Icons", "ImgGamesPosL")
    DesktopImage(5).Top = GetSetting("ExcellOS", "Icons", "ImgGamesPosT")
    DesktopLabel(5).Left = GetSetting("ExcellOS", "Icons", "LblGamesPosL")
    DesktopLabel(5).Top = GetSetting("ExcellOS", "Icons", "LblGamesPosT")
    
ErrorHandle:
    SaveSetting "ExcellOS", "Icons", "ImgExplorerPosL", "240"
    SaveSetting "ExcellOS", "Icons", "ImgExplorerPosT", "1005"
    SaveSetting "ExcellOS", "Icons", "LblExplorerPosL", "240"
    SaveSetting "ExcellOS", "Icons", "LblExplorerPosT", "1485"
        
    SaveSetting "ExcellOS", "Icons", "ImgRunPosL", "240"
    SaveSetting "ExcellOS", "Icons", "ImgRunPosT", "6525"
    SaveSetting "ExcellOS", "Icons", "LblRunPosL", "240"
    SaveSetting "ExcellOS", "Icons", "LblRunPosT", "7005"
        
    SaveSetting "ExcellOS", "Icons", "ImgFindPosL", "240"
    SaveSetting "ExcellOS", "Icons", "ImgFindPosT", "4290"
    SaveSetting "ExcellOS", "Icons", "LblFindPosL", "240"
    SaveSetting "ExcellOS", "Icons", "LblFindPosT", "4770"
        
    SaveSetting "ExcellOS", "Icons", "ImgMediaPosL", "240"
    SaveSetting "ExcellOS", "Icons", "ImgMediaPosT", "5325"
    SaveSetting "ExcellOS", "Icons", "LblMediaPosL", "240"
    SaveSetting "ExcellOS", "Icons", "LblMediaPosT", "5805"
    
    SaveSetting "ExcellOS", "Icons", "ImgWebPosL", "240"
    SaveSetting "ExcellOS", "Icons", "ImgWebPosT", "2970"
    SaveSetting "ExcellOS", "Icons", "LblWebPosL", "240"
    SaveSetting "ExcellOS", "Icons", "LblWebPosT", "3450"
        
    SaveSetting "ExcellOS", "Icons", "ImgGamesPosL", "240"
    SaveSetting "ExcellOS", "Icons", "ImgGamesPosT", "1920"
    SaveSetting "ExcellOS", "Icons", "LblGamesPosL", "240"
    SaveSetting "ExcellOS", "Icons", "LblGamesPosT", "2400"
    
    SaveSetting "ExcellOS", "Misc", "TbarBackColour", "14737632"
    SaveSetting "ExcellOS", "Misc", "TbarForeColour", "12582912"
    SaveSetting "ExcellOS", "Misc", "SSFS", "78"
    SaveSetting "ExcellOS", "Misc", "Sounds", "1"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    For i = 0 To 2
        Label3(i).FontUnderline = False
        Label3(i).ForeColor = &H0&
    Next i
    
    Label5.Caption = "0"
    Label6.Caption = "0"
End Sub

Private Sub Image2_Click()
    For Num = 0 To 5
        DesktopLabel(Num).FontBold = False
    Next Num
    
    For i = 0 To 2
        Label3(i).Visible = False
        Shape9(i).Visible = False
    Next i
    
    If ChangingCaption = True Then
        DesktopLabel(Selected).Caption = ChangeLabelCap.Text
        ChangeLabelCap.Visible = False
        DesktopLabel(Selected).Visible = True
    Else
    End If
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    For i = 0 To 2
        Label3(i).FontUnderline = False
        Label3(i).ForeColor = &H0&
    Next i
    
    Label5.Caption = "0"
    Label6.Caption = "0"
End Sub

Private Sub ImgDoc_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Selected = ImgDoc(Index).Index
    Select Case Index
        Case Selected: OldX = X: OldY = Y: IsMoving = True
    End Select
End Sub

Private Sub ImgDoc_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Selected = ImgDoc(Index).Index
    Select Case Index
        Case Selected
            If IsMoving Then
                ImgDoc(Selected).Top = ImgDoc(Selected).Top - (OldY - Y)
                ImgDoc(Selected).Left = ImgDoc(Selected).Left - (OldX - X)
        
                LblDoc(Selected).Top = LblDoc(Selected).Top - (OldY - Y)
                LblDoc(Selected).Left = LblDoc(Selected).Left - (OldX - X)
            End If
    End Select
End Sub

Private Sub ImgDoc_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Selected = ImgDoc(Index).Index
    Select Case Index
        Case Selected: IsMoving = False
    End Select
End Sub

Private Sub ImgFolder_Click(Index As Integer)
    Selected = ImgFolder(Index).Index

    For i = 0 To 2
        Label3(i).Visible = True
        Shape9(i).Visible = True
    Next i
End Sub

Private Sub ImgFolder_DblClick(Index As Integer)
    FrmFolder.Visible = True
End Sub

Private Sub ImgFolder_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Selected = ImgFolder(Index).Index
    Select Case Index
        Case Selected: OldX = X: OldY = Y: IsMoving = True
    End Select
End Sub

Private Sub ImgFolder_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Selected = ImgFolder(Index).Index
    Select Case Index
        Case Selected
            If IsMoving Then
                ImgFolder(Selected).Top = ImgFolder(Selected).Top - (OldY - Y)
                ImgFolder(Selected).Left = ImgFolder(Selected).Left - (OldX - X)
        
                LblFolder(Selected).Top = LblFolder(Selected).Top - (OldY - Y)
                LblFolder(Selected).Left = LblFolder(Selected).Left - (OldX - X)
            End If
    End Select
End Sub

Private Sub ImgFolder_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Selected = ImgFolder(Index).Index
    Select Case Index
        Case Selected: IsMoving = False
    End Select
End Sub

Private Sub imgsd_Click()
    End
End Sub

Private Sub ImgTrash_DblClick(Index As Integer)
    If Stuffin = True Then
        TrashContents.Visible = True
    End If
End Sub

Private Sub Label3_Click(Index As Integer)
    If Label3(Index).Index = 0 Then
        If FrmOptions.Check1.Value = Checked Then PlaySound (App.Path + "\select.wav") Else
        MsgVar = MsgBox("Are you sure you want to delete " & Chr(34) & DesktopLabel(Selected).Caption & Chr(34) & " to the trash?", vbYesNo + vbQuestion, "Delete Confermation")
        Select Case MsgVar
            Case vbYes
                DesktopLabel(Selected).Visible = False
                DesktopImage(Selected).Visible = False
                If ImgTrash(0).Visible = True Then
                    ImgTrash(0).Visible = False
                    ImgTrash(1).Visible = True
                End If
                Stuffin = True
            Case vbNo
                Cancel = Not ReadyToDelete
        End Select
    ElseIf Label3(Index).Index = 1 Then
        If FrmOptions.Check1.Value = Checked Then PlaySound (App.Path + "\select.wav") Else
    ElseIf Label3(Index).Index = 2 Then
        If FrmOptions.Check1.Value = Checked Then PlaySound (App.Path + "\select.wav") Else
        ChangeLabelCap.Left = DesktopLabel(Selected).Left
        ChangeLabelCap.Top = DesktopLabel(Selected).Top
        ChangeLabelCap.Visible = True
        DesktopLabel(Selected).Visible = False
        ChangeLabelCap.SetFocus
    End If
End Sub

Private Sub Label3_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
        Case 0: ChangeFormat (0)
        Case 1: ChangeFormat (1)
        Case 2: ChangeFormat (2)
    End Select
End Sub

Private Sub Mnu1_Click(Index As Integer)
    Select Case Index
    Case 0
            If Mnu1(0).BorderStyle = 1 Then
                Mnu1(0).BorderStyle = 0
                Frame2.Visible = False
                Line6.Visible = False
                Line7.Visible = False
                Line9.Visible = False
                Line10.Visible = False
                Line11.Visible = False
                Line12.Visible = False
                Line13.Visible = False
                Shape11.Visible = False
                Shape12.Visible = False
                For i = 1 To 5
                    Mnu1(i).Visible = False
                Next i
            ElseIf Mnu1(0).BorderStyle = 0 Then
                Mnu1(0).BorderStyle = 1
                Frame2.Visible = True
                Line6.Visible = True
                Line7.Visible = True
                Line9.Visible = True
                Line10.Visible = True
                Line11.Visible = True
                Line12.Visible = True
                Line13.Visible = True
                Shape11.Visible = True
                Shape12.Visible = True
                For i = 1 To 5
                    Mnu1(i).Visible = True
                Next i
            End If
            
        Case 1:
        Case 2:
        Case 3:
        Case 4: FrmOptions.Visible = True
        Case 5: EndApp
    End Select
End Sub

Private Sub Mnu1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    For i = 1 To 5
        Mnu1(i).BackColor = &HFFFFFF
    Next i
    
    If Mnu1(Index).Index <> 0 Then
        Mnu1(Index).BackColor = vbRed
    End If
    
    Select Case Index
        Case 0: SetNormal
        Case 1
            Frame5.Visible = True
            Frame4.Visible = False
            Frame3.Visible = False
            Shape11.Visible = False
            Shape12.Visible = False
        Case 2
            Frame3.Visible = True
            Frame4.Visible = False
            Frame5.Visible = False
            Shape11.Visible = False
            Shape12.Visible = False
        Case 3
            Frame4.Visible = True
            Frame3.Visible = False
            Frame5.Visible = False
            Shape11.Visible = False
            Shape12.Visible = False
            
        Case 4: SetNormal
        Case 5: SetNormal
    End Select
End Sub

Private Sub MnuAlign_Click(Index As Integer)
    If MnuAlign(Index).Index = 1 Then
        For i = 0 To 5
            DesktopImage(i).Left = Shape6.Left
            DesktopLabel(i).Left = Shape6.Left
        Next i
    ElseIf MnuAlign(Index).Index = 2 Then
        For i = 0 To 5
            DesktopImage(i).Left = Shape13.Left
            DesktopLabel(i).Left = Shape13.Left
        Next i
    End If
End Sub

Private Sub MnuAlign_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    For i = 1 To 2
        MnuAlign(i).BackColor = &HFFFFFF
    Next i
    
    MnuAlign(Index).BackColor = vbRed
End Sub

Private Sub MnuClock_Click(Index As Integer)
    Select Case Index
        Case 0
            Label1.Caption = Time
            TimerTime.Enabled = True
            Timer1.Enabled = False
        Case 1
            Label1.Caption = Date & " " & Time
            Timer1.Enabled = True
            TimerTime.Enabled = False
        Case 2
            Label1.Caption = Date
            Timer1.Enabled = False
            TimerTime.Enabled = False
    End Select
End Sub

Private Sub MnuClock_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    For i = 0 To 2
        MnuClock(i).BackColor = &HFFFFFF
    Next i
    
    MnuClock(Index).BackColor = vbRed
End Sub

Private Sub MnuHelp_Click()
    If MnuHelp.BorderStyle = 1 Then
        MnuHelp.BorderStyle = 0
    ElseIf MnuHelp.BorderStyle = 0 Then
        MnuHelp.BorderStyle = 1
        Text1.Text = "Click a desktop icon to make this window disappear"
    End If
End Sub

Private Sub MnuNew_Click(Index As Integer)
    Select Case Index
        Case 1
            Static i As Integer
            i = i + 1
            Load ImgFolder(i)
            Load LblFolder(i)
            
            ImgFolder(i).Left = ImgFolder(i - 1).Left + 200
            ImgFolder(i).Top = ImgFolder(i - 1).Top + 600
        
            LblFolder(i).Left = LblFolder(i - 1).Left + 200
            LblFolder(i).Top = LblFolder(i - 1).Top + 600
        
            LblFolder(i).Caption = "New Folder #" & i
            ImgFolder(i).Visible = True
            LblFolder(i).Visible = True
            
        Case 2
            Static J As Integer
            J = J + 1
            Load ImgDoc(J)
            Load LblDoc(J)
            
            ImgDoc(J).Left = ImgDoc(J - 1).Left + 200
            ImgDoc(J).Top = ImgDoc(J - 1).Top + 600
        
            LblDoc(J).Left = LblDoc(J - 1).Left + 200
            LblDoc(J).Top = LblDoc(J - 1).Top + 600
        
            LblDoc(J).Caption = "New doc #" & J
            ImgDoc(J).Visible = True
            LblDoc(J).Visible = True
    End Select
End Sub

Private Sub MnuNew_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    For i = 1 To 2
        MnuNew(i).BackColor = &HFFFFFF
    Next i
    
    MnuNew(Index).BackColor = vbRed
End Sub

Private Sub Option1_Click(Index As Integer)
    'Option1(i).Left = Option1(i - 1).Left + 1500
    'Option1(i).Top = Option1(i - 1).Top
    'Option1(i).Caption = DesktopLabel(Selected)
    'Option1(i).Visible = True
    'TaskbarComponent = DesktopLabel(Selected)
End Sub

Private Sub ChangeFormat(SelectedL As Integer)
    Select Case SelectedL
        Case 0
            Label3(SelectedL).FontUnderline = True
            Label3(SelectedL).ForeColor = &HFF0000
            Num = SelectedL + 1
            Label3(Num).FontUnderline = False
            Label3(Num).ForeColor = &H0&
            Num = Num + 1
            Label3(Num).FontUnderline = False
            Label3(Num).ForeColor = &H0&
        Case 1
            Label3(SelectedL).FontUnderline = True
            Label3(SelectedL).ForeColor = &HFF0000
            Num = SelectedL - 1
            Label3(Num).FontUnderline = False
            Label3(Num).ForeColor = &H0&
            Num = Num + 2
            Label3(Num).FontUnderline = False
            Label3(Num).ForeColor = &H0&
        Case 2
            Label3(SelectedL).FontUnderline = True
            Label3(SelectedL).ForeColor = &HFF0000
            Num = SelectedL - 1
            Label3(Num).FontUnderline = False
            Label3(Num).ForeColor = &H0&
            Num = Num - 1
            Label3(Num).FontUnderline = False
            Label3(Num).ForeColor = &H0&
    End Select
End Sub

Private Sub ChangeStyle(Num As Integer)
    For i = 0 To 5
        DesktopLabel(i).BorderStyle = 0
        DesktopLabel(i).FontBold = False
    Next i
            
    If FrmOptions.Check2.Value = Checked Then DesktopLabel(Num).FontBold = True Else
    If FrmOptions.Check1.Value = Checked Then PlaySound (App.Path + "\select.wav") Else

    For i = 0 To 2
        Label3(i).Visible = True
        Shape9(i).Visible = True
    Next i
End Sub

Public Sub PlaySound(strFileName As String)
    sndPlaySound strFileName, 1
End Sub

Private Sub Timer1_Timer()
    Label1.Caption = Date & " " & Time
End Sub

Private Sub Timer2_Timer()
    Label6.Caption = Label6.Caption + 1
    If Label6.Caption = "60" Then
        Label6.Caption = "0"
        Label5.Caption = Label5.Caption + 1
        If Label5.Caption = FrmOptions.Text4.Text Then
            FrmSS.Visible = True
        End If
    End If
End Sub

Private Sub TimerTime_Timer()
    Label1.Caption = Time
End Sub

Private Sub SetNormal()
    Frame4.Visible = False
    Frame3.Visible = False
    Frame5.Visible = False
    Shape11.Visible = True
    Shape12.Visible = True
End Sub
