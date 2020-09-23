VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmLoading 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5040
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6000
      Top             =   240
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   255
      Left            =   1800
      TabIndex        =   0
      Top             =   3600
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   $"FrmLoading.frx":0000
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   2
      Top             =   4080
      Width           =   6495
   End
   Begin VB.Image Image1 
      Height          =   2685
      Left            =   840
      Picture         =   "FrmLoading.frx":0126
      Top             =   360
      Width           =   5190
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Loading: Object Goes Here!!"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   3240
      Visible         =   0   'False
      Width           =   6495
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00C0C0FF&
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   2
      Height          =   4815
      Left            =   120
      Top             =   120
      Width           =   6735
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Height          =   4860
      Left            =   105
      Top             =   105
      Width           =   6780
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Height          =   4920
      Left            =   75
      Top             =   75
      Width           =   6840
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      Height          =   4965
      Left            =   45
      Top             =   45
      Width           =   6900
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   5040
      Left            =   0
      Top             =   0
      Width           =   6975
   End
End
Attribute VB_Name = "FrmLoading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    frmMain.Visible = True
    
    For i = 0 To 5
        frmMain.DesktopImage(i).Visible = False
        frmMain.DesktopLabel(i).Visible = False
    Next i
    
    For i = 0 To 1
        frmMain.ImgTrash(i).Visible = False
    Next i
    
    For i = 0 To 2
        frmMain.Shape9(i).Visible = False
        frmMain.Label3(i).Visible = False
    Next i
    
    frmMain.Frame1.Visible = False
    frmMain.lblDeskTrash.Visible = False
    frmMain.Enabled = False
    
    PB1.Visible = True
    Label1.Visible = True
    Timer1.Enabled = True
    
End Sub

Private Sub Timer1_Timer()
    PB1 = PB1 + 1
    
    Select Case PB1.Value
        Case 1: Label1.Caption = "Loading: Sounds"
        Case 10: Label1.Caption = "Retrieving registry settings"
        Case 20: Label1.Caption = "Loading: Icons"
            
        Case 21: frmMain.DesktopImage(0).Visible = True: frmMain.DesktopLabel(0).Visible = True
        Case 23: frmMain.DesktopImage(1).Visible = True: frmMain.DesktopLabel(1).Visible = True
        Case 24: frmMain.DesktopImage(2).Visible = True: frmMain.DesktopLabel(2).Visible = True
        Case 26: frmMain.DesktopImage(3).Visible = True: frmMain.DesktopLabel(3).Visible = True
        Case 27: frmMain.DesktopImage(4).Visible = True: frmMain.DesktopLabel(4).Visible = True
        Case 29: frmMain.DesktopImage(5).Visible = True: frmMain.DesktopLabel(5).Visible = True
    
        Case 32: Label1.Caption = "Loading: TaskHandler"
        Case 35: frmMain.Frame1.Visible = True:
        Case 37: frmMain.Option1(0).Visible = True
        
        Case 100
            Timer1.Enabled = False
            Me.Visible = False
            frmMain.Enabled = True
            Timer1.Enabled = True
            Timer1.Interval = "2000"
            frmMain.Visible = True
            Timer1.Enabled = False
    End Select
End Sub
