VERSION 5.00
Begin VB.Form FrmSS 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   12960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17280
   LinkTopic       =   "Form1"
   ScaleHeight     =   12960
   ScaleWidth      =   17280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   720
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Interval        =   3
      Left            =   240
      Top             =   120
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      Left            =   6120
      TabIndex        =   0
      Top             =   5640
      Width           =   4740
   End
End
Attribute VB_Name = "FrmSS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ScreenWidth As Integer
Dim ScreenHeight As Integer

Private Sub Form_Click()
    If Timer1.Enabled = True Then
        Timer1.Enabled = False
        Timer2.Enabled = True
    ElseIf Timer2.Enabled = True Then
        Timer2.Enabled = False
        Timer1.Enabled = True
    End If
    
End Sub

Private Sub Form_Load()
    ScreenWidth = FrmSS.Width
    ScreenHeight = FrmSS.Height
    
    Label1.FontSize = FrmOptions.Text2.Text
    
    Label1.Caption = FrmOptions.Text3.Text
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Unload Me
End Sub

Private Sub Label1_Click()
Unload Me
End Sub

Private Sub Timer1_Timer()
    Label1.Move Label1.Left - 50
    If (Label1.Left + Label1.Width) < 0 Then Label1.Left = Me.ScaleWidth + 10
End Sub

Private Sub Timer2_Timer()
    A = Int(Rnd * ScreenHeight)
    B = Int(Rnd * ScreenWidth)
    
    Label1.Top = A
    Label1.Left = B
End Sub
