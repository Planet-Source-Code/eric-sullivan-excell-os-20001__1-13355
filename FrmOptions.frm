VERSION 5.00
Begin VB.Form FrmOptions 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   ScaleHeight     =   3840
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   240
      TabIndex        =   118
      Top             =   480
      Width           =   6255
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   2655
         Left            =   0
         TabIndex        =   172
         Top             =   0
         Width           =   6255
         Begin VB.Label Label11 
            BackColor       =   &H00FFFFFF&
            Caption         =   $"FrmOptions.frx":0000
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2295
            Left            =   120
            TabIndex        =   173
            Top             =   240
            Width           =   6015
         End
         Begin VB.Shape Shape9 
            Height          =   2535
            Left            =   0
            Top             =   120
            Width           =   6255
         End
      End
      Begin VB.TextBox Text1 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4080
         TabIndex        =   171
         Text            =   "60"
         Top             =   2280
         Width           =   375
      End
      Begin VB.CommandButton Command7 
         Caption         =   "q"
         BeginProperty Font 
            Name            =   "Wingdings 3"
            Size            =   5.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   119
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Change background colour every        minutes"
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
         Left            =   120
         TabIndex        =   170
         Top             =   2280
         Width           =   6015
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Titlebar Colour:"
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
         Height          =   255
         Left            =   600
         TabIndex        =   169
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label BGColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   168
         Top             =   600
         Width           =   255
      End
      Begin VB.Label BGColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   167
         Top             =   600
         Width           =   255
      End
      Begin VB.Label BGColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   166
         Top             =   600
         Width           =   255
      End
      Begin VB.Label BGColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   1320
         TabIndex        =   165
         Top             =   600
         Width           =   255
      End
      Begin VB.Label BGColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   1560
         TabIndex        =   164
         Top             =   600
         Width           =   255
      End
      Begin VB.Label BGColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   1800
         TabIndex        =   163
         Top             =   600
         Width           =   255
      End
      Begin VB.Label BGarColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   2040
         TabIndex        =   162
         Top             =   600
         Width           =   255
      End
      Begin VB.Label BGColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   2280
         TabIndex        =   161
         Top             =   600
         Width           =   255
      End
      Begin VB.Label BGColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   8
         Left            =   600
         TabIndex        =   160
         Top             =   840
         Width           =   255
      End
      Begin VB.Label BGColours 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   9
         Left            =   840
         TabIndex        =   159
         Top             =   840
         Width           =   255
      End
      Begin VB.Label BGColours 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   10
         Left            =   1080
         TabIndex        =   158
         Top             =   840
         Width           =   255
      End
      Begin VB.Label BGColours 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   11
         Left            =   1320
         TabIndex        =   157
         Top             =   840
         Width           =   255
      End
      Begin VB.Label BGColours 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   12
         Left            =   1560
         TabIndex        =   156
         Top             =   840
         Width           =   255
      End
      Begin VB.Label BGColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   13
         Left            =   1800
         TabIndex        =   155
         Top             =   840
         Width           =   255
      End
      Begin VB.Label BGColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   14
         Left            =   2040
         TabIndex        =   154
         Top             =   840
         Width           =   255
      End
      Begin VB.Label BGColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF80FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   15
         Left            =   2280
         TabIndex        =   153
         Top             =   840
         Width           =   255
      End
      Begin VB.Label BGColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   16
         Left            =   600
         TabIndex        =   152
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label BGColours 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   17
         Left            =   840
         TabIndex        =   151
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label BGColours 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   18
         Left            =   1080
         TabIndex        =   150
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label BGColours 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   19
         Left            =   1320
         TabIndex        =   149
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label BGColours 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   20
         Left            =   1560
         TabIndex        =   148
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label BGColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   21
         Left            =   1800
         TabIndex        =   147
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label BGColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   22
         Left            =   2040
         TabIndex        =   146
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label BGColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF00FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   23
         Left            =   2280
         TabIndex        =   145
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label BGColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   24
         Left            =   600
         TabIndex        =   144
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label BGColours 
         Appearance      =   0  'Flat
         BackColor       =   &H000000C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   25
         Left            =   840
         TabIndex        =   143
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label BGColours 
         Appearance      =   0  'Flat
         BackColor       =   &H000040C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   26
         Left            =   1080
         TabIndex        =   142
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label BGColours 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   27
         Left            =   1320
         TabIndex        =   141
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label BGColours 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   28
         Left            =   1560
         TabIndex        =   140
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label BGColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   29
         Left            =   1800
         TabIndex        =   139
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label BGColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   30
         Left            =   2040
         TabIndex        =   138
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label BGColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00C000C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   31
         Left            =   2280
         TabIndex        =   137
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label BGColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   32
         Left            =   600
         TabIndex        =   136
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label BGColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   33
         Left            =   840
         TabIndex        =   135
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label BGColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00004080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   34
         Left            =   1080
         TabIndex        =   134
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label BGColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00008080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   35
         Left            =   1320
         TabIndex        =   133
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label BGColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   36
         Left            =   1560
         TabIndex        =   132
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label BGColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   37
         Left            =   1800
         TabIndex        =   131
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label BGColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   38
         Left            =   2040
         TabIndex        =   130
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label BGColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   39
         Left            =   2280
         TabIndex        =   129
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label BGColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   40
         Left            =   600
         TabIndex        =   128
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label BGColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00000040&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   41
         Left            =   840
         TabIndex        =   127
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label BGColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00404080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   42
         Left            =   1080
         TabIndex        =   126
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label BGColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00004040&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   43
         Left            =   1320
         TabIndex        =   125
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label BGColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   44
         Left            =   1560
         TabIndex        =   124
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label BGColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00404000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   45
         Left            =   1800
         TabIndex        =   123
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label BGColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   46
         Left            =   2040
         TabIndex        =   122
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label BGColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00400040&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   47
         Left            =   2280
         TabIndex        =   121
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Wingdings 3"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2040
         TabIndex        =   120
         Top             =   240
         Width           =   495
      End
      Begin VB.Line Line5 
         X1              =   120
         X2              =   6120
         Y1              =   2160
         Y2              =   2160
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   6255
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Bold selected desktop icon"
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
         Height          =   255
         Left            =   1920
         TabIndex        =   109
         Top             =   2280
         Width           =   2655
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Play Sounds"
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
         Height          =   255
         Left            =   120
         TabIndex        =   108
         Top             =   2280
         Width           =   1455
      End
      Begin VB.CommandButton Command5 
         Caption         =   "q"
         BeginProperty Font 
            Name            =   "Wingdings 3"
            Size            =   5.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5280
         TabIndex        =   59
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton Command4 
         Caption         =   "q"
         BeginProperty Font 
            Name            =   "Wingdings 3"
            Size            =   5.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   55
         Top             =   240
         Width           =   255
      End
      Begin VB.Line Line3 
         X1              =   1680
         X2              =   1680
         Y1              =   2280
         Y2              =   2520
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   6120
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Label TBarTcolour 
         Appearance      =   0  'Flat
         BackColor       =   &H00400040&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   47
         Left            =   5280
         TabIndex        =   107
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label TBarTcolour 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   46
         Left            =   5040
         TabIndex        =   106
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label TBarTcolour 
         Appearance      =   0  'Flat
         BackColor       =   &H00404000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   45
         Left            =   4800
         TabIndex        =   105
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label TBarTcolour 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   44
         Left            =   4560
         TabIndex        =   104
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label TBarTcolour 
         Appearance      =   0  'Flat
         BackColor       =   &H00004040&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   43
         Left            =   4320
         TabIndex        =   103
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label TBarTcolour 
         Appearance      =   0  'Flat
         BackColor       =   &H00404080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   42
         Left            =   4080
         TabIndex        =   102
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label TBarTcolour 
         Appearance      =   0  'Flat
         BackColor       =   &H00000040&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   41
         Left            =   3840
         TabIndex        =   101
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label TBarTcolour 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   40
         Left            =   3600
         TabIndex        =   100
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label TBarTcolour 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   39
         Left            =   5280
         TabIndex        =   99
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label TBarTcolour 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   38
         Left            =   5040
         TabIndex        =   98
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label TBarTcolour 
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   37
         Left            =   4800
         TabIndex        =   97
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label TBarTcolour 
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   36
         Left            =   4560
         TabIndex        =   96
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label TBarTcolour 
         Appearance      =   0  'Flat
         BackColor       =   &H00008080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   35
         Left            =   4320
         TabIndex        =   95
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label TBarTcolour 
         Appearance      =   0  'Flat
         BackColor       =   &H00004080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   34
         Left            =   4080
         TabIndex        =   94
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label TBarTcolour 
         Appearance      =   0  'Flat
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   33
         Left            =   3840
         TabIndex        =   93
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label TBarTcolour 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   32
         Left            =   3600
         TabIndex        =   92
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label TBarTcolour 
         Appearance      =   0  'Flat
         BackColor       =   &H00C000C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   31
         Left            =   5280
         TabIndex        =   91
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label TBarTcolour 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   30
         Left            =   5040
         TabIndex        =   90
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label TBarTcolour 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   29
         Left            =   4800
         TabIndex        =   89
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label TBarTcolour 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   28
         Left            =   4560
         TabIndex        =   88
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label TBarTcolour 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   27
         Left            =   4320
         TabIndex        =   87
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label TBarTcolour 
         Appearance      =   0  'Flat
         BackColor       =   &H000040C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   26
         Left            =   4080
         TabIndex        =   86
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label TBarTcolour 
         Appearance      =   0  'Flat
         BackColor       =   &H000000C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   25
         Left            =   3840
         TabIndex        =   85
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label TBarTcolour 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   24
         Left            =   3600
         TabIndex        =   84
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label TBarTcolour 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF00FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   23
         Left            =   5280
         TabIndex        =   83
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label TBarTcolour 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   22
         Left            =   5040
         TabIndex        =   82
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label TBarTcolour 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   21
         Left            =   4800
         TabIndex        =   81
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label TBarTcolour 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   20
         Left            =   4560
         TabIndex        =   80
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label TBarTcolour 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   19
         Left            =   4320
         TabIndex        =   79
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label TBarTcolour 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   18
         Left            =   4080
         TabIndex        =   78
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label TBarTcolour 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   17
         Left            =   3840
         TabIndex        =   77
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label TBarTcolour 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   16
         Left            =   3600
         TabIndex        =   76
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label TBarTcolour 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF80FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   15
         Left            =   5280
         TabIndex        =   75
         Top             =   840
         Width           =   255
      End
      Begin VB.Label TBarTcolour 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   14
         Left            =   5040
         TabIndex        =   74
         Top             =   840
         Width           =   255
      End
      Begin VB.Label TBarTcolour 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   13
         Left            =   4800
         TabIndex        =   73
         Top             =   840
         Width           =   255
      End
      Begin VB.Label TBarTcolour 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   12
         Left            =   4560
         TabIndex        =   72
         Top             =   840
         Width           =   255
      End
      Begin VB.Label TBarTcolour 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   11
         Left            =   4320
         TabIndex        =   71
         Top             =   840
         Width           =   255
      End
      Begin VB.Label TBarTcolour 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   10
         Left            =   4080
         TabIndex        =   70
         Top             =   840
         Width           =   255
      End
      Begin VB.Label TBarTcolour 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   9
         Left            =   3840
         TabIndex        =   69
         Top             =   840
         Width           =   255
      End
      Begin VB.Label TBarTcolour 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   8
         Left            =   3600
         TabIndex        =   68
         Top             =   840
         Width           =   255
      End
      Begin VB.Label TBarTcolour 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   5280
         TabIndex        =   67
         Top             =   600
         Width           =   255
      End
      Begin VB.Label TBarTcolour 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   5040
         TabIndex        =   66
         Top             =   600
         Width           =   255
      End
      Begin VB.Label TBarTcolour 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   4800
         TabIndex        =   65
         Top             =   600
         Width           =   255
      End
      Begin VB.Label TBarTcolour 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   4560
         TabIndex        =   64
         Top             =   600
         Width           =   255
      End
      Begin VB.Label TBarTcolour 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   4320
         TabIndex        =   63
         Top             =   600
         Width           =   255
      End
      Begin VB.Label TBarTcolour 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   4080
         TabIndex        =   62
         Top             =   600
         Width           =   255
      End
      Begin VB.Label TBarTcolour 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   3840
         TabIndex        =   61
         Top             =   600
         Width           =   255
      End
      Begin VB.Label TBarTcolour 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   3600
         TabIndex        =   60
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Wingdings 3"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5040
         TabIndex        =   58
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Titlebar text colour:"
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
         Height          =   255
         Left            =   3240
         TabIndex        =   57
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Wingdings 3"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2040
         TabIndex        =   54
         Top             =   240
         Width           =   495
      End
      Begin VB.Label TBarColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00400040&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   47
         Left            =   2280
         TabIndex        =   53
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label TBarColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   46
         Left            =   2040
         TabIndex        =   52
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label TBarColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00404000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   45
         Left            =   1800
         TabIndex        =   51
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label TBarColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   44
         Left            =   1560
         TabIndex        =   50
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label TBarColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00004040&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   43
         Left            =   1320
         TabIndex        =   49
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label TBarColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00404080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   42
         Left            =   1080
         TabIndex        =   48
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label TBarColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00000040&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   41
         Left            =   840
         TabIndex        =   47
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label TBarColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   40
         Left            =   600
         TabIndex        =   46
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label TBarColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   39
         Left            =   2280
         TabIndex        =   45
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label TBarColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   38
         Left            =   2040
         TabIndex        =   44
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label TBarColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   37
         Left            =   1800
         TabIndex        =   43
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label TBarColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   36
         Left            =   1560
         TabIndex        =   42
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label TBarColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00008080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   35
         Left            =   1320
         TabIndex        =   41
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label TBarColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00004080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   34
         Left            =   1080
         TabIndex        =   40
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label TBarColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   33
         Left            =   840
         TabIndex        =   39
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label TBarColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   32
         Left            =   600
         TabIndex        =   38
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label TBarColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00C000C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   31
         Left            =   2280
         TabIndex        =   37
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label TBarColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   30
         Left            =   2040
         TabIndex        =   36
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label TBarColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   29
         Left            =   1800
         TabIndex        =   35
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label TBarColours 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   28
         Left            =   1560
         TabIndex        =   34
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label TBarColours 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   27
         Left            =   1320
         TabIndex        =   33
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label TBarColours 
         Appearance      =   0  'Flat
         BackColor       =   &H000040C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   26
         Left            =   1080
         TabIndex        =   32
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label TBarColours 
         Appearance      =   0  'Flat
         BackColor       =   &H000000C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   25
         Left            =   840
         TabIndex        =   31
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label TBarColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   24
         Left            =   600
         TabIndex        =   30
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label TBarColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF00FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   23
         Left            =   2280
         TabIndex        =   29
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label TBarColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   22
         Left            =   2040
         TabIndex        =   28
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label TBarColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   21
         Left            =   1800
         TabIndex        =   27
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label TBarColours 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   20
         Left            =   1560
         TabIndex        =   26
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label TBarColours 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   19
         Left            =   1320
         TabIndex        =   25
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label TBarColours 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   18
         Left            =   1080
         TabIndex        =   24
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label TBarColours 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   17
         Left            =   840
         TabIndex        =   23
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label TBarColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   16
         Left            =   600
         TabIndex        =   22
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label TBarColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF80FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   15
         Left            =   2280
         TabIndex        =   21
         Top             =   840
         Width           =   255
      End
      Begin VB.Label TBarColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   14
         Left            =   2040
         TabIndex        =   20
         Top             =   840
         Width           =   255
      End
      Begin VB.Label TBarColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   13
         Left            =   1800
         TabIndex        =   19
         Top             =   840
         Width           =   255
      End
      Begin VB.Label TBarColours 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   12
         Left            =   1560
         TabIndex        =   18
         Top             =   840
         Width           =   255
      End
      Begin VB.Label TBarColours 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   11
         Left            =   1320
         TabIndex        =   17
         Top             =   840
         Width           =   255
      End
      Begin VB.Label TBarColours 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   10
         Left            =   1080
         TabIndex        =   16
         Top             =   840
         Width           =   255
      End
      Begin VB.Label TBarColours 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   9
         Left            =   840
         TabIndex        =   15
         Top             =   840
         Width           =   255
      End
      Begin VB.Label TBarColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   8
         Left            =   600
         TabIndex        =   14
         Top             =   840
         Width           =   255
      End
      Begin VB.Label TBarColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   2280
         TabIndex        =   13
         Top             =   600
         Width           =   255
      End
      Begin VB.Label TBarColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   2040
         TabIndex        =   12
         Top             =   600
         Width           =   255
      End
      Begin VB.Label TBarColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   1800
         TabIndex        =   11
         Top             =   600
         Width           =   255
      End
      Begin VB.Label TBarColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   1560
         TabIndex        =   10
         Top             =   600
         Width           =   255
      End
      Begin VB.Label TBarColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   1320
         TabIndex        =   9
         Top             =   600
         Width           =   255
      End
      Begin VB.Label TBarColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   8
         Top             =   600
         Width           =   255
      End
      Begin VB.Label TBarColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   7
         Top             =   600
         Width           =   255
      End
      Begin VB.Label TBarColours 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   6
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Titlebar Colour:"
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
         Height          =   255
         Left            =   600
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
      Begin VB.Line Line1 
         X1              =   3000
         X2              =   3000
         Y1              =   240
         Y2              =   2160
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   240
      TabIndex        =   56
      Top             =   480
      Visible         =   0   'False
      Width           =   6255
      Begin VB.CheckBox Check4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Randomly place text"
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
         Left            =   120
         TabIndex        =   117
         Top             =   1080
         Width           =   2055
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Scroll text"
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
         Left            =   120
         TabIndex        =   116
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         TabIndex        =   115
         Text            =   "I Love Excell OS!"
         Top             =   710
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         TabIndex        =   113
         Text            =   "72"
         Top             =   260
         Width           =   735
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2280
         TabIndex        =   112
         Text            =   "5"
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Screen saver text:"
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
         Left            =   120
         TabIndex        =   114
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Start screen saver after       minutes"
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
         Left            =   120
         TabIndex        =   111
         Top             =   480
         Width           =   4815
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Font Size:"
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
         Height          =   255
         Left            =   120
         TabIndex        =   110
         Top             =   270
         Width           =   975
      End
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "General"
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
      Height          =   255
      Index           =   2
      Left            =   3120
      MouseIcon       =   "FrmOptions.frx":008C
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   3210
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Screen Saver"
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
      Height          =   255
      Index           =   1
      Left            =   1680
      MouseIcon       =   "FrmOptions.frx":0396
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   3210
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Background"
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
      Height          =   255
      Index           =   0
      Left            =   360
      MouseIcon       =   "FrmOptions.frx":06A0
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   3210
      Width           =   1095
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   2
      Height          =   3615
      Left            =   120
      Top             =   120
      Width           =   6615
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Height          =   3660
      Left            =   105
      Top             =   105
      Width           =   6660
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Height          =   3720
      Left            =   75
      Top             =   75
      Width           =   6720
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      Height          =   3765
      Left            =   45
      Top             =   45
      Width           =   6780
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   3840
      Left            =   0
      Top             =   0
      Width           =   6855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "  ExcellOS Settings"
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
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   1560
      Shape           =   4  'Rounded Rectangle
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   3000
      Shape           =   4  'Rounded Rectangle
      Top             =   2760
      Width           =   1335
   End
End
Attribute VB_Name = "FrmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BGTimer As Integer

Private Sub Check1_Click()
    Select Case Check1.Value
        Case 0: SoundsEnabled = False
        Case 1: SoundsEnabled = True
    End Select
    SaveSetting "ExcellOS", "Misc", "Sounds", Check1.Value
End Sub

Private Sub Command4_Click()
    For i = 0 To 47
        TBarColours(i).Visible = True
    Next i
End Sub

Private Sub Command5_Click()
    For i = 0 To 47
        TBarTcolour(i).Visible = True
    Next i
End Sub

Private Sub Form_Load()
    For i = 0 To 47
        TBarColours(i).Visible = False
        TBarTcolour(i).Visible = False
    Next i
    
    Check1.Value = GetSetting("ExcellOS", "Misc", "Sounds")
    Text1.Text = GetSetting("ExcellOS", "bg", "Timer")
    
    Label4.BackColor = GetSetting("ExcellOS", "Misc", "TbarBackColour")
    Label7.BackColor = GetSetting("ExcellOS", "Misc", "TbarForeColour")
    
    Label1.BackColor = FrmOptions.Label4.BackColor
    Label1.ForeColor = FrmOptions.Label7.BackColor
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    For i = 0 To 2
        Label2(i).FontUnderline = False
        Label2(i).ForeColor = &H0&
    Next i
End Sub

Private Sub Label2_Click(Index As Integer)
    Select Case Index
        Case 0
            Frame1.Visible = True
            Frame2.Visible = False
            Frame3.Visible = False
            Frame4.Visible = False
            Label2(0).FontItalic = True
            Label2(1).FontItalic = False
            Label2(2).FontItalic = False
            Label2(0).Enabled = False
            Label2(1).Enabled = True
            Label2(2).Enabled = True
        Case 1
            Frame1.Visible = False
            Frame2.Visible = False
            Frame3.Visible = True
            Frame4.Visible = False
            Label2(0).FontItalic = False
            Label2(1).FontItalic = True
            Label2(2).FontItalic = False
            Label2(0).Enabled = True
            Label2(1).Enabled = False
            Label2(2).Enabled = True
        Case 2
            Frame1.Visible = False
            Frame2.Visible = True
            Frame3.Visible = False
            Frame4.Visible = False
            Label2(0).FontItalic = False
            Label2(1).FontItalic = False
            Label2(2).FontItalic = True
            Label2(0).Enabled = True
            Label2(1).Enabled = True
            Label2(2).Enabled = False
    End Select
End Sub

Private Sub Label2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
        Case 0: ChangeFormat (0)
        Case 1: ChangeFormat (1)
        Case 2: ChangeFormat (2)
    End Select
End Sub

Private Sub TBarColours_Click(Index As Integer)
    Label4.BackColor = TBarColours(Index).BackColor
    SaveSetting "ExcellOS", "Misc", "TbarBackColour", Label4.BackColor
    
    For i = 0 To 47
        TBarColours(i).Visible = False
    Next i
End Sub

Private Sub TBarTcolour_Click(Index As Integer)
    Label7.BackColor = TBarTcolour(Index).BackColor
    SaveSetting "ExcellOS", "Misc", "TbarForeColour", Label7.BackColor
    
    For i = 0 To 47
        TBarTcolour(i).Visible = False
    Next i
End Sub

Private Sub Text1_Change()
    SaveSetting "ExcellOS", "bg", "Timer", Text1.Text
End Sub

Private Sub ChangeFormat(SelectedL As Integer)
    Select Case SelectedL
        Case 0
            Label2(SelectedL).FontUnderline = True
            Label2(SelectedL).ForeColor = &HFF0000
            Num = SelectedL + 1
            Label2(Num).FontUnderline = False
            Label2(Num).ForeColor = &H0&
            Num = Num + 1
            Label2(Num).FontUnderline = False
            Label2(Num).ForeColor = &H0&
        Case 1
            Label2(SelectedL).FontUnderline = True
            Label2(SelectedL).ForeColor = &HFF0000
            Num = SelectedL - 1
            Label2(Num).FontUnderline = False
            Label2(Num).ForeColor = &H0&
            Num = Num + 2
            Label2(Num).FontUnderline = False
            Label2(Num).ForeColor = &H0&
        Case 2
            Label2(SelectedL).FontUnderline = True
            Label2(SelectedL).ForeColor = &HFF0000
            Num = SelectedL - 1
            Label2(Num).FontUnderline = False
            Label2(Num).ForeColor = &H0&
            Num = Num - 1
            Label2(Num).FontUnderline = False
            Label2(Num).ForeColor = &H0&
    End Select
End Sub

Private Sub Text2_Change()
    If IsNumeric(Text2.Text) = False Then
        MsgBox "You may only use numbers in this data field", vbCritical, "Excell OS processing error"
    Else
        SaveSetting "ExcellOS", "Misc", "SSFS", Text2.Text
        FrmSS.Label1.FontSize = Text2.Text
    End If
    
    If Text2.Text > "80" Then
        MsgBox "Highest font number they may be used is '80'", vbCritical, "Excell OS processing error"
        Text2.Text = "80"
    End If
End Sub
