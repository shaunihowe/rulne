VERSION 5.00
Begin VB.Form mainform 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rulne"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   5130
   Icon            =   "mainform.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   5130
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer cputurn 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6000
      Top             =   2640
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5520
      Top             =   2640
   End
   Begin VB.Label debugbox 
      BackStyle       =   0  'Transparent
      Caption         =   "Debug Info..."
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   5175
      Left            =   5160
      TabIndex        =   104
      Top             =   240
      Width           =   2055
      WordWrap        =   -1  'True
   End
   Begin VB.Label turncur 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "->"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2280
      TabIndex        =   102
      Top             =   5040
      Width           =   615
   End
   Begin VB.Label sco 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   1
      Left            =   3120
      TabIndex        =   101
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Label sco 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   0
      Left            =   720
      TabIndex        =   100
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   99
      Left            =   4560
      TabIndex        =   99
      Top             =   4560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   98
      Left            =   4080
      TabIndex        =   98
      Top             =   4560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   97
      Left            =   3600
      TabIndex        =   97
      Top             =   4560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   96
      Left            =   3120
      TabIndex        =   96
      Top             =   4560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   95
      Left            =   2640
      TabIndex        =   95
      Top             =   4560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   94
      Left            =   2160
      TabIndex        =   94
      Top             =   4560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   93
      Left            =   1680
      TabIndex        =   93
      Top             =   4560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   92
      Left            =   1200
      TabIndex        =   92
      Top             =   4560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   91
      Left            =   720
      TabIndex        =   91
      Top             =   4560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   90
      Left            =   240
      TabIndex        =   90
      Top             =   4560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   89
      Left            =   4560
      TabIndex        =   89
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   88
      Left            =   4080
      TabIndex        =   88
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   87
      Left            =   3600
      TabIndex        =   87
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   86
      Left            =   3120
      TabIndex        =   86
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   85
      Left            =   2640
      TabIndex        =   85
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   84
      Left            =   2160
      TabIndex        =   84
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   83
      Left            =   1680
      TabIndex        =   83
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   82
      Left            =   1200
      TabIndex        =   82
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   81
      Left            =   720
      TabIndex        =   81
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   80
      Left            =   240
      TabIndex        =   80
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   79
      Left            =   4560
      TabIndex        =   79
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   78
      Left            =   4080
      TabIndex        =   78
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   77
      Left            =   3600
      TabIndex        =   77
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   76
      Left            =   3120
      TabIndex        =   76
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   75
      Left            =   2640
      TabIndex        =   75
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   74
      Left            =   2160
      TabIndex        =   74
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   73
      Left            =   1680
      TabIndex        =   73
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   72
      Left            =   1200
      TabIndex        =   72
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   71
      Left            =   720
      TabIndex        =   71
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   70
      Left            =   240
      TabIndex        =   70
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   69
      Left            =   4560
      TabIndex        =   69
      Top             =   3120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   68
      Left            =   4080
      TabIndex        =   68
      Top             =   3120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   67
      Left            =   3600
      TabIndex        =   67
      Top             =   3120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   66
      Left            =   3120
      TabIndex        =   66
      Top             =   3120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   65
      Left            =   2640
      TabIndex        =   65
      Top             =   3120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   64
      Left            =   2160
      TabIndex        =   64
      Top             =   3120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   63
      Left            =   1680
      TabIndex        =   63
      Top             =   3120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   62
      Left            =   1200
      TabIndex        =   62
      Top             =   3120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   61
      Left            =   720
      TabIndex        =   61
      Top             =   3120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   60
      Left            =   240
      TabIndex        =   60
      Top             =   3120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   59
      Left            =   4560
      TabIndex        =   59
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   58
      Left            =   4080
      TabIndex        =   58
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   57
      Left            =   3600
      TabIndex        =   57
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   56
      Left            =   3120
      TabIndex        =   56
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   55
      Left            =   2640
      TabIndex        =   55
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   54
      Left            =   2160
      TabIndex        =   54
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   53
      Left            =   1680
      TabIndex        =   53
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   52
      Left            =   1200
      TabIndex        =   52
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   51
      Left            =   720
      TabIndex        =   51
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   50
      Left            =   240
      TabIndex        =   50
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   49
      Left            =   4560
      TabIndex        =   49
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   48
      Left            =   4080
      TabIndex        =   48
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   47
      Left            =   3600
      TabIndex        =   47
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   46
      Left            =   3120
      TabIndex        =   46
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   45
      Left            =   2640
      TabIndex        =   45
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   44
      Left            =   2160
      TabIndex        =   44
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   43
      Left            =   1680
      TabIndex        =   43
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   42
      Left            =   1200
      TabIndex        =   42
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   41
      Left            =   720
      TabIndex        =   41
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   40
      Left            =   240
      TabIndex        =   40
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   39
      Left            =   4560
      TabIndex        =   39
      Top             =   1680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   38
      Left            =   4080
      TabIndex        =   38
      Top             =   1680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   37
      Left            =   3600
      TabIndex        =   37
      Top             =   1680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   36
      Left            =   3120
      TabIndex        =   36
      Top             =   1680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   35
      Left            =   2640
      TabIndex        =   35
      Top             =   1680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   34
      Left            =   2160
      TabIndex        =   34
      Top             =   1680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   33
      Left            =   1680
      TabIndex        =   33
      Top             =   1680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   32
      Left            =   1200
      TabIndex        =   32
      Top             =   1680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   31
      Left            =   720
      TabIndex        =   31
      Top             =   1680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   30
      Left            =   240
      TabIndex        =   30
      Top             =   1680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   29
      Left            =   4560
      TabIndex        =   29
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   28
      Left            =   4080
      TabIndex        =   28
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   27
      Left            =   3600
      TabIndex        =   27
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   26
      Left            =   3120
      TabIndex        =   26
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   25
      Left            =   2640
      TabIndex        =   25
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   24
      Left            =   2160
      TabIndex        =   24
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   23
      Left            =   1680
      TabIndex        =   23
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   22
      Left            =   1200
      TabIndex        =   22
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   21
      Left            =   720
      TabIndex        =   21
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   20
      Left            =   240
      TabIndex        =   20
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   19
      Left            =   4560
      TabIndex        =   19
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   18
      Left            =   4080
      TabIndex        =   18
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   17
      Left            =   3600
      TabIndex        =   17
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   16
      Left            =   3120
      TabIndex        =   16
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   15
      Left            =   2640
      TabIndex        =   15
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   14
      Left            =   2160
      TabIndex        =   14
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   13
      Left            =   1680
      TabIndex        =   13
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   12
      Left            =   1200
      TabIndex        =   12
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   11
      Left            =   720
      TabIndex        =   11
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   10
      Left            =   240
      TabIndex        =   10
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   9
      Left            =   4560
      TabIndex        =   9
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   8
      Left            =   4080
      TabIndex        =   8
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   7
      Left            =   3600
      TabIndex        =   7
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   6
      Left            =   3120
      TabIndex        =   6
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   5
      Left            =   2640
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   4
      Left            =   2160
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   3
      Left            =   1680
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   2
      Left            =   1200
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   720
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape frame 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FFFF&
      Height          =   4815
      Left            =   120
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label stat 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Press 'F5' to start a new game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   0
      TabIndex        =   103
      Top             =   5520
      Width           =   5145
   End
   Begin VB.Menu menu_game 
      Caption         =   "&Game"
      Begin VB.Menu menu_new 
         Caption         =   "&New Game"
         Begin VB.Menu menu_newc 
            Caption         =   "&Two Player"
            Index           =   0
            Shortcut        =   +{F5}
         End
         Begin VB.Menu menu_newc 
            Caption         =   "&Easy"
            Index           =   1
            Shortcut        =   {F5}
         End
         Begin VB.Menu menu_newc 
            Caption         =   "&Normal"
            Checked         =   -1  'True
            Index           =   2
            Shortcut        =   {F6}
         End
         Begin VB.Menu menu_newc 
            Caption         =   "&Hard"
            Index           =   3
            Shortcut        =   {F7}
         End
         Begin VB.Menu menu_newc 
            Caption         =   "E&xpert"
            Index           =   4
            Shortcut        =   {F8}
         End
         Begin VB.Menu menu_newc 
            Caption         =   "&Computer Vs Computer"
            Index           =   5
         End
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu menu_highscore_show 
         Caption         =   "&Show Highscores"
         Shortcut        =   {F12}
      End
      Begin VB.Menu menu_highscore_condense 
         Caption         =   "&Condense Scores"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu menu_about_debug 
         Caption         =   "&Debug Info"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu menu_exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu menu_about 
      Caption         =   "&Help"
      Begin VB.Menu menu_about_rulne 
         Caption         =   "About& Rulne"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "mainform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub box_Click(Index As Integer)
If Current_Board.Player(Current_Board.Turn) <> Player_Human Then Exit Sub
If Current_Board.GameIn = False Then Exit Sub
DoMove Current_Board, Index
DrawBoard Current_Board
End Sub

Private Sub cputurn_Timer()
Dim cstime As Long
cputurn.Enabled = False
If Current_Board.GameIn = False Then Exit Sub
cstime = Timer
DebugInfo = vbCrLf & vbCrLf & "Move Evaulation..." & vbCrLf
DoEvents
If Current_Board.Player(Current_Board.Turn) = Player_CPUEASY Then
    hg = -10
    hgn = -1
    For a = 0 To 99
        If Current_Board.moves(a) = True And Current_Board.Allowed(a) = True And Current_Board.Value(a) > hg Then hg = Current_Board.Value(a): hgn = a
    Next a
    movv = hgn
    debuginfo2$ = debuginfo2$ & "Highest Only..." & vbCrLf
    DebugInfo$ = DebugInfo$ & "Best: " & Current_Board.Value(hgn) & vbCrLf
Else
    Nodes = 0
    movv = DoCompMove(Current_Board, -10, DiffLevel(Current_Board.Player(Current_Board.Turn)).BasePlys)
    debuginfo2$ = debuginfo2$ & "Minimax Search..." & vbCrLf & "Base Depth: " & DiffLevel(Current_Board.Player(Current_Board.Turn)).BasePlys & vbCrLf & _
    "Extended Depth: " & 0 - Current_Board.maxplys(Current_Board.Turn) & vbCrLf
End If
If movv > -1 Then DoMove Current_Board, movv
cetime = Timer - cstime
DoEvents
If cetime < 0 Then cetime = 0 - cetime
If cetime = 0 Then cetime = 1
If Current_Board.Turn = 0 Then ene = 1 Else ene = 0
debuginfo2$ = debuginfo2$ & "Time taken: " & Round(cetime, 2) & vbCrLf
debuginfo2$ = debuginfo2$ & "Nps: " & Int(Nodes / cetime) & vbCrLf
debuginfo2$ = debuginfo2$ & DebugInfo
If Current_Board.Player(ene) = Player_CPUEXPT Then
    If cetime < 0.3 Then
        'If (0 - Current_Board.maxplys(ene)) < 2 Then
            Current_Board.maxplys(ene) = Current_Board.maxplys(ene) - 2
        'Else
            'DiffLevel(Current_Board.Player(ene)).BasePlys = DiffLevel(Current_Board.Player(ene)).BasePlys + 2
            'Current_Board.maxplys(ene) = 0
        'End If
    End If
End If
mainform.debugbox.Caption = debuginfo2$
DrawBoard Current_Board
End Sub

Private Sub Form_Load()
LoadScores
End Sub

Private Sub Form_Resize()
dbg = menu_about_debug.Checked
DiffLevel(0).Name = "Two Player"
DiffLevel(1).Name = "Easy"
DiffLevel(2).Name = "Normal"
DiffLevel(3).Name = "Hard"
DiffLevel(4).Name = "All this PC's might"
DiffLevel(5).Name = "All this PC's might (Beta)"
For a = 0 To 4
    menu_newc(a).Caption = DiffLevel(a).Name
Next a
End Sub

Private Sub Form_Terminate()
gameclosed = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form_Terminate
End Sub

Private Sub menu_about_debug_Click()
If menu_about_debug.Checked = True Then
    menu_about_debug.Checked = False
    mainform.Width = 5220
    mainform.stat.Width = 5145
Else
    menu_about_debug.Checked = True
    mainform.Width = 7410
    mainform.stat.Width = 7320
End If
dbg = menu_about_debug.Checked
End Sub

Private Sub menu_about_rulne_Click()
MsgBox "Rulne - Programmed by Shaun Howe", , "About Rulne"
End Sub

Private Sub menu_exit_Click()
gameclosed = True
End
End Sub

Private Sub menu_highscore_condense_Click()
For a = 0 To 3
    condense a
    ScoreBoard(a).LoadHighScoreData
Next a
End Sub


Public Sub menu_highscore_show_Click()
For a = 0 To 4
    msg$ = msg$ & DiffLevel(a + 1).Name & ": " & vbCrLf
    For b = 1 To ScoreBoard(a).NumberOfEntrys
        sp$ = Space$(40 - Len(ScoreBoard(a).Name(b)))
        msg = msg$ & b & "  " & Smooth(ScoreBoard(a).Name(b)) & sp$ & Str(ScoreBoard(a).Score(b)) & vbCrLf
    Next b
    msg$ = msg$ & vbCrLf
Next a
MsgBox msg$, vbInformation, "Highscores"
End Sub

Private Sub menu_newc_Click(Index As Integer)
menu_newc(Index).Checked = True
For a = 0 To 5
    If a <> Index Then menu_newc(a).Checked = False
Next a
If menu_newc(0).Checked = True Then lev = Player_Human
If menu_newc(1).Checked = True Then lev = Player_CPUEASY
If menu_newc(2).Checked = True Then lev = Player_CPUNORM
If menu_newc(3).Checked = True Then lev = Player_CPUHARD
If menu_newc(4).Checked = True Then lev = Player_CPUEXPT
If menu_newc(5).Checked = True Then lev = Player_TWOCPUS

If ClockCycles = 0 Then
    Dim stime, etime As Long
    MsgBox "Starting a quick speed test to determine" & vbCrLf & "how fast your compter is." & vbCrLf & "it will take about 5 seconds.", , "Rulne - Speed Test"
    mainform.stat.Caption = "Timing Computer Clock Speed..."
    DoEvents ' give windows a chance to breath before we start
    stime = Timer
    Do
        ClockCycles = ClockCycles + 1
        etime = Timer - stime
    Loop Until etime > 4
    ClockCycles = Int(ClockCycles / 5)
End If
SetBoard Current_Board, lev
DrawBoard Current_Board
End Sub

Private Sub Timer1_Timer()
random = random Xor Int(Rnd(1) * 32000) + 1
End Sub
