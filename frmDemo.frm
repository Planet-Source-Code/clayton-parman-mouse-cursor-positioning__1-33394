VERSION 5.00
Begin VB.Form frmDemo 
   Caption         =   "Cursor Positioning Demo"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   4860
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   4860
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd 
      Caption         =   "&Start Over"
      Height          =   420
      Index           =   5
      Left            =   270
      TabIndex        =   5
      Top             =   4545
      Width           =   1140
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Exit"
      Height          =   420
      Index           =   4
      Left            =   270
      TabIndex        =   4
      Top             =   3870
      Width           =   1140
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Click Me"
      Height          =   420
      Index           =   3
      Left            =   270
      TabIndex        =   3
      Top             =   3180
      Width           =   1140
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Click Me"
      Height          =   420
      Index           =   2
      Left            =   270
      TabIndex        =   2
      Top             =   2490
      Width           =   1140
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Click Me"
      Height          =   420
      Index           =   1
      Left            =   270
      TabIndex        =   1
      Top             =   1800
      Width           =   1140
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Click Me"
      Height          =   420
      Index           =   0
      Left            =   270
      TabIndex        =   0
      Top             =   1110
      Width           =   1140
   End
   Begin VB.Label lblNote 
      Alignment       =   2  'Center
      Caption         =   "and is screen resolution independent"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   285
      TabIndex        =   12
      Top             =   495
      Width           =   4290
   End
   Begin VB.Label lblNote 
      Alignment       =   2  'Center
      Caption         =   "Works on any control with an hwnd property."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   285
      TabIndex        =   11
      Top             =   180
      Width           =   4290
   End
   Begin VB.Label lblPositon 
      Caption         =   "Cursor is centered just ""below"" the button"
      Height          =   285
      Index           =   4
      Left            =   1530
      TabIndex        =   10
      Top             =   3915
      Width           =   3165
   End
   Begin VB.Label lblPositon 
      Caption         =   "Cursor is just ""inside"" of Top - Left corner"
      Height          =   285
      Index           =   3
      Left            =   1530
      TabIndex        =   9
      Top             =   3240
      Width           =   3165
   End
   Begin VB.Label lblPositon 
      Caption         =   "Cursor is centered at Top of button"
      Height          =   285
      Index           =   2
      Left            =   1530
      TabIndex        =   8
      Top             =   2565
      Width           =   3165
   End
   Begin VB.Label lblPositon 
      Caption         =   "Cursor is in Center of button"
      Height          =   285
      Index           =   1
      Left            =   1530
      TabIndex        =   7
      Top             =   1845
      Width           =   3165
   End
   Begin VB.Label lblPositon 
      Caption         =   "Cursor is at Bottom-Right corner of button"
      Height          =   285
      Index           =   0
      Left            =   1530
      TabIndex        =   6
      Top             =   1170
      Width           =   3165
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Edge designations are:  T=Top, B=Bottom, L=Left, R=Right
'or in combinations of:  TL, TR, BL, BR -  LT, RT, LB, RB
  
'If no edge is specified the Mouse cursor will be centered.


Private Sub Form_Activate()
   MoveMouseCursorTo cmd(0), "BR"                   'Bottom-Right corner.
End Sub


Private Sub cmd_Click(Index As Integer)
   Select Case Index
      Case 0:  MoveMouseCursorTo cmd(1)             'Centered (default)
      Case 1:  MoveMouseCursorTo cmd(2), "T"        'Centered at Top
      Case 2:  MoveMouseCursorTo cmd(3), "TL", 3, 3 'Top-Left w/Adjust
      Case 3:  MoveMouseCursorTo cmd(4), "B", 6     'Centered below Bottom.
      Case 4:  End
      Case 5:  MoveMouseCursorTo cmd(0), "BR"       'Bottom-Right corner.
   End Select
End Sub
