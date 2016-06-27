VERSION 5.00
Begin VB.Form frmMessage 
   Caption         =   "Message from MAP"
   ClientHeight    =   2085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2085
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton btnexit 
      Height          =   735
      Left            =   3840
      Picture         =   "frmMessage.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Left            =   360
      Top             =   0
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Zentriert
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "lllllllll jjjjjjjjjjjjjjjj hhhhhhhhhhhhhhhhh iiiiiiiiiiiiiiiiiiiiiiiiii hhhhhhhhhhhhhhhh pppppppppppppppppppppppppppp"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = App.ExeName & " Message " & g_strProgramVersion
    lblMessage.Caption = g_strMessage
    Timer1.Interval = 10000
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    Unload Me
End Sub
