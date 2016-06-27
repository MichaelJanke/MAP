VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmInfo 
   Caption         =   "Info"
   ClientHeight    =   8595
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14235
   Icon            =   "frmInfo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   573
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   949
   StartUpPosition =   3  'Windows-Standard
   Begin VB.ComboBox cboOrg 
      Height          =   315
      Left            =   6720
      TabIndex        =   37
      Text            =   "cboOrg"
      Top             =   690
      Width           =   1515
   End
   Begin VB.ComboBox cboBundesland 
      Height          =   315
      Left            =   0
      TabIndex        =   36
      Top             =   0
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CommandButton btnNurAbwesend 
      Height          =   555
      Left            =   11520
      Picture         =   "frmInfo.frx":0442
      Style           =   1  'Grafisch
      TabIndex        =   29
      Top             =   7980
      Width           =   795
   End
   Begin VB.Timer T 
      Left            =   120
      Top             =   480
   End
   Begin VB.CommandButton btnHelp 
      Default         =   -1  'True
      Height          =   555
      Left            =   13260
      Picture         =   "frmInfo.frx":0884
      Style           =   1  'Grafisch
      TabIndex        =   28
      ToolTipText     =   "Funktionen dieses Formulars"
      Top             =   7980
      Width           =   795
   End
   Begin VB.CommandButton btnSecurity 
      Height          =   555
      Left            =   12420
      Picture         =   "frmInfo.frx":0CC6
      Style           =   1  'Grafisch
      TabIndex        =   27
      ToolTipText     =   "Sicherheits-Einstellungen"
      Top             =   7980
      Width           =   795
   End
   Begin VB.CommandButton btnGenehmigen 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   10200
      Picture         =   "frmInfo.frx":1108
      Style           =   1  'Grafisch
      TabIndex        =   26
      ToolTipText     =   "Genehmigen/Ablehnen"
      Top             =   180
      Width           =   795
   End
   Begin VB.HScrollBar ScrollOrg 
      Height          =   255
      Left            =   6720
      Max             =   10
      TabIndex        =   23
      Top             =   720
      Value           =   1
      Width           =   1515
   End
   Begin VB.HScrollBar ScrollQuartal 
      Height          =   255
      Left            =   8400
      Max             =   4
      Min             =   1
      TabIndex        =   2
      Top             =   720
      Value           =   1
      Width           =   1755
   End
   Begin VB.CommandButton btnExit 
      Cancel          =   -1  'True
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   11040
      Picture         =   "frmInfo.frx":154A
      Style           =   1  'Grafisch
      TabIndex        =   3
      ToolTipText     =   "Programm verlassen"
      Top             =   180
      Width           =   795
   End
   Begin MSFlexGridLib.MSFlexGrid FG 
      Height          =   6735
      Left            =   180
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1080
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   11880
      _Version        =   393216
      Rows            =   5
      Cols            =   20
      BackColorBkg    =   -2147483633
      AllowBigSelection=   -1  'True
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      SelectionMode   =   1
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblText 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Seminar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   12120
      TabIndex        =   21
      Top             =   3840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblText 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Krank"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   12120
      TabIndex        =   22
      Top             =   4080
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblText 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Dienstreise"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   12120
      TabIndex        =   20
      Top             =   3480
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblText 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Feiertag"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   12000
      TabIndex        =   19
      Top             =   3120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblText 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "FAKO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   12000
      TabIndex        =   18
      Top             =   2760
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblText 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Krank"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   12240
      TabIndex        =   33
      Top             =   4320
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblText 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Krank"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   12240
      TabIndex        =   30
      Top             =   4560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblText 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Urlaub"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   12000
      TabIndex        =   17
      Top             =   2400
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblTyp 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   15
      Left            =   12960
      TabIndex        =   35
      Top             =   7200
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label lblTyp 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   14
      Left            =   12840
      TabIndex        =   34
      Top             =   6840
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label lblTyp 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   13
      Left            =   12960
      TabIndex        =   32
      Top             =   6480
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label lblTyp 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   16
      Left            =   13080
      TabIndex        =   31
      Top             =   7560
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label lblStatus 
      Caption         =   "."
      Height          =   195
      Left            =   120
      TabIndex        =   25
      Top             =   7920
      Width           =   11535
   End
   Begin VB.Label lblOrg 
      Alignment       =   2  'Zentriert
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Org"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   435
      Left            =   6720
      TabIndex        =   24
      Top             =   180
      Width           =   1515
   End
   Begin VB.Label lblTyp 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   7
      Left            =   11880
      TabIndex        =   16
      Top             =   7200
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label lblTyp 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   8
      Left            =   12000
      TabIndex        =   15
      Top             =   7560
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label lblTyp 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   9
      Left            =   12960
      TabIndex        =   14
      Top             =   5040
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label lblTyp 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   10
      Left            =   12840
      TabIndex        =   13
      Top             =   5400
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label lblTyp 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   11
      Left            =   12840
      TabIndex        =   12
      Top             =   5760
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label lblTyp 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   12
      Left            =   12840
      TabIndex        =   11
      Top             =   6120
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label lblTyp 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   6
      Left            =   11880
      TabIndex        =   10
      Top             =   6840
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label lblTyp 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   5
      Left            =   11880
      TabIndex        =   9
      Top             =   6480
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label lblTyp 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   4
      Left            =   11880
      TabIndex        =   8
      Top             =   6120
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label lblTyp 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   11880
      TabIndex        =   7
      Top             =   5760
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label lblTyp 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   11880
      TabIndex        =   6
      Top             =   5280
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label lblTyp 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   11880
      TabIndex        =   5
      ToolTipText     =   "lblTyp(1)"
      Top             =   4920
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label lblQuartal 
      Alignment       =   2  'Zentriert
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Q1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   435
      Left            =   8400
      TabIndex        =   4
      Top             =   180
      Width           =   1755
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Info Abwesenheit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   795
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   6375
   End
   Begin VB.Menu mnuAbwesenheit 
      Caption         =   "Abwesenheit"
      Visible         =   0   'False
      Begin VB.Menu mnuBeantragen 
         Caption         =   "Beantragen"
      End
      Begin VB.Menu mnuChange 
         Caption         =   "Ändern"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Löschen"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuListe 
         Caption         =   "Zeige Liste"
      End
   End
   Begin VB.Menu mnuUsername 
      Caption         =   "Username"
      Visible         =   0   'False
      Begin VB.Menu mnuHome 
         Caption         =   "Change HOME"
      End
      Begin VB.Menu mnuChangeToUser 
         Caption         =   "Change to User"
      End
      Begin VB.Menu mnuSOY 
         Caption         =   "Show Sum of Year"
      End
   End
   Begin VB.Menu mnuAddon 
      Caption         =   "Addon"
      Visible         =   0   'False
      Begin VB.Menu mnuAddonOpenOptionForm 
         Caption         =   "Open Option Form"
      End
      Begin VB.Menu mnuAddonUser 
         Caption         =   "Open User Form"
      End
      Begin VB.Menu mnuAddonOrgItem 
         Caption         =   "Open OrgItem Form"
      End
      Begin VB.Menu mnuAddonChangeLocale 
         Caption         =   "Change locale"
      End
      Begin VB.Menu mnuAddonBildschirmInfo 
         Caption         =   "Bildschirm Info"
      End
      Begin VB.Menu mnuAddonJahressumme 
         Caption         =   "JahresSumme"
      End
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_dtmStart As Date, m_dtmEnde As Date, m_dtmRun As Date
Private i As Long
Private m_aCOLORS(ABW_URLAUB To ABW_KRANKHEIT, COL_A_GEPLANT To COL_A_GENEHMIGT) As Long
Private m_aUserHash() As Integer
Private m_aFG() As Long   ' Matrix mit den aktuellen Abwesenheiten / Feiertagen
Private m_booInit As Boolean
Private m_booSelected As Boolean
Private m_lngSelStart As Long, m_lngSelEnd As Long
Private m_dtmLastAction As Date

Private m_booNurAbwesenheit As Boolean
Private m_booOrgSlider As Boolean
Private m_lngIdxOrg As Long
Private m_selectedLine As Long

Private m_cAb As clsAbw     ' Speicher für das Kontextmenue
Private m_cUs As clsUser    ' Speicher für das Kontextmenue

Public Event CloseApplication()

Private Sub btnNurAbwesend_Click()
    m_dtmLastAction = Now()
    If m_booNurAbwesenheit = True Then    ' Umschalten auf NormalSicht
        m_booNurAbwesenheit = False
        ScrollQuartal.Enabled = True
        ScrollOrg.Enabled = True
        btnGenehmigen.Enabled = True
        btnSecurity.Enabled = True
        g_db.FillAbwesenheitCollection
        FillFG
        FG_Set_Today
    Else                                ' Umschalten auf AlleAbwesend
        m_booNurAbwesenheit = True
        ScrollQuartal.Enabled = False
        ScrollOrg.Enabled = False
        btnGenehmigen.Enabled = False
        btnSecurity.Enabled = False
        g_db.FillAbwesenheitCollection "WHERE tblAbwesenheit.dtmEnde > now() -1 or dtmstart > now() -1 ", "ORDER BY tblAbwesenheit.dtmStart, tblAbwesenheit.lngGVN, tblAbwesenheit.dtmEnde"
        FillFG
        FG_Set_Today
    End If
End Sub

Private Sub btnSecurity_Click()
    m_dtmLastAction = Now()
    frmSecurity.Show vbModal
    g_db.FillUserCollection   ' für den Kalender User-Daten ebenfalls noch mal einlesen
    g_db.setOrgUserVisibility ' stelle die (veränderte) Sichtbarkeit der org und user für diesen Benutzer her.
    Arrange_Elements          ' Securityrelevante Darstellungselemente updaten
    FillFG
End Sub

'##########################################################################################################
Private Sub btnExit_Click()
    Unload Me
End Sub
'##########################################################################################################
Private Sub btnGenehmigen_Click()
    m_dtmLastAction = Now()
    frmManager.Show vbModal
    FillFG
    FG_Set_Today
End Sub

Private Sub btnHelp_Click()
    m_dtmLastAction = Now()
    g_InfoText = g_db.GetString(1002) & vbCrLf & vbCrLf & _
        g_db.GetString(1003) & ":" & vbCrLf & g_db.GetString(1004) & vbCrLf & g_db.GetString(1005) & vbCrLf & vbCrLf & _
        g_db.GetString(1006) & ":" & vbCrLf & g_db.GetString(1007) & vbCrLf & vbCrLf & _
        g_db.GetString(1008) & ":" & vbCrLf & g_db.GetString(1009) & vbCrLf & vbCrLf & _
        g_db.GetString(1010) & ":" & vbCrLf & g_db.GetString(1011) & vbCrLf & vbCrLf & _
        g_db.GetString(1012) & ":" & vbCrLf & g_db.GetString(1013)
    If g_CU.booIsSek Then    ' Per Doppelclick Identität JEDES Benutzers
        g_InfoText = g_InfoText & vbCrLf & vbCrLf & g_db.GetString(1014) & ":" & vbCrLf & g_db.GetString(1026)
    End If
    If g_CU.lngUserLevel < Benutzer Then   ' Per Doppelclick Identität eines MA
        g_InfoText = g_InfoText & vbCrLf & vbCrLf & g_db.GetString(1014) & ":" & vbCrLf & g_db.GetString(1015)
    End If
    If g_CU.IsPrivileged Then    ' Verlaufsinfo
        g_InfoText = g_InfoText & vbCrLf & vbCrLf & g_db.GetString(1016) & ":" & vbCrLf & g_db.GetString(1017)
    End If
    If g_db.KrankheitUserBeantragbar Then
        g_InfoText = g_InfoText & vbCrLf & vbCrLf & g_db.GetString(1018) & ":" & vbCrLf & g_db.GetString(1019)
    Else
        If g_CU.IsPrivileged Then    ' AL oder Guru
            g_InfoText = g_InfoText & vbCrLf & vbCrLf & g_db.GetString(1018) & ":" & vbCrLf & g_db.GetString(1027)
        End If
    End If

    g_PopUpMode = "Text"
    frmPopUp.Show vbModal
End Sub
Private Sub btnHelp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuAddon
    End If
End Sub

Private Sub cboBundesland_Click()
    m_dtmLastAction = Now()
    basFormUtils.SetRegData "Bundesland", cboBundesland.ListIndex
    FillFG
End Sub

'##########################################################################################################
Private Sub FG_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    m_dtmLastAction = Now()
    cboBundesland.Visible = False
    If m_booNurAbwesenheit Then Exit Sub

    Dim lngFGMRow As Long, lngFGMCol As Long
    lngFGMCol = FG.MouseCol
    lngFGMRow = FG.MouseRow
    lblStatus = ""

    ' Selektieren NUR in der eigenen Reihe
    If lngFGMRow <= ROW_TAG Or lngFGMRow = FG.Rows - 1 Then Exit Sub ' Nicht im Userbereich geclickt
    
    Dim lngIdxUser As Long
    lngIdxUser = m_aFG(0, lngFGMRow)
    If lngIdxUser > 0 Then  ' Hier steht überhaupt was
        Dim cUs As clsUser
        Set m_cUs = g_db.GetUserByID(lngIdxUser)      ' der "gewünschte" Benutzer ... merken für Kontextmenue-feuern
        If m_cUs Is Nothing Then Exit Sub
    End If
    
    
    If lngFGMCol = 0 And Button = vbLeftButton Then
        ' mark selected line(2)
        If m_selectedLine <> 0 Then FG.RowHeight(m_selectedLine) = FG.RowHeight(m_selectedLine) * 0.66
        m_selectedLine = lngFGMRow
        If m_selectedLine <> 0 Then FG.RowHeight(m_selectedLine) = FG.RowHeight(m_selectedLine) * 1.5
        Exit Sub
    End If
    
    ' Benutzername - Kontextmenue aufrufen
    If lngFGMCol = 0 And Button = vbRightButton Then         ' In Namens-Spalte
        If SetupMnuUsername Then PopupMenu mnuUsername
        Exit Sub
    End If
    
    If m_aUserHash(g_CU.lngIdxUser) = 0 Then Exit Sub              ' Benutzer ist nicht auf dieser Seite - keine Zeiten eintragen
    
    If (m_dtmStart + lngFGMCol - OFFSET_FGDATE < Date) Then
        If Not g_CU_Login.IsSekOf(g_CU) And Not g_CU_Login.IsChefOf(g_CU) Then
            lblStatus = g_db.GetString(1020)     ' Termin liegt in der Vergangenheit
            Exit Sub
        End If
    End If
    
    FG.Row = m_aUserHash(g_CU.lngIdxUser)
    If m_aFG(lngFGMCol, FG.Row) > 0 Then ' Hier ist bereits Abwesenheit geplant
        If Button = vbRightButton Then
            Set m_cAb = g_db.GetAbwesenheit(m_aFG(lngFGMCol, FG.Row))   ' Speichern - auf dieser Abwesenheit wurde geklickt

            SetupMnuAbwesenheit
            PopupMenu mnuAbwesenheit
        End If
        Exit Sub
    End If
    If m_aFG(lngFGMCol, FG.Row) < 0 Then ' Feiertag / Zwangsabwesenheit
        m_dtmRun = m_dtmStart + lngFGMCol - OFFSET_FGDATE
        Dim cF As clsFeiertag
        Set cF = g_db.GetFeiertagByDate(m_dtmRun)
        If cF.lngIdxAbwesenheitsArt = basGlobals.ABW_FEIERTAG Then Exit Sub
    End If
    If m_booSelected = False Then
        m_booSelected = True
        m_lngSelStart = lngFGMCol
        m_lngSelEnd = lngFGMCol
        FG.col = lngFGMCol
        FG.CellBackColor = COLOR_PLANEN
    End If
    FG.ToolTipText = Format(m_dtmStart + lngFGMCol - OFFSET_FGDATE, "yyyy-mm-dd")
End Sub
Private Sub FG_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim cF As clsFeiertag, calStart As Date, calEnde As Date, booGueltigesDatum As Boolean, varval As Variant
    If Not m_booSelected Then Exit Sub
    If m_booNurAbwesenheit Then Exit Sub

    m_dtmLastAction = Now()
    m_lngSelEnd = FG.MouseCol
    ' Angewählte Abwesenheit: Start vor heute und nicht besondere Rechte? -> Weg damit
    If (m_dtmStart + m_lngSelEnd - OFFSET_FGDATE < Date) And Not g_CU_Login.booIsSek And Not g_CU_Login.lngIdxUser = g_CU.cOrg.lngIdxChef Then
        lblStatus = g_db.GetString(1020)
        If m_booSelected Then
            FG.col = m_lngSelStart
            FG.Row = m_aUserHash(g_CU.lngIdxUser)
            FG.CellBackColor = COLOR_FREI
        End If
        m_booSelected = False
        Exit Sub    ' Ende in der Vergangenheit
    Else
        lblStatus = ""
    End If
    
    m_booSelected = False
    If m_lngSelStart > m_lngSelEnd Then
        i = m_lngSelEnd:        m_lngSelEnd = m_lngSelStart:        m_lngSelStart = i ' Vertauschen
    End If

    calStart = m_dtmStart + m_lngSelStart - OFFSET_FGDATE
    calEnde = m_dtmStart + m_lngSelEnd - OFFSET_FGDATE
    calStart = basFunktionen.StartGueltig(calStart) ' Wenn auf einen Tag geclickt wird, der als Starttag nicht in Frage kommt, dann gehe so weit wie nötig vorwärts
    calEnde = basFunktionen.EndeGueltig(calEnde)    ' Wenn auf einen Tag geclickt wird, der als Endetag nicht in Frage kommt, dann gehe so weit wie nötig rückwärts

    FG.Redraw = False
    For i = m_lngSelStart To m_lngSelEnd
        FG.col = i
        If FG.CellBackColor = 0 Or FG.CellBackColor = COLOR_SONSTIGES Then FG.CellBackColor = COLOR_PLANEN
    Next i
    FG.Redraw = True

    m_booInit = True    ' KEIN FG_Redraw
    Dim cAb As New clsAbw
    cAb.dtmStart = m_lngSelStart + m_dtmStart - OFFSET_FGDATE
    cAb.dtmEnde = m_lngSelEnd + m_dtmStart - OFFSET_FGDATE
    cAb.lngIdxStatus = AbwStatus.UNDEFINED
    frmPlanen.EnterForm cAb
    frmPlanen.Show vbModal
    g_db.FillAbwesenheitCollection  ' Daten frisch reinholen
    m_booInit = False
    FillFG
End Sub

Private Sub FG_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' 1. Excel-Zieh-Effekt bei Markieren nachbilden
' 2. Wenn über Abwesenheit, dann Info darüber anzeigen
' 3. Wenn über Namen und SEK, dann Userinfo anzeigen

    If m_booNurAbwesenheit Then Exit Sub

    m_dtmLastAction = Now()
    Dim lngFGMRow As Long, lngFGMCol As Long
    lngFGMRow = FG.MouseRow
    lngFGMCol = FG.MouseCol

    ' geplanten Bereich dynamisch anzeigen
    If Button = vbLeftButton And m_booSelected Then
        '   m_lngSelEnd ist die bisherige Grenze
        If m_lngSelEnd <> lngFGMCol Then                      ' Hat sich überhaupt eine Änderung ergeben ?
            FG.Redraw = False                               ' es wird geändert - kein Zeichnen ...
            ' Die Erfassung der Mausposition erfolgt nicht kontinierlich - es können Sprünge entstehen.
            ' Deshalb merken wir uns in m_lngSelEnd, bis wohin die aktuelle Markierung reicht.
            '
            
            If lngFGMCol = m_lngSelStart Or _
               Sgn(m_lngSelEnd - m_lngSelStart) = -Sgn(lngFGMCol - m_lngSelStart) _
               Then                 ' zurück zum Ursprung  oder Vorzeichenwechsel  -> Entfärben
                If m_lngSelEnd > m_lngSelStart Then             ' von "oben"
                    ClearMark m_lngSelStart + 1, m_lngSelEnd    ' Alle Zellen über m_lngSelStart ...
                Else                                        ' von "unten"
                    ClearMark m_lngSelEnd, m_lngSelStart - 1    ' Alle Zellen unter m_lngSelStart ...
                End If
                m_lngSelEnd = m_lngSelStart
                FG.col = m_lngSelStart
            End If

            ' m_lngSelEnd kann links oder rechts von m_lngSelStart sein - Vorzeichen ...
            If lngFGMCol > m_lngSelStart Then                 ' positives Vorzeichen ...
                If lngFGMCol > m_lngSelEnd Then               ' größerer Bereich - evtl Einfärben
                    SetMark m_lngSelEnd + 1, lngFGMCol        ' nach bisherigen MarkEnde bis zu dieser Position
                Else                                        ' kleinerer Bereich
                    ClearMark lngFGMCol + 1, m_lngSelEnd      ' Alle Zellen über lngFGMCol ...
                End If
            Else                                            ' negatives Vorzeichen ...
                If lngFGMCol < m_lngSelEnd Then               ' größerer Bereich - evtl Einfärben
                    SetMark lngFGMCol, m_lngSelEnd - 1        ' nach bisherigen MarkEnde bis zu dieser Position
                Else                                        ' kleinerer Bereich
                    ClearMark m_lngSelEnd, lngFGMCol - 1      ' Alle Zellen unter lngFGMCol ...
                End If
            End If
            m_lngSelEnd = lngFGMCol
            FG.col = m_lngSelEnd
            FG.Redraw = True
        End If          ' Hat sich was geändert ?
        Exit Sub    ' Bei gedrückter Maustaste nichts weiter anzeigen
    End If  ' Linke Maustaste gedrückt ...

    If m_aFG(0, lngFGMRow) = 0 Then Exit Sub  ' Trennzeile - hier gibt es nichts anzuzeigen

    ' Anzeigen Info über Abwesenheit
    If lngFGMCol > OFFSET_FGDATE Then
        If m_aFG(lngFGMCol, lngFGMRow) > 0 Then             ' Positive Zahlen -> Abwesenheiten
            Dim cAb As clsAbw, lngIdxAbwesenheit As Long
            lngIdxAbwesenheit = m_aFG(lngFGMCol, lngFGMRow)
            Set cAb = g_db.GetAbwesenheit(lngIdxAbwesenheit)
            Dim strText As String
            strText = cAb.UserInfo & g_clsAb.Compose(cAb) & " " & g_clsAb.StatusText(cAb.lngIdxAbwesenheit, g_CU.IsSekOf(cAb.cUs))
            
            FG.ToolTipText = strText
            Set cAb = Nothing
        ElseIf m_aFG(lngFGMCol, lngFGMRow) < 0 Then         ' Negative Zahlen: Feiertage
            Dim lngIdxFeiertag As Long
            lngIdxFeiertag = -m_aFG(lngFGMCol, lngFGMRow)
            Dim cF As clsFeiertag
            Set cF = g_db.GetFeiertag(lngIdxFeiertag)
            If cF.strFeiertag <> "" Then
                FG.ToolTipText = cF.strFeiertag
            Else
                If cF.lngIdxAbwesenheitsArt = ABW_URLAUB Then FG.ToolTipText = g_db.GetString(1029)
                If cF.lngIdxAbwesenheitsArt = ABW_FAKO Then FG.ToolTipText = g_db.GetString(1030)
            End If
        Else                                                ' Wert ist 0 - nichts zu erklären
            FG.ToolTipText = ""
        End If
    ElseIf lngFGMCol = 0 Then   ' Info über Benutzer
        Dim cUs As clsUser
        Set cUs = g_db.GetUserByID(m_aFG(0, lngFGMRow))
        FG.ToolTipText = cUs.Fullname
        If cUs.colFunctions.Count > 0 Then
            FG.ToolTipText = FG.ToolTipText & "   Functions"
            Dim cFu As clsFunction
            For Each cFu In cUs.colFunctions
                FG.ToolTipText = FG.ToolTipText & ":" & cFu.strFunction
            Next
        End If
    Else
        FG.ToolTipText = ""
    End If
End Sub
Private Sub ClearMark(lngStepStart, lngStepEnd)
    For i = lngStepStart To lngStepEnd
        FG.col = i
        If m_aFG(i, m_aUserHash(g_CU.lngIdxUser)) = 0 And FG.CellBackColor <> COLOR_FREI Then FG.CellBackColor = COLOR_FREI     ' Entfärben
    Next i
End Sub
Private Sub SetMark(lngStepStart, lngStepEnd)
    For i = lngStepStart To lngStepEnd
        FG.col = i
        If m_aFG(i, m_aUserHash(g_CU.lngIdxUser)) = 0 And FG.CellBackColor <> COLOR_PLANEN Then FG.CellBackColor = COLOR_PLANEN
    Next i
End Sub
'##########################################################################################################
Private Sub SetupMnuAbwesenheit()
    If m_cAb Is Nothing Then Exit Sub   ' Beim Kontextmenueklick gespeicherte Abwesenheit
    mnuBeantragen.Visible = False:  mnuBeantragen.Enabled = False
    mnuChange.Visible = True:       mnuChange.Enabled = True
    mnuDelete.Visible = True:       mnuDelete.Enabled = True  ' alles zulassen
    Dim booZeige As Boolean
    booZeige = g_db.GetItem("ShowUserlist", "False", "in frmInfo:Show user list of absences - Default:False")
    mnuSep1.Visible = booZeige
    mnuListe.Visible = booZeige
    
    If basGlobals.g_iLCID = basGlobals.LocaleGerman Then
        mnuBeantragen.Caption = "Beantragen"
        mnuChange.Caption = "Ändern"
        mnuDelete.Caption = "Löschen"
        mnuListe.Caption = "Zeige Liste"
    Else
        mnuBeantragen.Caption = "Apply"
        mnuChange.Caption = "Change"
        mnuDelete.Caption = "Delete"
        mnuListe.Caption = "Show List"
    End If
    
    If m_cAb.lngIdxStatus = AbwStatus.PLANUNG Then
        If g_db.GetAbwesenheitsart(m_cAb.lngIdxAbwesenheitsArt).booUserBeantragbar Then
            mnuBeantragen.Visible = True:       mnuBeantragen.Enabled = True
        End If
    ElseIf m_cAb.lngIdxStatus = AbwStatus.BEANTRAGT_1 Then
    ElseIf m_cAb.lngIdxStatus = AbwStatus.BEANTRAGT_2 Then
    ElseIf m_cAb.lngIdxStatus = AbwStatus.GENEHMIGT Then
    End If
    
    Dim booEnabled As Boolean
End Sub
Private Function SetupMnuUsername() As Boolean
    If basGlobals.g_iLCID = basGlobals.LocaleGerman Then
        mnuSOY.Caption = "Zeige Tagessumme dieses Jahres"
        mnuChangeToUser.Caption = "Wechsel zu Benutzerkonto"
        mnuHome.Caption = "Wechsel zu eigenem Benutzer"
    Else
        mnuSOY.Caption = "Show sum of absence days this year"
        mnuChangeToUser.Caption = "Change to another account"
        mnuHome.Caption = "Change back to login account"
    End If

    Dim booC2U As Boolean, booSoy As Boolean, booHome As Boolean
    booC2U = CheckUserChangeOK() And m_cUs.lngIdxUser <> g_CU.lngIdxUser ' Darf ich zu Benutzer wechseln und bin ich es nicht selbst
    booSoy = CheckShowSoy()                                              ' Ist das Ziel in meiner Org?
    booHome = g_CU.lngIdxUser <> g_CU_Login.lngIdxUser                   ' Habe ich eine andere Identität?
    
    If booC2U Or booSoy Or booHome Then
        mnuUsername.Visible = False
        mnuSOY.Visible = True: mnuChangeToUser.Visible = True: mnuHome.Visible = True
        mnuChangeToUser.Visible = booC2U:     mnuChangeToUser.Enabled = booC2U: mnuHome.Visible = booHome
        mnuSOY.Visible = booSoy:              mnuSOY.Enabled = booSoy:          mnuHome.Enabled = booHome
        SetupMnuUsername = True     ' Menue zeigen
    Else
        mnuUsername.Visible = False
        SetupMnuUsername = False    ' Menue nicht zeigen
    End If
End Function

Private Sub JS()
    Dim strJahr As Integer
    strJahr = Format(Now(), "yyyy")
    strJahr = InputBox("Auswertung für welches Jahr?", App.ExeName & " V" & basGlobals.g_strProgramVersion, strJahr)
    Dim Jahr As Long
    Jahr = CInt(strJahr)
    GenJS (Jahr)
    MsgBox "Paste the content of the clipboard to an empty textfile, give it the extension .csv and open it in XL. (Or open c:\temp\Auswertung_Jahressumme.csv)", vbInformation + vbOKOnly, "kumulierte Abwesenheiten"
End Sub
Private Function GenJS(Jahr As Long) As String
    Dim fso As FileSystemObject, txStr As TextStream
    GenJS = g_db.WriteJahresSumme(Jahr)
    Set fso = New FileSystemObject
    Set txStr = fso.CreateTextFile("c:\temp\Auswertung_Jahressumme_" & Jahr & ".csv")
    txStr.Write GenJS
    txStr.Close
    Set txStr = Nothing
    Set fso = Nothing
    Clipboard.SetText GenJS
End Function
'#############################################################################################################################
Private Sub OpenUserForm()
    If g_CU_Login.booIsSek Then
        T.Enabled = False
        frmUser.Show vbModal
        T.Enabled = True
    End If
End Sub
Private Sub OpenOrgItemsForm()
    T.Enabled = False
    frmOrgItems.Show vbModal
    T.Enabled = True
End Sub
Private Sub OpenFormOptions()
        Dim f As frmOptions
        Set f = New frmOptions
        f.Show vbModal
        Set f = Nothing
        FillFG ' mit geänderten Optionen
End Sub
Private Sub ChangeLocale()
    If basGlobals.g_iLCID = basGlobals.LocaleEnglish Then
        basGlobals.g_iLCID = basGlobals.LocaleGerman
    Else
        basGlobals.g_iLCID = basGlobals.LocaleEnglish
    End If
    g_db.FillAllCollections
    g_db.setOrgUserVisibility
    Setup_Form_Info
End Sub

'#####  mnuAddon  ########################################################################################################################
Private Sub mnuAddonChangeLocale_Click()
    ChangeLocale
End Sub
Private Sub mnuAddonJahressumme_Click()
    JS
End Sub
Private Sub mnuAddonOpenOptionForm_Click()
    OpenFormOptions
End Sub
Private Sub mnuAddonOrgItem_Click()
    OpenOrgItemsForm
End Sub
Private Sub mnuAddonUser_Click()
    OpenUserForm
End Sub
Private Sub mnuaddonbildschirminfo_click()
    Dim wr As basFormUtils.RECT
    If basFormUtils.GetWorkArea(wr) Then
        wr.Top = wr.Top * Screen.TwipsPerPixelY    ' spart das spätere Umrechnen
        wr.Bottom = wr.Bottom * Screen.TwipsPerPixelY
        wr.Left = wr.Left * Screen.TwipsPerPixelX
        wr.Right = wr.Right * Screen.TwipsPerPixelX
        
        Dim f As Form
        Set f = Me
        MsgBox "BS: TBLR:" & wr.Top & "/" & wr.Bottom & "/" & wr.Left & "/" & wr.Right & vbCrLf & _
               "FM: THLW:" & f.Top & "/" & f.Height & "/" & f.Left & "/" & f.Width
    End If
End Sub
'#####  mnuAbwesenheit  ########################################################################################################################
Private Sub mnuBeantragen_Click()
    m_dtmLastAction = Now()
    g_clsAb.Beantragen m_cAb.lngIdxAbwesenheit
    g_db.FillAbwesenheitCollection  ' Daten frisch reinholen
    FillFG
End Sub
Private Sub mnuChange_Click()
    m_dtmLastAction = Now()
    m_cAb.strOutlookText = g_clsAb.GenSubject(m_cAb)    ' merken, falls OutlookAppointment geändert werden muss
    frmPlanen.EnterForm m_cAb
    frmPlanen.Show vbModal
    g_db.FillAbwesenheitCollection  ' Daten frisch reinholen
    FillFG
End Sub
Private Sub mnuDelete_Click()
    m_dtmLastAction = Now()
    g_clsAb.Zurueckziehen m_cAb.lngIdxAbwesenheit
    g_db.FillAbwesenheitCollection  ' Daten frisch reinholen
    FillFG
End Sub
Private Sub mnuListe_click()
    m_dtmLastAction = Now()
    frmBeantragen.Show vbModal
    g_db.FillAbwesenheitCollection  ' Daten frisch reinholen
    FillFG
End Sub
'#####  mnuUsername  ########################################################################################################################
Private Sub mnuHome_Click()
    m_dtmLastAction = Now()
    Set m_cUs = g_CU_Login
    ChangeToUser
End Sub
Private Sub mnuChangeToUser_Click()
    m_dtmLastAction = Now()
    ChangeToUser
End Sub
Private Sub mnuSOY_Click()
    m_dtmLastAction = Now()
    Dim Jahr As Long
    Jahr = Year(Now())
    GenJS Jahr      ' erstelle Text + PopUp-Informationen
    frmPopUp.Show vbModal  ' Tabelle wurde befüllt
End Sub
'#############################################################################################################################
Private Sub ChangeToUser()
    If CheckUserChangeOK() Then     ' Sollte überflüssig sein, weil im Kontextmenue-Build schon abgefragt
        Set g_CU = g_db.GetUserByAccountname(m_cUs.strAccountname)      ' Wechseln Benutzer - g_CU setzen aus strAccountName
        g_db.setOrgUserVisibility   ' Stelle die andere Sichtbarkeit für diesen Benutzer her.
        Arrange_Elements
        FillFG
        FG_Set_Today
    End If
End Sub

Private Function CheckUserChangeOK() As Boolean   ' DARF zu diesem m_cUs = Zielbenutzer gewechselt werden?
    ' Per Kontextmenu auf einen Benutzernamen wird dessen Identität eingenommen. Ab jetzt wird im Namen von  ... gehandelt
    ' Das dürfen Personen mit g_CU.strAccountName immer und alle Vorgesetzten nur in Richtung Mitarbeiter, nicht wieder zurück
    ' SuperTLs dürfen auch MA anderer Orgs anklicken

    Dim booSuperOL As Boolean          ' Darf OL Identität seiner MA annehmen ?
    Dim booOLSubstitute As Boolean       ' OL dürfen sich in derselben Ebene vertreten
    booSuperOL = g_db.GetItem("booSuperOL", "False", "Leader of org can enter data for members.")
    booOLSubstitute = g_db.GetItem("booOLSubstitute", "False", "Leader of org can subtitute org leader colleagues of same level in same org.")

    Dim booOK As Boolean:       booOK = False
    ' Ist dies ein TL und ( ist SuperTL eingestellt oder ist Click-User.Org = TL.Org ) ?
    If g_CU_Login.IsSekOf(m_cUs) Then    ' SEK darf immer alles
        booOK = True
    ElseIf m_cUs.lngIdxUser = g_CU_Login.lngIdxUser Then   ' zurückklicken auf eigenen Benutzer
        booOK = True                                           ' darf ich immer
    ElseIf booSuperOL And g_CU_Login.IsChefOf(m_cUs) Then ' OL darf Identität der MA annehmen - Ist der angeklickte Benutzer ein direkter MA von g_CU?
        booOK = True
    ElseIf booSuperOL And booOLSubstitute Then             ' Chefs können sich vertreten
        If g_CU_Login.IsOrgChef(m_cUs.cOrg, True) Then ' Der eingeloggte User ist ein Org-Chef oder Deputy
            Dim cC1 As clsUser
            For Each cC1 In g_db.colUser
                If cC1.lngIdxChef = g_CU_Login.lngIdxChef Or cC1.lngIdxChef = g_CU_Login.lngIdxChef2 Then  ' derselbe Chef
                    If cC1.cOrg.lngOrgLevel = g_CU_Login.cOrg.lngOrgLevel Then    ' dieselbe Hierarchiestufe
                        If cC1.IsOrgChef(m_cUs.cOrg, True) Then   ' auch OrgChef oder Deputy, also Kollege !
                            If cC1.IsChefOf(m_cUs) Then          ' ist cC1 Chef von m_cUs ?
                                booOK = True
                                Exit For    ' brauchen wir nicht weiter zu suchen
                            End If
                        End If
                    End If
                End If
            Next
        End If
    End If

    CheckUserChangeOK = booOK
End Function
Private Function CheckShowSoy() As Boolean    ' m_cUs in eigener Org? dann Kontextmenue zeigen
    Dim booOK As Boolean:       booOK = False

    If g_CU.IsSekOf(m_cUs) Then
        booOK = True                   ' Sek darf immer und überall
    ElseIf m_cUs.lngIdxUser = g_CU.lngIdxUser Then
        booOK = True                    ' target: it's me
    ElseIf g_CU_Login.IsChefOf(m_cUs) Then
        booOK = True                    ' target: member of my org
    End If
    CheckShowSoy = booOK
End Function
'#############################################################################################################################
'#############################################################################################################################

Private Sub Form_Load()
    Me.Caption = App.ExeName & " Info"
    ReadWindowPosition Me, booPositionOnly:=False
    Setup_Form_Info
    m_dtmLastAction = Now()
End Sub
Private Sub Setup_Form_Info()
    If g_db Is Nothing Then Exit Sub
    Me.Caption = g_db.GetString(1001) & " Info " & g_strProgramVersion
    m_booInit = True
    m_booOrgSlider = CBool(g_db.GetItem("booOrgSlider", "False", "Slider (True) oder DropDown(False)"))
    g_db.FillAbwesenheitCollection
    m_aCOLORS(ABW_URLAUB, COL_A_GENEHMIGT) = COLOR_URLAUB:        m_aCOLORS(ABW_URLAUB, COL_A_GEPLANT) = COLOR_URLAUB_B
    m_aCOLORS(ABW_FAKO, COL_A_GENEHMIGT) = COLOR_FAKO:            m_aCOLORS(ABW_FAKO, COL_A_GEPLANT) = COLOR_FAKO_B
    m_aCOLORS(ABW_SEMINAR, COL_A_GENEHMIGT) = COLOR_SEMINAR:      m_aCOLORS(ABW_SEMINAR, COL_A_GEPLANT) = COLOR_SEMINAR_B
    m_aCOLORS(ABW_DIENSTREISE, COL_A_GENEHMIGT) = COLOR_DIENSTREISE:  m_aCOLORS(ABW_DIENSTREISE, COL_A_GEPLANT) = COLOR_DIENSTREISE_B
    m_aCOLORS(ABW_SONSTIGES, COL_A_GENEHMIGT) = COLOR_SONSTIGES:  m_aCOLORS(ABW_SONSTIGES, COL_A_GEPLANT) = COLOR_SONSTIGES_B
    m_aCOLORS(ABW_FEIERTAG, COL_A_GENEHMIGT) = COLOR_FEIERTAG:    m_aCOLORS(ABW_FEIERTAG, COL_A_GEPLANT) = COLOR_FEIERTAG
    m_aCOLORS(ABW_SCHULFERIEN, COL_A_GENEHMIGT) = COLOR_SCHULFERIEN: m_aCOLORS(ABW_SCHULFERIEN, COL_A_GEPLANT) = COLOR_SCHULFERIEN
    m_aCOLORS(ABW_KRANKHEIT, COL_A_GENEHMIGT) = COLOR_KRANKHEIT:   m_aCOLORS(ABW_KRANKHEIT, COL_A_GEPLANT) = COLOR_KRANKHEIT

    lblText(ABW_URLAUB).Caption = g_db.GetString(1031)
    lblText(ABW_FAKO).Caption = g_db.GetString(1032)
    lblText(ABW_SEMINAR).Caption = g_db.GetString(1033)
    lblText(ABW_DIENSTREISE).Caption = g_db.GetString(1034)
    lblText(ABW_SONSTIGES).Caption = g_db.GetString(1035)
    lblText(ABW_FEIERTAG).Caption = g_db.GetString(1036)
    lblText(ABW_SCHULFERIEN).Caption = g_db.GetString(1012)
    lblText(ABW_KRANKHEIT).Caption = g_db.GetString(1037)

    lblText(ABW_URLAUB).ToolTipText = g_db.GetString(1038)
    lblText(ABW_FAKO).ToolTipText = g_db.GetString(1038)
    lblText(ABW_SEMINAR).ToolTipText = g_db.GetString(1038)
    lblText(ABW_DIENSTREISE).ToolTipText = g_db.GetString(1038)
    lblText(ABW_SONSTIGES).ToolTipText = g_db.GetString(1038)

    lblHeader.Caption = g_db.GetString(1039)
    Me.btnExit.ToolTipText = g_db.GetString(1107)
    Me.btnSecurity.ToolTipText = g_db.GetString(1146)
    Me.btnHelp.ToolTipText = g_db.GetString(1147)
    Me.btnGenehmigen.ToolTipText = g_db.GetString(1148)

    '###############   FG Setup  ########################
    FG.MergeCells = flexMergeFree
    FG.MergeRow(ROW_MONAT) = True       ' Monate
    FG.MergeRow(ROW_KW) = True       ' KW

    i = g_db.GetItem("WidthUserColumn", 2000, "Minimum Width of first column in frmInfo: username")
    If i < 2000 Then i = 2000   ' Minimum ColWidth
    FG.ColWidth(0) = i    ' Name
    '#########   Sichtbarkeit 2. Spalte - SortierStrings    ###############################################
    i = 0   ' default: not visible
    If g_booShowSort Then
        i = g_db.GetItem("lngBreiteSort", 500, "Width of sort column")
    End If
    FG.ColWidth(1) = i
    '#########   Sichtbarkeit 2. Spalte - SortierStrings    ###############################################

    FG.RowHeight(0) = 300       ' Überschriftenzeile
    '###############   FG Setup  Ende ########################

    Dim lngUpperLimit As Long
    If g_db.KrankheitUserBeantragbar Then      ' Sek, AL, TL...
        lngUpperLimit = ABW_KRANKHEIT
    Else
        lngUpperLimit = ABW_SCHULFERIEN
    End If
    
    Dim lngLeft As Long
    lngLeft = 12
    For i = ABW_URLAUB To lngUpperLimit
        lblText(i).Left = lngLeft
        lblText(i).Width = 90
        lblText(i).Visible = True
        lblTyp(i).Left = lngLeft
        lblTyp(i).Width = 45
        lblTyp(i).BackColor = m_aCOLORS(i, COL_A_GEPLANT)
        lblTyp(i).Visible = True
        lblTyp(i + ABW_KRANKHEIT).Left = lngLeft + 45
        lblTyp(i + ABW_KRANKHEIT).Width = 45
        lblTyp(i + ABW_KRANKHEIT).BackColor = m_aCOLORS(i, COL_A_GENEHMIGT)
        lblTyp(i + ABW_KRANKHEIT).Visible = True
        lngLeft = lngLeft + 100
    Next i
    Arrange_Elements
    m_booNurAbwesenheit = False   ' Normale Ansicht
    m_booInit = False     ' Bis hierher war noch kein FILL_FG erlaubt
    FillCboBundesländer
    FillFG
    FG_Set_Today
End Sub

Private Sub Arrange_Elements()
Dim dtmErstesQ As Date
Dim ii As Long
Dim lQ As Long, lM As Long, lJ As Long
    Dim AnzQuartale As Integer, QuartaleZurueck As Integer
    AnzQuartale = 8
    QuartaleZurueck = 3
    ' Scroll: Ein Slider für 8 Quartale
    ' QuartaleZurueck Q vor aktuellem Q und AnzQuartale-4 Q voraus     = AnzQuartale Quartale
    dtmErstesQ = Date - QuartaleZurueck * 365 / 4
    lQ = Fix((Format(dtmErstesQ, "MM") - 1) / 3) + 1 ' Anfangsquartal
    
    lM = (lQ - 1) * 3  ' Anfangsmonat 0,3,6,9
    lJ = Format(dtmErstesQ, "yyyy")  ' Anfangsjahr
    g_Quartal(1, COL_QUARTAL) = lQ:    g_Quartal(1, COL_MONAT) = lM:    g_Quartal(1, COL_JAHR) = lJ
    For ii = 2 To AnzQuartale
        lQ = lQ + 1:    lM = lM + 3
        If lQ = 5 Then
            lQ = 1: lM = 0: lJ = lJ + 1
        End If
        g_Quartal(ii, COL_QUARTAL) = lQ:    g_Quartal(ii, COL_MONAT) = lM:    g_Quartal(ii, COL_JAHR) = lJ
    Next ii

    ScrollQuartal.Min = 1:  ScrollQuartal.Max = AnzQuartale
    For ii = ScrollQuartal.Max To ScrollQuartal.Min + 1 Step -1
        If DateSerial(g_Quartal(ii, COL_JAHR), g_Quartal(ii, COL_MONAT), 1) < Date Then Exit For
    Next ii
    ScrollQuartal.value = ii
'''     ii = GetScrollStartMonat    ' 0,3,6,9
'''     If Month(Now()) - ii = 3 Then ScrollQuartal.Value = ScrollQuartal.Value + 1 ' In den Monaten 3,6,9,12 gleich einen weiter

    cboOrg.Visible = False
    ScrollOrg.Visible = False
    If m_booOrgSlider Then
        ScrollOrg.Min = 1:     ScrollOrg.Max = g_db.maxOrgID:    ScrollOrg.value = g_db.GetOrgPosition(g_CU.lngIdxOrg)

        If (g_CU.lngVerteiler >= SICHTBAR_ABT) Then ' And Not (g_CU.lngUserLevel <= Abteilung And g_booAlleMaSichtbar) Then   ' Scrollen in Org-Einheiten: ab AL-Sekretariat immer
            If ScrollOrg.Max > 1 Then ScrollOrg.Visible = True
        End If
    Else
        Dim cOrg As clsOrg
        Me.cboOrg.Clear
        For Each cOrg In g_db.colOrg
            If cOrg.booVisible Then
                If g_booAlleMaSichtbar And g_CU.lngUserLevel <= Abteilung Then ' keine Teams eintragen
                    If cOrg.lngOrgLevel <= basGlobals.OrgLevel.Abteilung Then
                        Me.cboOrg.AddItem cOrg.strOrg
                    End If
                Else
                    Me.cboOrg.AddItem cOrg.strOrg
                End If
            End If
        Next
        ' Auswahl sichtbar, wenn Sichtbarkeit >= SichtbatAbt(3) (Chef(1) und Team(2) NICHT) für alle
        ' Auswahl NICHT sichtbar, wenn Level <= AL und g_booAlleMaSichtbar
        If (g_CU.lngVerteiler >= SICHTBAR_ABT) Then     ' And Not (g_CU.lngUserLevel <= Abteilung And g_booAlleMaSichtbar) Then
            If cboOrg.ListCount > 1 Then cboOrg.Visible = True
        End If
        cboOrg = g_CU.cOrg.strOrg ' HeimatOrg des Logins setzen
    End If
    m_lngIdxOrg = g_CU.lngIdxOrg
    
    ' Genehmigen + Übersicht - ab TL
    If g_CU.IsPrivileged And g_CU.lngIdxUser = g_CU_Login.lngIdxUser Then ' nur mit eigener Identität oder Sek
        btnGenehmigen.Visible = True
        btnNurAbwesend.Visible = True
    Else
        btnGenehmigen.Visible = False
        btnNurAbwesend.Visible = False
    End If
    
    m_booSelected = False
    T.Interval = 60000
    T.Enabled = True
End Sub

Private Sub Form_Resize()
Dim l As Long
    m_dtmLastAction = Now()
    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.Width < 12450 Then Me.Width = 12450
    If Me.Height < 9000 Then Me.Height = 9000
    ' ScaleHeight=573   TopText=544 TopLblTyp=545  TopFG=76 HeightFG=449 = Scaleheight-124
    For i = ABW_URLAUB To ABW_KRANKHEIT
        lblText(i).Top = Me.ScaleHeight - 29
        lblTyp(i).Top = Me.ScaleHeight - 28
        lblTyp(i).Height = 20
        lblTyp(i + ABW_KRANKHEIT).Top = lblTyp(i).Top
        lblTyp(i + ABW_KRANKHEIT).Height = lblTyp(i).Height
    Next i
    FG.Height = Me.ScaleHeight - 124
    'Knöpfe ...
    l = btnExit.Width + 1   ' Breite der Knöpfe
    btnExit.Left = Me.ScaleWidth - l - 9:                   btnExit.Top = 12
    btnGenehmigen.Left = Me.ScaleWidth - 2 * l - 9:         btnGenehmigen.Top = btnExit.Top
    
    btnHelp.Left = btnExit.Left:                            btnHelp.Top = Me.ScaleHeight - 41
    btnSecurity.Left = btnGenehmigen.Left:                  btnSecurity.Top = btnHelp.Top
    btnNurAbwesend.Left = Me.ScaleWidth - 3 * l - 9:       btnNurAbwesend.Top = btnSecurity.Top
    ' Labels und Scrolls
    lblQuartal.Left = btnGenehmigen.Left - lblQuartal.Width - 3:   ScrollQuartal.Left = lblQuartal.Left
    lblOrg.Left = lblQuartal.Left - lblOrg.Width - 3:            ScrollOrg.Left = lblOrg.Left:  cboOrg.Left = lblOrg.Left
    lblHeader.Width = lblOrg.Left - lblHeader.Left - 3
    cboBundesland.Left = lblHeader.Left
    cboBundesland.Top = lblHeader.Top
    cboBundesland.Visible = False
    ' FG
    FG.Width = Me.ScaleWidth - FG.Left - 9
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWindowPosition Me
    RaiseEvent CloseApplication
End Sub
Private Function GetScrollJahr() As Long
    GetScrollJahr = g_Quartal(ScrollQuartal, COL_JAHR)
End Function
Private Function GetScrollStartMonat() As Long
    GetScrollStartMonat = g_Quartal(ScrollQuartal, COL_MONAT)
End Function
Private Function GetScrollQuartal() As Long
    GetScrollQuartal = g_Quartal(ScrollQuartal, COL_QUARTAL)
End Function
Private Sub FillFG()
Dim lngJahr As Long, lngMonat As Long, lngWeeks As Long
Dim lngCounter As Long, lngOldLevel As Long, strOldSortOrder As String, lngOldOrg As Long
Dim varval As Variant, booTrennerOrgs As Boolean, lngAnzTrenner As Long
Dim cF As clsFeiertag, cSF As clsSchulferien
Dim lngBackColor As Long, lngIdxFeiertag As Long
Dim lngNextLine As Long
Dim cAb As clsAbw, cUs As clsUser, strDisclaimer As String, strVerteiler As String
Dim cAbt As clsAbw
    Dim OldColor As Long
    OldColor = Me.lblQuartal.BackColor

    If g_db Is Nothing Then Exit Sub
    If m_booInit Then Exit Sub
    Me.MousePointer = vbHourglass
    Me.lblQuartal.BackColor = vbRed
    
    If m_booNurAbwesenheit Then   ' Alle Abwesenheiten innerhalb xx Wochen ab heute
        lngJahr = Year(Now())
        lngMonat = Month(Now())
        m_dtmStart = Date
        m_dtmEnde = Date + VIEW_WOCHEN * 7 - 1
        lngWeeks = VIEW_WOCHEN
    Else
        '################    Monate/Wochen/Tage    ################################################
        ' Es kann Quartalsweise die Mitte des Beobachtungsraums gesetzt werden
        ' Von diesem Quartal wird zusätzlich ein Monat vorher und nachher gezeigt
        lngJahr = GetScrollJahr
        lngMonat = GetScrollStartMonat  ' 0,3,6,9
        If lngMonat = 0 Then
            lngMonat = 12
            lngJahr = lngJahr - 1
        End If
        m_dtmStart = DateSerial(lngJahr, lngMonat, 1)

        lngMonat = lngMonat + 4  ' 4,7,10,13
        If lngMonat > 12 Then
            lngMonat = lngMonat - 12
            lngJahr = lngJahr + 1
        End If
        m_dtmEnde = DateSerial(lngJahr, lngMonat + 1, 1) - 1  ' der letzte Tag des Monats
        lngWeeks = DateDiff("w", m_dtmStart, m_dtmEnde) + 1
    End If

    ' FG malen - los geht's
    FG.Redraw = False
    On Error Resume Next
    FG.FixedRows = 3    ' Monat, KW, Tag             ( Damit Merge funktioniert  ???)
    FG.FixedCols = 2    ' Username, Sortierung
    On Error GoTo 0

    FG.Rows = ROW_TAG + 1:    FG.Cols = 2     ' Lösche alle Inhalte, aber Zeitskala nicht

    FG.TextMatrix(ROW_MONAT, 0) = g_db.GetString(1040)
    FG.TextMatrix(ROW_KW, 0) = g_db.GetString(1041)
    FG.TextMatrix(ROW_TAG, 0) = ""
'################    Monate/Wochen/Tage    ################################################
    FG.Cols = lngWeeks * 7 + 7
'''sglCounter = Timer()
    m_dtmRun = m_dtmStart
    lngCounter = OFFSET_FGDATE      ' Hier fängt im FG das Datum an
    FG.Row = ROW_MONAT
    While m_dtmRun <= m_dtmEnde
        FG.col = lngCounter
        FG.TextMatrix(ROW_MONAT, lngCounter) = Format(m_dtmRun, "mmmm yyyy")   ' Monat
        FG.TextMatrix(ROW_KW, lngCounter) = g_db.GetString(1041) & Format(m_dtmRun, "ww", g_db.FirstWorkingDay, vbFirstFourDays)
        FG.TextMatrix(ROW_TAG, lngCounter) = Format(m_dtmRun, "d")
        If Weekday(m_dtmRun) = g_db.WeekendFirstDay Or Weekday(m_dtmRun) = g_db.WeekendSecondDay Then ' Samstag oder Sonntag
            If g_booPlanEveryDay Then
                FG.ColWidth(lngCounter) = 110
            Else
                FG.ColWidth(lngCounter) = 20
            End If
        Else
            FG.ColWidth(lngCounter) = 270
        End If
        If m_dtmRun = Date Then
            FG.Row = ROW_TAG
            FG.CellFontBold = True
            FG.CellFontSize = 9.75
            FG.ColWidth(lngCounter) = FG.ColWidth(lngCounter) + 45
        End If
        If Format(m_dtmRun, "ww") = Format(Date, "ww") Then
            FG.Row = ROW_KW
            FG.CellFontBold = True
            FG.CellFontSize = 9.75
        End If
        If Format(m_dtmRun, "MM") = Format(Date, "MM") Then
            FG.Row = ROW_MONAT
            FG.CellFontBold = True
            FG.CellFontSize = 9.75
        End If
        m_dtmRun = m_dtmRun + 1
        lngCounter = lngCounter + 1
    Wend
    FG.Cols = lngCounter    ' Der endgültige Wert
'''Debug.Print Timer() - sglCounter
'################    Monate/Wochen/Tage    ################################################

'#########   Userdaten   ##############################################################################
    ReDim m_aUserHash(0)  ' Alles löschen
    ReDim m_aUserHash(g_db.MaxUserID)
    ReDim m_aFG(0, 0)
    ReDim m_aFG(FG.Cols, 5)
    lngNextLine = ROW_TAG + 1
    
    If m_booNurAbwesenheit Then   ' User werden nach Abwesenheitsdatum geordnet ...
        i = 4
        Dim cAb1 As clsAbw
        ' in colAbwesenheiten sind die Abwesenheiten nach Startdatum geordnet.
        ' Ziel in dieser Ansicht: Startzeiten, wobei heute der früheste Zeitpunkt ist
        '       1. Tag der Anwesenheit
        Dim colAbw As Collection
        Dim booEingetragen As Boolean, dStart As Date, dStart1 As Date
        Set colAbw = New Collection
        For Each cAb In g_db.colAbwesenheiten   ' Alle Abwesenheiten in neue Collection eintragen
            If cAb.dtmStart >= m_dtmStart And cAb.dtmStart <= m_dtmEnde Or _
               cAb.dtmEnde >= m_dtmStart And cAb.dtmEnde <= m_dtmEnde Then
               ' Eintrag ist relevant
               booEingetragen = False
'''               Debug.Print "..."
'''               Debug.Print cAb.dtmStart, cAb.dtmEnde, cAb.lngIdxAbwesenheit, cAb.strAbwesenheitsart
                If colAbw.Count = 0 Then    ' noch leer
'''                    Debug.Print "Start."
                    colAbw.Add cAb, "A_" & cAb.lngIdxAbwesenheit      ' einfach eintragen
                Else
                    ' Sortieren:
                    ' bisher eingetragene durchgehen und feststellen, ob aktuelle Abw vor einer bestehenden
                    ' eingetragen werden muss. Falls die Abw zum Schluss noch nicht eingetragen wurde, dann
                    ' am Ende anhängen.

                    For Each cAb1 In colAbw ' alle eingetragenen Abw durchgehen
                    ' Wenn aktuelle Abw später beginnt und später endet, dann eintragen
                        dStart = IIf(cAb.dtmStart > m_dtmStart, cAb.dtmStart, m_dtmStart)
                        dStart1 = IIf(cAb1.dtmStart > m_dtmStart, cAb1.dtmStart, m_dtmStart)
                        If dStart <= dStart1 Then
                            If cAb.dtmEnde < cAb1.dtmEnde Then
'''                               Debug.Print "vor " & cAb1.lngIdxAbwesenheit
                                colAbw.Add cAb, "A_" & cAb.lngIdxAbwesenheit, "A_" & cAb1.lngIdxAbwesenheit       ' Eintrag "before" cAb1
                                GoTo nextCab
                            End If
                        End If
                    Next
                    If Not booEingetragen Then  ' ganz hinten dranhängen
'''                        Debug.Print "Dranhängen"
                        colAbw.Add cAb, "A_" & cAb.lngIdxAbwesenheit
                    End If
                End If
            End If
nextCab:
        Next cAb
        For Each cAb In colAbw         ' Alle Abw datumssortiert durchlaufen
                ' Darstellen
                m_dtmRun = IIf(cAb.dtmStart > m_dtmStart, cAb.dtmStart, m_dtmStart)
                If m_aUserHash(cAb.lngIdxUser) = 0 Then       ' Benutzer noch nicht angelegt
                    FG.Rows = lngNextLine + 1               ' ... es kommt jemand dazu
                    i = i + 1                               ' Zähler für die Trennlinien
                    If i = 5 Then
                        FG.RowHeight(lngNextLine) = 21      ' Trennlinie zeichnen
                        lngNextLine = lngNextLine + 1       ' ... plus ein PseudoUser
                        FG.Rows = lngNextLine + 1           ' ... und Tabelle auch eine Zeile mehr
                        i = 0
                    End If
                    ReDim Preserve m_aFG(FG.Cols, FG.Rows)    ' Tabelle wird größer ...
                    m_aUserHash(cAb.lngIdxUser) = lngNextLine ' Schneller Zugriff von User auf Zeile
                    m_aFG(0, lngNextLine) = cAb.lngIdxUser    ' Schneller Zugriff von Zeile auf User
                    Set cUs = g_db.GetUserByID(cAb.lngIdxUser)
                    FG.TextMatrix(lngNextLine, 0) = cUs.strNachname & ", " & cUs.strVorname
                    FG.TextMatrix(lngNextLine, 1) = cUs.strSortOrder
                    lngNextLine = lngNextLine + 1
                End If
                FG.Row = m_aUserHash(cAb.lngIdxUser)
                If cAb.lngIdxStatus = AbwStatus.GENEHMIGT Then      ' genehmigt
                    lngBackColor = m_aCOLORS(cAb.lngIdxAbwesenheitsArt, COL_A_GENEHMIGT)
                Else ' geplant, beantragt0 bis 2
                    lngBackColor = m_aCOLORS(cAb.lngIdxAbwesenheitsArt, COL_A_GEPLANT)
                End If
                While m_dtmRun <= cAb.dtmEnde And m_dtmRun <= m_dtmEnde    ' Bis zum Ende der Abwesenheit, maximal zum Ende des Zeitraums
                        FG.col = m_dtmRun - m_dtmStart + OFFSET_FGDATE
                        m_aFG(FG.col, FG.Row) = cAb.lngIdxAbwesenheit

                        If FG.CellBackColor = 0 Or FG.CellBackColor = COLOR_SONSTIGES Then  ' "leer" -> Zellen überschreiben
                            FG.ForeColor = vbBlack
                            FG.CellBackColor = lngBackColor  ' Nur schreiben, wenn leer (Feiertage, Muss-Abwesenheit)
                            If cAb.strGVN <> "" Then FG.Text = cAb.strGVN: FG.CellFontBold = True

                            ' Im Genehmigungsprozess
                            If cAb.lngIdxStatus = AbwStatus.BEANTRAGT_1 Then
                                FG.Text = FG.Text & "/"
                            ElseIf cAb.lngIdxStatus = AbwStatus.BEANTRAGT_2 Then
                                FG.Text = FG.Text & "//"
                            End If
                        End If

                    m_dtmRun = m_dtmRun + 1
                Wend
'            End If  ' Datum
        Next cAb
    Else    ' User werden nach Org sortiert
        lngOldLevel = -1 ' Das Anfangslevel, es geht abwärts
        lngOldOrg = -1
        strOldSortOrder = ""

        booTrennerOrgs = CBool(g_db.GetItem("booTrennerOrgs", "True", "Separate org with double lines"))

        Dim colUser As Collection
        Set colUser = g_db.GetUserCollection(m_lngIdxOrg)   ' Definiere die Benutzer, die bei dieser Org-Anwahl erscheinen sollen
        For Each cUs In colUser
            If (cUs.dtmFirst <= m_dtmStart And cUs.dtmLast >= m_dtmStart) Or (cUs.dtmFirst <= m_dtmEnde And cUs.dtmLast >= m_dtmEnde) Then
                FG.Rows = lngNextLine + 1
                lngAnzTrenner = 0
                ' Trenner bei Level-Wechsel
                If cUs.lngUserLevel <> lngOldLevel Then
                    lngOldLevel = cUs.lngUserLevel          ' Level hat gewechselt - Merken
                    FG.RowHeight(lngNextLine) = 15      ' Trennlinie 1
                    lngNextLine = lngNextLine + 1
                    FG.Rows = lngNextLine + 1
                End If

                ' Trenner bei SortWechsel
                If basGlobals.g_booShowSort Then
                    If cUs.strSortOrder <> strOldSortOrder Then
                        strOldSortOrder = cUs.strSortOrder  ' Sortierung hat gewechselt - Merken
                        FG.RowHeight(lngNextLine) = 15      ' Trennlinie
                        lngNextLine = lngNextLine + 1
                        FG.Rows = lngNextLine + 1
                    End If
                End If

                ' Spezielle Trenner bei OrgWechsel
                If booTrennerOrgs And cUs.lngIdxOrg <> lngOldOrg Then
                    lngOldOrg = cUs.lngIdxOrg     ' Org hat gewechselt - Merken
                    If cUs.lngUserLevel = Benutzer Then
                        FG.RowHeight(lngNextLine) = 30      ' Trennlinie
                        lngNextLine = lngNextLine + 1
                        FG.Rows = lngNextLine + 1
                    ElseIf cUs.lngUserLevel = Team Then
                        FG.RowHeight(lngNextLine) = 30      ' Trennlinie
                        lngNextLine = lngNextLine + 1
                        FG.Rows = lngNextLine + 1
                    End If
                End If
    
                ' Userdaten eintragen
                    ReDim Preserve m_aFG(FG.Cols, FG.Rows)     ' Erweitern
                m_aUserHash(cUs.lngIdxUser) = lngNextLine ' Schneller Zugriff von User auf Zeile
                m_aFG(0, lngNextLine) = cUs.lngIdxUser    ' Schneller Zugriff von Zeile auf User
                strDisclaimer = IIf(cUs.dtmDisclaimer <> 0, "", "d")
                
                Select Case cUs.lngVerteiler
                    Case basGlobals.SICHTBAR_CHEF
                        strVerteiler = "#"
                    Case basGlobals.SICHTBAR_Org
                        strVerteiler = ""
                    Case basGlobals.SICHTBAR_ABT
                        strVerteiler = "_"
                    Case basGlobals.SICHTBAR_BER
                        strVerteiler = "~"
                    Case SICHTBAR_RES
                        strVerteiler = ":"
                    Case Else
                        strVerteiler = ""
                End Select

                FG.TextMatrix(lngNextLine, 0) = Liste(" ", cUs.lngUserLevel) & cUs.strNachname & ", " & cUs.strVorname & " " & strVerteiler & strDisclaimer
                FG.TextMatrix(lngNextLine, 1) = cUs.strSortOrder
                If cUs.lngIdxUser = g_CU.lngIdxUser Then
                    FG.Row = lngNextLine
                    FG.col = 0
                    FG.CellFontBold = True
                End If
                lngNextLine = lngNextLine + 1
            End If  ' im Zeitraum?
        Next cUs
        FG.Rows = FG.Rows + 1
        FG.RowHeight(FG.Rows - 1) = 15  ' Die letzte Zeile als Trenner
    End If

'#########   Feiertage   ##############################################################################
    For Each cF In g_db.colFeiertage
        If cF.dtmFeiertag >= m_dtmStart And cF.dtmFeiertag <= m_dtmEnde Then
            Select Case cF.lngIdxAbwesenheitsArt
                Case ABW_URLAUB  ' Urlaub
                    lngBackColor = COLOR_URLAUB
                Case ABW_FAKO  ' FAKO
                    lngBackColor = COLOR_FAKO
                Case ABW_FEIERTAG  ' Feiertag
                    lngBackColor = COLOR_FEIERTAG
                Case ABW_SONSTIGES
                    lngBackColor = COLOR_SONSTIGES
            End Select
            lngIdxFeiertag = cF.lngIdxFeiertag
            FG.col = cF.dtmFeiertag - m_dtmStart + OFFSET_FGDATE
            For i = ROW_TAG To FG.Rows - 1
                FG.Row = i
                If lngBackColor <> COLOR_SONSTIGES Or m_aFG(FG.col, i) = 0 Then
                    FG.CellBackColor = lngBackColor
                    m_aFG(FG.col, i) = -lngIdxFeiertag     ' Feiertage sind negativ dargestellt
                End If
            Next i
        End If
    Next cF
'#########   Feiertage   ##############################################################################
'#########   SchulFerien ##############################################################################
    For Each cSF In g_db.colSchulferien
        If (cSF.dtmStart >= m_dtmStart And cSF.dtmStart <= m_dtmEnde Or _
            cSF.dtmEnde >= m_dtmStart And cSF.dtmEnde <= m_dtmEnde) And _
            cboBundesland.Text = cSF.strBundesland Then
            m_dtmRun = cSF.dtmStart
            If m_dtmRun < m_dtmStart Then m_dtmRun = m_dtmStart
            While m_dtmRun <= cSF.dtmEnde And m_dtmRun <= m_dtmEnde    ' Bis zum Ende der Abwesenheit, maximal zum Ende des Zeitraums
                FG.col = m_dtmRun - m_dtmStart + OFFSET_FGDATE
                FG.Row = ROW_TAG
                FG.CellBackColor = COLOR_SCHULFERIEN
                m_dtmRun = m_dtmRun + 1
            Wend
        End If
    Next cSF
    FG.TextMatrix(2, 0) = "(" & cboBundesland.Text & ")":    FG.col = 0:    FG.Row = 2:    FG.CellFontSize = 4
'#########   SchulFerien  ##############################################################################

'#########   Userdaten   ##############################################################################
'#########   Abwesenheiten   ##########################################################################
    If m_booNurAbwesenheit = False Then
        Dim ShowAbwesenheit As Boolean: ShowAbwesenheit = False ' Zeige Abwesenheit
        For Each cAb In g_db.colAbwesenheiten
            If UBound(m_aUserHash) >= cAb.lngIdxUser Then
                If m_aUserHash(cAb.lngIdxUser) <> 0 Then   ' Diesen Benutzer darstellen
                    Set cUs = g_db.GetUserByID(cAb.lngIdxUser)
                    If g_db.UserSeesUser(cUs, g_CU) Then
                        ' Start oder Ende innerhalb des Anzeigezeitraums oder Start vor und Ende nach Anzeigezeitraum
                        If cAb.dtmStart >= m_dtmStart And cAb.dtmStart <= m_dtmEnde Or _
                           cAb.dtmEnde >= m_dtmStart And cAb.dtmEnde <= m_dtmEnde Or _
                           cAb.dtmStart <= m_dtmStart And cAb.dtmEnde >= m_dtmEnde _
                           Then
                            ' Darstellen
                            ' Krankheit nur für: Superuser, AL und u.U. TL
                            If cAb.lngIdxAbwesenheitsArt = ABW_KRANKHEIT Then
                                ShowAbwesenheit = g_db.KrankheitSichtbar(cAb)    ' Sek, AL, TL...
                            ElseIf cAb.lngIdxAbwesenheitsArt <> ABW_SCHULFERIEN Then
                                ShowAbwesenheit = True
                            End If  ' Krankheit
                            
                            If ShowAbwesenheit Then
                                m_dtmRun = cAb.dtmStart
                                If m_dtmRun < m_dtmStart Then m_dtmRun = m_dtmStart
                                FG.Row = m_aUserHash(cAb.lngIdxUser)
                                If cAb.lngIdxStatus = AbwStatus.GENEHMIGT Then      ' genehmigt
                                    lngBackColor = m_aCOLORS(cAb.lngIdxAbwesenheitsArt, COL_A_GENEHMIGT)
                                Else ' geplant, beantragt0 bis 2
                                    lngBackColor = m_aCOLORS(cAb.lngIdxAbwesenheitsArt, COL_A_GEPLANT)
                                End If
                                While m_dtmRun <= cAb.dtmEnde And m_dtmRun <= m_dtmEnde    ' Bis zum Ende der Abwesenheit, maximal zum Ende des Zeitraums
                                    FG.col = m_dtmRun - m_dtmStart + OFFSET_FGDATE
                                    m_aFG(FG.col, FG.Row) = cAb.lngIdxAbwesenheit
            
                                    If FG.CellBackColor <> COLOR_FEIERTAG Then   ' "leer" -> Zellen überschreiben
                                        FG.ForeColor = vbBlack
                                        FG.CellBackColor = lngBackColor  ' Nur schreiben, wenn leer (Feiertage, Muss-Abwesenheit)
                                        If cAb.strGVN <> "" Then FG.Text = cAb.strGVN: FG.CellFontBold = True
            
                                        ' Im Genehmigungsprozess
                                        If cAb.lngIdxStatus = AbwStatus.BEANTRAGT_1 Then
                                            FG.Text = FG.Text & "/"
                                        ElseIf cAb.lngIdxStatus = AbwStatus.BEANTRAGT_2 Then
                                            FG.Text = FG.Text & "//"
                                        End If
                                    End If
            
                                    m_dtmRun = m_dtmRun + 1
                                Wend
                            End If  ' ShowAbwesenheit
                        End If  ' Datum
                    Else
                        'Stop
                    End If  ' UserSeesUser
                Else
                    'Stop
                End If  ' User im m_aUserHash ?
            Else
                'Stop
            End If  ' Hashtabelle groß genug?
        Next cAb
    End If
'#########   Abwesenheiten   ##########################################################################

    If m_selectedLine > 0 Then
        FG.RowHeight(m_selectedLine) = FG.RowHeight(m_selectedLine) * 1.5
    End If

    FG.Redraw = True
    Me.MousePointer = vbNormal
    Me.lblQuartal.BackColor = vbGreen
    Sleep 100
    Me.lblQuartal.BackColor = OldColor
End Sub
Private Sub FG_Set_Today()
    On Error Resume Next
    If Date >= m_dtmStart And Date <= m_dtmEnde Then    ' Sind wir im laufenden Quartal ?
        FG.col = Date - m_dtmStart + OFFSET_FGDATE    ' Setze Col auf HEUTE
        FG.LeftCol = FG.col
    End If
End Sub

Private Sub FillCboBundesländer()
    Dim strLänder As String, aLänder() As String, i As Integer
    strLänder = g_db.Bundesländer()
    cboBundesland.Clear
    If strLänder <> "" Then
        aLänder = Split(strLänder, "#")
        For i = LBound(aLänder) To UBound(aLänder)
            cboBundesland.AddItem aLänder(i)
        Next i
    End If
    i = basFormUtils.GetRegData("Bundesland")
    If i >= 0 Then
        On Error GoTo Set_0
        cboBundesland.ListIndex = i
    Else
Set_0:
        cboBundesland.ListIndex = -1     ' Erster Eintrag
    End If
End Sub

Private Sub lblHeader_DblClick()
    ' Ferien Bundesland verstellen ... per cboBundesland
    cboBundesland.Visible = True
End Sub

'#############################################################################################################################
Private Sub ScrollQuartal_Change()
    m_dtmLastAction = Now()
    lblQuartal.Caption = "Q" & GetScrollQuartal & "/" & GetScrollJahr
    FillFG
    FG_Set_Today
End Sub

Private Sub ScrollOrg_Change()
    m_dtmLastAction = Now()
    m_lngIdxOrg = g_db.GetOrgIdxAtScrollPosition(ScrollOrg)
    lblOrg.Caption = g_db.colOrg("cOrg_" & m_lngIdxOrg).strOrg
    FillFG
    FG_Set_Today
End Sub
Private Sub cboOrg_Change()
    m_dtmLastAction = Now()
    m_lngIdxOrg = g_db.GetOrgIndexFromOrgName(cboOrg)
End Sub
Private Sub cboOrg_Click()
    m_dtmLastAction = Now()
    If cboOrg.DataChanged Then
        m_lngIdxOrg = g_db.GetOrgIndexFromOrgName(cboOrg)
        FillFG
        FG_Set_Today
    End If
End Sub

Private Sub T_Timer()
    If m_booSelected Then Exit Sub  ' kein Refresh beim Termin-Eintragen
    
    If Not g_db Is Nothing Then ' nachsehen: DB-Signal
        If CBool(g_db.GetItem("booNoRun", "False", "Stop application")) Then g_StopApp = "NoRun"
    End If
    If Now - m_dtmLastAction > TimeSerial(6, 0, 0) Then g_StopApp = "Idle" ' 6 Stunden nichts gemacht.
    
    If g_StopApp <> "" Then
        ' Ist dies das oberste Fenster ?
        If Forms.Count > 1 Then
            Dim iFenster As Integer
            For iFenster = 1 To Forms.Count - 1
                Unload Forms(iFenster)
            Next
        End If
        Unload Me
        Exit Sub
    End If
    
    If Not g_db Is Nothing Then g_db.FillAbwesenheitCollection  ' Daten frisch reinholen
    FillFG
End Sub
