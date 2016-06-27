VERSION 5.00
Begin VB.Form frmSecurity 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Sichtbarkeit"
   ClientHeight    =   9495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9495
   ScaleWidth      =   9195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton btnExit 
      Cancel          =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   8160
      Picture         =   "frmSecurity.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   3
      ToolTipText     =   "Programm verlassen"
      Top             =   240
      Width           =   795
   End
   Begin VB.CommandButton btnUpgrade 
      Height          =   1695
      Left            =   6960
      Picture         =   "frmSecurity.frx":014A
      Style           =   1  'Grafisch
      TabIndex        =   2
      ToolTipText     =   "Nächst höhere Sichtbarkeitsstufe"
      Top             =   7560
      Width           =   2055
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Zentriert
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Es geht um die Sichtbarkeit Ihrer Abwesenheitszeiten"
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
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   7695
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Label1"
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
      TabIndex        =   1
      Top             =   7560
      Width           =   6735
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblInfo 
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   8895
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub btnUpgrade_Click()
Dim varval As Variant
Dim strMessage As String
    strMessage = g_db.GetString(1125) '  Sie möchten zur Ebene wechseln: Sichtbarkeit auf
    Select Case g_CU.lngVerteiler
        Case basGlobals.SICHTBAR_CHEF
            strMessage = strMessage & g_db.GetString(1126)   ' auf Teamebene
        Case basGlobals.SICHTBAR_Org
            strMessage = strMessage & g_db.GetString(1127)   ' auf Abteilungsebene
        Case basGlobals.SICHTBAR_ABT
            strMessage = strMessage & g_db.GetString(1128)   ' auf Bereichsebene
        Case basGlobals.SICHTBAR_BER
            strMessage = strMessage & g_db.GetString(1142)   ' auf Ressortebene
    End Select
    strMessage = strMessage & """ ?"
    varval = MsgBox(strMessage, vbYesNo + vbDefaultButton2 + vbQuestion)
    If varval = vbYes Then
        g_CU.lngVerteiler = g_CU.lngVerteiler + 1
        g_db.WriteDisclaimer g_CU.lngVerteiler     ' Eins mehr
    End If
    FillStatus
End Sub

Private Sub Form_Load()
    Me.Caption = App.ExeName & " Security " & g_strProgramVersion
    lblInfo.Caption = g_db.GetString(1129) & vbCrLf & vbCrLf & _
        g_db.GetString(1130) & vbCrLf & _
        g_db.GetString(1131) & vbCrLf & _
        g_db.GetString(1132) & vbCrLf & vbCrLf & vbCrLf & _
        g_db.GetString(1133)
    FillStatus
End Sub

Private Sub FillStatus()
    lblStatus.Caption = g_db.GetString(1134) & ": "     ' Ihre aktuelle Sichtbarkeitsstufe:
    '1062:nur für Ihren Chef  1063:auch für Ihre Teamkollegen  1064:auch für Ihre Abteilungskollegen   1145:auch für Ihre Bereichskollegen 1141:auch für Ihre Ressortkollegen 1065:einsehbar.
    '1135:Niedrigste Sichtbarkeitebene 1136:Normale Sichtbarkeitsebene 1137:Hohe Sichtbarkeitsebene 1143:Sehr hohe Sichtbarkeitsebene 1138:Höchste Sichtbarkeitsebene
    '1139:Wenn Sie den Knopf rechts drücken, wechseln sie zu 1140:Sichtbarkeit auf
    '1126:Teamebene 1127:Abteilungsebene 1128:Bereichsebene 1142:Ressortebene
    Select Case g_CU.lngVerteiler
        Case SICHTBAR_CHEF
            lblStatus.Caption = lblStatus.Caption & """" & g_db.GetString(1062) & g_db.GetString(1065) & vbCrLf & g_db.GetString(1135) & vbCrLf & _
                g_db.GetString(1139) & vbCrLf & """" & g_db.GetString(1140) & " " & g_db.GetString(1126) & """."
            btnUpgrade.Enabled = True
        Case SICHTBAR_Org
            lblStatus.Caption = lblStatus.Caption & """" & g_db.GetString(1063) & g_db.GetString(1065) & vbCrLf & g_db.GetString(1136) & vbCrLf & _
                g_db.GetString(1139) & vbCrLf & """" & g_db.GetString(1140) & " " & g_db.GetString(1127) & """."
            btnUpgrade.Enabled = True
        Case SICHTBAR_ABT
            lblStatus.Caption = lblStatus.Caption & """" & g_db.GetString(1064) & g_db.GetString(1065) & vbCrLf & g_db.GetString(1137) & vbCrLf & _
                g_db.GetString(1139) & vbCrLf & """" & g_db.GetString(1140) & " " & g_db.GetString(1128) & """."
            btnUpgrade.Enabled = True
        Case SICHTBAR_BER
            lblStatus.Caption = lblStatus.Caption & """" & g_db.GetString(1145) & g_db.GetString(1065) & vbCrLf & g_db.GetString(1143) & vbCrLf & _
                g_db.GetString(1139) & vbCrLf & """" & g_db.GetString(1140) & " " & g_db.GetString(1142) & """."
            btnUpgrade.Enabled = True
        Case SICHTBAR_RES
            lblStatus.Caption = lblStatus.Caption & """" & g_db.GetString(1141) & g_db.GetString(1065) & vbCrLf & g_db.GetString(1138) & vbCrLf & _
                g_db.GetString(1144)
            btnUpgrade.Enabled = False
    End Select
End Sub
