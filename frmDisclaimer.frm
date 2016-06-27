VERSION 5.00
Begin VB.Form frmDisclaimer 
   Caption         =   "Disclaimer"
   ClientHeight    =   8955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9585
   ControlBox      =   0   'False
   Icon            =   "frmDisclaimer.frx":0000
   LinkTopic       =   "Form1"
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8955
   ScaleWidth      =   9585
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Frame Frame1 
      Height          =   2595
      Left            =   120
      TabIndex        =   6
      Top             =   6060
      Width           =   5415
      Begin VB.OptionButton optBer 
         Caption         =   "BereichssKollegen auch"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   9
         Top             =   2160
         Width           =   2595
      End
      Begin VB.OptionButton optAbt 
         Caption         =   "AbteilungsKollegen auch"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   8
         Top             =   1800
         Width           =   2595
      End
      Begin VB.OptionButton optOrg 
         Caption         =   "TeamKollegen auch"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   4
         Top             =   1440
         Width           =   2115
      End
      Begin VB.OptionButton optChef 
         Caption         =   "nur Chef"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   1500
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   $"frmDisclaimer.frx":0442
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1155
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   5115
      End
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Ich möchte nicht teilnehmen"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2475
      Left            =   7620
      Picture         =   "frmDisclaimer.frx":04F5
      Style           =   1  'Grafisch
      TabIndex        =   1
      Top             =   6180
      Width           =   1755
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "Ich erkläre mich einverstanden"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2475
      Left            =   5640
      Picture         =   "frmDisclaimer.frx":0937
      Style           =   1  'Grafisch
      TabIndex        =   2
      Top             =   6180
      Width           =   1875
   End
   Begin VB.Label lblDisclaimer 
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "lblDisclaimer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   9255
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Zentriert
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Willkommen zum Programm MAP - MitarbeiterAbwesenheitsPlanung"
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
      Height          =   1155
      Left            =   120
      TabIndex        =   0
      Top             =   300
      Width           =   9255
   End
End
Attribute VB_Name = "frmDisclaimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancel_Click()
    g_db.WriteDisclaimer    ' KEIN Disclaimer
    MsgBox g_db.GetString(1060)
    Unload Me
End Sub

Private Sub btnOK_Click()
Dim varval As Variant
    Dim strMessage As String
    If optChef Then strMessage = g_db.GetString(1062)
    If optOrg Then strMessage = g_db.GetString(1063)
    If optAbt Then strMessage = g_db.GetString(1064)
    If optBer Then strMessage = g_db.GetString(1145)    ' auch für Ihre Bereichskollegen
    ' 1061:Sie möchten an diesem Programm teilnehmen. Ihre Daten sind           1065:einsehbar    1066:Sind diese Angaben in Ihrem Sinne ?
    strMessage = g_db.GetString(1061) & " " & strMessage & " " & g_db.GetString(1065) & vbCrLf & vbCrLf & g_db.GetString(1066)
    varval = MsgBox(strMessage, vbYesNo + vbDefaultButton2)
    If varval = vbYes Then
        g_db.WriteDisclaimer IIf(optChef, 1, IIf(optOrg, 2, IIf(optAbt, 3, 4)))
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Me.Caption = "Disclaimer " & g_strProgramVersion
    optChef = True
    Me.lblHeader.Caption = g_db.GetString(1110)
    Me.Label1.Caption = g_db.GetString(1111)
    Me.btnOK.Caption = g_db.GetString(1112)
    Me.btnCancel.Caption = g_db.GetString(1113)
    Me.optChef.Caption = g_db.GetString(1114)
    Me.optOrg.Caption = g_db.GetString(1115)
    Me.optAbt.Caption = g_db.GetString(1116)
    Me.optBer.Caption = g_db.GetString(1117)
    lblDisclaimer.Caption = _
        g_db.GetString(1067) & vbCrLf & vbCrLf & _
        g_db.GetString(1068) & vbCrLf & vbCrLf & _
        g_db.GetString(1069) & vbCrLf & vbCrLf & _
        g_db.GetString(1070) & vbCrLf & vbCrLf & _
        g_db.GetString(1071)
End Sub

