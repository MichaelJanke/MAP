VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPlanen 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Planen"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6750
   Icon            =   "frmPlanen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox txtWeitereInfo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      HideSelection   =   0   'False
      Left            =   1680
      TabIndex        =   14
      Top             =   5520
      Width           =   1935
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "&Verwerfen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3780
      TabIndex        =   10
      Top             =   5400
      Width           =   1935
   End
   Begin VB.ListBox listAbwesenheitsArt 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1560
      Left            =   60
      TabIndex        =   2
      Top             =   3780
      Width           =   1995
   End
   Begin VB.CommandButton btnPlanen 
      Caption         =   "&Planen"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3780
      MaskColor       =   &H00000000&
      TabIndex        =   1
      Top             =   4620
      Width           =   1935
   End
   Begin MSComCtl2.MonthView calStart 
      CausesValidation=   0   'False
      Height          =   2370
      Left            =   60
      TabIndex        =   0
      Top             =   480
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483635
      BackColor       =   -2147483633
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MultiSelect     =   -1  'True
      StartOfWeek     =   16646146
      CurrentDate     =   37693
   End
   Begin MSComCtl2.MonthView calEnde 
      CausesValidation=   0   'False
      Height          =   2370
      Left            =   3480
      TabIndex        =   7
      Top             =   480
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483635
      BackColor       =   -2147483633
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      StartOfWeek     =   16646146
      CurrentDate     =   37693
   End
   Begin VB.Frame frameGVN 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1755
      Left            =   2130
      TabIndex        =   3
      Top             =   3600
      Width           =   1575
      Begin VB.OptionButton optVormittag 
         Caption         =   "&vormittag"
         Height          =   375
         Left            =   60
         TabIndex        =   6
         Top             =   600
         Width           =   1455
      End
      Begin VB.OptionButton optNachmittag 
         Caption         =   "&nachmittag"
         Height          =   375
         Left            =   60
         TabIndex        =   5
         Top             =   930
         Width           =   1455
      End
      Begin VB.OptionButton optGanztags 
         Caption         =   "&ganztags"
         Height          =   375
         Left            =   60
         TabIndex        =   4
         Top             =   300
         Width           =   1395
      End
   End
   Begin VB.Label lblWeitereInfo 
      Caption         =   "Weitere Info"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Endtermin"
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
      Left            =   3480
      TabIndex        =   12
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Starttermin"
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
      Index           =   0
      Left            =   60
      TabIndex        =   11
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label lblAnzahlTage 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Anzahl Tage:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   735
      Left            =   3780
      TabIndex        =   9
      Top             =   3780
      Width           =   1935
   End
   Begin VB.Label lblFeiertage 
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Feiertage: "
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   6180
      Width           =   5595
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmPlanen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private booGueltigesDatum As Boolean
Private varval As Variant
Private cAb As clsAbw

Private Sub btnCancel_Click()
    Me.Hide
End Sub

Private Sub btnPlanen_Click()
    If listAbwesenheitsArt = "" Then
        MsgBox g_db.GetString(1072), vbInformation
        Exit Sub
    End If
    
    ' Erstanlage
    If cAb.lngIdxAbwesenheit = 0 Then   ' neue Abwesenheit, dann auf Überschneidungen testen
        If g_clsAb.CheckDates(calStart, calEnde) = False Then
            MsgBox g_db.GetString(1073), vbCritical
            Exit Sub
        End If
    
        cAb.lngIdxUser = g_CU.lngIdxUser
        Set cAb.cUs = g_db.GetUserByID(cAb.lngIdxUser)
        
        cAb.strStatus = g_db.GetStatusText(cAb.lngIdxStatus)
    End If

    If cAb.lngIdxStatus = AbwStatus.PLANUNG Or cAb.lngIdxStatus = AbwStatus.UNDEFINED Then
        'Schreibe das, was in den Feldern steht, in die Datenbank
        cAb.lngIdxAbwesenheitsArt = g_db.GetAbwesenheitsartIndex(listAbwesenheitsArt)
        cAb.dtmStart = calStart
        cAb.dtmEnde = calEnde
        
        cAb.strAbwesenheitsart = listAbwesenheitsArt.Text
        
        If Me.optVormittag Then cAb.lngGVN = ABW_VORMITTAGS
        If Me.optNachmittag Then cAb.lngGVN = ABW_NACHMITTAGS
        If Me.optGanztags Then cAb.lngGVN = ABW_GANZTAGS
    End If
    
    cAb.strText = txtWeitereInfo.Text   ' Immer Änderung zulassen

    g_clsAb.Planen cAb
    g_db.FillAbwesenheitCollection
    Me.Hide
End Sub
Public Sub EnterForm(cAbO As clsAbw)
    Set cAb = cAbO
    If cAb.lngIdxAbwesenheit = 0 Then   ' neue Abwesenheit
        cAb.strText = ""
    Else                                ' Ändern
        ' Abwesenheitsart in der Listbox selektieren
        FillListeAbwesenheitsArt
        Dim strAbwesenheitsart As String
        strAbwesenheitsart = g_db.GetAbwesenheitsart(cAb.lngIdxAbwesenheitsArt).strAbwesenheitsart
        listAbwesenheitsArt = strAbwesenheitsart
        
        Dim i As Integer
'        For i = 0 To listAbwesenheitsArt.ListCount - 1
'            If listAbwesenheitsArt.List(i) = strAbwesenheitsart Then
'                listAbwesenheitsArt.Selected(i) = True
'                Exit For
'            End If
'        Next
        
        Select Case cAb.lngGVN
            Case ABW_VORMITTAGS
                Me.optVormittag = True
            Case ABW_NACHMITTAGS
                Me.optNachmittag = True
            Case ABW_GANZTAGS
                Me.optGanztags = True
        End Select
    End If
    
    
    calStart.value = cAb.dtmStart
    calEnde.value = cAb.dtmEnde
    calStart_DateClick cAb.dtmStart
    calEnde_DateClick cAb.dtmEnde
    
    txtWeitereInfo.Text = cAb.strText
    
    Dim booChange As Boolean
    booChange = cAb.lngIdxStatus = AbwStatus.UNDEFINED Or cAb.lngIdxStatus = AbwStatus.PLANUNG
    
    ' in Planung alles zulassen. danach nut noch den Info-Text ändern
    Me.calStart.Enabled = booChange
    Me.calEnde.Enabled = booChange
    Me.listAbwesenheitsArt.Enabled = booChange
    Me.txtWeitereInfo.Enabled = True    ' immer änderbar
    Me.frameGVN.Enabled = booChange
    Me.optGanztags.Enabled = booChange
    Me.optNachmittag.Enabled = booChange
    Me.optVormittag.Enabled = booChange
End Sub

'##########################################################################################################
Public Sub calStart_DateClick(ByVal DateClicked As Date)
    calStart = StartGueltig(calStart)
    If calStart > calEnde Then      ' Neuer Start liegt hinter altem Ende -> Ende vorverlegen
        calEnde.value = calStart.value  ' Neues Ende = neuer Start
    End If
    cAb.dtmStart = calStart
    CalcTage
End Sub
Public Sub calEnde_DateClick(ByVal DateClicked As Date)
    calEnde = EndeGueltig(calEnde)
    If calStart > calEnde Then      ' Start liegt nach Ende -> Start vorverlegen
        If calEnde > Date Then      ' Ist das Ende eigenlich in der Zukunft ?
            calStart.value = calEnde.value  ' Ja, also Starttermin = Endtermin
        Else                        ' Nein, Endtermin lieg in der Vergangenheit
                                    ' Suche den erten möglichen Starttermin ...
            calStart = Date         ' Korrigiere den Starttermin auf mindestens heute
            calStart = StartGueltig(calStart)            ' Verschiebe, falls notwendig
            calEnde = calStart      ' Korrigiere Endtermin ebenfalls auf den frühestmöglichen Starttermin
        End If
    End If
    cAb.dtmEnde = calEnde
    CalcTage
End Sub

Private Sub Form_Activate()
    btnPlanen.SetFocus
End Sub

Private Sub Form_Load()
    Me.Caption = App.ExeName & " Planing"
    ReadWindowPosition Me, booPositionOnly:=True
    Me.Height = 6975    ' OriginalHoehe
    Me.Caption = g_db.GetString(1006) & g_strProgramVersion
    Me.Label1(0).Caption = g_db.GetString(1100)  ' Starttermin
    Me.Label1(1).Caption = g_db.GetString(1101)  ' Endtermin
    Me.btnPlanen.Caption = g_db.GetString(1006)  ' Planen
    Me.btnCancel.Caption = g_db.GetString(1102)  ' Verwerfen
    Me.lblWeitereInfo.Caption = g_db.GetString(1103) ' WeitereInfo
    Me.optGanztags.Caption = g_db.GetString(1104)    ' ganztags
    Me.optVormittag.Caption = g_db.GetString(1045)   ' vormittags
    Me.optNachmittag.Caption = g_db.GetString(1044)  ' nachmittags
    
    Me.calStart.StartOfWeek = g_db.FirstWorkingDay
    Me.calEnde.StartOfWeek = g_db.FirstWorkingDay
    
    FillListeAbwesenheitsArt
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If g_booShutdown Then
        SaveWindowPosition Me
        Unload Me
    Else
        Me.Hide
    End If
End Sub

Private Sub FillListeAbwesenheitsArt()
    Dim c1 As Collection, cABA As clsAbwesenheitsart, s As String
    Set c1 = New Collection
    For Each cABA In g_db.colAbwesenheitsarten
        If cABA.booUserBeantragbar Or g_CU_Login.booIsSek Then c1.Add cABA.strAbwesenheitsart
    Next
    fillListBoxFromCollection listAbwesenheitsArt, c1
End Sub

Private Sub listAbwesenheitsArt_Click()
    cAb.lngIdxAbwesenheitsArt = g_db.GetAbwesenheitsartIndex(listAbwesenheitsArt)
    cAb.strAbwesenheitsart = listAbwesenheitsArt
    CalcTage
End Sub

Private Sub CalcTage()
Dim lngAnzFAKO As Long, lngAnzUrlaub As Long, lngAnzTage As Long, strAnzahlTage As String, strFeiertage As String
    strFeiertage = ":"
    CalculateTage cAb, False, strFeiertage
    lblAnzahlTage = cAb.lngAnzahl
    If strFeiertage <> ":" Then
        lblFeiertage.Visible = True
        lblFeiertage.Caption = g_db.GetString(1074) & ": " & strFeiertage
    Else
        lblFeiertage.Visible = False
    End If
    If cAb.lngAnzahl = 1 Then
        If cAb.lngIdxAbwesenheitsArt = basGlobals.ABW_URLAUB And _
         CBool(g_db.GetItem("booUrlaubGanztags", "True", "Vacation just whole day possible.(Default=True)")) Then
            Me.optGanztags = True    ' override bei eventuellem Umstellen FAKO->Urlaub Nur ganze Tage erlauben.
            frameGVN.Visible = False
        Else
            frameGVN.Visible = True
        End If
    Else
        frameGVN.Visible = False
    End If
End Sub
