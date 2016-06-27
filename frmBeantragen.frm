VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmBeantragen 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Ich glaub, Ich brauch mal wieder Urlaub ..."
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7875
   ControlBox      =   0   'False
   ForeColor       =   &H8000000D&
   Icon            =   "frmBeantragen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   7875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton btnExit 
      Cancel          =   -1  'True
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
      Height          =   795
      Left            =   7020
      Picture         =   "frmBeantragen.frx":0442
      Style           =   1  'Grafisch
      TabIndex        =   5
      ToolTipText     =   "Formular verlassen"
      Top             =   180
      Width           =   795
   End
   Begin MSFlexGridLib.MSFlexGrid FG 
      Height          =   6315
      Left            =   180
      TabIndex        =   4
      Top             =   2160
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   11139
      _Version        =   393216
      BackColorBkg    =   -2147483633
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
   Begin VB.CommandButton btnZurückziehen 
      Caption         =   "&Zurück- ziehen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   4800
      TabIndex        =   3
      ToolTipText     =   "Abwesenheit komplett streichen"
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton btnBeantragen 
      Caption         =   "&Beantragen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   3120
      TabIndex        =   2
      ToolTipText     =   "Beantragen der geplanten Abwesenheit"
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label lblBenutzer 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Aktueller Benutzer:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   795
      Left            =   180
      TabIndex        =   1
      Top             =   1140
      Width           =   2835
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Organisieren Abwesenheit"
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
      Height          =   735
      Left            =   195
      TabIndex        =   0
      Top             =   180
      Width           =   6165
   End
End
Attribute VB_Name = "frmBeantragen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lngFGMouseCol As Long, lngFGCol As Long, lngFGRow As Long
Private lngFGOldRow As Long, varval As Variant, strOrder As String, strText As String
Private dtmRun As Date
'##########################################################################################################
Private Sub btnExit_Click()
    Unload Me
End Sub
'##########################################################################################################
Private Sub btnBeantragen_Click()
    If FG.Row > 0 And FG.TextMatrix(FG.Row, COL_IDXSTAT) = AbwStatus.PLANUNG Then
        g_clsAb.Beantragen FG.TextMatrix(FG.Row, COL_IDX)
        FillListeAlle
    End If
End Sub
Private Sub btnZurückziehen_Click()
    If FG.Row > 0 Then
        g_clsAb.Zurueckziehen FG.TextMatrix(FG.Row, COL_IDX)
        FillListeAlle
    End If
End Sub

'##########################################################################################################
'##########################################################################################################
Private Sub Form_Load()
    Me.Caption = g_db.GetString(1008) & " " & g_db.GetString(1055)
    ReadWindowPosition Me, booPositionOnly:=True
    Me.lblHeader.Caption = g_db.GetString(1028)
    Me.lblBenutzer.Caption = g_db.GetString(1056) & ": " & vbCrLf & g_CU.Fullname
    Me.btnBeantragen.Caption = g_db.GetString(1097)
    Me.btnBeantragen.ToolTipText = g_db.GetString(1108)
    Me.btnZurückziehen.Caption = g_db.GetString(1098)
    Me.btnZurückziehen.ToolTipText = g_db.GetString(1109)
    Me.btnExit.ToolTipText = g_db.GetString(1107)
    strOrder = "dtmStart"
    FillListeAlle
End Sub
Private Sub Form_Unload(Cancel As Integer)
'    If g_booShutdown Then
        SaveWindowPosition Me
        Unload Me
'    Else
'        Me.Hide
'    End If
End Sub
'##########################################################################################################
'###############   Flexgrid   #############################################################################
'##########################################################################################################
Private Sub FillListeAlle()
Dim i As Long, lngAnzahl As Long, lngAnzahlUrlaub As Long, lngAnzahlFAKO As Long
    FG.SelectionMode = flexSelectionByRow
' Zeige alle eigenen Urlaube...
    FG.Cols = COL_MAX + 1
    FG.Rows = 1
    FG.ColWidth(COL_IDX) = 0
    FG.ColWidth(COL_IDXUSER) = 0
    FG.ColWidth(COL_USER) = 0
    FG.ColWidth(COL_IDXTYP) = 0
    FG.ColWidth(COL_ORG) = 0
    FG.ColWidth(COL_TYP) = 1100:    FG.TextMatrix(0, COL_TYP) = g_db.GetString(1057)
    FG.ColWidth(COL_START) = 1300:  FG.TextMatrix(0, COL_START) = g_db.GetString(1046)
    FG.ColWidth(COL_ENDE) = 1300:   FG.TextMatrix(0, COL_ENDE) = g_db.GetString(1047)
    FG.ColWidth(COL_GVN) = 400:     FG.TextMatrix(0, COL_GVN) = ""
    FG.ColWidth(COL_LAENGE) = 450: FG.TextMatrix(0, COL_LAENGE) = g_db.GetString(1059)
    FG.ColWidth(COL_IDXSTAT) = 0
    FG.ColWidth(COL_STAT) = 1500:   FG.TextMatrix(0, COL_STAT) = g_db.GetString(1058)

    FG.Width = 150
    For i = 0 To FG.Cols - 1
        FG.Width = FG.Width + FG.ColWidth(i)
    Next i
    
    ' Abwesenheiten sind nicht nach Startdatum sortiert, sollen hier aber so angezeigt werden.
    ' Die Kandidaten aus colAbwesenheiten in eine Collection colA einsortieren
    
    Dim colA As New Collection, cA As clsAbw, iRun As Integer, iFound As Integer
    Dim cAb As clsAbw, dtmS As Date
    
    For Each cAb In basGlobals.g_db.colAbwesenheiten
        If cAb.lngIdxUser = g_CU.lngIdxUser And cAb.dtmStart > Now() - 180 And cAb.lngIdxStatus <> AbwStatus.ABGELEHNT And cAb.lngIdxStatus <> AbwStatus.ZURUCKGEZOGEN Then
            If colA.Count = 0 Then  ' Wir fangen erst an, einfach rein
                colA.Add cAb
            Else
                iFound = -1
                For iRun = 1 To colA.Count ' im temporären Array suchen nach dem letzten Element mit kleinerem Datum
                    If colA(iRun).dtmStart < cAb.dtmStart Then iFound = iRun
                Next
                If iFound = -1 Then
                    colA.Add cAb, , 1  ' ganz vorn
                Else
                    colA.Add cAb, , , iFound ' nach diesem einordnen
                End If
            End If
        End If
    Next
    
    lngFGRow = 0: FG.Rows = lngFGRow + 1
    For Each cAb In colA
        lngFGRow = lngFGRow + 1: FG.Rows = lngFGRow + 1
        
        FG.TextMatrix(lngFGRow, COL_IDX) = cAb.lngIdxAbwesenheit
        FG.TextMatrix(lngFGRow, COL_IDXTYP) = cAb.lngIdxAbwesenheitsArt
        FG.TextMatrix(lngFGRow, COL_TYP) = cAb.strAbwesenheitsart
        FG.TextMatrix(lngFGRow, COL_START) = Format(cAb.dtmStart, "yyyy-mm-dd")
        FG.TextMatrix(lngFGRow, COL_ENDE) = Format(cAb.dtmEnde, "yyyy-mm-dd")
        FG.TextMatrix(lngFGRow, COL_GVN) = cAb.strGVN
        FG.TextMatrix(lngFGRow, COL_LAENGE) = cAb.lngAnzahl
        FG.TextMatrix(lngFGRow, COL_IDXSTAT) = cAb.lngIdxStatus
        FG.TextMatrix(lngFGRow, COL_STAT) = cAb.strStatus
    Next
    lngFGOldRow = 999
    FG.Row = FG.Rows - 1    ' Setze Cursor in die letzte Zeile
    FG.col = 1
    FG_Click
End Sub

'##########################################################################################################
Private Sub FG_Click()
Dim booDatumInZukunft As Boolean
    lngFGRow = FG.Row   ' Merken, weil LineMark FG.Row ändert
    lngFGCol = FG.col

    If lngFGOldRow <> lngFGRow Then
        FG.Redraw = False
        If lngFGOldRow < FG.Rows Then LineMark lngFGOldRow, False
        If FG.Rows > 1 Then
            LineMark lngFGRow, True
        Else
            btnBeantragen.Enabled = False
            btnZurückziehen.Enabled = False
        End If
        lngFGOldRow = lngFGRow     ' Merken
        FG.Redraw = True

        If lngFGRow > 0 Then    ' Stati der Knöpfe
            booDatumInZukunft = CDate(FG.TextMatrix(lngFGRow, COL_START)) >= Date
            btnBeantragen.Enabled = (FG.TextMatrix(lngFGRow, COL_IDXSTAT) = AbwStatus.PLANUNG) And booDatumInZukunft
            btnZurückziehen.Enabled = True ' 2.21 alle benutzer dürfen alles zurückziehen
            
            Dim cAb As clsAbw:  Set cAb = g_db.GetAbwesenheit(FG.TextMatrix(lngFGRow, COL_IDX))
            FG.ToolTipText = g_clsAb.StatusText(cAb.lngIdxAbwesenheit, g_CU.IsSekOf(cAb.cUs)) ' Beschreibung
            If cAb.strText <> "" Then FG.ToolTipText = FG.ToolTipText & "  <" & cAb.strText & ">"   ' evtl verlängern
        End If
    End If
    If lngFGRow = 0 And FG.Rows > 1 Then ' Sortieren
        strOrder = "dtmStart"   ' Default-Ordering
        If lngFGCol = COL_TYP Then strOrder = "strAbwesenheitsart, dtmStart"
        If lngFGCol = COL_START Then strOrder = "dtmStart"
        If lngFGCol = COL_ENDE Then strOrder = "dtmEnde"
        If lngFGCol = COL_STAT Then strOrder = "strStatus, dtmEnde"
        FillListeAlle       ' Damit wechselt die MouseRow
    End If
End Sub
Private Sub LineMark(lngLine As Long, booMark As Boolean)
        FG.Row = lngLine
        For lngFGCol = 1 To COL_MAX
            FG.col = lngFGCol:  FG.CellFontBold = booMark
        Next lngFGCol
End Sub
'##########################################################################################################
