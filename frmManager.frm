VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmManager 
   BorderStyle     =   5  'Änderbares Werkzeugfenster
   Caption         =   "Zustimmen kostet Manpower, Ablehnen bringt Ärger"
   ClientHeight    =   9915
   ClientLeft      =   60
   ClientTop       =   225
   ClientWidth     =   14490
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmManager.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   661
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   966
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CheckBox chkInDirect 
      DownPicture     =   "frmManager.frx":0442
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   12240
      Picture         =   "frmManager.frx":1C4C
      Style           =   1  'Grafisch
      TabIndex        =   11
      Top             =   120
      Value           =   1  'Aktiviert
      Width           =   975
   End
   Begin VB.CheckBox chkDeputy 
      DownPicture     =   "frmManager.frx":2FF6
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   11160
      Picture         =   "frmManager.frx":50C0
      Style           =   1  'Grafisch
      TabIndex        =   10
      Top             =   120
      Value           =   1  'Aktiviert
      Width           =   975
   End
   Begin VB.CommandButton btnSortuser 
      Caption         =   "Sort User"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6900
      Picture         =   "frmManager.frx":704A
      Style           =   1  'Grafisch
      TabIndex        =   9
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton btnSortDate 
      Caption         =   "Sort Datum"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8700
      Picture         =   "frmManager.frx":748C
      Style           =   1  'Grafisch
      TabIndex        =   8
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton btnInfoKW 
      Caption         =   "&Überschneidungen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   180
      Picture         =   "frmManager.frx":78CE
      Style           =   1  'Grafisch
      TabIndex        =   7
      ToolTipText     =   "Zeige alle Abwesenheiten in diesem Zeitraum"
      Top             =   2340
      Width           =   2355
   End
   Begin VB.CommandButton btnInfoUser 
      Caption         =   "&Benutzer alles"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   180
      Picture         =   "frmManager.frx":7D10
      Style           =   1  'Grafisch
      TabIndex        =   6
      ToolTipText     =   "Zeige alle Abwesenheiten dieses Benutzers"
      Top             =   1200
      Width           =   2355
   End
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
      Left            =   13560
      Picture         =   "frmManager.frx":8152
      Style           =   1  'Grafisch
      TabIndex        =   5
      ToolTipText     =   "Formular verlassen"
      Top             =   180
      Width           =   795
   End
   Begin VB.CommandButton btnNurOffene 
      Caption         =   "Beantragte"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4440
      Picture         =   "frmManager.frx":829C
      Style           =   1  'Grafisch
      TabIndex        =   4
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton btnAlleZeigen 
      Caption         =   "Alle"
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
      Height          =   855
      Left            =   2640
      Picture         =   "frmManager.frx":8B66
      Style           =   1  'Grafisch
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton btnAblehnen 
      Caption         =   "&Ablehnen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   180
      Picture         =   "frmManager.frx":9430
      Style           =   1  'Grafisch
      TabIndex        =   2
      Top             =   5520
      Width           =   2355
   End
   Begin VB.CommandButton btnZustimmen 
      Caption         =   "&Zustimmen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   180
      Picture         =   "frmManager.frx":973A
      Style           =   1  'Grafisch
      TabIndex        =   1
      Top             =   4320
      Width           =   2355
   End
   Begin MSFlexGridLib.MSFlexGrid FG 
      Height          =   8655
      Left            =   2580
      TabIndex        =   0
      Top             =   1140
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   15266
      _Version        =   393216
      BackColorBkg    =   -2147483633
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      ScrollBars      =   2
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS As ADODB.Recordset
Dim lngFGCol As Long, lngFGRow As Long
Dim lngFGOldRow As Long ' Die aktuelle Zeile vor FG_Click
Dim varval As Variant
Dim strOrder As String
Dim booZeigenBeantragte As Boolean
Dim strWhere As String
Dim lngCurrentRow As Long
Dim lngAnzUrlaub As Long, lngAnzFAKO As Long, lngAnzTage As Long

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub btnInfoKW_Click()
Dim lngRows As Long, lngCols As Long, strOrgOld As String

    If FG.Row = 0 Then Exit Sub
    
    Dim lngIdxUser As Long:             lngIdxUser = FG.TextMatrix(FG.Row, COL_IDXUSER)
    Dim lngIdxAbwesenheit As Long:      lngIdxAbwesenheit = FG.TextMatrix(FG.Row, COL_IDX)
    
    Dim colUserAbw As Collection:       Set colUserAbw = g_db.FillUserListeKW(lngIdxUser, lngIdxAbwesenheit)

    If colUserAbw.Count > 0 Then
        g_PopUpMode = "Tabelle"
        g_InfoText = g_db.GetString(1021) & " " & Format(FG.TextMatrix(FG.Row, COL_START), "dd.mm") & " " & g_db.GetString(1047) & " " & Format(FG.TextMatrix(FG.Row, COL_ENDE), "dd.mm.yy")
        lngRows = colUserAbw.Count + 1
        lngCols = 6
        ReDim g_InfoTabelle(lngCols)   ' 7 Spaltenbreiten
        ReDim g_DataTabelle(lngCols, lngRows)
        g_InfoTabelle(0) = 1200
        g_InfoTabelle(1) = 1500
        g_InfoTabelle(2) = 1500
        g_InfoTabelle(3) = 1050
        g_InfoTabelle(4) = 1050
        g_InfoTabelle(5) = 1000
        g_InfoTabelle(6) = 1200
        g_DataTabelle(0, 0) = g_db.GetString(1075)
        g_DataTabelle(1, 0) = g_db.GetString(1076)
        g_DataTabelle(2, 0) = g_db.GetString(1077)
        g_DataTabelle(3, 0) = g_db.GetString(1046)
        g_DataTabelle(4, 0) = g_db.GetString(1047)
        g_DataTabelle(5, 0) = g_db.GetString(1059)
        g_DataTabelle(6, 0) = g_db.GetString(1058)
        lngFGRow = 0
        strOrgOld = colUserAbw(1).cUs.cOrg.strOrg   ' Org des ersten Benutzers als Startwert
        Dim cAb As clsAbw
        For Each cAb In colUserAbw
            lngFGRow = lngFGRow + 1
            If strOrgOld <> cAb.cUs.cOrg.strOrg Then  ' Leerzeile einfügen
                strOrgOld = cAb.cUs.cOrg.strOrg
                lngFGRow = lngFGRow + 1
                lngRows = lngRows + 1
                g_InfoTabelle(0) = lngRows     ' Rows
                ReDim Preserve g_DataTabelle(lngCols, lngRows)
            End If
            g_DataTabelle(0, lngFGRow) = cAb.strAbwesenheitsart
            g_DataTabelle(6, lngFGRow) = cAb.strStatus
            g_DataTabelle(1, lngFGRow) = cAb.cUs.strNachname
            g_DataTabelle(2, lngFGRow) = cAb.cUs.strVorname
            g_DataTabelle(3, lngFGRow) = cAb.dtmStart
            g_DataTabelle(4, lngFGRow) = cAb.dtmEnde
            g_DataTabelle(5, lngFGRow) = cAb.lngAnzahl
        Next
        
        frmPopUp.Show vbModal
    End If
End Sub

Private Sub btnInfoUser_Click()
    If FG.Row = 0 Then Exit Sub
    
    Dim lngIdxUser As Long
    lngIdxUser = FG.TextMatrix(FG.Row, COL_IDXUSER)
    
    Dim colUserAbw As Collection
    Set colUserAbw = g_db.FillUserListe(lngIdxUser)
    If colUserAbw.Count > 0 Then
        g_PopUpMode = "Tabelle"
        g_InfoText = g_db.GetString(1022) & " " & colUserAbw(1).cUs.strNachname & ", " & colUserAbw(1).cUs.strVorname
        ReDim g_InfoTabelle(3)  ' rows, cols, 4 Spaltenbreiten
        ReDim g_DataTabelle(3, colUserAbw.Count)
        g_InfoTabelle(0) = 1500
        g_InfoTabelle(1) = 1200
        g_InfoTabelle(2) = 1200
        g_InfoTabelle(3) = 1000
        g_DataTabelle(0, 0) = g_db.GetString(1075)
        g_DataTabelle(1, 0) = g_db.GetString(1046)
        g_DataTabelle(2, 0) = g_db.GetString(1047)
        g_DataTabelle(3, 0) = g_db.GetString(1059)
        lngFGRow = 0
        Dim cAb As clsAbw
        For Each cAb In colUserAbw
            g_DataTabelle(0, lngFGRow) = cAb.strAbwesenheitsart
            g_DataTabelle(1, lngFGRow) = cAb.dtmStart
            g_DataTabelle(2, lngFGRow) = cAb.dtmEnde
            g_DataTabelle(3, lngFGRow) = cAb.lngAnzahl
            lngFGRow = lngFGRow + 1
        Next
        frmPopUp.Show vbModal
    End If
End Sub

' Anzeigen alle (sprich Org) oder
Private Sub btnAlleZeigen_Click()
    booZeigenBeantragte = False
    FillForm
End Sub
Private Sub btnNurOffene_Click()
    booZeigenBeantragte = True
    FillForm
End Sub

Private Sub btnSortDate_Click()
    strOrder = "dtmStart, strNachname"
    FillForm
End Sub

Private Sub btnSortuser_Click()
    strOrder = "strNachname, dtmStart"
    FillForm
End Sub

Private Sub btnAblehnen_Click()
    If lngFGOldRow > 0 And lngFGOldRow < 999 Then
        varval = MsgBox(g_db.GetString(1078), vbInformation + vbOKCancel + vbDefaultButton2) ' Wirklich ablehnen?
        If varval = vbOK Then
            Dim strBegründung As String
            strBegründung = InputBox(g_db.GetString(1079), g_strProgramVersion & g_db.GetString(1081), g_db.GetString(1080))  ' Begründung abfragen

            g_clsAb.Ablehnen FG.TextMatrix(FG.Row, COL_IDX), strBegründung
            FillListeAlle
        End If
    End If
End Sub

Private Sub btnZustimmen_Click()
    If lngFGOldRow > 0 And lngFGOldRow < 999 Then
        g_clsAb.Freigeben FG.TextMatrix(FG.Row, COL_IDX)
        FillListeAlle
    End If
End Sub

Private Sub chkDeputy_Click()
    If Me.chkDeputy.value = 0 Then  ' as boss
        Me.chkDeputy.Visible = g_CU_Login.booIsOrgChef2
        Me.chkInDirect.Visible = g_CU.lngUserLevel <= Abteilung
    Else                            ' as deputy
        ' wen vertrete ich?
        Dim ol As clsOrgLevel
        Dim topMostLevel As OrgLevel
        topMostLevel = Insel
        
        For Each ol In g_CU.colOrg
            If ol.Level = 2 Then    ' hier bin ich Deputy
                If ol.cOrg.lngOrgLevel < topMostLevel Then topMostLevel = ol.cOrg.lngOrgLevel
            End If
        Next
        Me.chkInDirect.Visible = (topMostLevel <= Abteilung)
    End If
    FillListeAlle
End Sub

Private Sub chkInDirect_Click()
    If Me.chkInDirect.value = 0 Then    ' direct only
    Else                                ' direct and indirect
    End If
    FillListeAlle
End Sub

Private Sub FG_Click()
    If FG.Row < 1 Then Exit Sub
    If FG.TextMatrix(FG.Row, COL_IDX) = "" Then
        Exit Sub
    End If
    lngFGRow = FG.Row
    FG.SelectionMode = flexSelectionByRow
    If lngFGOldRow <> lngFGRow Then
        FG.Redraw = False
        If lngFGOldRow < FG.Rows Then LineMark lngFGOldRow, False
        LineMark lngFGRow, True
        lngFGOldRow = lngFGRow     ' Merken
        FG.Redraw = True
    End If
    FG.ToolTipText = g_clsAb.StatusText(FG.TextMatrix(FG.Row, COL_IDX), True)
    
    ' Welche Rolle spiele ich hier - was darf ich mit den angezeigten Abwesenheitsn machen?
    Dim cAb As clsAbw
    Set cAb = basGlobals.g_db.GetAbwesenheit(FG.TextMatrix(FG.Row, COL_IDX))
    If cAb Is Nothing Then
        Me.btnZustimmen.Enabled = False
        Me.btnAblehnen.Enabled = False
    Else
        Me.btnZustimmen.Enabled = True
        Me.btnAblehnen.Enabled = True
        If cAb.lngIdxStatus = AbwStatus.GENEHMIGT Then
            Me.btnZustimmen.Enabled = False
        Else
            ' direkter Chef?
            If cAb.cUs.lngIdxChef = g_CU.lngIdxUser Or cAb.cUs.lngIdxChef2 = g_CU.lngIdxUser Then       ' direkter Chef 1/2
                Me.btnZustimmen.Enabled = True
            ElseIf g_db.GetUserByID(cAb.cUs.lngIdxChef).lngIdxChef = g_CU.lngIdxUser Then   ' nächster Chef
                If cAb.lngIdxStatus = AbwStatus.PLANUNG Then ' noch nicht beantragt
                    Me.btnZustimmen.Enabled = False
                Else
                    Me.btnZustimmen.Enabled = True
                End If
            End If
        End If
    End If
End Sub
Private Sub LineMark(lngLine As Long, booMark As Boolean)
        FG.Row = lngLine
        For lngFGCol = 1 To COL_MAX
            FG.col = lngFGCol:  FG.CellFontBold = booMark
        Next lngFGCol
End Sub

Private Sub FG_DblClick()
    If FG.Row > 1 Then Exit Sub
    lngFGRow = FG.Row
    lngFGCol = FG.col
    Stop
End Sub

Private Sub Form_Load()
    Me.Caption = g_db.GetString(1085) & " Manager"
    ReadWindowPosition Me, True
    Me.btnAlleZeigen.Caption = g_db.GetString(1091)
    Me.btnNurOffene.Caption = g_db.GetString(1092)
    Me.btnInfoUser.Caption = g_db.GetString(1093)
    Me.btnInfoUser.ToolTipText = g_db.GetString(1105)
    Me.btnInfoKW.Caption = g_db.GetString(1094)
    Me.btnInfoKW.ToolTipText = g_db.GetString(1106)
    Me.btnZustimmen.Caption = g_db.GetString(1095)
    Me.btnAblehnen.Caption = g_db.GetString(1096)
    Me.btnExit.ToolTipText = g_db.GetString(1107)
    strOrder = "dtmStart"
    booZeigenBeantragte = False
    Me.chkDeputy.value = False
    Me.chkInDirect = False
    Me.chkDeputy.Visible = g_CU_Login.booIsOrgChef2
    Me.chkInDirect.Visible = g_CU.lngUserLevel <= Abteilung
    FillForm    ' Fülle Grid und gehe in die letzte Zeile
End Sub

Private Sub Form_Resize()
    If Me.Visible = False Then Exit Sub
    If Me.Height = 465 Then Exit Sub        ' minimiert
    If Me.Height < 8000 Then
        Me.Height = 8000
    End If
    Me.FG.Height = Me.ScaleHeight - Me.FG.Top - 9
    If Me.Width < 12300 Then
        Me.Width = 12300
    End If
    Me.FG.Width = Me.ScaleWidth - Me.FG.Left - 9
    If FG.Cols >= COL_TEXT Then
        Me.FG.ColWidth(COL_TEXT) = Me.FG.Width * 15 - 8500
    End If
    Me.btnExit.Left = Me.FG.Left + Me.FG.Width - Me.btnExit.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWindowPosition Me
End Sub
Private Sub FillForm()
    FillListeAlle
End Sub

Private Sub FillListeAlle()
    FG.Cols = COL_MAX + 1
    FG.Rows = 1
    FG.ColWidth(COL_IDX) = 0
    FG.ColWidth(COL_IDXUSER) = 0
    FG.ColWidth(COL_USER) = 2400:       FG.TextMatrix(0, COL_USER) = g_db.GetString(1086)
    FG.ColWidth(COL_IDXTYP) = 0
    FG.ColWidth(COL_ORG) = 1000:        FG.TextMatrix(0, COL_ORG) = "Org"
    FG.ColWidth(COL_TYP) = 1100:        FG.TextMatrix(0, COL_TYP) = g_db.GetString(1057)
    FG.ColWidth(COL_START) = 1300:      FG.TextMatrix(0, COL_START) = g_db.GetString(1046)
    FG.ColWidth(COL_ENDE) = 1300:       FG.TextMatrix(0, COL_ENDE) = g_db.GetString(1047)
    FG.ColWidth(COL_GVN) = 200
    FG.ColWidth(COL_LAENGE) = 300:      FG.TextMatrix(0, COL_LAENGE) = "#"
    FG.ColWidth(COL_IDXSTAT) = 0
    FG.ColWidth(COL_STAT) = 1500:       FG.TextMatrix(0, COL_STAT) = g_db.GetString(1058)
    FG.ColWidth(COL_TEXT) = 3270:       FG.TextMatrix(0, COL_TEXT) = "Info"

    Dim colShow As Collection
    Set colShow = g_db.FillManagerListe(booZeigenBeantragte, chkDeputy.value = 1, chkInDirect.value = 1, strOrder)
    If colShow Is Nothing Then
        MsgBox (g_db.GetString(1087))   ' Fehler bei der Zusammenstellung
        FG.Rows = 1
        Exit Sub
    End If
    
    Dim lngFGRow As Long
    lngFGRow = 1
    FG.Rows = colShow.Count + 1
    
    Dim cAb As clsAbw
    For Each cAb In colShow
        FG.TextMatrix(lngFGRow, COL_IDX) = cAb.lngIdxAbwesenheit
        FG.TextMatrix(lngFGRow, COL_IDXUSER) = cAb.lngIdxUser
        FG.TextMatrix(lngFGRow, COL_USER) = cAb.cUs.strNachname & " " & cAb.cUs.strVorname
        FG.TextMatrix(lngFGRow, COL_IDXTYP) = cAb.lngIdxAbwesenheitsArt
        FG.TextMatrix(lngFGRow, COL_ORG) = g_db.GetUserByID(cAb.lngIdxUser).cOrg.strOrg
        FG.TextMatrix(lngFGRow, COL_TYP) = cAb.strAbwesenheitsart
        FG.TextMatrix(lngFGRow, COL_START) = cAb.dtmStart
        FG.TextMatrix(lngFGRow, COL_ENDE) = cAb.dtmEnde
        FG.TextMatrix(lngFGRow, COL_GVN) = cAb.strGVN
        FG.TextMatrix(lngFGRow, COL_IDXSTAT) = cAb.lngIdxStatus
        FG.TextMatrix(lngFGRow, COL_STAT) = cAb.strStatus
        FG.TextMatrix(lngFGRow, COL_LAENGE) = cAb.lngAnzahl
        FG.TextMatrix(lngFGRow, COL_TEXT) = cAb.strText
        
        lngFGRow = lngFGRow + 1
    Next

    lngFGOldRow = 999
    FG.Row = FG.Rows - 1    ' Setze Cursor in die letzte Zeile
    FG.col = 1
    FG_Click
End Sub

