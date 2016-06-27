VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public m_conn As New ADODB.Connection
Private m_MaxUserID As Long
Private aOrgHash() As Long, aOrgVisible() As Long
Private m_strText As String

Public colFeiertage As Collection
Private aFeiertage() As clsFeiertag
Private dtmFeiertagStart As Date
Public colSchulferien As Collection
Public colAbwesenheiten As Collection
Public colOrg As Collection     ' alle Organisationseinheiten, später die sichtbaren Orgs
Public colUser As Collection    ' die sichtbaren Benutzer
Public colItem As Collection
Public colUserItems As Collection   ' Die Namen der User-veränderbaren Items
Public colAbwesenheitsarten As Collection
Public colStatus As Collection
Public colTexte As Collection
Public colOrgItem As Collection
Public colXOrgItem As Collection
Public colFunctions As Collection

Public WeekendFirstDay As Integer   ' 1=Sunday ... 7=Saturday
Public WeekendSecondDay As Integer   ' 1=Sunday ... 7=Saturday
Public WeekendNumDays As Integer
Public FirstWorkingDay As Integer

Public Enum OrgItem ' Enumeration für die möglichen Org-Tags
    Sekretariat = 1
    ICSMailAnChef = 2
    ICSMailAnSek = 3
    ICSMailAnChefChef = 4
    E4siehtKranktheit = 5
    OrgSlider = 6
    ShowSortColumn = 7
    SuperL = 8
    AnzTageLevel1 = 9
    ShowNoColors = 10
    SuperMA = 12
    ChangePast = 13
End Enum

Private m_dtmItemUpdate As Date

Public Enum AbwStatus       ' Status einer Abwesenheit
    UNDEFINED = -1
    PLANUNG = 0
    BEANTRAGT_1 = 1
    BEANTRAGT_2 = 2
    GENEHMIGT = 3
    ZURUCKGEZOGEN = 4
    ABGELEHNT = 5
End Enum

Public Property Get MaxUserID() As Long
    MaxUserID = m_MaxUserID
End Property
Public Property Get GetDbPath() As String
    Dim DataSourceName As String
    DataSourceName = m_conn.Properties("Data Source")
    GetDbPath = Left(DataSourceName, InStrRev(DataSourceName, "\"))
End Property
Public Property Get RSSchemaTables() As ADODB.Recordset
    If OpenDB() Then
        Set RSSchemaTables = m_conn.OpenSchema(adSchemaTables, Array(Empty, Empty, Empty, "TABLE"))   ' Nur Tables
    Else
        Set RSSchemaTables = Nothing
    End If
End Property
Public Property Get RSSchemaFields(strTable As String) As ADODB.Recordset
    If OpenDB() Then
        Set RSSchemaFields = m_conn.OpenSchema(adSchemaColumns, Array(Empty, Empty, strTable, Empty))
    Else
        Set RSSchemaFields = Nothing
    End If
End Property

Private Sub InitWeekend()
    Dim txtWeekendFirstDay As String
    txtWeekendFirstDay = GetItem("WeekendFirstDay", "sat", "First Day of Weekend - please enter first three letters english : mon tue wed thu fri sat sun default=sat")
    WeekendNumDays = GetItem("WeekendNumDays", 2, "Number of days für weekend - default and maximum = 2")
    Select Case txtWeekendFirstDay
        Case "sat"
            WeekendFirstDay = mvwSaturday
            WeekendSecondDay = mvwSunday
        Case "sun"
            WeekendFirstDay = mvwSunday
            WeekendSecondDay = mvwMonday
        Case "mon"
            WeekendFirstDay = mvwMonday
            WeekendSecondDay = mvwTuesday
        Case "tue"
            WeekendFirstDay = mvwTuesday
            WeekendSecondDay = mvwWednesday
        Case "wed"
            WeekendFirstDay = mvwWednesday
            WeekendSecondDay = mvwThursday
        Case "thu"
            WeekendFirstDay = mvwThursday
            WeekendSecondDay = mvwFriday
        Case "fri"
            WeekendFirstDay = mvwFriday
            WeekendSecondDay = mvwSaturday
        Case Default
            WeekendFirstDay = mvwSaturday
            WeekendSecondDay = mvwSunday
    End Select
    If WeekendNumDays < 1 Then WeekendFirstDay = 0
    If WeekendNumDays < 2 Then WeekendSecondDay = 0
    
    FirstWorkingDay = WeekendFirstDay + WeekendNumDays
    FirstWorkingDay = (FirstWorkingDay - 1) Mod 7 + 1     ' result should be 1=Sun to 7=Sat
End Sub

Private Sub Class_Terminate()
    CloseDB
End Sub

Public Function OpenDB() As Boolean
    Dim strDBName As String, strSystemDB As String
    On Error Resume Next
    If m_conn.State = adStateOpen Then
        OpenDB = True
    Else
        Dim ExeName As String
        strDBName = App.Path & "\" & App.ExeName & ".mdb"
        If Dir(strDBName) = "" Then
            MsgBox strDBName & " not found.":        Return
        End If
    
        strSystemDB = App.Path & "\sys2k.mdw"
        If Dir(strSystemDB) = "" Then
            MsgBox "sys2k.mdw " & " not found.":     Return
        End If
    
        ' m_conn ist bereits initialisiert
        m_conn.Provider = "Microsoft.Jet.OLEDB.4.0"
        m_conn.Properties("Jet OLEDB:System database") = strSystemDB
        m_conn.Open "Data Source=" & strDBName, DBUSER, DBPASSWORD
        If Err <> 0 Then
            MsgBox "clsData - Conn.Open:Error: " & Err.Description
        End If
        OpenDB = (m_conn.State = adStateOpen)
        
    End If
End Function
Public Function InitDB() As Boolean
    If RSSchemaFields("tblText").BOF Then   ' das geht ja mal gar nicht ...
        MsgBox "database is incomplete - tblText missing. closing application"
        InitDB = False
        Return
    End If
    FillTextCollection      ' Alle Texte in allen Sprachen
    InitDB = Checkfields ' Einmalig die DB checken
End Function
Private Function CloseDB() As Boolean   ' nach öffentlich aufrufbaren DB-Aktionen aufrufen (FillAllCollections und FillAbwesenheitCollection)
    On Error Resume Next
    If m_conn.State = adStateOpen Then m_conn.Close
    CloseDB = True
End Function
Public Function FillAllCollections() As Boolean
    FillAllCollections = False
    On Error GoTo err_FillAllCollections
    
    ' Keine Voraussetzungen:
    FillStatusCollection    ' StatusTexte - entsprechend der gewählten Sprache
    FillOrgCollection       ' alle Teams etc
    FillItemCollection      ' Programmeinstellungen
    FillFeiertagCollection  ' Feiertage
    FillSchulferienCollection   ' Schulferien
    FillFunctions           ' Funktionen einlesen
    FillAbwesenheitsartCollection   ' für den Lookup nach Strings
    
    ' Voraussetzung: Org
    UpdateOrgSetTopOrg      ' Erste Erweiterungen in Org
    FillUserCollection      ' alle User erstmalig
    
    ' Voraussetzung: Org + Item
    FillOrgItemCollection   ' Programmeinstellungen Org-Zuordnungen
    UpdateItemCollectionForCU    ' Setze die Items gemäß OrgItem + UserItem
    
    ' Voraussetzung: Org + User
    UpdateOrgCollection     ' Collections in colOrg
    
    ' Voraussetzung: Org + User + Functions
    FillFunctionsUser       ' FunktionsBenutzer einlesen   - User müssen schon da sein
    FillFunctionsOrg        ' Funktionen durch Org-Chef konstruieren
    
    FillAllCollections = True
    Exit Function
err_FillAllCollections:
    m_strText = "DB ERROR FillAllCollections " & Err.Number & ":" & Err.Description
    MsgBox m_strText
    FillAllCollections = False
    Exit Function
End Function

Private Function OpenRecordset(strSQL As String, Optional curOpenMode As ADODB.CursorTypeEnum = adOpenKeyset) As ADODB.Recordset
    On Error Resume Next
    Set OpenRecordset = Nothing
    If Not OpenDB() Then
        MsgBox ("Unable to access database. Restart program." & vbCrLf & "Query:" & strSQL)
        Exit Function
    End If
    
    Set OpenRecordset = New ADODB.Recordset
    OpenRecordset.Open strSQL, m_conn, curOpenMode, adLockOptimistic
    If Err <> 0 Then
        m_strText = "DB ERROR Openrecordset(" & strSQL & ") " & Err.Number & ":" & Err.Description
        MsgBox m_strText
        Logging "DB ERROR", , m_strText
    End If
End Function

Private Sub CloseRecordset(RS As ADODB.Recordset)
    On Error Resume Next
    If Not RS Is Nothing Then
        If RS.State = adStateOpen Then RS.Close
        Set RS = Nothing
    End If
End Sub

Private Function Execute(strSQL As String) As Boolean
    If Not OpenDB() Then
        MsgBox ("Unable to access database. Restart program." & vbCrLf & "Query:" & strSQL)
        Exit Function
    End If
    
    On Error Resume Next
    Execute = True
    m_conn.Execute strSQL
    If Err.Number <> 0 Then
        m_strText = "DB ERROR Execute: " & strSQL & " Error:" & Err.Number & ":" & Err.Description
        Logging "DB ERROR ", , m_strText
        If Not InStr(Err.Description, "tblItems") Then MsgBox m_strText
        Execute = False
    End If
End Function

Private Function Lookup(strField As String, strTable As String, strWhere As String) As Variant
    Lookup = Null
    Dim RS As ADODB.Recordset
    Set RS = OpenRecordset("select " & strField & " from " & strTable & " where " & strWhere, adOpenKeyset)
    If RS Is Nothing Then Exit Function
    If RS.RecordCount = 1 Then Lookup = RS.Fields(0)
End Function

Private Function LookupString(strField As String, strTable As String, strWhere As String) As String
    Dim varval As Variant
    varval = Lookup(strField, strTable, strWhere)
    LookupString = IIf(IsNull(varval), STR_UNKNOWN, varval)
End Function
Private Function LookupLong(strField As String, strTable As String, strWhere As String) As Long
    Dim varval As Variant
    varval = Lookup(strField, strTable, strWhere)
    LookupLong = IIf(IsNull(varval), LNG_UNKOWN, varval)
End Function

'Public Function Update(strField As String, strTable As String, strWhere As String, varValue As Variant) As Boolean
'    On Error GoTo err_Update
'Dim strKey As String, strVal As String, iEQ As Long, dtDate As Date
'    Dim RS As ADODB.Recordset
'    Set RS = OpenRecordset(strTable, adOpenKeyset)
'    If RS Is Nothing Then Exit function
'    RS.Find strWhere
'    If RS.EOF Then ' exitiert nicht -> Neu anlegen
'        RS.AddNew
'        iEQ = InStr(strWhere, "=")
'        strKey = Left(strWhere, iEQ - 1)
'        strVal = Mid(strWhere, iEQ + 1)
'        RS.Fields(strKey).value = strVal
'    End If
'    RS.Fields(strField).value = varValue
'    RS.Update
'    Update = True
'
'exit_Update:
'    CloseRecordset RS
'    Exit Function
'err_Update:
'    m_strText = "DB ERROR Update err=" & Err.Number & ":" & Err.Description
'    Logging "DB ERROR", , m_strText
'    MsgBox m_strText
'    Update = False
'    Resume exit_Update
'End Function

Public Sub WriteDisclaimer(Optional lngVerteiler As Long = -1)
    If g_CU Is Nothing Then Exit Sub
    Logging "WriteDisclaimer", , "User:" & g_CU.strAccountname & " Level:" & lngVerteiler
    Dim RS As ADODB.Recordset
    Set RS = OpenRecordset("SELECT * FROM tblUser WHERE lngIdxUser=" & g_CU.lngIdxUser, adOpenKeyset)
    If RS Is Nothing Then Exit Sub
    If RS.RecordCount = 1 Then
        If lngVerteiler <> -1 Then
            RS!dtmDisclaimer = Now()
            If lngVerteiler < SICHTBAR_CHEF Then lngVerteiler = SICHTBAR_CHEF ' Nur Vorgesetzter
            If lngVerteiler > basGlobals.SICHTBAR_RES Then lngVerteiler = SICHTBAR_RES   ' Alle BereichsMitglieder
            RS!lngVerteiler = lngVerteiler
            g_CU.lngVerteiler = RS!lngVerteiler
            g_CU.dtmDisclaimer = RS!dtmDisclaimer
        Else
            RS!dtmDisclaimer = Null
            RS!lngVerteiler = Null
            g_CU.lngVerteiler = 0
            g_CU.dtmDisclaimer = "00:00:00"
        End If
        RS.Update
    End If
    CloseRecordset RS
End Sub

Private Sub FillFeiertagCollection()
Dim cF As clsFeiertag, lFeiertag As Long
    On Error GoTo err_FillFeiertagCollection
    
    Set colFeiertage = New Collection
    
    Dim RS As ADODB.Recordset
    Set RS = OpenRecordset("SELECT * FROM tovFeiertag WHERE dtmFeiertag is not NULL ORDER BY dtmFeiertag", adOpenKeyset)
    If RS Is Nothing Then Exit Sub
    If RS.RecordCount = 0 Then  ' überhaupt keine Feiertage eingetragen
        ReDim aFeiertage(1) ' damit überhaupt ein Array existiert - sonst Fehler bei GetFeiertagByDate
    Else
        dtmFeiertagStart = RS!dtmFeiertag   ' auf diesen bezieht sich die Feiertagscollection - erster definierter Feiertag
            
        While Not RS.EOF
            Set cF = New clsFeiertag
            cF.lngIdxFeiertag = RS!lngIdxFeiertag
            cF.dtmFeiertag = RS!dtmFeiertag
            cF.strFeiertag = IIf(Not IsNull(RS!strFeiertag), RS!strFeiertag, "")
            cF.lngIdxAbwesenheitsArt = IIf(Not IsNull(RS!lngIdxAbwesenheitsArt), RS!lngIdxAbwesenheitsArt, 0)
            On Error Resume Next
    '            colFeiertage.Add cF, "cF_" & Format(rs!dtmFeiertag, "yyyy-mm-dd")
            colFeiertage.Add cF, "cF_" & cF.lngIdxFeiertag
            On Error GoTo err_FillFeiertagCollection
            
            lFeiertag = CLng(cF.dtmFeiertag - dtmFeiertagStart)
            ReDim Preserve aFeiertage(lFeiertag)
            Set aFeiertage(lFeiertag) = cF  ' Verweis im Array speichern für schnellen Zugriff
            RS.MoveNext
        Wend
    End If
exit_FillFeiertagCollection:
    CloseRecordset RS
    Exit Sub
err_FillFeiertagCollection:
    m_strText = "DB ERROR FillFeiertagCollection err=" & Err.Number & ":" & Err.Description
    MsgBox m_strText
    Logging "DB ERROR", , m_strText
    Resume exit_FillFeiertagCollection
End Sub
Public Function GetFeiertagByDate(dtm As Date) As clsFeiertag
    Dim lFeiertag As Long
    lFeiertag = CLng(dtm - dtmFeiertagStart)
    If lFeiertag > UBound(aFeiertage) Or lFeiertag < 0 Then
        Set GetFeiertagByDate = Nothing
        Exit Function
    End If
    If Not aFeiertage(lFeiertag) Is Nothing Then
        Set GetFeiertagByDate = aFeiertage(lFeiertag)
    Else
        Set GetFeiertagByDate = Nothing
    End If
End Function
Public Function GetFeiertag(lngIdxFeiertag As Long) As clsFeiertag
    Set GetFeiertag = Nothing
    On Error Resume Next
    Set GetFeiertag = colFeiertage("cF_" & lngIdxFeiertag)
End Function


Public Sub FillAbwesenheitCollection(Optional strWhere As String = "", Optional strSort As String = "")
Dim cAb As clsAbw, lngAnzahl As Long, lngAnzahlFAKO As Long, lngAnzahlUrlaub As Long, varRet As Variant

    On Error GoTo err_FillAbwesenheitCollection
    Set colAbwesenheiten = New Collection           ' Alles vorher löschen
    
    strSQL = "SELECT * FROM tblAbwesenheit "
    If strWhere = "" Then
        strSQL = strSQL & "WHERE dtmStart > now() - " & g_db.GetItem("DatenTage", "180", "Maximum age of data")
    Else
        strSQL = strSQL & strWhere
    End If
    strSQL = strSQL & " AND lngIdxStatus <> " & AbwStatus.ABGELEHNT & " AND lngIdxStatus <> " & AbwStatus.ZURUCKGEZOGEN & " "   ' Keine gelöschten/zurückgezogenen
    If strSort = "" Then
        strSQL = strSQL & "ORDER BY dtmStart"
    Else
        strSQL = strSQL & strSort
    End If

    Dim RS As ADODB.Recordset
    Set RS = OpenRecordset(strSQL, adOpenKeyset)
    If RS Is Nothing Then Exit Sub
    If RS.RecordCount > 0 Then
        While Not RS.EOF
            Set cAb = New clsAbw

            cAb.lngIdxAbwesenheit = RS!lngIdxAbwesenheit
            cAb.lngIdxAbwesenheitsArt = RS!lngIdxAbwesenheitsArt
            cAb.lngIdxUser = RS!lngIdxUser
            
            cAb.dtmStart = RS!dtmStart
            cAb.dtmEnde = RS!dtmEnde
            cAb.lngGVN = RS!lngGVN
            cAb.lngIdxStatus = RS!lngIdxStatus
            cAb.dtmErstellung = IIf(IsNull(RS!dtmErstellung), "00:00:00", RS!dtmErstellung)
            cAb.dtmAntrag = IIf(IsNull(RS!dtmAntrag), "00:00:00", RS!dtmAntrag)
            cAb.dtmGenehmigt1 = IIf(IsNull(RS!dtmGenehmigt1), "00:00:00", RS!dtmGenehmigt1)
            cAb.dtmGenehmigt2 = IIf(IsNull(RS!dtmGenehmigt2), "00:00:00", RS!dtmGenehmigt2)
            cAb.strText = IIf(IsNull(RS!strText), "", RS!strText)
            
            ' Berechnungen
            cAb.strAbwesenheitsart = g_db.GetAbwesenheitsart(RS!lngIdxAbwesenheitsArt).strAbwesenheitsart
            cAb.strStatus = g_db.GetStatusText(RS!lngIdxStatus)
            Set cAb.cUs = g_db.GetUserByID(cAb.lngIdxUser)
            If RS!lngGVN = ABW_GANZTAGS Then
                cAb.strGVN = "" ' Ganztags
            ElseIf RS!lngGVN = ABW_VORMITTAGS Then
                cAb.strGVN = GetString(1042)
            ElseIf RS!lngGVN = ABW_NACHMITTAGS Then
                cAb.strGVN = GetString(1043)
            End If
            
            ' Fülle Anzahl Abwesenheitstage
            CalculateTage cAb, booDetails:=True
            colAbwesenheiten.Add cAb, "cAb_" & cAb.lngIdxAbwesenheit
            RS.MoveNext
        Wend
    End If
exit_FillAbwesenheitCollection:
    CloseRecordset RS
    CloseDB     ' periodisch Datenbank schliessen
    Exit Sub
err_FillAbwesenheitCollection:
    m_strText = "DB ERROR FillAbwesenheitCollection err=" & Err.Number & ":" & Err.Description
    Logging "DB ERROR", , m_strText
    MsgBox m_strText
    Resume exit_FillAbwesenheitCollection
End Sub
Public Function GetAbwesenheit(lngIdxAbwesenheit) As clsAbw
    Set GetAbwesenheit = Nothing
    On Error Resume Next
    Set GetAbwesenheit = colAbwesenheiten("cAb_" & lngIdxAbwesenheit)
End Function
Public Sub UpdateAbwesenheit(cAb As clsAbw, lngIdxStatusNeu As AbwStatus)
    On Error GoTo err_UpdateAbwesenheit

    Dim strVorgang As String
    cAb.lngIdxStatusOld = cAb.lngIdxStatus  ' Merken
    
    Dim booChange As Boolean
    booChange = False
    ' StateMachine
    If lngIdxStatusNeu = cAb.lngIdxStatus Then                      ' Ändern
        Logging "Aendern", cAb
        booChange = False
    Else
        Select Case cAb.lngIdxStatus
            Case AbwStatus.UNDEFINED
                If lngIdxStatusNeu = PLANUNG Then booChange = True
            Case AbwStatus.PLANUNG
                If lngIdxStatusNeu = BEANTRAGT_1 Then booChange = True      ' Planen -> 1
                If lngIdxStatusNeu = ZURUCKGEZOGEN Then booChange = True    ' Planen -> Zurückgezogen
                If lngIdxStatusNeu = GENEHMIGT Then booChange = True        ' Planen -> Genehmigt
            Case AbwStatus.BEANTRAGT_1
                If lngIdxStatusNeu = BEANTRAGT_2 Then booChange = True      ' Beantragt_1 -> 2
                If lngIdxStatusNeu = GENEHMIGT Then booChange = True        ' Beantragt_1 -> Genehmigt
                If lngIdxStatusNeu = ABGELEHNT Then booChange = True        ' Beantragt_1 -> Abgelehnt
                If lngIdxStatusNeu = ZURUCKGEZOGEN Then booChange = True    ' Beantragt_1 -> Zurückgezogen
            Case AbwStatus.BEANTRAGT_2
                If lngIdxStatusNeu = GENEHMIGT Then booChange = True        ' Beantragt_2 -> Genehmigt
                If lngIdxStatusNeu = ABGELEHNT Then booChange = True        ' Beantragt_2 -> Abgelehnt
                If lngIdxStatusNeu = ZURUCKGEZOGEN Then booChange = True    ' Beantragt_2 -> Zurückgezogen
            Case AbwStatus.GENEHMIGT
                If lngIdxStatusNeu = ZURUCKGEZOGEN Then booChange = True    ' Genehmigt -> Zurückgezogen
                If lngIdxStatusNeu = ABGELEHNT Then booChange = True        ' Genehmigt -> Abgelehnt
        End Select
    End If

    If booChange Then       ' Statuswechsel - Datum merken
        cAb.lngIdxStatus = lngIdxStatusNeu  ' Setzen aktuell
        cAb.strStatus = Me.GetStatusText(lngIdxStatusNeu)
        
        Select Case lngIdxStatusNeu
            Case AbwStatus.PLANUNG
                Logging "Planen", cAb:          cAb.dtmErstellung = Now
            Case AbwStatus.BEANTRAGT_1
                Logging "Beantragen_1", cAb:    cAb.dtmAntrag = Now
            Case AbwStatus.BEANTRAGT_2
                Logging "Beantragen_2", cAb:    cAb.dtmGenehmigt1 = Now
            Case AbwStatus.GENEHMIGT
                Logging "Genehmigen", cAb:      cAb.dtmGenehmigt2 = Now
            Case AbwStatus.ABGELEHNT
                Logging "Ablehnen", cAb:        cAb.dtmGenehmigt2 = Now
            Case AbwStatus.ZURUCKGEZOGEN
                Logging "Zurueckziehen", cAb:   cAb.dtmGenehmigt2 = Now
        End Select
    End If
    
    Dim RS As ADODB.Recordset
    Set RS = OpenRecordset("tblAbwesenheit", dbOpenDynaset)
    If RS Is Nothing Then Exit Sub
    RS.Find "lngIdxAbwesenheit=" & cAb.lngIdxAbwesenheit
    If RS.EOF Then
        RS.AddNew  ' noch nichts da
        cAb.lngIdxAbwesenheit = RS!lngIdxAbwesenheit  ' Wenn neu, dann ist neuer Record Current -> Hole neue ID
    End If

    RS!lngIdxUser = cAb.lngIdxUser
    RS!lngIdxStatus = cAb.lngIdxStatus

    If cAb.dtmAntrag > 0 Then RS!dtmAntrag = cAb.dtmAntrag
    If cAb.dtmGenehmigt1 > 0 Then RS!dtmGenehmigt1 = cAb.dtmGenehmigt1
    If cAb.dtmGenehmigt2 > 0 Then RS!dtmGenehmigt2 = cAb.dtmGenehmigt2
    RS!lngIdxAbwesenheitsArt = cAb.lngIdxAbwesenheitsArt
    RS!dtmStart = cAb.dtmStart
    RS!dtmEnde = cAb.dtmEnde
    RS!lngGVN = cAb.lngGVN
    RS!dtmErstellung = cAb.dtmErstellung
    If cAb.strText <> "" Then RS!strText = cAb.strText  ' bei Ändern zugelassen

    RS.Update
    Logging cAb.strStatus, cAb, "Updated"

exit_UpdateAbwesenheit:
    CloseRecordset RS
    Exit Sub
err_UpdateAbwesenheit:
    m_strText = "DB ERROR UpdateAbwesenheit err=" & Err.Number & ":" & Err.Description
    Logging "DB ERROR", , m_strText
    MsgBox m_strText
    Resume exit_UpdateAbwesenheit
End Sub
Public Function IstAbwesend(IdxUser As Long) As Boolean
    IstAbwesend = False
    Dim cAb As clsAbw
    For Each cAb In colAbwesenheiten
        If cAb.lngIdxUser = IdxUser Then
            If cAb.dtmStart < Now And cAb.dtmEnde > Now - 1 Then
                IstAbwesend = True
                Exit Function
            End If
        End If
    Next
End Function
Private Sub FillAbwesenheitsartCollection()
Dim cABA As clsAbwesenheitsart

    On Error GoTo err_FillAbwesenheitsartCollection
    Set colAbwesenheitsarten = New Collection           ' Alles vorher löschen
    
    If basGlobals.g_iLCID = basGlobals.LocaleGerman Then
        strSQL = "SELECT lngIdxAbwesenheitsart, strAbwesenheitsArt, booUserBeantragbar FROM tovAbwesenheitsart"
    Else
        strSQL = "SELECT lngIdxAbwesenheitsart, strAbwesenheitsArtE, booUserBeantragbar FROM tovAbwesenheitsart"
    End If

    Dim RS As ADODB.Recordset
    Set RS = OpenRecordset(strSQL, adOpenKeyset)
    If RS Is Nothing Then Exit Sub
    If RS.RecordCount > 0 Then
        While Not RS.EOF
            Set cABA = New clsAbwesenheitsart

            cABA.lngIdxAbwesenheitsArt = RS(0)
            cABA.strAbwesenheitsart = RS(1)
            cABA.booUserBeantragbar = RS(2)
            colAbwesenheitsarten.Add cABA, "cABA" & cABA.lngIdxAbwesenheitsArt
            RS.MoveNext
        Wend
    End If
exit_FillAbwesenheitsartCollection:
    CloseRecordset RS
    Exit Sub
err_FillAbwesenheitsartCollection:
    m_strText = "DB ERROR FillAbwesenheitsartCollection err=" & Err.Number & ":" & Err.Description
    Logging "DB ERROR", , m_strText
    MsgBox m_strText
    Resume exit_FillAbwesenheitsartCollection
End Sub
Public Function GetAbwesenheitsart(lngIdxAbwesenheitsArt As Long) As clsAbwesenheitsart
    On Error Resume Next
    Set GetAbwesenheitsart = Nothing
    Set GetAbwesenheitsart = Me.colAbwesenheitsarten("cABA" & lngIdxAbwesenheitsArt)
End Function
Public Function GetAbwesenheitsartIndex(strAbwesenheitsart As String) As Long
    Dim cABA As clsAbwesenheitsart
    For Each cABA In colAbwesenheitsarten
        If cABA.strAbwesenheitsart = strAbwesenheitsart Then
            GetAbwesenheitsartIndex = cABA.lngIdxAbwesenheitsArt
            Exit For
        End If
    Next
End Function

Private Sub FillTextCollection()
Dim cT As clsText

    On Error GoTo err_FillTextCollection
    Set colTexte = New Collection           ' Alles vorher löschen
    If RSSchemaFields("tblText").BOF Then Exit Sub
    
    Dim RS As ADODB.Recordset
    Set RS = OpenRecordset("tblText", adOpenKeyset)
    If RS Is Nothing Then Exit Sub
    If RS.RecordCount > 0 Then
        While Not RS.EOF
            Set cT = New clsText
            
            cT.lngID = RS(0)
            cT.strTextD = IIf(IsNull(RS(1)), "TextD_" & cT.lngID, RS(1))
            cT.strTextE = IIf(IsNull(RS(2)), "TextE_" & cT.lngID, RS(2))
            colTexte.Add cT, "cT_" & cT.lngID
            RS.MoveNext
        Wend
    End If
exit_FillTextCollection:
    CloseRecordset RS
    Exit Sub
err_FillTextCollection:
    m_strText = "DB ERROR FillTextCollection err=" & Err.Number & ":" & Err.Description
    Logging "DB ERROR", , m_strText
    MsgBox m_strText
    Resume exit_FillTextCollection
End Sub

Public Function GetString(ID As Long, Optional LCID As Long = 0) As String
    ' Liefere den Text in der aktuell eingstellten Sprache - Möglichkeit Overwrite mi LCID
    Dim cT As clsText
    On Error Resume Next
    Set cT = colTexte("cT_" & ID)
    On Error GoTo 0
    If cT Is Nothing Then
        GetString = "String_" & ID
    Else
        If LCID = 0 Then LCID = basGlobals.g_iLCID
        If LCID = basGlobals.LocaleGerman Then
            GetString = cT.strTextD
        Else
            GetString = cT.strTextE
        End If
    End If
End Function

Private Sub FillStatusCollection()
Dim cS As clsStatus
    On Error GoTo err_FillStatusCollection
    Set colStatus = New Collection           ' Alles vorher löschen
    
    If basGlobals.g_iLCID = basGlobals.LocaleGerman Then
        strSQL = "SELECT lngIdxStatus, strStatus FROM tovStatus"
    Else
        strSQL = "SELECT lngIdxStatus, strStatusE FROM tovStatus"
    End If

    Dim RS As ADODB.Recordset
    Set RS = OpenRecordset(strSQL, adOpenKeyset)
    If RS Is Nothing Then Exit Sub
    If RS.RecordCount > 0 Then
        While Not RS.EOF
            Set cS = New clsStatus

            cS.lngIdxStatus = RS(0)
            cS.strStatus = RS(1)
            colStatus.Add cS, "cS_" & cS.lngIdxStatus
            RS.MoveNext
        Wend
    End If
exit_FillStatusCollection:
    CloseRecordset RS
    Exit Sub
err_FillStatusCollection:
    m_strText = "DB ERROR FillStatusCollection err=" & Err.Number & ":" & Err.Description
    Logging "DB ERROR", , m_strText
    MsgBox m_strText
    Resume exit_FillStatusCollection
End Sub
Public Function GetStatusText(IdxStatus As Long) As String
    On Error Resume Next
    GetStatusText = ""
    GetStatusText = Me.colStatus("cS_" & IdxStatus).strStatus
End Function

Private Function Checkfields() As Boolean
    On Error GoTo err_Checkfields:

    Dim DB_Version As Integer
    
    Dim query As String
    query = "SELECT valItem FROM tblItems WHERE strItem=""DB_Version"""
    
    Dim RS As ADODB.Recordset
    Set RS = OpenRecordset(query, adOpenKeyset)
    If RS Is Nothing Then Exit Function
    If RS.RecordCount = 1 Then
        DB_Version = CInt(RS.Fields(0))
    Else
        query = "INSERT INTO tblItems(strItem, valItem, strDescription) VALUES (""DB_Version"", ""1"",""State of tables and columns"")": Execute query
        DB_Version = 1
    End If
    
    If DB_Version < 2 Then
'        Schema_InsertTblFunctions
'        Schema_InsertTbxFunctionsUser
        DB_Version = 2
        query = "UPDATE tblItems set valItem = """ & DB_Version & """ WHERE stritem=""DB_Version""": Execute query
    End If
    
    If DB_Version < 3 Then
        Content_UpdateTexte
        DB_Version = 3
        query = "UPDATE tblItems set valItem = """ & DB_Version & """ WHERE stritem=""DB_Version""": Execute query
    End If
    
    If DB_Version < 4 Then
        DB_Version = 4
        query = "UPDATE tblItems set valItem = """ & DB_Version & """ WHERE stritem=""DB_Version""": Execute query
    End If
    
    If DB_Version < 5 Then
'        Create_TbxOrgItems
        DB_Version = 5
        query = "UPDATE tblItems set valItem = """ & DB_Version & """ WHERE stritem=""DB_Version""": Execute query
    End If
    
    If DB_Version < 6 Then
        Content_SetStatusText
        Content_FillAbwesenheitsarten
        Content_SetAbwesenheitsartText
        DB_Version = 6
        query = "UPDATE tblItems set valItem = """ & DB_Version & """ WHERE stritem=""DB_Version""": Execute query
    End If

    Checkfields = True
exit_Checkfields:
    CloseRecordset RS
    Exit Function
err_Checkfields:
    m_strText = "DB ERROR Checkfields err=" & Err.Number & ":" & Err.Description & " query:" & query
    Logging "DB ERROR", , m_strText
    MsgBox m_strText
    Checkfields = False
    Resume exit_Checkfields
End Function
'Private Sub Schema_InsertTblFunctions()
'    If Not TableExists("tblFunctions") Then
'        Dim cat As New ADOX.Catalog
'        Set cat.ActiveConnection = m_conn
'        Dim newtable As New ADOX.Table
'        newtable.Name = "tblFunctions"
'        With newtable.Columns
'            .Append "lngIdxFunction", adInteger
'            .Append "strFunction", adVarWChar, 30
'        End With
'        With newtable.Columns(0)
'            Set .ParentCatalog = cat
'            .Properties("Autoincrement") = True
'        End With
'        cat.Tables.Append newtable
'
'        Dim ind As ADOX.Index
'        Set ind = New ADOX.Index
'        ind.Name = "PrimaryKey"
'        ind.PrimaryKey = True
'        ind.Columns.Append "lngIdxFunction"
'
'        Dim tbl As ADOX.Table
'        Set tbl = cat.Tables("tblFunctions")
'        tbl.Indexes.Append ind
'
'        Set tbl = Nothing
'        Set ind = Nothing
'    End If
'End Sub
'Private Sub Schema_InsertTbxFunctionsUser()
'    If Not TableExists("tbxFunctionsUser") Then
'        Dim cat As New ADOX.Catalog
'        Set cat.ActiveConnection = m_conn
'        Dim newtable As New ADOX.Table
'        newtable.Name = "tbxFunctionsUser"
'        With newtable.Columns
'            .Append "lngIdxFunctionsUser", adInteger
'            .Append "lngIdxFunction", adInteger
'            .Append "lngIdxUser", adInteger
'        End With
'        With newtable.Columns(0)
'            Set .ParentCatalog = cat
'            .Properties("Autoincrement") = True
'        End With
'        cat.Tables.Append newtable
'
'        Dim ind As ADOX.Index
'        Set ind = New ADOX.Index
'        ind.Name = "PrimaryKey"
'        ind.PrimaryKey = True
'        ind.Columns.Append "lngIdxFunctionsUser"
'
'        Dim tbl As ADOX.Table
'        Set tbl = cat.Tables("tbxFunctionsUser")
'        tbl.Indexes.Append ind
'
'        Set tbl = Nothing
'        Set ind = Nothing
'    End If
'End Sub
'Private Sub Create_TbxOrgItems()
'    On Error GoTo error_Create_TbxOrgItems
'
'    If Not TableExists("tbxOrgItems") Then
'        Dim cat As New ADOX.Catalog
'        Set cat.ActiveConnection = m_conn
'        Dim newtable As New ADOX.Table
'        newtable.Name = "tbxOrgItems"
'        With newtable.Columns
'            .Append "lngIdxOrg", adInteger
'            .Append "strItem", adVarChar, 50
'            .Append "valItem", adVarChar, 50
'        End With
'
'        Dim db As Database
'        Dim tdf As TableDef
'        Dim fld As Field
'
'        Dim ind As ADOX.Index
'        Set ind = New ADOX.Index
'        ind.Name = "PrimaryKey"
'        ind.PrimaryKey = True
'        ind.Columns.Append "lngIdxOrg"
'        ind.Columns.Append "stItem"
'        ind.Unique = True
'
'        newtable.Indexes.Append ind
'
'        cat.Tables.Append newtable
'        Set newtable = Nothing
'        Set newtable = Nothing
'        Set newtable = Nothing
'    End If
'exit_Create_TbxOrgItems:
'        Exit Sub
'
'error_Create_TbxOrgItems:
'    m_strText = "DB ERROR Create_TbxOrgItems err=" & Err.Number & ":" & Err.Description
'    Logging "DB ERROR", , m_strText
'    MsgBox m_strText
'    Resume exit_Create_TbxOrgItems
'End Sub


Private Function TableExists(TableName As String) As Boolean
    Dim rstSchema As ADODB.Recordset
    Set rstSchema = RSSchemaTables  ' Schema für alle Tabellen
    
    Dim booFound As Boolean
    booFound = False
    While Not rstSchema.EOF
       If rstSchema!TABLE_NAME = TableName Then
            booFound = True
            rstSchema.MoveLast
       End If
       rstSchema.MoveNext
    Wend
    CloseRecordset rstSchema
    TableExists = booFound
End Function
Private Function ColumnsExists(ColName As String, tbl As ADOX.Table) As Boolean
    Dim i As Integer
    Dim booFound As Boolean
    booFound = False
    For i = 0 To tbl.Columns.Count - 1
        If tbl.Columns(i).name = ColName Then
            booFound = True
            Exit For
        End If
    Next
    ColumnsExists = booFound
End Function
Private Sub Content_SetStatusText()
    Dim query As String
    If GetString(1023, basGlobals.LocaleGerman) <> "String_1023" Then
        If LookupString("strStatus", "tovStatus", "lngIdxStatus=" & AbwStatus.PLANUNG) <> GetString(1023, basGlobals.LocaleGerman) Then
            query = "UPDATE tovStatus SET strStatus = """ & GetString(1023, basGlobals.LocaleGerman) & """ WHERE lngIdxStatus=" & AbwStatus.PLANUNG:        Execute query
            query = "UPDATE tovStatus SET strStatus = """ & GetString(1024, basGlobals.LocaleGerman) & "_1"" WHERE lngIdxStatus=" & AbwStatus.BEANTRAGT_1:  Execute query
            query = "UPDATE tovStatus SET strStatus = """ & GetString(1024, basGlobals.LocaleGerman) & "_2"" WHERE lngIdxStatus=" & AbwStatus.BEANTRAGT_2:  Execute query
            query = "UPDATE tovStatus SET strStatus = """ & GetString(1025, basGlobals.LocaleGerman) & """ WHERE lngIdxStatus=" & AbwStatus.GENEHMIGT:      Execute query
            query = "UPDATE tovStatus SET strStatus = """ & GetString(1149, basGlobals.LocaleGerman) & """ WHERE lngIdxStatus=" & AbwStatus.ABGELEHNT:      Execute query
            query = "UPDATE tovStatus SET strStatus = """ & GetString(1150, basGlobals.LocaleGerman) & """ WHERE lngIdxStatus=" & AbwStatus.ZURUCKGEZOGEN:  Execute query
            query = "UPDATE tovStatus SET strStatusE = """ & GetString(1023, basGlobals.LocaleEnglish) & """ WHERE lngIdxStatus=" & AbwStatus.PLANUNG:        Execute query
            query = "UPDATE tovStatus SET strStatusE = """ & GetString(1024, basGlobals.LocaleEnglish) & "_1"" WHERE lngIdxStatus=" & AbwStatus.BEANTRAGT_1:  Execute query
            query = "UPDATE tovStatus SET strStatusE = """ & GetString(1024, basGlobals.LocaleEnglish) & "_2"" WHERE lngIdxStatus=" & AbwStatus.BEANTRAGT_2:  Execute query
            query = "UPDATE tovStatus SET strStatusE = """ & GetString(1025, basGlobals.LocaleEnglish) & """ WHERE lngIdxStatus=" & AbwStatus.GENEHMIGT:      Execute query
            query = "UPDATE tovStatus SET strStatusE = """ & GetString(1149, basGlobals.LocaleEnglish) & """ WHERE lngIdxStatus=" & AbwStatus.ABGELEHNT:      Execute query
            query = "UPDATE tovStatus SET strStatusE = """ & GetString(1150, basGlobals.LocaleEnglish) & """ WHERE lngIdxStatus=" & AbwStatus.ZURUCKGEZOGEN:  Execute query
        End If
    End If
End Sub
Private Sub Content_FillAbwesenheitsarten()
    Dim query As String
    If LookupLong("lngIdxAbwesenheitsart", "tovAbwesenheitsart", "lngIdxAbwesenheitsart=" & basGlobals.ABW_SCHULFERIEN) = LNG_UNKOWN Then
        query = "INSERT INTO tovAbwesenheitsart(lngIdxAbwesenheitsart) VALUES (" & basGlobals.ABW_SCHULFERIEN & ")":    Execute query
    End If
End Sub
Private Sub Content_SetAbwesenheitsartText()
    Dim query As String
    If GetString(1031, basGlobals.LocaleGerman) <> "String_1031" Then
        If LookupString("strAbwesenheitsartE", "tovAbwesenheitsart", "lngIdxAbwesenheitsart=" & basGlobals.ABW_SCHULFERIEN) <> GetString(1012, basGlobals.LocaleEnglish) Then
            query = "UPDATE tovAbwesenheitsart SET strAbwesenheitsart = """ & GetString(1031, basGlobals.LocaleGerman) & """ WHERE lngIdxAbwesenheitsart=" & basGlobals.ABW_URLAUB:    Execute query
            query = "UPDATE tovAbwesenheitsart SET strAbwesenheitsart = """ & GetString(1032, basGlobals.LocaleGerman) & """ WHERE lngIdxAbwesenheitsart=" & basGlobals.ABW_FAKO:    Execute query
            query = "UPDATE tovAbwesenheitsart SET strAbwesenheitsart = """ & GetString(1033, basGlobals.LocaleGerman) & """ WHERE lngIdxAbwesenheitsart=" & basGlobals.ABW_SEMINAR:    Execute query
            query = "UPDATE tovAbwesenheitsart SET strAbwesenheitsart = """ & GetString(1034, basGlobals.LocaleGerman) & """ WHERE lngIdxAbwesenheitsart=" & basGlobals.ABW_DIENSTREISE:    Execute query
            query = "UPDATE tovAbwesenheitsart SET strAbwesenheitsart = """ & GetString(1035, basGlobals.LocaleGerman) & """ WHERE lngIdxAbwesenheitsart=" & basGlobals.ABW_SONSTIGES:    Execute query
            query = "UPDATE tovAbwesenheitsart SET strAbwesenheitsart = """ & GetString(1036, basGlobals.LocaleGerman) & """ WHERE lngIdxAbwesenheitsart=" & basGlobals.ABW_FEIERTAG:    Execute query
            query = "UPDATE tovAbwesenheitsart SET strAbwesenheitsart = """ & GetString(1037, basGlobals.LocaleGerman) & """ WHERE lngIdxAbwesenheitsart=" & basGlobals.ABW_KRANKHEIT:    Execute query
            query = "UPDATE tovAbwesenheitsart SET strAbwesenheitsart = """ & GetString(1012, basGlobals.LocaleGerman) & """ WHERE lngIdxAbwesenheitsart=" & basGlobals.ABW_SCHULFERIEN:    Execute query
            query = "UPDATE tovAbwesenheitsart SET strAbwesenheitsartE = """ & GetString(1031, basGlobals.LocaleEnglish) & """ WHERE lngIdxAbwesenheitsart=" & basGlobals.ABW_URLAUB:    Execute query
            query = "UPDATE tovAbwesenheitsart SET strAbwesenheitsartE = """ & GetString(1032, basGlobals.LocaleEnglish) & """ WHERE lngIdxAbwesenheitsart=" & basGlobals.ABW_FAKO:    Execute query
            query = "UPDATE tovAbwesenheitsart SET strAbwesenheitsartE = """ & GetString(1033, basGlobals.LocaleEnglish) & """ WHERE lngIdxAbwesenheitsart=" & basGlobals.ABW_SEMINAR:    Execute query
            query = "UPDATE tovAbwesenheitsart SET strAbwesenheitsartE = """ & GetString(1034, basGlobals.LocaleEnglish) & """ WHERE lngIdxAbwesenheitsart=" & basGlobals.ABW_DIENSTREISE:    Execute query
            query = "UPDATE tovAbwesenheitsart SET strAbwesenheitsartE = """ & GetString(1035, basGlobals.LocaleEnglish) & """ WHERE lngIdxAbwesenheitsart=" & basGlobals.ABW_SONSTIGES:    Execute query
            query = "UPDATE tovAbwesenheitsart SET strAbwesenheitsartE = """ & GetString(1036, basGlobals.LocaleEnglish) & """ WHERE lngIdxAbwesenheitsart=" & basGlobals.ABW_FEIERTAG:    Execute query
            query = "UPDATE tovAbwesenheitsart SET strAbwesenheitsartE = """ & GetString(1037, basGlobals.LocaleEnglish) & """ WHERE lngIdxAbwesenheitsart=" & basGlobals.ABW_KRANKHEIT:    Execute query
            query = "UPDATE tovAbwesenheitsart SET strAbwesenheitsartE = """ & GetString(1012, basGlobals.LocaleEnglish) & """ WHERE lngIdxAbwesenheitsart=" & basGlobals.ABW_SCHULFERIEN:    Execute query
       End If
    End If
End Sub
Private Sub Content_UpdateTexte()
    Dim query As String
    query = "UPDATE tbltext set strTextD=""Per Kontextmenue auf einen Benutzernamen können Sie die Identität dieses Benutzers annehmen, sofern dieser ein Mitarbeiter ist."" WHERE lngIdxText = 1015": Execute query
    query = "UPDATE tbltext set strTextE=""Right clicking via contect menu on a user name allows you to take on their identity, as long as they are an employee."" WHERE lngIdxText = 1015": Execute query
End Sub
Private Sub FillOrgCollection()     ' Hole alle Orgs aus der DB
Dim cOrg As clsOrg      ', aOrgs() As Integer
    On Error GoTo err_FillOrgCollection
    Set colOrg = Nothing
    Set colOrg = New Collection
    Dim strErrors As String

    ' Orgs werden in Level-, dann in alphabetischer Reihenfolge gelesen.
    strSQL = "SELECT * FROM tovOrg ORDER BY strOrg"

    Dim cUs As clsUser
    Dim RS As ADODB.Recordset
    Set RS = OpenRecordset(strSQL, adOpenKeyset)
    If RS Is Nothing Then Exit Sub
    If RS.RecordCount > 0 Then
        While Not RS.EOF
            If Not IsNull(RS!strOrg) Then
                If GetOrgIndexFromOrgName(RS!strOrg) > -1 Then
                    If strErrors = "" Then strErrors = "Please open Database table tovOrg." & vbCrLf
                    strErrors = strErrors & vbCrLf & "ID:" & RS!lngIdxOrg & " Orgname:" & RS!strOrg & " not unique - SKIP!"
                Else
                    Set cOrg = New clsOrg
    
                    cOrg.lngIdxOrg = RS!lngIdxOrg
                    cOrg.strOrg = RS!strOrg
                    If IsNull(RS!lngOrgLevel) Then
                        strErrors = strErrors & vbCrLf & "Please open Database table Org." & vbCrLf & vbCrLf & "Enter a Level for Org <" & RS!strOrg & ">"
                        cOrg.lngOrgLevel = Team
                    Else
                        cOrg.lngOrgLevel = RS!lngOrgLevel
                    End If
                    If IsNull(RS!lngIdxTopOrg) Then
                        strErrors = strErrors & vbCrLf & "Please open Database table tovOrg." & vbCrLf & vbCrLf & "Enter a valid TopOrg for Org <" & RS!strOrg & ">"
                        cOrg.lngIdxTopOrg = RS!lngIdxOrg    ' Verweist auf sich selbst
                    Else
                        cOrg.lngIdxTopOrg = RS!lngIdxTopOrg
                    End If
                    cOrg.lngIdxChef = IIf(IsNull(RS!lngIdxChef), 0, RS!lngIdxChef)
                    If cOrg.lngIdxChef = 0 Then
                        If strErrors = "" Then strErrors = "Please open Database table tovOrg." & vbCrLf
                        strErrors = strErrors & vbCrLf & "Enter a manager for <" & cOrg.strOrg & ">"
                    End If
                    
                    ' Referenzen zu anderen Benutzern
                    cOrg.lngIdxChef2 = IIf(IsNull(RS!lngIdxChef2), 0, RS!lngIdxChef2)
                    cOrg.lngIdxSek = IIf(IsNull(RS!lngIdxSek), 0, RS!lngIdxSek)
                    cOrg.lngIdxSek2 = IIf(IsNull(RS!lngIdxSek2), 0, RS!lngIdxSek2)
                    
                    cOrg.booVisible = False
    
                    colOrg.Add cOrg, "cOrg_" & cOrg.lngIdxOrg      ' in Collection sammeln
                End If
            End If
            RS.MoveNext
        Wend
    End If
    
exit_FillOrgCollection:
    CloseRecordset RS
    If strErrors <> "" Then MsgBox strErrors, vbOKOnly + vbExclamation, "FillOrgCollection"
    Exit Sub
err_FillOrgCollection:
    m_strText = "DB ERROR FillOrgCollection err=" & Err.Number & ":" & Err.Description
    Logging "DB ERROR", , m_strText
    MsgBox m_strText
    Resume exit_FillOrgCollection
End Sub
Private Sub CheckOrgUser(colU As Collection)
    ' Referenzen zu anderen Benutzern - nachpruefen, ob diese existieren
    Dim cOrg As clsOrg
    For Each cOrg In colOrg
        If GetUserByIDinCol(cOrg.lngIdxChef, colU) Is Nothing Then
            Logging "OrgError", , "Org:" & cOrg.strOrg & " lngIdxChef=" & cOrg.lngIdxChef & " - invalid User"
        End If
        If GetUserByIDinCol(cOrg.lngIdxChef2, colU) Is Nothing And cOrg.lngIdxChef2 <> 0 Then
            Logging "OrgError", , "Org:" & cOrg.strOrg & " lngIdxChef2=" & cOrg.lngIdxChef2 & " - invalid User"
            cOrg.lngIdxChef2 = 0
            UpdateOrg cOrg
        End If
        If GetUserByIDinCol(cOrg.lngIdxSek, colU) Is Nothing And cOrg.lngIdxSek <> 0 Then
            Logging "OrgError", , "Org:" & cOrg.strOrg & " lngIdxSek=" & cOrg.lngIdxSek & " - invalid User"
            cOrg.lngIdxSek = 0
            UpdateOrg cOrg
        End If
        If GetUserByIDinCol(cOrg.lngIdxSek2, colU) Is Nothing And cOrg.lngIdxSek2 <> 0 Then
            Logging "OrgError", , "Org:" & cOrg.strOrg & " lngIdxSek2=" & cOrg.lngIdxSek2 & " - invalid User"
            cOrg.lngIdxSek2 = 0
            UpdateOrg cOrg
        End If
    Next
End Sub
Private Sub UpdateOrg(cOrg As clsOrg)
    Dim RS As ADODB.Recordset
    Set RS = OpenRecordset("tovOrg", dbOpenDynaset)
    If RS Is Nothing Then Exit Sub
    RS.Find "lngIdxOrg=" & cOrg.lngIdxOrg
    If Not RS.EOF Then
        If cOrg.lngIdxChef2 = 0 And Not IsNull(RS!lngIdxChef2) Then
            Execute "UPDATE tovOrg Set lngIdxChef2=Null Where lngIdxOrg=" & cOrg.lngIdxOrg
        End If
        If cOrg.lngIdxSek = 0 And Not IsNull(RS!lngIdxSek) Then
            Execute "UPDATE tovOrg Set lngIdxSek=Null Where lngIdxOrg=" & cOrg.lngIdxOrg
        End If
        If cOrg.lngIdxSek2 = 0 And Not IsNull(RS!lngIdxSek2) Then
            Execute "UPDATE tovOrg Set lngIdxSek2=Null Where lngIdxOrg=" & cOrg.lngIdxOrg
        End If
    End If
End Sub
Public Function GetOrg(lngIdxOrg) As clsOrg
    Set GetOrg = Nothing
    On Error Resume Next
    Set GetOrg = colOrg("cOrg_" & lngIdxOrg)
End Function
'#######################################################################################################
Public Property Get maxOrgID() As Long  ' max für Schieber zur Org-Auswahl
    maxOrgID = UBound(aOrgVisible)
End Property
Public Function GetOrgPosition(lngOrgIndex As Long) As Long ' gegeben: cUs.lngIdxOrg gesucht: Position von ScrollOrg
    If UBound(aOrgHash) >= lngOrgIndex Then      ' Element ist vorhanden
        GetOrgPosition = aOrgHash(lngOrgIndex)
    Else
        GetOrgPosition = -1           ' Das war nix !
    End If
End Function
Public Function GetOrgIdxAtScrollPosition(lngPosition As Long) As Long
    If UBound(aOrgVisible) >= lngPosition Then      ' Element ist vorhanden
        GetOrgIdxAtScrollPosition = aOrgVisible(lngPosition)
    Else
        GetOrgIdxAtScrollPosition = -1           ' Das war nix !
    End If
End Function
Public Function GetOrgIdxFromOrgName(strOrgName As String) As Long
    Dim cOrg As clsOrg
    For Each cOrg In colOrg
        If cOrg.strOrg = strOrgName Then
            GetOrgIdxFromOrgName = cOrg.lngIdxOrg
            Exit Function
        End If
    Next
    GetOrgIdxFromOrgName = -1
End Function
'#######################################################################################################

Public Sub FillUserCollection()
    ' Voraussetzung: Datenbank offen
    ' Voraussetzung: Org sind eingelesen (Collections gefüllt)

Dim cUs As clsUser, cOrg As clsOrg
    On Error GoTo err_FillUserCollection
    
    Set colUser = New Collection
    Dim colTemp As New Collection

    FillUserIntoCollection colTemp ' alle User aus der DB lesen
    
    ' Voraussetzung Org + User
    CheckOrgUser colTemp           ' Gibt es die Benutzer, auf die in clsOrg verwiesen wird?
    
    UpdateUserCollection colTemp
    sortIntoColUser colTemp
    
    Set colTemp = Nothing

exit_FillUserCollection:
    Exit Sub
err_FillUserCollection:
    m_strText = "DB ERROR FillUserCollection err=" & Err.Number & ":" & Err.Description
    Logging "DB ERROR", , m_strText
    MsgBox m_strText
    Resume exit_FillUserCollection
End Sub
Private Function FillUserIntoCollection(col As Collection) As Boolean    ' Alle Benutzer aus der DB lesen in eine temporäre Collection
    ' Voraussetzung: Datenbank offen
    
    FillUserIntoCollection = False
    On Error GoTo err_FillUserIntoCollection

    strSQL = "SELECT * FROM tblUser ORDER BY lngIdxOrg, strSortOrder, strNachname, strVorname"

    m_MaxUserID = -1
    Dim RS As ADODB.Recordset:    Set RS = OpenRecordset(strSQL, adOpenKeyset)
    If RS Is Nothing Then Exit Function
    Dim cUs As clsUser
    If RS.RecordCount > 0 Then
        RS.MoveFirst
        While Not RS.EOF
            If Not IsNull(RS!lngIdxOrg) Then
                Set cUs = New clsUser
    
                If RS!lngIdxUser > m_MaxUserID Then m_MaxUserID = RS!lngIdxUser
                cUs.lngIdxUser = RS!lngIdxUser
                cUs.strAccountname = Trim(LCase(IIf(IsNull(RS!strAccountname), "", RS!strAccountname)))
                cUs.strAccountname2 = Trim(LCase(IIf(IsNull(RS!strAccountname2), "", RS!strAccountname2)))
                cUs.strNachname = Trim(IIf(IsNull(RS!strNachname), "", RS!strNachname))
                cUs.strVorname = Trim(IIf(IsNull(RS!strVorname), "", RS!strVorname))
                cUs.strMailName = Trim(IIf(IsNull(RS!strMailName), "", RS!strMailName))
                cUs.strSortOrder = Trim(IIf(IsNull(RS!strSortOrder), "", RS!strSortOrder))
                
                cUs.dtmDisclaimer = IIf(IsNull(RS!dtmDisclaimer), "00:00:00", RS!dtmDisclaimer)
                cUs.dtmFirst = IIf(IsNull(RS!dtmFirst), "1900-01-01", RS!dtmFirst)
                cUs.dtmLast = IIf(IsNull(RS!dtmLast), "2025-12-31", RS!dtmLast)
    
                cUs.lngIdxOrg = RS!lngIdxOrg
                cUs.lngPN = IIf(IsNull(RS!lngPN), 0, RS!lngPN)
                cUs.lngVerteiler = IIf(IsNull(RS!lngVerteiler), 0, RS!lngVerteiler)
                
                Set cUs.colMA = New Collection
                Set cUs.colFunctions = New Collection
    
                col.Add cUs, "cUs_" & cUs.lngIdxUser
            End If
            RS.MoveNext
        Wend
    End If

exit_FillUserIntoCollection:
    CloseRecordset RS
    Exit Function
err_FillUserIntoCollection:
    m_strText = "DB ERROR FillUserIntoCollection err=" & Err.Number & ":" & Err.Description
    Logging "DB ERROR", , m_strText
    MsgBox m_strText
    Resume exit_FillUserIntoCollection
End Function
Public Function GetUserByAccountname(strAccountname As String) As clsUser
    Set GetUserByAccountname = Nothing

' Erkenne aus strAccountName (LoginName oder angenommenem Account) den entsprechenden Eintrag in der Datenbank
    Dim cUs As clsUser
    For Each cUs In colUser ' als erstes die Account2 abprüfen
        If cUs.strAccountname2 = strAccountname Then
            Set GetUserByAccountname = cUs
            Exit For
        End If
    Next
    For Each cUs In colUser ' im zweien Durchgang die Erstaccounts prüfen
        If cUs.strAccountname = strAccountname Then
            Set GetUserByAccountname = cUs
            Exit For
        End If
    Next
End Function

Public Function GetUserByID(lngIdxUser As Long) As clsUser
    On Error Resume Next
    Set GetUserByID = Nothing
    Set GetUserByID = colUser("cUs_" & lngIdxUser)
End Function
Public Function GetUserByIDinCol(lngIdxUser As Long, colU As Collection) As clsUser
    On Error Resume Next
    Set GetUserByIDinCol = Nothing
    Set GetUserByIDinCol = colU("cUs_" & lngIdxUser)
End Function
Public Function GetUserSekIdx(lngIdxUser As Long) As Long
    Dim cO As clsOrg:    Set cO = GetOrg(lngIdxUser)
    GetUserSekIdx = cO.lngIdxSek
End Function
Public Function GetOrgChef(IdxOrg As Long) As Long
    GetOrgChef = 0
    Dim aOrg As clsOrg:    Set aOrg = GetOrg(IdxOrg)
    If Not aOrg Is Nothing Then GetOrgChef = aOrg.lngIdxChef
End Function
Public Function IsUserSek(lngIdxUser As Long, lngIdxUserSek As Long) As Boolean
    Dim cO As clsOrg:    Set cO = GetOrg(lngIdxUser)
    If cO.lngIdxSek = lngIdxUserSek Or cO.lngIdxSek2 = lngIdxUserSek Then
        IsUserSek = True
    Else
        IsUserSek = False
    End If
End Function
Public Function HatVertreter(IdxUser As Long) As Boolean
    Dim cUs As clsUser:    Set cUs = GetUserByID(IdxUser)
    If Not cUs Is Nothing Then
        If cUs.cOrg.lngIdxChef = IdxUser Then
            If cUs.cOrg.lngIdxChef2 > 0 Then
                HatVertreter = True
                Exit Function
            End If
        End If
        If cUs.cOrg.lngIdxChef2 = IdxUser Then
            If cUs.cOrg.lngIdxChef > 0 Then
                HatVertreter = True
                Exit Function
            End If
        End If
    End If
    HatVertreter = False
End Function
Public Function GetVertreter(IdxUser As Long) As clsUser
    Dim cUs As clsUser:    Set cUs = GetUserByID(IdxUser)
    If Not cUs Is Nothing Then
        If cUs.cOrg.lngIdxChef = IdxUser Then
            If cUs.cOrg.lngIdxChef2 > 0 Then
                Set GetVertreter = GetUserByID(cUs.cOrg.lngIdxChef2)
                Exit Function
            End If
        End If
        If cUs.cOrg.lngIdxChef2 = IdxUser Then
            If cUs.cOrg.lngIdxChef > 0 Then
                Set GetVertreter = GetUserByID(cUs.cOrg.lngIdxChef)
                Exit Function
            End If
        End If
    End If
    Set GetVertreter = Nothing
End Function
Public Function UserSeesUser(cUsToSee As clsUser, ByVal cUsLooking As clsUser) As Boolean
    ' cUsLooking schaut auf cUsToSee. Frage: Kann Looking die Abwesenheiten von ToSee sehen?
    
    ' cUsToSee hat einen Abstand zu cUsLooking (= cUsToSee.lngAbstand). Beide sehen sich nur, wenn beider Sichtbarkeit(Verteiler) >= Abstand ist
    UserSeesUser = (cUsToSee.lngAbstand <= cUsLooking.lngVerteiler And cUsToSee.lngAbstand <= cUsToSee.lngVerteiler)
    If UserSeesUser Then Exit Function
    
    ' Ich sehe meine eigenen Abwesenheiten immer
    UserSeesUser = cUsToSee.lngIdxUser = cUsLooking.lngIdxUser
    If UserSeesUser Then Exit Function
    
    ' Der Chef/Dep sieht seine MA direkt/indirekt immer
    UserSeesUser = cUsLooking.IsChefOf(cUsToSee, True, True, True, True)
    If UserSeesUser Then Exit Function
     
    ' Eine Sek sieht alle schutzbefohlenen MA
    UserSeesUser = cUsLooking.IsSekOf(cUsToSee)
    If UserSeesUser Then Exit Function
     
End Function
'###################   Functions  #####################################
Private Sub FillFunctions()
    ' Voraussetzung: Datenbank offen
    
    On Error GoTo err_FillFunctions
    Dim cFu As clsFunction
    Dim RS As ADODB.Recordset

    Set colFunctions = New Collection
    
    strSQL = "SELECT * FROM tblFunctions"
    Set RS = OpenRecordset(strSQL, adOpenKeyset)
    If RS Is Nothing Then Exit Sub
    If RS.RecordCount > 0 Then
        While Not RS.EOF
            If Not IsNull(RS!lngIdxFunction) Then
                If Not IsNull(RS!strFunction) Then
                    Set cFu = New clsFunction
                    cFu.IdxFunction = RS!lngIdxFunction
                    cFu.strFunction = RS!strFunction
                    Set cFu.colUser = New Collection
                    colFunctions.Add cFu, "cFu_" & cFu.IdxFunction
                End If
            End If
            RS.MoveNext
        Wend
    End If
    
exit_FillFunctions:
    CloseRecordset RS
    Exit Sub
err_FillFunctions:
    m_strText = "DB ERROR FillFunctions err=" & Err.Number & ":" & Err.Description
    Logging "DB ERROR", , m_strText
    MsgBox m_strText
    Resume exit_FillFunctions
End Sub
Private Sub FillFunctionsUser()
    ' Voraussetzung: Datenbank offen
    ' Voraussetzung: Functions und User sind eingelesen (Collections gefüllt)
    
    On Error GoTo err_FillFunctionsUser
    Dim cFu As clsFunction, cUs As clsUser
    Dim RS As ADODB.Recordset
    
    strSQL = "SELECT * FROM tbxFunctionsUser"
    
    Set RS = OpenRecordset(strSQL, adOpenKeyset)
    If RS Is Nothing Then Exit Sub
    If RS.RecordCount > 0 Then
        While Not RS.EOF
            If Not IsNull(RS!lngIdxFunction) Then
                If Not IsNull(RS!lngIdxUser) Then
                    ' Funktion und User ermitteln
                    On Error Resume Next
                    Set cFu = colFunctions("cFu_" & RS!lngIdxFunction)
                    Set cUs = GetUserByID(CLng(RS!lngIdxUser))
                    On Error GoTo err_FillFunctionsUser
                    If Not cFu Is Nothing And Not cUs Is Nothing Then
                        cFu.colUser.Add cUs, "cUs_" & cUs.lngIdxUser   ' User zu den Functions hinzu
                        cUs.colFunctions.Add cFu, "cFu_" & cFu.IdxFunction  ' Funktion zum User dazu
                    End If
                End If
            End If
            RS.MoveNext
        Wend
    End If
    
exit_FillFunctionsUser:
    CloseRecordset RS
    Exit Sub
err_FillFunctionsUser:
    m_strText = "DB ERROR FillFunctionsUser err=" & Err.Number & ":" & Err.Description
    Logging "DB ERROR", , m_strText
    MsgBox m_strText
    Resume exit_FillFunctionsUser
End Sub
Private Sub FillFunctionsOrg()
    ' Voraussetzung: Functions und User sind eingelesen (Collections gefüllt)
    ' Voraussetzung: FunctionsUser sind eingelesen (Collections gefüllt)
    
    On Error GoTo err_FillFunctionsOrg
    
    ' Füge zu den tbxFunctionsUser die Funktionen der Org-Chefs hinzu
    Dim cUs As clsUser, cOrg As clsOrg
    For Each cOrg In colOrg
        Dim cFu As clsFunction
        Set cFu = GetFunctionByName("Chief " & cOrg.strOrg)
        
        Set cUs = GetUserByID(cOrg.lngIdxChef)  ' Chef1
        If cUs Is Nothing Then Exit For     ' Kein Chef eingetragen
        
        cFu.colUser.Add cUs, "cUs_" & cUs.lngIdxUser   ' User zu den Functions hinzu
        cUs.colFunctions.Add cFu, "cFu_" & cFu.IdxFunction  ' Funktion zum User dazu
    
        If cOrg.lngIdxChef2 > 0 Then
            Set cUs = GetUserByID(cOrg.lngIdxChef2)  ' Chef2
            If Not cUs Is Nothing Then
                cFu.colUser.Add cUs, "cUs_" & cUs.lngIdxUser   ' User zu den Functions hinzu
                cUs.colFunctions.Add cFu, "cFu_" & cFu.IdxFunction  ' Funktion zum User dazu
            End If
        End If
    Next

exit_FillFunctionsOrg:
    Exit Sub
err_FillFunctionsOrg:
    m_strText = "DB ERROR FillFunctionsOrg err=" & Err.Number & ":" & Err.Description
    Logging "DB ERROR", , m_strText
    MsgBox m_strText
    Resume exit_FillFunctionsOrg

End Sub

Private Sub UpdateOrgSetTopOrg()    ' Einmaliger Durchlauf
    ' Voraussetzung: Datenbank offen
    ' Voraussetzung: Org sind eingelesen (Collections gefüllt)
    
    On Error GoTo err_UpdateOrgSetTopOrg
    ' cTopOrg wird bei UpdateUserCollection benötigt
    Dim cOrg As clsOrg
    For Each cOrg In colOrg
        Set cOrg.cTopOrg = GetOrg(cOrg.lngIdxTopOrg) ' setzte die TopOrg für jede Org
        If Not cOrg.cTopOrg Is Nothing Then
            cOrg.lngIdxTopTopOrg = cOrg.cTopOrg.lngIdxTopOrg
        End If
    Next
exit_UpdateOrgSetTopOrg:
    Exit Sub
err_UpdateOrgSetTopOrg:
    m_strText = "DB ERROR UpdateOrgSetTopOrg err=" & Err.Number & ":" & Err.Description
    Logging "DB ERROR", , m_strText
    MsgBox m_strText
    Resume exit_UpdateOrgSetTopOrg
End Sub
Private Sub UpdateOrgCollection()   ' Einmaliger Durchlauf
    ' Voraussetzung: Datenbank offen
    ' Voraussetzung: Org und User sind eingelesen (Collections gefüllt)
    
    On Error GoTo err_UpdateOrgCollection
' Fülle die Collection der Unterorganisationen und die Collection der direkten MA, Vererbe die SekIDs, definiere für jede Org den Platz im Baum
    Dim cOrg As clsOrg, cOrgSub As clsOrg, cUs As clsUser
    Dim colOrgHi As Collection, minLevel As Integer    ' die höchste Organisation in der Sammlung
    minLevel = 5:
    For Each cOrg In colOrg
        cOrg.lngIdxRes = -1
        cOrg.lngIdxBer = -1
        cOrg.lngIdxAbt = -1
        cOrg.lngIdxTeam = -1
        cOrg.lngIdxInsel = -1
        
        ' Collection mit den höchsten Orgs erstellen - eine oder auch mehrere
        If cOrg.lngOrgLevel = minLevel Then ' gleich groß -> sammeln
            colOrgHi.Add cOrg    ' höchste Org merken
        ElseIf cOrg.lngOrgLevel < minLevel Then ' kleiner als bisher gefunden
            Set colOrgHi = New Collection   ' Alte Collection löschen
            minLevel = cOrg.lngOrgLevel     ' Auf neuer Ebene anfangen
            colOrgHi.Add cOrg    ' höchste Org merken
        End If
        
        ' Untergeordnete Orgs einer Org in einer Collection sammeln
        Set cOrg.colSubOrg = New Collection
        For Each cOrgSub In colOrg
            If cOrgSub.lngIdxTopOrg = cOrg.lngIdxOrg Then
                If cOrgSub.lngIdxOrg <> cOrgSub.lngIdxTopOrg Then
                    cOrg.colSubOrg.Add cOrgSub
                End If
            End If
        Next
        
        ' Alle MA in dieser Org: Leiter + Stab in Collection sammeln
        Set cOrg.colMember = New Collection
        For Each cUs In colUser
            If cUs.lngIdxOrg = cOrg.lngIdxOrg Then
                cOrg.colMember.Add cUs
            End If
        Next
    Next
    
    ' Von oben nach unten - rekursiv Sekretärinnen verteilen
    For Each cOrg In colOrgHi   ' Jede Top-Org ...
        SetSek cOrg
    Next
    
    For Each cOrg In colOrgHi   ' Die Orgs mit der kleinsten Nummer - 1 oder mehrere
        SetOrgPositionIndices cOrg
    Next
    
    ' Das Ergebnis
    For Each cOrg In colOrg
        Debug.Print cOrg.lngIdxOrg, cOrg.strOrg, cOrg.lngIdxRes, cOrg.lngIdxBer, cOrg.lngIdxAbt, cOrg.lngIdxTeam, cOrg.lngIdxInsel
    Next

exit_UpdateOrgCollection:
    Exit Sub
err_UpdateOrgCollection:
    m_strText = "DB ERROR UpdateOrgCollection err=" & Err.Number & ":" & Err.Description
    Logging "DB ERROR", , m_strText
    MsgBox m_strText
    Resume exit_UpdateOrgCollection
End Sub
Private Sub SetOrgPositionIndices(cOrg As clsOrg)
    Dim cOrgTop As clsOrg, cOrgSub As clsOrg
    ' Vererben von oben
    If cOrg.lngIdxTopOrg <> cOrg.lngIdxOrg Then ' ich bin nicht der höchste - vererben von oben
        Set cOrgTop = GetOrg(cOrg.lngIdxTopOrg)
        cOrg.lngIdxRes = cOrgTop.lngIdxRes
        cOrg.lngIdxBer = cOrgTop.lngIdxBer
        cOrg.lngIdxAbt = cOrgTop.lngIdxAbt
        cOrg.lngIdxTeam = cOrgTop.lngIdxTeam
        cOrg.lngIdxInsel = cOrgTop.lngIdxInsel
    End If
    ' neue Ebene definieren
    Select Case cOrg.lngOrgLevel
            Case basGlobals.OrgLevel.Ressort
                cOrg.lngIdxRes = cOrg.lngIdxOrg
            Case basGlobals.OrgLevel.Bereich
                cOrg.lngIdxBer = cOrg.lngIdxOrg
            Case basGlobals.OrgLevel.Abteilung
                cOrg.lngIdxAbt = cOrg.lngIdxOrg
            Case basGlobals.OrgLevel.Team
                cOrg.lngIdxTeam = cOrg.lngIdxOrg
            Case basGlobals.OrgLevel.Insel
                cOrg.lngIdxInsel = cOrg.lngIdxInsel
    End Select
    ' vererben nach unten
    If Not cOrg.colSubOrg Is Nothing Then
        For Each cOrgSub In cOrg.colSubOrg
            SetOrgPositionIndices cOrgSub
        Next
    End If
End Sub
Private Sub SetSek(cOrg As clsOrg)
    Dim cSubOrg As clsOrg
    For Each cSubOrg In cOrg.colSubOrg  ' für meine Org sind die Seks klar, setze für die "kinder"
        If cSubOrg.lngIdxSek = 0 Then     ' keine Sek1 eingetragen
            If cOrg.lngIdxSek > 0 Then
                cSubOrg.lngIdxSek = cOrg.lngIdxSek ' Sek1 übertragen
            End If
        End If
        If cSubOrg.lngIdxSek = 0 Then     ' immer noch keine Sek1 eingetragen
            If cOrg.lngIdxSek2 > 0 Then cSubOrg.lngIdxSek = cOrg.lngIdxSek2 ' Sek2 übertragen nach Sek1, wenn vorige Zeile kein Ergebnis hatte
        End If
        If cSubOrg.lngIdxSek > 0 Then ' Sek1 ist gefüllt, Sek2 versuchen
            If cSubOrg.lngIdxSek2 = 0 Then    ' Keine Sek2 eingetragen
                If cOrg.lngIdxSek2 > 0 And cOrg.lngIdxSek <> cSubOrg.lngIdxSek2 Then    ' cOrg.sekID2 steht nicht schon als Sek1 drin
                    cSubOrg.lngIdxSek2 = cOrg.lngIdxSek2 ' Sek2 übertragen
                End If
            End If
        End If
        If cSubOrg.colSubOrg.Count > 0 Then
            SetSek cSubOrg   ' für die untergeordneten Orgs genauso verfahren
        End If
    Next
End Sub
Private Sub UpdateUserCollection(col As Collection)
    ' Voraussetzung: Datenbank offen
    ' Voraussetzung: Org und User sind eingelesen (Collections gefüllt)

' Aus Userdaten und Org-Daten die Hierarchie herstellen
' Sortkey festlegen für die Darstellungsreihenfolge
' lngOrgLevel: Level meiner Org: Team,Abt,Bereich,Ressort
' lngUserLevel: MEIN Level: User, TL, AL, BL, RL
' strSortLevel: BL, BL_SEK,  AL, AL_SEK, TL, TL_SEK, USER
    On Error GoTo err_UpdateUserCollection
    
    Dim cUs As clsUser
    For Each cUs In col
        Set cUs.cOrg = GetOrg(cUs.lngIdxOrg)                    ' MEINE Organisation
        If cUs.cOrg Is Nothing Then
            MsgBox "Die Datenbank ist defekt: Benutzer " & cUs.strAccountname & " ist keiner Organisation zugeordnet. " & App.ExeName & " funktioniert NICHT.", vbCritical
            Exit Sub
        End If
    Next

    
    ' Aus allen Org die Chefs, Deputys und Seks kennzeichnen
    Dim op As clsOrgLevel
    Dim cOrg As clsOrg
    For Each cOrg In colOrg
        For Each cUs In col
            If cUs.lngIdxUser = cOrg.lngIdxChef Then
                Set op = New clsOrgLevel
                Set op.cOrg = cOrg
                op.Level = 1
                cUs.colOrg.Add op
                cUs.booIsOrgChef = True ' Chef irgendeiner Org - oder auch mehrerer
            End If
            If cUs.lngIdxUser = cOrg.lngIdxChef2 Then
                Set op = New clsOrgLevel
                Set op.cOrg = cOrg
                op.Level = 2
                cUs.colOrg.Add op
                cUs.booIsOrgChef2 = True    ' Deputy irgendeiner Org - oder auch mehrerer - kann gleichzeitig OrgChef sein
            End If
            If cUs.lngIdxUser = cOrg.lngIdxSek Or cUs.lngIdxUser = cOrg.lngIdxSek2 Then
                cUs.booIsSek = True
            End If
        Next
    Next
    
    Dim cUsChef As clsUser
    For Each cUs In col
        cUs.lngIdxChef = cUs.cOrg.lngIdxChef                    ' Default:Chef dieser Org ist mein Chef
        cUs.lngIdxChef2 = cUs.cOrg.lngIdxChef2                  ' Default:Chef dieser Org ist mein Chef
        'If cUs.strAccountname = "zwehl" Then Stop
        
        ' ----------------------    Chef oder User ?    -------------------------------------------------------------
        'If cUs.cOrg.lngIdxChef = cUs.lngIdxUser Then cUs.booIsOrgChef = True    ' in dieser Org
        'If cUs.cOrg.lngIdxChef2 = cUs.lngIdxUser Then cUs.booIsOrgChef2 = True  ' in dieser Org
        If cUs.IsOrgChef(cUs.cOrg, False) Then
            ' TODO Wenn Chef in mehreren Orgs, dann die höchste nehmen
            cUs.lngIdxChef = cUs.cOrg.cTopOrg.lngIdxChef    ' mein Chef ist der Chef der übergeordneten Org
            cUs.lngIdxChef2 = cUs.cOrg.cTopOrg.lngIdxChef2    ' mein Chef ist der Chef der übergeordneten Org
        End If
        
        If cUs.IsOrgChef(cUs.cOrg, False) Then ' ich bin in dieser Org Chef, also TL, AL, BL, ..., nicht Deputy
            cUs.lngUserLevel = cUs.cOrg.lngOrgLevel         ' das Org-Level dieser Org ist auch meiner
            If cUs.lngUserLevel = Bereich Then          ' Sortierung
                cUs.strSortLevel = basGlobals.SORTORDER_BL
            ElseIf cUs.lngUserLevel = Abteilung Then
                cUs.strSortLevel = basGlobals.SORTORDER_AL
            ElseIf cUs.lngUserLevel = Team Then
                cUs.strSortLevel = basGlobals.SORTORDER_TL
            End If
        Else                                                 ' default - Sachbearbeiter
            cUs.lngUserLevel = Benutzer
            cUs.strSortLevel = basGlobals.SORTORDER_USER
        
            ' Eintragen in MA-Liste(n)
            If cUs.lngIdxChef > 0 Then  ' Sollte immer der Fall sein, kann aber fehlerhaft weggelassen werden
                Set cUsChef = col("cUs_" & cUs.lngIdxChef)
                If Not cUsChef Is Nothing Then cUsChef.colMA.Add cUs    ' Ich bin MA bei diesem Chef
            End If
            If cUs.lngIdxChef2 > 0 Then  ' Ist manchmal der Fall
                On Error Resume Next
                Set cUsChef = col("cUs_" & cUs.lngIdxChef2)
                If Not cUsChef Is Nothing Then cUsChef.colMA.Add cUs    ' Ich bin MA bei diesem Vertreter
                On Error GoTo err_UpdateUserCollection
            End If
        End If
        
        ' ----------------------    Sek ?    -------------------------------------------------------------
        If cUs.IsOrgSek(cUs.cOrg) Then           ' ich bin hier Sek
            ' TODO Wenn Sek in mehreren Orgs, dann die höchste nehmen
            cUs.lngUserLevel = cUs.cOrg.lngOrgLevel                     ' das Org-Level dieser Org ist auch meiner
            If cUs.cOrg.lngOrgLevel = basGlobals.OrgLevel.Bereich Then       ' Sortierung
                cUs.strSortLevel = basGlobals.SORTORDER_BL_SEK
            ElseIf cUs.cOrg.lngOrgLevel = basGlobals.OrgLevel.Abteilung Then
                cUs.strSortLevel = basGlobals.SORTORDER_AL_SEK
            ElseIf cUs.cOrg.lngOrgLevel = basGlobals.OrgLevel.Team Then
                cUs.strSortLevel = basGlobals.SORTORDER_TL_SEK
            End If
        End If
        
        Select Case cUs.cOrg.lngOrgLevel
            Case basGlobals.OrgLevel.Ressort
                cUs.lngIdxRes = cUs.cOrg.lngIdxOrg
                cUs.lngIdxBer = -1
                cUs.lngIdxAbt = -1
                cUs.lngIdxTeam = -1
                cUs.lngIdxInsel = -1
            Case basGlobals.OrgLevel.Bereich  ' Bereich
                cUs.lngIdxRes = GetOrgID(cUs.cOrg.lngIdxOrg, basGlobals.OrgLevel.Ressort)
                cUs.lngIdxBer = cUs.cOrg.lngIdxOrg
                cUs.lngIdxAbt = -1
                cUs.lngIdxTeam = -1
                cUs.lngIdxInsel = -1
            Case basGlobals.OrgLevel.Abteilung  ' Abteilung
                cUs.lngIdxRes = GetOrgID(cUs.cOrg.lngIdxOrg, basGlobals.OrgLevel.Ressort)
                cUs.lngIdxBer = GetOrgID(cUs.cOrg.lngIdxOrg, basGlobals.OrgLevel.Bereich)
                cUs.lngIdxAbt = cUs.cOrg.lngIdxOrg
                cUs.lngIdxTeam = -1
                cUs.lngIdxInsel = -1
            Case basGlobals.OrgLevel.Team  ' Team
                cUs.lngIdxRes = GetOrgID(cUs.cOrg.lngIdxOrg, basGlobals.OrgLevel.Ressort)
                cUs.lngIdxBer = GetOrgID(cUs.cOrg.lngIdxOrg, basGlobals.OrgLevel.Bereich)
                cUs.lngIdxAbt = GetOrgID(cUs.cOrg.lngIdxOrg, basGlobals.OrgLevel.Abteilung)
                cUs.lngIdxTeam = cUs.cOrg.lngIdxOrg
                cUs.lngIdxInsel = -1
            Case basGlobals.OrgLevel.Insel  ' Insel
                cUs.lngIdxRes = GetOrgID(cUs.cOrg.lngIdxOrg, basGlobals.OrgLevel.Ressort)
                cUs.lngIdxBer = GetOrgID(cUs.cOrg.lngIdxOrg, basGlobals.OrgLevel.Bereich)
                cUs.lngIdxAbt = GetOrgID(cUs.cOrg.lngIdxOrg, basGlobals.OrgLevel.Abteilung)
                cUs.lngIdxTeam = GetOrgID(cUs.cOrg.lngIdxOrg, basGlobals.OrgLevel.Team)
                cUs.lngIdxInsel = cUs.cOrg.lngIdxOrg
        End Select

        cUs.strSortKey = cUs.cOrg.strOrg & "!" & cUs.strSortLevel & "!" & cUs.strSortOrder & "!" & cUs.strNachname & "!" & cUs.strVorname
        Debug.Print cUs.strSortKey, "R:" & cUs.lngIdxRes & "B:" & cUs.lngIdxBer, "A:" & cUs.lngIdxAbt, "T:" & cUs.lngIdxTeam, "Org:" & cUs.lngIdxOrg
NextUser:
    Next
    
    Exit Sub
err_UpdateUserCollection:
    m_strText = "DB ERROR UpdateUserCollection: err=" & Err.Number & ":" & Err.Description
    Logging "DB ERROR", , m_strText
    MsgBox m_strText
    Exit Sub
End Sub
Private Sub sortIntoColUser(col As Collection)
    ' Verschiebe die Benutzer aus Col in colUser - Reihenfolge ist cUs.strSortKey
    Dim strTopSort As String, lngIdxTopUser As Long
    Dim cUs As clsUser
ReStart:
    strTopSort = "ZZZZZZ"
    If col.Count > 0 Then   ' ist noch was drin
        For Each cUs In col     ' alle durchlaufen und kleinsten suchen
            If cUs.strSortKey < strTopSort Then
                strTopSort = cUs.strSortKey
                lngIdxTopUser = cUs.lngIdxUser
            End If
        Next
        ' Gefunden - in colUser einfügen und aus col löschen
        Set cUs = col("cUs_" & lngIdxTopUser)
        colUser.Add cUs, "cUs_" & cUs.lngIdxUser
        col.Remove ("cUs_" & lngIdxTopUser)
        GoTo ReStart
    End If
End Sub
Private Function GetFunctionByName(name As String) As clsFunction
    Dim cFu As clsFunction, MaxID As Integer
    For Each cFu In colFunctions
        If cFu.IdxFunction > MaxID Then MaxID = cFu.IdxFunction
        If cFu.strFunction = name Then
            Set GetFunctionByName = cFu
            Exit Function
        End If
    Next
    
    ' nicht gefunden, anlegen
    Set cFu = New clsFunction
    cFu.IdxFunction = MaxID + 1
    cFu.strFunction = name
    Set cFu.colUser = New Collection
    colFunctions.Add cFu, "cFu_" & cFu.IdxFunction
    Set GetFunctionByName = cFu
End Function
Private Function GetOrgID(lngIdxOrg As Long, OrgLevel As Long) As Long
' liefere ID der OU zu dieser OU_ID - verwendet, um Abteilung/Bereich zu Team/Abt zu finden
    GetOrgID = -1
    Dim cOrg As clsOrg
    Set cOrg = GetOrg(lngIdxOrg)
    If cOrg.lngOrgLevel = OrgLevel Then
        GetOrgID = cOrg.lngIdxOrg
        Exit Function
    End If
    
    ' zum gewünschten Orglevel "aufsteigen"
    While cOrg.lngIdxTopOrg <> cOrg.lngIdxOrg   ' noch nicht oben angekommen
        If cOrg.cTopOrg.lngOrgLevel = OrgLevel Then
            GetOrgID = cOrg.cTopOrg.lngIdxOrg
            Exit Function
        End If
        Set cOrg = cOrg.cTopOrg     ' die nächsthöhere versuchen
    Wend
End Function
Public Function GetOrgIndexFromOrgName(strName As String) As Long
    Dim cOrg As clsOrg
    For Each cOrg In colOrg
        If cOrg.strOrg = strName Then
            GetOrgIndexFromOrgName = cOrg.lngIdxOrg
            Exit Function
        End If
    Next
    GetOrgIndexFromOrgName = -1
End Function
Public Function IsSekID(lngIdxUser) As Boolean  ' Ist diese UserID eine Sek-ID?
    IsSekID = False
    Dim cOrg As clsOrg
    For Each cOrg In colOrg
        If cOrg.lngIdxSek = lngIdxUser Or cOrg.lngIdxSek2 = lngIdxUser Then
            IsSekID = True
            Exit Function
        End If
    Next
End Function
Public Function GetSekOrgID(lngIdxOrg As Long, colUs As Collection) As Long    ' nächste OU, in der es eine Sek gibt
    GetSekOrgID = -1
    Dim cOrg As clsOrg
    Dim cUs As clsUser
    
    On Error Resume Next
    Set cOrg = GetOrg(lngIdxOrg)          ' Organisation
    If cOrg Is Nothing Then Exit Function
    If cOrg.lngIdxSek > 0 Then    ' In der Organisation ist eine Sekretärin eingetragen
        Set cUs = Nothing
        Set cUs = colUs("cUs_" & cOrg.lngIdxSek)    ' Es gibt den Benutzer wirklich
        If Not cUs Is Nothing Then
            GetSekOrgID = cOrg.lngIdxOrg
            Exit Function
        End If
    End If
    While cOrg.lngIdxTopOrg <> cOrg.lngIdxOrg                   ' noch nicht oben angekommen
        Set cOrg = GetOrg(cOrg.lngIdxTopOrg)          ' der nächste oberhalb
        If cOrg Is Nothing Then Exit Function                   ' Das war ne Pleite - sollte gar nicht vorkommen
        If cOrg.lngIdxSek > 0 Then                            ' In der Organisation ist eine Sekretärin eingetragen
            Set cUs = Nothing
            Set cUs = colUs("cUs_" & cOrg.lngIdxSek)          ' Es gibt den Benutzer wirklich
            If Not cUs Is Nothing Then
                GetSekOrgID = cOrg.lngIdxOrg
                Exit Function
            End If
        End If
    Wend
End Function
Public Function GetE4OrgID(lngIdxorgID As Long, colUs As Collection) As Long    ' alle TL dieser Org
    Dim cOrg As clsOrg
    Dim cUs As clsUser
    On Error GoTo err_GetE4OrgID
    
    For Each cOrg In colOrg
        If cOrg.lngIdxTopOrg = lngIdxorgID And cOrg.lngOrgLevel = Team Then ' TL in dieser Org
            Set cUs = GetUserByID(cOrg.lngIdxChef)
            If Not cUs Is Nothing Then
                colUs.Add cUs
            End If
        End If
    Next
   
    GetE4OrgID = colUs.Count
    Exit Function
err_GetE4OrgID:
    GetE4OrgID = -1
    Exit Function
End Function
Public Function GetAL(cUs As clsUser) As clsUser
    On Error Resume Next
    Set GetAL = Nothing
    Dim cAbt As clsOrg: Set cAbt = GetOrg(cUs.lngIdxAbt)
    If Not cAbt Is Nothing Then Set GetAL = GetUserByID(cAbt.lngIdxChef)
End Function
Private Sub SetOrgAbstand()
    Dim cOrg As clsOrg
    For Each cOrg In colOrg     ' initialisierung - noch nicht erledigt
        cOrg.intAbstand = -1
    Next
    
    Set cOrg = GetOrg(g_CU.lngIdxOrg)
    cOrg.intAbstand = 0
    SetOrgAbstandWandern cOrg
    
    Debug.Print "CurrentOrg = " & cOrg.strOrg
    For Each cOrg In colOrg
        Debug.Print cOrg.lngIdxOrg, cOrg.strOrg, cOrg.lngIdxRes, cOrg.lngIdxBer, cOrg.lngIdxAbt, cOrg.lngIdxTeam, cOrg.lngIdxInsel, cOrg.intAbstand
    Next

End Sub
Public Function KrankheitUserBeantragbar() As Boolean
        Dim cABA As clsAbwesenheitsart:    Set cABA = GetAbwesenheitsart(ABW_KRANKHEIT)
        If cABA Is Nothing Then
            KrankheitUserBeantragbar = False
        Else
            KrankheitUserBeantragbar = cABA.booUserBeantragbar
        End If
End Function
Public Function KrankheitSichtbar(Optional cAb As clsAbw) As Boolean
    ' 1. Hürde: im System überhaupt zugelassen?
    Dim booShow As Boolean: booShow = False
    
    If g_CU.booIsSek Then                                   ' angewählter User = SEK -> darf alles
        booShow = True
    Else
        If Not cAb Is Nothing Then
            If g_CU.lngIdxUser = cAb.cUs.lngIdxUser Then    ' Der Benutzer sieht seine eigenen Krankheiten
                booShow = True
            End If
        End If
        
        If g_CU.IsOrgChef(cAb.cUs.cOrg, True) Then          ' Führungskräfte oder Deputy
            Dim booIsMyChief As Boolean
                        
            If cAb Is Nothing Then    ' unspezifisch
                booIsMyChief = True
            Else
                booIsMyChief = g_CU.IsChefOf(cAb.cUs)
            End If
            
            If booIsMyChief Then
                ' eigene AL und aufwärts
                If g_CU.lngUserLevel <= Abteilung Then      ' Ab AL Aufwärts darf Krankheit sehen
                    booShow = True
                ' eigene TL und IL, wenn DB eingestellt
                ElseIf g_CU.lngUserLevel >= Team Then       ' TL/IL und DB-Einstellung
                    booShow = g_db.GetItem("booE4SiehtKrankheit", "False", "Team leader can see sickness? Be careful!")
                End If
            End If
        End If
    End If
    KrankheitSichtbar = booShow
End Function

Private Sub SetOrgAbstandWandern(cOrg As clsOrg)
' Berechne den Abstand aller anderen Org-Einheiten zur Org de g_CU
' Jede Org besteht aus Leader und MA in den Ebenen 1=Ressort bis 4=Team oder 5=Insel
    Dim cTopOrg As clsOrg, cSubOrg As clsOrg
    
    ' hoch
    If cOrg.lngIdxTopOrg <> cOrg.lngIdxOrg Then
        Set cTopOrg = GetOrg(cOrg.lngIdxTopOrg)
        If cTopOrg.intAbstand = -1 Then ' noch nicht bearbeitet
            cTopOrg.intAbstand = cOrg.intAbstand + 1
            SetOrgAbstandWandern cTopOrg
        End If
    End If
    
    ' runter
    If Not cOrg.colSubOrg Is Nothing Then
        For Each cSubOrg In cOrg.colSubOrg
            If cSubOrg.intAbstand = -1 Then
                cSubOrg.intAbstand = cOrg.intAbstand + 1
                SetOrgAbstandWandern cSubOrg
            End If
        Next
    End If
End Sub
Private Sub SetUserAbstand()
    ' Berechne den Abstand aller anderen Benutzer innerhalb der Collection zu diesem Benutzer g_CU
    ' Variante 1: Im selben Team/Stab: Abstand 2
    '             In derselben Abteilung: Abstand 3
    '             Im Selben Bereich: Abstand 4
    '             Im selben Ressort: Abstand 5
    On Error GoTo err_SetUserAbstand
    
    Dim cUs As clsUser
    For Each cUs In colUser
        If cUs.lngIdxOrg = g_CU.lngIdxOrg Then                                          '### dasselbe Team/Abt/Ber/Res-OU
            If cUs.lngIdxUser = g_CU.lngIdxUser Then
                cUs.lngAbstand = 1                          ' ich bin es selbst
            Else
                cUs.lngAbstand = 2                          ' ein Kollege
            End If
            '#################################################################################################################################
        ElseIf g_CU.cOrg.lngOrgLevel = basGlobals.OrgLevel.Team Then                                            ' g_CU im Team: TL oder TeamMember
            If cUs.lngIdxAbt = g_CU.lngIdxAbt And cUs.lngIdxAbt <> -1 Then              '####  dieselbe Abteilung, anderes Team
                If cUs.lngUserLevel = Abteilung Then ' AL
                ElseIf cUs.lngUserLevel = Team Then ' TL
                Else
                End If
                
                cUs.lngAbstand = SICHTBAR_ABT
            ElseIf cUs.lngIdxBer = g_CU.lngIdxBer And cUs.lngIdxBer <> -1 Then          '#### derselbe Bereich, andere Abteilung
                cUs.lngAbstand = SICHTBAR_BER
            Else
                cUs.lngAbstand = SICHTBAR_RES
            End If
            '#################################################################################################################################
        ElseIf g_CU.cOrg.lngOrgLevel = basGlobals.OrgLevel.Abteilung Then                      ' g_CU in Abteilung: AL oder AbteilungsStab
            If cUs.lngIdxAbt = g_CU.lngIdxAbt And cUs.lngIdxAbt <> -1 Then          '####  dieselbe Abteilung, anderes Team
                If cUs.lngUserLevel = Team Then   ' Teamleiter in der eigenen Abteilung
                    cUs.lngAbstand = SICHTBAR_Org
                Else                                                ' TeamMitarbeiter
                    cUs.lngAbstand = SICHTBAR_ABT
                End If
            ElseIf cUs.lngIdxBer = g_CU.lngIdxBer And cUs.lngIdxBer <> -1 Then          '#### derselbe Bereich, andere Abteilung
                If cUs.cOrg.lngOrgLevel = basGlobals.OrgLevel.Team Then    ' aus der TeamEbene
                    If cUs.lngUserLevel = Team Then   ' TL einer Nachbar-Abteilung
                        cUs.lngAbstand = SICHTBAR_ABT
                    Else                                                ' MA einer Nachbar-Abteilung
                        cUs.lngAbstand = SICHTBAR_BER
                    End If
                ElseIf cUs.cOrg.lngOrgLevel = basGlobals.OrgLevel.Abteilung Then    ' AL einer Nachbar-Abteilung
                    cUs.lngAbstand = SICHTBAR_Org
                ElseIf cUs.cOrg.lngOrgLevel = basGlobals.OrgLevel.Bereich Then
                    If cUs.lngUserLevel = Bereich Then   ' mein Chef Bereichsleiter oder Ber-Sek
                        cUs.lngAbstand = SICHTBAR_Org
                    Else                                                ' StabsMA
                        cUs.lngAbstand = SICHTBAR_ABT
                    End If
                ElseIf cUs.cOrg.lngOrgLevel = basGlobals.OrgLevel.Ressort Then
                    cUs.lngAbstand = SICHTBAR_BER
                End If
            Else
                cUs.lngAbstand = SICHTBAR_RES
            End If
            '#################################################################################################################################
        ElseIf g_CU.cOrg.lngOrgLevel = basGlobals.OrgLevel.Bereich Then                                             ' g_CU in Bereich: BL oder BereichStab
            If cUs.lngIdxBer = g_CU.lngIdxBer Then  ' im selben Bereich aber nicht BL/BerStab
                If cUs.cOrg.lngOrgLevel = basGlobals.OrgLevel.Abteilung Then
                    If cUs.lngUserLevel = Abteilung Then              ' AL
                        cUs.lngAbstand = SICHTBAR_ABT
                    Else                                                ' Abt Stab
                        cUs.lngAbstand = SICHTBAR_ABT
                    End If
                ElseIf cUs.cOrg.lngOrgLevel = basGlobals.OrgLevel.Team Then
                    If cUs.lngUserLevel = Team Then    ' TL
                        cUs.lngAbstand = SICHTBAR_BER
                    Else                                                ' Team User
                        cUs.lngAbstand = SICHTBAR_BER
                    End If
                End If
            Else    ' nicht im selben Bereich
                ' gleiches Ressort und Bereichsebene:
                cUs.lngAbstand = SICHTBAR_RES
            End If
            '#################################################################################################################################
        ElseIf g_CU.cOrg.lngOrgLevel = basGlobals.OrgLevel.Ressort Then                                             ' g_CU in Ressort: RL oder RessortStab
            '#################################################################################################################################
        End If
    Next
    
exit_SetUserAbstand:
    On Error Resume Next
    Exit Sub
err_SetUserAbstand:
    m_strText = "DB ERROR SetUserAbstand err=" & Err.Number & ":" & Err.Description
    Logging "DB ERROR", , m_strText
    MsgBox m_strText
    Resume exit_SetUserAbstand
End Sub
Public Sub setOrgUserVisibility()
    On Error GoTo err_setOrgUserVisibility
    
    SetOrgAbstand
    SetUserAbstand
    ' Setze die Sichtbarkeit der Org-Member
    
    ' Org, die bei dieser User-ID angezeigt werden sollen -> frmInfo
    
    ' Gegeben: UserID -> ORD_ID
    
    ' sichtbar sein sollen :
    ' 1. Zeile: alle Benutzer in diesem Org
    ' 2. Zeile: alle TL/AL  innerhalb dieser Abteilung
    ' 3. Zeile: Wenn Superuser alle Benutzer sehen soll und der Benutzer SEK oder AL ist
'    If cUs.lngIdxOrg = m_lngIdxOrg _
'        Or cUs.corg.lngOrgLevel <= orglevel.team And cUs.lngIdxAbt = basGlobals.g_CU.lngIdxAbt _
'        Or g_booAlleMaSichtbar And (g_CU.booIsOrgChef) _


    Dim cUs As clsUser, cCOrg As clsOrg
    Set cCOrg = GetOrg(g_CU.lngIdxOrg)    ' die aktuelle Organisation
    
    ' Variablen zur Zuordnung OrgSlider -> OrgIdx
    Dim intAnzOrg As Integer
    intAnzOrg = 0
    ReDim aOrgHash(0)       ' löschen
    ReDim aOrgVisible(0)    ' löschen
    
    Dim cOrg As clsOrg  ' Laufvariable
    For Each cOrg In colOrg     ' Sichtbarkeit der Organisationen - erst mal unsichtbar
        cOrg.booVisible = False
        If cOrg.lngIdxOrg = g_CU.lngIdxOrg Then     ' Meine Org IMMER sichtbar
            cOrg.booVisible = True
        Else
            For Each cUs In cOrg.colMember          ' Alle Member in der Org
                If cUs.lngAbstand <= g_CU.lngVerteiler Then   ' Ist einer für mich sichtbar?
                    cOrg.booVisible = True
                    Exit For
                End If
            Next
        End If
        If g_CU.IsOrgSek(cOrg) Then   ' eventuell zusätzliche Sichtbarkeit als Sek
            cOrg.booVisible = True
        End If
'        For Each cUs In cOrg.colMember
'            cUs.booVisible = cOrg.booVisible And g_CU.lngVerteiler > basGlobals.SICHTBAR_CHEF    ' UserSichtbar=OrgSichtbar, wenn nicht Sichtbar=0
'        Next
        
        If cOrg.booVisible Then     ' OrgSlider ->OrgIndex
            intAnzOrg = intAnzOrg + 1
            ReDim Preserve aOrgVisible(intAnzOrg)
            aOrgVisible(intAnzOrg) = cOrg.lngIdxOrg             ' an Scroll-Position intAnzOrg steht corg.lngIdxOrg
            If UBound(aOrgHash) < cOrg.lngIdxOrg Then ReDim Preserve aOrgHash(cOrg.lngIdxOrg)
            aOrgHash(cOrg.lngIdxOrg) = intAnzOrg                ' für die Reverse-Suche: Auf welchem Platz ist die Org-ID?
        End If
    Next

exit_setOrgUserVisibility:
    On Error Resume Next
    Exit Sub
err_setOrgUserVisibility:
    m_strText = "DB ERROR setOrgUserVisibility err=" & Err.Number & ":" & Err.Description
    Logging "DB ERROR", , m_strText
    MsgBox m_strText
    Resume exit_setOrgUserVisibility
End Sub
Public Function GetUserCollection(lngIdxThisOrg As Long) As Collection
    ' Benutzer, die bei der Anwahl DIESER IdxOrg angezeigt wrden
    On Error GoTo err_GetUserCollection
    
    Set GetUserCollection = Nothing
    Dim colUserOut As Collection
    Dim cThisOrg As clsOrg
    Set cThisOrg = GetOrg(lngIdxThisOrg)
    
    Dim cUs As clsUser, cOrgUser As clsOrg  ' Laufvariable
    
    Dim booAdd As Boolean
    Dim booShowDeputySec As Boolean
    booShowDeputySec = g_db.GetItem("ShowDeputySec", "False", "Show deputy and sec")
    
    Set colUserOut = New Collection
    For Each cUs In colUser
        booAdd = False
        If basGlobals.g_booAlleMaSichtbar And g_CU.lngUserLevel <= Abteilung Then   ' Spezialfall: Ab AL aufwärts werden Abteilungen als Ganzes angezeigt
            If cThisOrg.lngOrgLevel = basGlobals.OrgLevel.Abteilung Then    ' wenn ich AL oder Abt-Sek bin, dann die ganze Abteilung
                booAdd = (cUs.lngIdxAbt = cThisOrg.lngIdxOrg)
            ElseIf cThisOrg.lngOrgLevel <= basGlobals.OrgLevel.Bereich Then  ' Ressort oder BL oder BL-Sek
                If cUs.lngIdxOrg = cThisOrg.lngIdxOrg Then                   ' Kollegen in DERSELBEN ORG
                    booAdd = True
                Else
                    If cUs.IsOrgChef(cUs.cOrg, booShowDeputySec) Then                         ' Chef einer Org
                        If cUs.lngIdxOrg = cThisOrg.lngIdxTopOrg Then          ' 1 Level drüber
                            booAdd = True                                       ' Chef der ÜBERGEORDNETEN Org
                        ElseIf cUs.cOrg.lngIdxTopOrg = cThisOrg.lngIdxOrg Then  ' 1 Level drunter
                            booAdd = True                                       ' Chef(s) der NÄCHSTNIEDRIGEEN Org
                        ElseIf cUs.cOrg.lngIdxTopOrg = cThisOrg.lngIdxTopOrg Then  ' Gemeinsame TopOrg
                                booAdd = True  ' Der gleiche Chef - selbe Ebene    Chef(s) der PARALLELEN Org
                        ElseIf cUs.cOrg.lngIdxTopTopOrg = cThisOrg.lngIdxOrg Then    ' 2 Level drunter
                                booAdd = True     ' Chef(s) der NÄCHST-NÄCHST-NIEDRIGEEN Org
                        End If
                    End If
                End If
            End If
'            booAdd = InOrg(cThisOrg, cUs)
'            booAdd = UserSeesUser(g_CU, cUs)
        Else  ' Standardfall: Jede Organisation wird einzeln angeboten
            If cUs.lngIdxOrg = cThisOrg.lngIdxOrg Then              ' Kollegen in DERSELBEN ORG
                booAdd = True
            Else
                If cUs.IsOrgChef(cUs.cOrg, booShowDeputySec) Or booShowDeputySec And cUs.IsOrgSek(cUs.cOrg) Then ' Chef/Deputy/Sek
                    If cUs.lngIdxOrg = cThisOrg.lngIdxTopOrg Then       ' cUs ist in TopOrg
                        booAdd = True   ' TopOrg: Mein Chef und die Sekretariate anzeigen ' Chef der ÜBERGEORDNETEN Org
                    ElseIf cUs.cOrg.lngIdxTopOrg = cThisOrg.lngIdxOrg Then   ' cUs ist in SubOrg
                        booAdd = True                                           ' Chef der NÄCHSTNIEDRIGEEN Org
                    ElseIf cUs.cOrg.lngIdxTopOrg = cThisOrg.lngIdxTopOrg Then ' Gemeinsame TopOrg
                        booAdd = True  ' Der gleiche Chef - selbe Ebene       ' Chef der PARALLELEN Org
                    End If
                End If
            End If
        End If

        If booAdd Then
            colUserOut.Add cUs                       ' Ja, dann auch diesen Benutzer rein
        End If
    Next
    
    Set GetUserCollection = colUserOut
    
exit_GetUserCollection:
    On Error Resume Next
    Exit Function
err_GetUserCollection:
    m_strText = "DB ERROR GetUserCollection err=" & Err.Number & ":" & Err.Description
    Logging "DB ERROR", , m_strText
    MsgBox m_strText
    Resume exit_GetUserCollection
End Function
Public Function InOrg(cOrg As clsOrg, cUs As clsUser) As Boolean
    Select Case cOrg.lngOrgLevel
        Case basGlobals.OrgLevel.Insel
            InOrg = (cOrg.lngIdxOrg = cUs.lngIdxInsel)
        Case basGlobals.OrgLevel.Team
            InOrg = (cOrg.lngIdxOrg = cUs.lngIdxTeam)
        Case basGlobals.OrgLevel.Abteilung
            InOrg = (cOrg.lngIdxOrg = cUs.lngIdxAbt)
        Case basGlobals.OrgLevel.Bereich
            InOrg = (cOrg.lngIdxOrg = cUs.lngIdxBer)
        Case basGlobals.OrgLevel.Ressort
            InOrg = (cOrg.lngIdxOrg = cUs.lngIdxRes)
        Case Else
            InOrg = False
    End Select
End Function
Private Sub FillSchulferienCollection()
Dim cSF As clsSchulferien
    On Error GoTo err_FillSchulferienCollection
    If Not colSchulferien Is Nothing Then
        For Each cSF In colSchulferien
            Set cSF = Nothing
        Next cSF
        Set colSchulferien = Nothing
    End If
    Set colSchulferien = New Collection
    
    Dim RS As ADODB.Recordset
'    Set rs = OpenRecordset("SELECT * FROM tovSchulferien WHERE dtmSchulferien > #" & Format(Date, "yyyy-mm-dd") & "# ORDER BY dtmSchulferien", adOpenKeyset)
    Set RS = OpenRecordset("SELECT tblSchulferien.*, tovBundesland.strBundesland " & _
                            " FROM tblSchulferien INNER JOIN tovBundesland ON tblSchulferien.lngIdxBundesland = tovBundesland.lngIdxBundesland " & _
                            "ORDER BY dtmStart, strBundesland", adOpenKeyset)
    If RS Is Nothing Then Exit Sub
    If RS.RecordCount > 0 Then
        While Not RS.EOF
            Set cSF = New clsSchulferien
            cSF.lngIdxSchulferien = RS!lngIdxSchulferien
            cSF.dtmStart = IIf(Not IsNull(RS!dtmStart), RS!dtmStart, 0)
            cSF.dtmEnde = IIf(Not IsNull(RS!dtmEnde), RS!dtmEnde, 0)
            cSF.lngIdxBundesland = IIf(Not IsNull(RS!lngIdxBundesland), RS!lngIdxBundesland, 0)
            cSF.strBundesland = IIf(Not IsNull(RS!strBundesland), RS!strBundesland, "")
            On Error Resume Next
            colSchulferien.Add cSF, "cSF_" & RS!lngIdxSchulferien
            On Error GoTo err_FillSchulferienCollection
            RS.MoveNext
        Wend
    End If
exit_FillSchulferienCollection:
    CloseRecordset RS
    Exit Sub
err_FillSchulferienCollection:
    m_strText = "DB ERROR FillSchulferienCollection err=" & Err.Number & ":" & Err.Description
    Logging "DB ERROR", , m_strText
    MsgBox m_strText
    Resume exit_FillSchulferienCollection
End Sub

Public Function Bundesländer() As String
    Bundesländer = ""
    
    On Error GoTo err_Bundesländer
    Dim RS As ADODB.Recordset
    Set RS = OpenRecordset(" SELECT * FROM tovBundesland ORDER By strBundesland")
    If RS Is Nothing Then Exit Function
    If RS.RecordCount > 0 Then
        While Not RS.EOF
            Bundesländer = Bundesländer & RS!strBundesland & "#"
            RS.MoveNext
        Wend
    End If
exit_Bundesländer:
    CloseRecordset RS
    Exit Function
err_Bundesländer:
    m_strText = "DB ERROR Bundesländer err=" & Err.Number & ":" & Err.Description
    Logging "DB ERROR", , m_strText
    MsgBox m_strText
    Resume exit_Bundesländer
End Function

Public Sub Logging(strAction As String, Optional cAb As clsAbw = Nothing, Optional strText As String = "")
    On Error GoTo err_Logging
    
    Dim RS As ADODB.Recordset
    Set RS = OpenRecordset("tblLog", adOpenDynamic)
    If RS Is Nothing Then Exit Sub
    RS.AddNew
    RS!strComputername = g_strProgramVersion & " " & basSysUtils.HostName()
    If g_CU_Login Is Nothing Then
        RS!strUsername = basSysUtils.CurrentUser
    Else
    
        If basSysUtils.CurrentUser <> g_CU_Login.strAccountname Then
            RS!strUsername = g_CU_Login.strAccountname & "." & basSysUtils.CurrentUser
        Else
            RS!strUsername = g_CU_Login.strAccountname
        End If
    End If
    RS!dtmTimeStamp = Now()
    RS!strAction = strAction
    If g_CU Is Nothing Then
        RS!lngIdxUser = 0
    Else
        RS!lngIdxUser = g_CU.lngIdxUser
    End If
    If Not cAb Is Nothing Then
        RS!lngIdxAbwesenheit = cAb.lngIdxAbwesenheit
        RS!lngIdxAbwesenheitsArt = cAb.lngIdxAbwesenheitsArt
        RS!dtmStart = cAb.dtmStart
        RS!dtmEnde = cAb.dtmEnde
        RS!lngGVN = cAb.lngGVN
        RS!lngIdxStatus = cAb.lngIdxStatus
    End If
    If strText <> "" Then RS!strText = strText
    RS.Update

exit_Logging:
    CloseRecordset RS
    Exit Sub
err_Logging:
    m_strText = "ToLog::" & strText & vbCrLf & "DB ERROR Logging err=" & Err.Number & ":" & Err.Description
    MsgBox m_strText
    Resume exit_Logging
End Sub
Public Function FillManagerListe(booZeigenBeantragte As Boolean, booDeputy As Boolean, booInDirect As Boolean, strOrder As String) As Collection
' called by frmManager
' booZeigenBeantragte:  True: Status Abwesenheit = beantragt1 oder beantragt 2
' booDeputy:            True: Zeige NUR MA, die ich als Funktion "stellvertretender XL" sehen soll
' booInDirect:          False: Zeige NUR MA, die mit direkt zugeordnet sind     True: Zeige auch MA, die mir nicht direkt zugeordnet sind

    On Error GoTo err_FillManagerListe
    Set FillManagerListe = New Collection   ' Collection anlegen
    Dim colChef As New Collection           ' Die Liste der Chefs - kann bei Sekretariaten >1 sein

' Liste mit den direkten und indirekten MA einer Person zusammenstellen für frmManager
    Dim cOrg As clsOrg, cOrgSub As clsOrg, cUs As clsUser, cUSub As clsUser
    Set cOrg = GetOrg(g_CU.lngIdxOrg)     ' Start.Org
    
    Dim CU_Org_is_TopOrg As Boolean
    CU_Org_is_TopOrg = (g_CU.lngIdxOrg = GetOrg(g_CU.lngIdxOrg).lngIdxTopOrg)
    If g_CU.IsSek Then
        Dim cOrgRun As clsOrg
        For Each cOrgRun In colOrg
            If g_CU.IsOrgSek(cOrgRun) Then
                ' ich bin Sek -> Zeige MA meines(r) Chefs
                ' bin ich bereits in der übergeordneten Org Sek, dann Skip, es sei denn, meine Org ist die TopOrg
                If Not g_CU.IsOrgSek(GetOrg(cOrgRun.lngIdxTopOrg)) Or CU_Org_is_TopOrg Then
                    colChef.Add GetUserByID(cOrgRun.lngIdxChef)
                End If
            End If
        Next
    Else
        ' selber Chef - zeige meine MA - Alle TL, AL, BL, ...
        colChef.Add GetUserByID(g_CU.lngIdxUser)
    End If
    
    Dim cC As clsUser

    Dim strUIDs1 As String:    strUIDs1 = ""    ' Direkte
    Dim strUIDs2 As String:    strUIDs2 = ""    ' Indirekte
    
    ' Haben wir es mit einem Vorgesetzten zu tun?
    ' stelle alle direkten MA fest: Stabsangehörige sowie eventuell Leiter niedrigerer Orgs
    For Each cUs In colUser
        For Each cC In colChef
            If cC.IsChefOf(cUs, booBoss:=Not booDeputy, booStellvertreter:=booDeputy, booDirekt:=True, booIndirekt:=False) Then
                strUIDs1 = strUIDs1 & cUs.lngIdxUser & ","
            End If
            If booInDirect Then   ' nicht nur direkte
                If cC.IsChefOf(cUs, booBoss:=Not booDeputy, booStellvertreter:=booDeputy, booDirekt:=False, booIndirekt:=True) Then
                    strUIDs2 = strUIDs2 & cUs.lngIdxUser & ","
                End If
            End If
        Next
    Next
    If strUIDs1 <> "" Then strUIDs1 = Left(strUIDs1, Len(strUIDs1) - 1) ' remove trailing ","
    If strUIDs2 <> "" Then strUIDs2 = Left(strUIDs2, Len(strUIDs2) - 1) ' remove trailing ","
    
    ' Wenn Chef der Top-Org, dann muss OrgChef leider selbst genehmigen
    If cOrg.lngIdxTopOrg = cOrg.lngIdxOrg Or CBool(GetItem("booBossSelfApprove", "False", "Boss approves his own absences. (def=false)")) Then
        If g_CU.IsOrgChef(cOrg, False) Then
            If strUIDs1 <> "" Then strUIDs1 = strUIDs1 & ","
            strUIDs1 = strUIDs1 & cOrg.lngIdxChef
        End If
    End If
    
    If strUIDs1 = "" And strUIDs2 = "" Then Exit Function
    
    Dim strSQLselect As String, strSQLfrom As String, strSQLwhere As String, strSQLorder As String
    strSQLselect = "SELECT tblAbwesenheit.lngIdxAbwesenheit "
    strSQLfrom = " FROM tovStatus " & _
        "INNER JOIN (tovAbwesenheitsArt " & _
        "INNER JOIN (tblUser " & _
        "INNER JOIN tblAbwesenheit " & _
        "ON tblUser.lngIdxUser = tblAbwesenheit.lngIdxUser) " & _
        "ON tovAbwesenheitsArt.lngIdxAbwesenheitsArt = tblAbwesenheit.lngIdxAbwesenheitsArt) " & _
        "ON tovStatus.lngIdxStatus = tblAbwesenheit.lngIdxStatus"
    If booZeigenBeantragte Then
        Dim strW1 As String, strW2 As String
        If strUIDs1 <> "" Then strW1 = "(tblUser.lngIdxUser in (" & strUIDs1 & ") AND tblAbwesenheit.lngIdxStatus=" & AbwStatus.BEANTRAGT_1 & ")"
        If strUIDs2 <> "" Then
            strW2 = "(tblUser.lngIdxUser in (" & strUIDs2 & ") AND tblAbwesenheit.lngIdxStatus=" & AbwStatus.BEANTRAGT_1 & " OR tblAbwesenheit.lngIdxStatus=" & AbwStatus.BEANTRAGT_2 & ")"
            If CBool(g_db.GetItem("TravellingRequestToE3", "False", "Travelling request to E3 instead of E4")) Then
                strW2 = "(tblUser.lngIdxUser in (" & strUIDs2 & ") AND (tblAbwesenheit.lngIdxStatus=" & AbwStatus.BEANTRAGT_1 & " OR tblAbwesenheit.lngIdxStatus=" & AbwStatus.BEANTRAGT_2 & " OR tblAbwesenheit.lngIdxAbwesenheitsart=" & basGlobals.ABW_DIENSTREISE & "))"
            End If
        End If
        
        strSQLwhere = " WHERE( "
        If strW1 <> "" Then strSQLwhere = strSQLwhere & strW1
        If strW2 <> "" And strW1 <> "" Then strSQLwhere = strSQLwhere & " OR "
        If strW2 <> "" Then strSQLwhere = strSQLwhere & strW2
        strSQLwhere = strSQLwhere & ")"
    Else
        strSQLwhere = " WHERE (tblUser.lngIdxUser in ("
        If strUIDs1 <> "" Then strSQLwhere = strSQLwhere & strUIDs1
        If strUIDs2 <> "" And strUIDs1 <> "" Then strSQLwhere = strSQLwhere & ","
        If strUIDs2 <> "" Then strSQLwhere = strSQLwhere & strUIDs2
        strSQLwhere = strSQLwhere & "))"
    End If

    Dim intDays As Integer
    intDays = GetItem("NumDaysBackApprove", 1, "Absences with end date today minus this can be approved")

    strSQLwhere = strSQLwhere & " AND tblAbwesenheit.dtmEnde >= now() -" & intDays   ' Auch den heutigen Tag mitnehmen (eigentlich >= date)
    strSQLorder = " ORDER BY " & strOrder
    strSQL = strSQLselect & strSQLfrom & strSQLwhere & strSQLorder
    
    Dim cAb As clsAbw
    Dim RS As ADODB.Recordset
    Set RS = OpenRecordset(strSQL, adOpenStatic)
    If RS Is Nothing Then Exit Function
    If RS.RecordCount > 0 Then
        While Not RS.EOF
            Set cAb = GetAbwesenheit(RS!lngIdxAbwesenheit)
            If Not cAb Is Nothing Then FillManagerListe.Add cAb
            RS.MoveNext
        Wend
    End If

exit_FillManagerListe:
    CloseRecordset RS
    Exit Function
err_FillManagerListe:
    m_strText = "DB ERROR FillManagerListe err=" & Err.Number & ":" & Err.Description
    Logging "DB ERROR", , m_strText
    MsgBox m_strText
    Resume exit_FillManagerListe
End Function
Public Function FillUserListe(lngIdxUser As Long) As Collection
    Set FillUserListe = New Collection
    Dim cAb As clsAbw
    For Each cAb In colAbwesenheiten
        If cAb.lngIdxUser = lngIdxUser Then
            FillUserListe.Add cAb
        End If
    Next
End Function
Public Function FillUserListeKW(lngIdxUser As Long, lngIdxAbwesenheit As Long) As Collection
    ' Überschneidungen dieser Abwesenheit für die Org dieses Benutzers
    Dim cA As clsAbw:   Set cA = GetAbwesenheit(lngIdxAbwesenheit)
    Dim cUs As clsUser:  Set cUs = GetUserByID(lngIdxUser)

    Set FillUserListeKW = New Collection
    Dim cAb As clsAbw
    For Each cAb In colAbwesenheiten
        If cAb.cUs.lngIdxOrg = cUs.lngIdxOrg Then    ' dieselbe Org?
            ' 1. Zeile: fängt vor dem Start an und hört nach dem Start auf  ' um den Start herum
            ' 2. Zeile: fängt vor dem Ende an und hört nach dem Ende auf    ' um das Ende herum
            ' 3. Zeile: fängt vor dem Start an und hört nach dem Ende auf   ' um alles herum
            If cAb.dtmStart <= cA.dtmStart And cAb.dtmEnde >= cA.dtmStart Or _
               cAb.dtmStart <= cA.dtmEnde And cAb.dtmEnde >= cA.dtmEnde Or _
               cAb.dtmStart <= cA.dtmStart And cAb.dtmEnde >= cA.dtmEnde _
               Then
                FillUserListeKW.Add cAb
            End If
        End If
    Next
    
'    strSQL = strSQL & " WHERE ("
'    strSQL = strSQL & "    dtmStart <= " & strDatumStart & " AND dtmEnde >= " & strDatumStart
'    strSQL = strSQL & " OR dtmStart <= " & strDatumEnd & " AND dtmEnde >= " & strDatumEnd
'    strSQL = strSQL & " OR dtmStart <= " & strDatumStart & " AND dtmEnde >= " & strDatumEnd
'    strSQL = strSQL & " )"
End Function
Public Function WriteJahresSumme(lngYear As Long) As String
Dim lngAnzTage As Long, lngAnzFAKO As Long, lngAnzUrlaub As Long
Dim lngUrlaub As Long, lngFAKO As Long, lngTage As Long
Dim aAbw() As Single
Dim aYear(366) As Long, aUserY(366) As Long
Dim strSQL As String

    WriteJahresSumme = ""

    Dim dtmFirstDay As Date:    dtmFirstDay = DateSerial(lngYear, 1, 1)
    Dim dtmLastDay As Date:     dtmLastDay = DateSerial(lngYear, 12, 31)
    Dim strFirstDay As String:
    strFirstDay = "#" & Format(dtmFirstDay, "yyyy\/mm\/dd") & "#"
    Dim strLastDay As String:
    strLastDay = "#" & Format(dtmLastDay, "yyyy\/mm\/dd") & "#"
    
    ' Basistabelle aYear mit Feiertagen füllen
    Dim i As Integer
    Dim cF As clsFeiertag
    For Each cF In colFeiertage
        If Year(cF.dtmFeiertag) = lngYear Then
            i = DateDiff("d", dtmFirstDay, cF.dtmFeiertag) + 1
            aYear(i) = cF.lngIdxAbwesenheitsArt
        End If
    Next
    
    Dim colU As New Collection
    If g_CU.IsOrgChef(g_CU.cOrg, True) Or g_CU.IsOrgSek(g_CU.cOrg) Then  ' orgInfo für Chef, Deputy  oder Sek
        Dim cUs As clsUser
        For Each cUs In colUser
            If InOrg(g_CU.cOrg, cUs) Then colU.Add cUs
        Next
    Else
        colU.Add g_CU   ' EinzelInfo
    End If
    
    
    g_PopUpMode = "Tabelle"
    g_InfoText = "Sum of absences"
    Dim Spaltenzahl As Integer: Spaltenzahl = 7
    ReDim g_InfoTabelle(Spaltenzahl)  ' Spaltenbreiten
    ReDim g_DataTabelle(Spaltenzahl, colU.Count + 1)
    g_InfoTabelle(0) = 1000:    g_InfoTabelle(1) = 1000:    g_InfoTabelle(2) = 1000:    g_InfoTabelle(3) = 1000
    g_InfoTabelle(4) = 1000:    g_InfoTabelle(5) = 1000:    g_InfoTabelle(6) = 1000:    g_InfoTabelle(7) = 1000
    
    g_DataTabelle(0, 0) = "Org":    g_DataTabelle(1, 0) = "Name"
    g_DataTabelle(2, 0) = g_db.GetString(1031):    g_DataTabelle(3, 0) = g_db.GetString(1032):    g_DataTabelle(4, 0) = g_db.GetString(1033)
    g_DataTabelle(5, 0) = g_db.GetString(1034):    g_DataTabelle(6, 0) = g_db.GetString(1035):    g_DataTabelle(7, 0) = "pct vac"
    
    Dim lngFGRow As Long
    lngFGRow = 0

    
    Dim SumLevel() As Single ' Summe Urlaub, Anzahl User
    ReDim SumLevel(2, colOrg.Count) ' wenn die Args dicht gepackt sind, reicht die Anzahl. Sonst erweitern
    
    WriteJahresSumme = WriteJahresSumme & "Team; User; Urlaub; FAKO; Seminar; Dienstreise; Sonstiges;"
    WriteJahresSumme = WriteJahresSumme & " Working Days until now; Working Days from tomorrow; PN;"
    WriteJahresSumme = WriteJahresSumme & vbCrLf
    
    Dim cAb As clsAbw
    For Each cUs In colU
        For i = 1 To 366
            aUserY(i) = aYear(i)    ' Reset
        Next
        
        strSQL = "SELECT * FROM tblAbwesenheit "
        strSQL = strSQL & "WHERE dtmStart >= " & strFirstDay & " AND dtmEnde <= " & strLastDay
        strSQL = strSQL & " AND lngIdxUser=" & cUs.lngIdxUser
        strSQL = strSQL & " AND lngIdxStatus <> " & AbwStatus.ABGELEHNT & " "   ' Keine gelöschten
        strSQL = strSQL & " AND lngIdxStatus <> " & AbwStatus.ZURUCKGEZOGEN & " "   ' Keine zurückgezogenen
        strSQL = strSQL & "ORDER BY dtmStart"

        Dim RS As ADODB.Recordset
        Set RS = OpenRecordset(strSQL, adOpenKeyset)
        If RS Is Nothing Then Exit Function
        If RS.RecordCount > 0 Then
            While Not RS.EOF
                WriteAbwToArray RS!lngIdxAbwesenheitsArt, RS!dtmStart, RS!dtmEnde, RS!lngGVN, dtmFirstDay, aUserY
                RS.MoveNext
            Wend
        End If
        
        ReDim aAbw(8)   ' löschen
        SumAbw dtmFirstDay, aUserY, aAbw
        
        ' String befüllen
        WriteJahresSumme = WriteJahresSumme & cUs.cOrg.strOrg & ";" & cUs.strAccountname & ";" & aAbw(ABW_URLAUB) & ";" & aAbw(ABW_FAKO) & ";" & aAbw(ABW_SEMINAR) & ";" & aAbw(ABW_DIENSTREISE) & ";" & aAbw(ABW_SONSTIGES)
        WriteJahresSumme = WriteJahresSumme & ";" & aAbw(7) & ";" & aAbw(8) & ";" & cUs.lngPN
        WriteJahresSumme = WriteJahresSumme & vbCrLf
        ' Tabelle befüllen
        lngFGRow = lngFGRow + 1
        g_DataTabelle(0, lngFGRow) = cUs.cOrg.strOrg
        g_DataTabelle(1, lngFGRow) = cUs.strAccountname
        g_DataTabelle(2, lngFGRow) = aAbw(ABW_URLAUB)
        g_DataTabelle(3, lngFGRow) = aAbw(ABW_FAKO)
        g_DataTabelle(4, lngFGRow) = aAbw(ABW_SEMINAR)
        g_DataTabelle(5, lngFGRow) = aAbw(ABW_DIENSTREISE)
        g_DataTabelle(6, lngFGRow) = aAbw(ABW_SONSTIGES)
        g_DataTabelle(7, lngFGRow) = Fix(aAbw(ABW_URLAUB) / 3 * 100) / 10   ' Prozent Urlaub von 30 AT
        
        Debug.Print cUs.cOrg.strOrg, cUs.strAccountname, aAbw(ABW_URLAUB), aAbw(ABW_SEMINAR), aAbw(ABW_DIENSTREISE), aAbw(ABW_SONSTIGES), Fix(aAbw(ABW_URLAUB) / 3 * 100) / 10
        
        ' für diese Org aufsummieren
        If cUs.cOrg.lngIdxOrg > UBound(SumLevel, 2) Then    ' vielleicht reicht das Array nicht ...
            ReDim Preserve SumLevel(2, cUs.cOrg.lngIdxOrg)
        End If
        SumLevel(1, cUs.cOrg.lngIdxOrg) = SumLevel(1, cUs.cOrg.lngIdxOrg) + aAbw(ABW_URLAUB)
        SumLevel(2, cUs.cOrg.lngIdxOrg) = SumLevel(2, cUs.cOrg.lngIdxOrg) + 1
    Next
    
ExitCalc:
    CloseRecordset RS
    
    If g_CU.IsOrgChef(g_CU.cOrg, True) Or g_CU.IsOrgSek(g_CU.cOrg) Then   ' orgInfo - Zusammenfassung
        Dim Run As Integer
        Dim pct As Single
        ReDim Preserve g_DataTabelle(Spaltenzahl, UBound(g_DataTabelle, 2) + 1)   ' Leerzeile
        ReDim Preserve g_DataTabelle(Spaltenzahl, UBound(g_DataTabelle, 2) + 1)   ' Leerzeile
        lngFGRow = UBound(g_DataTabelle, 2)
        g_DataTabelle(0, lngFGRow) = "Org"
        g_DataTabelle(1, lngFGRow) = "#member"
        g_DataTabelle(2, lngFGRow) = "#vacation"
        g_DataTabelle(3, lngFGRow) = "pct"
        WriteJahresSumme = WriteJahresSumme & vbCrLf & "Org;#member;#vacation;pct" & vbCrLf
        
        Dim cOrg As clsOrg
        For Run = 0 To UBound(SumLevel, 2)
            If SumLevel(2, Run) > 0 Then
                On Error Resume Next
                Set cOrg = GetOrg(Run)
                If Not cOrg Is Nothing Then
                    pct = Fix(SumLevel(1, Run) / SumLevel(2, Run) / 3 * 100) / 10
                    WriteJahresSumme = WriteJahresSumme & cOrg.strOrg & ";" & SumLevel(2, Run) & ";" & SumLevel(1, Run) & ";" & pct & vbCrLf
                    ReDim Preserve g_DataTabelle(Spaltenzahl, UBound(g_DataTabelle, 2) + 1)
                    lngFGRow = UBound(g_DataTabelle, 2)
                    g_DataTabelle(0, lngFGRow) = cOrg.strOrg
                    g_DataTabelle(1, lngFGRow) = SumLevel(2, Run)
                    g_DataTabelle(2, lngFGRow) = SumLevel(1, Run)
                    g_DataTabelle(3, lngFGRow) = pct
                    
                    Debug.Print cOrg.strOrg, SumLevel(1, Run), SumLevel(2, Run), pct
                End If
            End If
        Next
    End If

    Exit Function
ErrCalc:
    MsgBox "DB ERROR WriteJahresSumme: " & Err.Description
    Resume ExitCalc
End Function

Private Sub WriteAbwToArray(lngIdxAbwesenheitsArt As Long, dtmCalStart As Date, dtmCalEnde As Date, GVN As Long, dtmFirstDay As Date, ByRef aYear() As Long)
Dim iStart As Integer, iEnd As Integer, iRun As Integer
    iStart = DateDiff("d", dtmFirstDay, dtmCalStart) + 1
    iEnd = iStart + dtmCalEnde - dtmCalStart
    If iStart < 1 Then
        iStart = 1   ' Übergang Jahreswechsel
    End If
    If iEnd > 366 Then
        iEnd = 366
    End If
    If GVN <> ABW_GANZTAGS Then
        lngIdxAbwesenheitsArt = lngIdxAbwesenheitsArt + 100
    End If
    For iRun = iStart To iEnd
        If aYear(iRun) < ABW_FEIERTAG Then aYear(iRun) = lngIdxAbwesenheitsArt    ' Überschreiben, wenn nicht Feiertag
    Next
End Sub
Private Sub SumAbw(dtmFirstDay As Date, aYear() As Long, ByRef aAbw() As Single)
Dim iRun As Integer, iMaxRun As Integer
Dim iDOY As Integer
Dim iSaturday As Integer

    ReDim aAbw(8)
    iSaturday = 8 - Weekday(dtmFirstDay)   ' 7 = saturday
    iDOY = DateDiff("d", CDate("1/1/" & Year(dtmFirstDay)), Date) + 1
    iMaxRun = DateDiff("d", CDate("1/1/" & Year(dtmFirstDay)), CDate("1/1/" & Year(dtmFirstDay) + 1))
    For iRun = 1 To iMaxRun
        If iRun = iSaturday Then
            iSaturday = iSaturday + 7
            '''If aYear(iRun) = 0 Or aYear(iRun) = ABW_URLAUB Or aYear(iRun) = ABW_FAKO Then   ' da zählt Sa/So nicht, bei Dienstreise, Krankheit, Sonstiges schon
                iRun = iRun + 2
                If iRun > iMaxRun Then Exit For
            '''End If
        End If
CheckDay:
        'Debug.Print "Add:" & DateAdd("d", iRun, CDate("31.12.2012"))
        If aYear(iRun) = 0 Then ' Arbeitstag
            If iRun <= iDOY Then        ' Bereits vergangen inklusive Heute
                aAbw(7) = aAbw(7) + 1
            Else                        ' Die Zukunft
                aAbw(8) = aAbw(8) + 1
            End If
        ElseIf aYear(iRun) < 100 Then
            aAbw(aYear(iRun)) = aAbw(aYear(iRun)) + 1
        Else
            aAbw(aYear(iRun) - 100) = aAbw(aYear(iRun) - 100) + 0.5
            If iRun <= iDOY Then        ' Bereits vergangen inklusive Heute
                aAbw(7) = aAbw(7) + 0.5
            Else                        ' Die Zukunft
                aAbw(8) = aAbw(8) + 0.5
            End If
        End If
    Next
End Sub


Public Function Fehltage(cAb As clsAbw) As String
' called by clsAbwesenheit.Anyway
Dim strSQL As String

    Fehltage = ""

    Dim strFirstDay As String:
    strFirstDay = "#" & Format(cAb.dtmStart, "yyyy\/mm\/dd") & "#"
    Dim strLastDay As String:
    strLastDay = "#" & Format(cAb.dtmEnde, "yyyy\/mm\/dd") & "#"
    Dim ArrayLength As Integer
    ArrayLength = cAb.dtmEnde - cAb.dtmStart + 1
    
    ' Basistabelle aYear mit Feiertagen füllen
    Dim i As Integer
    
    Dim aTime() As Long, aFeiertage() As Long
    ReDim aTime(ArrayLength)
    ReDim aFeiertage(ArrayLength)
    
    Dim cF As clsFeiertag
    For Each cF In colFeiertage
        If cF.dtmFeiertag >= cAb.dtmStart And cF.dtmFeiertag <= cAb.dtmEnde Then
            i = DateDiff("d", cAb.dtmStart, cF.dtmFeiertag) + 1
            aFeiertage(i) = -1
        End If
    Next
    
    Dim cA As clsAbw, cUs As clsUser, cFu As clsFunction, ListUser As String
    
    Dim TextTage As String
    For Each cFu In cAb.cUs.colFunctions ' Die Funktionen des Benutzers
        For i = 1 To ArrayLength
            aTime(i) = aFeiertage(i)    ' Reset
        Next
        
        ' Liste der Benutzer für diese Funktion zusammenstellen
        ListUser = ""
        For Each cUs In cFu.colUser
            If ListUser <> "" Then ListUser = ListUser & ","
            ListUser = ListUser & cUs.lngIdxUser
        Next
        
        ' Abwesenheiten innerhalb des Zeitraums der FunktionsBenutzer
        strSQL = "SELECT * FROM tblAbwesenheit WHERE "
        strSQL = strSQL & "(   dtmStart >= " & strFirstDay & " AND dtmStart <= " & strLastDay  ' Start innerhalb oder ...
        strSQL = strSQL & " OR dtmEnde  >= " & strFirstDay & " AND dtmEnde  <= " & strLastDay   ' ... Ende innerhalb oder ...
        strSQL = strSQL & " OR dtmStart <= " & strFirstDay & " AND dtmEnde  >= " & strLastDay  ' ... drüber
        strSQL = strSQL & ") "
        
        strSQL = strSQL & " AND lngIdxUser IN (" & ListUser & ")"
        strSQL = strSQL & " AND lngIdxStatus <> " & AbwStatus.ABGELEHNT & " "   ' Keine gelöschten
        strSQL = strSQL & " AND lngIdxStatus <> " & AbwStatus.ZURUCKGEZOGEN & " "   ' Keine zurückgezogenen
        strSQL = strSQL & "ORDER BY dtmStart"

        Dim RS As ADODB.Recordset
        Set RS = OpenRecordset(strSQL, adOpenKeyset)
        If RS Is Nothing Then Exit Function
        If RS.RecordCount > 0 Then
            While Not RS.EOF
                If RS!lngIdxAbwesenheit <> cAb.lngIdxAbwesenheit Then
                    Dim iStart As Integer, iEnd As Integer, iRun As Integer
                    iStart = RS!dtmStart - cAb.dtmStart + 1
                    iEnd = RS!dtmEnde - cAb.dtmStart + 1
                    
                    If iStart < 1 Then iStart = 1   ' Vor Abwesenheitsbeginn
                    If iEnd > ArrayLength Then iEnd = ArrayLength   ' nach Abwesenheitsende
                    
                    For iRun = iStart To iEnd
                        If aTime(iRun) <> -1 Then aTime(iRun) = aTime(iRun) + 1
                    Next
                End If
                
                RS.MoveNext
            Wend
        
        
            TextTage = ""
            Dim dtmAway As Date
            Dim dtmAwayPrev As Date
            Dim dtmAwayStart As Date, dtmAwayEnd As Date
            dtmAwayStart = 0
            For iRun = 1 To ArrayLength
                If aTime(iRun) = cFu.colUser.Count - 1 Then ' Alle weg
                    dtmAway = cAb.dtmStart + iRun - 1
                    If Weekday(dtmAway) <> g_db.WeekendFirstDay And Weekday(dtmAway) <> g_db.WeekendSecondDay Then
                        If dtmAwayStart = 0 Then    ' neuer Abschnitt
                            dtmAwayStart = dtmAway
                            dtmAwayEnd = dtmAway
                        ElseIf dtmAway = dtmAwayEnd + 1 Then    ' Anschluss an letzten Abschnitt
                            dtmAwayEnd = dtmAway
                        Else                                ' Ausgeben und zurücksetzen
                            If TextTage <> "" Then TextTage = TextTage & " / "
                            If dtmAwayStart = dtmAwayEnd Then
                                TextTage = TextTage & Format(dtmAwayStart, "DD.MM")
                            Else
                                TextTage = TextTage & Format(dtmAwayStart, "DD.MM") & "-" & Format(dtmAwayEnd, "DD.MM")
                            End If
                            dtmAwayStart = dtmAway  ' neuer Abschnitt
                            dtmAwayEnd = dtmAway
                        End If
                    End If
                End If
            Next
            If dtmAwayStart <> 0 Then
                If TextTage <> "" Then TextTage = TextTage & " / "
                If dtmAwayStart = dtmAwayEnd Then
                    TextTage = TextTage & Format(dtmAwayStart, "DD.MM")
                Else
                    TextTage = TextTage & Format(dtmAwayStart, "DD.MM") & "-" & Format(dtmAwayEnd, "DD.MM")
                End If
            End If
            
            ' String befüllen
            
            If TextTage <> "" Then
                Fehltage = Fehltage & "Function group: " & cFu.strFunction & vbTab & TextTage & vbCrLf
            End If
        End If
    Next
    
ExitCalc:
    CloseRecordset RS
    Exit Function
ErrCalc:
    MsgBox "DB ERROR Fehltage: " & Err.Description
    Resume ExitCalc
End Function



'#######################################################################################################
'###########   Items  #############################################################################

Public Function GetItem(strItem As String, Optional varDefault As Variant, Optional strDescription As String) As Variant
Dim cI As clsItem
    GetItem = False
    
    If Now() - m_dtmItemUpdate > TimeSerial(0, 3, 0) Then
        FillItemCollection
        FillOrgItemCollection
        UpdateItemCollectionForCU   ' UserItems
    End If
    On Error Resume Next
    Set cI = Nothing
    Set cI = colItem.Item("cI_" & strItem)
    On Error GoTo err_GetItem
    If cI Is Nothing Then   ' Item selbst nicht gefunden -> anlegen mit Default Value
        If Not IsMissing(varDefault) Then
            GetItem = varDefault    ' das wird das Ergebnis
            Dim RS As ADODB.Recordset
            Set RS = OpenRecordset("tblItems", adOpenDynamic)
            If RS Is Nothing Then Exit Function
            RS.AddNew
            RS!strItem = strItem
            RS!valItem = varDefault
            RS!strDescription = strDescription
            RS.Update
            CloseRecordset RS
        End If
    Else    ' Item gefunden
        GetItem = cI.valItem
        If strDescription <> "" Then
            If strDescription <> cI.strDescription Then
                Execute ("UPDATE tblItems SET strDescription=""" & strDescription & """ WHERE strItem=""" & strItem & """")
            End If
        End If
    End If

exit_GetItem:
    Exit Function
err_GetItem:
    m_strText = "DB ERROR GetItem err=" & Err.Number & ":" & Err.Description
    Logging "DB ERROR", , m_strText
    GetItem = False
    Resume exit_GetItem

End Function
Private Function GetItemValForUser(cU As clsUser, cI As clsItem) As String
    GetItemValForUser = cI.ValDefault  ' Default
    If cI.colValues.Count = 0 Then Exit Function
    ' kleinste mögliche Org in cI.colValues suchen
    If GetOrgItem(cU.lngIdxOrg, cI, GetItemValForUser) Then Exit Function
    If GetOrgItem(cU.lngIdxInsel, cI, GetItemValForUser) Then Exit Function
    If GetOrgItem(cU.lngIdxTeam, cI, GetItemValForUser) Then Exit Function
    If GetOrgItem(cU.lngIdxAbt, cI, GetItemValForUser) Then Exit Function
    If GetOrgItem(cU.lngIdxBer, cI, GetItemValForUser) Then Exit Function
    If GetOrgItem(cU.lngIdxRes, cI, GetItemValForUser) Then Exit Function
End Function
Private Function GetOrgItem(lngIdxOrg As Long, cI As clsItem, ByRef value As String) As Boolean
    On Error Resume Next
    Dim cOI As clsOrgItem
    Set cOI = cI.colValues("cOI_" & lngIdxOrg)
    If cOI Is Nothing Then
        GetOrgItem = False
    Else
        value = cOI.valItem
        GetOrgItem = True
    End If
End Function
Public Function WriteOrgItem(lngIdxOrg As Long, strItem As String, valItem As String) As Boolean
    WriteOrgItem = False
    On Error GoTo err_WriteOrgItem
    
    ' ToDo: Ist dieses ohnehin der Default-Value? Dann Eintrag löschen
    Dim cI As clsItem
    Set cI = colItem("cI_" & strItem)   ' MUSS existieren
    Dim RS As ADODB.Recordset
    Set RS = OpenRecordset("tbxOrgItems", adOpenDynamic)
    If RS Is Nothing Then Exit Function
    RS.AddNew
    RS!lngIdxOrg = lngIdxOrg
    RS!strItem = strItem
    RS!valItem = valItem
    RS.Update
    
    WriteOrgItem = True
exit_WriteOrgItem:
    CloseRecordset RS
    FillItemCollection  ' komplett neu aufbauen
    FillOrgItemCollection
    UpdateItemCollectionForCU
    Exit Function
err_WriteOrgItem:
    m_strText = "DB ERROR WriteOrgItem err=" & Err.Number & ":" & Err.Description
    Logging "DB ERROR", , m_strText
    Resume exit_WriteOrgItem
End Function
Private Sub FillOrgItemCollection()
Dim cOI As clsOrgItem, rsV As ADODB.Recordset
    On Error GoTo err_FillOrgItemCollection
    Set colOrgItem = New Collection
    If RSSchemaFields("tbxOrgItems").BOF Then Exit Sub
    
    Dim RS As ADODB.Recordset
    Set RS = OpenRecordset("tbxOrgItems", adOpenKeyset)
    If RS Is Nothing Then Exit Sub
    If RS.RecordCount > 0 Then
        While Not RS.EOF
            Set cOI = New clsOrgItem
            cOI.lngIdxOrg = RS!lngIdxOrg
            cOI.valItem = RS!valItem
            
            ' Element an colItem.colValues anhängen
            ' Item finden
            Dim cI As clsItem
            On Error Resume Next
            Set cI = colItem("cI_" & RS!strItem)
            On Error GoTo err_FillOrgItemCollection
            If Not cI Is Nothing Then
                cI.colValues.Add cOI, "cOI_" & cOI.lngIdxOrg
            End If
            
            RS.MoveNext
        Wend
    End If
exit_FillOrgItemCollection:
    CloseRecordset RS
    Exit Sub
err_FillOrgItemCollection:
    m_strText = "DB ERROR FillOrgItemCollection err=" & Err.Number & ":" & Err.Description
    Logging "DB ERROR", , m_strText
    MsgBox m_strText
    Resume exit_FillOrgItemCollection
End Sub
Public Sub UpdateItemCollectionForCU()
    If g_CU_Login Is Nothing Then Exit Sub
    ' Nach FillItem + FillOrg + FillOrgItem -> Wert userspezifisch aufbereiten
    Dim cI As clsItem
    For Each cI In colItem
        ' a) org-spezifischen Wert ermitteln
        cI.valItem = GetItemValForUser(g_CU_Login, cI)
        ' b) user-spezifisch ?
        If IsUserItem(cI.strItem) Then  ' schau in der Liste nach, ob dieses Item Benutzeränderbar ist
            ReadUserOption cI
        End If
    Next
    InitWeekend ' how weekend is specified
End Sub
Private Sub FillItemCollection()
Dim cI As clsItem
    On Error GoTo err_FillItemCollection
    Set colItem = New Collection
    
    If RSSchemaFields("tblItems").BOF Then Exit Sub
    
    InitUserItems   ' Namen definieren
    Dim RS As ADODB.Recordset
    Set RS = OpenRecordset("tblItems", adOpenKeyset)
    If RS Is Nothing Then Exit Sub
    If RS.RecordCount > 0 Then
        While Not RS.EOF
            Set cI = New clsItem
            cI.strItem = RS!strItem
            cI.valItem = IIf(Not IsNull(RS!valItem), RS!valItem, "")
            cI.ValDefault = cI.valItem ' valItem kann überschrieben werden
            cI.strDescription = IIf(Not IsNull(RS!strDescription), RS!strDescription, "")
            Set cI.colValues = New Collection
            On Error Resume Next
            colItem.Add cI, "cI_" & RS!strItem
            On Error GoTo err_FillItemCollection
            RS.MoveNext
        Wend
    End If
exit_FillItemCollection:
    On Error Resume Next
    m_dtmItemUpdate = Now()
    CloseRecordset RS
    Exit Sub
err_FillItemCollection:
    m_strText = "DB ERROR FillItemCollection err=" & Err.Number & ":" & Err.Description
    Logging "DB ERROR", , m_strText
    MsgBox m_strText
    Resume exit_FillItemCollection
End Sub
Public Function GetOrgValue(lngIdxOrg As Long, strItem As String) As String
    GetOrgValue = ""
    On Error Resume Next
    Dim cI As clsItem
    Set cI = colItem("cI_" & strItem)
    If Not cI Is Nothing Then   ' Item gefunden ...
        GetOrgValue = cI.valItem
        Dim cOI As clsOrgItem
        Set cOI = cI.colValues("cOI_" & lngIdxOrg)
        If Not cOI Is Nothing Then
            GetOrgValue = cOI.valItem
        End If
    End If
End Function
Private Sub InitUserItems()
    Set colUserItems = New Collection
    colUserItems.Add "booAlleMASichtbar", "booAlleMASichtbar"
    colUserItems.Add "booOrgSlider", "booOrgSlider"
    colUserItems.Add "booPlanEveryDay", "booPlanEveryDay"
    colUserItems.Add "booShowSort", "booShowSort"
    colUserItems.Add "lngBreiteSort", "lngBreiteSort"
    colUserItems.Add "ShowUserList", "ShowUserList"
    colUserItems.Add "WidthUserColumn", "WidthUserColumn"
End Sub
Private Function IsUserItem(name As String) As Boolean
    Dim s As Variant
    On Error Resume Next
    s = colUserItems(name)
    If Err = 0 Then
        IsUserItem = True
    Else
        IsUserItem = False
    End If
End Function
Public Sub ReadUserOption(cI As clsItem)
    ' called FillItemCollection
    Dim var As Variant
    var = basFormUtils.GetRegData(cI.strItem)   ' Hat der Benutzer etwas gespeichert?
    If var <> -1 Then                           ' UserOption gefunden
        cI.valItem = var                        ' ÜBERSCHREIBEN
    End If
End Sub
Public Static Sub WriteUserOption(cI As clsItem)
    On Error Resume Next
    Dim ci1 As clsItem
    Set ci1 = Nothing
    Set ci1 = colItem.Item("cI_" & cI.strItem)  ' Aktueller Wert (aus DB + Registry)
    
    Dim var As Variant
    var = basFormUtils.GetRegData(cI.strItem)   ' aktuell in Registry geschrieben
    
    If var = -1 Then                ' UserOption nicht gefunden
        basFormUtils.SetRegData cI.strItem, cI.valItem
    Else                ' ... gefunden
        If Not ci1 Is Nothing Then  ' Standard gefunden
            If var <> ci1.valItem Then  ' Ist nicht Standard
                If var <> cI.valItem Then   ' Nicht schon geschrieben
                    basFormUtils.SetRegData cI.strItem, cI.valItem
                End If
            Else            ' identisch zu general -> Löschen
                basFormUtils.CleanRegData (cI.strItem)
            End If
        End If
    End If
    
    If Not ci1 Is Nothing Then
        ci1.valItem = cI.valItem
    End If

End Sub
