VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAbwesenheit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents m_oApp As Outlook.Application
Attribute m_oApp.VB_VarHelpID = -1
Dim booAdvancedSearchCompleteFired As Boolean, m_StrTextBegruendung As String

Public Sub Planen(cAb As clsAbw)

    If Not Anyway(cAb, "Plan") Then Exit Sub    ' Funktion nicht besetzt?
    
    If cAb.lngIdxAbwesenheit > 0 Then   ' Bestehende Abwesenheit
        Update cAb, cAb.lngIdxStatus    ' Keine Statusänderung
    Else                                ' Neue Abwesenheit
        cAb.lngIdxStatus = AbwStatus.UNDEFINED   ' Abwesenheit hat noch keinen Status - Relevant für Statuswechsel in clsData
        Update cAb, AbwStatus.PLANUNG   ' Neuer Status: Planung
    End If
    AddOutlookAbwesenheit cAb   ' eventuell Outlook Appointment ändern
End Sub

Public Sub Beantragen(lngIdxAbwesenheit As Long)
    Dim cAb As clsAbw
    Set cAb = g_db.GetAbwesenheit(lngIdxAbwesenheit)
    
    If Not Anyway(cAb, "Apply for") Then Exit Sub
    
    Update cAb, AbwStatus.BEANTRAGT_1
End Sub

Public Sub Freigeben(lngIdxAbwesenheit As Long)  ' btnGenehmigen in frmManager
    Dim cAb As clsAbw
    Set cAb = g_db.GetAbwesenheit(lngIdxAbwesenheit)
    
    If Not Anyway(cAb, "Approve") Then Exit Sub
    
    Dim lngIdxNewStatus As Long
    
    If cAb.dtmGenehmigt1 = "00:00:00" Then  ' 1. Genehmigung
        cAb.dtmGenehmigt1 = Now()
        ' Wenn die Level1Tage überschritten sind und NICHT AL und NICHT SEK -> Setzen in Level2-beantragt
        If cAb.lngAnzahl > g_db.GetItem("lngAnzTageLevel1", 99, "If 2 level approval: number of days level 1") And g_CU.lngUserLevel > basGlobals.OrgLevel.Abteilung And Not g_CU.IsSekOf(cAb.cUs) Then
            lngIdxNewStatus = AbwStatus.BEANTRAGT_2

        Else    ' Level1Tage unterschritten oder >= AL oder SEK -> Genehmigung
            lngIdxNewStatus = AbwStatus.GENEHMIGT
        End If
    Else        ' Genehmigung #2
        cAb.dtmGenehmigt2 = Now()
        lngIdxNewStatus = AbwStatus.GENEHMIGT
    End If
    
    Update cAb, lngIdxNewStatus
End Sub
Private Function Anyway(cAb As clsAbw, Action As String) As Boolean
    Anyway = True
    
    Dim Text As String
    Text = g_db.Fehltage(cAb)
    If Text <> "" Then
        If MsgBox("This absence will result in not fulfilling function(s):" & vbCrLf & vbCrLf & Text & vbCrLf & vbCrLf & Action & " anyway?", vbYesNo) = vbNo Then
            Anyway = False
        End If
    End If
End Function
Public Sub Zurueckziehen(lngIdxAbwesenheit As Long)
Dim cAb As clsAbw
    On Error GoTo err_Zurueckziehen
    
    Set cAb = g_db.GetAbwesenheit(lngIdxAbwesenheit)
    If cAb Is Nothing Then
        MsgBox "This absence doesn't exist anymore ..."
        Exit Sub
    End If
    
'    If False And cAb.dtmEnde < Date And cAb.lngIdxStatus > AbwStatus.PLANUNG Then  ' Version 2.2.21 - Vergangenheit geht auch...
'        MsgBox g_db.GetString(1020):     Exit Sub    ' Vergangenheit
'    End If
    
    Dim strText As String, varval As Object
    strText = g_clsAb.Compose(cAb) & vbCrLf & g_clsAb.StatusText(lngIdxAbwesenheit, True) & vbCrLf    ' Art+Zeitraum, bekannte Daten
    If MsgBox(strText & vbCrLf & g_db.GetString(1050), vbQuestion + vbOKCancel + vbDefaultButton2) <> vbOK Then Exit Sub ' Frage an Benutzer: Wirklich?

    ' ... Abwesenheit löschen ...
    DeleteOutlookAbwesenheit cAb            ' als erstes versuchen, einen eventuellen Outlook-Eintrag zu löschen
    Update cAb, AbwStatus.ZURUCKGEZOGEN
    Exit Sub
err_Zurueckziehen:
    MsgBox "clsAbwesenheit:Zurueckziehen" & Err.Description
End Sub
Public Sub Ablehnen(lngIdxAbwesenheit As Long, strTextBegruendung As String)
    On Error GoTo err_Ablehnen
    Dim cAb As clsAbw
    Set cAb = g_db.GetAbwesenheit(lngIdxAbwesenheit)
    
    m_StrTextBegruendung = strTextBegruendung
    
    DeleteOutlookAbwesenheit cAb    ' als erstes versuchen, einen eventuellen Outlook-Eintrag zu löschen

    Update cAb, AbwStatus.ABGELEHNT
    Exit Sub
err_Ablehnen:
    MsgBox "clsAbwesenheit:Ablehnen" & Err.Description
End Sub

'##########################################################################################################
Private Sub Update(cAb As clsAbw, IdxStatus As Long)
    If Not cAb Is Nothing Then
        g_db.UpdateAbwesenheit cAb, IdxStatus
        SendMail cAb
    End If
End Sub
' Generiere Texte für die Mail
Public Sub SendMail(cAb As clsAbw) ' letzte Aktion in clsData.UpdateAbwesenheit
    If cAb.lngIdxStatus = cAb.lngIdxStatusOld Then Exit Sub     ' Keine Status-Änderung, nur Inhalts-Änderung
    
    Dim strAbwText As String, strUserInfo As String, strMessage As String, strBetreff As String
    Dim booCC As Boolean    ' Sekretariate auch informieren?
    booCC = False           ' Standard: Keine Info an Sek
    strAbwText = g_clsAb.Compose(cAb)   ' <Abwesenheit> vom <...> bis zum <...>       - ohne Ihr/Ihre
    
    strUserInfo = cAb.UserInfo       ' <Nachname>, <Vorname> -

    Select Case cAb.lngIdxStatus
        Case AbwStatus.PLANUNG
            Exit Sub    ' noch keine Mail
        Case AbwStatus.BEANTRAGT_1, AbwStatus.BEANTRAGT_2
            ' Mail an Vorgesetzten - immer
            strBetreff = strUserInfo & g_db.GetString(1049) & " " & strAbwText ' <Nachname>, <Vorname> - Antrag <Abwesenheit> vom <...> bis zum <...>
            strMessage = strUserInfo & g_db.GetString(1024) & ": " & strAbwText              ' Beantragt: <Abwesenheit> vom <...> bis zum <...>
            Dim IdxChef As Long
            IdxChef = g_CU.lngIdxChef   ' Default
            If CBool(g_db.GetItem("TravellingRequestToE3", "False", "Travelling request to E3 instead of E4")) Then
                If g_CU.cOrg.lngOrgLevel = Team Then ' gilt nur für MA, alle andern halten den Weg ein
                    IdxChef = g_db.GetOrgChef(g_CU.lngIdxAbt)   ' E3
                End If
            End If
            
            ' Mail an den "original-Chef"
            If Not AutoMail(IdxChef, strBetreff, strMessage, cAb.lngIdxAbwesenheit, False) Then MsgBox g_db.GetString(1048), vbInformation
            
            ' Ist der Vorgesetzte gerade abwesend und hat einen Vertreter, dann Mail auch an diesen
            If g_db.HatVertreter(IdxChef) And g_db.IstAbwesend(IdxChef) Then
                IdxChef = g_db.GetVertreter(IdxChef).lngIdxUser
                If Not AutoMail(IdxChef, strBetreff, strMessage, cAb.lngIdxAbwesenheit, False) Then MsgBox g_db.GetString(1048), vbInformation
            End If
            
            If cAb.lngIdxStatus = AbwStatus.BEANTRAGT_2 Then    ' Mail vom Genehmiger_1 auch an Benutzer, wenn 2-stufig
                ' Mail an Benutzer und Mail an Vorgesetzten, jeweils ohne CC    1025:Genehmigt 1049:Antrag
                strBetreff = strUserInfo & g_db.GetString(1025) & " Level1 " & g_db.GetString(1152) & ":" & strAbwText   ' Genehmigt Level1 und wird weitergeleitet. <Abwesenheit...
                strMessage = strAbwText & " " & g_db.GetString(1153) & "."  ' ... wurde genehmigt und an die naechste...
                If Not AutoMail(cAb.lngIdxUser, strBetreff, strMessage, cAb.lngIdxAbwesenheit, False) Then MsgBox g_db.GetString(1084), vbInformation
            End If
        
            MailToFunctionColleagues strBetreff, strMessage, cAb
        Case AbwStatus.GENEHMIGT
            ' Mail an Benutzer mit CC
            strBetreff = strUserInfo & g_db.GetString(1025) & " " & strAbwText ' <Nachname>, <Vorname> - Genehmigt <Abwesenheit> vom <...> bis zum <...>
            strMessage = strUserInfo & g_db.GetString(1025) & ": " & strAbwText              ' Beantragt: <Abwesenheit> vom <...> bis zum <...>
            If Not AutoMail(cAb.lngIdxUser, strBetreff, strMessage, cAb.lngIdxAbwesenheit, True) Then MsgBox g_db.GetString(1084), vbInformation

            MailToFunctionColleagues strBetreff, strMessage, cAb
        Case AbwStatus.ABGELEHNT            'TODO: die ganze Strecke informieren, nicht nur den direkten Vorgesetzten
            strBetreff = strUserInfo & g_db.GetString(1082) & " " & strAbwText ' <Nachname>, <Vorname> - abgelehnt <Abwesenheit> vom <...> bis zum <...>
            
            strMessage = strAbwText & " - " & GetMessageForOldStat(cAb.lngIdxStatusOld) & " " & g_db.GetString(1083)     ' wurde abgelehnt
            strMessage = strMessage & vbCrLf & "Reason:" & m_StrTextBegruendung   ' wurde abgelehnt

            If cAb.lngIdxStatusOld = AbwStatus.GENEHMIGT Then booCC = True
            If Not AutoMail(cAb.lngIdxUser, strBetreff, strMessage, cAb.lngIdxAbwesenheit, booCC) Then MsgBox g_db.GetString(1084), vbInformation
            
            MailToFunctionColleagues strBetreff, strMessage, cAb
        Case AbwStatus.ZURUCKGEZOGEN            'TODO: die ganze Strecke informieren, nicht nur den direkten Vorgesetzten
            If cAb.lngIdxStatusOld = AbwStatus.PLANUNG Then Exit Sub                ' vor Antrag keine Mail
            If cAb.lngIdxStatusOld = AbwStatus.GENEHMIGT Then booCC = True
            
            strBetreff = strUserInfo & g_db.GetString(1054) & " " & strAbwText ' <Nachname>, <Vorname> - zurückgezogen <Abwesenheit> vom <...> bis zum <...>
            
            strMessage = strAbwText & " - " & GetMessageForOldStat(cAb.lngIdxStatusOld) & " " & g_db.GetString(1054)     ' zurückgezogen

            If Not AutoMail(g_CU.lngIdxChef, strBetreff, strMessage, cAb.lngIdxAbwesenheit, booCC) Then MsgBox g_db.GetString(1048), vbInformation
            
            MailToFunctionColleagues strBetreff, strMessage, cAb
    End Select
End Sub
Private Function MailToFunctionColleagues(strBetreff As String, strMessage As String, cAb As clsAbw) As Boolean
    If cAb.cUs.colFunctions.Count = 0 Then Exit Function
    
    Dim cFu As clsFunction, cUs As clsUser   ' Laufvariablen
    Dim MessagetoSend As String
    For Each cFu In cAb.cUs.colFunctions
        MessagetoSend = strMessage & vbCrLf & vbCrLf & "This mail was sent to you because you are member of function group " & cFu.strFunction & " together with " & g_CU.Fullname
        For Each cUs In cFu.colUser
            ' nicht an Vorgesetzten senden, das geschieht ohnehin. An sich selbst ebenfalls nicht
            If g_CU_Login.lngIdxUser <> cUs.lngIdxChef And g_CU_Login.lngIdxUser <> cUs.lngIdxUser Then
                If Not AutoMail(cUs.lngIdxUser, strBetreff, MessagetoSend, cAb.lngIdxAbwesenheit, False) Then MsgBox g_db.GetString(1048), vbInformation
            End If
        Next
    Next

End Function
Private Function GetMessageForOldStat(lngIdxOldStat As Long)
    If lngIdxOldStat = AbwStatus.UNDEFINED Then GetMessageForOldStat = "undefined"
    If lngIdxOldStat = AbwStatus.BEANTRAGT_1 Then GetMessageForOldStat = g_db.GetString(1051)    ' beantragte
    If lngIdxOldStat = AbwStatus.BEANTRAGT_2 Then GetMessageForOldStat = g_db.GetString(1052)     ' beantragte teilgenehmigte
    If lngIdxOldStat = AbwStatus.GENEHMIGT Then GetMessageForOldStat = g_db.GetString(1053)     ' genehmigte Abwesenheit
End Function


'##########################################################################################################
'##########################################################################################################
'##########################################################################################################
Private Sub AddOutlookAbwesenheit(cAb As clsAbw)    ' called by "Planen"
    If g_CU.lngIdxUser <> g_CU_Login.lngIdxUser Then Exit Sub  ' Nur Outlook-Eintrag, wenn Username und Loginname übereinstimmen
    If g_CU_Login.lngIdxUser <> cAb.lngIdxUser Then Exit Sub    ' Keine Outlook-Abwesenheit anlegen, wenn nicht für mich
    
    Dim lngIdxAbwesenheitsArt As Long, strStatus As String, strErr As String, numErr As Integer

    strStatus = "OutlookStart"
    Set m_oApp = Nothing
    On Error Resume Next
    strStatus = "Outlook New Outlook.Application"
    Set m_oApp = New Outlook.Application
    If m_oApp Is Nothing Or Err.Number <> 0 Then GoTo err_Outlook
    
    Dim oAppt As Outlook.AppointmentItem
    
    ' bei Änderungen gibt es bereits einen bestehenden Eintrag
    If cAb.strOutlookText <> "" Then
        SearchOutlookAppointment cAb.strOutlookText, oAppt  ' gibt es wirklich ?
    End If
    
    If oAppt Is Nothing Then
        strStatus = "Outlook New Outlook.Appointment"
        Set oAppt = m_oApp.CreateItem(olAppointmentItem)
        If oAppt Is Nothing Or Err.Number <> 0 Then GoTo err_Outlook
        
        On Error GoTo err_Outlook
        strStatus = "Outlook New AppStatus"
        oAppt.MeetingStatus = olNonMeeting
    End If
    
    strStatus = "Outlook New App Subject"
    oAppt.Subject = GenSubject(cAb)     ' aus den Daten die Signatur erzeugen

    Dim timStart As Date, timEnd As Date
    Select Case cAb.lngGVN
        Case ABW_GANZTAGS
            timStart = TimeSerial(0, 0, 0)
            timEnd = TimeSerial(24, 0, 0)
        Case ABW_VORMITTAGS
            timStart = TimeSerial(8, 0, 0)
            timEnd = TimeSerial(12, 0, 0)
        Case ABW_NACHMITTAGS
            timStart = TimeSerial(12, 0, 0)
            timEnd = TimeSerial(18, 0, 0)
    End Select
    strStatus = "Outlook New App Start"
    oAppt.Start = cAb.dtmStart + timStart
    oAppt.End = cAb.dtmEnde + timEnd
    oAppt.ReminderSet = True
    oAppt.ReminderMinutesBeforeStart = 12 * 60
    oAppt.BusyStatus = olOutOfOffice
    oAppt.IsOnlineMeeting = False
    oAppt.AllDayEvent = (cAb.lngGVN = ABW_GANZTAGS)
    strStatus = "Outlook New App Save"
    oAppt.Save
exit_Outlook:
    On Error Resume Next
    Set oAppt = Nothing
    Set m_oApp = Nothing
    Exit Sub

err_Outlook:
    numErr = Err.Number
    strErr = Err.Description
    g_db.Logging "Add", cAb, "err_Outlook:" & strStatus & " err:" & numErr & " Desc:" & strErr
    MsgBox g_db.GetString(1099) & "  " & strStatus & vbCrLf & strErr
    Resume exit_Outlook
End Sub
'################################################################################
Private Function SearchOutlookAppointment(SearchText As String, AppItem As Outlook.AppointmentItem) As Boolean
    ' Im Outlook nach diesem Eintrag suchen

    Dim strStatus As String
    On Error GoTo err_SearchOutlookAppointment
    
    Const strScope As String = "Calendar"
    
    Dim strFilter As String
    strFilter = "urn:schemas:httpmail:subject = '" & SearchText & "'"

    ' Suche vorbereiten
    Dim booSearchComplete As Boolean:     booSearchComplete = False
    booAdvancedSearchCompleteFired = False  ' wird in m_oApp_AdvancedSearchComplete gesetzt
    
    Set m_oApp = New Outlook.Application
    Dim oSrch As Outlook.Search
    Set oSrch = m_oApp.AdvancedSearch(strScope, strFilter)
    
    Dim iMaxCounter As Integer:    iMaxCounter = 15     ' 15*200ms= 3 Sekunden
    
'    MousePointer = vbHourglass
    Dim iCounter As Integer:    iCounter = 0
    While booSearchComplete = False     ' Warten auf Event
        DoEvents
        If booAdvancedSearchCompleteFired Then
            booSearchComplete = True
        Else
            iCounter = iCounter + 1
            If iCounter > iMaxCounter Then booSearchComplete = True      ' Raus ...
            Sleep 200
        End If
    Wend
'    frmInfo.MousePointer = vbNormal
    
    Debug.Print "Search:" & strFilter & " Result#=" & oSrch.Results.Count
    
    strStatus = "Absence was found in outlook. Number of hits:" & oSrch.Results.Count
    
    If oSrch.Results.Count = 0 Then
        strStatus = strStatus & vbCrLf & "Absence with text " & SearchText & " not found in outlook."
        MsgBox strStatus
        SearchOutlookAppointment = False
    ElseIf oSrch.Results.Count = 1 Then
        On Error Resume Next
        Set AppItem = oSrch.Results.Item(1)
        SearchOutlookAppointment = True
    Else
        strStatus = strStatus & vbCrLf & "Multiple absences with text " & SearchText & " were found in outlook. Nothing changed."
        MsgBox strStatus
        SearchOutlookAppointment = False
    End If
    Exit Function
    
err_SearchOutlookAppointment:
    MsgBox "Error searching absence in outlook:" & strStatus & " Err:" & Err.Description
    SearchOutlookAppointment = False
End Function
Private Sub m_oApp_AdvancedSearchComplete(ByVal SearchObject As Search)
'    MsgBox "The AdvancedSearchComplete Event fired"
    booAdvancedSearchCompleteFired = True
End Sub
'################################################################################
Private Sub DeleteOutlookAbwesenheit(cAb As clsAbw)
    If g_CU.lngIdxUser <> g_CU_Login.lngIdxUser Then Exit Sub  ' Nur Outlook-Eintrag, wenn Username und Loginname übereinstimmen
    If g_CU_Login.lngIdxUser <> cAb.lngIdxUser Then Exit Sub    ' Bei Ablehnung - keine Outlook-Abwesenheit löschen

    Dim strStatus As String
    Dim AppItem As Outlook.AppointmentItem
    If SearchOutlookAppointment(GenSubject(cAb), AppItem) Then    ' was gefunden
        On Error Resume Next
        AppItem.DELETE
        If Err.Number > 0 Then
            strStatus = "Absence couldn't be deleted from outlook calendar. " & Err.Description
            MsgBox strStatus
        Else
            strStatus = "Absence was deleted from outlook calendar succsessfully."
        End If
    Else
        strStatus = strStatus & vbCrLf & "Absences with text " & GenSubject(cAb) & " not found in outlook. Nothing deleted."
        MsgBox strStatus
    End If
end_Delete_Outlook:
    On Error Resume Next
    Set m_oApp = Nothing  ' Outlook Application
    Exit Sub
End Sub

'------------------------------------------------------------------------------------------------------------------------------------------------
Public Function GenSubject(cAb As clsAbw) As String
    Dim strSubject As String            ' mit diesem Betreff wurde eingetragen
    If cAb.strAbwesenheitsart = "" Then
        g_db.Logging "GenSubject", cAb, "strAbw ist leer. IdxAbwArt=" & cAb.lngIdxAbwesenheitsArt
        If cAb.lngIdxAbwesenheitsArt = 0 Then
            g_db.Logging "GenSubject", cAb, "lngIdxAbw ist 0"
        Else
            cAb.strAbwesenheitsart = g_db.GetAbwesenheitsart(cAb.lngIdxAbwesenheitsArt).strAbwesenheitsart
            If cAb.strAbwesenheitsart = "" Then
                g_db.Logging "GenSubject", cAb, "Abwesenheitsart konnte nicht ermittelt werden."
            Else
                g_db.Logging "GenSubject", cAb, "Abwesenheitsart gefunden."
            End If
        End If
    End If
    strSubject = App.ExeName & " " & cAb.strAbwesenheitsart & " - " & cAb.dtmStart & IIf(cAb.dtmStart = cAb.dtmEnde, "", "-" & cAb.dtmEnde)
    If cAb.strText <> "" Then strSubject = strSubject & "  " & cAb.strText
    If CBool(g_db.GetItem("booOutlookPlan", "False", "Enter Oulook Appointment as planned.(default=false)?")) Then
        strSubject = strSubject & " - " & g_db.GetString(1023)  ' Geplant
    End If
    
    GenSubject = strSubject
End Function

Public Function GenAbwText(cAb As clsAbw) As String
    GenAbwText = IIf(cAb.lngIdxAbwesenheit = ABW_DIENSTREISE, g_db.GetString(1151), g_db.GetString(1088)) & " "
    GenAbwText = GenAbwText & Compose(cAb)
End Function
Public Function Compose(cAb As clsAbw) As String    ' Abwesenheitsart von xxx bis xxx vormittags
    Dim strText As String
    strText = g_db.GetAbwesenheitsart(cAb.lngIdxAbwesenheitsArt).strAbwesenheitsart & " "
    
    If cAb.dtmStart = cAb.dtmEnde Then
        strText = strText & g_db.GetString(1156) & " " & cAb.dtmStart  ' vom
    Else
        strText = strText & g_db.GetString(1046) & " " & cAb.dtmStart & " " & g_db.GetString(1047) & " " & cAb.dtmEnde     ' vom bis
    End If
    
    If cAb.strGVN <> "" Then strText = strText & " " & cAb.strGVN
    If cAb.strText <> "" Then strText = strText & "  <" & cAb.strText & ">"
    
    Compose = strText
End Function

Public Function CheckDates(dtmStart As Date, dtmEnde As Date) As Boolean
' Gibt es Überschneidungen mit bereits eingetragenen eigenen Entries ?
    CheckDates = True
    Dim cAb As clsAbw
    For Each cAb In g_db.colAbwesenheiten
        If cAb.lngIdxUser = g_CU.lngIdxUser Then
            If cAb.lngIdxStatus <> AbwStatus.ZURUCKGEZOGEN And cAb.lngIdxStatus <> AbwStatus.ABGELEHNT Then
                If dtmStart >= cAb.dtmStart And dtmStart <= cAb.dtmEnde Or _
                   dtmEnde >= cAb.dtmStart And dtmEnde <= cAb.dtmEnde Then
                   CheckDates = False   ' Neue Abw. liegt innerhalb bestehender Abw.
                   Exit For ' Nicht weitermachen
                End If
                If cAb.dtmStart >= dtmStart And cAb.dtmStart <= dtmEnde Or _
                   cAb.dtmEnde >= dtmStart And cAb.dtmEnde <= dtmEnde Then
                   CheckDates = False      ' Alte Abw. liegt innerhalb neuer Abw.
                   Exit For  ' Nicht weitermachen
                End If
            End If
        End If
    Next
End Function

Public Function StatusText(lngIdxAbwesenheit As Long, booTimes As Boolean) As String
    Dim cAb As clsAbw, strText As String
    Set cAb = basGlobals.g_db.GetAbwesenheit(lngIdxAbwesenheit)
    If cAb Is Nothing Then
        StatusText = ""
        Exit Function
    End If
    Select Case cAb.lngIdxStatus
        Case AbwStatus.UNDEFINED
            strText = "undefined"
        Case AbwStatus.PLANUNG
            strText = g_db.GetString(1023)               ' Geplant   / planned
        Case AbwStatus.BEANTRAGT_1
            strText = g_db.GetString(1024) & " Level1"   ' Beantragt /
        Case AbwStatus.BEANTRAGT_2
            strText = g_db.GetString(1024) & " Level2"
        Case AbwStatus.GENEHMIGT
            strText = g_db.GetString(1025)               ' Genehmigt /
        Case AbwStatus.ABGELEHNT
            strText = g_db.GetString(1049)               ' Abgelehnt / refused
        Case AbwStatus.ZURUCKGEZOGEN
            strText = g_db.GetString(1048)               ' Zurückgezogen /
    End Select

    If booTimes Then
        If cAb.dtmErstellung > 0 Then strText = strText & "<P:" & cAb.dtmErstellung & ">"
        If cAb.dtmAntrag > 0 Then strText = strText & "<A:" & cAb.dtmAntrag & ">"
        If cAb.dtmGenehmigt1 > 0 Then strText = strText & "<OK1:" & cAb.dtmGenehmigt1 & ">"   ' Genehmigt
        If cAb.dtmGenehmigt2 > 0 Then strText = strText & "<OK2:" & cAb.dtmGenehmigt2 & ">"
    End If
    StatusText = strText
End Function
