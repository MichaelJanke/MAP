Attribute VB_Name = "basFunktionen"
Option Explicit

Private booGueltigesDatum As Boolean, varval As Variant

Public Function EndeGueltig(ByRef calEnde As Date) As Date
    ' Wenn ein Tag angegeben wird, der als Endetag nicht in Frage kommt, dann gehe so weit wie nötig rückwärts
    On Error Resume Next
    If g_booPlanEveryDay Then
        EndeGueltig = calEnde
    Else
        booGueltigesDatum = False
        Do
            If Weekday(calEnde) = g_db.WeekendSecondDay Then calEnde = calEnde - 2                  ' 1=Sonntag
            If Weekday(calEnde) = g_db.WeekendFirstDay Then calEnde = calEnde - 1                ' 7=Samstag
            Dim cF As clsFeiertag:         Set cF = g_db.GetFeiertagByDate(calEnde)
            If cF Is Nothing Then   ' kein Feiertag
                booGueltigesDatum = True
            Else    ' Feiertag
                If cF.lngIdxAbwesenheitsArt = basGlobals.ABW_FEIERTAG Then
                    calEnde = calEnde - 1
                Else
                    booGueltigesDatum = True
                End If
            End If
        Loop Until booGueltigesDatum = True
        EndeGueltig = calEnde
    End If
End Function
Public Function StartGueltig(ByRef calStart As Date) As Date
    ' Wenn ein Tag angegeben wird, der als Starttag nicht in Frage kommt, dann gehe so weit wie nötig vorwärts

    On Error Resume Next
    If g_booPlanEveryDay Then
        StartGueltig = calStart
    Else
        booGueltigesDatum = False
        Do
            If Weekday(calStart) = g_db.WeekendSecondDay Then calStart = calStart + 1              ' 1=Sonntag
            If Weekday(calStart) = g_db.WeekendFirstDay Then calStart = calStart + 2            ' 7=Samstag
            Dim cF As clsFeiertag:         Set cF = g_db.GetFeiertagByDate(calStart)
            If cF Is Nothing Then
                booGueltigesDatum = True
            Else
                If cF.lngIdxAbwesenheitsArt = basGlobals.ABW_FEIERTAG Then
                    calStart = calStart + 1
                Else
                    booGueltigesDatum = True
                End If
            End If
        Loop Until booGueltigesDatum = True
        StartGueltig = calStart
    End If
End Function

Public Function CalculateTage(ByRef cAb As clsAbw, Optional booDetails As Boolean = False, Optional ByRef strFeiertage As String) As String
Dim dtmRun As Date, cF As clsFeiertag
Dim strStatus As String
Dim booFeiertage As Boolean     ' strAbwesenheitstage und strFeiertage auch Füllen
    On Error GoTo err_calcTage
    dtmRun = cAb.dtmStart
    If strFeiertage = ":" Then
        booFeiertage = True
    Else
        booFeiertage = False
    End If
    strFeiertage = ""
    cAb.lngAnzahl = 0           ' Alle Tage
    cAb.lngAnzahlFAKO = 0       ' Nur die FAKO-Tage
    cAb.lngAnzahlUrlaub = 0     ' Nur die Urlaubs-Tage
    While dtmRun <= cAb.dtmEnde ' Gehe alle geplanten Tage durch und überprüfe, ob überhaupt Urlaub/Fako nötig
        strStatus = "   ---   dtmRun=<" & dtmRun & ">"
        Set cF = g_db.GetFeiertagByDate(dtmRun)
        If Not cF Is Nothing Then   ' gefunden als Feiertag
            If cF.lngIdxAbwesenheitsArt > 0 Then
                strStatus = strStatus & " Varval=TRUE"
                If cF.lngIdxAbwesenheitsArt <> basGlobals.ABW_FEIERTAG Then    ' Wirklich Feiertag und nicht FAKO oder Urlaub ?
                    GoTo KeinFeiertag
                End If
                If booFeiertage Then
                    strFeiertage = strFeiertage & " - " & dtmRun & ": "
                    If cF.strFeiertag <> "" Then strFeiertage = strFeiertage & cF.strFeiertag
                    strFeiertage = strFeiertage & "(" & cAb.strAbwesenheitsart & ")"
                End If
            End If
        Else    ' Kein Pflicht-Abwesenheitstag
KeinFeiertag:
            strStatus = strStatus & " Varval=FALSE"
            If Weekday(dtmRun) <> g_db.WeekendFirstDay And Weekday(dtmRun) <> g_db.WeekendSecondDay Then       ' 1=Sonntag, 7=Samstag
                If booDetails Then
                    If cAb.lngIdxAbwesenheitsArt = basGlobals.ABW_URLAUB Then cAb.lngAnzahlUrlaub = cAb.lngAnzahlUrlaub + 1
                    If cAb.lngIdxAbwesenheitsArt = basGlobals.ABW_FAKO Then cAb.lngAnzahlFAKO = cAb.lngAnzahlFAKO + 1
                End If
                cAb.lngAnzahl = cAb.lngAnzahl + 1
            End If
        End If
        Set cF = Nothing
        dtmRun = dtmRun + 1
    Wend

    If booDetails Then
        Dim strAbwesenheitsTage As String
        strAbwesenheitsTage = ""
        If cAb.lngAnzahlUrlaub > 0 Then strAbwesenheitsTage = strAbwesenheitsTage & g_db.GetString(1031) & ":" & cAb.lngAnzahlUrlaub & " "
        If cAb.lngAnzahlFAKO > 0 Then strAbwesenheitsTage = strAbwesenheitsTage & g_db.GetString(1032) & ":" & cAb.lngAnzahlFAKO & " "
        If cAb.strAbwesenheitsart = "" Then
            strAbwesenheitsTage = strAbwesenheitsTage & g_db.GetString(1124) & ":" & cAb.lngAnzahl & " "
        Else
            If cAb.lngIdxAbwesenheitsArt <> basGlobals.ABW_FAKO And cAb.lngIdxAbwesenheitsArt <> basGlobals.ABW_URLAUB Then
                strAbwesenheitsTage = strAbwesenheitsTage & cAb.strAbwesenheitsart & ":" & cAb.lngAnzahl & " "
            End If
        End If
        CalculateTage = strAbwesenheitsTage
    Else
        CalculateTage = cAb.lngAnzahl
    End If
exit_calcTage:
    Exit Function
err_calcTage:
    MsgBox "Error in CalcTage: " & Err.Number & "  Desc: " & Err.Description & strStatus
    CalculateTage = ""
End Function

