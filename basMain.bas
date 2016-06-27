Attribute VB_Name = "basMain"
Option Explicit
Private clsM As clsMain

Public Sub Main()
    g_strProgramVersion = "" & App.ExeName & " V_" & App.Major & "." & App.Minor & "." & App.Revision
    basGlobals.g_iLCID = basLocale.GetUserLCID
'    If basGlobals.g_iLCID <> basGlobals.LocaleGerman And basGlobals.g_iLCID <> basGlobals.LocaleEnglish Then
'        MsgBox "Dieses Programm unterstützt nur zwei Spracheinstellungen: Deutsch/Deutschland und English/USA. Bitte stellen Sie eines davon ein." & vbCrLf & vbCrLf & _
'            "This program supports two language settings only: Deutsch/Deuschland and English/USA. Please setup your computer."
'        Exit Sub
'    End If
    If Not basCheckAccessRights.CanWrite() Then
        MsgBox "Sie haben keinen Schreibzugriff auf das folgende Verzeichnis/You need write permission for this folder:" & vbCrLf & vbCrLf & App.Path & vbCrLf & vbCrLf & "Setzen Sie sich mit Ihrem " & App.ExeName & "-Betreuer in Verbindung / Please contact your " & App.ExeName & " admin.", vbCritical
        Exit Sub
    End If
    
    Set g_db = New clsData
    If Not g_db.OpenDB Then Exit Sub ' Öffnen ist fehlgeschlagen
    If Not g_db.InitDB Then Exit Sub    ' Initialisieren ist fehlgeschlagen
    g_db.FillAllCollections ' Alle Daten einlesen
    '--------------------------------- einmalig holen, immer wieder verwenden
    Dim strCurrentUser As String
    strCurrentUser = basSysUtils.CurrentUser()
    '---------------------------------
    
    Set g_CU_Login = g_db.GetUserByAccountname(strCurrentUser)   ' Setzt einmalig  g_CU_Login mit strCurrentUser
    
    If g_CU_Login Is Nothing Then     ' strCurrentUser nicht in der DB gefunden
'        Ihr LoginName  & strCurrentUser &  wurde in der Datenbank nicht gefunden. Bitte wenden Sie sich an Ihren Orgleiter.
        g_strMessage = g_db.GetString(1118) & " " & strCurrentUser & " " & g_db.GetString(1119)
        frmMessage.Show vbModal
        g_LoginSucceeded = False
    Else    ' g_CU konnte gesetzt werden
        g_LoginSucceeded = True
        Set g_CU = g_CU_Login       ' am Anfang ist der angezeigte Benutzer = der eingeloggte Benutzer
        
        ' erstmalig UserItems setzen
        g_db.UpdateItemCollectionForCU
        
        Set g_clsAb = New clsAbwesenheit    ' die Funtionenklasse zu Abwesenheiten

        If g_db.GetItem("booCheckMinVersion", "False", "Check minimum version?") Then
            Dim strCurrentRevision As String
            Dim strMinVersion As String
            strCurrentRevision = Format(App.Major, "00") & "." & Format(App.Minor, "00") & "." & Format(App.Revision, "00")
            strMinVersion = g_db.GetItem("MinVersion", "02.50.0", "Minimum db version for current sw")
            If strCurrentRevision < strMinVersion Then
                MsgBox g_db.GetString(1120) & ": " & strCurrentRevision & vbCrLf & g_db.GetString(1121) & " " & strMinVersion & " " & g_db.GetString(1122)
                g_LoginSucceeded = False
                Exit Sub
            End If
        End If
        If g_CU.lngVerteiler = 0 Then
            frmDisclaimer.Show vbModal
            If g_CU.lngVerteiler = 0 Then   ' immer noch -> User konnte sich nicht entscheiden
                Exit Sub
            End If
        End If
        g_booShutdown = False
        g_booShowSort = g_db.GetItem("booShowSort", "True", "Show sort column.")
        g_booAlleMaSichtbar = g_db.GetItem("booAlleMaSichtbar", "False", "Members of Department are shown on one screen(Leader Dept + Sec)")
        g_booPlanEveryDay = g_db.GetItem("booPlanEveryDay", "False", "Saturday/sunday/holiday are valid dates")
        g_db.setOrgUserVisibility ' stelle die Sichtbarkeit der Org und user für diesen Benutzer her.
        Set clsM = New clsMain
        clsM.CreateForm
    End If
End Sub

Public Sub Main_Close_Application()
Dim strText As String
    g_booShutdown = True
    Unload frmPlanen
    Unload frmBeantragen
    Unload frmInfo
    Set g_clsAb = Nothing
    If g_StopApp = "NoRun" Then
        strText = g_db.GetItem("strNoRun", g_db.GetString(1123), "Reason for program stop")
    Else
        strText = "Idle Timeout."
    End If
    Set g_db = Nothing  ' Erst Datenbank schliessen, dann Message
    If g_StopApp <> "" Then
        g_strMessage = strText
        frmMessage.Show vbModal
        Exit Sub
    End If
End Sub
