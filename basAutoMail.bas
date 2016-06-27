Attribute VB_Name = "basAutoMail"
Option Explicit

Public Function AutoMail(ByVal IdxRecipient As Long, ByVal Betreff As String, ByVal Message As String, ByVal IdxAbwesenheit As Long, Optional booCC As Boolean = False) As Boolean
    On Error GoTo err_AutoMail
    
    Dim strStatus As String:    strStatus = "0_Start"

    Dim cAb As clsAbw:          Set cAb = g_db.GetAbwesenheit(IdxAbwesenheit)
    Dim Recipient As clsUser:   Set Recipient = g_db.GetUserByID(IdxRecipient)

    If g_CU.strMailName = STR_UNKNOWN Or Recipient.strMailName = STR_UNKNOWN Then GoTo exit_AutoMail

    strStatus = "0_GenCDO"
    Dim msg As CDO.Message:         Set msg = GenCdoMessage()

    Dim strTo As String:    strTo = Recipient.strMailName
    Dim strFrom As String:  strFrom = g_CU.strMailName  ' eingeloggter Benutzer - muss nicht der Vorgesetzte sein
    Dim strBCC As String
    If CBool(g_db.GetItem("booMailBCC", "True", "Send mail to sender too(default=true)?")) Then    ' Default - anschalten
        strBCC = strFrom     ' Auch an sich selbst schicken
    End If
    
'#######################################################################################################################################
Mail_Part_1:
    strStatus = "1_RecipientCC"
    On Error GoTo Mail_Part_1_error
    Dim strCC As String     ' Wenn einmal im CC etwas steht, kann kein weiterer Name hinzugefügt werden. Also Adressen sammeln und dann eintragen
    strCC = ""
    If booCC = True Then
        Dim OrgAbwUser As clsOrg:   Set OrgAbwUser = g_db.colOrg("cOrg_" & cAb.cUs.lngIdxOrg)
        g_db.Logging "Mail", cAb, strStatus & " cO:" & OrgAbwUser.strOrg & " IdxSek1/2:" & OrgAbwUser.lngIdxSek & "/" & OrgAbwUser.lngIdxSek2

        strStatus = "1_RecipientCC1" '##################################   Sek 1    ###############################################
        Dim cUsSek As clsUser
        If OrgAbwUser.lngIdxSek > 0 Then
            Set cUsSek = g_db.GetUserByID(OrgAbwUser.lngIdxSek)
            If Not cUsSek Is Nothing Then If cUsSek.strMailName <> "" Then strCC = cUsSek.strMailName
        End If  ' SekID1

        strStatus = "1_RecipientCC2" '##################################   Sek 2    ###############################################
        If OrgAbwUser.lngIdxSek2 > 0 Then   ' gibt es eine zweite Sek?
            Set cUsSek = g_db.GetUserByID(OrgAbwUser.lngIdxSek2)
            If Not cUsSek Is Nothing Then
                If cUsSek.strMailName <> "" Then
                    If strCC <> "" Then strCC = strCC & ";"
                    strCC = strCC & cUsSek.strMailName
                End If
            End If
        End If  ' SekID2 > 0
    End If  ' booCC
    
'#######################################################################################################################################
Mail_Part_2:
    On Error GoTo Mail_Part_2_error
    strStatus = "2_Genehmigt"
    If cAb.lngIdxStatus = AbwStatus.GENEHMIGT Or cAb.lngIdxStatus = AbwStatus.ABGELEHNT Then
        ' ist BCC der Vorgesetzte? Sonst diesen ebenfalls in BCC
        If g_CU.lngIdxUser <> cAb.cUs.lngIdxChef Then   ' Eingeloggter Benutzer ist nicht Chef
            Dim cUsChef As clsUser: Set cUsChef = g_db.GetUserByID(cAb.cUs.lngIdxChef)
            If strBCC <> "" Then strBCC = strBCC & ";"
            strBCC = strBCC & cUsChef.strMailName
        End If

        ' Wenn Abwesenheitsn von MA genehmigt werden, dann eventuell auch Mail an AL des MA
        If cAb.cUs.lngUserLevel = Benutzer Or cAb.cUs.lngUserLevel = Insel Then   ' lohnt sich nur für MA
            Dim cAbt As clsOrg:     Set cAbt = g_db.GetOrg(cAb.cUs.lngIdxAbt)   ' Abteilung des Abwesenheits-Benutzers
            If Not cAbt Is Nothing Then                                         ' Eigene Abteilung gefunden, dann ...
                strStatus = "2_Genehmigt_AddAL"
                If CBool(g_db.GetItem("MailImmerAL", "False", "Confirmation mail to department leader(default=false)?")) Then
                    ' Genehmigungs-Mail an AL
                    Dim cUsCh As clsUser:   Set cUsCh = g_db.GetAL(cAb.cUs)
                    If Not cUsCh Is Nothing Then
                        If strCC <> "" Then strCC = strCC & ";"
                        strCC = strCC & cUsCh.strMailName
                    End If
                End If

                strStatus = "2_Genehmigt_AddTL"
                If CBool(g_db.GetItem("MailAlleE4", "False", "Confirmation mail to every team leader(default=false)?")) Then
                    'Genehmigungs-Mail an alle TL
                    Dim colUs As New Collection
                    If g_db.GetE4OrgID(cAb.cUs.lngIdxAbt, colUs) > 0 Then
                        For Each cUsCh In colUs
                            If strCC <> "" Then strCC = strCC & ";"
                            strCC = strCC & cUsCh.strMailName
                        Next
                    End If
                End If
            End If
        End If
    End If
'#######################################################################################################################################
Mail_Part_3:
    On Error GoTo Mail_Part_3_error
    strStatus = "3_Genehmigt_ICS"
    ' Mail des ICS-Files an Vorgesetzten + optional ChefChef
    If cAb.lngIdxStatus = AbwStatus.GENEHMIGT Or cAb.lngIdxStatus = AbwStatus.ABGELEHNT Or cAb.lngIdxStatus = AbwStatus.ZURUCKGEZOGEN Then
        If CBool(g_db.GetItem("booMailICS", "True", "Send Mail with ICS attachment(default=true)?")) Then    ' Default - anschalten
            ' ICS-string in File schreiben
            strStatus = "3_GenICS"
            Dim strICS As String:           strICS = GenIcsFile(cAb)
            If strICS <> "" Then
                Dim strFilename As String:      strFilename = GetTempFolder
                If strFilename = "" Then    ' kein Temp-Folder
                    g_db.Logging "Mail", cAb, "Unable to locate TempFolder. Skip ICS"
                Else
                    strFilename = strFilename & "\map" & Format(Now(), "yyyymmdd_hhMMss") & ".ics"
                    Dim fso As FileSystemObject:    Set fso = New FileSystemObject
                    Dim txtStream As TextStream:    Set txtStream = fso.CreateTextFile(strFilename, True)
                    txtStream.Write strICS:         txtStream.Close
        
                    strStatus = "3_GenMessage"
        
                    Dim iBp As CDO.IBodyPart
                    Set iBp = msg.AddAttachment(strFilename)
                    Set iBp = Nothing   ' TODO kann iBp jetzt schon gelöscht werden?
                    ' TODO: kann jetzt das File schon gelöscht werden?
                End If
            End If
        End If
    End If
'#######################################################################################################################################
Mail_Part_4:
    On Error GoTo Mail_Part_4_error
    strStatus = "4_FillMsg"

    msg.Subject = App.ExeName & " ... " & Betreff
    Dim UNCPath As String
    UNCPath = basUNCPath.GetUNCPath(App.Path & "\")
    msg.TextBody = Message & vbCrLf & vbCrLf & "Gezeichnet " & g_CU.Fullname & vbCrLf & "Link to " & App.ExeName & ": file:" & UNCPath & App.ExeName & ".exe"
    msg.From = strFrom
    msg.ReplyTo = strFrom
    msg.To = strTo
    msg.cC = strCC
    msg.BCC = strBCC
    g_db.Logging "Mail", cAb, "From=" & strFrom & " To=" & strTo & " CC=" & strCC & " Bcc=" & strBCC & " Betreff=" & msg.Subject & " Message=" & msg.TextBody

    If CBool(g_db.GetItem("booSendMail", "True", "Send mail at all(default=true)?")) Then
        strStatus = "4_SendMail"
        On Error Resume Next
        Err.Clear
        msg.Send
        Dim intErrNumber As Integer:    intErrNumber = Err.Number
    
        If intErrNumber > 0 Then
            g_db.Logging "Mail", cAb, "ERROR /" & strStatus & "/Err:" & intErrNumber & " Desc:" & Err.Description
            AutoMail = False
        Else
            g_db.Logging "Mail", cAb, "SUCCESS /" & strStatus
            AutoMail = True
        End If
    Else
        g_db.Logging "Mail", cAb, "skipped"
    End If

exit_AutoMail:
    On Error Resume Next
    Set iBp = Nothing
    Err.Clear
    If strFilename <> "" Then fso.DeleteFile strFilename  ' Nach Send erst löschen oder früher ?
    If Err > 0 Then Debug.Print Err.Description
    Set fso = Nothing

    Set msg = Nothing
    Exit Function
'#######################################################################################################################################
err_AutoMail:
    MsgBox "ERROR /" & strStatus & "/Err:" & Err.Number & " Desc:" & Err.Description, , "AutoMail V" & basGlobals.g_strProgramVersion
    g_db.Logging "Mail", cAb, "ERROR /" & strStatus & "/Err:" & Err.Number & " Desc:" & Err.Description
    AutoMail = False
    Resume exit_AutoMail
Mail_Part_1_error:
    MsgBox "ERROR /" & strStatus & "/Err:" & Err.Number & " Desc:" & Err.Description, , "AutoMail V" & basGlobals.g_strProgramVersion
    g_db.Logging "Mail_Part1", cAb, "ERROR /" & strStatus & "/Err:" & Err.Number & " Desc:" & Err.Description
    Resume Mail_Part_2
Mail_Part_2_error:
    MsgBox "ERROR /" & strStatus & "/Err:" & Err.Number & " Desc:" & Err.Description, , "AutoMail V" & basGlobals.g_strProgramVersion
    g_db.Logging "Mail_Part2", cAb, "ERROR /" & strStatus & "/Err:" & Err.Number & " Desc:" & Err.Description
    Resume Mail_Part_3
Mail_Part_3_error:
    MsgBox "ERROR /" & strStatus & "/Err:" & Err.Number & " Desc:" & Err.Description, , "AutoMail V" & basGlobals.g_strProgramVersion
    g_db.Logging "Mail_Part3", cAb, "ERROR /" & strStatus & "/Err:" & Err.Number & " Desc:" & Err.Description
    Resume Mail_Part_4
Mail_Part_4_error:
    MsgBox "ERROR /" & strStatus & "/Err:" & Err.Number & " Desc:" & Err.Description, , "AutoMail V" & basGlobals.g_strProgramVersion
    g_db.Logging "Mail_Part4", cAb, "ERROR /" & strStatus & "/Err:" & Err.Number & " Desc:" & Err.Description
    AutoMail = False
    Resume exit_AutoMail
End Function


Private Function GenCdoMessage() As CDO.Message
    Dim cdomsg As CDO.Message
    Set cdomsg = New CDO.Message
    cdomsg.Configuration.Fields(cdoSMTPServer) = g_db.GetItem("SMTP-Server", "smtp.mtu-online.com", "URL for smtp-Server")
    cdomsg.Configuration.Fields(cdoSMTPAuthenticate) = cdoAnonymous
    cdomsg.Configuration.Fields(cdoSMTPConnectionTimeout) = 30
    cdomsg.Configuration.Fields(cdoSendUsingMethod) = cdoSendUsingPort
    cdomsg.Configuration.Fields.Update
    Set GenCdoMessage = cdomsg
End Function
Private Function GetTempFolder() As String
    If Environ("TEMP") <> "" Then
        GetTempFolder = Environ("TEMP")
        Exit Function
    End If
    If Environ("TMP") <> "" Then
        GetTempFolder = Environ("TMP")
        Exit Function
    End If
    If Environ("USERPROFILE") <> "" Then
        GetTempFolder = Environ("USERPROFILE")
        Exit Function
    End If
    Dim fso As FileSystemObject:    Set fso = New FileSystemObject
    If fso.FolderExists("c:\temp") Then
        GetTempFolder = "c:\temp"
        Exit Function
    End If
    GetTempFolder = ""
End Function
Private Function GenIcsFile(cAb As clsAbw) As String
    On Error GoTo err_GenIcsFile
    
    GenIcsFile = "BEGIN:VCALENDAR" & vbCrLf
    GenIcsFile = GenIcsFile & "PRODID:-//mtu electronics//" & App.ExeName & " 2.2//GE" & vbCrLf
    GenIcsFile = GenIcsFile & "VERSION:2.0" & vbCrLf
    
    If cAb.lngIdxStatus = AbwStatus.GENEHMIGT Then
        GenIcsFile = GenIcsFile & "METHOD: PUBLISH" & vbCrLf
    Else
        GenIcsFile = GenIcsFile & "METHOD: CANCEL" & vbCrLf
    End If
    
    GenIcsFile = GenIcsFile & "X-MS-OLK-FORCEINSPECTOROPEN:FALSE" & vbCrLf

    GenIcsFile = GenIcsFile & "BEGIN:VTIMEZONE" & vbCrLf & "TZID:FN" & vbCrLf & "BEGIN:STANDARD" & vbCrLf & "DTSTART:16011028T030000" & vbCrLf & "RRULE:FREQ=YEARLY;BYDAY=-1SU;BYMONTH=10" & vbCrLf & "TZOFFSETFROM:+0200" & vbCrLf & "TZOFFSETTO:+0100" & vbCrLf & "END:STANDARD" & vbCrLf & "BEGIN:DAYLIGHT" & vbCrLf & "DTSTART:16010325T020000" & vbCrLf & "RRULE:FREQ=YEARLY;BYDAY=-1SU;BYMONTH=3" & vbCrLf & "TZOFFSETFROM:+0100" & vbCrLf & "TZOFFSETTO:+0200" & vbCrLf & "END:DAYLIGHT" & vbCrLf & "END:VTIMEZONE" & vbCrLf

    GenIcsFile = GenIcsFile & "BEGIN:VEVENT" & vbCrLf
    GenIcsFile = GenIcsFile & "UID:map" & Format(cAb.dtmErstellung, "yyyymmddhhmmss") & vbCrLf
    If cAb.lngIdxStatus = AbwStatus.GENEHMIGT Then
        GenIcsFile = GenIcsFile & "SEQUENCE:0" & vbCrLf
    Else
        GenIcsFile = GenIcsFile & "SEQUENCE:1" & vbCrLf
    End If
    GenIcsFile = GenIcsFile & "CATEGORIES:" & App.ExeName & vbCrLf
    GenIcsFile = GenIcsFile & "CLASS:PUBLIC" & vbCrLf

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
    timStart = cAb.dtmStart + timStart
    timEnd = cAb.dtmEnde + timEnd

    GenIcsFile = GenIcsFile & "DTEND:" & Format(timEnd, "yyyymmddThhmmss") & vbCrLf
    GenIcsFile = GenIcsFile & "DTSTAMP:" & Format(Now(), "yyyymmddThhmmssZ") & vbCrLf
    GenIcsFile = GenIcsFile & "DTSTART:" & Format(timStart, "yyyymmddThhmmss") & vbCrLf
    
    GenIcsFile = GenIcsFile & "SUMMARY:" & cAb.UserInfo & cAb.strAbwesenheitsart
    If cAb.strText <> "" Then GenIcsFile = GenIcsFile & " - " & cAb.strText
    GenIcsFile = GenIcsFile & vbCrLf
    
    If cAb.lngIdxStatus = AbwStatus.GENEHMIGT Then
        GenIcsFile = GenIcsFile & "DESCRIPTION:Publish - Autogenerated by " & App.ExeName & "(c)" & vbCrLf
    Else
        GenIcsFile = GenIcsFile & "DESCRIPTION:Cancel - Autogenerated by " & App.ExeName & "(c)" & vbCrLf
    End If

    GenIcsFile = GenIcsFile & "TRANSP:TRANSPARENT" & vbCrLf
    GenIcsFile = GenIcsFile & "X-MICROSOFT-CDO-BusyStatus:FREE" & vbCrLf
    GenIcsFile = GenIcsFile & "X-MICROSOFT-CDO-IMPORTANCE:1" & vbCrLf
    GenIcsFile = GenIcsFile & "X-MICROSOFT-DISALLOW-COUNTER:FALSE" & vbCrLf
    GenIcsFile = GenIcsFile & "X-MS-OLK-ALLOWEXTERNCHECK:TRUE" & vbCrLf
    GenIcsFile = GenIcsFile & "X-MS-OLK-AUTOSTARTCHECK:FALSE" & vbCrLf
    GenIcsFile = GenIcsFile & "X-MS-OLK-CONFTYPE:0" & vbCrLf

    GenIcsFile = GenIcsFile & "END:VEVENT" & vbCrLf

    GenIcsFile = GenIcsFile & "END:VCALENDAR" & vbCrLf
    
    GenIcsFile = ConvertToASCII(GenIcsFile)
    Exit Function
err_GenIcsFile:
    GenIcsFile = ""
    g_db.Logging "Mail_GenIcsFile", cAb, Err.Description
    Exit Function
End Function
