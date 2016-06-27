Attribute VB_Name = "basLocale"
Option Explicit

Private Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long
Private Declare Function GetUserDefaultLCID Lib "kernel32.dll" () As Long
Private Declare Function GetLocaleInfo Lib "kernel32" _
        Alias "GetLocaleInfoA" (ByVal Locale As Long, _
        ByVal LCType As Long, ByVal lpLCData As String, _
        ByVal cchData As Long) As Long

Public Const LOCALE_ILANGUAGE = &H1
Public Const LOCALE_SLANGUAGE = &H2
Public Const LOCALE_SENGLANGUAGE = &H1001
Public Const LOCALE_SABBREVLANGNAME = &H3
Public Const LOCALE_SNATIVELANGNAME = &H4
Public Const LOCALE_ICOUNTRY = &H5
Public Const LOCALE_SCOUNTRY = &H6
Public Const LOCALE_SENGCOUNTRY = &H1002
Public Const LOCALE_SABBREVCTRYNAME = &H7
Public Const LOCALE_SNATIVECTRYNAME = &H8
Public Const LOCALE_IDEFAULTLANGUAGE = &H9
Public Const LOCALE_IDEFAULTCOUNTRY = &HA
Public Const LOCALE_IDEFAULTCODEPAGE = &HB
Public Const LOCALE_SLIST = &HC
Public Const LOCALE_IMEASURE = &HD
Public Const LOCALE_SDECIMAL = &HE
Public Const LOCALE_STHOUSAND = &HF
Public Const LOCALE_SGROUPING = &H10
Public Const LOCALE_IDIGITS = &H11
Public Const LOCALE_ILZERO = &H12
Public Const LOCALE_SNATIVEDIGITS = &H13
Public Const LOCALE_SCURRENCY = &H14
Public Const LOCALE_SINTLSYMBOL = &H15
Public Const LOCALE_SMONDECIMALSEP = &H16
Public Const LOCALE_SMONTHOUSANDSEP = &H17
Public Const LOCALE_SMONGROUPING = &H18
Public Const LOCALE_ICURRDIGITS = &H19
Public Const LOCALE_IINTLCURRDIGITS = &H1A
Public Const LOCALE_ICURRENCY = &H1B
Public Const LOCALE_INEGCURR = &H1C
Public Const LOCALE_SDATE = &H1D
Public Const LOCALE_STIME = &H1E
Public Const LOCALE_SSHORTDATE = &H1F
Public Const LOCALE_SLONGDATE = &H20
Public Const LOCALE_STIMEFORMAT = &H1003
Public Const LOCALE_IDATE = &H21
Public Const LOCALE_ILDATE = &H22
Public Const LOCALE_ITIME = &H23
Public Const LOCALE_ICENTURY = &H24
Public Const LOCALE_ITLZERO = &H25
Public Const LOCALE_IDAYLZERO = &H26
Public Const LOCALE_IMONLZERO = &H27
Public Const LOCALE_S1159 = &H28
Public Const LOCALE_S2359 = &H29
Public Const LOCALE_SDAYNAME1 = &H2A
Public Const LOCALE_SDAYNAME2 = &H2B
Public Const LOCALE_SDAYNAME3 = &H2C
Public Const LOCALE_SDAYNAME4 = &H2D
Public Const LOCALE_SDAYNAME5 = &H2E
Public Const LOCALE_SDAYNAME6 = &H2F
Public Const LOCALE_SDAYNAME7 = &H30
Public Const LOCALE_SABBREVDAYNAME1 = &H31
Public Const LOCALE_SABBREVDAYNAME2 = &H32
Public Const LOCALE_SABBREVDAYNAME3 = &H33
Public Const LOCALE_SABBREVDAYNAME4 = &H34
Public Const LOCALE_SABBREVDAYNAME5 = &H35
Public Const LOCALE_SABBREVDAYNAME6 = &H36
Public Const LOCALE_SABBREVDAYNAME7 = &H37
Public Const LOCALE_SMONTHNAME1 = &H38
Public Const LOCALE_SMONTHNAME2 = &H39
Public Const LOCALE_SMONTHNAME3 = &H3A
Public Const LOCALE_SMONTHNAME4 = &H3B
Public Const LOCALE_SMONTHNAME5 = &H3C
Public Const LOCALE_SMONTHNAME6 = &H3D
Public Const LOCALE_SMONTHNAME7 = &H3E
Public Const LOCALE_SMONTHNAME8 = &H3F
Public Const LOCALE_SMONTHNAME9 = &H40
Public Const LOCALE_SMONTHNAME10 = &H41
Public Const LOCALE_SMONTHNAME11 = &H42
Public Const LOCALE_SMONTHNAME12 = &H43
Public Const LOCALE_SABBREVMONTHNAME1 = &H44
Public Const LOCALE_SABBREVMONTHNAME2 = &H45
Public Const LOCALE_SABBREVMONTHNAME3 = &H46
Public Const LOCALE_SABBREVMONTHNAME4 = &H47
Public Const LOCALE_SABBREVMONTHNAME5 = &H48
Public Const LOCALE_SABBREVMONTHNAME6 = &H49
Public Const LOCALE_SABBREVMONTHNAME7 = &H4A
Public Const LOCALE_SABBREVMONTHNAME8 = &H4B
Public Const LOCALE_SABBREVMONTHNAME9 = &H4C
Public Const LOCALE_SABBREVMONTHNAME10 = &H4D
Public Const LOCALE_SABBREVMONTHNAME11 = &H4E
Public Const LOCALE_SABBREVMONTHNAME12 = &H4F
Public Const LOCALE_SABBREVMONTHNAME13 = &H100F
Public Const LOCALE_SPOSITIVESIGN = &H50
Public Const LOCALE_SNEGATIVESIGN = &H51
Public Const LOCALE_IPOSSIGNPOSN = &H52
Public Const LOCALE_INEGSIGNPOSN = &H53
Public Const LOCALE_IPOSSYMPRECEDES = &H54
Public Const LOCALE_IPOSSEPBYSPACE = &H55
Public Const LOCALE_INEGSYMPRECEDES = &H56
Public Const LOCALE_INEGSEPBYSPACE = &H57

Public Const LOCALE_USER_DEFAULT = &H400
Public Const LOCALE_SYSTEM_DEFAULT As Long = &H400

Public Sub PrintAll()
  'List1.Clear
  
  GLI LOCALE_SLIST, "Listentrennzeichen"
  GLI LOCALE_IMEASURE, "0=metrisch, 1=US"
  GLI LOCALE_SDECIMAL, "Dezimaltrennzeichen"
  GLI LOCALE_STHOUSAND, "Tausendertrennzeichen"
  GLI LOCALE_SGROUPING, "Gruppierung links vom Komma"
  GLI LOCALE_IDIGITS, "Zahlen hinter dem Komma"
  GLI LOCALE_ILZERO, "führende Nullen"
  GLI LOCALE_SCURRENCY, "Währungsymbol"
  GLI LOCALE_SINTLSYMBOL, "Währung nach ISO 4217"
  GLI LOCALE_SMONDECIMALSEP, "Währungstrennzeichen"
  GLI LOCALE_SMONTHOUSANDSEP, "Währungstausendertrennzeichen"
  GLI LOCALE_SMONGROUPING, "Währungsgruppierung"
  GLI LOCALE_ICURRDIGITS, "Zahlen hinter dem Komma (Pf)"
  GLI LOCALE_ICURRENCY, "Anzeige des Währungssymbols"
  GLI LOCALE_INEGCURR, "Negatives Währungsvorzeichen"
  GLI LOCALE_SDATE, "Datumstrennzeichen"
  GLI LOCALE_STIME, "Zeittrennzeichen"
  GLI LOCALE_SSHORTDATE, "Kurzes Datumsformat"
  GLI LOCALE_SLONGDATE, "Langes Datumsformat"
  GLI LOCALE_STIMEFORMAT, "Zeitformat"
  GLI LOCALE_ITIME, "12/24 Stunden"
  GLI LOCALE_S1159, "AM-Zeichen"
  GLI LOCALE_S2359, "PM-Zeichen"
  GLI LOCALE_SPOSITIVESIGN, "Positives Vorz."
  GLI LOCALE_SNEGATIVESIGN, "Negatives Vorz."
  GLI LOCALE_ILANGUAGE, "Sprach ID"
  GLI LOCALE_SLANGUAGE, "Lokalisierter Sprachname"
  GLI LOCALE_SENGLANGUAGE, "Engl. Äquivalent"
  GLI LOCALE_SABBREVLANGNAME, "Abgekürzt"
  GLI LOCALE_SNATIVELANGNAME, "Sprache in Landessprache"
  GLI LOCALE_ICOUNTRY, "Ländercode"
  GLI LOCALE_SCOUNTRY, "Ländername"
  GLI LOCALE_SENGCOUNTRY, "Ländername in Engl."
  GLI LOCALE_SABBREVCTRYNAME, "Abgekürzt"
  GLI LOCALE_SNATIVECTRYNAME, "Land in Landessprache"
  GLI LOCALE_IDEFAULTLANGUAGE, "Standard Sprach-ID"
  GLI LOCALE_IDEFAULTCOUNTRY, "Standard Landes-ID"
  GLI LOCALE_IDEFAULTCODEPAGE, "Standard Codeseite"
  GLI LOCALE_SNATIVEDIGITS, "gebräuchliche Zahlen"
  GLI LOCALE_IINTLCURRDIGITS, "Zahlen hinter Komma nach ISO"
  GLI LOCALE_IDATE, "Datums Gruppierung"
  GLI LOCALE_ILDATE, "Reihenfolge langes Datumsformat"
  GLI LOCALE_ICENTURY, "Jahr in 2/4 Ziffern"
  GLI LOCALE_ITLZERO, "führende Null für Zeiten"
  GLI LOCALE_IDAYLZERO, "führende Null für Tage"
  GLI LOCALE_IMONLZERO, "führende Null für Monate"
  GLI LOCALE_SDAYNAME1, "Langer Name für Mo"
  GLI LOCALE_SDAYNAME2, "Langer Name für Di"
  GLI LOCALE_SDAYNAME3, "Langer Name für Mi"
  GLI LOCALE_SDAYNAME4, "Langer Name für Do"
  GLI LOCALE_SDAYNAME5, "Langer Name für Fr"
  GLI LOCALE_SDAYNAME6, "Langer Name für Sa"
  GLI LOCALE_SDAYNAME7, "Langer Name für So"
  GLI LOCALE_SABBREVDAYNAME1, "Abgk. Name für Mo"
  GLI LOCALE_SABBREVDAYNAME2, "Abgk. Name für Di"
  GLI LOCALE_SABBREVDAYNAME3, "Abgk. Name für Mi"
  GLI LOCALE_SABBREVDAYNAME4, "Abgk. Name für Do"
  GLI LOCALE_SABBREVDAYNAME5, "Abgk. Name für Fr"
  GLI LOCALE_SABBREVDAYNAME6, "Abgk. Name für Sa"
  GLI LOCALE_SABBREVDAYNAME7, "Abgk. Name für So"
  GLI LOCALE_SMONTHNAME1, "Langer Name für Jan"
  GLI LOCALE_SMONTHNAME2, "Langer Name für Feb"
  GLI LOCALE_SMONTHNAME3, "Langer Name für Mae"
  GLI LOCALE_SMONTHNAME4, "Langer Name für Mai"
  GLI LOCALE_SMONTHNAME5, "Langer Name für Apr"
  GLI LOCALE_SMONTHNAME6, "Langer Name für Jun"
  GLI LOCALE_SMONTHNAME7, "Langer Name für Jul"
  GLI LOCALE_SMONTHNAME8, "Langer Name für Aug"
  GLI LOCALE_SMONTHNAME9, "Langer Name für Sep"
  GLI LOCALE_SMONTHNAME10, "Langer Name für Okt"
  GLI LOCALE_SMONTHNAME11, "Langer Name für Nov"
  GLI LOCALE_SMONTHNAME12, "Langer Name für Dez"
  GLI LOCALE_SABBREVMONTHNAME1, "Abgk. Name für Jan"
  GLI LOCALE_SABBREVMONTHNAME2, "Abgk. Name für Feb"
  GLI LOCALE_SABBREVMONTHNAME3, "Abgk. Name für Mae"
  GLI LOCALE_SABBREVMONTHNAME4, "Abgk. Name für Apr"
  GLI LOCALE_SABBREVMONTHNAME5, "Abgk. Name für Mai"
  GLI LOCALE_SABBREVMONTHNAME6, "Abgk. Name für Jun"
  GLI LOCALE_SABBREVMONTHNAME7, "Abgk. Name für Jul"
  GLI LOCALE_SABBREVMONTHNAME8, "Abgk. Name für Aug"
  GLI LOCALE_SABBREVMONTHNAME9, "Abgk. Name für Sep"
  GLI LOCALE_SABBREVMONTHNAME10, "Abgk. Name für Okt"
  GLI LOCALE_SABBREVMONTHNAME11, "Abgk. Name für Nov"
  GLI LOCALE_SABBREVMONTHNAME12, "Abgk. Name für Dez"
  GLI LOCALE_IPOSSIGNPOSN, "Format. für pos. Währung"
  GLI LOCALE_INEGSIGNPOSN, "Format. für neg. Währung"
  GLI LOCALE_IPOSSYMPRECEDES, "Präfix für pos. Währungsvorzeichen"
  GLI LOCALE_IPOSSEPBYSPACE, "Trennz. bei pos. Währungsbetrag"
  GLI LOCALE_INEGSYMPRECEDES, "Präfix für neg. Währungsvorzeichen"
  GLI LOCALE_INEGSEPBYSPACE, "Trennz. bei neg. Währungsbetrag"
End Sub

Private Sub GLI(ID&, Text$)
    Debug.Print ID & ":" & Text & vbTab & GetEntry(ID)
'  List1.AddItem Text & ":  " & GetEntry(ID)
'  List1.ItemData(List1.NewIndex) = ID
End Sub

Public Function GetSystemLCID() As Integer
    GetSystemLCID = GetSystemDefaultLCID()
End Function
Public Function GetUserLCID() As Integer
    GetUserLCID = GetUserDefaultLCID()
End Function
Public Function GetEntry(ID&) As String
  Dim LCID&, Result&, Buffer$, Length&
    
    LCID = GetUserDefaultLCID()
    Length = GetLocaleInfo(LCID, ID, Buffer, 0) - 1
    Buffer = Space(Length + 1)
    Result = GetLocaleInfo(LCID, ID, Buffer, Length)
    GetEntry = Left$(Buffer, Length)
End Function


