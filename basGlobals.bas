Attribute VB_Name = "basGlobals"
Option Explicit

Public g_db As clsData
Public g_clsAb As clsAbwesenheit  ' Die Tools zu Abwesenheit
Public g_strProgramVersion As String
Public strSQL As String
Public g_booShutdown As Boolean ' Zustand: Applikation runterfahren und Fenster schließen
Public g_StopApp As String  ' Die Notbremse

'----------------------------------------------------------------------------- statisch per Login festgelegt
Public g_CU_Login As clsUser    ' Info zu eingeloggtem Benutzer
'----------------------------------------------------------------------------- dynamisch
Public g_CU As clsUser          ' Info zu dem aktuellen Benutzer aus der Datenbank

Public g_debugCurrentUser As String ' zum Testen - als welcher User erscheinen ?
Public g_PopUpMode As String ' Was soll angezeigt werden - Text oder Tabelle
Public g_InfoText As String ' Text, der auf dem PopUpFormular angezeigt wird
Public g_InfoTabelle() As Long    ' Tabellendaten
Public g_DataTabelle() As String  ' Steuerdaten zur Tabellenansicht
Public g_Quartal(8, 2) As Long ' Im QuartalSlider: AnfangsMonat und Jahr
Public g_booPlanEveryDay As Boolean  ' Auch Samstag/Sonntag/Feiertag planen?

' Sichtbar für Vorgesetzten / Kollegen
Public Const SICHTBAR_CHEF As Long = 1
Public Const SICHTBAR_Org As Long = 2
Public Const SICHTBAR_ABT As Long = 3
Public Const SICHTBAR_BER As Long = 4
Public Const SICHTBAR_RES As Long = 5

Public g_LoginSucceeded As Boolean
Public g_strMessage As String         ' Parameter für die frmMessage
Public g_booShowSort As Boolean       ' Zeige 2. Spalte - SortierStrings
Public g_booAlleMaSichtbar As Boolean  ' Superuser sieht die ganze Abteilung auf einen Blick

'Rollen, mit denen man in's Programm kommen kann - Sortierung Absteigend darstellen
Public Const SORTORDER_V As String = "A"        ' Sprecher der GF
Public Const SORTORDER_V_SEK As String = "B"
Public Const SORTORDER_GF As String = "C"       ' GF
Public Const SORTORDER_GF_SEK As String = "D"
Public Const SORTORDER_BL As String = "E"       ' Bereichleiter
Public Const SORTORDER_BL_SEK As String = "F"   ' Bereichleiter Sekretariat
Public Const SORTORDER_AL As String = "G"       ' Abteilungsleiter
Public Const SORTORDER_AL_SEK As String = "H"   ' Abteilungsleiter Sekretariat
Public Const SORTORDER_TL As String = "I"       ' Teamleiter
Public Const SORTORDER_TL_SEK As String = "J"   ' Teamleiter Sekretariat
Public Const SORTORDER_USER As String = "K"     ' Der normale Benutzer

' Verwendet für OrgLevel und UserLevel
Public Enum OrgLevel
    Ressort = 1     ' Ressortleiter E1
    Bereich = 2     ' Bereichsleiter E2
    Abteilung = 3   ' Abteilungsleiter E3
    Team = 4        ' Teamleiter E4
    Insel = 5       ' Inselleiter E5
    Benutzer = 6    ' TeamMember E6
End Enum


' Abwesenheiten
Public Const ABW_URLAUB As Long = 1
Public Const ABW_FAKO As Long = 2
Public Const ABW_SEMINAR As Long = 3
Public Const ABW_DIENSTREISE As Long = 4
Public Const ABW_SONSTIGES As Long = 5
Public Const ABW_FEIERTAG As Long = 6
Public Const ABW_SCHULFERIEN As Long = 7
Public Const ABW_KRANKHEIT As Long = 8

' Ganztägig ?
Public Const ABW_GANZTAGS As Long = 0
Public Const ABW_VORMITTAGS As Long = 1
Public Const ABW_NACHMITTAGS As Long = 2

Public Const DBUSER As String = "map"
Public Const DBPASSWORD As String = "mapelektronik"

Public Const STR_UNKNOWN As String = "???"
Public Const LNG_UNKOWN As Long = &H10000000

Public Const OFFSET_FGDATE As Long = 2  ' Die Spalte des erten Tages im FlexGrid FG in der Form frmInfo
Public Const COL_A_GEPLANT As Long = 1
Public Const COL_A_GENEHMIGT As Long = 2

Public Const ROW_MONAT As Long = 0      ' Zeilen in frmInfo
Public Const ROW_KW As Long = 1
Public Const ROW_TAG As Long = 2

'Formulare frmBeantragen + frmManager
Public Const COL_IDX As Long = 0        ' Index Abwesenheit
Public Const COL_IDXUSER As Long = 1    ' Index Benutzer
Public Const COL_USER As Long = 2       ' Name Benutzer
Public Const COL_IDXTYP As Long = 3     ' Index AbwesenheitsTyp
Public Const COL_ORG As Long = 4        ' Orgname des Benutzers
Public Const COL_TYP As Long = 5        ' Name AbwesenheitsTyp
Public Const COL_START As Long = 6      ' Datum
Public Const COL_ENDE As Long = 7       ' Datum
Public Const COL_GVN As Long = 8        ' Ganztags Vormittag, Nachmittag
Public Const COL_LAENGE As Long = 9     ' Anzahl Tage
Public Const COL_IDXSTAT As Long = 10   ' Index Status
Public Const COL_STAT As Long = 11      ' Name Status
Public Const COL_TEXT As Long = 12      ' Zusatz-Text
Public Const COL_MAX As Long = 12

Public Const COLOR_URLAUB As Long = 49152       '&HC000     darkGreen   ' vbGreen     ' &H00FF00
Public Const COLOR_FAKO As Long = vbBlue        ' &HFF0000
Public Const COLOR_FEIERTAG As Long = vbRed     ' &H0000FF
Public Const COLOR_DIENSTREISE As Long = 49344  ' &HC0C0    darkYellow  '  vbYellow   ' &H00FFFF
Public Const COLOR_SEMINAR As Long = 12632064   ' &HC0C000  darkCyan    ' vbCyan     ' &HFFFF00
Public Const COLOR_SONSTIGES As Long = vbMagenta    ' &HFF00FF
Public Const COLOR_SCHULFERIEN As Long = 10539263   ' 15790320 &H00F0F0F0&   10539263 &H00A0D0FF
Public Const COLOR_KRANKHEIT As Long = 8438015      ' &H0080C0FF&   ' orange
Public Const COLOR_PLANEN As Long = vbYellow
Public Const COLOR_FREI As Long = 0

Public Const COLOR_URLAUB_B As Long = 8454016       ' &H80FF80 ' lightGreen
Public Const COLOR_FAKO_B As Long = 16744576        ' &HFF8080 ' lightBlue
Public Const COLOR_DIENSTREISE_B As Long = 8454134  ' &H80FFFF ' lightYellow
Public Const COLOR_SEMINAR_B As Long = 16777088     ' &HFFFF80 ' lightCyan
Public Const COLOR_SONSTIGES_B As Long = 16744703       ' &HFF80FF ' lightMagenta

Public Const COLOR_GRAY As Long = 13421772          ' &HCCCCCC

Public Const COL_QUARTAL = 0
Public Const COL_MONAT = 1
Public Const COL_JAHR = 2

Public Const VIEW_WOCHEN = 8

Public g_iLCID As Integer
Public Const LocaleGerman = 1031
Public Const LocaleEnglish = 1033
