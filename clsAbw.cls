VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAbw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Aus DB
Public lngIdxAbwesenheit As Long
Public lngIdxAbwesenheitsArt As Long
Public lngIdxUser As Long
Public dtmStart As Date
Public dtmEnde As Date
Public lngGVN As Long
Public lngIdxStatus As Long
Public lngIdxStatusOld As Long
Public dtmErstellung As Date
Public dtmAntrag As Date
Public dtmGenehmigt1 As Date
Public dtmGenehmigt2 As Date
Public strText As String

' Generiert
Public strAbwesenheitsart As String
Public strStatus As String
Public cUs As clsUser
Public strGVN As String

Public lngAnzahl As Long
Public lngAnzahlUrlaub As Long
Public lngAnzahlFAKO As Long

' Für die Änderung con Abwesenheiten
Public strOutlookText As String

Public Function UserInfo() As String
    Dim Output As String
    Output = cUs.strNachname & ", " & cUs.strVorname & " - "
    If cUs.lngPN <> 0 Then Output = Output & cUs.lngPN & " - "
    UserInfo = Output
End Function
