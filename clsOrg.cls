VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOrg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Aus DB
Public lngIdxOrg As Long        ' laufende Nummer
Public strOrg As String         ' Klartext
Public lngOrgLevel As OrgLevel  ' Ebene laut Tognum 0-4
Public lngIdxTopOrg As Long     ' OrgIndex der übergeordneten Organisation
Public lngIdxChef As Long       ' UserIndex Leiter1 dieser Org
Public lngIdxChef2 As Long      ' UserIndex Leiter2 dieser Org
Public lngIdxSek As Long        ' UserIndex Sek1 optional
Public lngIdxSek2 As Long       ' UserIndex Sek2 optional
Public lngTageLevel1 As Long    ' Level1-Genehmigung in dieser Org bis tage

'Abgeleitete Eigenschaften
Public booVisible As Boolean

Public lngIdxTopTopOrg As Long  ' OrgIndex der nächsten übergeordneten Organisation
Public cTopOrg As clsOrg        ' Referenz auf TopOrg
Public colSubOrg As Collection  ' Sammlung der untergeordneten Organisationen
Public colMember As Collection  ' Sammlung aller MA in dieser Organisation

' Einordnung in den Baum
Public lngIdxRes As Long        ' Vererbt aus tblOrg
Public lngIdxBer As Long        ' Vererbt aus tblOrg
Public lngIdxAbt As Long        ' Vererbt aus tblOrg
Public lngIdxTeam As Long       ' Vererbt aus tblOrg
Public lngIdxInsel As Long      ' Vererbt aus tblOrg

' Bewertung für den g_CU
Public intAbstand As Integer    ' Abstand der anderen Org von der Org des g_CU
