VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' frmInfo
Public ShowUserList As Boolean
Public OrgSlider As Boolean
Public WidthUserColumn As Integer
Public ShowSort As Boolean
Public BreiteSort As Integer
Public AlleMASichtbar As Boolean
Public PlanEveryDay As Boolean

Public DatenTage As Integer

' clsAbwesenheit
'lngAnzTageLevel1

' basAutoMail
'booMailBCC
'MailImmerAL
'MailAlleE4
'booMailICS
'booSendMail
'SMTP-Server

' basMain
'booCheckMinVersion
'MinVersion
'booE4SiehtKrankheit
'booSuperOL
'booShowSort
'booAlleMaSichtbar
'booPlanEveryDay
'strNoRun

' fromInfo
'ShowUserlist
'booOrgSlider
'WidthUserColumn
'lngBreiteSort
'booTrennerOrgs
'booNoRun

' clsData
'DatenTage
'TravellingRequestToE3
'NumDaysBackApprove
'
'
'
'
