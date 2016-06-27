VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' aus DB
Public lngIdxUser As Long
Public strAccountname As String
Public strAccountname2 As String
Public strNachname As String
Public strVorname As String
Public strMailName As String
Public lngPN As Long
Public lngIdxOrg As Long
Public dtmFirst As Date
Public dtmLast As Date
Public strSortOrder As String   ' zum Sortieren 3
Public dtmDisclaimer As Date    ' Datum Einverständniserklärung
Public lngVerteiler As Long     ' Sichtbar für 1=Chef 2=Org

' aus Org - eigener oder übergeordneter
Public cOrg As clsOrg           ' Referenz zur Org

Public lngIdxChef As Long       ' als MA: Chef der Org, als Org-Chef: Chef der übergeordneten Org
Public lngIdxChef2 As Long
Public strSortLevel As String   '   zum Sortieren 2

' abgeleitet
Public lngUserLevel As OrgLevel ' User,IL,TL,AL,BL,RL
Public lngIdxRes As Long        ' Vererbt aus tblOrg
Public lngIdxBer As Long        ' Vererbt aus tblOrg
Public lngIdxAbt As Long        ' Vererbt aus tblOrg
Public lngIdxTeam As Long       ' Vererbt aus tblOrg
Public lngIdxInsel As Long      ' Vererbt aus tblOrg

' aus Org
Public booIsSek As Boolean      ' Benutzer hat SekStatus
Public booIsOrgChef As Boolean  ' Ist Chef seiner eigenen Organitsation
Public booIsOrgChef2 As Boolean ' Ist Stellvertretender Chef seiner eigenen Organitsation

Public colMA As Collection      ' MEINE MA
Public colFunctions As Collection ' meine Funktionen

Public lngAbstand As Long       ' berechneter Abstand dieses Benutzers zum g_CU
Public strSortKey As String

Public colOrg As New Collection ' in diesen Orgs bin ich Chef

Public Function Fullname() As String
    Fullname = strVorname & " " & strNachname & ", " & cOrg.strOrg
End Function
Public Function IsOrgChef(cOrg As clsOrg, booDep As Boolean) As Boolean
    ' cOrg  Organisation, um die es geht
    ' booDep Wird Stellvertreter auch als Chef gewertet?
    If cOrg.lngIdxChef = lngIdxUser Or booDep And cOrg.lngIdxChef2 = lngIdxUser Then
        IsOrgChef = True:  Exit Function   '########################################################
    End If
    IsOrgChef = False:  Exit Function   '########################################################
End Function
Public Function IsOrgSek(cOrg As clsOrg) As Boolean
    IsOrgSek = (cOrg.lngIdxSek = lngIdxUser Or cOrg.lngIdxSek2 = lngIdxUser)
End Function
Public Property Get IsSek() As Boolean
    IsSek = Me.booIsSek
End Property
Public Property Let IsSek(value As Boolean)
    Me.booIsSek = value
End Property
Public Property Get IsPrivileged() As Boolean
    IsPrivileged = Me.booIsOrgChef Or Me.booIsOrgChef2 Or Me.booIsSek
End Property

Public Function IsSekOf(cUs As clsUser) As Boolean
    IsSekOf = False
    If Not booIsSek Then Exit Function
    
    Dim cOrg As clsOrg
    Set cOrg = cUs.cOrg
    While True   ' Abbruchkriterium: oberste Org
        If cOrg.lngIdxSek = lngIdxUser Or cOrg.lngIdxSek2 = lngIdxUser Then ' Ich bin Org-Sek
            IsSekOf = True:        Exit Function   '########################################################
        End If
        If cOrg.lngIdxOrg <> cOrg.lngIdxTopOrg Then
            If g_db Is Nothing Then ' ich kann nicht hochklettern. Das hier muss reichen
                IsSekOf = False:  Exit Function   '########################################################
            Else
                Set cOrg = g_db.GetOrg(cOrg.lngIdxTopOrg)
            End If
        Else
            IsSekOf = False:  Exit Function   '########################################################
        End If
    Wend
    IsSekOf = False:  Exit Function   '########################################################
End Function

Public Function IsChefOf(cUs As clsUser, Optional booBoss As Boolean = True, Optional booStellvertreter As Boolean = True, Optional booDirekt As Boolean = True, Optional booIndirekt As Boolean = True) As Boolean
' Ist dieser Benutzer oder Stellvertreter Chef von cUs?
    If Not booDirekt And Not booIndirekt Then   ' eines von beiden muss sein
        IsChefOf = False:        Exit Function   '########################################################
    End If
    If Not booBoss And Not booStellvertreter Then   ' eines von beiden muss sein
        IsChefOf = False:        Exit Function   '########################################################
    End If
    
    ' Direkte MA, wenn cu1.IdxChef = cU1.IdUser True und noch nicht weitergelaufen wurde
    
    Dim booD As Boolean: booD = True    ' ohne Iteration gefunden -> Direkter MA
    Dim cU1 As clsUser
    Set cU1 = cUs       ' Startwert: zu untersuchender Benutzer, dann Baum aufsteigend Cheflinie
    While cU1.lngIdxChef <> cU1.lngIdxUser      ' Abbruchkriterium: oberster Chef ist sein eigener Chef
        If cU1.lngIdxChef = Me.lngIdxUser And booBoss Or _
           cU1.lngIdxChef2 = Me.lngIdxUser And booStellvertreter Then      ' User.Chef = Chef.ID
            If booDirekt And booD Or booIndirekt And Not booD Then
                IsChefOf = True:  Exit Function   '########################################################
            End If
        End If
        
        booD = False    ' im ersten Anlauf keim MA!
        
        If Not booIndirekt Then   ' indirekte zählen nicht -> return false
            IsChefOf = False:  Exit Function   '########################################################
        End If
        Set cU1 = g_db.GetUserByID(cU1.lngIdxChef)           ' Eine Hierarchieebene höher
    Wend
    IsChefOf = False:  Exit Function   '########################################################
End Function
Public Function IsAL() As Boolean
    IsAL = lngUserLevel = Abteilung And Me.IsOrgChef(Me.cOrg, False)
End Function

