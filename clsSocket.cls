VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public SMTPHost As String
Public From As String
Public ReplyToAddress As String
Public BccRecipient As String
Public Recipient As String
Public CcRecipient As String
Public Delimiter As String
Public Subject As String
Public Message As String

Public WithEvents wSock As Winsock
Attribute wSock.VB_VarHelpID = -1
Private booConnected As Boolean
Private booSendComplete As Boolean
Private booDataArrival As Boolean
Private lngTimerStart As Long

Public Sub Send()
On Error GoTo Send_err
    booConnected = False
    Set wSock = New MSWinsockLib.Winsock
    wSock.RemoteHost = SMTPHost
    wSock.RemotePort = 25
    SocketConnect
    
    SendData "HELO " & wSock.LocalIP                'Identify to the mail server
    SendData "MAIL FROM: " & From                   'Identify who the mail is from
    SendData "RCPT TO: " & Left(Recipient & Delimiter, InStr(Recipient & Delimiter, Delimiter) - 1)     '1st recipient
    SendData "DATA"                                 'Tell the server we're ready to send the mail body and subject
    
    SendText "From: " & From
    SendText "To:" & Recipient
    If CcRecipient <> "" Then SendText "Cc:" & CcRecipient
    SendText "Bcc:" & BccRecipient & Delimiter & "Michael.Janke@mtu-online.com"
'    SendText "Bcc:" & BccRecipient
    SendText "SUBJECT: " & Subject
    SendText vbCrLf & Message                       'Send the mail body
    
    SendData "."                                    'End the mail body    'Wait for server to catch up
    SendData "QUIT"
Send_Exit:
    On Error Resume Next
    If Not wSock Is Nothing Then
        wSock.Close                                     'Close connection to the server
        DoEvents                                        'Befehl ist wichtig!
        Set wSock = Nothing                             'Winsock entfernen
    End If
    Exit Sub
Send_err:
    Log "Send:err:" & Err.Number & " Desc:" & Err.Description, True
    Resume Send_Exit
End Sub
Private Sub SendData(strText As String)    'Sende text und warte auf Antwort
    Log "#SendData:" & strText, False
    booDataArrival = False
    booSendComplete = False
    wSock.SendData strText & vbCrLf
    WaitForSendComplete
    WaitForAnswer
End Sub
Private Sub SendText(strText As String) ' Sende Text ohne Antwort
    Log "#SendText:" & strText, False
    booSendComplete = False
    wSock.SendData strText & vbCrLf
    WaitForSendComplete
End Sub

Private Sub SocketConnect()
    booDataArrival = False
    wSock.Connect
    Do Until wSock.State = 7 Or booConnected
        DoEvents
    Loop
    WaitForAnswer    ' Begrüßungsnachricht des Servers abwarten
End Sub
Private Sub WaitForAnswer()
    lngTimerStart = Timer
    Do Until booDataArrival Or Timer - lngTimerStart > 3
        DoEvents
    Loop
    If Timer - lngTimerStart > 3 Then Log "WaitForAnswer:Timeout", True
End Sub
Private Sub WaitForSendComplete()
    lngTimerStart = Timer
    Do Until booSendComplete Or Timer - lngTimerStart > 3
        DoEvents
    Loop
    If Timer - lngTimerStart > 3 Then Log "WaitForSendComplete:Timeout", True
End Sub

'######################################################################################################################
' Section EventHandling for WinSock
Private Sub wSock_Connect()
    Log ">>>Connect", False
    booConnected = True
End Sub
Private Sub wSock_Close()
    Log ">>>Close", False
End Sub

Private Sub wSock_DataArrival(ByVal bytesTotal As Long)
Dim Buffer As String
'Get incoming data, set that data to the Buffer variable
    wSock.GetData Buffer
    Log ">>>DataArrival:" & Buffer, False
    booDataArrival = True
End Sub
Private Sub wSock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'Error ocurred, display in Data textbox
    Log ">>>Error:Number" & Number & " Desc:" & Description, True
End Sub
Private Sub wSock_SendComplete()
'    Log ">>>SendComplete", False
    booSendComplete = True
End Sub
'######################################################################################################################
Private Sub Log(strLog As String, booErr As Boolean)
    Debug.Print Format(Now, "yyMMdd HHmmSS") & "." & Format((Timer - Fix(Timer)) * 1000, "000") & ": " & strLog
    If booErr Then MsgBox strLog
End Sub
