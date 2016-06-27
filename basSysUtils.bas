Attribute VB_Name = "basSysUtils"
Option Explicit

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Function CurrentUser() As String     ' Holt NT-UserName
    CurrentUser = Environ("USERNAME")
    CurrentUser = LCase(CurrentUser)
End Function

Public Function HostName() As String     ' Holt WorkstationName
    HostName = Environ("COMPUTERNAME")
End Function

Public Function ConvertToASCII(InString As String) As String
    On Error GoTo err_ConvertToASCII
    Dim i As Integer, OutString As String, c As Byte
    Dim ch As String
    OutString = ""
    For i = 1 To Len(InString)
        ch = Mid(InString, i, 1)
        If ch <= "z" Then
            OutString = OutString & ch
        Else
            Select Case ch
                Case "�"
                    OutString = OutString & "ae"
                Case "�"
                    OutString = OutString & "oe"
                Case "�"
                    OutString = OutString & "ue"
                Case "�"
                    OutString = OutString & "Ae"
                Case "�"
                    OutString = OutString & "Oe"
                Case "�"
                    OutString = OutString & "Ue"
                Case "�"
                    OutString = OutString & "ss"
            End Select
        End If
    Next
    ConvertToASCII = OutString
    Exit Function
err_ConvertToASCII:
    ConvertToASCII = InString
End Function

Public Function Liste(s As String, N As Integer) As String
    Dim i As Integer
    Liste = ""
    For i = 1 To N
        Liste = Liste & s
    Next
End Function
