Attribute VB_Name = "basFormUtils"
Option Explicit
' API-Prototyp für den Tausendsassa SystemParametersInfo:
Private Declare Function SystemParametersInfo _
  Lib "user32.dll" Alias "SystemParametersInfoA" ( _
  ByVal SPI_Action As Long, _
  ByVal uiParam As Long, _
  ByRef pvParam As Any, _
  ByVal fWinIni As Long _
  ) As Long

' SPI_GETWORKAREA bewegt SystemParametersInfo dazu, im
' Parameter pvParam die nutzbare Arbeitsfläche des Desktops
' in einer RECT-Struktur in pvParam zurückzugeben:
Private Const SPI_GETWORKAREA As Long = 48&

Public Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Public Sub fillListBoxFromCollection(objListBox As ListBox, col As Collection)
    objListBox.Clear
    Dim i As Integer
    For i = 1 To col.Count
        objListBox.AddItem col.Item(i)
    Next
End Sub

Public Sub ReadWindowPosition(f As Form, Optional booPositionOnly As Boolean = False)
Dim l As Long
    l = GetScreenItem(f.name, "Top"):     If l <> -1 Then f.Top = l
    l = GetScreenItem(f.name, "Left"):    If l <> -1 Then f.Left = l
    If Not booPositionOnly Then
        l = GetScreenItem(f.name, "Width"):  If l <> -1 Then f.Width = l: If f.Width < 100 Then f.Width = 3000
        l = GetScreenItem(f.name, "Height"): If l <> -1 Then f.Height = l: If f.Height < 100 Then f.Height = 3000
    End If
    l = GetScreenItem(f.name, "WindowState"):  If l <> -1 Then f.WindowState = l
    
    f.Caption = f.Caption & " " & basGlobals.g_strProgramVersion

    ' Wenn der linke obere Punkt nicht in der WorkArea liegt, dann in den sichtbaren Bereich verschieben
    Dim wr As RECT  ' verfügbare Arbeitsfläche
    If GetWorkArea(wr) Then
        ' form.Top Left etc in Twips, WorkArea in Pixel
        wr.Top = wr.Top * Screen.TwipsPerPixelY    ' spart das spätere Umrechnen
        wr.Bottom = wr.Bottom * Screen.TwipsPerPixelY
        wr.Left = wr.Left * Screen.TwipsPerPixelX
        wr.Right = wr.Right * Screen.TwipsPerPixelX
        
'        MsgBox "BS: TBLR:" & wr.Top & "/" & wr.Bottom & "/" & wr.Left & "/" & wr.Right & vbCrLf & _
'               "FM: THLW:" & f.Top & "/" & f.Height & "/" & f.Left & "/" & f.Width
        
        If f.Top < wr.Top Then f.Top = wr.Top
        If f.Left < wr.Left Then f.Left = wr.Left
        If f.Top + f.Height > wr.Bottom - wr.Top Then   ' zu weit unten? zu groß?
            If f.Height < wr.Bottom - wr.Top Then   ' kann hochgeschoben werden
                f.Top = wr.Bottom - f.Height
            Else    ' f.Height ist größer als Screen -> kleine machen
                f.Top = wr.Top  ' ganz oben
                f.Height = wr.Bottom - wr.Top    ' maximale Größe
            End If
        End If
        If f.Left + f.Width > wr.Right Then     ' zu weit rechts? zu breit?
            If f.Width < wr.Right - wr.Left Then ' kann nach links geschoben werden
                f.Left = wr.Right - f.Width
''            Else
''                f.Left = wr.Left
''                f.Width = wr.Right - wr.Left
            End If
        End If
'        MsgBox "BS: TBLR:" & wr.Top & "/" & wr.Bottom & "/" & wr.Left & "/" & wr.Right & vbCrLf & _
'               "FM: THLW:" & f.Top & "/" & f.Height & "/" & f.Left & "/" & f.Width
    End If
End Sub
Public Sub SaveWindowPosition(f As Form)
    SetScreenItem f.name, "Top", f.Top
    SetScreenItem f.name, "Left", f.Left
    SetScreenItem f.name, "Width", f.Width
    SetScreenItem f.name, "Height", f.Height
    SetScreenItem f.name, "WindowState", f.WindowState
End Sub

Public Function GetScreenItem(strScreen As String, strItem As String) As Long
    GetScreenItem = GetSetting(App.Title, strScreen, strItem, -1)
End Function
Public Static Sub SetScreenItem(strScreen As String, strItem As String, lngItem As Long)
    SaveSetting App.Title, strScreen, strItem, lngItem
End Sub

Public Static Function GetRegData(strDataName As String) As Variant
    GetRegData = GetSetting(App.Title, "Data", strDataName, -1)
End Function
Public Static Sub SetRegData(strDataName As String, varData As Variant)
    SaveSetting App.Title, "Data", strDataName, varData
End Sub
Public Static Function CleanRegData(strDataName As String) As Boolean
    DeleteSetting App.Title, "Data", strDataName
End Function

' Hier die zugehörige RECT-Struktur:
Public Function GetWorkArea(udtRect As RECT) As Boolean
  ' Ermittlung der verfügbaren Desktop-Arbeitsfläche:
  Call SystemParametersInfo(SPI_GETWORKAREA, 0&, udtRect, 0&)
'  ' Ausgabe des Ergebnisses
'  MsgBox "Position und Größe des Arbeitsbereichs:" & vbNewLine & _
'   "Links:  " & CStr(udtRect.Left) & vbNewLine & _
'   "Oben:   " & CStr(udtRect.Top) & vbNewLine & _
'   "Rechts: " & CStr(udtRect.Right) & vbNewLine & _
'   "Unten:  " & CStr(udtRect.Bottom) & vbNewLine, _
'   vbOKOnly + vbInformation, "Desktop-Arbeitsfläche"
   GetWorkArea = True
End Function

