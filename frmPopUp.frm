VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPopUp 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "PopUp"
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8910
   Icon            =   "frmPopUp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton btnClose 
      Cancel          =   -1  'True
      Default         =   -1  'True
      Height          =   735
      Left            =   8100
      Picture         =   "frmPopUp.frx":0442
      Style           =   1  'Grafisch
      TabIndex        =   2
      ToolTipText     =   "Formular verlassen"
      Top             =   180
      Width           =   675
   End
   Begin VB.TextBox Text 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6975
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "frmPopUp.frx":058C
      Top             =   180
      Width           =   7875
   End
   Begin MSFlexGridLib.MSFlexGrid FG 
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   180
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   12303
      _Version        =   393216
      ScrollTrack     =   -1  'True
      ScrollBars      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmPopUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim lngCol As Long, lngRow As Long
Dim lngWidth As Long, lngHeight As Long
    Me.Caption = g_db.GetString(1085) & " Info " & g_strProgramVersion
    
    Me.btnClose.ToolTipText = g_db.GetString(1107)
    If g_PopUpMode = "Text" Then
        Text.Visible = True
        FG.Visible = False
        Text.Text = g_InfoText
    Else
        Text.Visible = False
        FG.Visible = True
        FG.Redraw = False
        Me.Caption = g_InfoText
        FG.Rows = UBound(g_DataTabelle, 2) + 1
        FG.Cols = UBound(g_DataTabelle, 1) + 1
        lngWidth = 0:   lngHeight = 0
        
        ' Spaltenbreiten festlegen, Gesamtbreite berechnen und Daten eintragen
        For lngCol = 0 To FG.Cols - 1
            lngWidth = lngWidth + g_InfoTabelle(lngCol)
            FG.ColWidth(lngCol) = g_InfoTabelle(lngCol)
            For lngRow = 0 To FG.Rows - 1
                FG.TextMatrix(lngRow, lngCol) = g_DataTabelle(lngCol, lngRow)
            Next lngRow
        Next lngCol
        
        ' Zeilen mit Text/ohne Text
        For lngRow = 0 To FG.Rows - 1
            If FG.TextMatrix(lngRow, 0) <> "" Then
                FG.RowHeight(lngRow) = 285
            Else
                FG.RowHeight(lngRow) = 30
            End If
            lngHeight = lngHeight + FG.RowHeight(lngRow)
        Next lngRow
        
        ' FlexGrid Breite, Hoehe, Eigenschaften
        FG.Width = lngWidth + 120
        FG.Height = lngHeight + 120
        FG.ScrollBars = flexScrollBarNone
        Me.Height = FG.Height + 1065    '''
        If Me.Height > 9000 Then
            Me.Height = 9000
            FG.Height = Me.Height - 1065
            FG.Width = FG.Width + 600   ' Scrollbalken
            FG.ScrollBars = flexScrollBarVertical
        End If
        Me.Width = FG.Width + 1125
        If Me.Width < 6000 Then Me.Width = 6000     ' 9000
        btnClose.Left = Me.Width - 900      ' Width=675, Rechts=225, Links=225 -> 1125
        FG.Redraw = True
    End If
End Sub
