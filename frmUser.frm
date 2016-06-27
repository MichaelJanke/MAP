VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmUser 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "User"
   ClientHeight    =   10785
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14520
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10785
   ScaleWidth      =   14520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin MSDataGridLib.DataGrid FG 
      Height          =   7335
      Left            =   120
      TabIndex        =   18
      Top             =   1080
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   12938
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1031
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1031
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton btnAddUser 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   120
      Picture         =   "frmUser.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   17
      ToolTipText     =   "Neuer Benutzer"
      Top             =   120
      Width           =   795
   End
   Begin VB.ComboBox cboOrg 
      Height          =   420
      Left            =   7920
      TabIndex        =   16
      Text            =   "Combo1"
      Top             =   8640
      Width           =   3495
   End
   Begin VB.TextBox strAccountname2 
      Height          =   420
      Left            =   2160
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   9120
      Width           =   3375
   End
   Begin VB.TextBox strSortOrder 
      Height          =   420
      Left            =   7920
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   10200
      Width           =   3375
   End
   Begin VB.TextBox strMailAddtress 
      Height          =   420
      Left            =   2160
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   10200
      Width           =   3375
   End
   Begin VB.TextBox strFirstName 
      Height          =   420
      Left            =   7920
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   9720
      Width           =   3375
   End
   Begin VB.TextBox strLastName 
      Height          =   420
      Left            =   2160
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   9720
      Width           =   3375
   End
   Begin VB.TextBox strAccountname 
      Height          =   420
      Left            =   2160
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   8640
      Width           =   3375
   End
   Begin VB.CommandButton btnDeleteUser 
      Cancel          =   -1  'True
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   1080
      Picture         =   "frmUser.frx":030A
      Style           =   1  'Grafisch
      TabIndex        =   2
      ToolTipText     =   "Benutzer löschen"
      Top             =   120
      Width           =   795
   End
   Begin VB.CommandButton btnSaveChanges 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   13440
      Picture         =   "frmUser.frx":0614
      Style           =   1  'Grafisch
      TabIndex        =   1
      ToolTipText     =   "save changes"
      Top             =   9840
      Width           =   795
   End
   Begin VB.CommandButton btnExit 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   13560
      Picture         =   "frmUser.frx":091E
      Style           =   1  'Grafisch
      TabIndex        =   0
      ToolTipText     =   "Programm verlassen"
      Top             =   120
      Width           =   795
   End
   Begin VB.Label lblOrg 
      Caption         =   "Org"
      Height          =   375
      Left            =   5760
      TabIndex        =   15
      Top             =   8640
      Width           =   1935
   End
   Begin VB.Label lblAccountName2 
      Caption         =   "Accountname2"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   9120
      Width           =   1935
   End
   Begin VB.Label lblSortOrder 
      Caption         =   "SortOrder"
      Height          =   375
      Left            =   5760
      TabIndex        =   11
      Top             =   10200
      Width           =   1935
   End
   Begin VB.Label lblMailAddress 
      Caption         =   "MailAddress"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   10200
      Width           =   1935
   End
   Begin VB.Label lblFirstName 
      Caption         =   "First name"
      Height          =   375
      Left            =   5760
      TabIndex        =   7
      Top             =   9720
      Width           =   1935
   End
   Begin VB.Label lblLastName 
      Caption         =   "Last name"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   9720
      Width           =   1935
   End
   Begin VB.Label lblAccountname 
      Caption         =   "Accountname"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   8640
      Width           =   1935
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub FG_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Debug.Print "RolColChange:Last:" & LastRow & "/" & LastCol & " New:" & FG.Row & "/" & FG.col
    
End Sub

Private Sub Form_Load()
    Me.Caption = App.ExeName & " UserManager " & g_strProgramVersion
    LoadData
End Sub

Private Sub LoadData()
Dim strDBName As String, strSystemDB As String
    strDBName = App.Path & "\" & App.ExeName & ".mdb"
    strSystemDB = App.Path & "\sys2k.mdw"
    strConnect = "Data Source=" & strDBName

Dim conn As ADODB.Connection
    Set conn = New ADODB.Connection
    conn.CursorLocation = adUseClient
    conn.Provider = "Microsoft.Jet.OLEDB.4.0"
    conn.Properties("Jet OLEDB:System database") = strSystemDB
    conn.Open strConnect, DBUSER, DBPASSWORD
    
Dim lOrg As Long
    lOrg = g_CU.lngIdxOrg
    
Dim strSQL As String
    strSQL = "SELECT * FROM tblUser ORDER BY lngIdxOrg, strSortOrder, strNachname"

Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    RS.Open strSQL, conn, adOpenDynamic, adLockOptimistic

'    Data1.DatabaseName = strDBName
'    Data1.RecordSource = "SELECT * FROM tblUser ORDER BY lngIdxChef, strSortOrder, strNachname"
'    Data1.Database = g_db.conn
'    Data1.RecordSource = ""
'    Set FG.DataSource = Data1
    
    Set FG.DataSource = RS
    
    FG.AllowAddNew = False
    FG.AllowDelete = False
    FG.AllowUpdate = False
    FG.RecordSelectors = True

    FG.Refresh
    FG.ColumnHeaders = True
    FG.DefColWidth = 0
    FG.EditActive = True
    FG.ScrollBars = dbgAutomatic
'    FG.Columns(0).Width = 1200
'    FG.Columns(1).Width = 1200
'    FG.Columns(2).Width = 1200
'    FG.Columns(3).Width = 1200
'    FG.Columns(4).Width = 1200
'    FG.Columns(5).Width = 1200
'    FG.Columns(6).Width = 1200
End Sub
Private Sub FillData()

End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' BUTTONS '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub btnExit_Click()
    Unload Me
End Sub
Private Sub btnAddUser_Click()
    Stop
End Sub
Private Sub btnDeleteUser_Click()
    Stop
End Sub
Private Sub btnSaveChanges_Click()
    Stop
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Inhalte  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboOrg_Change()
    Stop
End Sub
Private Sub strAccountname_Change()
    Stop
End Sub
Private Sub strAccountname2_Change()
    Stop
End Sub
Private Sub strFirstName_Change()
    Stop
End Sub
Private Sub strLastName_Change()
    Stop
End Sub
Private Sub strMailAddtress_Change()
    Stop
End Sub
Private Sub strSortOrder_Change()
    Stop
End Sub
