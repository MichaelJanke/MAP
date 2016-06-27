VERSION 5.00
Begin VB.Form frmOptions 
   Caption         =   "Options"
   ClientHeight    =   3105
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7635
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   7635
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CheckBox chkOrgSlider 
      Caption         =   "OrgSlider"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   3135
   End
   Begin VB.TextBox txtWidthSortColumn 
      Height          =   375
      Left            =   5640
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox txtWidthUserCol 
      Height          =   375
      Left            =   5640
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   840
      Width           =   1575
   End
   Begin VB.CheckBox chkPlanEveryDay 
      Caption         =   "PlanEveryDay"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   3135
   End
   Begin VB.CheckBox chkAlleMASichtbar 
      Caption         =   "AlleMASichtbar"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   3135
   End
   Begin VB.CheckBox chkShowSort 
      Caption         =   "Show Sort"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   3135
   End
   Begin VB.CheckBox chkShowUserList 
      Caption         =   "Show User List"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "WidthSortColumn"
      Height          =   375
      Left            =   3840
      TabIndex        =   8
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "WidthUserColumn"
      Height          =   375
      Left            =   3840
      TabIndex        =   6
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label labHeader 
      Caption         =   "Userdefinable Options for MAP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    FillForm
End Sub

Private Sub FillForm()
    Setval chkShowUserList, "ShowUserList"
    Setval chkShowSort, "booShowSort"
    Setval chkAlleMASichtbar, "booAlleMASichtbar"
    Setval chkPlanEveryDay, "PlanEveryDay"
    Setval chkOrgSlider, "booOrgSlider"
    Setval txtWidthUserCol, "WidthUserColumn"
    Setval txtWidthSortColumn, "lngBreiteSort"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckDiff chkShowUserList, "ShowUserList"
    CheckDiff chkShowSort, "booShowSort"
    CheckDiff chkAlleMASichtbar, "booAlleMASichtbar"
    CheckDiff chkPlanEveryDay, "PlanEveryDay"
    CheckDiff chkOrgSlider, "booOrgSlider"
    CheckDiff txtWidthUserCol, "WidthUserColumn"
    CheckDiff txtWidthSortColumn, "lngBreiteSort"
End Sub

Private Sub CheckDiff(ctl As Control, name As String)
    Dim cI As clsItem
    Set cI = New clsItem    ' vorsorglich eine neue Struktur anlegen
    cI.strItem = name       ' Namen kenne ich schon

    If TypeOf ctl Is CheckBox Then
        Dim b As Boolean
        b = CBool(g_db.GetItem(name))
        If ctl.Value = vbChecked And b = False Then
            cI.ValItem = True       ' Wert ist jetzt auch bekannt ...
            g_db.WriteUserOption cI ' ... schreiben
        ElseIf ctl.Value = vbUnchecked And b = True Then
            cI.ValItem = False
            g_db.WriteUserOption cI ' ... schreiben
        End If
    Else
        Dim l As Long
        l = CLng(g_db.GetItem(name))
        If l <> CLng(ctl.Text) Then
            cI.ValItem = ctl.Text
            g_db.WriteUserOption cI ' ... schreiben
        End If
    End If
End Sub
Private Sub Setval(ctl As Control, name As String)
    If TypeOf ctl Is CheckBox Then
        Dim b As Boolean
        b = CBool(g_db.GetItem(name))
        If b Then
            ctl.Value = vbChecked
        Else
            ctl.Value = vbUnchecked
        End If
    Else        ' Textbox
        ctl.Text = g_db.GetItem(name)
    End If
End Sub
