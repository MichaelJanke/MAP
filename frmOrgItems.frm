VERSION 5.00
Begin VB.Form frmOrgItems 
   Caption         =   "Organizational Items Editing"
   ClientHeight    =   9600
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11175
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
   ScaleHeight     =   9600
   ScaleWidth      =   11175
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox txtValue 
      Height          =   420
      Left            =   5160
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   4920
      Width           =   5895
   End
   Begin VB.ListBox lbOrg 
      Height          =   4860
      Left            =   240
      TabIndex        =   2
      Top             =   4440
      Width           =   4695
   End
   Begin VB.ListBox lbItem 
      Height          =   3660
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   10575
   End
   Begin VB.Label lblValue 
      Caption         =   "Please enter Value"
      Height          =   375
      Left            =   5160
      TabIndex        =   4
      Top             =   4440
      Width           =   5415
   End
   Begin VB.Label lblorgItems 
      Caption         =   "Item name"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "frmOrgItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_IdxOrg As Long, m_IdxItem As Long
Private m_cO As clsOrg

Private Sub Form_Load()
    Me.Caption = App.ExeName & " Items"
    ReadWindowPosition Me
    Setup_Form
End Sub
Private Sub Setup_Form()
    FillLbItems
    Me.txtValue.Visible = False
    Me.lblValue.Visible = False
End Sub
Private Sub FillLbItems()
    Dim cI As clsItem
    With lbItem
        .Clear
        For Each cI In g_db.colItem
            .AddItem cI.strItem & vbTab & "Default:" & cI.ValDefault & vbTab & "Desc:" & cI.strDescription
        Next
        .ListIndex = 0  ' auf ersten Eintrag setzen, Org füllen
    End With
End Sub
Private Sub FillLbOrg()
    ' Wenn ein Item angewählt ist, dann stelle für alle Orgs den Wert des Items dar. Sonst nur Org-Namen
    Dim strT As String, strV As String, iElement As Integer
    With lbOrg
        .Clear
        If m_IdxItem < 0 Then Exit Sub  ' kein Item angewählt

        For Each m_cO In g_db.colOrg
            strT = m_cO.strOrg
            If Me.lbItem.ListIndex <> -1 Then   ' Ein Item angewählt
                strV = g_db.GetOrgValue(m_cO.lngIdxOrg, GetTextBeforeTab(Me.lbItem))
                strT = strT & vbTab & ":" & strV
            End If
            .AddItem strT
            .ItemData(.NewIndex) = m_cO.lngIdxOrg   ' Merken, wenn geändert werden soll
        Next
        .ListIndex = -1
    End With
    
    Me.txtValue.Visible = False
    Me.lblValue.Visible = False
End Sub
Private Function GetTextBeforeTab(T As String) As String
    GetTextBeforeTab = Left(T, InStr(T, vbTab) - 1)
End Function

Private Sub Form_Resize()
    Me.lbItem.Height = Me.Height * 0.45
    Me.lbOrg.Height = Me.lbItem.Height
    Me.lbItem.Width = Me.Width - 1000
    Me.lbOrg.Width = Me.Width * 0.45
    Me.lbOrg.Top = Me.lbItem.Top + Me.lbItem.Height + 200
    Me.txtValue.Left = Me.lbOrg.Width + Me.lbOrg.Left + 200
    Me.lblValue.Left = Me.txtValue.Left
    Me.lblValue.Top = Me.lbOrg.Top
    Me.txtValue.Top = Me.lblValue.Top + Me.lblValue.Height + 200
End Sub

Private Sub lbOrg_Click()
    ' Org ist angeklickt -> Stelle mögliche Werte in cbo dar
    m_IdxOrg = lbOrg.ItemData(lbOrg.ListIndex)
    Set m_cO = g_db.colOrg("cOrg_" & m_IdxOrg)
    Me.txtValue.Text = g_db.GetOrgValue(m_cO.lngIdxOrg, GetTextBeforeTab(Me.lbItem))
    Me.txtValue.Visible = True
    Me.lblValue.Visible = True
End Sub

Private Sub lbItem_Click()
    ' Item ist angewählt -> Stelle Werte für die org-member dar
    m_IdxItem = lbItem.ItemData(lbItem.ListIndex)
    FillLbOrg
End Sub

'######################################################################################################
'##########  Entering Values   ########################################################################
Private Sub txtValue_Validate(Cancel As Boolean)
    ChangeValue txtValue.Text
End Sub
'######################################################################################################
' Value changed
Private Sub ChangeValue(ValItem As String)
    If g_db.WriteOrgItem(m_IdxOrg, GetTextBeforeTab(Me.lbItem), Me.txtValue) Then
        FillLbOrg
    End If
End Sub
Private Sub DeleteValue()
'    g_db.WriteOrgVal m_IdxOrg, m_IdxItem, m_Element, "", True
End Sub

'######################################################################################################
'####  cbo Management
Private Sub btnAdd_Click()
    Stop
End Sub

Private Sub btnDel_Click()
'    Stop
End Sub

