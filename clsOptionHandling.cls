VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOptionHandling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Static Function ReadOptions() As Boolean
    op.AlleMASichtbar = basFormUtils.GetRegData("AlleMASichtbar")
    op.BreiteSort = basFormUtils.GetRegData("BreiteSort")
    op.OrgSlider = basFormUtils.GetRegData("OrgSlider")
    op.PlanEveryDay = basFormUtils.GetRegData("PlanEveryDay")
    op.ShowSort = basFormUtils.GetRegData("ShowSort")
    op.ShowUserList = basFormUtils.GetRegData("ShowUserList")
    op.WidthUserColumn = basFormUtils.GetRegData("WidthUserColumn")
End Function
Private Function SetValue(Name As String) As Boolean
    Dim var As Variant
    var = basFormUtils.GetRegData(Name)
    If var <> -1 Then   ' steht was drin
        
    End If
End Function
