VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private WithEvents frmI As frmInfo
Attribute frmI.VB_VarHelpID = -1

Public Sub frmI_CloseApplication()
    Main_Close_Application
End Sub

Public Sub CreateForm()
    Set frmI = New frmInfo
    frmI.Show
End Sub
