VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "sToolTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim frmTT As frmToolTip


Private Sub Class_Initialize()
Dim INI&, TOP&
Set frmTT = New frmToolTip
INI = Screen.Height
frmTT.TOP = INI
frmTT.Show
TOP = Screen.Height - frmTT.Height - 500
frmTT.Left = Screen.Width - frmTT.Width
For I = INI To TOP Step -10
    frmTT.TOP = I
    Refresh
    DoEvents
Next
End Sub
