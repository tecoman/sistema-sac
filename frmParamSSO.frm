VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmParamSSO 
   Caption         =   "Parámtros S.S.O"
   ClientHeight    =   3780
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3780
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cmd 
      Caption         =   "&Aceptar"
      Height          =   420
      Index           =   1
      Left            =   3270
      TabIndex        =   7
      Top             =   3075
      Width           =   1170
   End
   Begin VB.CommandButton Cmd 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      Height          =   420
      Index           =   0
      Left            =   1995
      TabIndex        =   6
      Top             =   3075
      Width           =   1170
   End
   Begin MSDataListLib.DataCombo dtc 
      Height          =   315
      Index           =   0
      Left            =   405
      TabIndex        =   2
      Top             =   975
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dtc 
      Height          =   315
      Index           =   1
      Left            =   1470
      TabIndex        =   3
      Top             =   975
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dtc 
      Height          =   315
      Index           =   2
      Left            =   405
      TabIndex        =   4
      Top             =   2475
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dtc 
      Height          =   315
      Index           =   3
      Left            =   1470
      TabIndex        =   5
      Top             =   2475
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.Label lbl 
      Caption         =   $"frmParamSSO.frx":0000
      Height          =   795
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   1605
      Width           =   4170
   End
   Begin VB.Label lbl 
      Caption         =   "1.- Seleccione de la lista el Proveedor correspondiente al Seguro Social Obligatorio:"
      Height          =   510
      Index           =   0
      Left            =   210
      TabIndex        =   0
      Top             =   315
      Width           =   4200
   End
End
Attribute VB_Name = "frmParamSSO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstLocal(3) As New ADODB.Recordset

Private Sub cmd_Click(Index As Integer)
'variables locales
Select Case Index

    Case 0  'cancelar
        Unload Me
        Set frmParamSSO = Nothing
    
    Case 1  'aceptar
        cnnConexion.Execute "UPDATE ParamSSO SET CodPro='" & dtc(0) & "', CodGasto='" & _
        dtc(2) & "'"
        
End Select
End Sub

Private Sub dtc_Click(Index As Integer, Area As Integer)
'variables locales
If Index = 0 Then
    dtc(1) = dtc(0).BoundText
ElseIf Index = 2 Then
    dtc(3) = dtc(2).BoundText
End If
End Sub

Private Sub Form_Load()
'variables local
Dim strSql As String
Dim Cod0$, Cod2$

With FrmAdmin.objRst
    .MoveFirst
    If Not .EOF And Not .BOF Then
        .Find "Inactivo=False"
        If .EOF Then
            MsgBox "No existen registrados inmuebles activos", vbInformation, App.ProductName
            Me.Tag = 1
        End If
    Else
        MsgBox "No existen inmuebles registrados." & vbCrLf & "Debe registrar " & _
        "por lo menos un inmueble.", vbInformation, App.ProductName
        Me.Tag = 1
    End If
End With
If Me.Tag = "1" Then Exit Sub
'
rstLocal(0).Open "ParamSSO", cnnConexion, adOpenKeyset, adLockOptimistic, adCmdTable
If Not rstLocal(0).EOF And Not rstLocal(0).BOF Then
    Cod0 = rstLocal(0)("CodPro")
    Cod2 = rstLocal(0)("CodGasto")
End If
rstLocal(0).Close

'configura los controles dtc
'indice 0 muestra los códigos de los proveedores registrador
strSql = "SELECT * FROM Proveedores ORDER BY Codigo"
rstLocal(0).Open strSql, cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
dtc(0).ListField = "Codigo"
dtc(0).BoundColumn = "NombProv"
Set dtc(0).RowSource = rstLocal(0)
'indice 1 los lista por orden del nombre
strSql = "SELECT * FROM Proveedores ORDER BY NombProv"
rstLocal(1).Open strSql, cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
dtc(1).ListField = "NombProv"
dtc(1).BoundColumn = "Codigo"
Set dtc(1).RowSource = rstLocal(1)
'indice 2: muestra los gastos ordenador por código
'toma el catálogo de los gastos del primer inmueble registrado (activo)
strSql = "SELECT * FROM Tgastos IN '" & gcPath & FrmAdmin.objRst("Ubica") & "inm.mdb' ORDER BY " _
& "CodGasto"
rstLocal(2).Open strSql, cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
dtc(2).ListField = "CodGasto"
dtc(2).BoundColumn = "Titulo"
Set dtc(2).RowSource = rstLocal(2)
'indice 3: muestra los gastos ordenador por título
strSql = "SELECT * FROM Tgastos IN '" & gcPath & FrmAdmin.objRst("Ubica") & "inm.mdb' ORDER BY " _
& "Titulo"
rstLocal(3).Open strSql, cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
dtc(3).ListField = "Titulo"
dtc(3).BoundColumn = "CodGasto"
Set dtc(3).RowSource = rstLocal(3)
If Cod0 <> "" Then dtc(0) = Cod0: dtc_Click 0, 2
If Cod2 <> "" Then dtc(2) = Cod2: dtc_Click 2, 2

Me.Show vbModeless, FrmAdmin
End Sub

Private Sub Form_Unload(Cancel As Integer)
'variables locales
Dim I%

For I = 0 To 3
    rstLocal(I).Close
    Set rstLocal(I) = Nothing
Next

End Sub
