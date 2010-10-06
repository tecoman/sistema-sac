VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmCtaInm 
   Caption         =   "Cuenta de Nómina Inmueble"
   ClientHeight    =   2640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5325
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2640
   ScaleWidth      =   5325
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2400
      Left            =   75
      TabIndex        =   8
      Top             =   60
      Width           =   5130
      Begin VB.CommandButton cmd 
         Cancel          =   -1  'True
         Caption         =   "&Cerrar"
         Height          =   495
         Index           =   0
         Left            =   3735
         TabIndex        =   0
         Top             =   1770
         Width           =   1215
      End
      Begin VB.CommandButton cmd 
         Caption         =   "&Guardar"
         Height          =   495
         Index           =   1
         Left            =   2370
         TabIndex        =   7
         Top             =   1770
         Width           =   1215
      End
      Begin MSDataListLib.DataCombo dtc 
         Height          =   315
         Index           =   1
         Left            =   1095
         TabIndex        =   3
         ToolTipText     =   "Nombre del Inmueble"
         Top             =   795
         Width           =   3885
         _ExtentX        =   6853
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Nombre"
         BoundColumn     =   "CodInm"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc 
         Height          =   315
         Index           =   2
         Left            =   1095
         TabIndex        =   5
         ToolTipText     =   "Banco"
         Top             =   1185
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "NombreBanco"
         BoundColumn     =   "NumCuenta"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc 
         Height          =   315
         Index           =   3
         Left            =   2700
         TabIndex        =   6
         ToolTipText     =   "Nº de Cuenta"
         Top             =   1185
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "NumCuenta"
         BoundColumn     =   "NombreBanco"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc 
         Height          =   315
         Index           =   0
         Left            =   1095
         TabIndex        =   2
         ToolTipText     =   "Código del Inmueble"
         Top             =   405
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "CodInm"
         BoundColumn     =   "Nombre"
         Text            =   ""
      End
      Begin VB.Label lbl 
         Caption         =   "C&uenta:"
         Height          =   285
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   1215
         Width           =   1215
      End
      Begin VB.Label lbl 
         Caption         =   "&Inmueble:"
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   435
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmCtaInm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstCta(2) As New ADODB.Recordset
Dim IDCuenta As String
'
Private Sub cmd_Click(Index As Integer)
'variables locales
Dim ventana As Form
Dim ID As Integer
'
Select Case Index
    Case 0
        Unload Me
        Set frmCtaInm = Nothing
    
    Case 1  'guardar cambios
        If Error Then Exit Sub
        With rstCta(2)
            .MoveFirst
            .Find "NumCuenta=" & dtc(3)
            If Not .EOF Then
                ID = .Fields("IDBanco")
                rstCta(0).Update "CtaInm", .Fields("IDCuenta")
                For Each ventana In Forms
                    If ventana.Name = "FrmFichaEmp" Then
                        FrmFichaEmp.adoEmp(3).Refresh
                        FrmFichaEmp.dtcEmp(3) = dtc(2)
                        FrmFichaEmp.dtcEmp(3).Tag = ID
                        Exit For
                    End If
                Next
                Call rtnBitacora("Asignar Cuenta Inm.:'" & dtc(0) & "/" & dtc(3) & "'")
                MsgBox "Registro Actualizado"
            Else
                MsgBox "Imposible guardar este registro, consulte al administrador del sistema", _
                vbInformation, App.ProductName
            End If
            '
        End With
'
End Select
'
End Sub

Private Sub dtc_Click(Index As Integer, Area As Integer)
'variables locales
Dim Cadena_Conexion As String

'
If Area = 2 Then
Select Case Index
    Case 0, 1   'codigo - nombre de inmueble
        dtc(IIf(Index = 0, 1, 0)) = dtc(IIf(Index = 0, 0, 1)).BoundText
        dtc(2) = ""
        dtc(3) = ""
        With rstCta(0)
            .MoveFirst
            .Find "CodInm='" & dtc(0) & "'"
            IDCuenta = IIf(IsNull(.Fields("CtaInm")), "", .Fields("CtaInm"))
            If .Fields("Caja") = sysCodCaja Then
                Cadena_Conexion = cnnOLEDB & gcPath & "\" & sysCodInm & "\inm.mdb"
            Else
                Cadena_Conexion = cnnOLEDB & gcPath & .Fields("Ubica") & "inm.mdb"
            End If
            If rstCta(2).State = 1 Then rstCta(2).Close
            rstCta(2).Open "SELECT Bancos.NombreBanco, Bancos.IDBanco,Cuentas.NumCuenta,Cuentas.IDCuenta FROM " _
            & "Bancos INNER JOIN Cuentas ON Bancos.IDBanco = Cuentas.IDBanco;", Cadena_Conexion, adOpenKeyset, adLockOptimistic, adCmdText
            If Not IDCuenta = "" Then
                With rstCta(2)
                    .Filter = "IDCuenta=" & IDCuenta
                    If Not (.EOF And .BOF) Then
                        '.MoveFirst
                        '.Find "IDCuenta=" & IDCuenta
                        'If Not (.EOF And .BOF) Then
                        dtc(2) = .Fields("NumCuenta")
                        dtc(3) = .Fields("NombreBanco")
                        'End If
                    End If
                End With
            End If
            Set dtc(2).RowSource = rstCta(2)
            Set dtc(3).RowSource = rstCta(2)
            
        End With
        
    Case 3, 2   'banco - cuenta
        dtc(IIf(Index = 2, 3, 2)) = dtc(IIf(Index = 2, 2, 3)).BoundText
        
End Select

End If

End Sub

Private Sub Form_Activate()
If blnModif Then
    
End If
End Sub

Private Sub Form_Load()
'variables locales
'----------------------
rstCta(0).Open "SELECT * FROM Inmueble ORDER BY CodInm", cnnConexion, adOpenKeyset, _
adLockOptimistic, adCmdText
rstCta(1).Open "SELECT * FROM Inmueble ORDER BY Nombre", cnnConexion, adOpenKeyset, _
adLockOptimistic, adCmdText
Set dtc(0).RowSource = rstCta(0)
Set dtc(1).RowSource = rstCta(1)
'---------------------
Call CenterForm(frmCtaInm)
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
For I = 0 To 2
    rstCta(I).Close
    Set rstCta(I) = Nothing
Next
End Sub

Private Function Error() As Boolean
For I = 0 To 3
    If dtc(I) = "" Then
        Error = MsgBox("Falta '" & dtc(I).ToolTipText & "'")
        Exit For
    End If
Next
End Function
