VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmVacaSup 
   Caption         =   "Suplencia Vacaciones"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6420
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4215
   ScaleWidth      =   6420
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Fra 
      Caption         =   "Complete la siguiente información:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3210
      Left            =   240
      TabIndex        =   2
      Top             =   195
      Width           =   5955
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   1
         Left            =   1755
         TabIndex        =   12
         Top             =   2595
         Width           =   3885
      End
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   0
         Left            =   1755
         TabIndex        =   11
         Top             =   2100
         Width           =   3885
      End
      Begin MSDataListLib.DataCombo Dtc 
         Height          =   315
         Index           =   0
         Left            =   555
         TabIndex        =   9
         Top             =   1020
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo Dtc 
         Height          =   315
         Index           =   1
         Left            =   1755
         TabIndex        =   10
         Top             =   1020
         Width           =   3885
         _ExtentX        =   6853
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label lbl 
         Caption         =   "Cédula de Identidad:"
         Height          =   480
         Index           =   5
         Left            =   690
         TabIndex        =   8
         Top             =   2505
         Width           =   960
      End
      Begin VB.Label lbl 
         Caption         =   "Nombre:"
         Height          =   285
         Index           =   4
         Left            =   690
         TabIndex        =   7
         Top             =   2130
         Width           =   1215
      End
      Begin VB.Label lbl 
         Caption         =   "Nombre"
         Height          =   285
         Index           =   3
         Left            =   2160
         TabIndex        =   6
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lbl 
         Caption         =   "Cod."
         Height          =   285
         Index           =   2
         Left            =   690
         TabIndex        =   5
         Top             =   735
         Width           =   480
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Datos del Suplente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   285
         TabIndex        =   4
         Top             =   1785
         Width           =   1635
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Datos del Empleado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   285
         TabIndex        =   3
         Top             =   420
         Width           =   1710
      End
   End
   Begin VB.CommandButton Cmd 
      Caption         =   "&Procesar"
      Height          =   495
      Index           =   1
      Left            =   4920
      TabIndex        =   1
      Top             =   3585
      Width           =   1215
   End
   Begin VB.CommandButton Cmd 
      Caption         =   "&Cerrar"
      Height          =   495
      Index           =   0
      Left            =   3615
      TabIndex        =   0
      Top             =   3585
      Width           =   1215
   End
End
Attribute VB_Name = "frmVacaSup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstEmp As New ADODB.Recordset
Dim m As Double
Private Sub cmd_Click(Index As Integer)
Select Case Index

    Case 0  'cerrar ventana
        Unload Me
        Set frmVacaSup = Nothing
        
    Case 1  '
        Call Procesar
End Select
End Sub

Private Sub dtc_Click(Index As Integer, Area As Integer)
'variables locales
If Area = 2 Then
If Index = 0 Then
    dtc(1) = dtc(0).BoundText
    If dtc(1) = dtc(0) Then
        dtc(1) = ""
    Else
        m = Total
    End If
Else
    dtc(0) = dtc(1).BoundText
    If dtc(0) = dtc(1) Then
        dtc(0) = ""
    Else
        m = Total
    End If
    '
End If
End If
'
End Sub

Private Sub dtc_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If Index = 0 Then Call Validacion(KeyAscii, "0123456789")
If KeyAscii = 13 Then Call dtc_Click(Index, 2)
End Sub

Private Sub Form_Load()
''WHERE CodEmp IN " _
& "(SELECT CodEmp FROM Nom_vaca WHERE " _
& "Incorp )"
rstEmp.Open "SELECT *, Apellidos & ' ' & Nombres as Nom FROM Emp", cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
'
Set dtc(0).RowSource = rstEmp
Set dtc(1).RowSource = rstEmp
dtc(0).ListField = "CodEmp"
dtc(1).ListField = "Nom"
dtc(0).BoundColumn = "Nom"
dtc(1).BoundColumn = "CodEmp"
Call CenterForm(frmVacaSup)
'
'
End Sub

Private Sub Form_Unload(Cancel As Integer)
'variables locales
rstEmp.Close
Set rstEmp = Nothing
End Sub

Private Sub Procesar()
'variables locales
Dim strSql$, NDoc$, Cargo$, Inmueble$, CInm$, CGast$, strDesde$, strHasta$, lngSueldo&
Dim rstPro As New ADODB.Recordset
Dim rpReporte As ctlReport
Dim cNeto As Currency
'
'valida los campos mínimos requeridos
If Trim(Txt(0)) = "" Or Trim(Txt(1)) = "" Then
    MsgBox "Introduzca los datos del suplente", vbInformation, App.ProductName
    Exit Sub
End If
'
With rstPro

    strSql = "SELECT TOP 1 Emp.CodEmp, Emp_Cargos.NombreCargo, Emp.CodInm, Inmueble.Nombre" _
    & ", Emp.CodGasto, Nom_Vaca.Desde, Nom_Vaca.Hasta, (Nom_Vaca.Dias_Vaca " _
    & ") * (SueldoM / 30) as Neto, Emp.Sueldo + ((Emp.Sueldo * Emp.BonoNoc) /100) " _
    & "as SueldoM FROM (Inmueble INNER JOIN " _
    & "(Emp_Cargos INNER JOIN Emp ON Emp_Cargos.CodCargo = Emp.CodCargo) " _
    & "ON Inmueble.CodInm = Emp.CodInm) INNER JOIN Nom_Vaca ON Emp.CodEmp = " _
    & "Nom_Vaca.CodEmp WHERE Nom_Vaca.CodEmp=" & dtc(0) & " ORDER BY Nom_Vaca.Hasta DESC"
    
    
    .Open strSql, cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
    
    If Not .EOF And Not .BOF Then
        '   selecciona la información del empleado
        Cargo = !NombreCargo
        Inmueble = !Nombre & " (" & !CodInm & ")"
        CInm = !CodInm
        CGast = !codGasto
        strDesde = !Desde
        strHasta = !Hasta
        lngSueldo = !SueldoM
        '
    Else
        MsgBox "Imposible procesar esta suplencia. Consule al administrador del sistema." & vbCrLf _
        & vbCrLf & "No se encuentra el código del empleado en la tabla 'vacaciones'", vbInformation, _
        App.ProductName
        Exit Sub
    End If
    '
    cNeto = !Neto
    NDoc = Registro_Proveedor(Txt(0), True, "V" & !CodEmp, _
    "CANC.SUPLENCIA DE VACACIONES " & Cargo & ". " & Inmueble, cNeto, CInm)
    .Close
    
End With
    '   imprime el soporte de la suplencia
    If NDoc <> "" Then
        Set rpReporte = New ctlReport
        Dim ToWords As New clsNum2Let
        ToWords.Moneda = "Bs."
        ToWords.Numero = cNeto
        
        With rpReporte
            '
            .Reporte = gcReport + "nom_sup.rpt"
            .OrigenDatos(0) = gcPath & "\sac.mdb"
            .OrigenDatos(1) = gcPath & "\sac.mdb"
            .OrigenDatos(2) = gcPath & "\sac.mdb"
            .OrigenDatos(3) = gcPath & "\sac.mdb"
            .Formulas(0) = "NomSup='" & Txt(0) & "'"
            .Formulas(1) = "ciSup=" & Txt(1)
            .Formulas(2) = "Nfact='" & NDoc & "'"
            .Formulas(3) = "NEmp=" & dtc(0)
            .Formulas(4) = "Desde=date(" & Format(strDesde, "yyyy,mm,dd") & ")"
            .Formulas(5) = "Hasta=date(" & Format(strHasta, "yyyy,mm,dd") & ")"
            .Formulas(6) = "cLetras='" & ToWords.ALetra & "'"
            '.SelectionFormula = "{Emp.CodEmp}=" & dtc(0)
            .Salida = crPantalla
            .TituloVentana = "Suplencia de Vacaciones"
            '.CopiesToPrinter = 3
            .Imprimir
            
'            If errLocal <> 0 Then
'                Call rtnBitacora("Error al imprimir la constacia de suplencia de vacaciones.")
'                MsgBox .LastErrorString, vbCritical, "Error " & .LastErrorNumber
'            Else
'                Call rtnBitacora("Impreso Reporte Suplencia de Vacaciones Ok.")
                cnnConexion.Execute "UPDATE Nom_Vaca SET ciSup='" & Txt(1) & "', NomSup='" & Txt(0) & _
                "', SueldoSup='" & lngSueldo & "', UsuarioSup='" & gcUsuario & "', FechaSup=Date(), " _
                & "NFactSup='" & NDoc & "' WHERE CodEmp=" & dtc(0) & " AND Desde=#" _
                & Format(strDesde, "mm/dd/yyyy") & "# AND Hasta=#" & Format(strHasta, "mm/dd/yyyy") & "#"
                Call rtnBitacora("Actualizado Registro Tabla Vacacione Emp: " & dtc(0) & " Desde: " & _
                strDesde & " Hasta: " & strHasta)
                Unload Me
                Set frmVacaSup = Nothing
'            End If
    
        End With
        
    Else
        MsgBox "Ocurrio un error durante el proceso. Consulte al administrador del sistema", _
        vbInformation, App.ProductName
    End If

Set rstPro = Nothing

End Sub

Function Total()
'variables locals
Dim rstlocal As New ADODB.Recordset
'
rstlocal.Open "SELECT * FROM Nom_Vaca WHERE CodEmp=" & dtc(0) & " AND Incorpora" _
& ">=date()", cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
'
If Not rstlocal.EOF And Not rstlocal.BOF Then

    Total = (rstlocal!Dias_Vaca + rstlocal!Dias_Feri) * (rstlocal!SueldoMensual / 30)
    cmd(1).Enabled = True
    
Else

    MsgBox "No se puede establecer el monto de las vacaciones", vbInformation, _
    App.ProductName
    cmd(1).Enabled = False
    
End If
'
End Function


Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
'
If Index = 0 Then
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
ElseIf Index = 1 Then
    Call Validacion(KeyAscii, "0123456789")
End If

'
End Sub
