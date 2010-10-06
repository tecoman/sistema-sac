VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmLibroBanco 
   Caption         =   "Libro Diario"
   ClientHeight    =   1725
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   3210
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   1725
   ScaleWidth      =   3210
   ShowInTaskbar   =   0   'False
   Tag             =   "1"
   WindowState     =   2  'Maximized
   Begin VB.Frame fraLBanco 
      BorderStyle     =   0  'None
      Height          =   1515
      Index           =   2
      Left            =   555
      TabIndex        =   5
      Top             =   6645
      Width           =   10470
      Begin VB.Frame fraLBanco 
         Caption         =   "Filtar por:"
         Height          =   1215
         Index           =   0
         Left            =   0
         TabIndex        =   9
         Top             =   225
         Width           =   7620
         Begin VB.CheckBox chk 
            Caption         =   "Cheques Emitidos"
            Height          =   315
            Index           =   4
            Left            =   5850
            TabIndex        =   14
            Tag             =   "CHQ"
            Top             =   255
            Width           =   1635
         End
         Begin VB.CheckBox chk 
            Caption         =   "Cheques(Ingresos)"
            Height          =   315
            Index           =   3
            Left            =   4065
            TabIndex        =   13
            Tag             =   "CHE"
            Top             =   255
            Width           =   1635
         End
         Begin VB.CheckBox chk 
            Caption         =   "Efectivo"
            Height          =   315
            Index           =   2
            Left            =   3075
            TabIndex        =   12
            Tag             =   "EFE"
            Top             =   255
            Width           =   1380
         End
         Begin VB.CheckBox chk 
            Caption         =   "Transferencias"
            Height          =   315
            Index           =   1
            Left            =   1590
            TabIndex        =   11
            Tag             =   "TRA"
            Top             =   240
            Width           =   1380
         End
         Begin VB.CheckBox chk 
            Caption         =   "Depósitos"
            Height          =   315
            Index           =   0
            Left            =   375
            TabIndex        =   10
            Tag             =   "DEP"
            Top             =   255
            Width           =   1215
         End
         Begin MSMask.MaskEdBox MskFecha 
            Bindings        =   "frmLibroBanco.frx":0000
            Height          =   315
            Index           =   0
            Left            =   1005
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   705
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   12
            Format          =   "dd/MM/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskFecha 
            Bindings        =   "frmLibroBanco.frx":0022
            DataField       =   "FRecep"
            Height          =   315
            Index           =   1
            Left            =   3075
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   705
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   12
            Format          =   "dd/MM/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Aplicar Filtro »»"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   4
            Left            =   4530
            MouseIcon       =   "frmLibroBanco.frx":0044
            MousePointer    =   99  'Custom
            TabIndex        =   20
            Top             =   765
            Width           =   1080
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Hasta:"
            Height          =   330
            Index           =   3
            Left            =   2475
            TabIndex        =   19
            Top             =   750
            Width           =   930
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Desde:"
            Height          =   330
            Index           =   2
            Left            =   375
            TabIndex        =   18
            Top             =   750
            Width           =   930
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Desmarcar todo »»"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   5895
            MouseIcon       =   "frmLibroBanco.frx":0196
            MousePointer    =   99  'Custom
            TabIndex        =   15
            Top             =   765
            Width           =   1350
         End
      End
      Begin VB.CommandButton cmdLBanco 
         Caption         =   "Imprimir"
         Height          =   1020
         Index           =   0
         Left            =   7965
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   330
         Width           =   1170
      End
      Begin VB.CommandButton cmdLBanco 
         Caption         =   "Salir"
         Height          =   1020
         Index           =   1
         Left            =   9150
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   330
         Width           =   1170
      End
      Begin VB.CommandButton cmdLBanco 
         Caption         =   "Actualizar Libro"
         Height          =   1020
         Index           =   2
         Left            =   7965
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   330
         Visible         =   0   'False
         Width           =   1170
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridLBanco 
      Height          =   5295
      Left            =   615
      TabIndex        =   0
      Tag             =   "1000|3000|500|1200|1400|1400|1400"
      Top             =   1245
      Width           =   10320
      _ExtentX        =   18203
      _ExtentY        =   9340
      _Version        =   393216
      Cols            =   7
      RowHeightMin    =   100
      BackColorFixed  =   -2147483646
      ForeColorFixed  =   -2147483639
      BackColorBkg    =   -2147483636
      GridColor       =   -2147483633
      ScrollTrack     =   -1  'True
      GridLineWidthFixed=   2
      FormatString    =   "Fecha |Concepto|Tipo|Número|Debe|Haber|Saldo"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontWidthFixed  =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   7
      _Band(0).GridLineWidthBand=   1
   End
   Begin VB.Frame fraLBanco 
      Height          =   915
      Index           =   1
      Left            =   630
      TabIndex        =   1
      Top             =   165
      Width           =   10290
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   8730
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "0,00"
         Top             =   300
         Width           =   1215
      End
      Begin MSDataListLib.DataCombo dtcCuentas 
         DataField       =   "NumCuenta"
         Height          =   315
         Index           =   0
         Left            =   3360
         TabIndex        =   2
         Top             =   360
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "NumCuenta"
         BoundColumn     =   "NombreBanco"
         Text            =   "DataCombo1"
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
      Begin MSDataListLib.DataCombo dtcCuentas 
         DataField       =   "NombreBanco"
         Height          =   315
         Index           =   1
         Left            =   945
         TabIndex        =   3
         Top             =   360
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "NombreBanco"
         BoundColumn     =   "NumCuenta"
         Text            =   "DataCombo1"
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
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo Inicio:"
         Height          =   330
         Index           =   5
         Left            =   7665
         TabIndex        =   21
         Top             =   360
         Width           =   930
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta:"
         Height          =   330
         Index           =   0
         Left            =   105
         TabIndex        =   4
         Top             =   345
         Width           =   930
      End
   End
End
Attribute VB_Name = "frmLibroBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstLBanco(1) As New ADODB.Recordset
Dim cnnLBanco As New ADODB.Connection

Private Sub chk_Click(Index As Integer)
'opciones de filtrado
If chk(Index).Value = vbChecked Then Call muestra_libro
End Sub

Private Sub cmdLBanco_Click(Index As Integer)
Select Case Index
    Case 0: Call Print_report
    Case 1: Unload Me
    Case 2: Call rtnRefrescar
End Select
End Sub

Private Sub DtcCuentas_Click(Index As Integer, Area As Integer)
If Area = 2 And DtcCuentas(Index) <> "" Then
    If Index = 0 Then DtcCuentas(1) = DtcCuentas(0).BoundText
    If Index = 1 Then DtcCuentas(0) = DtcCuentas(1).BoundText
    Call muestra_libro
End If
End Sub

Private Sub Form_Load()
'
Dim strSQL$    'variables locales
cnnLBanco.Open cnnOLEDB + mcDatos   'abre la conexón al orígen de datos
'
strSQL = "SELECT Cuentas.*, Bancos.NombreBanco FROM Bancos INNER JOIN " & _
        "Cuentas ON Bancos.IDBanco" _
        & "= Cuentas.IDBanco;"
rstLBanco(1).Open strSQL, cnnLBanco, adOpenStatic, adLockReadOnly, adCmdText
'
Set DtcCuentas(0).RowSource = rstLBanco(1)
Set DtcCuentas(1).RowSource = rstLBanco(1)
'
'strSQL = "SELECT Libro_Banco.*, Cuentas.NumCuenta, Bancos.NombreBanco " & _
'        "FROM Bancos INNER JOIN (Cuentas INNER JOIN Libro_Banco " & _
'        "ON Cuentas.IDCuenta = " & _
'        "Libro_Banco.IDCuenta) ON Bancos.IDBanco " _
'        & "= Cuentas.IDBanco "
'
'Set rstLBanco(0) = ModGeneral.ejecutar_procedure("procLibroBanco", Me.dtcCuentas(0), _
'getIdCuenta(dtcCuentas(0)), gcCodInm)
'With rstLBanco(0)
'    .CursorLocation = adUseClient
'    .Open strSQL, cnnLBanco, adOpenKeyset, adLockOptimistic, adCmdText
'    .Filter = "Fecha > #31/12/2002#"
'    .Sort = "Fecha, IdCheque"
    Call Config_Grid    'configura la presentación del grid
'End With
cnnConexion.Execute "procActulizaTDFCheques " & gcCodInm & ", 'TRANSFERENCIA'"
'ModGeneral.ejecutar_procedure "procActulizaTDFCheques", gcCodInm, "TRANSFERENCIA"
cnnConexion.Execute "procActulizaTDFCheques " & gcCodInm & ", 'DEPOSITO'", N
'Carga las imagenes de los botones de comando
cmdLBanco(0).Picture = LoadResPicture("Print", vbResIcon)
cmdLBanco(1).Picture = LoadResPicture("Salir", vbResIcon)
cmdLBanco(2).Picture = LoadResPicture("Libro1", vbResIcon)
'

End Sub
Private Function getIdCuenta(numCuenta As String) As Integer
Dim rst As ADODB.Recordset
Set rst = cnnLBanco.Execute("select idcuenta from cuentas where numcuenta='" & numCuenta & "'")
If Not (rst.EOF And rst.BOF) Then
    getIdCuenta = rst(0).Value
End If
rst.Close
Set rst = Nothing
End Function

Private Sub Form_Resize()
gridLBanco.Height = Me.Height - gridLBanco.Top - fraLBanco(2).Height - fraLBanco(0).Height - fraLBanco(0).Top
fraLBanco(2).Top = 150 + gridLBanco.Top + gridLBanco.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rstLBanco(0) = Nothing
Set rstLBanco(1) = Nothing
cnnLBanco.Close
Set ObjCnn = Nothing
End Sub

Private Sub lbl_Click(Index As Integer)
If Index = 1 Then
    For I = 0 To chk.UBound
        chk(I).Value = vbUnchecked
    Next
    Call muestra_libro
ElseIf Index = 4 Then
    Call muestra_libro
End If

End Sub


'-------------------------------------------------------------------------------------------------
'   Rutina: muestra_libro
'
'   distribuye en el grid todos los registros del ADODB.Recordset (rstLbanco(0))
'-------------------------------------------------------------------------------------------------
Private Sub muestra_libro()
Dim l%, SaldoA@, debe@, haber@, Filtro$, Periodo$ 'variables locales
'
MousePointer = vbHourglass

Call RtnConfigUtility(True, "Libro Banco Inmueble " & gcNomInm, "Buscando Saldo anterior...", "Espere un momento por favor.....")
DoEvents
Text1 = "0,00"
If IsDate(MskFecha(0)) Then
    SaldoA = ModGeneral.ejecutar_procedure("procLibroBancoSaldoFecha", MskFecha(0), _
    Me.DtcCuentas(0), getIdCuenta(DtcCuentas(0)), gcCodInm).Fields(0)
    Text1 = Format(SaldoA, "#,##0.00")
End If
Call RtnConfigUtility(True, "Libro Banco Inmueble " & gcNomInm, "Seleccionando registros", "Espere un momento por favor.....")
DoEvents
Set rstLBanco(0) = ModGeneral.ejecutar_procedure("procLibroBanco", Me.DtcCuentas(0), _
getIdCuenta(DtcCuentas(0)), gcCodInm)
With rstLBanco(0)
    
    
    'opciones de filtrado
    Filtro = ""
    Periodo = ""
    If IsDate(MskFecha(0)) Then
         Periodo = "(Fecha >= #" & Format(MskFecha(0), "dd/mm/yyyy") & "#)"
    End If
    If IsDate(MskFecha(1)) Then
        If IsDate(MskFecha(0)) Then
            Periodo = Periodo & " and (Fecha <=#" & Format(MskFecha(1), "dd/mm/yyyy") & "#)"
        Else
            Periodo = "(Fecha <=#" & Format(MskFecha(1), "dd/mm/yyyy") & "#)"
        End If
    Else
        Periodo = Periodo & IIf(Periodo <> "", " and (Fecha <=#" & Format(Date, "dd/mm/yyyy") & "#)", "")
    End If
    '
    For I = 0 To chk.UBound
        If chk(I) Then
            Filtro = Filtro & IIf(Filtro <> "", " or ", "") & "(IDTipoMov='" & chk(I).Tag & "'%periodo%)"
            Filtro = Replace(Filtro, "%periodo%", IIf(Periodo = "", "", " AND " & Periodo))
        End If
    Next
'
    If Filtro = "" Then
        Filtro = Periodo
    End If
    .Filter = Filtro
    .Sort = "Fecha, IDCheque"
    Me.Tag = Filtro
    If Not .EOF Or Not .BOF Then
    gridLBanco.Rows = .RecordCount + 2
    .MoveFirst
    gridLBanco.Redraw = False
    Do
        l = l + 1
        Call RtnProUtility("Obteniendo información " & CLng(.AbsolutePosition * 6015 / .RecordCount * 100 / 6015) & "%", CLng(.AbsolutePosition * 6015 / .RecordCount))
        'Debug.Print CLng(.AbsolutePosition * 6015 / .RecordCount)
        DoEvents
        SaldoA = SaldoA + !debe - !haber
        gridLBanco.RowHeight(l) = 220
        gridLBanco.TextMatrix(l, 0) = Format(!fecha, "DD/mm/yy")
        gridLBanco.TextMatrix(l, 1) = !Beneficiario
        gridLBanco.TextMatrix(l, 2) = !IDTipoMov
        gridLBanco.TextMatrix(l, 3) = Format(!IDCheque, "000000   ")
        gridLBanco.TextMatrix(l, 4) = Format(!debe, "#,##0.00 ")
        gridLBanco.TextMatrix(l, 5) = Format(!haber, "#,##0.00 ")
        gridLBanco.TextMatrix(l, 6) = Format(SaldoA, "#,##0.00 ")
        SaldoA = CCur(gridLBanco.TextMatrix(l, 6))
        debe = debe + CCur(!debe)
        haber = haber + CCur(!haber)
        .MoveNext
    Loop Until .EOF
        gridLBanco.TextMatrix(l + 1, 0) = ""
        gridLBanco.TextMatrix(l + 1, 1) = "(" & l & ") Registros."
        gridLBanco.TextMatrix(l + 1, 2) = ""
        gridLBanco.TextMatrix(l + 1, 3) = ""
        gridLBanco.TextMatrix(l + 1, 4) = Format(debe, "#,##0.00")
        gridLBanco.TextMatrix(l + 1, 5) = Format(haber, "#,##0.00")
        gridLBanco.TextMatrix(l + 1, 6) = ""
    gridLBanco.Redraw = True
    'gridLBanco.Visible = True
    Else
        gridLBanco.Rows = 2
        Call rtnLimpiar_Grid(gridLBanco)
    End If
End With
Unload FrmUtility
MousePointer = vbDefault
End Sub

'-------------------------------------------------------------------------------------------------
'   Rutina: config_grid
'
'   configura la presentación del grid, ancho de columnas, alineación
'   centra el encabezado
'-------------------------------------------------------------------------------------------------
Private Sub Config_Grid()
Dim I%
With gridLBanco
    .ColAlignment(0) = flexAlignCenterCenter
    .ColAlignment(2) = flexAlignCenterCenter
    Call centra_titulo(gridLBanco, True)
End With
End Sub


Private Sub optLBanco_Click(Index As Integer)
Call muestra_libro
End Sub

Private Sub Print_report()
''variables locales
Dim strSQL$, Filtro$
Dim rpReporte As ctlReport
'
'Call clear_Crystal(ctlReport)
Set rpReporte = New ctlReport

With rpReporte
    .Reporte = gcReport + "banco.libro.rpt"
    .OrigenDatos(0) = gcPath & "\sac.mdb"
    .Formulas(0) = "Inmueble='" & gcCodInm & "-" & gcNomInm & "'"
    .Formulas(1) = "banco='" & DtcCuentas(1) & "'"
    .Formulas(2) = "numero-cuenta='" & DtcCuentas(0) & "'"
    .Formulas(3) = "hasta='" & IIf(IsDate(MskFecha(1)), MskFecha(1), Date) & "'"
    .Formulas(4) = "saldo-inicial=" & Replace(Replace(Text1, ".", ""), ",", ".")
    If chk(0) Then Filtro = "DEP"
    If chk(1) Then Filtro = Filtro & IIf(Filtro = "", "", ", ") & "TRA"
    If chk(2) Then Filtro = Filtro & IIf(Filtro = "", "", ", ") & "EFE"
    If chk(3) Then Filtro = Filtro & IIf(Filtro = "", "", ", ") & "CHE"
    If chk(4) Then Filtro = Filtro & IIf(Filtro = "", "", ", ") & "CHQ"
    If Filtro <> "" Then
        .Formulas(5) = "filtro='" & Filtro & "'"
    End If
    .Parametros(0) = DtcCuentas(0)
    .Parametros(1) = getIdCuenta(DtcCuentas(0))
    .Parametros(2) = gcCodInm
    Filtro = Replace(Replace(Me.Tag, ")", ""), "(", "")
    'reemplazamos todos los campos posibles de filtrado
    Filtro = Replace(Filtro, "IDTipoMov", "{procLibroBanco.IDTipoMov}")
    Filtro = Replace(Filtro, "Fecha", "{procLibroBanco.Fecha}")
    .FormuladeSeleccion = Filtro
    .TituloVentana = .Formulas(0)
    .Salida = crPantalla
    .Imprimir
    Call rtnBitacora("Print Libro Banco Inm.:" & gcCodInm)
    
End With
Set rpReporte = Nothing
'
End Sub

'-------------------------------------------------------------------------------------------------
'   Rutina: rtnRefrescar
'
'
'
'-------------------------------------------------------------------------------------------------
Private Sub rtnRefrescar()
''variables locales
'Dim rstRef As New ADODB.Recordset
'Dim strSQL$, StrBanco$, Cta%, NCta$, Bco%, NBco$
'Dim I%, K%
''------------------------------------------------
'On Error GoTo salir
'cmdLBanco(2).Picture = LoadResPicture("Libro2", vbResIcon)
'cmdLBanco(2).Refresh
'MousePointer = vbHourglass
'strSQL = "SELECT Cuentas.*, Bancos.NombreBanco FROM Bancos INNER JOIN Cuentas ON Bancos.IDBanco" _
'& "= Cuentas.IDBanco;"
''Selecciona la información de las cuentas correspondientes al inmueble
''y llena el vector cuentas
''With rstRef
''    .Open strSQL, cnnLBanco, adOpenStatic, adLockReadOnly, adCmdText
''    .Filter = "NumCuenta='" & Me.dtcCuentas(0) & "'"
''    If Not .EOF Or Not .BOF Then
''        Cta = !IDCuenta
''        NCta = !numCuenta
''        Bco = !IDBanco
''        NBco = !NombreBanco
''
''    cnnLBanco.BeginTrans    'Abre una transacción
''    'elimina toda la información de la tabla Libro_Banco correspondiente a la útlima
''    'fecha de actualización
''    cnnLBanco.Execute "DELETE * FROM Libro_Banco WHERE " _
''    & "IDCuenta=" & Cta
''    '
''    'agrega todos los cheques emitidos
''    strSQL = "INSERT INTO Libro_Banco (IDcuenta,Fecha,IDTipoMov,Concepto,Numeral,Debe,Haber,Hor" _
''    & "a) SELECT " & Cta & ", Cheque.FechaCheque, 'CH', Left(ChequeDetalle.Detalle,50), Cheque.IDCheque,0," _
''    & "ChequeDetalle.Monto, Cheque.Hora FROM Cheque INNER JOIN ChequeDetalle ON Cheque.IDC" _
''    & "heque=ChequeDetalle.IDCheque IN '" & gcPath & "\sac.mdb' WHERE Cheque.Cuenta = '" & NCta _
''    & "';"
''    cnnLBanco.Execute strSQL
''
''    '
''    'Agrega los cheques anulados tambien
''    strSQL = "INSERT INTO Libro_Banco (IDcuenta,Fecha,IDTipoMov,Concepto,Numeral,Debe,Haber,Hor" _
''    & "a) SELECT " & Cta & ",FechaCheque,'CH','ANULADO',IDCheque,0,0,Hora FROM ChequeAnulado IN '" & _
''    gcPath & "\sac.mdb' WHERE Cuenta='" & NCta & "'" _
''    & ";"
''    cnnLBanco.Execute strSQL
''    '
''    'Agrega los depositos efectuados por caja
''    strSQL = "INSERT INTO Libro_Banco (IDcuenta,Fecha,IDTipoMov, Concepto, Numeral,Debe,Haber,H" _
''    & "ora) SELECT " & Cta & ",TdfDepositos.Fecha,'DP','DEPOSITO Nº',TDFDepositos.IDDe" _
''    & "posito,Sum(TDFCheques.Monto),0,TIME() FROM TDFDepositos INNER JOIN TDFCheques ON TDFDepo" _
''    & "sitos.IDDeposito = TDFCheques.IDDeposito in '" & gcPath & "\sac.mdb' WHERE tdfdEPOSITOS." _
''    & "Cuenta='" & NCta & "' GROUP" _
''    & " BY TDFDepositos.IDDeposito,TdfDepositos.Fecha;"
''    cnnLBanco.Execute strSQL
''    '
''    'agrega los depositos recibidos por caja
''    If StrBanco <> NBco Then
''        '
''        cnnLBanco.Execute "DELETE * FROM Libro_Banco WHERE Fecha IN (SELECT FechaDoc FROM TDFCh" _
''        & "eques IN '" & gcPath & "\sac.mdb' WHERE Fpago='DEPOSITO' AND CodInmueble='" & _
''        gcCodInm & "' AND Banco='" & NBco & _
''        "') AND (IDTipoMov='DP' or IDTipoMov='TR')"
''        '
''        strSQL = "INSERT INTO Libro_Banco (IDcuenta,Fecha,IDTipoMov,Concepto,Numeral,Debe,Haber" _
''        & ",Hora) SELECT " & Cta & ",TDFCheques.FechaDoc,iif(TDFCheques.Fpago='Deposito','DP','TR'), TDFCheques.FPago & ' Nº',right(TDFCheques" _
''        & ".Ndoc,10),TDFCheques.Monto,0, MovimientoCaja.Hora FROM TDFCheques INNER JOIN MovimientoC" _
''        & "aja ON TDFCheques.IDRecibo = MovimientoCaja.IDRecibo IN '" & gcPath & "\sac.mdb' WHE" _
''        & "RE (TDFCheques.Fpago = 'deposito' or TDFCheques.Fpago='Transferencia') And TDFCheques.CodInmueble = '" & gcCodInm & "' and" _
''        & " TDFCheques.Banco='" & NBco & "'" _
''        & ";"
''        StrBanco = NBco
''        cnnLBanco.Execute strSQL
''    End If
''Else
''        MsgBox "Este inmueble no tiene cuentas registradas..", vbInformation, App.ProductName
''End If
''.Close
''End With
'salir:
'If Err.Number <> 0 Then 'si ocurre algún error
'    cnnLBanco.RollbackTrans
'    MsgBox "Imposible llevar a cabo la operación solicitada. Ha ocurrido el siguiente error: " _
'    & Err.Description, vbExclamation, "Error " & Err.Number
'Else
'    cnnLBanco.CommitTrans
'    MsgBox "Información actualizada..", vbInformation, App.ProductName
'    rstLBanco(0).Requery
'    Call muestra_libro
'End If
'MousePointer = vbDefault
'cmdLBanco(2).Picture = LoadResPicture("Libro1", vbResIcon)
'
End Sub
