VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmEdoCta 
   Caption         =   "Movimiento Cod. Gasto"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   Tag             =   "1"
   WindowState     =   2  'Maximized
   Begin VB.Frame fra 
      Caption         =   "Destino:"
      Height          =   1095
      Index           =   0
      Left            =   5505
      TabIndex        =   21
      Top             =   6240
      Width           =   2475
      Begin VB.OptionButton opt 
         Caption         =   "Impresora"
         Height          =   315
         Index           =   4
         Left            =   675
         TabIndex        =   23
         Top             =   675
         Width           =   1455
      End
      Begin VB.OptionButton opt 
         Caption         =   "Ventana"
         Height          =   315
         Index           =   3
         Left            =   675
         TabIndex        =   22
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Ver"
      Height          =   1005
      Index           =   2
      Left            =   9285
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6330
      Width           =   1170
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Imprimir"
      Enabled         =   0   'False
      Height          =   1005
      Index           =   1
      Left            =   8115
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6330
      Width           =   1170
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Salir"
      Height          =   1005
      Index           =   0
      Left            =   10455
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6330
      Width           =   1170
   End
   Begin VB.Frame fra 
      Caption         =   "Opciones: "
      ClipControls    =   0   'False
      Height          =   1665
      Index           =   1
      Left            =   90
      TabIndex        =   0
      Top             =   210
      Width           =   11610
      Begin VB.OptionButton opt 
         Caption         =   "Mes:"
         Height          =   315
         Index           =   0
         Left            =   1875
         TabIndex        =   8
         Top             =   1170
         Width           =   720
      End
      Begin VB.OptionButton opt 
         Caption         =   "Entre:"
         Height          =   315
         Index           =   1
         Left            =   4560
         TabIndex        =   7
         Top             =   1170
         Width           =   870
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         DataField       =   "Saldo1"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         DataSource      =   "AdoTfondos"
         Height          =   315
         Index           =   2
         Left            =   8505
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "0,00"
         Top             =   750
         Width           =   1260
      End
      Begin VB.ComboBox cmb 
         DataField       =   "TipoMovimientoCaja"
         Height          =   315
         ItemData        =   "frmEdoCta.frx":0000
         Left            =   2760
         List            =   "frmEdoCta.frx":002B
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1140
         Width           =   1290
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         DataField       =   "Saldo1"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         DataSource      =   "AdoTfondos"
         Height          =   315
         Index           =   0
         Left            =   2775
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "0,00"
         Top             =   750
         Width           =   1260
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         DataField       =   "Saldo1"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         DataSource      =   "AdoTfondos"
         Height          =   315
         Index           =   1
         Left            =   5460
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "0,00"
         Top             =   750
         Width           =   1260
      End
      Begin VB.OptionButton opt 
         Caption         =   "Todos"
         Height          =   315
         Index           =   2
         Left            =   8910
         TabIndex        =   1
         Top             =   1140
         Value           =   -1  'True
         Width           =   870
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   315
         Index           =   0
         Left            =   5460
         TabIndex        =   6
         Top             =   1140
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         CalendarTitleBackColor=   -2147483646
         CalendarTitleForeColor=   -2147483639
         Format          =   58589185
         CurrentDate     =   37603
      End
      Begin MSDataListLib.DataCombo dtc 
         DataField       =   "CodGasto"
         Height          =   315
         Index           =   0
         Left            =   1470
         TabIndex        =   9
         Top             =   360
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "CodGasto"
         Text            =   ""
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   315
         Index           =   1
         Left            =   7110
         TabIndex        =   10
         Top             =   1140
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         CalendarTitleBackColor=   -2147483646
         CalendarTitleForeColor=   -2147483639
         CustomFormat    =   "dd/th/yy"
         Format          =   58589185
         CurrentDate     =   37603
      End
      Begin MSDataListLib.DataCombo dtc 
         DataField       =   "Titulo"
         Height          =   315
         Index           =   1
         Left            =   2760
         TabIndex        =   17
         Top             =   360
         Width           =   7020
         _ExtentX        =   12383
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Titulo"
         Text            =   ""
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta Gasto:"
         Height          =   270
         Index           =   13
         Left            =   255
         TabIndex        =   15
         Top             =   420
         Width           =   1095
      End
      Begin VB.Label label1 
         BackStyle       =   0  'Transparent
         Caption         =   "y"
         Height          =   270
         Index           =   14
         Left            =   6855
         TabIndex        =   14
         Top             =   1185
         Width           =   240
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo al:"
         Height          =   270
         Index           =   15
         Left            =   7305
         TabIndex        =   13
         Top             =   795
         Width           =   990
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cargado:"
         Height          =   270
         Index           =   16
         Left            =   1575
         TabIndex        =   12
         Top             =   795
         Width           =   990
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pagado:"
         Height          =   270
         Index           =   17
         Left            =   4260
         TabIndex        =   11
         Top             =   795
         Width           =   990
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   4065
      Left            =   120
      TabIndex        =   16
      Tag             =   "1200|4800|1000|1400|1400|1400"
      Top             =   2100
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   7170
      _Version        =   393216
      ForeColor       =   -2147483646
      Cols            =   6
      FixedCols       =   0
      BackColorFixed  =   -2147483646
      ForeColorFixed  =   -2147483643
      BackColorSel    =   65280
      ForeColorSel    =   0
      BackColorBkg    =   -2147483636
      GridColor       =   -2147483633
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      MousePointer    =   99
      FormatString    =   "^Fecha |<Concepto |^Périodo|>Cargado|>Pagado |>Saldo"
      MouseIcon       =   "frmEdoCta.frx":0094
      _NumberOfBands  =   1
      _Band(0).Cols   =   6
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
End
Attribute VB_Name = "frmEdoCta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    'módulo de consulta de cargado vs. pagado de las cuentas de gastos
    'varibles locales a nivel de módulo
    Private rstCtl As ADODB.Recordset
Attribute rstCtl.VB_VarHelpID = -1
    Private objRpt As Object
    '
    Private Sub cmd_Click(Index As Integer)
    '
    Select Case Index
        'salir
        Case 0
            Unload Me
            Set frmEdoCta = Nothing
        
        'imprimir
        Case 1: Call Printer_Report
        
        'mostrar resultados en pantalla
        Case 2: Call ver_edocta
        
    End Select
    '
    End Sub

    Private Sub dtc_Click(Index As Integer, Area As Integer)
    'variables locales
    Dim Criterio As String
    '
    If Area = 2 Then
        
        If Index = 0 Then
            Criterio = "CodGasto ='" & dtc(0).Text & "'"
        Else
            Criterio = "Titulo ='" & dtc(1).Text & "'"
        End If
        rstCtl.Find Criterio
        If rstCtl.EOF Then
            rstCtl.MoveFirst
            rstCtl.Find Criterio
        End If
        dtc(0).Text = rstCtl("CodGasto")
        dtc(1).Text = rstCtl("Titulo")
        ver_edocta
    End If
    '
    End Sub

    Private Sub dtc_KeyPress(Index As Integer, KeyAscii As Integer)
    '
    Select Case Index
        '
        Case 0  'código del gasto
            Call Validacion(KeyAscii, "01234567890")
            
            If KeyAscii = 13 Then
                rstCtl.Find "Codgasto='" & dtc(0) & "'"
                If rstCtl.EOF Then
                    rstCtl.MoveFirst
                    rstCtl.Find "CodGasto='" & dtc(0) & "'"
                    If rstCtl.EOF Then
                        MsgBox "Código de gasto no registrado", vbInformation, App.ProductName
                        Exit Sub
                    End If
                End If
                dtc(0) = rstCtl("CodGasto")
                dtc(1) = rstCtl("Titulo")
                Call ver_edocta
            End If
        '
        Case 1  'titulo del gasto
            
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
            If KeyAscii = 13 Then
                rstCtl.Find "Titulo LIKE '%" & dtc(1) & "%'"
                If rstCtl.EOF Then
                    rstCtl.MoveFirst
                    rstCtl.Find "Titulo LIKE '%" & dtc(1) & "%'"
                    If rstCtl.EOF Then
                        MsgBox "Gasto no registrado", vbInformation, App.ProductName
                        Exit Sub
                    End If
                End If
                dtc(0) = rstCtl("CodGasto")
                dtc(1) = rstCtl("Titulo")
                Call ver_edocta
            End If
            
            '
    End Select
    '
    End Sub


    Private Sub Form_Load()
    'mcDatos = gcPath + gcUbica + "Inm.mdb"
    cmb.Visible = False
    dtp(0).Visible = False
    dtp(1).Visible = False
    Label1(14).Visible = False
    
    Set rstCtl = New ADODB.Recordset
    'carga la propiedad picture del objeto cmd del archivo de recursos
    cmd(0).Picture = LoadResPicture("Salir", vbResIcon)
    cmd(1).Picture = LoadResPicture("Print", vbResIcon)
    cmd(2).Picture = LoadResPicture("Ver", vbResIcon)
    '
    'obtiene el conjunto de registros de la tabla Tgastos
    rstCtl.Open "SELECT * FROM Tgastos ORDER BY CodGasto", cnnOLEDB + mcDatos, adOpenKeyset, _
    adLockOptimistic, adCmdText
    Set dtc(0).RowSource = rstCtl
    Set dtc(1).RowSource = rstCtl
    Set Grid.FontFixed = LetraTitulo(LoadResString(527), 7.5, , True)
    Set Grid.Font = LetraTitulo(LoadResString(528), 8)
    Call centra_titulo(Grid, True)
    cmb = cmb.List(0)
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    '   Rutina:     ver_edocta
    '
    '   muestra en el grid el detalle de la cuenta seleccionada, según parámtros
    '   también seleccionados por el usuario
    '---------------------------------------------------------------------------------------------
    Private Sub ver_edocta()
    'variables locales
    Dim strSQL As String
    Dim I As Long, K As Integer
    Dim rstEDOCTA As New ADODB.Recordset
    Dim curSaldo As Currency
    Dim strFiltro As String, Fecha1 As String, Fecha2 As String
    Const m$ = "#,##0.00;(#,##0.00)"
    Dim codGasto(2) As String
    
    '
    If dtc(0) = "" Or dtc(1) = "" Then
        MsgBox "Verifique el código y/o  concepto del gasto", vbCritical, App.ProductName
        Exit Sub
    End If
    MousePointer = vbHourglass
    
    For I = 1 To 2
        codGasto(I) = Left(dtc(0), 2) & I & Right(dtc(0), 3)
    Next
    codGasto(0) = dtc(0)
    If Mid(dtc(0), 3, 1) = 0 Then

        strSQL = "SELECT Fecha, Detalle, Periodo, Sum(Monto) as Cargado, 0 AS Pagado From DetFact " _
        & "WHERE CodGasto='" & codGasto(0) & "' or CodGasto='" & codGasto(1) & "' GROUP BY DetFact." _
        & "Fecha, DetFact.Detalle, DetFact.Periodo, 0, DetFact.CodGasto UNION " _
        & "SELECT Fecha,Descripcion,Cargado,0,Abs(Monto) FROM AsignaGasto WHERE CodGasto='" & codGasto(2) _
        & "' UNION SELECT Cheque.FechaCheque,  ChequeDetalle.Detalle & ' #' & " & _
        "Cheque.IDcheque, '',0,Sum(" _
        & "ChequeDetalle.Monto) FROM Cheque INNER JOIN ChequeDetalle ON Cheq" _
        & "ue.Clave = ChequeDetalle.Clave IN '" & gcPath & "\sac.mdb' GROUP BY Cheque.FechaCheque, " _
        & "ChequeDetalle.CodGasto, ChequeDetalle.Detalle,Cheque.IDCheque,ChequeDetalle.CodInm HAVIN" _
        & "G (ChequeDetalle.CodGasto='" & codGasto(0) & "' or CodGasto='" & codGasto(1) & _
        "') AND ChequeDetalle.CodInm='" & gcCodInm & "';"
    
    Else
        
        strSQL = "SELECT Fecha, Detalle, Periodo, Sum(Monto) as Cargado, 0 AS Pagado From DetFAct " _
        & "WHERE CodGasto='" & codGasto(0) & "' or CodGasto='" & codGasto(1) & "' GROUP BY DetFact." _
        & "Fecha, DetFact.Detalle, DetFact.Periodo, 0, DetFact.CodGasto UNION SELECT Cheque.FechaCheque," _
        & "ChequeDetalle.Detalle & ' #' & " & "Cheque.IDcheque, '',0,Sum(" _
        & "ChequeDetalle.Monto) FROM Cheque INNER JOIN ChequeDetalle ON Cheq" _
        & "ue.Clave = ChequeDetalle.Clave IN '" & gcPath & "\sac.mdb' GROUP BY Cheque.FechaCheque, " _
        & "ChequeDetalle.CodGasto, ChequeDetalle.Detalle,Cheque.IDCheque,ChequeDetalle.CodInm HAVIN" _
        & "G (ChequeDetalle.CodGasto='" & codGasto(0) & "' or CodGasto='" & codGasto(1) & _
        "') AND ChequeDetalle.CodInm='" & gcCodInm & "';"
    End If
    '
    rstEDOCTA.CursorLocation = adUseClient
    rstEDOCTA.Open strSQL, cnnOLEDB + mcDatos, adOpenKeyset, adLockOptimistic, adCmdText
    '
    'aplica los filtros seleccionados por los usuarios
    Fecha1 = cnnConexion.Execute("SELECT Min(Periodo) FROM Factura IN '" & _
        mcDatos & "' WHERE Fact<>'' AND Fact<>Null")(0)
    If opt(0).Value = True Then
        If CDate("01/" & cmb & "/" & Year(Date)) < Fecha1 Then
            Fecha1 = CDate("01/" & cmb & "/" & Year(Date))
        End If
        Fecha2 = DateAdd("m", 1, Fecha1)
        Fecha2 = DateAdd("d", -1, Fecha2)
        strFiltro = "Fecha >=#" & Fecha1 & "# AND Fecha <=#" & Fecha2 & "#"
    ElseIf opt(1).Value = True Then
        Fecha1 = IIf(CDate(dtp(0)) < Fecha1, Fecha1, dtp(0))
        strFiltro = "Fecha >=#" & Fecha1 & "# AND Fecha <=#" & dtp(1) & "#"
    Else
        strFiltro = "Fecha >=#" & Fecha1 & "#"
    End If
    rstEDOCTA.Filter = strFiltro
    '
    Txt(0) = "0,00"
    Txt(1) = "0,00"
    Txt(2) = "0,00"
    'grid.Rows = 2
    Grid.Visible = False
    Call rtnLimpiar_Grid(Grid)
    With rstEDOCTA
        If Not .EOF And Not .BOF Then
            cmd(1).Enabled = True
            .MoveFirst: I = 1
            Grid.Rows = .RecordCount + 1
            
            Do
                For K = 0 To .Fields.count - 1
                    If K = 2 Then
                        Grid.TextMatrix(I, K) = Format(.Fields(K), "mm/yyyy")
                    Else
                        Grid.TextMatrix(I, K) = IIf(K > 2, Format(.Fields(K), m), .Fields(K))
                    End If
                Next
                curSaldo = curSaldo + (.Fields("Cargado") - .Fields("Pagado"))
                Grid.TextMatrix(I, K) = Format(curSaldo, m)
                Txt(0) = Format(CCur(Txt(0)) + CCur(.Fields("Cargado")), m)
                Txt(1) = Format(CCur(Txt(1)) + CCur(.Fields("Pagado")), m)
                Txt(2) = Format(curSaldo, m)
                .MoveNext: I = I + 1
            Loop Until .EOF
        Else
            Grid.Visible = True
            cmd(1).Enabled = False
            MsgBox "Sin movimientos", vbInformation, "Gasto " & dtc(0)
        End If
        .Close
    End With
    Set rstEDOCTA = Nothing
    Grid.Visible = True
    MousePointer = vbDefault
    '
    End Sub

    Private Sub Form_Unload(Cancel As Integer)
    'cerrar el formulario
    On Error Resume Next
    Set objRpt = Nothing
    rstCtl.Close
    Set rstCtl = Nothing
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    '   Rutina:     Printer_Report
    '
    '   Imprime el reporte de los datos contenidos en el grid
    '---------------------------------------------------------------------------------------------
    Private Sub Printer_Report()
    'variables locales
    'Dim mdb As DBEngine
    Dim cnnTemp As New ADODB.Connection
    Dim aCont(5) As String
    Dim rpReporte As ctlReport, errLocal As Long
    '
    If Grid.TextArray(7) = "" And Grid.TextArray(8) = "" And Grid.TextArray(8) = "" Then
        MsgBox "No hay registros en pantalla.", vbInformation, App.ProductName
        Exit Sub
    End If
    'Genera la tabla para imprimir el reporte
    If Dir(App.Path & "\Temp.mdb") = "" Then
        'si no existe la bd temporal la crea en el directoria raiz
        DBEngine.CreateDatabase App.Path & "/Temp.mdb", dbLangSpanish
        cnnTemp.Open cnnOLEDB & App.Path & "\temp.mdb"
        
    End If
    '
    'si permanece cerrrada la conexion la abre
    If cnnTemp.State = 0 Then cnnTemp.Open cnnOLEDB + App.Path & "\temp.mdb"
    Dim blnExiste As Boolean
    Dim BD As Database
    Dim TDF As TableDef
    
    Set BD = DBEngine.OpenDatabase(App.Path & "\Temp.mdb")
    For Each TDF In BD.TableDefs
        If UCase(TDF.Name) = UCase("tdfTemp") Then blnExiste = True: Exit For
    Next
    If Not blnExiste Then
        cnnTemp.Execute "CREATE TABLE tdfTEmp (Fecha DATETIME,Concepto TEXT(200), Periodo TEXT(" _
        & "10),Cargado Currency, Pagado Currency,Saldo Currency)"
    End If
    'agrega los regristros a una tabla temporal primelo lo blanquea
    cnnTemp.Execute "DELETE * FROM tdfTemp"
    With Grid
    
        For I = 1 To .Rows - 1
        
            For K = 0 To 5  'llena una matriz de contenido
                aCont(K) = .TextMatrix(I, K)
            Next
            'anexa el registro
            cnnTemp.Execute "INSERT INTO tdfTEmp (Fecha,Concepto,Periodo,Cargado,Pagado,Saldo) " _
            & "VALUES ('" & aCont(0) & "','" & aCont(1) & "','" & aCont(2) & "','" & _
            CCur(aCont(3)) & "','" & CCur(aCont(4)) & "','" & CCur(aCont(5)) & "');"
        Next
    End With
    'imprime el reporte
    'Set objRpt = CreateObject("Crystal.CrystalReport")
    For I = 0 To 20000000
    Next
    Set rpReporte = New ctlReport
    With rpReporte
            .Reporte = gcReport + "cxp_CvsP.rpt"
            .OrigenDatos(0) = App.Path & "\temp.mdb"
            .Formulas(0) = "Subtitulo='" & dtc(0) & " - " & dtc(1) & "'"
            .Formulas(1) = "Inm='" & gcCodInm & " - " & gcNomInm & "'"
            .Salida = IIf(opt(3), crPantalla, crImpresora)
            If .Salida = crPantalla Then
                .TituloVentana = "Cargado vs. Pagado"
            End If
            errLocal = .Imprimir
            Call rtnBitacora("Imprimir Cargado vs. Pagado Inm:" & gcCodInm & "(" & dtc(0) _
            & ")")
            If errLocal <> 0 Then
                MsgBox "Ocurrio un error al tratar de imprimir el reporte", _
                vbExclamation, "Error " & Err
                Call rtnBitacora("Error " & Err.Description & " al imprimir el reporte")
            End If
    End With
    'cierra los objetos
    cnnTemp.Close
    Set cnnTemp = Nothing
    '
    End Sub

 
Private Sub Opt_Click(Index As Integer)
Select Case Index

       Case 0
             cmb.Visible = True
             dtp(0).Visible = False
             dtp(1).Visible = False
             Label1(14).Visible = False
      
       Case 1
            cmb.Visible = False
            dtp(0).Visible = True
            dtp(1).Visible = True
            Label1(14).Visible = True
            
       Case 2
            cmb.Visible = False
            dtp(0).Visible = False
            dtp(1).Visible = False
            Label1(14).Visible = False
    
End Select

End Sub
