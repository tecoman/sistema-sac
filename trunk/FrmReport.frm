VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmReport 
   Caption         =   "Reporte de Impresión"
   ClientHeight    =   3630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4560
   Icon            =   "FrmReport.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3630
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdPrint 
      Caption         =   "&Imprimir"
      Height          =   765
      Left            =   915
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2835
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc AdoCajas 
      Height          =   330
      Left            =   150
      Top             =   3195
      Visible         =   0   'False
      Width           =   2070
      _ExtentX        =   3651
      _ExtentY        =   582
      ConnectMode     =   4
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   1
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "Sac"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   270
      TabIndex        =   19
      Top             =   1635
      Visible         =   0   'False
      Width           =   3945
      Begin MSMask.MaskEdBox MskDesde 
         Bindings        =   "FrmReport.frx":0442
         DataField       =   "FechaMovimientoCaja"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   3
         EndProperty
         DataSource      =   "AdoMtoCaja"
         Height          =   315
         Left            =   735
         TabIndex        =   5
         Top             =   195
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/MM/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MskHasta 
         Bindings        =   "FrmReport.frx":0464
         DataField       =   "FechaMovimientoCaja"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   3
         EndProperty
         DataSource      =   "AdoMtoCaja"
         Height          =   315
         Left            =   2715
         TabIndex        =   7
         Top             =   195
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/MM/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Hasta :"
         Height          =   195
         Index           =   1
         Left            =   2115
         TabIndex        =   6
         Top             =   255
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Desde :"
         Height          =   195
         Index           =   0
         Left            =   75
         TabIndex        =   4
         Top             =   255
         Width           =   555
      End
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   765
      Left            =   2775
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2835
      Width           =   1005
   End
   Begin VB.Frame FraCaja 
      Caption         =   " Seleccione Rango de Fechas : "
      Height          =   1305
      Left            =   165
      TabIndex        =   17
      Top             =   1440
      Visible         =   0   'False
      Width           =   4155
      Begin MSDataListLib.DataCombo DCCaja 
         Bindings        =   "FrmReport.frx":0486
         DataField       =   "DescripCaja"
         DataSource      =   "AdoCajas"
         Height          =   315
         Left            =   1455
         TabIndex        =   9
         Top             =   885
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "DescripCaja"
         BoundColumn     =   "DescripCaja"
         Text            =   ""
      End
      Begin VB.Frame Frame2 
         Height          =   615
         Left            =   105
         TabIndex        =   18
         Top             =   195
         Visible         =   0   'False
         Width           =   3945
         Begin VB.ComboBox CmbPeriodo 
            Height          =   315
            ItemData        =   "FrmReport.frx":049D
            Left            =   1545
            List            =   "FrmReport.frx":04C5
            TabIndex        =   11
            Top             =   195
            Width           =   1590
         End
         Begin VB.TextBox TxtAno 
            DataField       =   "NumDocumento"
            Height          =   285
            Left            =   3240
            MaxLength       =   5
            TabIndex        =   12
            Text            =   " "
            Top             =   210
            Width           =   600
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Período Asignado :"
            Height          =   195
            Index           =   2
            Left            =   75
            TabIndex        =   10
            Top             =   240
            Width           =   1365
         End
      End
      Begin VB.Label Label2 
         Caption         =   "S&eleccione Caja :"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   930
         Width           =   1290
      End
   End
   Begin VB.Frame FrameOrden 
      Caption         =   " Ordenado Por: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2430
      TabIndex        =   16
      Top             =   60
      Width           =   1890
      Begin VB.OptionButton OptOrden 
         Caption         =   "&Alfabético"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   3
         Top             =   885
         Width           =   1200
      End
      Begin VB.OptionButton OptOrden 
         Caption         =   "&Código"
         Height          =   255
         Index           =   0
         Left            =   375
         TabIndex        =   2
         Top             =   405
         Value           =   -1  'True
         Width           =   1200
      End
   End
   Begin VB.Frame FrameDisp 
      Caption         =   " Salida Por: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   165
      TabIndex        =   15
      Top             =   90
      Width           =   1830
      Begin VB.OptionButton OptSal 
         Caption         =   "&Pantalla"
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   0
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton OptSal 
         Caption         =   "&Impresora"
         Height          =   195
         Index           =   1
         Left            =   285
         TabIndex        =   1
         Top             =   855
         Width           =   1065
      End
   End
End
Attribute VB_Name = "FrmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public apto$, Desde$, Hasta$, Hora$

Private Sub CmbPeriodo_Change()
    '
    If KeyAscii = 13 Then 'Permite Avanzar de campo con Enter
        If CmbPeriodo.Text = "" Then
            MsgBox "Debe Ingresar es Periodo", vbInformation, App.ProductName
            CmbPeriodo.SetFocus
        Else
            TxtAno.SetFocus
        End If
    End If
    '
End Sub

Private Sub CmdCancel_Click(): Unload Me
End Sub

    Private Sub CmdPrint_Click()
    'Variables locales
    Dim Fecha1$, Fecha2$, strDatos$, strSQL1$, strSQL2$, errLocal&
    Dim blnCAso As Boolean, intX As Integer
    Dim rstReg As ADODB.Recordset
    Dim rpReporte As ctlReport
    If no_valida Then Exit Sub
    MousePointer = vbHourglass
    Call rtnBitacora("Printer " & mcTitulo)
    'valida los valores mínimos necesarios para imprimir un reporte
    
    'selecciona el reporte a imprimir
    'strDatos = mcDatos
    'FrmAdmin.rptReporte.Reset
    
    'MsgBox "Debes seleccionar un Inmueble", vbInformation_
    
    Set rpReporte = New ctlReport
    'FrmAdmin.rptReporte.ReportFileName = gcReport + mcReport
    rpReporte.Reporte = gcReport + mcReport
    Select Case UCase(mcReport)
        'resumen de caja
        'Case "CAJARESUMEN.RPT": Call RtnPrintResumen: Exit Sub: MousePointer = vbDefault
        '
        'portadas de caja
        Case "CAJAP.RPT"
            blnCAso = True
            'FrmAdmin.rptReporte.DataFiles(0) = gcPath & "\sac.mdb"
            rpReporte.OrigenDatos(0) = gcPath & "\sac.mdb"
            Fecha1 = Format(MskDesde, "MM/DD/YYYY")
            Fecha2 = Format(MskHasta, "MM/DD/YYYY")
        
            strSQL1 = "SELECT MC.IDTaquilla, MC.IDRecibo, MC.InmuebleMovimientoCaja, I.Nombre, " _
            & "I.Caja,C.DescripCaja , MC.AptoMovimientoCaja, MC.CuentaMovimientoCaja, MC.Descri" _
            & "pcionMovimientoCaja, MC.MontoMovimientoCaja, P.CodGasto, P.Descripcion, P.Period" _
            & "o, P.Monto, D.CodGasto,D.Titulo, D.Monto, MC.FechaMovimientoCaja, MC.NumDocument" _
            & "oMovimientoCaja, MC.NumDocumentoMovimientoCaja1, MC.NumDocumentoMovimientoCaja2," _
            & "MC.BancoDocumentoMovimientoCaja, MC.BancoDocumentoMovimientoCaja1, MC.BancoDocum" _
            & "entoMovimientoCaja2, MC.FechaChequeMovimientoCaja, MC.FechaChequeMovimientoCaja1" _
            & ", MC.FechaChequeMovimientoCaja2, MC.FormaPagoMovimientoCaja, MC.MontoCheque, MC." _
            & "MontoCheque1, MC.MontoCheque2, MC.EfectivoMovimientoCaja , MC.FPago, MC.FPago1, " _
            & "MC.FPago2, 0, I.FondoAct, P.Facturado, MC.CodGasto FROM (((Caja as C INNER JOIN inmue" _
            & "ble as I ON C.CodigoCaja = I.Caja) INNER JOIN MovimientoCaja as MC ON I.CodInm" _
            & "= MC.InmuebleMovimientoCaja) LEFT JOIN Periodos as P ON MC.IDRecibo = P.IDRecibo" _
            & ") LEFT JOIN Deducciones as D ON P.IDPeriodos = D.IDPeriodos WHERE " & IIf(IntTaquilla = 99, "", "MC.IdTaquilla=" _
            & IntTaquilla & " AND ") & "MC.FechaMovimientoCaja BETWEEN #" & Fecha1 & "# AND #" & _
            Fecha2 & "# AND MC.HORA>=#" & Hora & "# ORDER BY I.Caja, MC.FormaPagoMovimientoC" _
            & "aja, MC.InmuebleMovimientoCaja , MC.AptoMovimientoCaja;"
            '--------
            Set rstReg = New ADODB.Recordset
            rstReg.Open strSQL1, cnnConexion, adOpenStatic, adLockReadOnly, adCmdText
            If rstReg.EOF Or rstReg.BOF Then
                MsgBox "Caja '" & IntTaquilla & "' sin movimientos..", vbInformation, _
                App.ProductName
                rstReg.Close
                Set rstReg = Nothing
                MousePointer = vbDefault
                Exit Sub
            End If
            rstReg.Close
            Set rstReg = Nothing
            '--------
            '
            Call rtnGenerator(gcPath & "\sac.mdb", strSQL1, "QdfCaja")
            
            strSQL1 = "SELECT T.Ndoc, T.Banco, 'VAR' AS Cuenta, I.Caja, T.Monto, T.FechaMov,T.I" _
            & "DTaquilla,T.FechaMov AS Efec FROM TDFCheques as T INNER JOIN Inmueble as I ON T." _
            & "CodInmueble = I.CodInm WHERE T.Fpago not in ('efectivo','cheque') AND T.FechaMov BETWEEN #" & _
            Fecha1 & "# AND #" & Fecha2 & "#  and T.IDtAQUILLA =" & IntTaquilla & " UNION SELEC" _
            & "T D.IDDeposito, D.Banco, D.Cuenta, D.Caja, Sum(C.Monto) AS Total, C.FechaMov,C.I" _
            & "DTaquilla,D.Fecha FROM TDFDepositos as D INNER JOIN TDFCheques as C ON D.IDDepos" _
            & "ito = C.IDDeposito WHERE C.FechaMov BETWEEN #" & Fecha1 & "# AND #" & Fecha2 & _
            "# " & IIf(IntTaquilla = 99, "", "AND C.IDTAQUILLA=" & IntTaquilla) & " GROUP BY D.IDDeposito, D.Banco, D.Cuenta, D" _
            & ".Caja,C.FechaMov,C.IDTaquilla,D.Fecha ORDER BY T.FechaMov DESC;"
            '-------
            Set rstReg = New ADODB.Recordset
            
            rstReg.Open strSQL1, cnnConexion, adOpenStatic, adLockReadOnly, adCmdText
            
            If rstReg.EOF Or rstReg.BOF Then
            
                MsgBox "Caja '" & IntTaquilla & "' sin movimientos..", vbInformation, _
                App.ProductName
                rstReg.Close
                Set rstReg = Nothing
                MousePointer = vbDefault
                Exit Sub
                
            End If
            
            rstReg.Close
            Set rstReg = Nothing
            '--------
'
            '------
            strSQL2 = "SELECT Caja.CodigoCaja, Caja.DescripCaja, Sum(M.Monto) AS Total, Caja.Cu" _
            & "enta1, Caja.Cuenta2, Caja.Cuenta3, M.FechaMov, M.IDTaquilla FROM (Caja INNER" _
            & " JOIN Inmueble ON Caja.CodigoCaja = Inmueble.Caja) INNER JOIN TDFCheques as " _
            & "M ON Inmueble.CodInm = M.CodInmueble GROUP BY Caja.CodigoCaja, Caja.DescripC" _
            & "aja, Caja.Cuenta1, Caja.Cuenta2, Caja.Cuenta3, M.FechaMov, M.IDtaquilla HAVI" _
            & "NG M.FechaMov BETWEEN #" & Fecha1 & "# AND #" & Fecha2 & "# AND M.IDTaquilla" _
            & "=" & IntTaquilla
            '
            For I = 1 To 2
                Call rtnGenerator(gcPath & "\sac.mdb", IIf(I = 1, strSQL1, strSQL2), _
                IIf(I = 1, "CajaPortada", "CajaTotal"))
            Next
            'para imprimir correctamente debe estar creada la consulta CajaP
            '
            If DCCaja.Text <> "TODOS" Then
                rpReporte.FormuladeSeleccion = "{CajaP.DescripCaja}='" & DCCaja & "'"
            End If
            
            
        Case "CAJAREPORT.RPT", "CAJARESUMEN.RPT"
            'generra la consulta para el reporte
            blnCAso = True
            Fecha1 = Format(MskDesde, "mm/dd/yyyy")
            Fecha2 = Format(MskHasta, "mm/dd/yyyy")
            '
            If DCCaja.Text <> "TODOS" Then mcCrit = "{QDFcaja.DescripCaja}='" & DCCaja & "'"
            '
            strSQL1 = "SELECT MC.IDTaquilla, MC.IDRecibo, MC.InmuebleMovimientoCaja, I.Nombre, " _
            & "I.Caja,C.DescripCaja , MC.AptoMovimientoCaja, MC.CuentaMovimientoCaja, MC.Descri" _
            & "pcionMovimientoCaja, MC.MontoMovimientoCaja, P.CodGasto, P.Descripcion, P.Period" _
            & "o, P.Monto, D.CodGasto,D.Titulo, D.Monto, MC.FechaMovimientoCaja, MC.NumDocument" _
            & "oMovimientoCaja, MC.NumDocumentoMovimientoCaja1, MC.NumDocumentoMovimientoCaja2," _
            & "MC.BancoDocumentoMovimientoCaja, MC.BancoDocumentoMovimientoCaja1, MC.BancoDocum" _
            & "entoMovimientoCaja2, MC.FechaChequeMovimientoCaja, MC.FechaChequeMovimientoCaja1" _
            & ", MC.FechaChequeMovimientoCaja2, MC.FormaPagoMovimientoCaja, MC.MontoCheque, MC." _
            & "MontoCheque1, MC.MontoCheque2, MC.EfectivoMovimientoCaja , MC.FPago, MC.FPago1, " _
            & "MC.FPago2, 0, I.FondoAct, P.Facturado, MC.CodGasto FROM (((Caja as C INNER JOIN inmue" _
            & "ble as I ON C.CodigoCaja = I.Caja) INNER JOIN MovimientoCaja as MC ON I.CodInm" _
            & "= MC.InmuebleMovimientoCaja) LEFT JOIN Periodos as P ON MC.IDRecibo = P.IDRecibo" _
            & ") LEFT JOIN Deducciones as D ON P.IDPeriodos = D.IDPeriodos WHERE " & IIf(IntTaquilla = 99, "", "MC.IdTaquilla=" _
            & IntTaquilla & " AND ") & "MC.FechaMovimientoCaja BETWEEN #" & Fecha1 & "# AND #" & _
            Fecha2 & "# AND MC.HORA>=#" & Hora & "# ORDER BY I.Caja, MC.FormaPagoMovimientoC" _
            & "aja, MC.InmuebleMovimientoCaja , MC.AptoMovimientoCaja;"
            '--------
            Set rstReg = New ADODB.Recordset
            rstReg.Open strSQL1, cnnConexion, adOpenStatic, adLockReadOnly, adCmdText
            If rstReg.EOF Or rstReg.BOF Then
                MsgBox "Caja '" & IntTaquilla & "' sin movimientos..", vbInformation, _
                App.ProductName
                rstReg.Close
                Set rstReg = Nothing
                MousePointer = vbDefault
                Exit Sub
            End If
            rstReg.Close
            Set rstReg = Nothing
            '--------
            '
            Call rtnGenerator(gcPath & "\sac.mdb", strSQL1, "QdfCaja")
            '
            If Trim(DCCaja.Text) = "TODOS" Then
                Set rpReporte = Nothing
                Call RtnPrintResumen
            Else
                If UCase(mcReport) = "CAJAREPORT.RPT" Then
                    If Respuesta("¿Desea Imprimir El Resumen de Caja?") Then Call RtnPrintResumen
                End If
            End If
            If UCase(mcReport) = "CAJARESUMEN.RPT" Then
                MousePointer = vbDefault
                If gcNivel < nuSUPERVISOR Then IntTaquilla = IntTemp
                Unload Me
                Set FrmReport = Nothing
                Exit Sub
            End If
            Set rpReporte = New ctlReport
            rpReporte.Reporte = gcReport + mcReport
            rpReporte.OrigenDatos(0) = gcPath & "\sac.mdb"
            rpReporte.Formulas(0) = "desde = date(" & Format(MskDesde, "yyyy,mm,dd") & ")"
            rpReporte.Formulas(1) = "hasta = date(" & Format(MskHasta, "yyyy,mm,dd") & ")"
            
            '
        Case "LISTAPRO.RPT", "FICHAPRO.RPT", "LISTAGAS.RPT", "AGENDAPROP.RPT", _
            "AGENDAJUNTACON.RPT", "FACT_LGNC.RPT", "CXC_ANAVEN.RPT"
                blnCAso = True
                'FrmAdmin.rptReporte.Formulas(0) = "Condominio = '" & gcNomInm & "'"
                rpReporte.Formulas(0) = "Condominio = '" & gcNomInm & "'"
                If UCase(mcReport) = "FICHAPRO.RPT" Then
                    'FrmAdmin.rptReporte.Formulas(1) = "Legal" & FrmAdmin.objRst.Fields("MesesMora")
                    rpReporte.Formulas(1) = "Legal=" & FrmAdmin.objRst.Fields("MesesMora")
                End If
                strDatos = mcDatos
            '
        Case "GESTION.RPT"
            blnCAso = True
            strDatos = mcDatos
            rpReporte.Formulas(0) = "Apartamento='" & apto & "'"
            rpReporte.Formulas(1) = "Condominio = '" & gcNomInm & "'"
            rpReporte.Formulas(2) = "Desde= Date(" & Desde & ")"
            rpReporte.Formulas(3) = "Hasta= Date(" & Hasta & ")"
             '
        Case "CXC_COBRADOR.RPT"
            'FrmAdmin.rptReporte.Formulas(0) = "Titulo='" & apto & "'"
            rpReporte.Formulas(0) = "Titulo='" & apto & "'"
            '
        Case "RECEPFC.RPT"
            blnCAso = True
            rpReporte.OrigenDatos(0) = gcPath & "\SAC.MDB"
            rpReporte.OrigenDatos(1) = gcPath & "\SAC.MDB"
            rpReporte.OrigenDatos(2) = gcPath & "\SAC.MDB"
            rpReporte.Formulas(0) = "desde = date(" & Format(MskDesde, "yyyy,mm,dd") & ")"
            rpReporte.Formulas(1) = "hasta = date(" & Format(MskHasta, "yyyy,mm,dd") & ")"
            '
        Case "LISCHEQDEV.RPT"
            rpReporte.OrigenDatos(0) = gcPath & "\sac.mdb"
            rpReporte.OrigenDatos(1) = gcPath & "\sac.mdb"
            rpReporte.Formulas(0) = "desde = date(" & Format(MskDesde, "yyyy,mm,dd") & ")"
            rpReporte.Formulas(1) = "hasta = date(" & Format(MskHasta, "yyyy,mm,dd") & ")"
            
            blnCAso = True
            '
        Case UCase("cfd_deuda.Rpt")
            mcOrdCod = ""
            mcOrdAlfa = ""
        
        Case "NOM_REPORT.RPT"
            rpReporte.Formulas(0) = "Titulo='" & mcTitulo & "'"
            rpReporte.Formulas(1) = "SubTitulo='" & apto & "'"
            mcOrdCod = "+{qdfNomina.CodEmp}"
            mcOrdAlfa = "+{qdfNomina.Apellidos}"
                    
        Case "NOM_NOV.RPT"
            Call Printer_Nom_Nov(IIf(OptSal(1), crImpresora, crPantalla), apto)
            MousePointer = vbDefault
            Exit Sub
            
        Case "EMP_FICHA.RPT"
            For intX = 0 To 5: rpReporte.OrigenDatos(intX) = gcPath + "/sac.mdb"
            Next
             
        Case "NOM_AGUI.RPT"
            rpReporte.ArbolGrupo = True
            
    End Select
    
    If Not blnCAso Then
        rpReporte.OrigenDatos(0) = gcPath & "\sac.mdb"
    End If
            
    '    If Frame2.Visible = True Then
    '        If CmbPeriodo.Text = "" Or TxtAno.Text = "" Then
    '            MsgBox ("Debe Ingresar datos del Periodo...")
    '            Exit Sub
    '        End If
    '
    '        Select Case CmbPeriodo.Text
    '            Case Is = "Enero"
    '                  mcPeriodo = "01" + Mid(Trim(TxtAno), 3, 2)
    '            Case Is = "Febrero"
    '                  mcPeriodo = "02" + Mid(Trim(TxtAno), 3, 2)
    '            Case Is = "Marzo"
    '                  mcPeriodo = "03" + Mid(Trim(TxtAno), 3, 2)
    '            Case Is = "Abril"
    '                  mcPeriodo = "04" + Mid(Trim(TxtAno), 3, 2)
    '            Case Is = "Mayo"
    '                  mcPeriodo = "05" + Mid(Trim(TxtAno), 3, 2)
    '            Case Is = "Junio"
    '                  mcPeriodo = "06" + Mid(Trim(TxtAno), 3, 2)
    '            Case Is = "Julio"
    '                  mcPeriodo = "07" + Mid(Trim(TxtAno), 3, 2)
    '            Case Is = "Agosto"
    '                  mcPeriodo = "08" + Mid(Trim(TxtAno), 3, 2)
    '            Case Is = "Septiembre"
    '                  mcPeriodo = "09" + Mid(Trim(TxtAno), 3, 2)
    '            Case Is = "Octubre"
    '                  mcPeriodo = "10" + Mid(Trim(TxtAno), 3, 2)
    '            Case Is = "Noviembre"
    '                  mcPeriodo = "11" + Mid(Trim(TxtAno), 3, 2)
    '            Case Is = "Diciembre"
    '                  mcPeriodo = "12" + Mid(Trim(TxtAno), 3, 2)
    '            End Select
    '    End If
    
'    If OptOrden(0).Value = True Then FrmAdmin.rptReporte.SortFields(0) = mcOrdCod
'    If OptOrden(1).Value = True Then FrmAdmin.rptReporte.SortFields(0) = mcOrdAlfa
        
    If mcCrit <> "" Then
        'FrmAdmin.rptReporte.SelectionFormula = mcCrit
        rpReporte.FormuladeSeleccion = mcCrit
    End If
    If OptSal(1).Value = True Then 'Activa Salida
        'FrmAdmin.rptReporte.Destination = crptToPrinter
        rpReporte.Salida = crImpresora
    Else
'        FrmAdmin.rptReporte.Destination = crptToWindow
'        FrmAdmin.rptReporte.WindowParentHandle = FrmAdmin.hWnd
'        FrmAdmin.rptReporte.WindowState = crptMaximized
'        FrmAdmin.rptReporte.WindowShowCloseBtn = True
        rpReporte.Salida = crPantalla

    End If
    'FrmReport.Caption = mcTitulo
    rpReporte.TituloVentana = mcTitulo
    '    If Frame1.Visible Then
    '            frmadmin.rptreporte.Formulas(1) = "desde = date(" & Format(MskDesde, "yyyy,mm,dd") & ")"
    '            frmadmin.rptreporte.Formulas(2) = "hasta = date(" & Format(MskHasta, "yyyy,mm,dd") & ")"
    '    End If
    
    '    If Frame2.Visible Then
    '
    '        frmadmin.rptreporte.Formulas(1) = "Inmueble = '" & gcNomInm & "'"
    '        frmadmin.rptreporte.Formulas(2) = "FondoReserva = " & Str(mnFondoAct)
    '        frmadmin.rptreporte.Formulas(3) = "Inm = '" & gcCodInm & "'"
    '        frmadmin.rptreporte.Formulas(4) = "Periodo = '" & mcPeriodo & "'"
    '        frmadmin.rptreporte.SelectionFormula = "{Gastos.Periodo}={@Periodo} and {Gastos.CodInm}={@Inm}"
    '
    '    End If
    
        
    '        If  Then
    '            Value = True
    '            'frmadmin.rptreporte.SelectionFormula = "{QDFcaja.IDTaquilla}=" & IntTaquilla & " And {QDFcaja.FechaMovimientoCaja}>={@Desde}AND{QDFcaja.FechaMovimientoCaja}<={@Hasta}"
    '        Else
    '            Value = False
    '            'frmadmin.rptreporte.SelectionFormula = "{QDFcaja.IDTaquilla}=" & IntTaquilla & " AND {QDFcaja.DescripCaja} = '" & DCCaja & "' and {QDFcaja.FechaMovimientoCaja}>={@desde} and {QDFcaja.FechaMovimientoCaja}<={@hasta}"
    '        End If
    If strDatos <> "" Then 'FrmAdmin.rptReporte.DataFiles(0) = strDatos
        rpReporte.OrigenDatos(0) = strDatos
    End If
    'FrmAdmin.rptReporte.WindowTitle = FrmReport.Caption

    'errLocal = FrmAdmin.rptReporte.PrintReport
    rpReporte.Imprimir
    
'    If errLocal <> 0 Then MsgBox FrmAdmin.rptReporte.LastErrorString, vbCritical, "Error " & _
'    FrmAdmin.rptReporte.LastErrorNumber
    '
    MousePointer = vbDefault
    If gcNivel < nuSUPERVISOR Then IntTaquilla = IntTemp
    Unload Me
    Set FrmReport = Nothing
    '
    End Sub

    Private Sub DCCaja_Click(Area As Integer): If Area = 2 Then CmdPrint.SetFocus
    End Sub
    
    Private Sub DCCaja_KeyPress(KeyAscii As Integer): If KeyAscii = 13 Then CmdPrint.SetFocus
    End Sub

    Private Sub Form_Load()
    AdoCajas.ConnectionString = cnnOLEDB + gcPath + "\sac.mdb"
    AdoCajas.RecordSource = "SELECT * FROM Caja ORDER BY DescripCaja"
    AdoCajas.Refresh
    MskDesde = Date
    MskHasta = Date
    If gcNivel < nuSUPERVISOR Then Frame1.Enabled = True
    DCCaja.Text = "TODOS"
    CmdCancel.Picture = LoadResPicture("Salir", vbResIcon)
    CmdPrint.Picture = LoadResPicture("Print", vbResIcon)
    CenterForm Me
    End Sub

    Private Sub MskDesde_KeyPress(KeyAscii As Integer)
    'Permite Avanzar de campo con Enter
    If KeyAscii = 13 Then
        If MskHasta.Enabled = True Then
            With MskHasta
                .SetFocus
                .SelStart = 0
                .SelLength = Len(MskHasta)
            End With
        Else
            DCCaja.SetFocus
        End If
    End If
    End Sub
    
    Private Sub MskHasta_KeyPress(KeyAscii As Integer)
    'Permite Avanzar de campo con Enter
    If KeyAscii = 13 Then
        If FraCaja.Visible = True Then
            DCCaja.SetFocus
        Else
            CmdPrint.SetFocus
        End If
    End If
    End Sub

    Private Sub optOrden_Click(Index As Integer)
    If Frame1.Visible = True Then
        RtnEnter
    Else
        If CmbPeriodo.Visible Then CmbPeriodo.SetFocus
    End If
    End Sub

    Private Sub OptSal_Click(Index As Integer)
    On Error Resume Next
    If Frame1.Visible = True Then
        RtnEnter
    Else
        CmbPeriodo.SetFocus
    End If
    End Sub

    Private Sub OptSal_KeyPress(Index As Integer, KeyAscii As Integer)
    If FraCaja.Visible Then RtnEnter
    End Sub
    Sub RtnEnter()
        
    If Frame1.Enabled And MskDesde.Enabled Then
        With MskDesde
            .SetFocus
            .SelStart = 0
            .SelLength = Len(MskDesde)
        End With
    Else
        DCCaja.SetFocus
    End If
        
    End Sub

    '---------------------------------------------------------------------------------------------
    '   Rutina: rtnPrintResumes
    '
    '   Imprime los resumenes del movimiento de caja
    '---------------------------------------------------------------------------------------------
    Private Sub RtnPrintResumen()
    'variables locales
    Dim strReporte$
    Dim rpReporte As ctlReport
    
    If FraCaja.Visible Then
        
        If DCCaja.Text = "TODOS" Then
            
            For I = 0 To 1
            
                strReporte = IIf(I = 0, "CajaResumen.rpt", "cajaresumen1.rpt")
                Set rpReporte = New ctlReport
                rpReporte.Reporte = gcReport + strReporte
                rpReporte.OrigenDatos(0) = gcPath + "\sac.mdb"
                rpReporte.Formulas(0) = "desde = date(" & Format(MskDesde, "yyyy,mm,dd") & ")"
                rpReporte.Formulas(1) = "hasta = date(" & Format(MskHasta, "yyyy,mm,dd") & ")"
                rpReporte.TituloVentana = IIf(I = 0, "Resumen de Pagos", " Resumen de Caja")
                If IntTaquilla = 99 Then
                    rpReporte.FormuladeSeleccion = IIf(I = 0, "{TDFCheques.FechaMov}>={@Desde} AND {TDFCheques.FechaMov}<={@Hasta}", "{QDFcaja.FechaMovimientoCaja}>={@Desde} AND {QDFcaja.FechaMovimientoCaja}<={@Hasta}")
                Else
                    rpReporte.FormuladeSeleccion = IIf(I = 0, "{TDFCheques.IDTaquilla}=" & IntTaquilla & " AND " _
                    & "{TDFCheques.FechaMov}>={@Desde} AND {TDFCheques.FechaMov}<={@Hasta}", "{QDFCa" _
                    & "ja.IDTaquilla}=" & IntTaquilla & "AND {QDFcaja.FechaMovimientoCaja}>={@Desde" _
                    & "} AND {QDFcaja.FechaMovimientoCaja}<={@Hasta}")
                End If
                
                rpReporte.Salida = IIf(OptSal(0), crPantalla, crImpresora)
                rpReporte.Imprimir
                
                Set rpReporte = Nothing
            Next
            
        Else
            
            For I = 0 To 1
                Set rpReporte = New ctlReport
                strReporte = IIf(I = 0, "CajaResumen.rpt", "cajaresumen1.rpt")
                rpReporte.Reporte = gcReport + strReporte
                rpReporte.Formulas(1) = "desde = date(" & Format(MskDesde, "yyyy,mm,dd") & ")"
                rpReporte.Formulas(2) = "hasta = date(" & Format(MskHasta, "yyyy,mm,dd") & ")"
                rpReporte.FormuladeSeleccion = IIf(I = 0, "{TDFCheques.FechaMov}>={@Desde} AND{TDFCheques.FechaMov}<={@Hasta} AND {QdfInmCaja.DescripCaja}='" & DCCaja & "'", "{QDFcaja.FechaMovimientoCaja}>={@Desde} AND {QDFcaja.FechaMovimientoCaja}<={@Hasta} AND {QdfCaja.DescripCaja}='" & DCCaja & "'")
                rpReporte.TituloVentana = IIf(I = 0, "Resumen de Pagos", " Resumen de Caja")
                rpReporte.Imprimir
                Set rpReporte = Nothing
            Next
            
        End If
        '
        '
        Set rpReporte = New ctlReport
        With rpReporte
        
            .Reporte = gcReport + "caja_rfd.rpt"
            .OrigenDatos(0) = gcPath & "\sac.mdb"
            .OrigenDatos(1) = gcPath & "\sac.mdb"
            .Formulas(0) = "desde = date(" & Format(MskDesde, "yyyy,mm,dd") & ")"
            .Formulas(1) = "hasta = date(" & Format(MskHasta, "yyyy,mm,dd") & ")"
            .Formulas(2) = "CodSys='" & sysCodInm & "'"
            .TituloVentana = "Caja Relación Fondo - Deuda " & MskDesde & " al " & MskHasta
            .Salida = IIf(OptSal(0), crPantalla, crImpresora)
            .Imprimir
        End With
        Set rpReporte = Nothing
    End If
    '
    End Sub


    Function no_valida() As Boolean
    '
    If Frame1.Visible = True Then
        If Not IsDate(MskDesde) Or Not IsDate(MskHasta) Then
            no_valida = MsgBox("Debe Ingresar el rango de fecha del reporte...", vbExclamation, _
            App.ProductName)
        End If
    End If
    '
    If FraCaja.Visible And DCCaja = "" Then
        no_valida = MsgBox("Falta seleccionar la caja...", vbExclamation, App.ProductName)
    End If
    '
    End Function

