VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmPagoWeb 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pagos en línea www.administradorasac.com"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14610
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   14610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmd 
      Caption         =   "&Imprimir"
      Height          =   495
      Index           =   2
      Left            =   11745
      TabIndex        =   4
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Salir"
      Height          =   495
      Index           =   1
      Left            =   13080
      TabIndex        =   2
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Procesar"
      Height          =   495
      Index           =   0
      Left            =   10425
      TabIndex        =   1
      Top             =   5895
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   4815
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   8493
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
      MergeCells      =   3
      AllowUserResizing=   1
      BorderStyle     =   0
      FormatString    =   "|^ID |^Fecha Transacción |^Período|>Monto |>Saldo"
      MouseIcon       =   "frmPagoWeb.frx":0000
      _NumberOfBands  =   1
      _Band(0).Cols   =   6
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   2
   End
   Begin ComctlLib.ProgressBar pBar 
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   6000
      Visible         =   0   'False
      Width           =   9570
      _ExtentX        =   16880
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   0
      Min             =   1e-4
   End
   Begin VB.Image img 
      Height          =   240
      Index           =   3
      Left            =   0
      Picture         =   "frmPagoWeb.frx":0162
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image img 
      Height          =   240
      Index           =   2
      Left            =   480
      Picture         =   "frmPagoWeb.frx":033A
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image img 
      Height          =   240
      Index           =   1
      Left            =   6240
      Picture         =   "frmPagoWeb.frx":051F
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image img 
      Height          =   240
      Index           =   0
      Left            =   5760
      Picture         =   "frmPagoWeb.frx":0704
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmPagoWeb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const scUserAgent = "test-SAC"
Const INTERNET_OPEN_TYPE_DIRECT = 1
Const INTERNET_OPEN_TYPE_PROXY = 3
Const INTERNET_FLAG_RELOAD = &H80000000
Const sURL = "http://www.administradorasac.com/pago.consultar.asp"
'Const sURL = "http://server-pronet21/administradorasac/pago.consultar.asp"
Const sServerFTP = "administradorasac.com"
Const sUser = "admras"
Const sPass = "dmn+str"
Const sDir = "httpdocs/cancelacion.gastos"

Private Declare Function InternetOpen Lib "wininet" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetCloseHandle Lib "wininet" (ByRef hInet As Long) As Long
Private Declare Function InternetReadFile Lib "wininet" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Private Declare Function InternetOpenUrl Lib "wininet" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal lpszUrl As String, ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
        
Private sCodPagoCondominio As String, sCodAbonoCuenta As String, sCodAbonoFuturo As String
Private sNombreInmueble As String, sCajaInmueble As String
Private aFacturasC() As String
Private cErrores As New Collection
Private sProceso As String

Private Sub cmd_Click(Index As Integer)
Select Case Index
    Case 1
        Unload Me
        Set frmPagoWeb = Nothing
    
    Case 0
        procesar_pago
        
    Case 2
        imprimir_reporte
End Select
End Sub

Private Sub imprimir_reporte()
Dim rpReporte As ctlReport
    
    '
    Set rpReporte = New ctlReport
    With rpReporte
    '
        .Reporte = gcReport + "pagosweb.rpt"
        .TituloVentana = "strTitulo"
        .Salida = crPantalla
        .Imprimir
        If Err <> 0 Then MsgBox Err.Description, vbCritical, Err
    '
    End With

End Sub


Private Function BoundedText(ByVal ptr As Object, ByVal Txt _
    As String, ByVal max_wid As Single) As String
    Do While ptr.TextWidth(Txt) > max_wid
        Txt = Left$(Txt, Len(Txt) - 1)
    Loop
    BoundedText = Txt
End Function
Private Sub Form_Load()
    'KPD-Team 1999
    'URL: http://www.allapi.net/
    'E-Mail: KPDTeam@Allapi.net
    MousePointer = vbHourglass
    Dim hOpen As Long, hFile As Long, sBuffer As String, Ret As Long
    Dim Pago, Detalle, hConnect As Long
    
    'Create a buffer for the file we're going to download
    'Create an internet connection
    hOpen = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
    'Open the url
    hFile = InternetOpenUrl(hOpen, sURL, vbNullString, ByVal 0&, INTERNET_FLAG_RELOAD, ByVal 0&)
    
    'InternetReadFile hFile, sBuffer, 1000, 0
    Dim bDoLoop             As Boolean
    Dim sReadBuffer         As String * 2048
    Dim lNumberOfBytesRead  As Long
    bDoLoop = True
    While bDoLoop
     sReadBuffer = vbNullString
     bDoLoop = InternetReadFile(hFile, _
        sReadBuffer, Len(sReadBuffer), lNumberOfBytesRead)
     sBuffer = sBuffer & _
          Left(sReadBuffer, lNumberOfBytesRead)
     If Not CBool(lNumberOfBytesRead) Then bDoLoop = False
    Wend

    'clean up
    InternetCloseHandle hFile
    InternetCloseHandle hOpen
    'sBuffer = "|9|3/8/2010 9:16:00 PM|D|0456789900|3/8/2010|249,27||banesco|" & _
    "01340013987000912344|P|edgar.messia@cantv.net|True|(0426)9554668|2551|" & _
    "PB03|249,27|0210551019|<br>|10|3/8/2010 9:18:00 PM|T|000989|3/7/2010|42,18|" & _
    "banesco|venezuela|01252344008990929091|P|info@pronet21.com|True|(0212)" & _
    "7519992|2511|001A|42,18|0703511001|<br>|11|3/8/2010 9:30:00 PM|D|0008990|" & _
    "3/8/2010|1,445,78||banesco|01340000123498000001|P|ynfantes@gmail.com|" & _
    "True|(0212)4841057|2532|0051|610,87|0110532021|2532|0051|367,57|" & _
    "0210532021|2532|0051|467,34|0809532021|<br>|12|3/8/2010 9:46:00 PM|T|" & _
    "0809|3/8/2010|371,19|venezolano de credito|banesco|01550003987898983048" & _
    "|P|ynfantes@gmail.com|True|(0212)5751892|2524|001A|186,06|0305524001|" & _
    "2524|001A|185,13|0405524001|"

    img(2).Picture = LoadResPicture("UnChecked", vbResBitmap)
    img(3).Picture = LoadResPicture("Checked", vbResBitmap)
    grid.FormatString = "Sel|ID|^Fecha|^FP|Documento|^Fecha Doc.|Monto|Banco Origen|Banco Destino|^Cuenta|^Est.|email|Env|Teléfono|Eli"
    grid.Tag = "320|0|1000|350|959,8111|1124,787 | 1000| 1440 | 1440 | 2129,953 | 404,7874 | 1680,284 | 300| 1200|300"
    centra_titulo grid, True
    If Left(sBuffer, 1) = "|" Then
        Listar (sBuffer)
    Else
        MsgBox "Problemas al descargar la información. " & _
        "Verifique su conexión a Internet e inténtelo nuevamente." & _
        vbCrLf & "Si el problema persiste, póngase en contacto con el administrador del sistema", _
        vbCritical, App.ProductName
        
    End If
    If IntTaquilla = 0 Then
        Dim cnn As ADODB.Connection
        Set cnn = New ADODB.Connection
        cnn.Open cnnOLEDB & gcPath & "\tablas.mdb"
        basSeguridad.rtnCajero cnn
        cnn.Close
        Set cnn = Nothing
    End If
    MousePointer = vbDefault
End Sub


Private Sub Listar(contenido As String)
Dim Reg() As String, maestro() As String, Filas As Integer
Dim Fila As Integer, valor As String

rtnLimpiar_Grid grid
Reg = Split(contenido, "<br>")
Filas = UBound(Reg)
grid.Rows = Filas + 1
For i = 0 To Filas - 1
    Fila = Fila + 1
    grid.Col = 0
    grid.Row = Fila
    Set grid.CellPicture = img(2)
    grid.CellPictureAlignment = flexAlignCenterCenter
    maestro = Split(Reg(i), "|")
    For j = 0 To 13
        valor = maestro(j)
        If IsDate(valor) And Not IsNumeric(valor) Then
            valor = Format(valor, "dd/mm/yyyy")
        ElseIf valor = "True" Or valor = "False" Then
            grid.Col = 12
            grid.Row = Fila
            grid.CellPictureAlignment = flexAlignCenterCenter
            Set grid.CellPicture = IIf(valor = "True", img(0), img(1))
            valor = ""
            grid.Col = 0
        ElseIf j = 9 Then
           valor = Format(valor, "0###-####-####-####-####")
        ElseIf j = 7 Or j = 8 Then
            valor = Trim(Replace(UCase(valor), "BANCO", ""))
        End If
        grid.TextMatrix(Fila, j) = valor
        
    Next
    
    grid.Col = 0
    grid.ColSel = grid.Cols - 1
    grid.FillStyle = flexFillRepeat
    grid.CellBackColor = RGB(240, 240, 240)
    grid.FillStyle = flexFillSingle
    'agregamos el detalle por factura
    For Z = 14 To UBound(maestro) - 1
        Fila = Fila + 1
        grid.AddItem "", Fila
        For X = 0 To 3
            grid.TextMatrix(Fila, X + 4) = maestro(Z + X)
        Next
        Z = Z + 3
    Next
Next
'
End Sub

Private Sub grid_Click()
Dim Fila As Integer
If grid.ColSel = 0 Then
    '   marcarmos|demarcamos pagos revisados en banco
    grid.Row = grid.RowSel
    If grid.CellPicture = img(2) Or grid.CellPicture = img(3) Then
        Set grid.CellPicture = IIf(grid.CellPicture = img(2), img(3), img(2))
        Fila = grid.RowSel
        If grid.CellPicture = img(3) And _
            grid.TextMatrix(Fila, 3) = "T" And _
            (grid.TextMatrix(Fila, 7) <> grid.TextMatrix(Fila, 8)) Then
                grid.TextMatrix(Fila + 1, 2) = _
                InputBox("Ingrese en número de operación registrada en banco." & vbCrLf & "SI EL NUMERO DE OPERACION EN BANCO COINCIDE CON LA REGISTRADA POR EL CLIENTE DEJE ESTE VALOR EN BLANCO", _
                "Número de Operación")
                If Trim(grid.TextMatrix(Fila + 1, 2)) = Trim(grid.TextMatrix(Fila, 4)) Then grid.TextMatrix(Fila + 1, 2) = ""
                
        End If
    End If
ElseIf grid.ColSel = 14 Then
    If grid.TextMatrix(grid.RowSel, 1) <> "" Then
        grid.Row = grid.RowSel
        grid.Col = 0
        Set grid.CellPicture = img(2)
        If grid.TextMatrix(grid.RowSel, 14) = "" Then
            grid.TextMatrix(grid.RowSel, 14) = "Sí"
        Else
            grid.TextMatrix(grid.RowSel, 14) = ""
        End If
    End If
End If
'
End Sub

Private Sub procesar_pago()
Dim i As Integer, ID As Integer, Fila As Integer
Dim Pago As Currency, mfactura As Currency, sql As String, dAbono As Double, mAbono As Double
Dim sINM As String, sFact As String, sFP As String, sRecibo As String
Dim cFactura As Double, sEmail As String, sApto As String, sNdoc As String, sFdoc As String
Dim sBanco As String, sNdoc2 As String
Dim bTrans As Boolean, n As Integer, sDescrip As String, u As Integer
Dim rstlocal As ADODB.Recordset


cmd(0).Enabled = False
cmd(1).Enabled = False
pBar.Visible = True
pBar.Max = grid.Rows - 1
Call rtnBitacora("Inicio procesar pagos web")


For i = 1 To grid.Rows - 1
    grid.Col = 0
    grid.Row = i
    pBar.Value = i
    
    'si el registro esta marcado para procesar, entramos en esta rutina
    If grid.CellPicture = img(3) And grid.TextMatrix(i, 10) = "P" Then
        
        'abrimos una transaccion para guardas este pago
        cnnConexion.BeginTrans
        bTrans = True
        On Error GoTo ReversarPago:
        Call rtnBitacora("-- Inicar transacción...")
        
        Fila = grid.RowSel
        grid.Row = i + 1
        
        ' seteamos las variables generales del pago
        Pago = grid.TextMatrix(i, 6)
        ID = grid.TextMatrix(i, 1)      'monto total del pago
        sFP = IIf(grid.TextMatrix(i, 3) = "D", "DEPOSITO", "TRANSFERENCIA")
        sEmail = grid.TextMatrix(i, 11)
        sNdoc = grid.TextMatrix(i, 4)
        sFdoc = grid.TextMatrix(i, 5)
        sNdoc2 = grid.TextMatrix(i + 1, 2)
        sBanco = grid.TextMatrix(i, 8)

        
        If formaPagoYaRegistrada(sNdoc, sBanco, sFP) Then
            sProceso = "Agregar Pago"
            Err.Raise -2147467259 + ID, "Registrar " & sFP, sFP & " " & sBanco & " " & sNdoc & " ya registrado(a)"
        End If
            
        mfactura = 0
        n = 0
        
        Do While (grid.CellPicture = 0)
            
            sDescrip = ""
            sINM = grid.TextMatrix(grid.RowSel, 4)
            sApto = grid.TextMatrix(grid.RowSel, 5)
            mAbono = CDbl(grid.TextMatrix(grid.RowSel, 6))
            sFact = grid.TextMatrix(grid.RowSel, 7)
            
            ' si la factura tiene saldo pendiente
            cFactura = monto_factura(sINM, sFact)
            If cFactura > 0 Then
            
                sDescrip = sDescrip & Left(sFact, 2) & "-" & Mid(sFact, 3, 2)
                
                '   efectuamos el registro en la tabla MovimientoCaja
                sRecibo = guardar_movimiento_caja(sINM, sApto, sFP, sNdoc, _
                        sFdoc, mAbono, sBanco, sDescrip, sCodPagoCondominio, _
                        "PAGO CONDOMINIO (via web)", sNdoc2)
                        
                'cFactura = monto_factura(sINM, sFact)
                mfactura = mfactura + cFactura
                
                ' actualizamos la tabla factura, propietarios, periodos, Inmueble
                actualizar_factura sINM, sFact, sRecibo, cFactura, sCodPagoCondominio, _
                    "PAGO CONDOMINIO (via web)"
            
            End If
            
            pBar.Value = i + 1
            DoEvents
            
            
            'Guardar_NumFact sFact, cFactura
            ReDim Preserve aFacturasC(n)
            
            aFacturasC(n) = sFact & "|" & cFactura
            
            If grid.Row + 1 < grid.Rows Then
                grid.Row = grid.Row + 1
            Else
                Exit Do
            End If
            n = n + 1
        
        Loop
        'sDescrip = Left(sDescrip, Len(sDescrip) - 1)
        
        
        'sql = "UPDATE MovimientoCaja SET MontoMovimientoCaja='" & Pago & _
              "', MontoCheque='" & Pago & "', DescripcionMovimientoCaja='" & sDescrip & _
              "' WHERE IDRecibo='" & sRecibo & "'"
        
'        cnnConexion.Execute sql, u
'        Call rtnBitacora("---- Actualizar (" & u & ") el movimiento de caja #" & sRecibo)
        '
        If mfactura = 0 Then
            
            grid.TextMatrix(Fila, 10) = "R"
            cnnConexion.RollbackTrans
            Err.Clear
        
        Else
            
            If Pago > mfactura Then
            
                dAbono = Pago - mfactura
                '   si existe alguna diferencia efectuamos el abono en la cuenta del cliente
                '   validamos si existe factura pendiente para abonarle a esa factura
                sql = "SELECT * FROM factura IN '" & gcPath & "\" & sINM & "\inm.mdb' " & _
                      "WHERE codprop='" & sApto & "' AND saldo > 0 ORDER BY periodo ASC"
                Set rstlocal = cnnConexion.Execute(sql)
                
                ' si hay facturas pendientes abonamos a esa factura
                If Not (rstlocal.EOF And rstlocal.BOF) Then
                    
                    sDescrip = "Abono a Cuenta " & Format(rstlocal("periodo"), "mm-yy")
                    
                    '   efectuamos el registro en la tabla MovimientoCaja
                    sRecibo = guardar_movimiento_caja(sINM, sApto, sFP, sNdoc, _
                        sFdoc, dAbono, sBanco, sDescrip, sCodAbonoCuenta, _
                        "ABONO A CUENTA (via web)", sNdoc2)
                        
                    ' actualizamos la tabla factura, propietarios, periodos, Inmueble
                    actualizar_factura sINM, sFact, sRecibo, dAbono, sCodAbonoCuenta, _
                    "ABONO A CUENTA (via web)"
                    
'                    sql = "UPDATE Factura IN '" & gcPath & "\" & sINM & "\inm.mdb' set Saldo = Saldo - '" & dAbono & "', usuario='" & _
'                    gcUsuario & "', fecha='" & Format(Time(), "hh:mm:ss") & "', pagado = pagado + '" & dAbono & "' WHERE FACT='" & rstlocal("FACT") & "'"
'                    cnnConexion.Execute sql, u
'                    Call rtnBitacora("---- Aplicar (" & u & ") abono factura " & rstlocal("FACT") & ".")
'                    '
'                    'registramos el movimiento en la tabla periodo
'                    sql = "INSERT INTO Periodos(IDRecibo,IDPeriodos,Periodo,CodGasto,Descripcion,Monto,Facturado) " & _
'                    "VALUES ('" & sRecibo & "','" & sRecibo & Format(rstlocal("periodo"), "mmyy") & "','" & _
'                    Format(rstlocal("periodo"), "mm-yy") & "','" & sCodAbonoCuenta & "','ABONO A CUENTA','" & dAbono & _
'                    "','" & rstlocal("facturado") & "')"
'                    cnnConexion.Execute sql, u
'
'                    sql = "UPDATE MovimientoCaja SET DescripcionMovimientoCaja = " & _
'                    "DescripcionMovimientoCaja & ' / Abono a Cuenta " & Format(rstlocal("periodo"), "mm-yy") & _
'                    "' WHERE IDRecibo='" & sRecibo & "'"
'
'                    cnnConexion.Execute sql, u
                     
                Else
                    Dim Periodo As String
                    
                    
                    sDescrip = "ABONO A PROX. FACTURACION (via web) " & Periodo
                    
                    '   efectuamos el registro en la tabla MovimientoCaja
                    sRecibo = guardar_movimiento_caja(sINM, sApto, sFP, sNdoc, _
                        sFdoc, dAbono, sBanco, sDescrip, sCodAbonoCuenta, _
                        "ABONO A FUTURO (via web)", sNdoc2)
                        
                    ' actualizamos la tabla factura, propietarios, periodos, Inmueble
                    actualizar_factura sINM, sFact, sRecibo, dAbono, sCodAbonoCuenta, _
                    "ABONO A CUENTA (via web)"
                    
'                    '   hacemos un abono a futuro, no hay facturas pendientes
'                    '   registramos el movimiento en la tabla periodo
'                    sql = "INSERT INTO Periodos(IDRecibo,IDPeriodos,Periodo,CodGasto,Descripcion,Monto,Facturado) " & _
'                    "VALUES('" & sRecibo & "','" & sRecibo & sCodAbonoFuturo & "','" & _
'                    Format(DateAdd("m", 1, CDate("01-" & Left(sFact, 2) & _
'                    "-" & Mid(sFact, 3, 2))), "MM-YY") & "'," & _
'                    "'" & sCodAbonoFuturo & "','Abono Próx. Facturación','" & dAbono & _
'                    "',0)"
'                    cnnConexion.Execute sql, u
'
'                    ' actualizamos la información del movimiento de caja
'                    sql = "UPDATE MovimientoCaja SET DescripcionMovimientoCaja = " & _
'                    "DescripcionMovimientoCaja & ' / Abono a Prox.Facturación' WHERE IDRecibo='" & _
'                    sRecibo & "'"
'                    cnnConexion.Execute sql, u
                                        
                End If
                '   insertamos la informacion del documento
            'Dim sCaja As String
                
                ' actualizamos la información del propietario
'                sql = "UPDATE propietarios IN '" & gcPath & "\" & sINM & "\inm.mdb'  SET Deuda = Deuda - '" & dAbono & "', " & _
'                "Ultpago = Ultpago + '" & dAbono & "' WHERE Codigo='" & sApto & "'"
'                cnnConexion.Execute sql, u

                

            End If
            '   ---------------------- INGRESAMOS EL DOCUMENTO A LA TABLA CHEQUES ---------------------
            ModGeneral.insertar_registro "procPagoAdd", sRecibo, CStr(IntTaquilla), sINM, Format(Date, "dd/mm/yy"), _
            sFP, sNdoc, sBanco, Format(CDate(sFdoc), "dd/mm/yy"), Pago, ""
            '   ------ SI ES TRANSFERENCIA REGISTRAMOS OTRO DOCUMENTO POR EL REGISTRADO EN BANCO ------
            If sNdoc2 <> "" And sFP = "TRANSFERENCIA" Then
                If formaPagoYaRegistrada(sNdoc2, sBanco, sFP) Then
                    sProceso = "Agregar Pago"
                    Err.Raise -2147467259 + ID, "Registrar " & sFP, sFP & " " & sBanco & " " & sNdoc2 & " ya registrado(a)"
                End If
                ModGeneral.insertar_registro "procPagoAdd", sRecibo, CStr(IntTaquilla), sINM, _
                Format(Date, "dd/mm/yy"), sFP, sNdoc2, sBanco, Format(CDate(sFdoc), "dd/mm/yy"), 0, ""
                
            End If
            '
ReversarPago:
            If Err <> 0 Then
                cnnConexion.RollbackTrans
                grid.Col = 10
                grid.Row = Fila
                grid.CellForeColor = vbRed
                grid.TextMatrix(Fila, 10) = "E"
                Err.Description = "Línea: " & Fila & ": " & Err.Description
                cErrores.Add Err.Description
                Call rtnBitacora("-- Error en operacion. ID: " & ID & ". " & Err.Description)
                Err.Clear
            Else
                cnnConexion.CommitTrans
                '   guardamos el movimiento de caja
                grid.TextMatrix(Fila, 10) = "A"
                Call rtnBitacora("-- Pago OK!!!. Cerrar Transacción.")
                ' enviamos la cancelación de gastos vía email
                enviar_recibos sINM, sRecibo, sEmail
                Call rtnBitacora("-- Enviar email de confirmación.")
                '
                grid.TextMatrix(Fila + 1, 8) = actualizarFTP(ID, "A")
            End If
            bTrans = False
            i = grid.RowSel - 1
        
        End If
        
    ElseIf grid.TextMatrix(i, 14) = "Sí" Then ' eliminar pago
        ID = grid.TextMatrix(i, 1)
        Call rtnBitacora("Pago ID: " & ID & " rechazado por el usuario")
        grid.TextMatrix(i + 1, 8) = actualizarFTP(ID, "R")
    End If
    
Next
If cErrores.Count > 0 Then
    If bTrans Then cnnConexion.RollbackTrans
    sDescrip = "El proceso se completo, pero ocurrieron los siguientes errores: " & vbCrLf
    'For Each error In cErrores
    For i = 1 To cErrores.Count
        sDescrip = sDescrip & "- " & cErrores.Item(i) & vbCrLf
    Next
    sDescrip = sDescrip & "Verifique las transacciones con estatus (E) en color rojo."
    MsgBox sDescrip
Else
    Call rtnBitacora("Fin proceso pago web.")
    MsgBox "Proceso finalizado con éxito.", vbInformation, App.ProductName
End If
pBar.Visible = False
cmd(0).Enabled = True
cmd(1).Enabled = True
End Sub

Private Function actualizarFTP(ID As Integer, Estatus As String) As String
Dim hOpen As Long, hFile As Long, sBuffer As String, Ret As Long
Dim Detalle, hConnect As Long, URL As String
URL = "http://www.administradorasac.com/pago.confirmar.asp?id=" & ID & "&estatus=" & Estatus
sBuffer = ""
' Create a buffer for the file we're going to download
' Create an internet connection
hOpen = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
' Open the url
hFile = InternetOpenUrl(hOpen, URL, vbNullString, ByVal 0&, INTERNET_FLAG_RELOAD, ByVal 0&)
Dim bDoLoop             As Boolean
Dim sReadBuffer         As String * 2048
Dim lNumberOfBytesRead  As Long
bDoLoop = True
While bDoLoop
 sReadBuffer = vbNullString
 bDoLoop = InternetReadFile(hFile, _
    sReadBuffer, Len(sReadBuffer), lNumberOfBytesRead)
 sBuffer = sBuffer & _
      Left(sReadBuffer, lNumberOfBytesRead)
 If Not CBool(lNumberOfBytesRead) Then bDoLoop = False
Wend
actualizarFTP = sBuffer
'clean up
InternetCloseHandle hFile
InternetCloseHandle hOpen

End Function

Private Function monto_factura(Inmueble As String, Factura As String) As Double
Dim sql As String, cnn As ADODB.Connection, rst As ADODB.Recordset

sProceso = "monto_factura"
Set cnn = New ADODB.Connection

cnn.Open cnnOLEDB$ & gcPath & "\" & Inmueble & "\inm.mdb"
sql = "select saldo from factura where fact='" & Factura & "'"
Set rst = cnn.Execute(sql)
If Not (rst.EOF And rst.BOF) Then monto_factura = rst("saldo")

rst.Close
Set rst = Nothing

End Function


Private Sub actualizar_factura(Inmueble As String, Factura As String, _
ReciboCaja As String, dPagado As Double, Codigo As String, Detalle As String)

Dim StrRutaInmueble As String, sql As String, u As Integer
sProceso = "actualizar_factura"
StrRutaInmueble = gcPath & "\" & Inmueble & "\inm.mdb"
'
'   actualizamos la deuda del propietario
'
sql = "UPDATE Propietarios INNER JOIN Factura ON Propietarios.Codigo = Factura.codprop " & _
      "IN '" & StrRutaInmueble & "' SET Propietarios.UltPago = '" & dPagado & "',Propietarios.Recibos = Recibos - 1, " & _
      "Propietarios.FecUltPag=date(), Propietarios.FecReg=date(), Propietarios.Usuario='" & _
      gcUsuario & "', Propietarios.Deuda=Propietarios.Deuda - Factura.Saldo WHERE Factura.FACT='" & Factura & "';"
cnnConexion.Execute sql, u
Call rtnBitacora("---- Actualizar (" & u & ") Deuda Propietario.")
'
'   agregamos los registros a la tabla Periodos
'
sql = "INSERT INTO Periodos(IDRecibo,IDPeriodos,Periodo,CodGasto, Descripcion, Monto, Facturado ) " & _
      "SELECT '" & ReciboCaja & "','" & ReciboCaja & "' & format(periodo, 'mmyy')," & _
      "format(periodo,'mm-yy'),'" & Codigo & "','" & Detalle & "'," & _
      "Saldo, Facturado FROM Factura IN '" & _
      StrRutaInmueble & "' WHERE Fact='" & Factura & "'"
cnnConexion.Execute sql, u

Call rtnBitacora("---- Registrar (" & u & ") Peridodo Cancelado.")
'
'   actualizamos la tabla factura
'
sql = "UPDATE Factura IN '" & StrRutaInmueble & "' SET " & "Pagado = Pagado + Saldo ," & _
      "Saldo = 0, freg=Date(), usuario='" & gcUsuario & _
      "', fecha=Format(Time(),'hh:mm:ss') WHERE Fact= '" & Factura & "'"

cnnConexion.Execute sql, u
Call rtnBitacora("---- Actualizar Saldo (" & u & ") factura Nº " & Factura)
'
'   actualizamos la deuda general del inmueble
sql = "UPDATE Inmueble INNER JOIN MovimientoCaja ON Inmueble.CodInm = MovimientoCaja." & _
    "InmuebleMovimientoCaja SET Inmueble.Deuda = Inmueble.Deuda-MovimientoCaja." & _
    "MontoMovimientoCaja WHERE (((MovimientoCaja.IDRecibo)='" & ReciboCaja & "'));"
cnnConexion.Execute sql, u
            
Call rtnBitacora("---- Actualizar (" & u & ") deuda inmueble.")
'
End Sub

Private Function guardar_movimiento_caja(Inmueble As String, _
    Apartamento As String, FormaPago As String, _
    NumeroDocumento As String, FechaDocumento As String, _
    Monto As Double, Banco As String, Descripcion As String, _
    Codigo As String, CuentaMovimiento As String, _
    Optional NumeroDocumento2 As String) As String

Dim strRecibo, nTransaccion As Integer
Dim rst As ADODB.Recordset, sql As String
Dim r As Integer, sCaja As Integer

sProceso = "CajaInmueble"

sCaja = CajaInmueble(Inmueble)
sCajaInmueble = sCaja

sProceso = "procNumTransaccion"
Set rst = ejecutar_procedure("procNumTransaccion", Date, sCaja)

sProceso = "guardar_movimiento_caja"

nTransaccion = rst("maximo") + 1

strRecibo = Right(Inmueble, 2) & Apartamento & Format(Date, "ddmmyy") & Format(nTransaccion, "00")

sql = "INSERT INTO MovimientoCaja(IDTaquilla, IDRecibo, NumeroMovimientoCaja, FechaMovimientoCaja, " & _
    "TipoMovimientoCaja, MontoMovimientoCaja, CodGasto, CuentaMovimientoCaja, InmuebleMovimientoCaja," & _
    "AptoMovimientoCaja, FormaPagoMovimientoCaja, BancoDocumentoMovimientoCaja, " & _
    "EfectivoMovimientoCaja, FPago, NumDocumentoMovimientoCaja, FechaChequeMovimientoCaja, " & _
    "MontoCheque, Usuario, Freg, Hora,DescripcionMovimientoCaja) VALUES(" & IntTaquilla & ",'" & strRecibo & "'," & nTransaccion & _
    ",Date(),'INGRESO','" & Monto & "','" & Codigo & "','" & CuentaMovimiento & "','" & Inmueble & "','" & _
    Apartamento & "','" & FormaPago & "','" & Banco & "',0,'" & FormaPago & "','" & NumeroDocumento & "','" & FechaDocumento & "','" & Monto & _
    "','" & gcUsuario & "',Date(),Time(),'" & Descripcion & "')"

cnnConexion.Execute sql, r

Call rtnBitacora("---- Guardar (" & r & ") Movimiento de Caja #" & strRecibo)

If r > 0 Then guardar_movimiento_caja = strRecibo

End Function

Private Function CajaInmueble(Inmueble As String) As Integer
Dim rst As ADODB.Recordset

Set rst = ejecutar_procedure("procBuscaCaja", Inmueble)
If Not (rst.EOF And rst.BOF) Then
    CajaInmueble = rst("Caja")
    sCodPagoCondominio = rst("CodPagoCondominio")
    sCodAbonoCuenta = rst("CodAbonoCta")
    sCodAbonoFuturo = rst("CodAbonoFut")
    sNombreInmueble = rst("Nombre")
End If
rst.Close
Set rst = Nothing
End Function

'---------------------------------------------------------------------------------------------
'   Rutina: Guardar_NumFact
'
'   Guarda el número del recibo cancelado y el monto real pagado
'---------------------------------------------------------------------------------------------
Private Sub Guardar_NumFact(strNumFact As String, Cancelado As Double)
Dim numFichero%   'variables locales
Dim strArchivo$
'
numFichero = FreeFile
strArchivo = App.Path & Archivo_Temp
On Error GoTo CerrarArchivo
10 Open strArchivo For Append As numFichero
Write #numFichero, strNumFact, Cancelado
Close numFichero
Exit Sub
CerrarArchivo:
Close numFichero
GoSub 10
'
End Sub

'---------------------------------------------------------------------------------------------
'   Rutina: Imprimir_Recibos
'
'   Recorre el archivo Archivo_Temp y genera una copia en formato pdf de la canelacion
'   y se la envia por email al cliente
'---------------------------------------------------------------------------------------------
Private Sub enviar_recibos(Inmueble As String, Recibo As String, email As String)
'Variables locales
Dim strArchivo$, Factura$
Dim Pago@, numFichero%
Dim Carpeta$, sDirectorioLocal$, sArchivo$
'Dim pEmail As New clsSendMail
Dim mFTP As New cFtp, Facturas() As String
'
'
'
'With pEmail
'    .SMTPHostValidation = VALIDATE_HOST_DNS
'    .EmailAddressValidation = VALIDATE_SYNTAX
'    .Delimiter = ";"
'    .SMTPHost = "mail.cantv.net"
'    .FromDisplayName = "Pago Condominio vía web"
'    .from = "pagoscondominio@administradorasac.com"
'    .Message = strBody()
'    .Recipient = email
'    .RecipientDisplayName = "Administrador"
'    .Subject = strSubject
'    ' adjuntamos los avisos de cobro
    sDirectorioLocal = Environ("temp") & "\"
    mFTP.SetModeActive
    mFTP.SetTransferBinary
    
    
    If mFTP.OpenConnection(sServerFTP, sUser, sPass) Then
        mFTP.SetFTPDirectory sDir
        'numFichero = FreeFile
        'strArchivo = App.Path & Archivo_Temp
        'Open strArchivo For Input As numFichero 'abre el archivo de recibos cancelados
        For i = 0 To UBound(aFacturasC)
            Facturas = Split(aFacturasC(i), "|")
            Carpeta = "\" & Inmueble & "\"
            sArchivo = Facturas(0) & ".pdf"
            Call Printer_Pago(Facturas(0), CCur(Facturas(1)), Carpeta, _
            Inmueble, sNombreInmueble, Recibo, True, 0, crArchivoDisco)
            If Err <> 0 Then
                Call rtnBitacora(Err.Description)
            Else
                '.Attachment = sDirectorio & sArchivo
                If Not (mFTP.SimpleFTPPutFile(sDirectorioLocal & sArchivo, sArchivo)) Then
                    MsgBox mFTP.GetLastErrorMessage
                End If
            End If
                '
        'Loop Until EOF(numFichero)
        Next
        Erase aFacturasC
        'Close numFichero
        'If Dir(strArchivo, vbArchive) <> "" Then Kill strArchivo
        mFTP.CloseConnection
        ' confirmamos en el pago al cliente, hacemos un llamado a la pagina web
        ' para que actualice el estatus del pago y envie una conirmación al cliente
    Else
        MsgBox "No se establecion conexion."
    End If
'    .Send
'End With
'
End Sub

Private Function formaPagoYaRegistrada(Numero_Documento$, Nombre_Banco$, Forma_Pago$) As Boolean
Dim rst As ADODB.Recordset

Set rst = ejecutar_procedure("procPagoExiste", Numero_Documento, Nombre_Banco, Forma_Pago)
formaPagoYaRegistrada = Not (rst.EOF And rst.BOF)
rst.Close
Set rst = Nothing
End Function
