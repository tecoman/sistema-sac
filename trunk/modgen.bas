Attribute VB_Name = "ModGeneral"
    '------------------------------------------------------------------------------------------
    'esta primera lines corresponde a las variables globales de la empresa en el sistema
    Public sysEmpresa As String, sysCodInm As String, sysCodPro As String, sysCodCaja As String
    Public mcTitulo As String                'Titulo del Formulario de Tablas Múltiples
    Public mcDatos As String               'Ruta y Nombre de Base de Datos en Reporte
    Public mcTablas As String               'Tabla Seleccionada para el
    Public gcUbica As String                'Ubicación de los datos del Sistema
    Public gcUbiGraf As String              'Ubicación de los Iconos del Sistema
    Public gcReport As String               'Ubicación de los Reportes del Sistema
    Public mcReport As String               'Nombre del Archivo Report
    Public mcOrdCod As String            'Orden del Reporte por Código
    Public mcOrdAlfa As String              'Orden del Reporte por Alfabeto
    Public gcCodInm As String               'Variable Global que guarda Código de Inmueble Activo.
    Public gcNomInm As String           'Variable Global que guarda Nombre de Inmueble Activo.
    Public gcUsuario As String          'Variable Global que guarda lOGIN de Usuario Activo.
    Public gcNombreCompleto As String ' Variable global nombre completo del usuario.
    Public gcCodFondo As String         'Código Fondo Reserva.
    Public gnPorcFondo As Long          'Pública Porcentaje de fondo.
    Public gnPorIntMora As Double    '.................
    Public gnMesesMora As Double    '.................
    Public gnCta As Long                'Variable Global Cuenta Inm/Cuenta Pote
    Public cnnConexion As ADODB.Connection    'Variable Publica de Conexion a SAC.mdb
    Public mcCrit As String               '
    Public gnIva As Double              'Alícuota Impuesto Valor Agregado
    Public gnIDB As Double              'Alícuota Impuesto al Débito Bancario
    Public gcPath As String               'Ubicación principal.
    Public ObjCnn As New ADODB.Connection   'Objeto Publico Conexión
    Public ObjCmd As New ADODB.Command      'Objeto Publico Comando
    Public Const CUENTA_INMUEBLE% = 0
    Public Const CUENTA_POTE% = 1
    '--------------------------------------------
    '   CONSTANTES SERVIDOR SMTP
    '--------------------------------------------
    Private Const SMTP_SERVER$ = "smtp.gmail.com"
    Private Const SMTP_SERVER_PORT& = 465
    Private Const SERVER_AUTH  As Boolean = True
    Private Const USER_NAME$ = "webmail.pronet21@gmail.com"
    Private Const Password$ = "admras5231/-"
    Private Const SSL As Boolean = True

    'Public Const INVERSIONES$ = "9999"  'Codigo de Inmueble Inversiones
    '
    Public Const cnnOLEDB$ = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
    Public strLlamada As String * 1             '
    Public Const Archivo_Temp = "\RECANSAC.log" '
    Public Matriz_A()               '
    Public DataServer As String     'nombre del servidor de datos
    Public IDMoneda As String       '
    Public nEnviados As Long
    '
    Public Enum NivelUsuario
        nuADSYS = 0
        nuAdministrador
        nuSUPERVISOR
        nuUSUARIO
        nuINACTIVO
    End Enum
    '
    Public Enum ESTATUS_CONVENIO
        '
        CONVENIO_ACTIVO = 1
        CONVENIO_INACTIVO
        CONVENIO_INCUMPLIDO
        CONVENIO_CUMPLIDO
        '
    End Enum
    '
    Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long
    Public Const NV_CLOSEMSGBOX As Long = &H5000&
    Public Declare Function SetTimer& Lib "user32" (ByVal hWnd&, ByVal nIDEvent&, _
    ByVal uElapse&, ByVal lpTimerFunc&)
    Public Declare Function FindWindow& Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName$, ByVal lpWindowName$)
    Public Declare Function LockWindowUpdate& Lib "user32" (ByVal hwndLock&)
    Public Declare Function SetForegroundWindow& Lib "user32" (ByVal hWnd&)
    Public Declare Function MessageBox& Lib "user32" Alias "MessageBoxA" _
    (ByVal hWnd&, ByVal lpText$, ByVal lpCaption$, ByVal wType&)
    Public Declare Function KillTimer& Lib "user32" (ByVal hWnd&, ByVal nIDEvent&)
    Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" _
    (ByVal nBufferLength As Long, ByVal lpbuffer As String) As Long
    Public Const API_FALSE As Long = 0&
    Type rSSO
        CodInm As String * 6
        Nfact As String * 8
        Monto As String * 11
    End Type
    Type Detalle_Cheque
        ID As String * 6
        Monto As String * 11
        Clave As String * 30
    End Type
    Public Enum crSalida
        crPantalla = 0
        crImpresora = 1
        crArchivoDisco = 2
        crEmail = 3
    End Enum
    Dim frmTT As frmToolTip
    Dim frmLocal As frmView
    
    Global Const SYNCHRONIZE = &H100000
    Global Const INFINITE = &HFFFFFFFF
    
    Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, _
    ByVal dwProcessId As Long) As Long
    Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
    Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
    
    Rem-------------------------------------------------------------------------------------------
    
    Public Sub TimerProc(ByVal hWnd&, ByVal uMsg&, ByVal idEvent&, ByVal dwTime&, ipCaption$)
    KillTimer hWnd, idEvent
    '
    Dim hMessageBox&
    '
    hMessageBox = FindWindow("#32770", App.ProductName)
    If hMessageBox Then
    Call SetForegroundWindow(hMessageBox)
    SendKeys "{enter}"
    End If
    Call LockWindowUpdate(API_FALSE)
    '
    End Sub

        Sub WaitForTerm(ByVal PID As Long)
        On Error GoTo Gestion_Error
    
        'Variables locales
        Dim phnd As Long
    
        phnd = OpenProcess(SYNCHRONIZE, 0, PID)
        If phnd <> 0 Then
            Call WaitForSingleObject(phnd, INFINITE)
            Call CloseHandle(phnd)
        End If
    Exit Sub
Gestion_Error:
        Call MsgBox(Err.Number & ": " & Err.Description)
    End Sub

    Public Sub Validacion(KeyAscii%, StrValido$) '
    'si no es codigo de control
    If KeyAscii > 26 Then If InStr(StrValido, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End Sub

    '---------------------------------------------------------------------------------------------
    '   Rutina:     BuscaInmueble
    '
    '   Busca datos inmuebles
    '---------------------------------------------------------------------------------------------
    Public Sub BuscaInmueble(StrConsulta$, data As DataCombo) '-
    'variables locales
    '
    With ObjCmd
    '
        .ActiveConnection = cnnConexion
        .CommandType = adCmdText
        .CommandText = "SELECT * FROM Monedas INNER JOIN (Caja INNER JOIN inmueble ON Caja.Codi" _
        & "goCaja = inmueble.Caja) ON Monedas.IdMoneda=Caja.MonedaCaja WHERE (((" & StrConsulta _
        & ") LIKE '%" & data.Text & "%'));"
    '
    End With
    '
    End Sub
    
    

    'Rev.05/09/2002-------------------------------------------------------------------------------
    Public Function BuscaProp(StrCampo$, StrRecord$, ado As Adodc) As Boolean
    '---------------------------------------------------------------------------------------------
    'BUSCA EL CODIGO Y NOMBRE DE UN PROPIETARIO SEGUN PARAMETROS ENVIADOS POR EL USUARIO
    '
    With ado.Recordset
    '
        If .EOF Or .BOF Then Exit Function
        .MoveFirst
        .Find StrCampo & " LIKE '*" & StrRecord & "*'"
        If Not .EOF Then BuscaProp = True
        
    '
    End With
    '
    End Function

    'Rev.-22/08/2002---Presentación de la barra de herramientas según el boton pulsado------------
    Public Sub RtnEstado(boton%, BarraH As ToolBar, Optional Rstsr As Boolean)
    '
    With BarraH
    '
        
        Select Case boton
    '
            Case 5, 10 'AGREGAR/ELIMINAR REGISTRO
    '       -------------------------------------------------------
                For I = 1 To 12
                    If I = 6 Or I = 8 Then
                        .Buttons(I).Enabled = True
                    Else
                        .Buttons(I).Enabled = False
                    End If
                Next
                If IntButton = 5 Then .Buttons(12).Enabled = False
    '
            Case 6, 8, 9 'GUARDAR/CANCELAR/ELIMINAR REGISTRO
    '       -------------------------------------------------------
                For I = 1 To 12
                    If I = 6 Or I = 8 Then
                        .Buttons(I).Enabled = False
                    Else
                        .Buttons(I).Enabled = True
                    End If
                Next
    '
        End Select
        'efectua una verificación de registros en el ADODB.Recordset
        'para activar los botones de la barra herramienta
        If Rstsr Then
            For I = 1 To 4: .Buttons(I).Enabled = False
            Next I
        End If
    '
    End With
    '
    End Sub

'12/08/2002---------------------------------------------Elimina un registro del movimiento de caja
Sub RtnEliminarMovCaja(strID$, strCod$)
'variables locales
Dim objRst As ADODB.Recordset   'Información Movimiento de Caja
Dim StrApto$, strCarpeta$, StrInmueble$ 'vairables cadena
Dim VecInforma(0 To 1)  As Currency     'Contine inf. total Bs, N° Recibos
Dim rstPago As ADODB.Recordset      '
Dim Reg As Long, N As Long                                '
'--------------------------------------------------------
Set objRst = New ADODB.Recordset       'Crea una instancia del ADODB.Recordset
'
On Error GoTo rtnCerrar 'Captura algún error producido durante el proceso
cnnConexion.BeginTrans  'Comienza la Transaccion
Call RtnProUtility("Procesando por lotes..", 1015)
'---------------------------------------------------------------------------------------------
'Selecciona La información necesaria del movimiento de caja
objRst.Open "SELECT Mc.IDTaquilla, Mc.IDRecibo, Mc.InmuebleMovimientoCaja, I.Nombre, I.Caja, " _
& "Caja.DescripCaja, Mc.AptoMovimientoCaja As Apto,  Mc.CuentaMovimientoCaja, Mc.Descripcio" _
& "nMovimientoCaja, Mc.MontoMovimientoCaja, Periodos.CodGasto, Periodos.Descripcion, Period" _
& "os.Periodo, Periodos.Monto, Deducciones.CodGasto, Deducciones.Titulo, Deducciones.Monto," _
& " Mc.FechaMovimientoCaja, Mc.NumDocumentoMovimientoCaja, Mc.NumDocumentoMovimientoCaja1, " _
& "Mc.NumDocumentoMovimientoCaja2, Mc.BancoDocumentoMovimientoCaja, Mc.BancoDocumentoMovimi" _
& "entoCaja1, Mc.BancoDocumentoMovimientoCaja2, Mc.FechaChequeMovimientoCaja, Mc.FechaChequ" _
& "eMovimientoCaja1, Mc.FechaChequeMovimientoCaja2, Mc.FormaPagoMovimientoCaja, Mc.MontoChe" _
& "que, Mc.MontoCheque1, Mc.MontoCheque2, Mc.EfectivoMovimientoCaja, Mc.FPago, Mc.FPago1, " _
& "Mc.FPago2, MC.CodGasto as ING, I.CodIngresosVarios, I.CodCCheq, Periodos.Facturado, " _
& "I.CodAbonoCta, I.CodRcheq,I.CodAbonoFut, I.CodPagoCondominio FROM (((Caja INNER JOIN inm" _
& "ueble  as I ON Caja.CodigoCaja = I.Caja) INNER JOIN MovimientoCaja AS Mc ON I.CodInm = M" _
& "c.InmuebleMovimientoCaja) LEFT JOIN Periodos ON Mc.IDRecibo = Periodos.IDRecibo) LEFT JO" _
& "IN Deducciones ON Periodos.IDPeriodos = Deducciones.IDPeriodos WHERE (((Mc.IDRecibo)='" _
& strID & "'))", cnnConexion, adOpenKeyset, adLockOptimistic
    'Llena el ADODB.Recordset
'
With objRst
'   Inicializa las variables locales--------------------------------------------------------------
    .MoveFirst
    strCodPC = !CodPagoCondominio
    strCodAbonoCta = !CodAbonoCta
    strCodAbonoFut = !CodAbonoFut
    strCodIV = !CodIngresosVarios
    strCodCCHeq = !CodCCheq
    strCodRCheq = !CodRcheq
    StrApto = objRst!apto
    StrInmueble = Trim(!InmuebleMovimientoCaja)
    strCarpeta = "\" + StrInmueble + "\"
    VecInforma(0) = 0
    VecInforma(1) = 0
    '-------------------------------------------------------------------------------------------------
    'si el pago afecta una cuenta de fondo hace el ajuste
    If booFondo(.Fields("Ing"), gcPath & strCarpeta & "inm.mdb") Then
    
        cnnConexion.Execute "UPDATE MovFondo IN '" & gcPath & strCarpeta & "inm.mdb' SET De" _
        & "l=True WHERE Fecha=#" & Format(.Fields("FechaMovimientoCaja"), "mm/dd/yy") & "# " _
        & "AND Concepto='" & .Fields("DescripcionMovimientoCAja") & "' AND Tipo='NC' AND Co" _
        & "dGasto='" & .Fields("ING") & "' AND Haber=" & CCur(.Fields("MontoMovimientoCaja"))
        Call rtnBitacora("Emilinando Movimiento del Fondo '" & .Fields("ING") & "'")

    End If
    '
    'Elimina la información de la tabla TDFcheques
    Call RtnProUtility("Eliminando Detalle del pago...", 3000)
    
    Set rstPago = New ADODB.Recordset
    
    If .Fields("FormaPagoMovimientoCaja") = "EFECTIVO" Then
    
        cnnConexion.Execute "DELETE * FROM tdfCheques WHERE IdRecibo='" & strID & "'", N
        If N > 0 Then Call rtnBitacora("Pago en efectivo...Eliminado")
        
    Else
'
        'cnnConexion.Execute "DELETE * FROM tdfCheques WHERE IdRecibo='" & strID & "'"
        For I = 18 To 20    'elimina c/documento relacionado con el pago
            If .Fields(I) <> "" Then
                rstPago.Open "SELECT * FROM TDFCheques WHERE Ndoc='" & .Fields(I) & "' AND Banc" _
                & "o='" & .Fields(3 + I) & "' AND FechaDoc=#" & Format(.Fields(6 + I), "mm/dd/yy") _
                & "#;", cnnConexion, adOpenKeyset, adLockOptimistic
                If Not rstPago.EOF And Not rstPago.BOF Then
                    If rstPago!Monto = .Fields(I + 10) Then
                        cnnConexion.Execute "DELETE * FROM tdfCheques WHERE Ndoc='" & .Fields(I) & "'", N
                        If N > 0 Then Call rtnBitacora(rstPago!FPago & " " & .Fields(I) & " Eliminado..")
                    Else
                        cnnConexion.Execute "UPDATE TdfCheques SET Monto = Monto - '" & CCur(.Fields(I + 10)) _
                            & "' WHERE Ndoc='" & .Fields(I) & "' AND Banco='" & .Fields(3 + I) & "' AND FechaDoc=#" & Format(.Fields(6 + I), "mm/dd/yy") & "#", N
                        If N > 0 Then Call rtnBitacora(N & " " & rstPago!FPago & " " & .Fields(I) & " Actualizado..")
                    End If
                End If
                rstPago.Close
            End If
        Next
        Set rstPago = Nothing
        cnnConexion.Execute "DELETE * FROM tdfCheques WHERE IdRecibo='" & strID & "' AND FPago='EFECTIVO'", N
        If N > 0 Then Call rtnBitacora(N & " Eliminado pago en efectivo")
        '
    End If
'   ----------------------------------------------------------------------------------------------
    Do Until .EOF   'Mientras no sea el fin del archivo
'
       If .Fields(10) = strCodPC Or .Fields(10) = strCodAbonoCta Or .Fields(10) = strCodRCheq Then
'          -----------------------------------------Inicializa otras variables locales
           VecInforma(0) = VecInforma(0) + .Fields(13)         'Total Cancelado
           VecInforma(1) = VecInforma(1) + 1                   '# Recibos cancelados
'          -------------------------------------------------------------------------------------------------
'           Se actualiza la tabla Facturas del Propietario especificado
            Dim datPeriodo As Date
            '
            datPeriodo = Format("01-" & .Fields(12), "MM-DD-YY")
            '
            cnnConexion.Execute "UPDATE Factura IN '" & gcPath + strCarpeta + "Inm.mdb" _
            & "' SET Pagado=Pagado - '" & .Fields(13) & "', Saldo=Saldo + '" & .Fields(13) _
            & "', freg=DATE(), Usuario='" & gcUsuario & "', Fecha=Format(Time(),'hh:mm:ss')" _
            & " WHERE codprop='" & StrApto & "' And periodo=#" & datPeriodo & "#"
            '
            'elimina las deducciones relacinadas a esta factura
            '
            cnnConexion.Execute "DELETE * FROM DetFAct IN '" & gcPath + strCarpeta + "Inm.mdb' " _
            & " WHERE Usuario & Fecha = (SELECT Usuario & Freg FROM MovimientoCaja WHERE IDRec" _
            & "ibo='" & strID & "') AND Periodo=#" & datPeriodo & "# AND Monto < 0", Reg
            '
            cnnConexion.Execute "DELETE * FROM Cpp WHERE FRecep =(SELECT Freg FROM Movimien" _
            & "toCaja WHERE IdRecibo ='" & strID & "') AND Tipo='**' AND Fact ='" & _
            Left(strID, 7) & "' AND Estatus ='PAGADO'", Reg
                
            If Reg = 0 Then
                'elimina cpp si genero una factura
                cnnConexion.Execute "DELETE * FROM Cpp WHERE FRecep =(SELECT Freg FROM Movimien" _
                & "toCaja WHERE IdREcibo ='" & strID & "') AND Tipo='**' AND Fact ='" & _
                Left(strID, 7) & "'", Reg
                If Reg > 0 Then Call rtnBitacora("Factura ** " & Left(strID, 7) & " eliminada...")
                
            Else
            
                MsgBox gcUsuario & ", la factura correspondiente a honorarios de abogado no se " _
                & " ha eliminado de la tabla cuentas por pagar", vbInformation, App.ProductName
                Call rtnBitacora("Factura de honorarios está pagada")
                
            End If
            '
            Call rtnBitacora("Actualiza Factura Período " & .Fields(12))
            Call RtnProUtility("Actualizando Facturas propietario " & StrApto & " Periodo " & _
            .Fields(12), 4000)
            
            If .Fields(10) = strCodRCheq Then  'Desmarca el cheque recuperado
                cnnConexion.Execute "UPDATE ChequeDevuelto IN '" & gcPath + strCarpeta & _
                "Inm.mdb' SET Recuperado=Not Recuperado WHERE NumCheque = '" _
                & Mid(.Fields(8), 5, 6) & "'"
            ElseIf .Fields(10) = strCodAbonoCta Then
                cnnConexion.Execute "DELETE * FROM DetFact IN '" & gcPath + strCarpeta + "Inm.m" _
                & "db' WHERE CodGasto='" & strCodAbonoCta & "' AND Fecha=#" & _
                Format(.Fields("FechaMovimientoCaja"), "mm/dd/yyyy") & "# AND Codigo='" & _
                StrApto & "';"
            End If
            
       ElseIf .Fields(10) = strCodAbonoFut Then    'Abono a futuro
'       -----------------------------------------Inicializa otras variables locales
            VecInforma(0) = VecInforma(0) + .Fields(13)         'Total Cancelado
            VecInforma(1) = VecInforma(1) + 1                   '# Recibos cancelados
'           Elimina el abono a futuro de TDFAbonos
            cnnConexion.Execute "DELETE * FROM TDFAbonos WHERE IDRecibo='" & strID & "'"
        End If
        .MoveNext
'
    Loop
    
    .Close
'
End With
'
'Incrementa Deuda Propietario/Deuda del Inmueble/'Recibos cancelados
'
If strCod <> strCodIV And strCod <> "" And strCod <> strCodCCHeq Then
'
    Call RtnProUtility("Actualizando Deuda Inmueble '" & StrInmueble _
    & "' Propietario '" & StrApto & "'", 5000)
    Call rtnBitacora("Actualiza Propietario / Inm Bs. " & Format(VecInforma(0), "#,##0.00"))
'
    cnnConexion.Execute "UPDATE Propietarios IN '" & gcPath + strCarpeta + "Inm.mdb' SET Deuda=" _
    & "Deuda + '" & VecInforma(0) & "', Recibos=Recibos + " & VecInforma(1) & " WHERE Codigo='" _
    & StrApto & "'"
    '-------------------------
    cnnConexion.Execute "UPDATE Inmueble SET Deuda = Deuda + '" & VecInforma(0) _
    & "' WHERE CodInm ='" & StrInmueble & "'"
    '------------------------
'
End If
'Consultamos los pagos del propietario sin tomar en cuenta el presente
Call RtnProUtility("Actualizando Pagos Anteriores", 5500)
objRst.Open "SELECT * From MovimientoCaja WHERE IDRecibo <>'" & strID & "' And IDRecibo Like '" _
& Left(strID, 6) & "%' ORDER BY IndiceMovimientoCaja DESC;", cnnConexion, adOpenKeyset, _
adLockReadOnly
'
'
If Not objRst.EOF Then  ' Se Actualiza Fecha y Monto de último pago del propietario
    '
    With objRst
        .MoveFirst
        cnnConexion.Execute "UPDATE Propietarios IN '" & gcPath + strCarpeta + "Inm.mdb' SET Fe" _
        & "cUltPag ='" & !FechaMovimientoCaja & "',UltPago= '" & !montomovimientocaja & "' , Fe" _
        & "cReg='" & Date & "',Usuario='" & gcUsuario & "' WHERE codigo='" & StrApto & "'"
        .Close
    End With
'
End If
'
'Elimina el registro del movimiento de caja y del TDFABONO
Call RtnProUtility("Finalizando Proceso", 6015)
cnnConexion.Execute "DELETE * FROM Deducciones WHERE IDPeriodos IN (SELECT IDPeriodos FROM Peri" _
& "odos WHERE IDRecibo ='" & strID & "')"    'elimina las deducciones de este pago
cnnConexion.Execute "DELETE * FROM MovimientoCaja WHERE IDRecibo='" & strID & "'"
cnnConexion.Execute "DELETE * FROM tdfAbonos WHERE IDRecibo='" & strID & "'"
cnnConexion.Execute "DELETE * FROM Recibos_Enviar WHERE IDRecibo='" & strID & "'"
'
rtnCerrar:
    If Err.Number <> 0 Then     'Si ocurre un error durante el proceso
'   ------------------------------
            cnnConexion.RollbackTrans
            MsgBox Err.Description & vbCrLf & "No se ha llevado con éxtio el próceso, pongase e" _
            & "n contacto con el administrador del sistema", vbCritical, Err.Number
    Else        'El proceso fué efectuado con éxito
'   -------------------------------
            Call rtnBitacora("Transacción #" & strID & " Eliminada...")
            cnnConexion.CommitTrans
    End If
    Set objRst = Nothing    'Destruye el ADODB.Recordset creado en este apartado
    Unload FrmUtility
    '
End Sub

    '20/08/2002----Rutina que resume la presentacion del Form Utility a tracvés de cualquier proc
    Public Sub RtnProUtility(strAccion$, Optional intValor%)
    '-----------------------el valor de la barra de progreso y el texto que muestra la barra de e
        If intValor > 6015 Then intValor = 6015
        With FrmUtility
            .Label1(2).Width = intValor    'ancho de la etiqueta de llenado
            .Label1(1) = strAccion   'Texto de acción
            .Label1(1).Refresh
            .Label1(2).Refresh
        End With
    '
    End Sub

    '20/08/2002 ----------------------Configura la presentación inicial del form Utility----------
    Public Sub RtnConfigUtility(lbl_Visible As Boolean, strTitulo$, strAccion$, strMsg$)
    '---------------------------------------------------------------------------------------------
    '
    With FrmUtility
        CenterForm FrmUtility
        .Show vbModeless, FrmAdmin
        For I = 0 To 2: .Label1(I).Visible = lbl_Visible
        Next
        .Label1(0) = strMsg
        .Caption = strTitulo
        .Label1(1) = strAccion
        .Refresh
    End With
    '
    End Sub

    '03/09/2002-----------------------------------------------------------------------------------
    '   Rutina:     RtnFlex
    '
    '   Entradas:   Codigo de Apartamento(strApto), Grid que muestra la informacion
    '               (Grid),intMora porcentaje int.mora, intMesMora a partir del mes
    '               que se cobra intereses, C ,txtMora control TextBox que muestra total int.
    '               cnnApto Conexion al origen de datos
    '
    '   Rutina que distribuye en el grid las facturas pendientes de pago, calculo honorarios
    '   de abogado y la deuda total. Además verifica el estatus del propietario, demandado,
    '   convenio, etc.
    '---------------------------------------------------------------------------------------------
    Public Sub RtnFlex(StrApto$, Grid As Control, intMesMora%, intMora%, C%, txtmora As TextBox, _
    cnnApto As ADODB.Connection, Optional Inm As String, Optional BsF As Boolean)
    'variales locales
    Dim Rfacturas As ADODB.Recordset
    Dim curDeuda@, strSQL$
    Dim cFactura@, cAbona@, cSaldo@
    Dim nFactor&
    
    nFactor = IIf(BsF, 1000, 1)
    '
    Set Rfacturas = New ADODB.Recordset
    '
    Rfacturas.Open "SELECT Factura.FACT, Factura.Periodo, Factura.Facturado, Factura.Pagado, " _
    & "Factura.Saldo From Factura WHERE (((Factura.codprop)='" & StrApto & "') AND " _
    & "((Factura.Saldo)>0.009)) ORDER BY Factura.Periodo", cnnApto, adOpenKeyset, adLockOptimistic
    '---------------------------------------------------------------------------------------------
    If Rfacturas.EOF Then  'Si no tiene información sale de la rutina
        Rfacturas.Close
        Set Rfacturas = Nothing
        Exit Sub
    End If
    Rfacturas.MoveFirst
    '
    With Grid
    '   Configura la presentación del Grid(Titulos, Ancho de columna, N° de Filas)
        .Rows = Rfacturas.RecordCount + 1
        .Cols = C
    '
    End With
    I = 1
    '
    curDeuda = 0
    
    Do While Not Rfacturas.EOF  'Hacer hasta el final del archivo
    '
        cFactura = Rfacturas("FACTURADO") * nFactor
        cAbono = Rfacturas("PAGADO") * nFactor
        cSaldo = Rfacturas("SALDO") * nFactor
        
        With Grid
    '
            .Col = 1
            .Row = I
            .CellAlignment = flexAlignCenterCenter
            .TextMatrix(I, 0) = IIf(IsNull(Rfacturas("fact")), "", Rfacturas("fact"))
            .TextMatrix(I, 1) = IIf(IsNull(Rfacturas("PERIODO")), "", _
            Format(Rfacturas("PERIODO"), "MM-YY"))
            .TextMatrix(I, 2) = Format(cFactura, "##,##0.00")
            .TextMatrix(I, 3) = Format(cAbono, "##,##0.00")
            .TextMatrix(I, 4) = Format(cSaldo, "##,##0.00")
            curDeuda = curDeuda + cSaldo
    '
            If I = 1 Then
                .TextMatrix(I, 5) = .TextMatrix(I, 4)
            Else
                .TextMatrix(I, 5) = Format(CCur(.TextMatrix(I - 1, 5)) + _
                CCur(.TextMatrix(I, 4)), "#,##0.00")
            End If
    '
            Rfacturas.MoveNext
            I = I + 1
    
        End With
    '
    Loop    'Punto de control {fin hasta}
    Grid.Col = 0
    Grid.ColSel = Grid.Cols - 1
    If Grid.Enabled And Grid.Visible Then Grid.SetFocus
    'Busca Honorarios-----------------------------------------------------------------------------
    If Rfacturas.RecordCount > intMesMora Then
        txtmora = Format(Round(curDeuda * intMora / 100), "#,##0.00")
    Else
        txtmora = "0,00"
    End If
    'aqui busca la información llama el formulario del convenio
'    strSql = "SELECT Convenio.*, Convenio_Detalle.* FROM Convenio INNER JOIN Convenio_Detalle O" _
'    & "N Convenio.IDConvenio = Convenio_Detalle.IDConvenio WHERE Convenio.CodInm='" & Inm & _
'    "' AND Convenio.CodProp ='" & StrApto & "' AND IDStatus = 1"
'    Rfacturas.Close
'    Rfacturas.Open strSql, cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
'    '
'    If Not Rfacturas.EOF And Not Rfacturas.BOF Then
'    '   'revisar estas lineas de código
'        frmConsultaConvenio.Inmueble = Inm
'        frmConsultaConvenio.apto = StrApto
'        frmConCon.Show vbModal, FrmAdmin
'
'    End If
    '
    '---------------------------------------------------------------------------------------------
    'CIERA Y DESCARGA LA CONEXION Y EL ADODB.Recordset
    Rfacturas.Close
    Set Rfacturas = Nothing
    '
    End Sub
    
    '12/09/2002.---Actualiza el menu Caja de acuerdo el estado de la caja Abierta / Cerrada-------
    Public Sub rtnEstadoCaja(strOpen As String) '
    '---------------------------------------------------------------------------------------------
    '
    With FrmAdmin   'Caja Abierta
        '
        .AC400(0).Checked = strOpen
        .AC400(1).Enabled = strOpen
        .AC400(3).Enabled = strOpen
        .AC400(4).Enabled = strOpen
        .AC400(6).Enabled = strOpen
        .AC400(8).Enabled = strOpen
        .AC400(9).Enabled = strOpen
        .AC400(12).Enabled = strOpen
        '
    End With
    '
    End Sub

    '13/09/2002-----------------------------------------------------------------------------------
    '   Rutina:     rtnGenerator
    '
    '   Entradas:   directorio de Trababajo, Instrucción SQL, Nombre de
    '               la consulta que se va a generar,
    
    '   Procedimiento que genera una consulta temporal siguiendo los parametos
    '   de entrada dados por el usuario
    '---------------------------------------------------------------------------------------------
    Public Sub rtnGenerator(strPath$, strQDF$, strName$) '
    '
    Dim N As Long
    Dim dbSac As Database
    Dim qdfTemp As QueryDef
    '
    Set dbSac = OpenDatabase(strPath)
    'Si existe una consulta con nombre similar la elimina
    For Each qdfTemp In dbSac.QueryDefs
    
        If UCase(qdfTemp.Name) = UCase(strName) Then
            dbSac.QueryDefs.Delete (qdfTemp.Name)
            Exit For
        End If
        
    Next
    'Crea la consulta con los parametros indicados en strQDF
    N = dbSac.QueryDefs.Count
    Set qdfTemp = dbSac.CreateQueryDef(strName, strQDF)
    Do
    dbSac.QueryDefs.Refresh
    DoEvents
    Loop Until dbSac.QueryDefs.Count = N + 1
    qdfTemp.Close
    Set qdfTemp = Nothing
    dbSac.Close
    Set dbSac = Nothing
    '
    End Sub

    '-----------------------------------------------------------
    ' SUB: CenterForm
    '
    ' Centra el Formulario en pantalla
    '-----------------------------------------------------------
    '
    Sub CenterForm(Frm As Form)
    Frm.Top = (FrmAdmin.Height * 0.85) \ 2 - Frm.Height \ 2
    Frm.Left = FrmAdmin.Width \ 2 - (Frm.Width \ 2)
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    '   Rutina:     rtnLimpiar_Grid
    '
    '   Borra el contenido de todas las celdas no fijas del MSFlexGrid
    '---------------------------------------------------------------------------------------------
    Public Sub rtnLimpiar_Grid(msGrid As Control)
    'variables locales
    With msGrid
        '
        .Rows = 2
        If .Rows = 1 Then Exit Sub
        .Col = 0
        .Row = 1
        .ColSel = msGrid.Cols - 1
        .RowSel = msGrid.Rows - 1
        .FillStyle = flexFillRepeat
        .Text = ""
        .FillStyle = flexFillSingle
        .ColSel = 0
        .RowSel = 1
        '
    End With
    '
    End Sub
    
    Sub Main()
    
    'valida que no se este ejecutando una nueva instancia de la aplicacion
    If App.PrevInstance Then
        MsgBox "Ya se está ejecutando " & App.ProductName, vbCritical, App.ProductName
        Exit Sub
    End If
    'buscamos actualizacion del programa de actualización de la aplicacion
    'Call verificar_actualizador
    'ejecutamos la aplicacion de actualización
    PID = Shell(App.Path & "\sacUpdate.exe", vbNormalFocus)
    If PID <> 0 Then
        'Esperar a que finalice
        WaitForTerm PID
    End If
    'rutina de inicio del programa
    On Error Resume Next
    'variables locales
    Dim FechaEntorno As Date
    Dim strArchivo As String
    
    DataServer = GetSetting(App.EXEName, "Conexion", "Server")
    SaveSetting App.EXEName, "Entorno", "Proveedor", cnnOLEDB
    
    '
10 If DataServer = "" Then

        DataServer = InputBox$("Por favor introduzca el nombre del servidor donde se encuentra " _
        & "la base de datos del sistema", App.ProductName)
        
    End If
    If DataServer = "" Then
        MsgBox App.ProductName & " se cerrará"
        End
    End If
    gcPath = GetSetting(App.EXEName, "Entorno", "Ruta", "")
    '
    If Dir$(gcPath & "\sac.mdb") = "" Then
        
        gcPath = InputBox$("Introduzca la ruta donde está guardada la base de datos principal. " _
        & vbCrLf & "Si ud. no está seguro consulte con el proveedor del sistema", App.ProductName, "C:\sac\datos")
        
        If Dir$(gcPath & "\sac.mdb") = "" Then
        
            MsgBox "Ruta inválida. " & App.ProductName & " se cerrará", vbCritical, _
            App.ProductName
            End
        Else
            SaveSetting App.EXEName, "Entorno", "Ruta", gcPath
        End If
    End If
    '
    strArchivo = App.Path & Archivo_Temp
    '
    Load frmAcceso
    On Error GoTo Cerrar
    'Configura la conexion general a la BD
    '
    Dim AdoAmbiente As ADODB.Recordset
    '
    'asigna las variables de entorno
    Set AdoAmbiente = New ADODB.Recordset
    AdoAmbiente.Open "Ambiente", cnnConexion, adOpenKeyset, adLockOptimistic, adCmdTable
    
    gnIva = AdoAmbiente("IVA")
    gnIDB = AdoAmbiente("IDB")
    IDMoneda = AdoAmbiente("IDMoneda")
    SaveSetting App.EXEName, "Conexion", "Server", AdoAmbiente("ServidorH")
    '
    If gcPath <> AdoAmbiente("Ruta") Then
        'gcPath = AdoAmbiente("Ruta")
        If gcNivel = nuADSYS Then
        
            'gcPath = AdoAmbiente("Ruta")
            SaveSetting App.EXEName, "Entorno", "Ruta", gcPath
            AdoAmbiente.Update "ruta", gcPath
            msg = "Se ha ajustado la ruta de acceso a datos. Consulte los parámetros administra" _
            & "tivos del sistema" & vbCrLf & "'" & gcPath & "'"
            
        Else
        
            msg = "Consulte con el administrador del sistema. Se ha reajustado la ruta de acces" _
            & "o a datos"
            
        End If
        MsgBox msg, vbInformation, App.ProductName
    End If
    
    '
    '********************************************************
    '   este apartado lo uso para ejecutar instrucciones
    '   sql para ajustar la base de datos en forma remota
    '********************************************************
'    cnnConexion.Execute "delete from tdfcheques where monto = 0", N
'    Call rtnBitacora("Borrados (" & N & ") registros en cero de la tabla TDFCheques")
    Dim sql As String
'    gcMAC = "REMOTO"
'    gcUsuario = "SUPERVISOR"
'
''
'    sql = "PARAMETERS [Cuentas] Text ( 255 ), [IDCuenta] Long, [Inmueble] Text ( 255 );" & _
'        "SELECT [IDCuenta] AS IDCuenta, Cheque.FechaCheque as Fecha, 'CHQ' AS IDTipoMov, Cheque.Beneficiario, " & _
'        "Cheque.IDCheque, 0 AS Debe, Cheque_Total.Total_Cheque AS Haber, Cheque.Hora " & _
'        "FROM Cheque INNER JOIN Cheque_Total ON Cheque.IDCheque = Cheque_Total.IDCheque " & _
'        "WHERE (((Cheque.Cuenta)=[Cuentas])) " & _
'        "Union All SELECT [IDCuenta], ChequeAnulado.FechaCheque, 'CHQ', '(CHQ. ANULADO) '+Beneficiario, " & _
'        "ChequeAnulado.IDCheque, 0, 0, ChequeAnulado.Hora " & _
'        "FROM ChequeAnulado Where (((Cuenta) = [Cuentas])) " & _
'        "Union All " & _
'        "SELECT [IDCuenta], MovimientoCaja.FechaMovimientoCaja,  Left(TDFCheques.Fpago,3), " & _
'        "MovimientoCaja.CuentaMovimientoCaja  & ' APTO: ' &  iif(MovimientoCaja.MontoMovimientocaja < TDFCheques.Monto, 'VARIOS',MovimientoCaja.AptoMovimientoCaja), " & _
'        "TDFCheques.Ndoc, TDFCheques.Monto, 0, MovimientoCaja.Hora " & _
'        "FROM MovimientoCaja RIGHT JOIN TDFCheques ON MovimientoCaja.IDRecibo = TDFCheques.IDRecibo " & _
'        "WHERE TDFCheques.CodInmueble & TDFCheques.IDDeposito=[Inmueble] & [IDCuenta] AND TDFCheques.Fpago<>'EFECTIVO' AND " & _
'        "TDFCheques.Fpago<>'CHEQUE' " & _
'        "Union All " & _
'        "SELECT [IDCuenta],MovimientoCaja.FechaMovimientoCaja, Left(TDFCheques.Fpago,3) , " & _
'        "MovimientoCaja.CuentaMovimientoCaja & ' APTO: ' &  iif(MovimientoCaja.MontoMovimientocaja < TDFCheques.Monto, 'VARIOS',MovimientoCaja.AptoMovimientoCaja), " & _
'        "TDFCheques.Ndoc, TDFCheques.Monto, 0, MovimientoCaja.Hora " & _
'        "FROM TDFDepositos RIGHT JOIN (MovimientoCaja RIGHT JOIN TDFCheques ON MovimientoCaja.IDRecibo " & _
'        "= TDFCheques.IDRecibo) ON TDFDepositos.IDDeposito = TDFCheques.IDDeposito " & _
'        "WHERE TDFCheques.CodInmueble & TDFDepositos.Cuenta=[Inmueble]&[Cuentas] AND (TDFCheques.Fpago='EFECTIVO' " & _
'        "or TDFCheques.Fpago='CHEQUE');"
'
'    rtnGenerator gcPath & "\sac.mdb", sql, "procLibroBanco"
'
'    Call rtnBitacora("Creada procedimiento procLibroBanco")

'    sql = "PARAMETERS [FECHA] DateTime; " & _
'        "SELECT Sum(procLibroBanco.Debe) - Sum(procLibroBanco.Haber) AS Saldo " & _
'        "from procLibroBanco " & _
'        "WHERE (((procLibroBanco.Fecha)<=[FECHA]));"
'
'    rtnGenerator gcPath & "\sac.mdb", sql, "procLibroBancoSaldoFecha"
'
'    Call rtnBitacora("Creado procedimiento procLibroBancoSaldoFecha")

    
    'aqui vamos a actualizar la tabla TDFCheques
'    SELECT TOP 1 * from (SELECT TOP 1 *
'from tdfcheques
'WHERE (((tdfcheques.IDDeposito)<>'') AND ((tdfcheques.Fpago)='deposito' Or (tdfcheques.Fpago)='transferencia'))
'ORDER BY FechaMov DESC) as p
'
'    sql = "PARAMETERS [RECIBO] Text ( 255 ); " & _
'          "SELECT Fpago, Ndoc, Banco, FechaDoc, Monto " & _
'          "from TDFCheques " & _
'          "Where (((TDFCheques.IDRecibo) = [Recibo])) " & _
'          "ORDER BY TDFCheques.Fpago;"
'
'    rtnGenerator gcPath & "\sac.mdb", sql, "procDetallePago"
'    Call rtnBitacora("Creada Consulta procBuscaCaja")
'-------validamos los cheques recibidos y procesados en caja
'    Dim Monto As Double
'    Dim rst As ADODB.Recordset
'    Dim rsm As ADODB.Recordset
'    Dim campo As String
'    Dim boIgual As Boolean
'    sql = "update propietarios in '" & gcPath & "\2530\inm.mdb' set cedula=9964627  where id=6"
'    cnnConexion.Execute sql, N
'    sql = "update propietarios in '" & gcPath & "\2523\inm.mdb' set cedula=15664484  where id=9"
'    cnnConexion.Execute sql, N
'    sql = "update tdfcheques set idrecibo='66B011180609401' where idrecibo='66B01118060901'"
'    cnnConexion.Execute sql, N
'    sql = "update movimientocaja set montomovimientocaja = montocheque ,TipoMovimientoCaja='INGRESO'" & _
'        "where montomovimientocaja <= 0 and FormaPagoMovimientoCaja <> 'EFECTIVO' and " & _
'        "montocheque > 0 and codgasto <> '900011'"
'    cnnConexion.Execute (sql), N
'
'    sql = "select sum(montomovimientocaja) from movimientocaja group by InmuebleMovimientoCaja having InmuebleMovimientoCaja='2566'"
'    Monto = cnnConexion.Execute(sql).Fields(0)
'    Debug.Print "Monto cobrado en caja: " & Format(Monto, "#,##0.00")
'    sql = "select sum(Monto) from tdfcheques group by CodInmueble having CodInmueble ='2566'"
'    Monto = cnnConexion.Execute(sql).Fields(0)
'    Debug.Print "Monto cheques registrados: " & Format(Monto, "#,##0.00")
'
'    sql = "select * from tdfcheques where codinmueble='2566'"
'    Set rst = cnnConexion.Execute(sql)
'    With rst
'        Do
'            sql = "select * from movimientocaja where inmueblemovimientocaja='2566' and idrecibo='" & _
'            rst("idrecibo") & "'"
'            Monto = cnnConexion.Execute(sql).Fields("montomovimientocaja")
'            If Monto <> rst("monto") Then
'                'Stop
'                'si el monto movimiento caja es mayor al cheque registrado
'                'revisamos cada documento recibido
'                boIgual = False
'                If Monto > rst("monto") Then
'                    For I = 0 To 3
'                        If I < 3 Then
'                            campo = "MontoCheque" & I
'                        Else
'                            campo = "EfectivoMovimientoCaja"
'                        End If
'                        campo = Replace(campo, "0", "")
'                        sql = "select " + campo + " from movimientocaja where " & _
'                                "inmueblemovimientocaja='2566' and idrecibo='" & _
'                                rst("idrecibo") & "'"
'
'                        Monto = cnnConexion.Execute(sql).Fields(0)
'                        'buscamos cualquier otro abono que tenga este cheque
'                        sql = "select montomovimientocaja from movimientocaja where inmueblemovimientocaja='2566' and idrecibo <>'" & rst("idrecibo") & "' and fechamovimientocaja=#" & Format(rst("fechamov"), "yyyy/mm/dd") & "# and numdocumentomovimientocaja ='" & rst("ndoc") & "'"
'                        If Not (cnnConexion.Execute(sql).EOF And cnnConexion.Execute(sql).BOF) Then
'                            Monto = Monto + cnnConexion.Execute(sql).Fields(0)
'                        End If
'                        If Monto = rst("monto") Or (Monto - rst("monto")) < 0.5 Then
'                            boIgual = True
'                            Exit For
'                        End If
'                    Next
'
'                Else
'                    'buscamos en la caja el monto faltante del cheque
'                    sql = "select * from movimientocaja where inmueblemovimientocaja='2566' " & _
'                          "and FechaMovimientoCaja=#" & Format(rst("FechaMov"), "YYYY/MM/DD") & "# and " & _
'                          " NumDocumentoMovimientoCaja = '" & _
'                          rst("Ndoc") & "'"
'                          'IDRecibo <> '" & rst("idrecibo") & "' and
'                    Set rsm = cnnConexion.Execute(sql)
'                    If Not (rsm.EOF And rsm.BOF) Then
'                        Monto = 0
'                        Do
'                            Monto = Monto + rsm("montomovimientocaja") - rsm("efectivomovimientocaja") - rsm("montocheque1")
'                            boIgual = Monto = rst("monto")
'                            rsm.MoveNext
'                        Loop Until rsm.EOF
'                    End If
'                    rsm.Close
'                    Set rsm = Nothing
'                End If
'                If Not boIgual Then
'                    Debug.Print "No cuadra recibo #" & rst("idrecibo")
'                End If
'            End If
'            .MoveNext
'        Loop Until .EOF
'    End With
    
'--------fin validacion
    
    gcMAC = ""
    gcUsuario = ""
        
'    '********************************************************
    AdoAmbiente.Close
    'aqui inicializa las variales de la empresa (sistema)
    AdoAmbiente.Open "Empresa", cnnConexion, adOpenKeyset, adLockReadOnly, adCmdTable
    sysEmpresa = IIf(IsNull(AdoAmbiente("nombre")), "", AdoAmbiente("nombre"))
    sysCodCaja = IIf(IsNull(AdoAmbiente("codcaja")), "", AdoAmbiente("codcaja"))
    sysCodInm = IIf(IsNull(AdoAmbiente("codinm")), "", AdoAmbiente("codinm"))
    sysCodPro = IIf(IsNull(AdoAmbiente("codprov")), "", AdoAmbiente("codprov"))
    AdoAmbiente.Close
    Set AdoAmbiente = Nothing
    
    If sysEmpresa = "" Or sysCodCaja = "" Or sysCodInm = "" Or sysCodPro = "" Then
        MsgBox "Falta algunos datos de la empresa, es posible que el sitema no funcione correctamente. Consulte con el administrador del sistema", vbInformation, App.ProductName
        Call rtnBitacora("Faltan datos de la empresa")
    End If
Cerrar:
    If Err.Number = -2147467259 Then
        MsgBox "Estimado Usuario su estación de Trabajo esta fuera de la red." & vbCrLf & vbCrLf _
        & "Para resolver el Problema intente lo siguiente:" & vbCrLf & "1.- Cierre la sesión (d" _
        & "esde el boton 'INICIO')" & vbCrLf & "2.- Reinicie la sesión y presione 'enter' en la" _
        & " ventana 'Contrasena de Red'" & vbCrLf & vbCrLf & "Si el problema persiste Consulte " _
        & "al Administrador del Sistema", vbInformation, App.ProductName
        End
    Else
        frmAcceso.fraAcceso(0).Visible = True
        frmAcceso.fraAcceso(3).Visible = True
    End If
    '
    End Sub
    
    Sub Print_PreRecibo(ByVal Periodo As Date)
    'VARIABLES LOCALES
    Dim strQuery As String
    '
    strQuery = "SELECT CodGasto, Cargado, Descripcion, Sum(Monto) AS Total, Cstr(Time()) as Hor" _
    & "a From AsignaGasto Where (((Comun) = True)) GROUP BY CodGasto, Cargado, Descripcion Havi" _
    & "ng (((Cargado) = #" & Periodo & "#) AND ((CodGasto)<>'" & gcCodFondo & "')) ORDER BY Cod" _
    & "Gasto;"
    Call rtnGenerator(mcDatos, strQuery, "qdfPreRecibo")
    '
    End Sub
    
    
    '---------------------------------------------------------------------------------------------
    '   Function:   Obtener_Nota
    '
    '   Entrada:    Ubicacion del archivo strUbica
    '
    '   Obtiene, si los tiene, del archivo notas.txt del inmueble un valor cadena que
    '   aparecerá en todos los avisos en el area de "nota importante"
    '---------------------------------------------------------------------------------------------
    Public Function Obtener_Nota(Folder As String) As String
    'variables locales
    Dim numFichero As Integer
    Dim strArchivo As String
    Dim strCadena As String
    '
    On Error GoTo Cerrar
    numFichero = FreeFile
    strArchivo = gcPath & Folder & "notas.txt"
    Open strArchivo For Input As numFichero
    Do Until EOF(numFichero)
        Input #numFichero, strCadena
        Obtener_Nota = Obtener_Nota + IIf(strCadena <> "" And Obtener_Nota <> "", ", " & strCadena, strCadena)
    Loop
    Close numFichero
Cerrar:
    If Err.Number <> 0 Then Obtener_Nota = ""
    '
    End Function

    '---------------------------------------------------------------------------------------------
    '   Rutina:     Marcar_Linea
    '
    '   Colorea la linea seleccionada   en un control MSHFlexGrid
    '---------------------------------------------------------------------------------------------
    Public Sub Marcar_Linea(ctlGrid As Control, color As Long)
    'variables locales
    '
    Dim intCol As Integer
    '
    With ctlGrid
    '
        intCol = .ColSel
        .Row = .RowSel
        .Col = 0
        .ColSel = .Cols - 1
        .FillStyle = flexFillRepeat
        .CellBackColor = color
        If color = vbGreen Then
            .CellFontBold = True
        Else
            .CellFontBold = False
        End If
        .Col = intCol
        .FillStyle = flexFillSingle
    '
    End With
    '
    End Sub


    '---------------------------------------------------------------------------------------------
    '
    '   Funcion Publica:    Respuesta
    '
    '   Devuelve verdadera si el usuario contesta si, de lo contrario devuelve falso
    '---------------------------------------------------------------------------------------------
    Public Function Respuesta(strQ As String) As Boolean
    Dim Pregunta As Integer
    Pregunta = MsgBox(strQ, vbYesNo + vbQuestion, App.ProductName)
    Respuesta = IIf(Pregunta = vbYes, True, False)
    End Function

    '---------------------------------------------------------------------------------------------
    '  function Total_Fondo
    '
    '   Entrada:        strInm / Codigo del Inmueble
    '                         strPer / Período facturación formato cadena
    '
    '   Devuelve el monto total del fondo de reserva de un inmueble
    '---------------------------------------------------------------------------------------------
    Public Function Total_Fondo(strInm As String, ByVal strPer As String) As Long
    'variables locales
    Dim strSQL As String
    Dim rstlocal As New ADODB.Recordset
    
    strSQL = "SELECT Pago_Inf.*,TGastos.Titulo FROM Pago_INF INNER JOIN TGastos ON " _
    & "Pago_Inf.CtaFondo=Tgastos.CodGasto WHERE Periodo=#" & strPer & "# AND " _
    & "CtaFondo='" & gcCodFondo & "'"
    '
    rstlocal.Open strSQL, cnnOLEDB + mcDatos, adOpenKeyset, adLockOptimistic, adCmdText
    
    If Not rstlocal.EOF And Not rstlocal.BOF Then
        Total_Fondo = rstlocal("SA") + rstlocal("CR") - rstlocal("DB")
    End If
    '
    rstlocal.Close
    Set rstlocal = Nothing
    '
    End Function

    
    '---------------------------------------------------------------------------------------------
    '   Rutina:     Printer_PaqueteCompleto
    '
    '   Entradas:   Periodo Facturado, control de Crystal Report,Total Factuación
    '               Salida del reporte, variable F (si se está ejecutando el proceso facturar)
    '
    '   Imprime todos los reportes del módulo de facturación
    '---------------------------------------------------------------------------------------------
    Public Sub Printer_PaqueteCompleto(ByVal datFacturado$, _
    ByVal curTotal@, Optional Salida As crSalida, Optional F As Boolean)
    '
    'Variables locales
    Dim strPeriodo As String
    Dim Direcciona As String
    Dim strRep As String
    Dim intPos As Integer
    Dim rpReporte As ctlReport, errLocal As Long
    '
    MousePointer = vbHourglass
    strPeriodo = UCase(Format(Format(CDate(datFacturado), "mm/dd/yyyy"), "MMM - yyyy"))
    If F Then
        FrmEmisionFactura.lblFactura(0) = "Imprimiendo el Análisis de Facturación"
        FrmEmisionFactura.lblFactura(1).Width = 9200
        For I = 0 To 1: FrmEmisionFactura.lblFactura(I).Refresh
        Next
    End If
    'Imprime el análisis de facturación
    Call Printer_Analisis_Facturacion(strPeriodo, datFacturado, curTotal, Salida)
    If F Then   'si está facturando msg en pantalla
        FrmEmisionFactura.lblFactura(0) = "Imprimiendo el Pre-Recibo"
        FrmEmisionFactura.lblFactura(1).Width = 9400
        For I = 0 To 1: FrmEmisionFactura.lblFactura(I).Refresh
        Next
    End If
    '
    'imprime el pre_recibo
    Call Printer_PreRecibo(strPeriodo, datFacturado, 1, Salida) 'Pre-Recibo
    If F Then   'si está fact. msg en pantalla
        FrmEmisionFactura.lblFactura(0) = "Imprimiendo Reporte Gastos No Comunes"
        FrmEmisionFactura.lblFactura(1).Width = 9500
        For I = 0 To 1: FrmEmisionFactura.lblFactura(I).Refresh
        Next
    End If
    '
    'imprime el reporte de gastos no comunes
    Call Printer_GNC(strPeriodo, datFacturado, Salida)
    
    If F Then   'si esta facturando msg en pantalla
        FrmEmisionFactura.lblFactura(0) = "Imprimiendo la Facturación Mensual"
        FrmEmisionFactura.lblFactura(1).Width = 9700
        For I = 0 To 1: FrmEmisionFactura.lblFactura(I).Refresh
        Next
    End If
    '
    'imprime el reporte de gastos mensuales
    Call FrmAdmin.Reporte_GM(CStr(datFacturado), Salida, Guarda_Copia:=True)
    '
    'imprime la facturación mensual 1 copia
    Call Printer_Facturacion_Mensual(strPeriodo, Salida)
    '
    If F Then
        FrmEmisionFactura.lblFactura(0) = "Imprimiendo el Control de Facturación"
        FrmEmisionFactura.lblFactura(1).Width = 9900
        For I = 0 To 1: FrmEmisionFactura.lblFactura(I).Refresh
        Next
    End If
    '
    'impresión carta paquete completo
    Set rpReporte = New ctlReport
    With rpReporte
            '
        '.Reset
        '.ProgressDialog = False
        strRep = "Carta Paquete Completo"
        .Reporte = gcReport + "fact_pm.rpt"
        .OrigenDatos(0) = mcDatos
        .FormuladeSeleccion = "{Propietarios.CarJunta}='PRESIDENTE'"
        .Formulas(0) = "Condominio='" & gcNomInm & "'"
        .Formulas(1) = "Fecha='" & Format(Date, "Long Date") & "'"
       .Formulas(2) = "Periodo='" & UCase(Format(Format(datFacturado, "MM/dd/yy"), "MMMM")) _
       & ",'"
        With FrmAdmin.objRst
            'If .State = 0 Then .Requery
            If Not .EOF Or Not .BOF Then
                .MoveFirst
                .Find "CodInm='" & gcCodInm & "'"
                If Not IsNull(!Direccion) Then
                    intPos = InStr(!Direccion, Chr(13))
                    Direcciona = !Direccion
                    If intPos <> 0 Then Direcciona = Left(Direcciona, intPos - 1)
                    If Not .EOF Then rpReporte.Formulas(3) = "Direccion='" & Direcciona & "'"
                End If
            End If
        End With
        .Salida = Salida
        .TituloVentana = "Carta Paquete Completo" & gcCodInm & "/" & Right(.Formulas(2), Len(.Formulas(2)) - 1)
        errLocal = .Imprimir
        GoSub Ocurre_Error
        '
    End With
    Set rpReporte = Nothing
    
    '
    'imprime el control de facturación
    Call Printer_Control_Facturacion(strPeriodo, datFacturado, Salida)
    If F Then
        FrmEmisionFactura.lblFactura(0) = "Imprimiendo Reportes Finales"
        FrmEmisionFactura.lblFactura(1).Width = 10000
        For I = 0 To 1: FrmEmisionFactura.lblFactura(I).Refresh
        Next
    End If
    '
    '
    Set rpReporte = New ctlReport
    With rpReporte
        'Impresion deuda confidencial junta
        strRep = "Deuda Confidencial"
        If Dir(gcPath & gcUbica & "Reportes\DC" & Format(datFacturado, "ddyyyy") _
        & ".rpt") = "" Then
            .Reporte = gcReport + "EdoCtaress.rpt"
            .OrigenDatos(0) = mcDatos
            '.FormuladeSeleccion = "{Propietarios.Recibos}>=4"
            .Formulas(0) = "Inmueble ='" & gcCodInm & " - " & gcNomInm & "'"
            .Salida = crArchivoDisco 'almacena un copia en la carpeta reportes del inm
            '.PrintFileType = crptCrystal
            .ArchivoSalida = gcPath & gcUbica & "Reportes\DC" & Format(datFacturado, "ddyyyy") _
            & ".rpt"
            errLocal = .Imprimir
            GoSub Ocurre_Error
        Else
            .Reporte = gcPath & gcUbica & "Reportes\DC" & Format(datFacturado, "ddyyyy") _
            & ".rpt"
        End If
        .Salida = Salida
        If .Salida = crPantalla Then .TituloVentana = "Deuda Confidencial " & gcCodInm
        errLocal = .Imprimir
        GoSub Ocurre_Error
    End With
    Set rpReporte = Nothing
    '
'    'guarda una copia de la deuda general hasta 3 meses
    Set rpReporte = New ctlReport
    With rpReporte
        strRep = "Deuda General"
        'si no ha sido creado
        If Dir(gcPath & gcUbica & "Reportes\DG" & Format(datFacturado, "ddyyyy") & _
        ".rpt") <> "" Then
'            .OrigenDatos(0) = mcDatos
'            .Reporte = gcReport + "EdoCtaInm.rpt"
'            .Formulas(0) = "Inmueble ='" & gcCodInm & " - " & gcNomInm & "'"
'            .FormuladeSeleccion = "{Factura.Saldo}<>0 and {Propietarios.Recibos}<=3"
'            'almacena una copoa del reporte
'            .Salida = crArchivoDisco
'            '.PrintFileType = crptCrystal
'            .ArchivoSalida = gcPath & gcUbica & "Reportes\DG" & Format(datFacturado, "ddyyyy") _
'            & ".rpt"
'            errLocal = .Imprimir
'            GoSub Ocurre_Error
            '
'        Else
            'Impresion deuda general hasta tres meses
            strRep = "Deuda General hasta tres meses"
            .Reporte = gcPath & gcUbica & "Reportes\DG" & Format(datFacturado, "ddyyyy") _
            & ".rpt"
'        End If
            .Salida = Salida
            If .Salida = crPantalla Then .TituloVentana = "Estado de Cuenta Inm. " & gcCodInm
            errLocal = .Imprimir
            GoSub Ocurre_Error
        End If
    End With
    Set rpReporte = Nothing
    '
    'almacena una copia de la deuda general
    'si no ha sido generada
'    Set rpReporte = New ctlReport
'    With rpReporte
'        If Dir(gcPath & gcUbica & "Reportes\DGT" & Format(datFacturado, "ddyyyy") _
'        & ".rpt") = "" Then
'            strRep = "Archivo Disco Deuda General"
'            .OrigenDatos(0) = mcDatos
'            .Reporte = gcReport + "EdoCtaInm.rpt"
'            .Formulas(0) = "Inmueble ='" & gcCodInm & " - " & gcNomInm & "'"
'            .FormuladeSeleccion = "{Factura.Saldo}<>0"
'            .Salida = crArchivoDisco
'            '.PrintFileType = crptCrystal
'            .ArchivoSalida = gcPath & gcUbica & "Reportes\DGT" & Format(datFacturado, "ddyyyy") & ".rpt"
'            errLocal = .Imprimir
'            GoSub Ocurre_Error
'        End If
'        '
'    End With
'    Set rpReporte = New ctlReport
    '
    'imprime otro reporte de facturación si se está facturando
    If F Then Call Printer_Facturacion_Mensual(strPeriodo, Salida)
    'imprime otro control de facturación si se esta facturando
    If F Then Call Printer_Control_Facturacion(strPeriodo, datFacturado, Salida)
    MousePointer = vbDefault
    Exit Sub
Ocurre_Error:
    If errLocal <> 0 Then
        MsgBox "Error al imprimir el reporte " & strRep & vbCrLf & Err.Description, vbInformation, "Error " & _
        Err.Number
        Call rtnBitacora("Error " & Err.Number & ", al imprimir " & strRep)
    End If
    Return

    
    End Sub

    '---------------------------------------------------------------------------------------------
    '   Function:   ftnFact
    '
    '   Funcion que devuelve un valor moneda que representa el total Gasto
    '   comun facturado o el fondo de Reserva facturado para un periodo
    '   determinado
    '---------------------------------------------------------------------------------------------
    Private Function ftnFact(intOpcion As Integer, Periodo As String) As Currency
    'variables locales
    Dim rstTotal As New ADODB.Recordset
    Dim CnnInm As New ADODB.Connection
    Dim Operador As String
    '-------------------------------
    CnnInm.Open cnnOLEDB & mcDatos
    If intOpcion = 0 Then
        Operador = "="
    Else
        Operador = "<>"
    End If
    rstTotal.Open "SELECT Sum(CCur(Monto)) as Total FROM AsignaGasto WHERE Comun=True AN" _
    & "D CodGasto " & Operador & "'" & gcCodFondo & "' AND Cargado=#" & Periodo & "# AND Not Is" _
    & "Null(Monto);", CnnInm, adOpenKeyset, adLockOptimistic
    ftnFact = IIf(IsNull(rstTotal!Total), 0, rstTotal!Total)
    Set rstTotal = Nothing
    Set CnnInm = Nothing
    '
    End Function


    Public Sub rtnMakeQuery(Periodo As Date, strGA$, strGM$, strGes$, strNotIn$)
        '---------------------------
    'Crea sendas consultas necesarios en el proceso siguiente
    Dim strSQL$, strFact$
    Dim vecGasto(2 To 5) As String
    
    'Consulta gastos comunes
    strSQL = "SELECT Codigo, Periodo, Clng(Sum(Monto)*100)/100 AS GC,Hora,Fecha FROM DetFact WHERE Detalle" _
    & " In (SELECT Descripcion FROM AsignaGasto WHERE Comun=True AND Cargado=#" & Periodo _
    & "#) AND Periodo=#" & Periodo & "# GROUP BY Codigo, Periodo,Hora,Fecha;"
    Call rtnGenerator(mcDatos, strSQL, "FACT1")
    '
    'Genera las consultas GastoAdmin,GastoMora y gestion
    vecGasto(2) = strGA
    vecGasto(3) = strGM
    vecGasto(4) = strGes
    vecGasto(5) = strNotIn
    '
    For I = 2 To 5
        strSQL = "SELECT CodApto, Periodo, Clng(CCUR(Sum(Monto))*100)/100 AS GNC FROM GastoNoComun WHERE CodG" _
        & "asto " & vecGasto(I) & " AND Periodo=#" & Periodo & "# AND CodApto <> 'U" & gcCodInm _
        & "' GROUP BY CodApto, Periodo UNION SELECT Codigo,'" & Format(Periodo, "mm/dd/yyyy") & _
        "' as P, 0 as G FROM Propietarios WHERE Codigo Not In (SELECT CodApto FROM GastoNoComun" _
        & " WHERE CodGasto " & vecGasto(I) & " AND Periodo=#" & Periodo & "#) AND Codigo <> 'U" _
        & gcCodInm & "';"
        Call rtnGenerator(mcDatos, strSQL, "FACT" & I)
    Next
    '/10 + 0.01)*10
    strSQL = "SELECT Propietarios.Codigo, Propietarios.Nombre, FACT1.Periodo, FACT1.GC, FACT2.G" _
    & "NC, FACT3.GNC, FACT4.GNC, FACT5.GNC, Clng(CCur((FACT1.GC+FACT2.GNC+FACT3.GNC+FACT4.GNC+FACT5." _
    & "GNC))* 100)/100 AS Facturado,Propietarios.Alicuota,Propietarios.ID,FACT1.Hora,FACT1.F" _
    & "echa FROM ((((FACT1 INNER JOIN Propietarios ON FACT1.Codigo = Propietarios.Codigo) INNER" _
    & " JOIN FACT2 ON FACT1.Codigo = FACT2.CodApto) INNER JOIN FACT3 ON FACT1.Codigo = FACT3.Co" _
    & "dApto) INNER JOIN FACT4 ON FACT1.Codigo = FACT4.CodApto) INNER JOIN FACT5 ON FACT1.Codig" _
    & "o= FACT5.CodApto;"
    Call rtnGenerator(mcDatos, strSQL, "FACT6")
    '
    '---------------------------------------------------------------------------------------------
    End Sub


    '---------------------------------------------------------------------------------------------
    '  function Total_Deuda
    '
    '   Devuelve el monto total del deuda del inmeuble
    '---------------------------------------------------------------------------------------------
    Public Function Total_Deuda(strInm$) As Currency
    '
    With FrmAdmin.objRst
        .Requery
        .MoveFirst
        .Find "CodInm='" & strInm & "'"
        If Not .EOF Then Total_Deuda = CCur(!Deuda)
    
    End With
    '
    End Function

    '---------------------------------------------------------------------------------------------
    '   Funcion:    Config_Query
    '
    '   Entrada:    Periodo
    '
    '   Funcion que genera consultas necearias para el proceso, si ocurre un error
    '   durante su ejecución devuelve 'verdadero', de lo contrario devuelve 'falso'
    '---------------------------------------------------------------------------------------------
    Function Config_Query(ByVal Periodo As Date, Path$) As Boolean
    '
    On Error Resume Next
    'variables locales
    Dim str5 As String
    Dim str6 As String
    Dim strSQL As String
    Dim QdfName As String
    '
    str5 = "SELECT CodGasto, Cargado, Descripcion, Comun, Sum(Monto) AS Total FROM AsignaGasto " _
    & "GROUP BY CodGasto, Cargado, Descripcion, Comun HAVING (Comun=True AND Cargado=#" & _
    Periodo & "#);"
    '
    str6 = "SELECT DF.Codigo, DF.Fact, DF.CodGasto, DF.Detalle, DF.CodGasto, Sum(DF.Monto) AS " _
    & "SumaDeMonto, DF.Hora, Propietarios.Nombre, Propietarios.Alicuota, Propietarios.Deuda, " _
    & "Propietarios.Recibos, DF.Periodo, DF.Fecha, AG.Total,iif(Propietarios.AC,1,0) as bAviso FROM (DetFact AS DF INNER JOIN Prop" _
    & "ietarios ON DF.Codigo = Propietarios.Codigo) LEFT JOIN AG ON (DF.CodGasto = AG.CodGasto)" _
    & " AND (DF.Detalle = AG.Descripcion) GROUP BY DF.Codigo,DF.Fact, DF.CodGasto, DF.Detalle, " _
    & "DF.CodGasto, DF.Hora, Propietarios.Nombre, Propietarios.Alicuota, Propietarios.Deuda, " _
    & "Propietarios.Recibos, DF.Periodo, DF.Fecha, AG.Total,Propietarios.AC HAVING (((DF.Periodo)=#" & Periodo & "#));"
    'esta linea va antes de la clausula GROUP BY
    'WHERE (((Propietarios.AC)=False))
    For K = 5 To 6
        strSQL = IIf(K = 5, str5, str6)
        QdfName = IIf(K = 5, "AG", "AC")
        Call rtnGenerator(Path, strSQL, QdfName)
    Next
    If Err.Number <> 0 Then Config_Query = MsgBox(Err.Description, vbInformation)
    '
    End Function

    Public Function Deuda_Total(Mes As Date)
    'variables locles
    Dim rstDeuda As New ADODB.Recordset
    Dim cnn As New ADODB.Connection
    '
    cnn.Open cnnOLEDB & mcDatos
    rstDeuda.Open "SELECT Sum(Facturado) FROM Factura WHERE Periodo=#" & Mes & "# and fact n" _
    & "ot lIKE 'CH%'", cnn, adOpenKeyset, adLockOptimistic
    Deuda_Total = IIf(rstDeuda.BOF Or rstDeuda.EOF, 0, rstDeuda.Fields(0))
    rstDeuda.Close
    Set rstDeuda = Nothing
    cnn.Close
    Set cnn = Nothing
    End Function

    '---------------------------------------------------------------------------------------------
    '
    '   Funcion AjsuteGC
    '---------------------------------------------------------------------------------------------
    Function AjusteGC() As Currency
    'variables locales
    Dim rstAjuste As New ADODB.Recordset
    Dim cnnAjuste As New ADODB.Connection
    '
    cnnAjuste.Open cnnOLEDB + mcDatos
    rstAjuste.Open "SELECT ccur(SUM(GC)) FROM FACT6;", cnnAjuste, adOpenStatic, adLockReadOnly
    AjusteGC = IIf(IsNull(rstAjuste.Fields(0)), 0, rstAjuste.Fields(0))
    '
    Set rstAjuste = Nothing
    Set cnnAjuste = Nothing
    End Function
    
    '---------------------------------------------------------------------------------------------
    '   Funcion:    LetraTitulo
    '
    '   Devuelve un tipo de letra determinado por los parametros recibios
    '
    '---------------------------------------------------------------------------------------------
    Public Function LetraTitulo(Nombre$, Tamaño%, Optional Negrita As Boolean, _
    Optional Subrayado As Boolean) As StdFont
    '
    Set LetraTitulo = New StdFont
    LetraTitulo.Name = Nombre
    LetraTitulo.Size = Tamaño
    LetraTitulo.Underline = Subrayado
    LetraTitulo.Bold = Negrita
    '
    End Function

    
    '---------------------------------------------------------------------------------------------
    '   Function:   booFondo
    '
    '   Funcion que devuelve True si el código asignado al gasto corresponde
    '   con algún fondo de reserva activado para el inmueble, de lo contrario
    '   devuelve False
    '---------------------------------------------------------------------------------------------
    Public Function booFondo(Codigo As String, Optional Inm As String) As Boolean
    'variables locales
    Dim rstFondos As New ADODB.Recordset
    Dim strCnn As String
    '
    If Inm = "" Then
        strCnn = cnnOLEDB + mcDatos
    Else
        strCnn = cnnOLEDB + Inm
    End If
    rstFondos.Open "SELECT * FROM Tgastos WHERE CodGasto='" & Codigo _
    & "' AND Fondo=True;", strCnn, adOpenKeyset, adLockOptimistic
    If Not rstFondos.EOF Or Not rstFondos.BOF Then booFondo = True
    rstFondos.Close
    Set rstFondos = Nothing
    End Function


    '---------------------------------------------------------------------------------------------
    '   Rutina: centra_titulo
    '
    '   centra el encabezado de las celdas de un grid
    '---------------------------------------------------------------------------------------------
    Public Sub centra_titulo(Grid As Control, Optional ancho As Boolean)
    'variables locales
    Dim I%, Ncol() As Integer, N%, strTag$ 'variables locales
    
    '
    With Grid
        '
        .Visible = False
        .FormatString = .FormatString
        .Refresh
        strTag = .Tag
        .Row = 0
        N = 1
        For I = 0 To .Cols - 1
            .Col = I
            .CellAlignment = flexAlignCenterCenter
            If ancho Then
                N = InStr(strTag, "|")
                .ColWidth(I) = Left(strTag, IIf(N = 0, Len(strTag), N - 1))
                If N > 0 Then strTag = Right(strTag, Len(strTag) - N)
            End If
        Next I
        .Col = 0
        .Row = 0
        .Visible = True
        
    End With
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    '   Funcion:    Subject
    '
    '   Devuelve el texto que forma parte del cuerpo del mensaje enviado a los clientes
    '   adjuntandole su aviso de cobro.-
    '---------------------------------------------------------------------------------------------
    Public Function Subjet() As String
    'Variables locales
    Dim numFichero As Integer
    Dim I As Integer
    '
    On Error Resume Next
    numFichero = FreeFile
    '
    Open gcPath & "\email.txt" For Input As numFichero
    I = 1
    Do
        Linea = Input(1, #numFichero)
        Subjet = Subjet + Linea
        I = I + 1
    Loop Until EOF(numFichero)
    Subject = Subjet + vbCrLf
    '
    Close numFichero
    '
    End Function

    '---------------------------------------------------------------------------------------------
    '   Rutina:     Enviar_ACemail
    '
    '   Entradas:    Periodo, (Opcional) Código de Propietario
    '
    '   Envia todos el aviso de cobro correspondiente al periodo determinado
    '   por el argumento 'Periodo'y del propietario señalado en el argumento
    '   'propietario' a su dirección de correo electrónico
    '---------------------------------------------------------------------------------------------
    Public Sub Enviar_ACemail(Periodo As Date, Optional Propietario$, _
    Optional Fact As Boolean, Optional Ver As Boolean)
    'variables locales
    Dim strSQL As String
    Dim numFichero As Integer
    Dim strArchivo As String, IDArchivo As String * 10
    Dim rstEmail(1) As New ADODB.Recordset
    Dim Mes$, apto$, Naviso$, Nombre$, Alic@, MP$, Facturado$
    Dim Total@, Deuda@, Comun@, m&, F&
    Dim Dir1$, Dir2$, PyC&
    Dim mPeriodo As Date
    Dim gsTEMPDIR As String
    Dim lchar As Long
    Dim m_report As CRAXDRT.Report
    Dim m_app As CRAXDRT.Application
    '----------------------------------
    mPeriodo = Format(Periodo, "mm/dd/yy")
    
    'selecciona todos los propietarios del inmueble que tengan e-mail
    If Not Ver Then
        strSQL = "SELECT * FROM Propietarios WHERE email <>'' AND Demanda = False;"
    Else
        strSQL = "SELECT * FROM Propietarios WHERE Codigo Not Like 'U%'"
    End If
    
    With rstEmail(0)
    '
        .Open strSQL, cnnOLEDB + mcDatos, adOpenStatic, adLockReadOnly, adCmdText
        'Si pasa algún propietario filtra por ese valor
        If Propietario <> "" Then .Filter = "Codigo='" & Propietario & "'"
        '
        If Not .EOF Or Not .BOF Then
            .MoveFirst
            If mPeriodo <= CDate("01/07/2005") Then
                'Genera una consulta de los totales de gastos comunes por gasto en el período señalado
                strSQL = "SELECT CodGasto, Cargado, Descripcion, Comun, Sum(Monto) AS Total FROM " _
                & "AsignaGasto GROUP BY CodGasto, Cargado, Descripcion, Comun HAVING (Comun=True " _
                & "AND Cargado=#" & Periodo & "#);"
                Call rtnGenerator(mcDatos, strSQL, "AG")
                'Busca el monto del fondo de reserva
                F = Total_Fondo(gcCodInm, Periodo)
            Else
                gsTEMPDIR = String$(255, 0)
                lchar = GetTempPath(255, gsTEMPDIR)
                gsTEMPDIR = Left(gsTEMPDIR, lchar)
                Set m_app = New CRAXDRT.Application
                Set m_report = New CRAXDRT.Report
                
            End If
            '
            Do
                IDArchivo = Format(mPeriodo, "mmyy") + Right(gcCodInm, 3) + Format(!ID, "000")
                Mes = UCase(Format(mPeriodo, "mmm-yyyy"))
                If mPeriodo <= CDate("01/07/2005") Then
                    'Selecciona el detalle de la factura señalada en el argumento 'periodo'
                    strSQL = "SELECT DF.Codigo, DF.Fact, DF.Detalle, DF.CodGasto, Sum(DF.Monto) AS SubT" _
                    & "otal, DF.Hora, DF.Periodo, DF.Fecha, AG.Total, Factura.Facturado FROM (DetFact A" _
                    & "S DF LEFT JOIN AG ON (DF.Detalle = AG.Descripcion) AND (DF.CodGasto = AG.CodGast" _
                    & "o)) INNER JOIN Factura ON DF.Fact = Factura.FACT GROUP BY DF.Codigo, DF.Fact, DF" _
                    & ".Detalle, DF.CodGasto, DF.Hora, DF.Periodo, DF.Fecha, AG.Total, DF.CodGasto, Fac" _
                    & "tura.Facturado HAVING (((DF.Codigo)='" & !Codigo & "') AND ((DF.Periodo)=#" _
                    & Periodo & "#)) ORDER BY DF.CodGasto;"
                    '
                    With rstEmail(1)
                    '
                        .Open strSQL, cnnOLEDB + mcDatos, adOpenStatic, adLockReadOnly, adCmdText
                        '
                        If Not .EOF Or Not .BOF Then
                            .MoveFirst
                            Nombre = IIf(IsNull(rstEmail(0)!Nombre), "", rstEmail(0)!Nombre)
                            'Mes = UCase(Format(.Fields("Periodo"), "mmm-yyyy"))
                            Naviso = .Fields("Fact")
                            
                            If Dir(gcPath & gcUbica & "\reportes\" & .Fields("Fact") & ".html") <> "" Then
                                .Close
                                If Ver Then    'si no esta facturando
                                    Exit Sub
                                Else
                                    GoTo 10
                                End If
                            End If
                            'asigna valores a las variables locales
                            
                            Alic = rstEmail(0)!Alicuota
                            MP = rstEmail(0)!Recibos
                            apto = rstEmail(0)!Codigo
                            Deuda = IIf(rstEmail(0)!Recibos > gnMesesMora, rstEmail(0)!Deuda + _
                            (rstEmail(0)!Deuda * gnPorIntMora / 100), rstEmail(0)!Deuda)
                            Comun = 0
                            Facturado = .Fields("Fecha")
                            Total = .Fields("Facturado")
                            
                            'genera el archivo formato html (encabezado)
                            numFichero = FreeFile
                            strArchivo = gcPath & gcUbica & "reportes\" & Naviso & ".html"
                            Open strArchivo For Output As numFichero
                                Print #numFichero, Encabezado(Mes, gcCodInm, gcNomInm, apto, Naviso, Nombre, Alic, _
                                Facturado, Format(Total, "#,##0.00 "), Format(Deuda, "#,##0.00 "), MP, Format(CCur(F), "#,##0.00 "))
                            Close numFichero
                            '
                            Do  'genera el detalle de la notificación de gastos
                                Open strArchivo For Append As numFichero
                                    If IsNull(!Total) Then
                                        m = 0
                                    Else
                                        m = CLng(!Total)
                                    End If
                                    Print #numFichero, detalle(!codGasto, !detalle, m, !subTotal)
                                Close numFichero
                                Comun = Comun + m
                                .MoveNext
                            Loop Until .EOF
                            '----------------------------------
                            'Cierra el archivo html
                            Open strArchivo For Append As numFichero
                                Print #numFichero, cierre(Total, Comun)
                            Close numFichero
                            '
                        End If
                        .Close
                    End With
                Else
                    'envia el archivo en formato pdf
                    Set m_report = m_app.OpenReport(gcPath & gcUbica & "reportes\AC" & _
                    UCase(Format(mPeriodo, "MMMYY")) & ".rpt", 1)
                    m_report.RecordSelectionFormula = "{AC.Codigo}='" & rstEmail(0)!Codigo & "'"
                    m_report.DisplayProgressDialog = False
                    m_report.ExportOptions.DestinationType = crEDTDiskFile
                    m_report.ExportOptions.FormatType = crEFTPortableDocFormat
                    strArchivo = gsTEMPDIR & IDArchivo & ".pdf"
                    m_report.ExportOptions.DiskFileName = strArchivo
                    m_report.Export (False)
                End If
        'Envia el archivo generado vía e-amil
10         If Not Ver And !email <> "" Then
                '
                If frmSelecInm.MAPm.SessionID = 0 Then
                    frmSelecInm.MAPs.SignOn
                    frmSelecInm.MAPm.SessionID = frmSelecInm.MAPs.SessionID
                End If
                '
                If InStr(!email, ";") = 0 Then
                    Dir1 = !email
                    Dir2 = ""
                Else
                    PyC = InStr(!email, ";")
                    Dir1 = Left(!email, PyC - 1)
                    Dir2 = Trim(Mid(!email, PyC + 1, 200))
                End If
                frmSelecInm.MAPm.Compose
                frmSelecInm.MAPm.RecipIndex = 0
                frmSelecInm.MAPm.RecipAddress = Dir1
                frmSelecInm.MAPm.RecipDisplayName = Nombre
                frmSelecInm.MAPm.RecipType = mapToList
                
                '
                If Dir2 <> "" Then
                    frmSelecInm.MAPm.RecipIndex = 1
                    frmSelecInm.MAPm.RecipAddress = Dir2
                    frmSelecInm.MAPm.RecipDisplayName = Nombre
                    frmSelecInm.MAPm.RecipType = mapCcList
                End If
                '
                frmSelecInm.MAPm.MsgNoteText = Subjet
                frmSelecInm.MAPm.MsgSubject = "Aviso de Cobro Período: " & Mes
                frmSelecInm.MAPm.AttachmentPosition = Len(frmSelecInm.MAPm.MsgNoteText) - 1
                frmSelecInm.MAPm.AttachmentName = UCase(Right(strArchivo, Len(strArchivo) - InStrRev(strArchivo, "\")))
                frmSelecInm.MAPm.AttachmentPathName = strArchivo 'gcPath & gcUbica & "\reportes\" _
                & Naviso & ".html"
                frmSelecInm.MAPm.Send False
                If mPeriodo > CDate("01/07/2005") Then Kill strArchivo
                nEnviados = nEnviados + 1
                'frmselecinm.MAPs.SignOff
                '
            End If
            .MoveNext
            Loop Until .EOF
            If Not Ver Then
                If frmSelecInm.MAPm.SessionID <> 0 Then
                    frmSelecInm.MAPs.SignOff
                    frmSelecInm.MAPm.SessionID = 0
                End If
            End If
        End If
        .Close
    End With
    If mPeriodo > CDate("01/07/2005") Then
        Set m_report = Nothing
        Set m_app = Nothing
    End If
    '
    Set rstEmail(0) = Nothing
    Set rstEmail(1) = Nothing
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    '
    '   Función:    Encabezado
    '
    '   Entradas:   Apto,Nombre_Propietario,N_Aviso,Alicuota,Fecha_Impresion
    '               Neto_Pagar, Deuda_Acumulada,MP
    '
    '   Devuelve una cadena con el encabezado del archio htm.
    '---------------------------------------------------------------------------------------------
    Private Function Encabezado(ParamArray data()) As String
    'apto$, Nombre_Propietario$, N_Aviso$, Alicuota@, _
    Fecha_Impresion$, Neto_Pagar@, Deuda_Acumulada@, MP$, Periodo$, FondoR@
    'variables locales
    Dim N As Long, Contador As Long, Dato As Byte
    '------------------------------------'-------
    '
    '<script language="javascript">
    'window.open('http://i.domaindlx.com/ynfantes/leido.asp?Periodo=01/05/2005&CodInm=2501&CodPro=011A&Fecha=29/06/2005&Clave=0102032501','','width=300,height=120,Top=250,Left=300');
    '</script>
    '
    '------------------------------------'-------
    N = FreeFile
    Open gcPath & "\encabezado.txt" For Binary As #N
        Do While Not EOF(N)
            Get N, , Dato
            If Dato <> 13 And Dato <> 10 Then
                If Dato = 63 Then
                    Encabezado = Encabezado & data(Contador)
                    Contador = Contador + 1
                Else
                    Encabezado = Encabezado & Chr(Dato)
                End If
            End If
        Loop
        
    Close #N
    Contador = 0
    '
    'Encabezado = "<HTML>" & vbCrLf & "<head>" & vbCrLf & Space(3) & "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>" & vbCrLf & Space(3) _
    & "<title>Aviso de cobro</title>" & vbCrLf & "<style type='text/css'><!--body,td,th { font-family: Geneva, Arial, Helvetica, sans-serif;} --></style></head>" & vbCrLf & "<body text='#000000' BGCOLOR='#FFFFFF'>&nbsp;" & vbCrLf & "<table border=1 WIDTH=80% align='center' bordercolor='#FFFFFF' cellpadding=0 " _
    & "cellspacing=1 style='font-family:Verdana, Arial, Helvetica, sans-serif; font-size:xx-small'>" & vbCrLf & Space(3) & "<tr>" & vbCrLf & Space(3) & "<td COLSPAN=3 Align='Center'><Font Size=+3 Color='#008080'><b>S A C</b></font></td>" & vbCrLf & Space(3) & "<td align=CENTER COLSPAN=3 BGCOLOR='#008080' bordercolor='#FFFFFF'><b><font size='+1' color='#FFFFFF'>NOTIFICACI&Oacute;N DE GASTOS</font>" _
    & "</b></td>" & vbCrLf & "</tr>" & vbCrLf & "<tr>" & vbCrLf & Space(3) & "<td Align=Center COLSPAN=3><Font face='ARIAL NARROW' SIZE=1 Color=#008080><u>SERVICIO DE ADMINISTRACION DE CONDOMINIO</u></font></td>" & vbCrLf & Space(3) & "<td align=CENTER COLSPAN='3' BGCOLOR='#008080' bordercolor='#FFFFFF'><b><font size='-1' color='#FFFFFF'>PERIODO: " & vbCrLf & Space(3) _
    & Periodo & "</font></b></td>" & vbCrLf & Space(3) & "</tr>" & vbCrLf & "<TR>" & vbCrLf & Space(3) & "<TD COLSPAN=6 HEIGHT=5></TD>" & vbCrLf & "</TR>" & vbCrLf & "<tr bordercolor='#000000'>" & vbCrLf & Space(3) & "<td align=CENTER WIDTH='10%' BGCOLOR='#008080'><b><font face='Arial Narrow' " _
    & "color='#FFFFFF' size='-1'>C&oacute;digo</font></b></td>" & vbCrLf & Space(3) & "<td align=CENTER WIDTH='60%' COLSPAN='3' BGCOLOR='#008080'><b><font face='Arial Narrow' color='#FFFFFF' size='-1'>Inmueble</font></b></td>" & vbCrLf & Space(3) & "<td align=CENTER WIDTH='10%' BGCOLOR='#008080'><b><font face='Arial Narrow' " _
    & "color='#FFFFFF' size='-1'>Apto.</font></b></td>" & vbCrLf & Space(3) & "<td align=CENTER WIDTH='20%' BGCOLOR='#008080'><b><font face='Arial Narrow' color='#FFFFFF' size=-1>N&ordm;Aviso</font></b></td>" & vbcrl & Space(3) & "</tr>" & vbCrLf & "<tr bordercolor='#000000'>" & vbCrLf & Space(3) & "<td align=CENTER " _
    & "WIDTH='10%' BGCOLOR='#F7EFDE'>" & gcCodInm & "</td>" & vbCrLf & Space(3) & "<td align=CENTER WIDTH='60%' COLSPAN='3' BGCOLOR='#F7EFDE'>" & gcNomInm & "</td>" & vbCrLf & Space(3) & "<td align=CENTER WIDTH='10%' BGCOLOR='#F7EFDE'>" & apto & "</td>" & vbCrLf & Space(3) & "<td align=CENTER " _
    & "WIDTH='20%' BGCOLOR='#F7EFDE'>" & N_Aviso & "</td>" & vbCrLf & "</tr>" & vbCrLf & "<tr>" & vbCrLf & Space(3) & "<td valign=TOP COLSPAN='6' height='5'></td>" & vbCrLf & "</tr>" & vbCrLf & "<tr bordercolor='#000000'>" & vbCrLf & Space(3) & "<td align=CENTER WIDTH='60%' COLSPAN=4 BGCOLOR='#008080'><b><font face='Arial Narrow' color='#FFFFFF' " _
    & "size='-1'>Propietario</font></b></td>" & vbCrLf & Space(3) & "<td align=CENTER WIDTH=10% BGCOLOR='#008080'><b><font face='Arial Narrow' color='#FFFFFF' size='-1'>Al&iacute;cuota</font></b></td>" & vbCrLf & Space(3) _
    & "<td align=CENTER WIDTH='30%' BGCOLOR='#008080'><b><font face='Arial Narrow' color='#FFFFFF' size='-1'>Fecha Impresi&oacute;n</font></b></td>" & vbCrLf & "</tr>" & vbCrLf & "<tr bordercolor='#000000'>" & vbCrLf & Space(3) _
    & "<td align=CENTER WIDTH='60%' COLSPAN=4 BGCOLOR='#F7EFDE'>" & Nombre_Propietario & "</td>" & vbCrLf & Space(3) & "<td align=CENTER WIDTH='10%' BGCOLOR='#F7EFDE'>" & Format(Alicuota, "#,##0.000000 ") & "</td>" & vbCrLf & Space(3) & "<td align=CENTER WIDTH='30%' " _
    & "BGCOLOR='#F7EFDE'>" & Fecha_Impresion & "</td>" & vbCrLf & "</tr>" & vbCrLf & "<tr>" & vbCrLf & Space(3) & "<td COLSPAN='6' height='5' WIDTH='100%'></td>" & vbCrLf & Space(3) & "</tr>" & vbCrLf & "<tr bordercolor='#000000'>" & vbCrLf & Space(3) _
    & "<td align=CENTER WIDTH='35%' COLSPAN='3' BGCOLOR='#008080'><b><font face='Arial Narrow' color='#FFFFFF' size='-1'>Neto a pagar</font></b></td>" & vbCrLf & Space(3) _
    & "<td align=CENTER WIDTH='35%' BGCOLOR='#008080'><b><font face='Arial Narrow' color='#FFFFFF' size='-1'>Deuda Acumulada</font></b></td>" & vbCrLf & Space(3) & "<td align=CENTER WIDTH='10%' BGCOLOR='#008080'><b><font face='Arial Narrow' " _
    & "color='#FFFFFF' size='-1'>M.P.</font></b></td>" & vbCrLf & Space(3) & "<td align=CENTER WIDTH='20%' BGCOLOR='#008080'><b><font face='Arial Narrow' color='#FFFFFF' size='-1'>Fondo Reserva</font></b></td>" _
    & vbCrLf & "</tr>" & vbCrLf & "<tr bordercolor='#000000'>" & vbCrLf & Space(3) & "<td align=CENTER WIDTH='35%' COLSPAN='3' BGCOLOR='#F7EFDE'>" & Format(Neto_Pagar, Monto) & "</td>" & vbCrLf & Space(3) _
    & "<td align=CENTER WIDTH='35%' BGCOLOR='#F7EFDE'>" & Format(Deuda_Acumulada, Monto) & "</td>" & vbCrLf & Space(3) & "<td align=CENTER WIDTH='10%' BGCOLOR='#F7EFDE'>" & MP & "</td>" & vbCrLf & Space(3) _
    & "<td align=CENTER WIDTH='20%' BGCOLOR='#F7EFDE'>" & Format(FondoR, "#,##0.00 ") & "</td>" & vbCrLf & "</tr>" & vbCrLf & "</table>" & vbCrLf & "<BR>" & vbCrLf & "<table WIDTH='80%' align='center' BORDER='1' bordercolor='#000000' cellspacing='0' cellpadding='0' style='font-family:Verdana, Arial, Helvetica, sans-serif; font-size:xx-small'>" & vbCrLf & Space(3) & "<tr>" _
    & vbCrLf & "<td ALIGN=CENTER WIDTH='10%' BGCOLOR='#008080' bordercolor='#008080'><b><font face='Arial Narrow' color='#FFFFFF' size='-1'>Cod.Gasto</font></b></td>" & vbCrLf & Space(3) & "<td ALIGN=CENTER WIDTH='66%' BGCOLOR='#008080' " _
    & "bordercolor='#008080'><b><font face='Arial Narrow' color='#FFFFFF' size='-1'>Descripci&oacute;n</font></b></td>" & vbCrLf & Space(3) & "<td ALIGN=CENTER WIDTH='12%' BGCOLOR='#008080' bordercolor='#008080'><b><font face='Arial Narrow' " _
    & "color='#FFFFFF' size='-1'>Monto</font></b></td>" & vbCrLf & Space(3) & "<td ALIGN=CENTER WIDTH='12%' BGCOLOR='#008080' bordercolor='#008080'><b><font face='Arial Narrow' " _
    & "color='#FFFFFF' size='-1'>X Alícuota</font></b></td>" & vbCrLf & "</tr>"
    '
    End Function


    '---------------------------------------------------------------------------------------------
    '
    '   Función:    Detalle
    '
    '   Entrada:    CodGasto,Descripción,Comun,Monto
    '
    '   Devuelve una cadena que genera una nueva linea en la tabla dentro del
    '   archivo *.html
    '---------------------------------------------------------------------------------------------
    Private Function detalle(codGasto$, Descripcion$, Comun&, Monto@) As String
    '
    detalle = "<tr class='sacDetalle'><td ALIGN=CENTER>" & codGasto & "</td>" _
   & "<td>" & Descripcion & "</td><td ALIGN=RIGHT >" & IIf(Comun = 0, "", _
   Format(Comun, "#,##0.00")) & "</td><td ALIGN=RIGHT >" & Format(Monto, "#,##0.00") _
   & "</td></tr>"
    'Detalle = "<tr>" & vbCrLf & "<td ALIGN=CENTER BGCOLOR='#F7EFDE' bordercolor='#FFFFFF'>" & CodGasto & "</td>" _
    & vbCrLf & Space(3) & "<td BGCOLOR='#F7EFDE' bordercolor='#FFFFFF'>" & Descripcion & "</td>" & _
    "<td ALIGN=RIGHT BGCOLOR='#F7EFDE' bordercolor='#FFFFFF'>" & IIf(Comun = 0, "", Format(Comun, "#,##0.00 ")) & "</td>" & vbCrLf _
    & "<td ALIGN=RIGHT BGCOLOR='#F7EFDE' bordercolor='#FFFFFF'>" & Format(Monto, "#,##0.00 ") & "</td>" _
    & "</tr>" & vbCrLf
    '
    End Function

    '---------------------------------------------------------------------------------------------
    '
    '   Función:    Cierre
    '
    '   Entrada:    Neto (Total a Pagar), Gasto_Comun(Monto de c/gasto común)
    '
    '   Devuelve una cadena que contiene el finald el archivo *.html
    '---------------------------------------------------------------------------------------------
    Private Function cierre(Neto@, Gasto_Comun@) As String
    'variables locales
    Dim strBenef$   '
    '
    strBenef = gcNomInm
    strCB = complemento
    '
    If gnCta = CUENTA_POTE Then strBenef = ""
'    Else
'        strBenef = "EMITIR CHEQUE NO ENDOSABLE A NOMBRE DE " & strBenef
'    End If
    cierre = "<tr valign='bottom' class='sacTitulo'><td COLSPAN='2' ALIGN=RIGHT>Totales&nbsp;" _
    & "</td><td ALIGN=RIGHT>" & Format(Gasto_Comun, "#,##0.00") & "</td><td ALIGN=RIGH" _
    & "T>" & Format(Neto, "#,##0.00") & "</tr></table><DIV  align='center' style='margin-top:7px ;" _
    & "font: bold 14px Verdana, Arial, Helvetica, sans-serif; color:#FF0000;  text-decoration: blink;'>" _
    & "CANCELE ANTES DEL 20 DE CADA MES Y EVITESE RECARG" _
    & "O DE INTERESES</DIV><div align='center' style='margin-top:7px; font:10px Arial, Helvetica, sans-serif'" _
    & ">" & IIf(strBenef = "", "", "FAVOR EMITIR CHEQUE NO ENDOSABLE A NOMBRE DE: " _
                & strBenef) & IIf(strCB = "" Or strNombre = "", "", " 0 ") & strCB & "" _
    & "</div><hr heiht=1><div style='font:10px Arial, Helvetica, sans-serif; color:#0099FF'>2004 - " _
    & "<a href='http:\\www.administradorasac.com.ve' target='_blank'>Servicio de Administracion de Condominios</a>" _
    & ". Todos los derechos Reservados</div><" _
    & "div style='font:bold 14px Times New Roman, Times, serif'>Este informe se ha creado util" _
    & "izando</div><marquee behavior='slide' direction='LEFT' width='500' loop='3'><Font Co" _
    & "lor='#0000CC'><B>SAC&reg;</B><i>(<b>S</b><font color='#0000CC'></font></i><i>" _
    & "istema de <b>A</b>dministraci&oacute;n de <b>C</b>ondominios)</i></font></ma" _
    & "rquee><div style='font:bold 14px Times New Roman, Times, serif'>una marca comercial" _
    & " de <B><a href='mailto:ynfantes@cantv.net'>SinaiTech, C.A</a></B></div></body></html>"
    
    'cierre = Space(3) & "<tr>" & vbCrLf & Space(3) & "<td ALIGN=RIGHT COLSPAN='2' BGCOLOR='#008080' bordercolor='#008080'><b><font face='Arial Narrow' color='#FFFFFF' " _
    & "size='-1'>Totales&nbsp;</font></b></td>" & vbCrLf & Space(3) & "<td ALIGN=RIGHT BGCOLOR='#008080' bordercolor='#008080'><b><font face='Arial Narrow' " _
    & "color='#FFFFFF' size='-1'>" & Format(Gasto_Comun, "#,##0.00 ") & "</font></b></td>" & vbCrLf & Space(3) & "<td ALIGN=RIGHT BGCOLOR='#008080' bordercolor='#008080'><b><font face='Arial Narrow' " _
    & "color='#FFFFFF' size='-1'>" & Format(Neto, "#,##0.00 ") & "</font></b></td>" & vbCrLf & "</tr>" & vbCrLf & "</table>" & vbCrLf & "<DIV ALIGN=CENTER><b><FONT face='Geneva, Arial, Helvetica, sans-serif' SIZE='-1'>CANCELE ANTES DEL 20 " _
    & "DE CADA MES Y EVITESE RECARGO DE INTERESES</FONT></b></DIV>" & vbCrLf & "<TABLE WIDTH=80% ALIGN=CENTER BORDER=1 CELLPADDING=0 CELLSPACING=0>" & vbCrLf & Space(3) & "<TR>" & vbCrLf & Space(6) & _
    "<TD ALIGN=Center BGCOLOR=#FFFFFF BORDERCOLOR=#008080 WIDTH=80%>" & vbCrLf & vbTab & "<FONT face='Arial, Helvetica, sans-serif' SIZE=-2>EMITIR CHEQUE NO ENDOSABLE A NOMBRE DE " & strBenef & "</FONT>" & vbCrLf & Space(6) & "</TD>" & vbCrLf & Space(3) & _
    "</TR</TABLE><TABLE WIDTH=80% ALIGN=CENTER><TR><TD ALIGN=CENTER WIDTH=80% BORDERCOLOR=#008080><FONT face='Geneva, Arial, Helvetica, sans-serif' SIZE=-2 ><B>" & UCase(Obtener_Nota(gcUbica)) & "</B></FONT>" & vbCrLf & "</TD>" & "</TR>" & vbCrLf & "</TR></TABLE>" & vbCrLf & "<br>" & vbCrLf & "<hr heiht=1>Este informe " _
    & "se ha creado utilizando<br><marquee behavior='slide' direction='LEFT' width='500' loop='3'><Font Color='#0000CC'><B>SAC&reg;</B><i>(<b>S</b><font color='#0000CC'></font></i><i>istema de <b>A</b>dministraci&oacute;n" _
    & " de <b>C</b>ondominios)</i></font></marquee><br>una marca comercial de <B><a href='mailto:ynfantes@cantv.net'>SinaiTech, C.A</a></B>" & vbCrLf & "</body>" _
    & vbCrLf & "</html>"
    '
    End Function

    
    Public Function strPSWD(datos$) As String
    'variables locales
    Dim Mascara$
    '
    Mascara = Chr(78) & Chr(134) & Chr(251) & Chr(236) & Chr(55) & Chr(93) & Chr(68) & _
    Chr(156) & Chr(250) & Chr(198) & Chr(94) & Chr(40) & Chr(230) & Chr(19)
    Open datos For Binary As #1
    Seek #1, &H42
    '
    For N = 1 To 14
        s1 = Mid(Mascara, N, 1)
        s2 = Input(1, 1)
        If (Asc(s1) Xor Asc(s2)) <> 0 Then
                strPSWD = strPSWD & Chr(Asc(s1) Xor Asc(s2))
        End If
    Next
    Close 1
    '
    End Function

    '---------------------------------------------------------------------------------------------
    '   Función:    Detalle_fondo
    '
    '   Entrada:    Código del inmueble
    '
    '   Devuelve Verdadero si la opción esta activa, de lo contrario retorna false
    '---------------------------------------------------------------------------------------------
    Private Function Detalle_Fondo(strInm$) As Boolean
    'variables locales
    Dim rstDF As New ADODB.Recordset
    '
    Set rstDF = cnnConexion.Execute("SELECT * FROM Inmueble WHERE CodInm='" & strInm & "'")
    Detalle_Fondo = rstDF.Fields("FactFV")
    rstDF.Close
    Set rstDF = Nothing
    '
    End Function

    '---------------------------------------------------------------------------------------------
    '
    '   Función:     Cierre_nomina
    '
    '   Efectua el cierre de la nomina, devuelve True si ocurre un error durante el proceso
    '   de lo contrario retorna False. El proceso se efectua como una transacciòn, si no
    
    '   ocurre ningún error se efectuan los cambios, sino se cancelan las modificaciones
    '---------------------------------------------------------------------------------------------
    Public Function cierre_nomina(ID&, subT$, Acredita$, pDF@) As Boolean
    'variables locales
    Dim rstNom As New ADODB.Recordset
    Dim rstCG As New ADODB.Recordset
    Dim rstTemp As New ADODB.Recordset
    Dim Rep(6) As String, strD As String
    Dim strSQL As String, Inm As String
    Dim CargadoA As Date
    Dim errLocal As Long
    Dim Resp As Long
    Dim numToLetras As New clsNum2Let
    Dim RepNom As String
    Dim codNom(2) As String
    Dim crReporte As ctlReport
    Dim errMsg As String
    
    '---------------------------------------------------------------------------------------------
    'varifica los valores de estado
    'On Error Resume Next
    With rstNom
        .Open "Nom_Calc", cnnConexion, adOpenKeyset, adLockOptimistic, adCmdTable
        If Not (.EOF And .BOF) Then
            If !CodDia_Feriado = 0 Or !codDia_Libre = 0 Or !codOtras_Asig = 0 Then
                cierre_nomina = MsgBox("Verifique los parámetros de la nómina. Consulte al administrador del sistema.", _
                vbInformation, App.ProductName)
                .Close
                Exit Function
            Else
                codNom(0) = !codDia_Libre
                codNom(1) = !CodDia_Feriado
                codNom(2) = !codOtras_Asig
            End If
        Else
            cierre_nomina = MsgBox("Verifique los parámetros de la nómina. Consulte al administrador del sistema.", _
            vbInformation, App.ProductName)
            .Close
            Exit Function
        End If
        .Close
        
        'verifica los ncuetas de los inmuebles
        strSQL = "SELECT Inmueble.CodInm, Inmueble.CtaInm  " & _
        "FROM Inmueble INNER JOIN (Emp_Cargos INNER JOIN (Emp INNER JOIN Nom_Detalle ON Emp.CodEmp = Nom_Detalle.CodEmp) ON Emp_Cargos.CodCargo = Emp.CodCargo) ON Inmueble.CodInm = Emp.CodInm " & _
        "Where (((Inmueble.Inactivo) = False) And ((Nom_Detalle.IDNom) = " & ID & ")) ORDER BY Inmueble.Caja, Inmueble.CodInm;"

        .Open strSQL, cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
        If Not (.EOF And .BOF) Then
            .MoveFirst
            
            Do
                If IsNull(.Fields("CtaInm")) Then
                    errMsg = errMsg + "- La cuenta del inmueble " & .Fields("CodInm") _
                    & " no está establecida." & vbCrLf
                    I = I + 1
                End If
                .MoveNext
            Loop Until .EOF
            If I > 0 Then
                cierre_nomina = MsgBox("No se puede procesar esta nómina: " & vbCrLf & vbCrLf & errMsg, _
                vbInformation, App.ProductName)
                .Close
                Exit Function
            End If
        End If
        .Close
        
    End With
    
    RepNom = gcPath & "\nomina\"
    Rem On Error GoTo Cerrar:   'al ocurrir algún error se va al final de la función
    '
    'configura la presentación del config
    '
    Call RtnConfigUtility(True, subT, "Iniando proceso....", "Espere un momento por favor....")
    
    Call rtnBitacora("Iniciando cierre nómina " & IDNomina)
    'en esta función se utiliza la conexión principal
    'abre una transacción.-
    cnnConexion.BeginTrans
    '
    'agrega el registro en la tabla Nom_Inf (Nominas procesadas)
    cnnConexion.Execute "INSERT INTO Nom_Inf(IDNomina,Fecha,Usuario,Efectivo) VALUES (" & ID & _
    ",Date()," & "'" & gcUsuario & "','" & Acredita & "');"
    '
    'agrega la información correspondiente al facturado por cada inmueble
    Call rtnBitacora("Cargando información de facturación")
    
    strSQL = "SELECT Emp.CodGasto, Emp.CodInm, Sum((Nom_Detalle.Sueldo/30*Nom_Detalle." _
    & "Dias_Trab)+((Nom_Detalle.Sueldo/30)*Nom_Detalle.Dias_NoTrab*-1)+(Nom_Detalle." _
    & "Bono_Noc)) AS TotalA, Sum((Nom_Detalle.Sueldo+(Emp.BonoNoc*Nom_detalle.Sueldo/100)" _
    & ")/30*Nom_Detalle.Dias_libres) AS DiasL, Sum((Nom_Detalle.Sueldo+(Emp.BonoNoc*" _
    & "Nom_detalle.Sueldo/100))/30*Nom_Detalle.Dias_Feriados*'" & pDF & "') AS diasF, Sum(Nom_Detalle" _
    & ".Otras_Asignaciones) AS OA, Inmueble.Ubica FROM Inmueble INNER JOIN (Nom_Detalle " _
    & "INNER JOIN Emp ON Nom_Detalle.CodEmp = Emp.CodEmp) ON Inmueble.CodInm = " _
    & "Emp.CodInm Where (((Nom_Detalle.IDNom) = " & ID & ")) GROUP BY Emp.CodGasto, " _
    & "Emp.CodInm, Inmueble.Ubica ORDER BY Emp.CodInm, Emp.CodGasto;"
    '
    'selecciona el conjunto de registros de la nómina
    rstNom.Open strSQL, cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
    '
    CargadoA = "01/" & Mid(ID, 2, 2) & "/" & Right(ID, 4)
    '
    If Not rstNom.EOF And Not rstNom.BOF Then   'si el ADODB.Recordset tiene registros
    '
        rstNom.MoveFirst
        Do
            
            Call RtnProUtility("Cargando información facturación inm. '" & rstNom("CodInm") & _
            "'", rstNom.AbsolutePosition * 6025 / rstNom.RecordCount)
            
            'selecciona el período a facturar
            rstCG.Open "SELECT max(Periodo) FROM Factura WHERE Fact Not Like 'CH%'", cnnOLEDB + _
            gcPath + rstNom("Ubica") + "inm.mdb", adOpenKeyset, adLockOptimistic, adCmdText
            If Not rstCG.EOF And Not rstCG.BOF Then
                If CargadoA = rstCG.Fields(0) Then
                    MsgBox "Este período ya fue facturado", vbCritical, App.ProductName
                    Err.Raise -214728236, , "Imposible cerrar nómina"
                End If
            End If
            rstCG.Close
            '
            'inserta el registro en la tabla facturación
            'cnnConexion.Execute "DELETE * FROM AsignaGasto IN '" & gcPath & rstNom("Ubica") & "inm.mdb' WHERE Ndoc='NOMINA' and Cargado=#" & Format(CargadoA, "mm/dd/yyyy") & "#"
            strSQL = "INSERT INTO AsignaGasto (Ndoc,CodGasto,Cargado,Descripcion,Fijo,Comun,Ali" _
            & "cuota,Monto,Usuario,Fecha,Hora) In '" & gcPath & rstNom("Ubica") & "inm.mdb' SEL" _
            & "ECT 'NOMINA','" & rstNom("CodGasto") & "','" & CargadoA & "',Titulo,False,True,T" _
            & "rue,'" & rstNom("TotalA") & "','" & gcUsuario & "',Date(),Time() FROM TGastos In" _
            & " '" & gcPath & rstNom("Ubica") & "Inm.mdb' WHERE CodGasto='" & rstNom("CodGasto") _
            & "'"
            '
            cnnConexion.Execute strSQL  'sueldo
            '
            If rstNom("DiasL") > 0 Then
            
                strSQL = "INSERT INTO AsignaGasto (Ndoc,CodGasto,Cargado,Descripcion,Fijo,Comun,Ali" _
                & "cuota,Monto,Usuario,Fecha,Hora) In '" & gcPath & rstNom("Ubica") & "inm.mdb' SEL" _
                & "ECT 'NOMINA','" & codNom(0) & "','" & CargadoA & "',Titulo,False,True,T" _
                & "rue,'" & rstNom("DiasL") & "','" & gcUsuario & "',Date(),Time() FROM TGastos In" _
                & " '" & gcPath & rstNom("Ubica") & "Inm.mdb' WHERE CodGasto='" & codNom(0) & "'"
                '
                
                cnnConexion.Execute strSQL  'dias libres
                
            End If
            '
            If rstNom("DiasF") > 0 Then
                strSQL = "INSERT INTO AsignaGasto (Ndoc,CodGasto,Cargado,Descripcion,Fijo,Comun,Ali" _
                & "cuota,Monto,Usuario,Fecha,Hora) In '" & gcPath & rstNom("Ubica") & "inm.mdb' SEL" _
                & "ECT 'NOMINA','" & codNom(1) & "','" & CargadoA & "',Titulo,False,True,T" _
                & "rue,'" & rstNom("DiasF") & "','" & gcUsuario & "',Date(),Time() FROM TGastos In" _
                & " '" & gcPath & rstNom("Ubica") & "Inm.mdb' WHERE CodGasto='" & codNom(1) & "'"
                '
                cnnConexion.Execute strSQL  'dias feriados
            End If
            '
            If rstNom("OA") > 0 Then
            
                strSQL = "INSERT INTO AsignaGasto (Ndoc,CodGasto,Cargado,Descripcion,Fijo,Comun,Ali" _
                & "cuota,Monto,Usuario,Fecha,Hora) In '" & gcPath & rstNom("Ubica") & "inm.mdb' SEL" _
                & "ECT 'NOMINA','" & codNom(2) & "','" & CargadoA & "',Titulo,False,True,T" _
                & "rue,'" & rstNom("OA") & "','" & gcUsuario & "',Date(),Time() FROM TGastos In" _
                & " '" & gcPath & rstNom("Ubica") & "Inm.mdb' WHERE CodGasto='" & codNom(2) & "'"
                '
                cnnConexion.Execute strSQL  'otras asignaciones
                
            End If
                
            rstNom.MoveNext
            '
        Loop Until rstNom.EOF
        '
    End If
    '
    rstNom.Close
    'carga el reintegro de seguro social
    strSQL = "SELECT Emp.CodInm, Sum(Nom_Detalle.SSO + Nom_Detalle.spf) AS sumSSO, Inmueble.Ubi" _
    & "ca FROM Inmueble INNER JOIN (Nom_Detalle INNER JOIN Emp ON Nom_Detalle.CodEmp = Emp.CodEm" _
    & "p) ON Inmueble.CodInm = Emp.CodInm WHERE (((Nom_Detalle.IDNom) =" & ID & ")) GROUP BY Em" _
    & "p.CodInm, Inmueble.Ubica ORDER BY Emp.CodInm;"
    '
    With rstNom
        '
        .Open strSQL, cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
        '
        If Not .EOF And Not .BOF Then
        
            .MoveFirst
            
            Do
              If !sumSSO <> 0 Then
                Call RtnProUtility("Cargando Reintegro Seguro Social inm. " & !CodInm, _
                .AbsolutePosition * 6015 / .RecordCount)
                
                strSQL = "INSERT INTO AsignaGasto (Ndoc,CodGasto,Cargado,Descripcion,Fijo,Comun" _
                & ",Alicuota,Monto,Usuario,Fecha,Hora) In '" & gcPath & !Ubica & "inm." _
                & "mdb' SELECT 'NOMINA','192020','" & CargadoA & "',Trim(Titulo) & " & "' " & _
                Mid(ID, 2, 2) & "/" & Right(ID, 4) & "',False,True,True,'" & !sumSSO * -1 & "','" & _
                gcUsuario & "',Date(),Time() FROM TGastos In '" & gcPath & _
                !Ubica & "Inm.mdb' WHERE CodGasto='192020'"
                
                cnnConexion.Execute strSQL
              End If
              .MoveNext
            '
            Loop Until .EOF
            '
        End If
        '
        .Close
        '
    End With
    '
    'cargar ahora el reintegro de poilítica habitacional
    '
    strSQL = "SELECT Emp.CodInm, Sum(Nom_Detalle.LPH) AS sumLPH, Inmueble.Ubica FROM Inmueble I" _
    & "NNER JOIN (Nom_Detalle INNER JOIN Emp ON Nom_Detalle.CodEmp = Emp.CodEmp) ON Inmueble.Co" _
    & "dInm = Emp.CodInm WHERE (((Nom_Detalle.IDNom) =" & ID & ")) GROUP BY Emp.CodInm, Inmuebl" _
    & "e.Ubica ORDER BY Emp.CodInm;"
    '
    With rstNom
    '
        .Open strSQL, cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
        
        If Not .EOF And Not .BOF Then
        '
            .MoveFirst
            '
            Do
                '
              If !sumLPH <> 0 Then
                '
                Call RtnProUtility("Cargando Reintegro Ley Política Hab. inm. " & !CodInm, _
                .AbsolutePosition * 6015 / .RecordCount)
                '
                strSQL = "INSERT INTO AsignaGasto (Ndoc,CodGasto,Cargado,Descripcion,Fijo,Comun" _
                & ",Alicuota,Monto,Usuario,Fecha,Hora) In '" & gcPath & !Ubica & "inm." _
                & "mdb' SELECT 'NOMINA','192026','" & CargadoA & "',Trim(Titulo) & " & "' " & _
                Mid(ID, 2, 2) & "/" & Right(ID, 4) & "',False,True,True,'" & !sumLPH * -1 & "','" & _
                gcUsuario & "',Date(),Time() FROM TGastos In '" & gcPath & _
                !Ubica & "Inm.mdb' WHERE CodGasto='192026'"
                cnnConexion.Execute strSQL
                '
              End If
            .MoveNext
            '
            Loop Until .EOF
            '
        End If
        '
        .Close
        '
    End With
    '
    Call rtnBitacora("Generando Cuentas por  Pagar")
    'genera las cpp de los cheques
    rstNom.Open "qdfNomina_Bco", cnnConexion, adOpenKeyset, adLockOptimistic, adCmdTable
    rstNom.Filter = "Cuenta='' or Cuenta = Null"
    '
    'impresion de reportes (cartas al banco y listado de cheques)
    Call RtnProUtility("Imprimiendo pagos en cheques...", 0)
    '
    If Not rstNom.EOF And Not rstNom.BOF Then
        '
        rstNom.MoveFirst
        Dim cppDetalle As String
        '
        Do
            
            Call RtnProUtility("Cargando Cpp inm. " & rstNom("CodInm"), rstNom.AbsolutePosition * 6025 / rstNom.RecordCount)
            'GENERA LA FACTURA DE CPP Y ASIGNA EL GASTO PARA SACAR EL CHEQUE
            '
            strD = FrmFactura.FntStrDoc
            cppDetalle = Descripcion(rstNom!codGasto, rstNom!CodInm)
            '
            cnnConexion.Execute "INSERT INTO Cpp(Tipo,Ndoc,Fact,CodProv,Benef,Detalle,Monto,Ivm" _
            & ",Total,FRecep,Fecr,Fven,CodInm,Moneda,Estatus,Usuario,Freg) VALUES('NO','" & _
            strD & "','" & Left(ID, 3) & Right(ID, 2) & _
            Format(rstNom.AbsolutePosition, "00") & "','" & sysCodPro & "','" & rstNom("Name") & "' ,'" _
            & cppDetalle & " " & subT & "','" & rstNom("Neto") & "',0,'" & _
            rstNom("Neto") & "',Date(),Date(),'" & DateAdd("d", 30, Date) & "','" & _
            rstNom("CodInm") & "','BS','ASIGNADO','" & gcUsuario & "',Date())"
            '
            'ingresa el cargado
            '
            cnnConexion.Execute "INSERT INTO Cargado(Ndoc,CodGasto,Detalle,Periodo,Monto,Fecha," _
            & "Hora,Usuario) IN '" & gcPath & "\" & rstNom("CodInm") & "\inm.mdb' SELECT '" & _
            strD & "','" & rstNom("CodGasto") & "',Titulo,'" & CargadoA & "','" & rstNom("Neto") _
            & "',Date(),Time(),'" & gcUsuario & "' FROM Tgastos IN '" & gcPath & "\" & _
            rstNom("CodInm") & "\inm.mdb' WHERE CodGasto='" & rstNom("CodGasto") & "'"
            '
            rstNom.MoveNext
            '
        Loop Until rstNom.EOF
        '
        Call rtnBitacora("Emitiendo listado de cheques")
        'Call clear_Crystal(FrmAdmin.rptReporte)
        'listado de cheque
        Set crReporte = New ctlReport
        crReporte.Reporte = gcReport + "nom_chq.rpt"
        crReporte.OrigenDatos(0) = gcPath & "\sac.mdb"
        crReporte.Formulas(0) = "Subtitulo='" & subT & "'"
        'guarda una copia del reporte
        crReporte.Salida = crArchivoDisco
        crReporte.ArchivoSalida = RepNom & "CH" & ID & ".rpt"
        crReporte.Imprimir
'        'impresión en papel
        crReporte.Salida = crImpresora
        crReporte.Imprimir
        Set crReporte = Nothing
        
    End If
    '
    'reporte generel del banco
    Set crReporte = New ctlReport
    crReporte.Reporte = gcReport + "nom_gen.rpt"
    crReporte.OrigenDatos(0) = gcPath + "\sac.mdb"
    crReporte.Formulas(0) = "subtitulo='" & subT & "'"
    'guarda una copia local
    crReporte.Salida = crArchivoDisco
    crReporte.ArchivoSalida = RepNom & "GB" & ID & ".rpt"
    crReporte.Imprimir
    'envia a la impresora
    crReporte.Salida = crImpresora
    crReporte.Imprimir
    Set crReporte = Nothing
    
    Call RtnProUtility("Imprimiendo Reporte General...", 0)
    rstNom.Close
    '
    'imprime el reporte general
    '
    Call rtnBitacora("Emitiendo reporte general")
    rstTemp.Open "qdfNomina", cnnConexion, adOpenStatic, adLockPessimistic, adCmdTable
    '
    Set crReporte = New ctlReport
    crReporte.Reporte = gcReport + "nom_report.rpt"
    crReporte.OrigenDatos(0) = gcPath + "\sac.mdb"
    crReporte.Formulas(0) = "Titulo='Nomina'"
    crReporte.Formulas(1) = "SubTitulo='" & subT & "'"
    'guarda una copia local
    crReporte.Salida = crArchivoDisco
    crReporte.ArchivoSalida = RepNom & "RG" & ID & ".rpt"
    crReporte.Imprimir
    'envia a la impresora
    crReporte.Salida = crImpresora
    crReporte.Imprimir
    Set crReporte = Nothing
    
    'imprime el reporte general filtrado por inmueble
    If Not rstTemp.EOF And Not rstTemp.BOF Then
    '
        rstTemp.MoveFirst: Inm = ""
        rstTemp.Sort = "CodInm"
        Do
        
        Call RtnProUtility("Imprimiendo reporte general inm. '" & rstTemp("CodInm") & "'", _
        rstTemp.AbsolutePosition * 6025 / rstTemp.RecordCount)
        '
        If Inm <> rstTemp("CodInm") Then    'imprime una copia por inmueble
            Call rtnBitacora("Imprimiendo reporte general inm. " & Inm)
            '
            Set crReporte = New ctlReport
            crReporte.Reporte = gcReport + "nom_report.rpt"
            crReporte.OrigenDatos(0) = gcPath + "\sac.mdb"
            crReporte.Formulas(0) = "Titulo='Nomina " & rstTemp("CodInm") & "'"
            crReporte.Formulas(1) = "SubTitulo='" & subT & "'"
            crReporte.FormuladeSeleccion = "{qdfNomina.CodInm}='" & rstTemp("CodInm") & "'"
'                'guarda una copia en disco
            crReporte.Salida = crArchivoDisco
            crReporte.ArchivoSalida = RepNom & "RG" & rstTemp("CodInm") & ID & ".rpt"
            crReporte.Imprimir
            'imprime el reporte
            crReporte.Salida = crImpresora
            crReporte.Imprimir
            Inm = rstTemp("CodInm")
            '
        End If
        rstTemp.MoveNext
        Loop Until rstTemp.EOF
        '
    End If
    rstTemp.Close
    '
    Call RtnProUtility("Imprimiendo Cuadre de Nómina...", 6025)
    '
    'IMPRIME EL REPORTE CUADRE DE NOMINA
    Call rtnBitacora("Emitiando cuadre de nómina")
    strSQL = "SELECT Sum(qdfNomina_Bco.Neto) AS Total,Caja.DescripCaja,qdfNomina_Bco.Caja,False" _
    & " as Cheque,iif(qdfNomina_Bco.Caja='77','77',iif(qdfNomina_Bco.Caja='99','99','25')) & qd" _
    & "fNomina_Bco.Caja as Inm FROM Caja INNER JOIN (qdfNomina_Bco INNER JOIN Inmueble ON qdfNo" _
    & "mina_Bco.CodInm = Inmueble.CodInm) ON Caja.CodigoCaja = Inmueble.Caja WHERE qdfNomina_Bc" _
    & "o.Cuenta <> '' GROUP BY Caja.DescripCaja,qdfNomina_Bco.Caja  UNION SELECT qdfNomina_Bco." _
    & "Neto, qdfNomina_Bco.Name & ' ' & qdfNomina_Bco.NombreCargo, qdfNomina_Bco.Caja,True,qdfN" _
    & "omina_Bco.CodInm FROM Caja INNER JOIN (qdfNomina_Bco INNER JOIN Inmueble ON qdfNomina_Bc" _
    & "o.CodInm = Inmueble.CodInm) ON Caja.CodigoCaja = Inmueble.Caja WHERE (((qdfNomina_Bco.Cu" _
    & "enta) Is Null Or (qdfNomina_Bco.Cuenta) = '')) ORDER BY Cheque DESC , Inm;"
    
    Call rtnGenerator(gcPath & "\SAC.MDB", strSQL, "qdfNom_Cuadre")
    
'    Call clear_Crystal(FrmAdmin.rptReporte)
    Set crReporte = New ctlReport
    crReporte.Reporte = gcReport + "nom_cuadre.rpt"
    crReporte.OrigenDatos(0) = gcPath & "\sac.mdb"
    crReporte.Formulas(0) = "subtitulo='" & subT & "'"
    crReporte.Salida = crArchivoDisco
    crReporte.ArchivoSalida = RepNom & "CN" & ID & ".rpt"
    crReporte.Imprimir
'    'envia a la impresora
    crReporte.Salida = crImpresora
    crReporte.Imprimir
    '
    'si es la segunda nómina consulta la impresión de los recibos de pago
    '
    If Left(ID, 1) = 2 Then
        
        'IMPRIME REPORTE DE NOMINA PARA FACTURACION
        Call rtnBitacora("Emitiendo reporte de nómina vs. facturación")
        Call RtnProUtility("Imprimiendo Reporte Nómina vs. Facturación...", 6025)
        '
        'genera la consulta para imprimir reportes condensados de nómina (todo el mes)
        strSQL = "SELECT Emp.CodInm, Nom_Detalle.CodEmp, Emp.Apellidos, Emp.Nombres, No" _
        & "m_Detalle.Sueldo,Nom_Detalle.Dias_Trab,(Nom_Detalle.Sueldo+(Nom_Detalle.Sueldo* Nom_Detalle.Porc_BonoNoc/10" _
        & "0)) /30 *  Nom_Detalle.Dias_libres as DL, (Nom_Detalle.Sueldo + (Nom_Detalle.Sueldo *  Nom_Detalle.Porc_BonoNoc" _
        & " / 100)) /30 *  Nom_Detalle.Dias_Feriados * '" & pDF & "'  as DF, Nom_" _
        & "Detalle.Bono_Noc, Nom_Detalle.Bono_Otros, Nom_Detalle.Otras_Asignaciones, No" _
        & "m_Detalle.Dias_NoTrab, Nom_Detalle.SSO, Nom_Detalle.SPF, Nom_Detalle.LPH, No" _
        & "m_Detalle.Otras_Deducciones, Nom_Detalle.Descuento,Emp.CodGasto FROM Emp INN" _
        & "ER JOIN Nom_Detalle ON Emp.CodEmp = Nom_Detalle.CodEmp WHERE Nom_Detal" _
        & "le.IDNom=" & ID & " or Nom_Detalle.IDNom=1" & Mid(ID, 2, 100) & " ORDER BY Nom_Detalle.Sueldo"
        '
        Call rtnGenerator(gcPath & "\sac.mdb", strSQL, "qdfNomina")
        '
        Set crReporte = New ctlReport
        crReporte.Reporte = gcReport + "nom_fact.rpt"
        crReporte.OrigenDatos(0) = gcPath + "\sac.mdb"
        crReporte.Formulas(0) = "subtitulo='" & UCase(MonthName(Mid(ID, 2, 2))) & "-" & Right(ID, 4) _
            & "'"
        crReporte.Formulas(1) = "codNom0='" & codNom(0) & "'"
        crReporte.Formulas(2) = "codNom1='" & codNom(1) & "'"
        crReporte.Formulas(3) = "codNom2='" & codNom(2) & "'"
        crReporte.Salida = crArchivoDisco
        crReporte.ArchivoSalida = RepNom & "FA" & ID & ".rpt"
        crReporte.Imprimir
        'envia a la impresora
        crReporte.Salida = crImpresora
        crReporte.Imprimir
        Set crReporte = Nothing
        '
        Call RtnProUtility("Imprimiendo Reporte General Mes " & UCase(MonthName(Mid(ID, 2, 2))) _
        & "-" & Right(ID, 4), 6025)
        '
        'nomina general del mes
        Set crReporte = New ctlReport
        crReporte.Reporte = gcReport + "nom_report_gen.rpt"
        crReporte.OrigenDatos(0) = gcPath + "\sac.mdb"
        crReporte.Formulas(0) = "Titulo='Nomina'"
        crReporte.Formulas(1) = "SubTitulo='" & UCase(MonthName(Mid(ID, 2, 2))) _
        & "-" & Right(ID, 4) & "'"
'        copia local
        crReporte.Salida = crArchivoDisco
        crReporte.ArchivoSalida = RepNom & "GE" & ID & ".rpt"
        crReporte.Imprimir
        'eniva a la impresora
        crReporte.Salida = crImpresora
        crReporte.Imprimir
        Set crReporte = Nothing
        '
        '
        'SQL Recibo de Pago
'        strSQL = "PARAMETERS [IdNomina] Long; " & _
'                 "SELECT Inmueble.Nombre, Emp.Apellidos, Emp.Nombres, Emp.CodEmp, Emp.Cedula, " & _
'                 "Nom_Detalle.Sueldo, Sum(Nom_Detalle.Dias_Trab) AS DT, " & _
'                 "Sum((Emp.Sueldo+(Emp.Sueldo*Emp.BonoNoc/100))/30*Nom_Detalle.Dias_libres) AS BsDL, " & _
'                 "Sum(Nom_Detalle.Dias_libres) AS DL, " & _
'                 "Sum((Emp.Sueldo+(Emp.Sueldo*Emp.BonoNoc/100))/30*Nom_Detalle.Dias_Feriados*" & _
'                 "nom_calc.Dia_Feriado) AS BsDF, Sum(Nom_Detalle.Dias_Feriados) AS DF, " & _
'                 "Sum(Nom_Detalle.Dias_NoTrab) AS DNT, Sum(Nom_Detalle.Bono_Noc) AS BN, " & _
'                 "Sum(Nom_Detalle.Bono_Otros) AS OB, Sum(Nom_Detalle.Otras_asignaciones) AS OA, " & _
'                 "Sum(Nom_Detalle.SSO) AS SS, Sum(Nom_Detalle.SPF) AS PF, Sum(Nom_Detalle.LPH) AS PH, " & _
'                 "Sum(Nom_Detalle.Otras_Deducciones) AS OD, Sum(Nom_Detalle.Descuento) AS D, " & _
'                 "Inmueble.Rif FROM Nom_calc, Inmueble INNER JOIN (Nom_Detalle INNER JOIN Emp " & _
'                 "ON Nom_Detalle.CodEmp = Emp.CodEmp) ON Inmueble.CodInm = Emp.CodInm " & _
'                 "Where (((Nom_Detalle.IDNom) = [idnomina])) GROUP BY Inmueble.Nombre, " & _
'                 "Emp.Apellidos, Emp.Nombres, Emp.CodEmp, Emp.Cedula, Nom_Detalle.Sueldo, Inmueble.Rif;"
'
'        '
'        Call rtnGenerator(gcPath & "\sac.mdb", strSQL, "qdfNom_RP")
        
        '
        Resp = MsgBox("Desea imprimir ahora los recibos de pago?" & vbCrLf & "Si selecciona no," _
        & " podrá hacerlo más tarde.", vbQuestion + vbYesNo + vbDefaultButton2, App.ProductName)
        '
        Call printer_recibo_pago(ID, True, Resp, crImpresora)
        Unload FrmUtility
'        Set crReporte = New ctlReport
'        crReporte.Reporte = gcReport + "nom_rec_pago.rpt"
'        crReporte.OrigenDatos(0) = gcPath & "\sac.mdb"
'        crReporte.OrigenDatos(1) = gcPath & "\sac.mdb"
'        crReporte.OrigenDatos(2) = gcPath & "\sac.mdb"
'        crReporte.Parametros(0) = ID
'        crReporte.Formulas(0) = "subtitulo='MES:" & MonthName(Mid(ID, 2, 2)) & "-" & Right(ID, 4) & "'"
'        crReporte.Salida = crArchivoDisco
'        crReporte.ArchivoSalida = RepNom & "RP" & ID & ".rpt"
'        crReporte.Imprimir
'
'        Unload FrmUtility
'        If resp = vbYes Then
'        '
'            crReporte.Salida = crImpresora
'            crReporte.Imprimir
'            '
'        End If
'        Set crReporte = Nothing
        
    Else
        Unload FrmUtility
    End If
    '
    Set rstNom = Nothing
    Set rstTemp = Nothing
    Set rstCG = Nothing
    '
Cerrar:
    If Err.Number <> 0 Then
        Unload FrmUtility
        cnnConexion.RollbackTrans   'echa para atrás los cambios
        cierre_nomina = MsgBox(Err.Description, vbCritical, "Error " & Err.Number)
    Else
        cnnConexion.CommitTrans
    End If
    '
    End Function

    '---------------------------------------------------------------------------------------------
    '   Rutina:     Printer_Nom_Nov
    '
    '   Entradas:   Destino del reporte (Salida); Matriz de parámetros con los códigos de las
    '               nóminas que desea listar, si se omite imprime todas las nominas
    '
    '   rutina que envia el reporte de novedanes al destino señalado por el usuario
    '
    '---------------------------------------------------------------------------------------------
    Public Sub Printer_Nom_Nov(Salida As crSalida, ParamArray Nom())
    'variables locales
    Dim strSQL As String
    Dim strFiltro As String
    Dim errLocal As Long
    Dim strSTitulo As String
    Dim rpReporte As ctlReport
    '
    If UBound(Nom) >= 0 Then
        For I = 0 To UBound(Nom)
            strFiltro = strFiltro & IIf(strFiltro = "", "", " AND ") & "Nom_Temp.IDnomina =" & Nom(I)
        Next
        strFiltro = "WHERE " & strFiltro
        If UBound(Nom) = 0 Then
            strSTitulo = Left(Nom(0), 1) & "º Quincena de " & UCase(MonthName(Mid(Nom(0), 2, 2)) _
            & "/" & Right(Nom(0), 4))
        Else
            strSTitulo = "Desde: " & Left(Nom(0), 1) & "º Q. " & _
            UCase(MonthName(Mid(Nom(0), 2, 2)) & "/" & Right(Nom(0), 4)) & " Hasta: " & _
            Left(Nom(I - 1), 1) & "º Q. " & UCase(MonthName(Mid(Nom(I - 1), 2, 2)) & "/" & _
            Right(Nom(I - 1), 4))
        End If
    Else
        
    End If
    strSQL = "SELECT Nom_Temp.*,Emp.Apellidos & ', ' & Emp.Nombres AS Emp,Inmueble.CodInm FROM " _
    & "Inmueble INNER JOIN (Nom_Temp INNER JOIN Emp ON Nom_Temp.CodEmp = Emp.CodEmp) ON Inmuebl" _
    & "e.CodInm = Emp.CodInm " & strFiltro 'WHERE Nom_Temp.IDnomina"
    '
    Call rtnGenerator(gcPath & "\sac.mdb", strSQL, "qdfNom_Nov")
    'Call clear_Crystal(FrmAdmin.rptReporte)
    Set rpReporte = New ctlReport
    'envia el reporte
    With rpReporte
        .Reporte = gcReport & "nom_nov.rpt"
        .OrigenDatos(0) = gcPath & "\sac.mdb"
        .Formulas(0) = "subtitulo='" & strSTitulo & "'"
        .Salida = Salida
        .Imprimir
    End With
    Set rpReporte = Nothing
    '
    End Sub

    
    '------------------------------------------------------------------------------------------
    '   Funcion:    Complemento
    '
    '   Devuelve una cadena que contiene la ifnormación de la cuenta y banco donde depositar
    '------------------------------------------------------------------------------------------
    Private Function complemento() As String
    'variables locales
    Dim rstInm As ADODB.Recordset
    Dim rstCB As ADODB.Recordset
    Dim strSQL As String
    '
    Set rstInm = New ADODB.Recordset
    
    With rstInm
        .Open "Inmueble", cnnConexion, adOpenKeyset, adLockOptimistic, adCmdTable
        .Filter = "CodInm='" & gcCodInm & "'"
        'si existe conincidencia
        If Not .EOF And Not .BOF Then
        
            If .Fields("FactCB") Then
            
                Set rstCB = New ADODB.Recordset
                
                strSQL = "SELECT Cuentas.NumCuenta,Bancos.NombreBanco,Cuentas.Pred,Cuentas.Titular FROM Bancos " _
                & "INNER JOIN Cuentas ON Bancos.IDBanco = Cuentas.IDBanco ORDER BY Pred,NombreB" _
                & "anco;"
                
                If gnCta = CUENTA_POTE Then
                    rstCB.Open strSQL, cnnOLEDB + gcPath + "\" & sysCodInm & "\inm.mdb", _
                    adOpenKeyset, adLockOptimistic, adCmdText
                    rstCB.Filter = "Pred=True"
                Else
                    rstCB.Open strSQL, cnnOLEDB + mcDatos, adOpenKeyset, adLockOptimistic, adCmdText
                End If
                
                If Not rstCB.EOF And Not rstCB.BOF Then
                    
                    complemento = "DEPOSITE EN LA CTA. Nº " & rstCB("NumCuenta") & " DEL BCO. " _
                    & rstCB("NombreBanco") & " A NOMBRE DE: " & rstCB("TITULAR")
                    
                End If
                rstCB.Close
                
            End If
            
        End If
        '
        .Close
    End With
    '
    Set rstInm = Nothing
    Set rstCB = Nothing
    '
    End Function
    
    '---------------------------------------------------------------------------------------------
    '   funcion:    Nombre_Reporte
    '
    '   Entrada:    Codigo del inmueble
    '
    '   Retorne True si el condominio es administrado por admininistradora sac, de lo contrario
    '   retorna false
    '
    '---------------------------------------------------------------------------------------------
    Private Function Nombre_Reporte(Inmueble As String) As Boolean
    'variables locales
    Dim rst As New ADODB.Recordset
    '
    rst.Open "SELECT * FROM Inmueble WHERE CodInm='" & Inmueble & "';", cnnConexion, _
    adOpenKeyset, adLockOptimistic, adCmdText
    If Not rst.EOF And Not rst.BOF Then
        Nombre_Reporte = rst("SAC")
    End If
    '
    rst.Close
    Set rst = Nothing
    
    End Function


'    Public Sub r()
'    Call clear_Crystal(FrmAdmin.rptReporte)
'        'nomina general del mes
'        With FrmAdmin.rptReporte
'        '
'            .ReportFileName = gcReport + "nom_report.rpt"
'            .DataFiles(0) = gcPath + "\sac.mdb"
'            .Formulas(0) = "Titulo='Nomina'"
'            '.Formulas(1) = "SubTitulo='" & UCase(MonthName(Mid(ID, 2, 2))) _
'            & "-" & Right(ID, 4) & "'"
'
'            '.SectionFormat(0) = "DETAIL;F;X;X;X;X;X;X"
'            .SectionFormat(0) = "GF1;T;x;x;X;X;X;X"
'            '.SectionFormat(2) = "GH1;T;X;X;X;X;X;X"
''            .Destination = crptToFile
''            .PrintFileType = crptCrystal
''            .PrintFileName = RepNom & "GE" & ID & ".rpt"
''            errlocal = .PrintReport
''            If errlocal <> 0 Then
''                MsgBox "Ocurrio el siguiente error cuando se trato de guardar en dicos el repor" _
''                & "te general de la nómina: " & .LastErrorString, vbCritical, App.ProductName
''            End If
'            .Destination = crptToWindow
'            errlocal = .PrintReport
'            If errlocal <> 0 Then
'                MsgBox "Ocurrio el siguiente error cuando se trato de imprimir el reporte gener" _
'                & "al de la nómina: " & .LastErrorString, vbCritical, App.ProductName
'            End If
'            '
'        End With
'    End Sub

    Private Function Deuda_Acumulada(Codigo_Inmueble$, Saldo@) As Currency
    'variables locales
    Dim rstlocal As New ADODB.Recordset
    '
    rstlocal.Open "SELECT * FROM Inmueble WHERE CodInm='" & Codigo_Inmueble & "'", cnnConexion, _
    adOpenKeyset, adLockOptimistic, adCmdText
    If Not rstlocal.EOF And Not rstlocal.BOF Then
        Deuda_Acumulada = (Saldo * rstlocal!HonoMorosidad / 100) + Saldo
    End If
    '
    rstlocal.Close
    Set rstlocal = nohting
    '
    End Function

    Public Function LoadResCustom(ID, Optional Tipo) As Variant
    'variables locales
    Dim bytes() As Byte, idf As Integer, snd As String
    '
    
    If Dir(App.Path & "tempimg.tmp", vbArchive) <> "" Then
        idf = FreeFile
        Open App.Path & "tempimg.tmp" For Binary As #idf
            bytes = LoadResData(ID, "CUSTOM")
        Put #idf, , bytes
        Close #idf
        Set LoadResCustom = LoadPicture(App.Path & "tempimg.tmp")
    End If
    '
    End Function

'    Public Sub aguinaldos()
'    Dim rstLocal As New ADODB.Recordset
'    Dim strSql As String
'    Dim N As Long
'    rstLocal.Open "Inmueble", cnnConexion, adOpenKeyset, adLockOptimistic, adCmdTable
'    rstLocal.Filter = "Inactivo = False"
'
'    If Not rstLocal.EOF And Not rstLocal.BOF Then
'        rstLocal.MoveFirst
'
'        Do
'            strSql = "UPDATE AsignaGasto IN '" & gcPath & rstLocal("Ubica") & "\inm.mdb' " _
'        & "SET Ndoc = 'AGU2004' WHERE CodGasto ='190002' AND Cargado =#10/01/200" _
'        & "4# AND Ndoc='NOV';"
'            cnnConexion.Execute strSql, N
'            Debug.Print rstLocal!CodInm & " - " & N
'            rstLocal.MoveNext
'        Loop Until rstLocal.EOF
'        MsgBox "Fin archivo"
'    End If
'    End Sub

    
Public Sub Configurar_ADO(ctlADO As Adodc, TipoComando As ADODB.CommandTypeEnum, _
OrigenDatos As String, Ubica_Datos As String, Optional CampoOrden As String)

With ctlADO

    .ConnectionString = cnnOLEDB & Ubica_Datos
    .CursorLocation = adUseClient
    .CommandType = TipoComando
    .Mode = adModeShareDenyNone
    .LockType = adLockOptimistic
    .RecordSource = OrigenDatos
    .Refresh
    If CampoOrden <> "" Then .Recordset.Sort = CampoOrden
    
End With

End Sub



    '---------------------------------------------------------------------------------------------
    '   Rutina:     Cabeza
    '
    '   Entradas:   Pagina (número de la página) entero largo
    '
    '   Escribe el encabezado en cada pág. del reporte
    '---------------------------------------------------------------------------------------------
    Public Sub cabeza(Pagina&, Titulo1, Titulo2$)
    '
    Call RtnProUtility("Procesando Pág." & Pagina)
    
    If Dir(gcUbiGraf & "\logo.bmp") <> "" Then   'imprime el logo de la empresa
        Printer.PaintPicture LoadPicture(gcUbiGraf & "\logo.bmp"), _
        Printer.ScaleTop + 10, Printer.ScaleLeft, 4290, 1365, , , , , vbSrcCopy
        Printer.Print
    End If
    
    Printer.FontSize = 24
    Printer.FontName = "Times New Roman"
    Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth(Titulo1) - 1000
    Printer.Print Titulo1
    Printer.FontSize = 12
    If Titulo2 <> "" Then
        Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth(Titulo2) - 1100
        Printer.Print Titulo2
    End If
    Printer.FontSize = 10
    Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("Fecha de Impresión: " & _
    Format(Date, "dd/mm/yy")) - 1100
    Printer.Print "Fecha de Impresión: " & Format(Date, "dd/mm/yy")
    Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("Pág.: " & Pagina) - 1100
    Printer.Print "Pág.: " & Pagina
    Printer.FontSize = 8
    Printer.CurrentY = 1400
    '
    End Sub


'---------------------------------------------------------------------------------------------
    '   Función:        linea_detalle
    '
    '   Entradas: cod. Inm. rst (recordset), linea
    '
    '   Imprime el registro activo y devuelve la linea final utilizada
    '---------------------------------------------------------------------------------------------
    Public Function linea_Detalle(Inm$, rst As ADODB.Recordset, Linea&, _
    Optional r As Boolean, Optional Facturado$) As Long
    'variables locales
    Dim Cuenta As String
    Dim Des As String
    Dim Cargado As String
    Dim Monto As String
    'asignación de las vairables
    If r Then
        Cuenta = Facturado
    Else
        Cuenta = Inm & "-" & rst("CodGasto")
    End If
    Des = rst("Detalle")
    Cargado = Format(rst("Periodo"), "MM/YYYY")
    Monto = Format(rst("Monto"), "#,##.00")
    '-------------------
    Printer.CurrentY = Linea
    Printer.CurrentX = 1200
    Printer.Print Cuenta    'print código de cuenta
    Printer.CurrentY = Linea
    Printer.CurrentX = 2400
    Printer.Print Cargado   'print cargado
    Printer.CurrentY = Linea
    Printer.CurrentX = 3400
    Printer.Print Des   'print. descripción
    Printer.CurrentY = Linea
    Printer.CurrentX = 9000 - Printer.TextWidth(Monto)
    Printer.Print Monto 'print. monto
    linea_Detalle = Printer.CurrentY
    '
    End Function

    '---------------------------------------------------------------------------------------------
    '   Función:    Registro_Proveedor
    '
    '   Entradas:   Nombre del proveedor, Cpp(si/no) registra cpp
    '
    '   Esta función registra a un proveedor y si opcional mente
    '   carga una factura al sistema
    '---------------------------------------------------------------------------------------------
    Public Function Registro_Proveedor(Nombre$, Cpp As Boolean, ParamArray Factura()) As String
    'variables locales
    Dim rstlocal As ADODB.Recordset
    Dim strSQL As String
    Dim NDoc As String
    Dim Npro As String
    Dim N As Long
    Call rtnBitacora("Registro de Proveedor" & IIf(Cpp, " y Cpp", ""))
    cnnConexion.BeginTrans
    On Error GoTo salir:
    '
    Set rstlocal = New ADODB.Recordset
    
    With rstlocal
    
        .Open "Proveedores", cnnConexion, adOpenKeyset, adLockOptimistic, adCmdTable
        
        If Not .EOF And Not .BOF Then
            .Find "NombProv='" & Nombre & "'"
            
            If .EOF Then
                'no registrado el proveedor
                .Close
                .Open "SELECT Max(Clng(Codigo)) FROM Proveedores", cnnConexion, adOpenKeyset, _
                adLockOptimistic, adCmdText
                Npro = .Fields(0) + 1
                Npro = Format(Npro, "0000")
                'registro del proveedor
                strSQL = "INSERT INTO Proveedores(Codigo,NombProv,Beneficiario,Usuario,Freg,Nac" _
                & "ional,FecReg) VALUES ('" & Npro & "','" & Nombre & "','" & Nombre & "','" & gcUsuario _
                & "',Date(),-1,Date())"
                cnnConexion.Execute strSQL, N
                Call rtnBitacora("Ingresado " & N & " proveedor- Nº Prov.: " & Npro)
            Else
                Npro = !Codigo
            End If
        Else
            Npro = 1
        End If
        Npro = Format(Npro, "0000")
        
        If Cpp Then 'registro de la factura en Cpp
            
            NDoc = FrmFactura.FntStrDoc
            '----------------------------------------------------------------
            '   valores de los indices de la matriz Factura
            '   0   Nº de Factura
            '   1   Detalle
            '   2   Monto
            '   3   CodInm
            '----------------------------------------------------------------
            strSQL = "INSERT INTO Cpp(Tipo,Ndoc,Fact,CodProv,Benef,Detalle,Monto,Ivm,Total,FRecep," _
            & "Fven,Fecr,CodInm,Moneda,Estatus,Usuario,Freg) VALUES('FA','" & NDoc & "','" & _
            Factura(0) & "','" & Npro & "','" & Nombre & "','" & Factura(1) & "','" & Factura(2) & _
            "'," & "0,'" & Factura(2) & "',Date(),Date(),Date(),'" & Factura(3) & "','Bs','PENDIENTE','" & _
            gcUsuario & "',Date())"
            
            cnnConexion.Execute strSQL
            Call rtnBitacora("Ingresado Factura #" & NDoc & " en Cpp")
            
        End If
        
salir: If Err <> 0 Then
            cnnConexion.RollbackTrans
            Call rtnBitacora("Transacción Fallida")
        Else
            cnnConexion.CommitTrans
            Call rtnBitacora("Transacción Ok.")
            Registro_Proveedor = NDoc
        End If
        '
    End With
    '
    End Function

Public Sub addToolTip(Contacto As String, Gestion As String, Por As String)
Dim INI&, lTop&
Set frmTT = New frmToolTip
frmTT.lbl(1) = Contacto
frmTT.lbl(2) = Gestion
frmTT.lbl(4) = Por
INI = Screen.Height
frmTT.Top = INI
frmTT.Show
'Call SetWindowPos(frmTT.hWnd, -1, 0, 0, 0, 0, 1 Or 2)
lTop = Screen.Height - frmTT.Height - 500
frmTT.Left = Screen.Width - frmTT.Width
For I = INI To lTop Step -10
    frmTT.Top = I
    DoEvents
Next

End Sub

Public Sub EliminaFrmGestion()
Dim F As Form
For Each F In Forms
    If F.Name = "frmToolTip" Then Unload F
Next
End Sub

Public Sub UpdateStatus(pic As PictureBox, ByVal sngPercent As Long, _
Optional ByVal fBorderCase)
Dim strPercent As String
Dim intX As Integer
Dim intY As Integer
Dim intWidth As Integer
Dim intHeight As Integer

If sngPercent > 100 Then sngPercent = 100
If IsMissing(fBorderCase) Then fBorderCase = False

'For this to work well, we need a white background and any color foreground (blue)
Const colBackground = &HFFFFFF ' white
Const colForeground = &H800000       'blue

pic.ForeColor = colForeground
pic.BackColor = colBackground

'
'Format percentage and get attributes of text
'
Dim intPercent
intPercent = sngPercent

'Never allow the percentage to be 0 or 100 unless it is exactly that value.  This
'prevents, for instance, the status bar from reaching 100% until we are entirely done.
If intPercent = 0 Then
    If Not fBorderCase Then
        intPercent = 1
    End If
ElseIf intPercent = 100 Then
    If Not fBorderCase Then
        intPercent = 99
    End If
End If

strPercent = Format$(intPercent) & "%"
intWidth = pic.TextWidth(strPercent)
intHeight = pic.TextHeight(strPercent)

'
'Now set intX and intY to the starting location for printing the percentage
'
intX = pic.Width / 2 - intWidth / 2
intY = (pic.Height / 2) - (intHeight / 2) - 35

'
'Need to draw a filled box with the pics background color to wipe out previous
'percentage display (if any)
'

'
'Now fill in the box with the ribbon color to the desired percentage
'If percentage is 0, fill the whole box with the background color to clear it
'Use the "Not XOR" pen so that we change the color of the text to white
'wherever we touch it, and change the color of the background to blue
'wherever we touch it.
'
pic.DrawMode = 10 ' Not XOR Pen
pic.Line (intX, intY)-Step(intWidth, intHeight), pic.BackColor, BF

pic.CurrentX = intX
pic.CurrentY = intY
pic.Print strPercent


If sngPercent > 0 Then
    pic.Line (0, 0)-(sngPercent * pic.Width / 100, pic.Height), colForeground, BF
Else
    pic.Line (0, 0)-(pic.Width, pic.Height), colForeground, BF
End If
pic.DrawMode = 13 ' Copy Pen
'pic.Line (intX, intY)-Step(intWidth, intHeight), pic.BackColor, BF

'
'Back to the center print position and print the text
'

pic.Refresh
End Sub

    
Public Sub Procesar_GastosMenores()
'variables locales
Dim rstInm As ADODB.Recordset
Dim rstGM As ADODB.Recordset
Dim rstlocal As ADODB.Recordset
Dim cnnGM As ADODB.Connection
Dim strSQL As String, NDoc As String
Dim Reporte(), I&
Dim mPeriodo As Date, Descripcion As String

Set rstInm = New ADODB.Recordset

rstInm.CursorLocation = adUseClient
rstInm.Open "Inmueble", cnnConexion, adOpenKeyset, adLockOptimistic, adCmdTable
rstInm.Filter = "Inactivo=False"

If Not (rstInm.EOF And rstInm.BOF) Then
    Call RtnConfigUtility(True, "Gastos Fijos Menores", "Iniciando...", _
    "Seleccionando Información")
    rstInm.MoveFirst
    Set rstGM = New ADODB.Recordset
    Set cnnGM = New ADODB.Connection
    Set rstlocal = New ADODB.Recordset
    rstGM.CursorLocation = adUseClient
    Do
        Call RtnProUtility("Procesando Inm: " & rstInm!CodInm, 6015 * _
        (rstInm.AbsolutePosition / rstInm.RecordCount))
        cnnGM.Open cnnOLEDB & gcPath & rstInm!Ubica & "inm.mdb"
        
        
        'If rstInm!CodInm = "2552" Then Stop
        strSQL = "SELECT GastosMenores.*, Tgastos.MontoFijo, TGastos.Fijo, TGastos.Comun, " _
        & "TGastos.Alicuota FROM GastosMenores INNER JOIN Tgastos ON GastosMenores.CodGasto " _
        & "= Tgastos.CodGasto;"

        rstGM.Open strSQL, cnnGM, adOpenKeyset, adLockOptimistic, adCmdText
        If Not (rstGM.EOF And rstGM.BOF) Then
            strSQL = "SELECT Max(Periodo) as D FROM factura WHERE Fact not LIKE 'CHD%'"
            Set rstlocal = cnnGM.Execute(strSQL)
            If (Not (rstlocal.EOF And rstlocal.BOF)) And Not IsNull(rstlocal("d")) Then
                mPeriodo = DateAdd("m", 1, rstlocal("D"))
            Else
20              mPeriodo = InputBox("Ingrese el período a facturar", "Gastos Fijos Menores")
            End If
            rstlocal.Close
            If Not IsDate(mPeriodo) Then GoTo 20
            rstGM.MoveFirst
            Do
                strSQL = "SELECT * FROM Cargado WHERE CodGasto='" & _
                rstGM!codGasto & "' AND NDoc<>'0' and Periodo=#" & _
                Format(mPeriodo, "mm/dd/yy") & "#"
                
                rstlocal.Open strSQL, cnnGM, adOpenKeyset, _
                adLockOptimistic, adCmdText
                
                If rstlocal.EOF And rstlocal.BOF Then
                    'agrega la Cpp
                    NDoc = FrmFactura.FntStrDoc
                    If rstGM!Descripcion Like "MANTENIMIENTO*" Then
                        Descripcion = "MTTO. " & Mid(rstGM!Descripcion, InStr(rstGM!Descripcion, " ") + 1, Len(rstGM!Descripcion)) & " MES " & Format(mPeriodo, "mm/yyyy")
                    Else
                        Descripcion = "CANC. " & rstGM!Descripcion
                    End If
                    
                    strSQL = "INSERT INTO Cpp(Tipo,Ndoc,Fact,CodProv,Benef,Detalle,Monto,Ivm," _
                    & "Total,FRecep,Fecr,FVen,CodInm,Moneda,Estatus,Usuario,Freg) VALUES('GM','" & _
                    NDoc & "','" & Right(rstInm!CodInm, 3) & Format(mPeriodo, "mmyy") & "','" & rstGM!codProv & _
                    "','" & rstGM!Nombre & "','" & Descripcion & " " & rstInm!Nombre & _
                    "(" & rstInm!CodInm & ")','" & Replace(CCur(rstGM!MontoFijo), ",", ".") & _
                    "',0,'" & Replace(CCur(rstGM!MontoFijo), ",", ".") & "',Date(),Date()," _
                    & "dateadd('d',30,Date()),'" & rstInm!CodInm & "','Bs.','ASIGNADO','" & _
                    gcUsuario & "',date())"
                    cnnConexion.Execute strSQL
                    '
                    'agrega la asignación del gasto
                    strSQL = "INSERT INTO AsignaGasto(NDoc,CodGasto,Cargado,Descripcion,Fijo," _
                    & "Comun,Alicuota,Monto,Usuario,Fecha,Hora) VALUES('" & NDoc & "','" & _
                    rstGM!codGasto & "','" & mPeriodo & "','" & rstGM!Descripcion & "'," & IIf(rstGM!Fijo, -1, 0) & _
                    "," & IIf(rstGM!Comun, -1, 0) & "," & IIf(rstGM!Alicuota, -1, 0) & ",'" & _
                    Replace(CCur(rstGM!MontoFijo), ",", ".") & "','" & gcUsuario & "',Date()," _
                    & "Time())"
                    cnnGM.Execute strSQL
                    'agrega el cargado
                    strSQL = "INSERT INTO Cargado(Ndoc,CodGasto,Detalle,Periodo,Monto,Fecha," & _
                    "Hora,Usuario) VALUES('" & NDoc & "','" & rstGM!codGasto & "','" & _
                    rstGM!Descripcion & "','" & mPeriodo & "','" & Replace(CCur(rstGM!MontoFijo), ",", ".") _
                    & "',Date(),Time(),'" & gcUsuario & "')"
                    cnnGM.Execute strSQL
                    '
                    ReDim Preserve Reporte(5, I)
                    Reporte(0, I) = rstInm!CodInm
                    Reporte(1, I) = rstInm!Nombre
                    Reporte(2, I) = NDoc
                    Reporte(3, I) = rstGM!Nombre
                    Reporte(4, I) = rstGM!Descripcion
                    Reporte(5, I) = Format(rstGM!MontoFijo, "#,##0.00")
                    DoEvents
                    I = I + 1
                    
                End If
                rstlocal.Close
                rstGM.MoveNext
            Loop Until rstGM.EOF
        End If
        rstGM.Close
        cnnGM.Close
        rstInm.MoveNext
    Loop Until rstInm.EOF
    Set rstGM = Nothing
    Set cnnGM = Nothing
End If
rstInm.Close

Set rstInm = Nothing
'impresión del reporte
If I > 0 Then
    Dim Y&
    Dim strNombre As String
    Dim strBenef As String
    Dim strDescrip As String
    Dim strMonto As String
    Dim Total As Double
    
    Printer.FontBold = True
    Printer.Print "Gastos Fijos Menores"
    Printer.Print "Procesados"
    Printer.Print "Fecha: " & Date
    Printer.FontBold = False
    Printer.Print
    Y = Printer.CurrentY
    Printer.Print "Cod.Inm"
    Printer.CurrentY = Y
    Printer.CurrentX = 1000
    Printer.Print "Nº Doc."
    Printer.CurrentX = 2000
    Printer.CurrentY = Y
    Printer.Print "Nombre/Razón Social"
    Printer.CurrentY = Y
    Printer.CurrentX = 5000
    Printer.Print "Beneficiario"
    Printer.CurrentY = Y
    Printer.CurrentX = 7500
    Printer.Print "Descripción"
    Printer.CurrentY = Y
    Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("Monto")
    Printer.Print "Monto"
    Printer.CurrentY = Printer.CurrentY + 50
    For I = 0 To UBound(Reporte, 2)
        Call RtnProUtility("Imprimiendo registros " & I + 1 & "/" & UBound(Reporte, 2) + 1)
        Y = Printer.CurrentY
        strMonto = Reporte(5, I)
        Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth(strMonto)
        Printer.Print strMonto
        Total = Total + CDbl(strMonto)
        Printer.CurrentY = Y
        Printer.CurrentX = 7400
        strDescrip = Reporte(4, I)
        If Printer.TextWidth(strDescrip) > 3000 Then
            Do
                strDescrip = Left(strDescrip, Len(strDescrip) - 1)
            Loop Until Not Printer.TextWidth(strDescrip) > 3000
        End If
        Printer.Print strDescrip
        
        Printer.CurrentY = Y
        Printer.CurrentX = 4900
        strBenef = Reporte(3, I)
        If Printer.TextWidth(strBenef) > 2400 Then
            Do
                strBenef = Left(strBenef, Len(strBenef) - 1)
            Loop Until Not Printer.TextWidth(strBenef) > 2400
        End If
        Printer.Print strBenef
        
        Printer.CurrentY = Y
        Printer.CurrentX = 1900
        strNombre = Reporte(1, I)
        'If Reporte(0, I) = "2560" Then Stop
        If Printer.TextWidth(strNombre) > 2800 Then
            Do
                strNombre = Left(strNombre, Len(strNombre) - 1)
            Loop Until Not Printer.TextWidth(strNombre) > 2800
        End If
        Printer.Print strNombre
        
        Printer.CurrentY = Y
        Printer.CurrentX = 1000
        Printer.Print Reporte(2, I)
        
        Printer.CurrentY = Y
        Printer.CurrentX = 100
        Printer.Print Reporte(0, I)
        
        Next
    Printer.Print
    Printer.FontBold = True
    Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("Total Bs. " & Format(Total, "#,##0.00"))
    Printer.Print "Total Bs. "; Format(Total, "#,##0.00")

    Printer.EndDoc
    
    Call RtnProUtility("Gastos Menores procesados (" & I & ")")
    MsgBox "Proceso llevado a cabo con éxito." & vbCrLf & _
    "Retire el reporte en la impresora.", vbInformation, App.ProductName
Else
    Call rtnBitacora("Gastos Menores procesados ya procesados")
    MsgBox "No se procesaron registros." & vbCrLf & _
    "Antes de consultar con el administrador del sistema," & vbCrLf & _
    "VERIFIQUE SI ESTE PROCESO FUE EJECUTADO CON ANTERIORIDAD", vbInformation, App.ProductName
End If
Unload FrmUtility

End Sub

Public Sub reporte_gm_general()
Dim rstlocal As ADODB.Recordset
Dim cnnLocal As ADODB.Connection
Dim strSQL As String

With FrmAdmin.objRst
    .Filter = "Inactivo=False"
    If Not (.EOF And .BOF) Then
        Call RtnConfigUtility(True, "Reporte General Gastos Menores", "Iniciando...", "Seleccionando Información")
        .MoveFirst
        Set rstlocal = New ADODB.Recordset
        Set cnnLocal = New ADODB.Connection
        Printer.FontBold = True
        Printer.Print "Gastos Fijos Menores"
        Printer.Print "Registrados"
        Printer.Print "Fecha: " & Date
        Printer.FontBold = False
        Printer.Print
        Y = Printer.CurrentY
        Printer.Print "Cod.Inm"
        Printer.CurrentY = Y
        Printer.CurrentX = 1000
        Printer.Print "Nº Doc."
        Printer.CurrentX = 2000
        Printer.CurrentY = Y
        Printer.Print "Nombre/Razón Social"
        Printer.CurrentY = Y
        Printer.CurrentX = 5000
        Printer.Print "Beneficiario"
        Printer.CurrentY = Y
        Printer.CurrentX = 7500
        Printer.Print "Descripción"
        Printer.CurrentY = Y
        Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("Monto")
        Printer.Print "Monto"
        Printer.CurrentY = Printer.CurrentY + 50
        Do
            strSQL = "SELECT GastosMenores.*,TGastos.* FROM GastosMenores INNER " _
            & "JOIN TGastos ON GastosMenores.CodGasto = TGastos.CodGasto"
            cnnLocal.Open cnnOLEDB & gcPath & !Ubica & "inm.mdb"
            rstlocal.Open strSQL, cnnLocal, adOpenKeyset, adLockOptimistic, adCmdText
            
            Call RtnProUtility("Procesando Inm: " & !CodInm, 6015 * (.AbsolutePosition / .RecordCount))
            DoEvents
            If Not (rstlocal.EOF And rstlocal.BOF) Then
                rstlocal.MoveFirst
                Do
                    Y = Printer.CurrentY
                    strMonto = Format(IIf(IsNull(rstlocal!MontoFijo), 0, rstlocal!MontoFijo), "#,##0.00")
                    Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth(strMonto)
                    Printer.Print strMonto
                    Total = Total + CDbl(strMonto)
                    Printer.CurrentY = Y
                    Printer.CurrentX = 7400
                    strDescrip = rstlocal!Descripcion
                    If Printer.TextWidth(strDescrip) > 3000 Then
                        Do
                            strDescrip = Left(strDescrip, Len(strDescrip) - 1)
                        Loop Until Not Printer.TextWidth(strDescrip) > 3000
                    End If
                    Printer.Print strDescrip
                    
                    Printer.CurrentY = Y
                    Printer.CurrentX = 4900
                    strBenef = rstlocal!Nombre
                    If Printer.TextWidth(strBenef) > 2400 Then
                        Do
                            strBenef = Left(strBenef, Len(strBenef) - 1)
                        Loop Until Not Printer.TextWidth(strBenef) > 2400
                    End If
                    Printer.Print strBenef
                    
                    Printer.CurrentY = Y
                    Printer.CurrentX = 1900
                    strNombre = !Nombre
                    If Printer.TextWidth(strNombre) > 2800 Then
                    Do
                        strNombre = Left(strNombre, Len(strNombre) - 1)
                        Loop Until Not Printer.TextWidth(strNombre) > 2800
                    End If
                    Printer.Print strNombre
                
                    Printer.CurrentY = Y
                    Printer.CurrentX = 1200
                    Printer.Print "--"
                
                    Printer.CurrentY = Y
                    Printer.CurrentX = 100
                    Printer.Print !CodInm
                    rstlocal.MoveNext
                Loop Until rstlocal.EOF
            End If
            rstlocal.Close
            cnnLocal.Close
            .MoveNext
        Loop Until .EOF
        Printer.Print
        Printer.FontBold = True
        Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("Total Bs. " & Format(Total, "#,##0.00"))
        Printer.Print "Total Bs. "; Format(Total, "#,##0.00")
    
        Printer.EndDoc
        Set rstlocal = Nothing
        Set cnnLocal = Nothing
    End If
    .Filter = 0
    .MoveFirst
    Unload FrmUtility
    MsgBox "Reporte impreso con éxito", vbInformation, App.ProductName
End With
End Sub

Function Descripcion(codGasto$, Inm$) As String
Dim rstlocal As ADODB.Recordset

Set rstlocal = New ADODB.Recordset

rstlocal.Open "SELECT * FROM TGastos WHERE CodGasto='" & codGasto & "'", cnnOLEDB + _
gcPath + "\" + Inm + "\inm.mdb", adOpenKeyset, adLockPessimistic, adCmdText

If Not (rstlocal.EOF And rstlocal.BOF) Then
    Descripcion = rstlocal("Titulo")
End If

rstlocal.Close
Set rstlocal = Nothing

End Function

'//**************************************************************************************
'// DE AQUI EN ADELANTE SE CODIFICARA LA EMISIÓN DE TODOS LOS REPORTES
'//**************************************************************************************

    '---------------------------------------------------------------------------------------------
    '   Rutina:     Printer_Analisis_Facturacion
    '
    '---------------------------------------------------------------------------------------------
    Public Sub Printer_Analisis_Facturacion(strPeriodo$, ByVal datFacturado$, curDeudaT@, _
    Salida As crSalida)
    'Variables locales---------------------------
    Dim archivo As String
    Dim Ruta As String
    Dim Nombre As String
    Dim rpReporte As ctlReport
    '--------------------------------------------
    Ruta = gcPath & gcUbica & "Reportes\"
    Nombre = "AN" & Left(strPeriodo, 3) & Right(strPeriodo, 2) & ".rpt"
    archivo = Ruta & Nombre
    '
    Set rpReporte = New ctlReport
    With rpReporte
        'Reporte Análisis de Facturación
        .Reporte = archivo
        If Dir(archivo) = "" Then
            '
            .Formulas(0) = "Inmueble='" & gcCodInm & " - " & gcNomInm & "'"
            .Formulas(1) = "Total_Fact=" & Replace(curDeudaT, ",", ".")
            .Formulas(2) = "GC=" & Replace(ftnFact(1, datFacturado), ",", ".")
            .Formulas(3) = "FR=" & Replace(ftnFact(0, datFacturado), ",", ".")
            .Formulas(4) = "Periodo='" & strPeriodo & "'"
            .Formulas(5) = "Ajus=" & Replace(AjusteGC, ",", ".")
            .Reporte = gcReport + "fact_analisis.rpt"
            .OrigenDatos(0) = mcDatos
            '.OrigenDatos(1) = gcPath & "\sac.mdb"
            .Salida = crArchivoDisco
            .ArchivoSalida = archivo
            .Imprimir
            '
        End If
        .Salida = Salida
        If Salida = crPantalla Then
            .TituloVentana = "Analisis Facturación " & gcCodInm & "/" & strPeriodo
        End If
        
        .Imprimir
        '
    End With
    '
    Set rpReporte = Nothing
    End Sub

    '---------------------------------------------------------------------------------------------
    '   Rutina:     Printer_PreRecibo
    '
    '---------------------------------------------------------------------------------------------
    Public Sub Printer_PreRecibo(strPeriodo$, ByVal datFacturado$, _
    Optional Proceso%, Optional Salida As crSalida)
    'variables locales
    Dim archivo$, Destino%, ErrorLocal&
    Dim Existe As Boolean
    Dim rpReporte As ctlReport
    '------------------------------------
    archivo = gcPath & gcUbica & "Reportes\" & "PR" & Left(strPeriodo, 3) & Right(strPeriodo, 2) _
    & ".rpt"
    If Not Dir(archivo) = "" Then Existe = True
    '
    'Call clear_Crystal(ctlReport)
    '
    Set rpReporte = New ctlReport
    With rpReporte  ''Pre-Recibo
        '.Reset
        '.ProgressDialog = False
        
        '.ReportFileName = IIf(Existe = True, Archivo, gcReport + "prerecibo.rpt")
        .Reporte = IIf(Existe = True, archivo, gcReport + "prerecibo.rpt")
        If Not Existe Then .OrigenDatos(0) = mcDatos '.DataFiles(0) = mcDatos
        .Formulas(0) = "Condominio ='" & gcCodInm & " - " & gcNomInm & "'"
        .Formulas(1) = "PorcFondo=" & Replace(gnPorcFondo, ",", ".")
        .Formulas(2) = "Periodo='" & strPeriodo & "'"
        If Not Existe And Proceso = 1 Then
            Call Print_PreRecibo(CDate(datFacturado))
            '.ReportFileName = Archivo
            '.Destination = crptToFile
            '.PrintFileName = Archivo
            '.PrintFileType = crptCrystal
            'ErrorLocal = .PrintReport
            .Reporte = archivo
            .Salida = crArchivoDisco
            .ArchivoSalida = archivo
            .Imprimir
'            If errlocal <> 0 Then
'                MsgBox ctlReport.LastErrorString, vbCritical, ctlReport.LastErrorNumber
'                Call rtnBitacora("Print Prerecibo /Err." & ctlReport.LastErrorNumber & " " & ctlReport.LastErrorString)
'                errlocal = 0
'            End If
        End If
        
        
        
        '.Destination = Salida
        .Salida = Salida
        If .Salida = crPantalla Then
            '.WindowState = crptMaximized
            '.WindowShowCloseBtn = True
            '.WindowTitle = "Precibo " & gcCodInm & "/" & strPeriodo
            '.WindowParentHandle = FrmAdmin.hWnd
            .TituloVentana = "Precibo " & gcCodInm & "/" & strPeriodo
        End If
        'ErrorLocal = .PrintReport
        .Imprimir
    End With
'PreRecibo:
'    If ErrorLocal <> 0 Then
'        MsgBox ctlReport.LastErrorString, vbCritical, "Error. " & ctlReport.LastErrorNumber
'        Call rtnBitacora("Print Prerecibo /Err." & ctlReport.LastErrorNumber & " " & ctlReport.LastErrorString)
'    End If
    End Sub

    '---------------------------------------------------------------------------------------------
    '   Rutina:     Printer_GNC
    '
    '---------------------------------------------------------------------------------------------
    Public Sub Printer_GNC(strPeriodo$, Periodo, Salida As crSalida)
    'variables locales
    Dim strGNC As String
    Dim archivo As String
    Dim rpReporte As ctlReport
    '----------------------------
    archivo = gcPath & gcUbica & "Reportes\NC" & Left(strPeriodo, 3) & Right(strPeriodo, 2) _
    & ".rpt"
    '
    Set rpReporte = New ctlReport
    With rpReporte  'Reporte Gasto No común
        .Reporte = archivo
        .Formulas(0) = "Condominio='" & gcCodInm & " - " & gcNomInm & "'"
        .Formulas(1) = "Periodo='" & strPeriodo & "'"
        If Dir(archivo) = "" Then
            
            'Crea una consulta Gastos No Comunes
            strGNC = "SELECT * FROM GastoNoComun WHERE Periodo=#" & Periodo & _
            "# ORDER BY CodApto, CodGasto;"
            Call rtnGenerator(mcDatos, strGNC, "qdfGNC")
            .Reporte = gcReport + "fact_nc.rpt"
            .OrigenDatos(0) = mcDatos
            .Salida = crArchivoDisco
            .ArchivoSalida = archivo
            .Imprimir
        End If
        .Salida = Salida
        If .Salida = crPantalla Then
            .TituloVentana = "Gastos No comunes " & gcCodInm & "/" & strPeriodo
        End If
        .Imprimir
    End With
    Set rpReporte = Nothing
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    '   Rutina:     Printer_Facturacion_Mensual
    '
    '   Este reporte contiene el detalle de la facturacion del mes
    '   detallado por apartamento, gasto comun, gasto no comun,
    '   otros gastos, total a cancelar
    '---------------------------------------------------------------------------------------------
    Public Sub Printer_Facturacion_Mensual(strPeriodo$, Salida As crSalida)
    'Variables locales
    Dim Archivo1$, Archivo2$, Ruta$, Nombre$
    Dim rpReporte As ctlReport, Destino%
    Dim rstFondo As New ADODB.Recordset
    Dim datPeriodo As Date
    '
    Ruta = gcPath & gcUbica & "Reportes\"
    Nombre = Left(strPeriodo, 3) & Right(strPeriodo, 2) & ".rpt"
    Archivo2 = Ruta & "F" & Nombre
    
    '
    Set rpReporte = New ctlReport
    With rpReporte
        'Reporte de Facturación Mensual (2 copias)
        .Formulas(0) = "Condominio='" & gcCodInm & " - " & gcNomInm & "'"
        .Formulas(1) = "Periodo='" & strPeriodo & "'"
        
        If strLlamada = "R" Then
        
            Archivo1 = Ruta & "R" & Nombre
            .Reporte = Archivo2
            .Formulas(2) = "Titulo='Reverso Facturación'"
            .Salida = crArchivoDisco
            .ArchivoSalida = Archivo1
            If Not Dir(Archivo1) = "" Then Kill Archivo1
            .TituloVentana = "Reverso de Facturuación " & gcCodInm
            .Imprimir
            
        Else
            .Reporte = Archivo2
            
            If Dir(Archivo2) = "" Then  'si no ha sido creado el reporte
                '
                datPeriodo = "01/" & strPeriodo
                datPeriodo = Format(datPeriodo, "mm/dd/yyyy")
                .Reporte = gcReport + "fact_mes.rpt"
                .OrigenDatos(0) = mcDatos
                '//Selecciona el movimiento del fondo durante el periodo facturado
                rstFondo.Open "SELECT * FROM Pago_Inf WHERE CtaFondo='" & gcCodFondo & "' AND P" _
                & "eriodo=#" & datPeriodo & "#;", cnnOLEDB + mcDatos, adOpenStatic, _
                adLockReadOnly, adCmdText
                '
                If Not rstFondo.EOF Or Not rstFondo.BOF Then
                
                    rstFondo.MoveFirst
                    .Formulas(2) = "FondoA=" & Replace(rstFondo!sa, ",", ".") 'Saldo anterior
                    .Formulas(3) = "FondoF=" & Replace(rstFondo!cr, ",", ".") 'Crédito durante el mes
                    .Formulas(4) = "Debitos=" & Replace((rstFondo!DB * -1), ",", ".") 'Débitos
                    '.Formulas(5) = "CheqDev=0"  'Total Cheques Devueltos
                    
                End If
                rstFondo.Close  '//cierra el ADODB.Recordset
                rstFondo.Open "SELECT Sum(Monto) FROM GastoNoComun WHERE CodGasto=(SELECT Cod" _
                & "ChDev FROM Inmueble IN '" & gcPath & "\sac.mdb' WHERE CodInm='" & gcCodInm _
                & "') and Periodo=#" & datPeriodo & "#;", cnnOLEDB + mcDatos, adOpenStatic, _
                adLockOptimistic
                If Not IsNull(rstFondo.Fields(0)) Then
                    .Formulas(5) = "CheqDev=" & Replace(rstFondo.Fields(0), ",", ".")
                Else
                    .Formulas(5) = "CheqDev=0"
                End If
                rstFondo.Close
                Set rstFondo = Nothing  '/descarga el ADODB.Recordset de memoria
                '
                .Salida = crArchivoDisco
                .ArchivoSalida = Archivo2
                .Imprimir
                '
            End If
            '
            .Salida = Salida
            .TituloVentana = "Facturación Mensual " & gcCodInm & "/" & strPeriodo
            .Imprimir
        End If
        Set rpReporte = Nothing
        '
       
        '
    End With
    '
    End Sub
    
    '---------------------------------------------------------------------------------------------
    '   Rutina:     Printer_Control_Facturación
    '
    '---------------------------------------------------------------------------------------------
    Public Sub Printer_Control_Facturacion(strPeriodo$, Periodo, Salida As crSalida)
    '
    'variables locales
    Dim strSQL$, rpReporte As ctlReport, m%, E%
    Dim strDir As String, rstPeriodo As New ADODB.Recordset
    '
    strPeriodo = UCase(strPeriodo)
    Set rpReporte = New ctlReport
    
    With rpReporte
        'Recibo Control de Facturacion
        strDir = gcPath & gcUbica & "Reportes\" & "CF" & Left(strPeriodo, 3) & Right(strPeriodo, 2) _
        & ".rpt"
        If Dir(strDir) = "" Then    'si no ha sido creado el reporte
            '
            .Reporte = gcReport & "Fact_control.rpt"
            .OrigenDatos(0) = mcDatos
            .Formulas(0) = "Inmueble='" & gcCodInm & " - " & gcNomInm & "'"
            .Formulas(1) = "Periodo='" & strPeriodo & "'"
            .Formulas(2) = "CodigoCondominio='" & gcCodInm & "'"
            .Formulas(3) = "FondoReserva=" & Replace(Total_Fondo(gcCodInm, Periodo), ",", ".")
            .Formulas(4) = "DI=" & Replace(Total_Deuda(gcCodInm), ",", ".")
            '
            If Detalle_Fondo(gcCodInm) Then
                strSQL = "SELECT Pago_Inf.*,TGastos.Titulo FROM Pago_INF INNER JOIN TGastos ON " _
                & "Pago_Inf.CtaFondo=Tgastos.CodGasto WHERE Periodo=#" & Periodo & "# AND CtaFond" _
                & "o IN (SELECT CodGasto FROM Fact_FondoInm IN '" & gcPath & "\sac.mdb' WHERE C" _
                & "odInm='" & gcCodInm & "')"
               
               rstPeriodo.Open strSQL, cnnOLEDB + mcDatos, adOpenKeyset, adLockReadOnly, adCmdText
            '
                If Not rstPeriodo.EOF Or Not rstPeriodo.BOF Then
                    rstPeriodo.MoveFirst
                    m = 5: E = 0
                    Do
                        E = E + 1
                        .Formulas(m) = "Cta" & E & "='" & rstPeriodo!Titulo & "'": m = m + 1
                        .Formulas(m) = "SA" & E & "=" & Replace(rstPeriodo!sa, ",", "."): m = m + 1
                        .Formulas(m) = "CR" & E & "=" & Replace(rstPeriodo!cr, ",", "."): m = m + 1
                        .Formulas(m) = "DB" & E & "=" & Replace(rstPeriodo!DB, ",", "."): m = m + 1
                        rstPeriodo.MoveNext
                    Loop Until rstPeriodo.EOF Or E = 4
                    '
                End If
                rstPeriodo.Close
                Set rstPeriodo = Nothing
            End If
            '.Formulas(5) = "Nota='" & Obtener_Nota(gcUbica) & "'"
            'Genera la consulta para el recibo de control de facturacion
            strSQL = "SELECT Propietarios.Codigo, Propietarios.Nombre, AsignaGasto.CodGasto, As" _
            & "ignaGasto.Cargado, AsignaGasto.Descripcion, Sum(AsignaGasto.Monto) AS Total, CSt" _
            & "r(Time()) AS Hora From AsignaGasto, Propietarios WHERE (((AsignaGasto.Comun)=Tru" _
            & "e) AND ((Propietarios.Codigo) Like 'U*') AND ((AsignaGasto.Cargado)=#" & Periodo _
            & "#)) GROUP BY Propietarios.Codigo, Propietarios.Nombre, AsignaGasto.CodGasto, Asi" _
            & "gnaGasto.Cargado, AsignaGasto.Descripcion ORDER BY AsignaGasto.CodGasto;"
            
            Call rtnGenerator(mcDatos, strSQL, "qdfCF")
            .Salida = crArchivoDisco
            .ArchivoSalida = strDir
            .Imprimir
            .Salida = Salida
            .TituloVentana = "Control de Facturación " & gcCodInm & "/" & strPeriodo
            .Imprimir
            
        Else
            .Reporte = strDir
            .Salida = Salida
            .TituloVentana = "Control de Facturación " & gcCodInm & "/" & strPeriodo
            .Imprimir
        End If
        
    End With
    Set rpReporte = Nothing
    End Sub

    
    '---------------------------------------------------------------------------------------------
    '   Rutina:     Printer_pago
    '
    '   Entradas:   # factura, Total a Pagar, Carpeta, Codigo de inmueble, Nombre
    '               Inmueble, # Recibo, Saldo Fondo de Reserva, Control Report, intECG
    '
    '   Entradas Opcionales:    Salida, Forma pago1, forma pago2, forma pago3
    '
    '   Imprime reportes de cancelación de gastos
    '---------------------------------------------------------------------------------------------
    Public Sub Printer_Pago(strNF$, curTotal@, Carpeta$, CodInm$, Inmueble$, NRecibo$, _
    Fondo As Boolean, Optional intECG%, Optional Destino As crSalida, Optional FP1$, _
    Optional Fp2$, Optional FP3$, Optional BsF As Boolean)
    '--------------------------------------------------
    '   valores de la variable intECG
    '   0 = Pago por Taquilla
    '   1 = Emision del Recibo de Pago (sin cancelar)
    '   2 = Reimpresion
    '--------------------------------------------------
    'On Error GoTo LocalErr
    'variables locales
    Dim rstPeriodo As ADODB.Recordset
    Dim cnnPeriodo As ADODB.Connection
    Dim errLocal&, strSQL$, strPer$, strFP1$, strFP2$, strFP3$
    Dim E%, m%, str5$, str6$, QdfName$, Temp%, Intento%
    Dim Mes As Date
    Dim crReport As CRAXDRT.Report
    Dim crApp As CRAXDRT.Application
    Dim crSubReportObj As CRAXDRT.Report
    Dim rstDP As ADODB.Recordset
    '
    Set crApp = New CRAXDRT.Application
    Set rstPeriodo = New ADODB.Recordset
    Set cnnPeriodo = New ADODB.Connection
    '-----------------------------------------------------------------------------------
    ' estos valores se setean en el caso de que algun parámetro requerido venga vacio
    '-----------------------------------------------------------------------------------
    If CodInm = "" Then
        CodInm = "2" & Mid(strNF, 5, 3)
        Carpeta = "\" & CodInm & "\"
    End If
    If Inmueble = "" Then
        Set rstDP = ModGeneral.ejecutar_procedure("procBuscaCaja", CodInm)
        If Not (rstDP.EOF And rstDP.BOF) Then Inmueble = rstDP("nombre")
        rstDP.Close
    End If
    '-----------------------------------------------------------------------------------
    '
    'Busca si el recibo esta en poder del cobrador para  no reimprimirlo
    strSQL = "SELECT * FROM Recibos_Emision WHERE Nfact='" & strNF & "';"
    
    Set rstPeriodo = cnnConexion.Execute(strSQL)
    '
    If Not rstPeriodo.EOF And Not rstPeriodo.BOF Then   'si está
    
        SetTimer hWnd, NV_CLOSEMSGBOX, 2000&, AddressOf TimerProc
        Call MessageBox(hWnd, "Este Recibo #" & strNF & " ya fué impreso", _
        App.ProductName, vbInformation)
    
        Call rtnBitacora("Printer Canc. Gastos: #" & strNF & " Impreso por " & _
        rstPeriodo!Impresopor & " " & rstPeriodo!fecha)
        'elimina el registro de la tabla
         cnnConexion.Execute "DELETE * FROM Recibos_Emision WHERE Nfact='" & strNF & "'"
        
        rstPeriodo.Close
        Set rstPeriodo = Nothing
        Exit Sub    'sale de la rutina sin imprimir
                    'el recibo
    End If
    rstPeriodo.Close
    strSQL = "SELECT Periodo FROM Factura WHERE Fact='" & strNF & "';"
    cnnPeriodo.Open cnnOLEDB & gcPath & Carpeta & "inm.mdb"
    rstPeriodo.Open strSQL, cnnPeriodo, adOpenStatic, adLockReadOnly
    strPer = Format(rstPeriodo!Periodo, "mm/dd/yyyy")
    Mes = rstPeriodo!Periodo
    rstPeriodo.Close
    If Config_Query(strPer, cnnPeriodo.Properties("Data Source")) Then Exit Sub
    '
    If intECG = 1 Then  'emision recibo de pago (sin cancelar)
    
        strSQL = "SELECT DF.Codigo, DF.Fact, DF.CodGasto, DF.Detalle, DF.CodGasto, Sum(DF.Monto) " _
        & "AS SumaDeMonto, DF.Hora, Propietarios.Nombre, Propietarios.Alicuota, Propietarios.Deuda, " _
        & "Propietarios.Recibos, DF.Periodo, DF.Fecha, AG.Total,Date() as freg,'" & gcUsuario & "' as usuario" & " FROM " _
        & "Propietarios INNER JOIN ((DetFact AS DF LEFT JOIN AG ON (DF.CodGasto = AG." & "CodGa" _
        & "sto) AND (DF.Detalle = AG.Descripcion)) INNER JOIN Factura ON DF.Fact = Factura.FACT) ON " _
        & "Propietarios.Codigo = DF.Codigo GROUP BY DF.Codigo, DF.Fact, DF.CodGasto, DF.Detalle, " _
        & "DF.CodGasto, DF.Hora, Propietarios.Nombre, Propietarios.Alicuota, Propietarios.Deuda, " _
        & "Propietarios.Recibos, DF.Periodo, DF.Fecha, AG.Total, Factura.freg, Factura.usuario " _
        & "HAVING (((DF.Fact)='" & strNF & "'));"
    Else
        strSQL = "SELECT DF.Codigo, DF.Fact, DF.CodGasto, DF.Detalle, DF.CodGasto, Sum(DF.Monto) " _
        & "AS SumaDeMonto, DF.Hora, Propietarios.Nombre, Propietarios.Alicuota, Propietarios.Deuda, " _
        & "Propietarios.Recibos, DF.Periodo, DF.Fecha, AG.Total, Factura.freg, Factura.usuario FROM " _
        & "Propietarios INNER JOIN ((DetFact AS DF LEFT JOIN AG ON (DF.CodGasto = AG.CodGasto) " _
        & "AND (DF.Detalle = AG.Descripcion)) INNER JOIN Factura ON DF.Fact = Factura.FACT) ON " _
        & "Propietarios.Codigo = DF.Codigo GROUP BY DF.Codigo, DF.Fact, DF.CodGasto, DF.Detalle, " _
        & "DF.CodGasto, DF.Hora, Propietarios.Nombre, Propietarios.Alicuota, Propietarios.Deuda, " _
        & "Propietarios.Recibos, DF.Periodo, DF.Fecha, AG.Total, Factura.freg, Factura.usuario " _
        & "HAVING (((DF.Fact)='" & strNF & "'));"
    End If
    '
    Call rtnGenerator(gcPath & Carpeta & "Inm.mdb", strSQL, "qdfPago")
    DoEvents
    '
    Set crReport = New CRAXDRT.Report

1   With crReport
        '
        If Nombre_Reporte(CodInm) Then
            'MsgBox "Introduzca el papel con logo para emitir el recibo de pago", _
            vbInformation, App.ProductName
            Set crReport = crApp.OpenReport(gcReport & "fact_pago_sac.rpt", 1)
        Else
            Set crReport = crApp.OpenReport(gcReport & "fact_pago.rpt", 1)
        End If
        
        crReport.FormulaFields.GetItemByName("CodigoCondominio").Text = "'" & CodInm & "'"
        crReport.FormulaFields.GetItemByName("Inmueble").Text = "'" & Inmueble & "'"
        If Intento = 0 Then
            crReport.FormulaFields.GetItemByName("Periodo").Text = _
            "'" & UCase(Format(Format(CDate(strPer), "mm/dd/yyyy"), "mmm - yyyy")) & "'"
        End If
        crReport.FormulaFields.GetItemByName("Total").Text = Replace(curTotal, ",", ".")
        crReport.FormulaFields.GetItemByName("NRecibo").Text = "'" & strNF & "'"
        E = 0
        '
        If Detalle_Fondo(CodInm) Then
            '---------------------------------------------------------------------------------
            'imprime el reporte de pagos el movimiento de las cuentas del fondo si está activo
            '---------------------------------------------------------------------------------
            '
            strSQL = "SELECT Pago_Inf.*,TGastos.Titulo FROM Pago_INF INNER JOIN TGastos ON " _
            & "Pago_Inf.CtaFondo=Tgastos.CodGasto WHERE Periodo=#" & strPer & "# AND CtaFond" _
            & "o IN (SELECT CodGasto FROM Fact_FondoInm IN '" & gcPath & "\sac.mdb' WHERE C" _
            & "odInm='" & CodInm & "')"
            rstPeriodo.Open strSQL, cnnPeriodo, adOpenStatic, adLockReadOnly, adCmdText

            If Not rstPeriodo.EOF Or Not rstPeriodo.BOF Then
                rstPeriodo.MoveFirst
                Do
                    E = E + 1
                    m = 4 * E
                    crReport.FormulaFields.GetItemByName("CTA" & E).Text = "'" & rstPeriodo!Titulo & "'"
                    crReport.FormulaFields.GetItemByName("SA" & E).Text = Replace(rstPeriodo!sa, ",", ".")
                    crReport.FormulaFields.GetItemByName("CR" & E).Text = Replace(rstPeriodo!cr, ",", ".")
                    crReport.FormulaFields.GetItemByName("DB" & E).Text = Replace(rstPeriodo!DB, ",", ".")
                    rstPeriodo.MoveNext
                Loop Until rstPeriodo.EOF
                
            End If
            rstPeriodo.Close
            '----------
        End If
        m = m + 4
        If intECG = 2 Then Temp = 2: intECG = 0
        
        If intECG = 0 Then  'si es la emisión de la cancelación de gastos
            intECG = Temp
            
                '
            If Fondo Then   'determina si es una emisión sin cancelar
                
                If intECG = 2 Then  'Reimpresión del recibo
                    
                    crReport.FormulaFields.GetItemByName("FP1").Text = "'" & IIf(FP1 = " -  - ", "", FP1) & "'"
                    crReport.FormulaFields.GetItemByName("FP2").Text = "'" & IIf(Fp2 = " -  - ", "", Fp2) & "'"
                    crReport.FormulaFields.GetItemByName("FP3").Text = "'" & IIf(FP3 = " -  - ", "", FP3) & "'"
                    strPer = ""
                    
                Else
                    Set rstDP = ModGeneral.ejecutar_procedure("procDetallePago", NRecibo)
                    
                    If BsF Then
                    
'                        strPer = IIf(FrmMovCajaBs.Txt(12) > 0 Or FrmMovCajaBs.cmb(0).Text = "EFECTIVO", "EFECTIVO", "")
'
'                        For E = 5 To 7
'                            If FrmMovCajaBs.cmb(E) <> "" And FrmMovCajaBs.Txt(E + 1) <> "" And FrmMovCajaBs.Txt(E - 4) <> "" Then
'                                crReport.FormulaFields.GetItemByName("FP" & E - 4).Text = "'" & FrmMovCajaBs.cmb(E) & " - " _
'                                & FrmMovCajaBs.Txt(E + 1) & " - " & FrmMovCajaBs.cmb(E - 3) & "'"
'                            End If
'                        Next
                        If Not (rstDP.EOF And rstDP.BOF) Then
                            rstDP.MoveFirst
                            E = 1
                            Do
                                If UCase(rstDP("FPago")) = "EFECTIVO" Then
                                    strPer = "EFECTIVO"
                                Else
                                    crReport.FormulaFields.GetItemByName("FP" & E).Text = "'" & rstDP("FPAGO") _
                                    & " - " & rstDP("ndoc") & " - " & rstDP("Banco") & " - " & _
                                    Format(rstDP("FechaDoc"), "dd/mm/yy") & "'"
                                    E = E + 1
                                End If
                                rstDP.MoveNext
                            Loop Until rstDP.EOF
                            rstDP.Close
                        End If
                        
                    Else
'                        strPer = IIf(FrmMovCaja.Txt(12) > 0 Or FrmMovCaja.cmb(0).Text = "EFECTIVO", "EFECTIVO", "")
'
'                        For E = 5 To 7
'                            If FrmMovCaja.cmb(E) <> "" And FrmMovCaja.Txt(E + 1) <> "" And FrmMovCaja.Txt(E - 4) <> "" Then
'                                crReport.FormulaFields.GetItemByName("FP" & E - 4).Text = "'" & FrmMovCaja.cmb(E) & " - " _
'                                & FrmMovCaja.Txt(E + 1) & " - " & FrmMovCaja.cmb(E - 3) & "'"
'                            End If
'                        Next
                        If Not (rstDP.EOF And rstDP.BOF) Then
                            rstDP.MoveFirst
                            E = 1
                            Do
                                If UCase(rstDP("FPago")) = "EFECTIVO" Then
                                    strPer = "EFECTIVO"
                                Else
                                    crReport.FormulaFields.GetItemByName("FP" & E).Text = "'" & rstDP("FPAGO") _
                                    & " - " & rstDP("ndoc") & " - " & rstDP("Banco") & " - " & _
                                    Format(rstDP("FechaDoc"), "dd/mm/yy") & "'"
                                    E = E + 1
                                End If
                                rstDP.MoveNext
                            Loop Until rstDP.EOF
                            rstDP.Close
                        End If
                    
                    End If
                    
                End If
                '
            Else
                strPer = "EFECTIVO"
            End If
            '
        Else
            strPer = ""
        End If
        cnnPeriodo.Close
        crReport.FormulaFields.GetItemByName("FP0").Text = "'" & IIf(IsDate(strPer), "", strPer) & "'"
        crReport.Database.Tables(1).Location = gcPath & Carpeta & "Inm.mdb"
        crReport.Database.Tables(2).Location = gcPath & "\tablas.mdb"
        '
        Set crSubReportObj = New CRAXDRT.Report
        Set crSubReportObj = crReport.OpenSubreport("fact_graf.rpt")
        
        crSubReportObj.Database.Tables(1).Location = gcPath & Carpeta & "Inm.mdb"
        If Destino = crImpresora Then
            crReport.DisplayProgressDialog = False
            crReport.PrintOut False
        ElseIf Destino = crArchivoDisco Then
            'guardamos la copia del archivo en formato pdf
            crReport.ExportOptions.DestinationType = crEDTDiskFile
            crReport.ExportOptions.FormatType = crEFTPortableDocFormat
            crReport.ExportOptions.DiskFileName = Environ("temp") & "\" & strNF & ".pdf"
            crReport.Export False
        Else
            
            Set frmLocal = New frmView
            frmLocal.crView.ReportSource = crReport
            frmLocal.crView.ViewReport
            If Screen.Width / Screen.TwipsPerPixelX = 1024 Then
                frmLocal.crView.Zoom 120
            Else
                frmLocal.crView.Zoom 1
            End If
            frmLocal.Caption = "Recibo Cancelación #" & strNF
            frmLocal.Show
        End If
        Call rtnBitacora("Printer Canc. Gastos: #" & strNF)
        
    End With
    '
LocalErr:
    If Err <> 0 Then
        If Intento >= 3 Then
            MsgBox LoadResString(539), vbExclamation, App.ProductName
        Else
            Intento = Intento + 1
            Err.Clear: GoTo 1
        End If
        '
    Else
        'marca el recibo como impreso
        If intECG = 2 Then   'si esta imprimiento recibos enviar
            cnnConexion.Execute "UPDATE Recibos_Enviar SET Print=True WHERE IDRecibo='" & _
            NRecibo & "' AND Fact='" & strNF & "';"
        End If
    End If
    'elimina de la memoria los objetos
    Set crReport = Nothing
    Set crApp = Nothing
    Set crSubReportObj = Nothing
    Set rstPeriodo = Nothing
    Set cnnPeriodo = Nothing
    '
    End Sub


'---------------------------------------------------------------------------------------------
    '   Rutina rtnAC
    '
    '   Entradas: Control Crytal Report, Periodo Formatio "mmm/yyyy",Periodo
    '               Formato fecha
    '
    '   Devuelve los avisos de cobro de un determinado condominio
    '---------------------------------------------------------------------------------------------
    Public Sub rtnAC(strPeriodo$, datPer As Date, Optional Salida As crSalida)     '
    '
    'variables locales
    Dim errLocal As Integer
    Dim strDir As String
    Dim Destino As Integer
    Dim strNombre As String
    Dim ProEmail As New ADODB.Recordset
    Dim strCB As String
    Static Inm
    Static Per
    Dim mReport As CRAXDRT.Report
    Dim mApp As CRAXDRT.Application
    '
    On Error GoTo errRtnAC
           '
    strDir = gcPath & gcUbica & "reportes\AC" & UCase(Left(strPeriodo, 3)) & _
    Right(strPeriodo, 2) & ".rpt"
    '
    Screen.MousePointer = vbHourglass
    'Call rtnBitacora("Imprimiendo Pre-Recibos...Inm.:" & gcCodInm)
    'ctlReport.Destination = Salida
    '
    If Dir$(strDir) = "" Then    'si es la primera vez que se pide este reporte
                                            'almacena una copia
        If Inm <> gcCodInm Or datPer <> Per Then If Config_Query(datPer, mcDatos) Then Exit Sub
        '
'        With ctlReport
            
'            If Inm <> gcCodInm Or datPer <> Per Then
'                'Call clear_Crystal(ctlReport)
'                .Reset
'                .ProgressDialog = False
'                .ReportFileName = gcReport + IIf(Nombre_Reporte(gcCodInm), "fact_aviso_sac.rpt", _
'                "fact_aviso.rpt")
'                .DataFiles(0) = mcDatos
'                '
'                .Formulas(0) = "CodigoCondominio='" & gcCodInm & "'"
'                .Formulas(2) = "FondoReserva =" & Total_Fondo(gcCodInm, datPer)
'                .Formulas(4) = "Inmueble='" & gcNomInm & "'"
'                .Formulas(3) = "Nota='" & Obtener_Nota(gcUbica) & "'"
'                .Formulas(6) = "MesesMora=" & gnMesesMora
'                .Formulas(9) = "pInteres='" & gnPorIntMora & "'"
'                '
'                If gnCta = CUENTA_POTE Then
'                    'strNombre = sysEmpresa
'                    strNombre = ""
'                Else
'                    strNombre = gcNomInm
'                End If
'                '
'
'                strCB = complemento
'                .Formulas(8) = "Pie_Pagina='" & IIf(strNombre = "", "", "FAVOR EMITIR CHEQUE NO ENDOSABLE A NOMBRE DE: " _
'                & strNombre) & IIf(strCB = "" Or strNombre = "", "", " 0 ") & strCB & "'"
'                .Formulas(7) = "Periodo='" & UCase(strPeriodo) & "'"
'                '-------------------------------------------------------------------------------
'                'activar estas línes para el nuevo aviso de cobro
'                '-------------------------------------------------------------------------------
'                .SubreportToChange = "fact_graf.rpt"
'                .DataFiles(0) = mcDatos
'                .DiscardSavedData = True
'                .SQLQuery = "SELECT * FROM FACTURA WHERE Periodo >=#" & _
'                DateAdd("m", -6, Format(datPer, "mm/dd/yy")) & "#"
'                .SubreportToChange = "fact_sub2.rpt"
'                .DataFiles(0) = gcPath & "\sac.mdb"
'                .DiscardSavedData = True
'                .SQLQuery = "SELECT * FROM MovimientoCaja WHERE CodigoInmuebleMovimientoCaja='" & _
'                gcCodInm & "'"
'                '------------------------------------------------------------------------------
'                .Destination = crptToFile
'                .PrintFileName = strDir
'                .PrintFileType = crptCrystal
'                .DiscardSavedData = True
'                errLocal = .PrintReport
'                .Destination = Salida
'                'SOLO IMPRIMO LOS PROPIETARIOS MARCADOS PARA RECIBIR EL AVISO
'                If Format(datPer, "mm/dd/yy") > CDate("01/07/2005") Then
'                    .SubreportToChange = ""
'                    .ReplaceSelectionFormula "{AC.bAviso}=0"
'                End If
'                If .Destination = crptToWindow Then
'                    .WindowParentHandle = FrmAdmin.hWnd
'                    .WindowTitle = "Avisos de Cobro " & gcCodInm & "/" & UCase(Format(datPer, "MMM-YYYY"))
'                    .WindowShowCloseBtn = True
'                    .WindowShowGroupTree = True
'                    .WindowBorderStyle = crptNoBorder
'                    .WindowState = crptMaximized
'                End If
'                errLocal = .PrintReport
'
'            Else
'                'SOLO IMPRIMO LOS PROPIETARIOS MARCADOS PARA RECIBIR EL AVISO
'
'                If Format(datPer, "mm/dd/yy") > CDate(Z) Then
'                    .SubreportToChange = ""
'                    .ReplaceSelectionFormula "{AC.bAviso}=0"
'                End If
'                .Destination = Destino
'                .WindowTitle = "Avisos de Cobro " & gcCodInm
'                errLocal = .PrintReport
'
'            End If
'
'        End With

        Set mApp = New CRAXDRT.Application
        Set mReport = mApp.OpenReport(gcReport + IIf(Nombre_Reporte(gcCodInm), "fact_aviso_sac.rpt", _
        "fact_aviso.rpt"), 1)
        mReport.DisplayProgressDialog = False
        mReport.DiscardSavedData
        mReport.Database.Tables(1).Location = mcDatos
        mReport.FormulaFields.GetItemByName("CodigoCondominio").Text = "'" & gcCodInm & "'"
        mReport.FormulaFields.GetItemByName("FondoReserva").Text = Total_Fondo(gcCodInm, datPer)
        mReport.FormulaFields.GetItemByName("Inmueble").Text = "'" & gcNomInm & "'"
        mReport.FormulaFields.GetItemByName("Nota").Text = "'" & Obtener_Nota(gcUbica) & "'"
        mReport.FormulaFields.GetItemByName("MesesMora").Text = gnMesesMora
        mReport.FormulaFields.GetItemByName("pInteres").Text = gnPorIntMora
        If gnCta = CUENTA_POTE Then
            'strNombre = sysEmpresa
            strNombre = ""
        Else
            strNombre = gcNomInm
        End If
        strCB = complemento
        mReport.FormulaFields.GetItemByName("Pie_Pagina").Text = "'" & IIf(strNombre = "", "", "FAVOR EMITIR CHEQUE NO ENDOSABLE A NOMBRE DE: " _
        & strNombre) & IIf(strCB = "" Or strNombre = "", "", " 0 ") & strCB & "'"
        mReport.FormulaFields.GetItemByName("Periodo").Text = "'" & UCase(strPeriodo) & "'"
        'subreporte
        Dim crSubReportObj As CRAXDRT.Report
        Dim crSubReportObj1 As CRAXDRT.Report

        Set crSubReportObj1 = mReport.OpenSubreport("fact_sub2.rpt")
        crSubReportObj1.Database.Tables(1).Location = gcPath & "\sac.mdb"
        crSubReportObj1.DiscardSavedData
        '
        Set crSubReportObj = mReport.OpenSubreport("fact_graf.rpt")
        crSubReportObj.Database.Tables(1).Location = mcDatos
        crSubReportObj.DiscardSavedData
        '
        'almacena una copia fiel del reporte
        mReport.ExportOptions.DestinationType = crEDTDiskFile
        mReport.ExportOptions.FormatType = crEFTCrystalReport
        mReport.ExportOptions.DiskFileName = strDir
        mReport.Export False
        'SOLO IMPRIMO LOS PROPIETARIOS MARCADOS PARA RECIBIR EL AVISO
        '
        If Format(datPer, "mm/dd/yy") > CDate("01/07/2005") Then
            mReport.RecordSelectionFormula = "{AC.bAviso}=0"
        End If
        
        If Salida = crptToWindow Then
            Dim mFrm As New frmView
            mFrm.Caption = "Aviso de Cobros " & gcCodInm
            mFrm.crView.ReportSource = mReport
            mFrm.crView.ViewReport
            While mFrm.crView.IsBusy
                DoEvents
            Wend
            mFrm.crView.DisplayGroupTree = True
            mFrm.crView.EnableGroupTree = True
            mFrm.crView.Zoom (IIf(mvarZoom = 0, 1, mvarZoom))
            mFrm.Show
            mFrm.Visible = True

        Else
            mReport.PrintOut False
        End If
    Else    'si se pidio anteriormente este reporte busca la copia almacenada para presentarla
            ' al usuario
        
        Set Reporte = New ctlReport
        Reporte.Reporte = strDir
        Reporte.Salida = crPantalla
        If Format(datPer, "mm/dd/yy") > CDate("01/07/2005") Then
            Reporte.FormuladeSeleccion = "{AC.bAviso}=0"
        End If
        Reporte.TituloVentana = "Avisos de Cobro " & gcCodInm & "/" _
        & UCase(strPeriodo)
        Reporte.ArbolGrupo = True
        Reporte.Imprimir
        Set Reporte = Nothing
    End If
    '
    Screen.MousePointer = vbDefault
    'For errLocal = 0 To 100000  'RETARDO
    'Next errLocal
    
    Call rtnBitacora("Emisión de Recibos de Cobro Inm.:" & gcCodInm)
    'captura el error (si ocurre)
errRtnAC:
    If Err <> 0 Then MsgBox Err.Description, vbCritical, App.ProductName
    If errLocal <> 0 Then MsgBox ctlReport.LastErrorString, vbCritical, App.ProductName
    Inm = gcCodInm
    Per = datPer
    '
    End Sub


    '---------------------------------------------------------------------------------------------
    '   Rutina:     Printer_Report
    '
    '   Entradas:   Cadena de Selección de registros, Titulo y Subtiulo del
    '               Reporte, Tipo de reporte
    
    '   Esta rutina desensadena la impresion de un reporte de acuerdo al tipo
    '   de reporte seleccionado por el usuario.
    '---------------------------------------------------------------------------------------------
    Public Sub Printer_Report(strSQL As String, Titulo As String, SubTitulo As String, _
    Optional IVA As Boolean)
    'variables locales
    Dim rstCxP As New ADODB.Recordset
    Dim rstCargado As New ADODB.Recordset
    Dim curTotal As String
    Dim Izq As Long, K As Integer
    Dim Lin As Long, booTotalizado As Boolean
    Dim Lintemp As Long, Pag As Long
    Dim vecSubtotal(10, 3)  'Acumula el total por gasto
    Dim TotalGeneral As Currency
    Const m$ = "#,##0.00"
    '
    Pag = 1
    '   origen de los registros seleccionados
    rstCxP.Open strSQL, cnnConexion, adOpenKeyset, adLockReadOnly, adCmdText
    '
    If Not rstCxP.EOF And Not rstCxP.BOF Then
    
        Unload FrmRepFact
        Call RtnConfigUtility(True, App.ProductName, "Generando Reporte", _
        "Espere un momento por favor....")
        'aqui genera el encabezado del documento
        Call cabeza(Pag, Titulo, SubTitulo)
        rstCxP.MoveFirst
        Do  'hacer hasta fin de archivo
            'origen de los registros seleccionados
            If Lin > 14000 Then 'nueva página
                Printer.NewPage
                Pag = Pag + 1
                Call cabeza(Pag, Titulo, SubTitulo)
                Lin = Printer.CurrentY
            End If
            '
'            If Not IVA Then
                strSQL = "SELECT * FROM Cargado WHERE Ndoc='" & rstCxP!NDoc & "'"
'            Else
'                strSql = "SELECT * FROM Cargado WHERE Ndoc='" & rstCxP!NDoc & "' AND CodGasto IN " _
'                & "(SELECT CodIVA FROM Inmueble IN '" & gcPath & "\sac.mdb' WHERE CodInm='" _
'                & rstCxP("Cpp.CodInm") & "')"
'            End If
            rstCargado.Open strSQL, cnnOLEDB + gcPath + rstCxP!Ubica + "Inm.mdb", adOpenKeyset, _
            adLockOptimistic, adCmdText
            Call RtnProUtility("Escribiendo Inf. Doc.:" & rstCxP("Ndoc"), _
            6015 * (rstCxP.AbsolutePosition / rstCxP.RecordCount))
            
            If Not rstCargado.EOF Or Not rstCargado.BOF Then
            
                rstCargado.MoveFirst
                If IVA Then
                    If Not FrmAdmin.objRst.EOF And Not FrmAdmin.objRst.BOF Then
                        FrmAdmin.objRst.MoveFirst
                        FrmAdmin.objRst.Find "CodInm='" & rstCxP("Cpp.CodInm") & "'"
                        If Not FrmAdmin.objRst.EOF Then
                            rstCargado.Find "CodGasto ='" & FrmAdmin.objRst("CodIVA") & "'"
                            If rstCargado.EOF Then
                                rstCargado.Close
                                
                                GoTo 100
                            Else
                                rstCargado.MoveFirst
                            End If
                        End If
                    End If
                End If
                'genera un encabezado por inmueble
                Printer.FontName = "Comic Sans MS"
                'ForeColor = &H8000000F
                Lin = Printer.CurrentY
                Printer.Line (1200, Lin)-(Printer.ScaleWidth - 800, Lin + 215), &H8000000F, BF
                Printer.Line (1200, Lin)-(Printer.ScaleWidth - 800, Lin + 215), vbBlack, B
                '---------------
                Printer.CurrentX = 1000 + ((1200 - Printer.TextWidth(rstCxP("Inmueble.CodInm"))) / 2)
                Printer.CurrentY = Lin
                Printer.Print rstCxP("Inmueble.CodInm")  'código del inm
                '---------------
                Printer.CurrentY = Lin
                Printer.CurrentX = 2200 + ((1200 - Printer.TextWidth(rstCxP("NDoc"))) / 2)
                Printer.Print rstCxP("NDoc")    'número de documento
                '--------------
                Printer.CurrentY = Lin
                Printer.CurrentX = 3400
                Printer.Print Left(rstCxP("Detalle"), 63) 'descripción
                '-------------
                Printer.CurrentY = Lin
                curTotal = Format(rstCxP("Total"), m)
                Printer.CurrentX = 9500 + 1200 - Printer.TextWidth(curTotal)
                Printer.Print curTotal  'monto
                '-----------
                Lin = Printer.CurrentY
                Printer.FontName = "Arial"
                '
                TotalGeneral = TotalGeneral + rstCxP("Total")
                
                Do
                    'aqui se genera el listado por inmueble
                    Lintemp = linea_Detalle(rstCxP("Inmueble.CodInm"), rstCargado, Lin, IVA, rstCxP("FRecep"))
                    Lin = Lintemp
                    'sub totaliza por cada gasto
                    For K = 0 To 10
                        'busca la coincidencia
                        If vecSubtotal(K, 0) = rstCargado("CodGasto") Then
                            vecSubtotal(K, 2) = vecSubtotal(K, 2) + rstCargado("Monto")
                            booTotalizado = True
                            Exit For
                        End If
                    Next
                    If Not booTotalizado Then   ' si no fue sumado
                        For K = 0 To 10
                            'si no hay coincidencia registra el cargado en el vector
                            If vecSubtotal(K, 0) = "" Then
                                vecSubtotal(K, 0) = rstCargado("CodGasto")
                                vecSubtotal(K, 1) = Left(rstCargado("Detalle"), 30)
                                vecSubtotal(K, 2) = rstCargado("Monto")
                                Exit For
                            End If
                        Next
                    Else
                        booTotalizado = False
                    End If
                    rstCargado.MoveNext
                Loop Until rstCargado.EOF
                'Printer.FontName = "Comic Sans MS"
            End If
            'cierra el objeto ADODB.Recordset
            rstCargado.Close
            Printer.Print
100       rstCxP.MoveNext
            
        Loop Until rstCxP.EOF
        '-----------
        'imprime el pie de página y cierra el documento
        Printer.FontName = "Comic Sans MS"
        '
        For K = 0 To 10
            If vecSubtotal(K, 0) <> "" Then
                curTotal = Format(vecSubtotal(K, 2), m)
                Printer.CurrentX = 4000
                Lin = Printer.CurrentY
                Printer.Print "TOTAL " & vecSubtotal(K, 1) & "(" & vecSubtotal(K, 0) & ")"
                Printer.CurrentY = Lin
                Printer.CurrentX = 9700 - Printer.TextWidth(curTotal)
                Printer.Print curTotal
            End If
        Next

        Printer.FontSize = 12
        Printer.CurrentX = 1000
        Printer.Print "TOTAL GENERAL: " & Format(TotalGeneral, m)
        Printer.EndDoc
        '------------
        Unload FrmUtility
        '
    Else
        MsgBox "No existen registros que imprimir", vbInformation, App.ProductName
    End If
    'cierra el objeto ADODB.Recordset
    rstCxP.Close
    Set rstCxP = Nothing
    Set rstCargado = Nothing

    End Sub

    '------------------------------------------------------------------------------------
    '   Rutina: imprimir_cheque
    '
    '   Imprime los cheques recibidos por caja.
    '------------------------------------------------------------------------------------
    Public Sub imprimir_cheque(Monto As Currency, Beneficiario As String, Pote As Boolean)
    
    Dim rpReporte As ctlReport
    Dim strMonto As String, strNeto As String, strBeneficiario As String
    Dim cAletra As clsNum2Let
    Dim errLocal As Long
    
    'valida los datos mínimos requeridos para la imrpesión del cheque
    
    Set cAletra = New clsNum2Let
        
    cAletra.Moneda = "Bolívares"
    cAletra.Numero = Monto
    strMonto = cAletra.ALetra
    'configura la presentación del monto
    strMonto = String(10, "x") & strMonto
    If Len(strMonto) <= 98 Then
        strMonto = strMonto & " " & String(10, "x") & " " & String(Len(strMonto), "x")
    End If
    strNeto = Format(Monto, "XXX#,##0.00XXX")
    If Pote Then
        strBeneficiario = sysEmpresa
    Else
        strBeneficiario = Beneficiario
    End If
    Set rpReporte = New ctlReport
    With rpReporte
        .Reporte = gcReport & "chq_general.rpt"
        .Formulas(0) = "Beneficiario='" & strBeneficiario & "'"
        .Formulas(1) = "Neto='" & strNeto & "'"
        .Formulas(2) = "aletras='" & strMonto & "'"
        .Salida = crImpresora
        errLocal = .Imprimir
    End With
    If errLocal <> 0 Then
        Call rtnBitacora("Error al imprimir el cheque: " & Err.Description)
    End If
    'descarga objetos de memoria
    Set rpReporte = Nothing
    Set cAletra = Nothing
    End Sub


    '13/08/2002--------------------------------------------Imprime Recibos:
    '----------act.06/02/2005
    Public Sub RtnImpresion(IDRecibo$, strForSel$, strEtiqueta$, RutaInmueble$)
    '-------------------------------------------------------------------------------
    Dim rpReporte As ctlReport
    Dim errLocal As Long
    Set rpReporte = New ctlReport
    '
    With rpReporte
    '---
        .Reporte = gcReport & IIf(strEtiqueta = "INGRESOS VARIOS", "reciboiv.rpt", "recibo.rpt")
        .OrigenDatos(0) = gcPath + "\sac.mdb"
        .OrigenDatos(1) = RutaInmueble
        .Formulas(0) = "LblTipo='" & strEtiqueta & "'"
        .FormuladeSeleccion = "{QdfCaja.IDRecibo} = '" & IDRecibo & "'" & strForSel
        .Salida = crImpresora
        .TituloVentana = "Impresión ID: " & IDRecibo
        errLocal = .Imprimir(IIf(strEtiqueta = "HONORARIOS DE ABOGADO", 2, 1))
        Call rtnBitacora("Impresión " & strEtiqueta)
        If errLocal <> 0 Then Call rtnBitacora(Err.Description & "-" & Err)
    '---
    End With
    'descarga objeto de memoria
    Set rpReporte = Nothing
    '
    End Sub

    Public Sub emitir_solvencia(ParamArray sol())
    'arreglo con la información del propietario
    ' Elemento - Contenido
    ' 0 = Apto
    ' 1 = Cancelacion
    ' 2 = facturacion
    ' 3 = Inmueble
    ' 4 = Propietario
    ' 5 = Usuario
    '-------------------------------------------
    Dim rpReporte As ctlReport
    Dim I As Integer
    Dim errLocal As Long
    
    Set rpReporte = New ctlReport
    With rpReporte
        .Reporte = gcReport + "cxc_solvencia.rpt"
        .OrigenDatos(0) = gcPath & "\tablas.mdb"
        '.OrigenDatos(1) = mcDatos
        .TituloVentana = "Solvencia de Condominio"
        .Salida = crPantalla
        
        .Formulas(0) = "apto='" & sol(0) & "'"
        .Formulas(1) = "cancelacion='" & sol(1) & "'"
        .Formulas(2) = "facturacion='" & sol(2) & "'"
        .Formulas(3) = "inmueble='" & sol(3) & "'"
        .Formulas(4) = "propietario='" & sol(4) & "'"
        .Formulas(5) = "usuario='" & sol(5) & "'"
        .Formulas(6) = "codinm='" & sol(6) & "'"
        Call rtnBitacora("Imprimir Solvencia Inm:" & sol(3) & "/Apto.:" & sol(0))
        errLocal = .Imprimir
        If errLocal <> 0 Then
            Call rtnBitacora("Ocurrio un error al imprimir el reporte." & Err.Description)
        End If
    End With
    Set rpReporte = Nothing
    End Sub


    Public Sub ImprimirConstancia(Codigo As Long, conSueldo As Boolean, _
    Salida As crSalida, Empleado As String, Sueldo As Double, _
    Optional Seleccion As String)
    
    Dim rpReporte As ctlReport
    Dim errLocal As Long
    Dim nom_report As String
    Dim cAletra As clsNum2Let
    Dim SueldoaLetra As String
    
    Set cAletra = New clsNum2Let
    
    nom_report = IIf(conSueldo, "nom_constancia_con_sueldo.rpt", "nom_constancia_sin_sueldo.rpt")
    cAletra.Moneda = "Bs. "
    cAletra.Numero = Sueldo
    SueldoaLetra = cAletra.ALetra
    
    Set rpReporte = New ctlReport
    Call rtnBitacora("Imprimiendo Constancia de Trabajo.")
    With rpReporte
        .Reporte = gcReport + nom_report
        .OrigenDatos(0) = gcPath + "\sac.mdb"
        .OrigenDatos(1) = gcPath + "\sac.mdb"
        .OrigenDatos(2) = gcPath + "\sac.mdb"
        .Formulas(0) = "sueldoenletras='" & SueldoaLetra & "'"
        .FormuladeSeleccion = "{Emp.CodEmp} =" & Codigo
        .Salida = Salida
        .TituloVentana = "Constancia de Trabajo [" & Empleado & "]"
        errLocal = .Imprimir
        If errLocal <> 0 Then
            Call rtnBitacora("Ocurrio un error al imprimir el reporte." & Err.Description)
        End If
    End With
    
    End Sub
    
    Public Function imprimir_Carta_Aguinaldo(strPeriodo As String)
    Dim cReport As ctlReport
    Set cReport = New ctlReport
    cReport.Reporte = gcReport & "nom_cartaagui.rpt"
    cReport.OrigenDatos(0) = gcPath & "\sac.mdb"
    cReport.OrigenDatos(0) = gcPath & "\sac.mdb"
    cReport.Formulas(0) = "ano='" & strPeriodo & "'"
    cReport.Salida = crPantalla
    cReport.ArbolGrupo = True
    cReport.TituloVentana = "Cartas Aguinaldos"
    cReport.Imprimir
    Set cReport = Nothing
    
    End Function
    Public Function reversar_nomina(IDNomina As Long) As Boolean
    Dim strSQL As String
    
    Call RtnConfigUtility(True, "Reverso Nómina " & IDNomina, "Iniando proceso....", "Espere un momento por favor....")
    Call rtnBitacora("Iniciando Reverso nómina " & IDNomina)
    
    Dim rstNom As ADODB.Recordset
    Dim codNom(2) As String
    Dim pDFT As Currency
    
    Set rstNom = New ADODB.Recordset
    strSQL = "SELECT Inmueble.CodInm, Inmueble.Caja, Inmueble.CtaInm, Emp.CodEm" _
    & "p, Emp.Apellidos & ', ' & Emp.Nombres AS Name,Emp.Cedula,Emp.Cuenta,(Nom" _
    & "_Detalle.Sueldo/30) AS D,(d*Nom_Detalle.Dias_Trab)+(Nom_Detalle.Sueldo + (Nom_Detalle.Sue" _
    & "ldo * Emp.BonoNoc / 100)) /30 *  Nom_Detalle.Dias_libres +(Nom_Detalle.Sueldo + " _
    & "(Nom_Detalle.Sueldo * Nom_Detalle.Porc_BonoNoc / 100)) /30 *  Nom_Detalle.Dias_Feriados * '" _
    & pDFT & "' +Nom_Detalle.Bono_Noc+Nom_Detalle.Bono_Otros+Nom_Detalle.Otras_" _
    & "Asignaciones-((d*Nom_Detalle.Dias_NoTrab)+Nom_Detalle.SSO+Nom_Detalle.SP" _
    & "F+Nom_Detalle.LPH+Nom_Detalle.Otras_Deducciones+Nom_Detalle.Descuento) A" _
    & "S Neto, Inmueble.Nombre, Emp_Cargos.NombreCargo,Emp.CodGasto FROM Emp_Ca" _
    & "rgos INNER JOIN (Inmueble INNER JOIN (Emp INNER JOIN Nom_Detalle ON Emp." _
    & "CodEmp=Nom_Detalle.CodEmp) ON Inmueble.CodInm = Emp.CodInm) ON Emp_Cargo" _
    & "s.CodCargo=Emp.CodCargo Where Inmueble.Inactivo = False And Nom_Detalle." _
    & "IDNom = " & IDNomina & " ORDER BY Inmueble.Caja, Inmueble.CodInm;"
    '
    Call rtnGenerator(gcPath & "\SAC.MDB", strSQL, "qdfNomina_Bco")
    '
    strSQL = "SELECT Emp.CodInm, Nom_Detalle.CodEmp, Emp.Apellidos, Emp.Nombres" _
    & ",  Nom_Detalle.Sueldo, Nom_Detalle.Dias_Trab,( Nom_Detalle.Sueldo + ( Nom_Detalle.Sueldo *  Nom_Detalle.Porc_BonoNoc" _
    & " / 100)) /30 *  Nom_Detalle.Dias_libres as DL, ( Nom_Detalle.Sueldo + ( " _
    & "Nom_Detalle.Sueldo *  Nom_Detalle.Porc_BonoNoc / 100)) /30 *  Nom_Detalle.Dias_Feriados * '" & pDFT & _
    "'  as DF, Nom_Detalle.Bono_Noc, Nom_Detalle.Bono_Otros, Nom_Detalle.Otras_" _
    & "Asignaciones, Nom_Detalle.Dias_NoTrab, Nom_Detalle.SSO, Nom_Detalle.SPF," _
    & "Nom_Detalle.LPH, Nom_Detalle.Otras_Deducciones, Nom_Detalle.Descuento,Em" _
    & "p.CodGasto FROM Emp INNER JOIN Nom_Detalle ON Emp.CodEmp = Nom_Detalle.C" _
    & "odEmp WHERE Nom_Detalle.IDNom=" & IDNomina
    'impresión
    Call rtnGenerator(gcPath & "\sac.mdb", strSQL, "qdfNomina")
    With rstNom
        .Open "Nom_Calc", cnnConexion, adOpenKeyset, adLockOptimistic, adCmdTable
        If Not (.EOF And .BOF) Then
            If !CodDia_Feriado = 0 Or !codDia_Libre = 0 Or !codOtras_Asig = 0 Then
                reversar_nomina = MsgBox("Verifique los parámetros de la nómina. Consulte al administrador del sistema.", _
                vbInformation, App.ProductName)
                .Close
                Exit Function
            Else
                codNom(0) = !codDia_Libre
                codNom(1) = !CodDia_Feriado
                codNom(2) = !codOtras_Asig
                pDFT = !Dia_Feriado
                
            End If
        Else
            reversar_nomina = MsgBox("Verifique los parámetros de la nómina. Consulte al administrador del sistema.", _
            vbInformation, App.ProductName)
            .Close
            Exit Function
        End If
        .Close
    End With
    
    cnnConexion.BeginTrans
    
    Call rtnBitacora("Eliminando información de facturación.")
    
    strSQL = "SELECT Emp.CodGasto, Emp.CodInm, Sum((Nom_Detalle.Sueldo/30*Nom_Detalle." _
    & "Dias_Trab)+((Nom_Detalle.Sueldo/30)*Nom_Detalle.Dias_NoTrab*-1)+(Nom_Detalle." _
    & "Bono_Noc)) AS TotalA, Sum((Nom_Detalle.Sueldo+(Emp.BonoNoc*Nom_detalle.Sueldo/100)" _
    & ")/30*Nom_Detalle.Dias_libres) AS DiasL, Sum((Nom_Detalle.Sueldo+(Emp.BonoNoc*" _
    & "Nom_detalle.Sueldo/100))/30*Nom_Detalle.Dias_Feriados*'" & pDFT & "') AS diasF, Sum(Nom_Detalle" _
    & ".Otras_Asignaciones) AS OA, Inmueble.Ubica FROM Inmueble INNER JOIN (Nom_Detalle " _
    & "INNER JOIN Emp ON Nom_Detalle.CodEmp = Emp.CodEmp) ON Inmueble.CodInm = " _
    & "Emp.CodInm Where (((Nom_Detalle.IDNom) = " & IDNomina & ")) GROUP BY Emp.CodGasto, " _
    & "Emp.CodInm, Inmueble.Ubica ORDER BY Emp.CodInm, Emp.CodGasto;"
    
    rstNom.Open strSQL, cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
    
    If Not (rstNom.EOF And rstNom.BOF) Then
        rstNom.MoveFirst
        Do
            Call RtnProUtility("Eliminando información facturación inm. '" & rstNom("CodInm") & _
            "'", rstNom.AbsolutePosition * 6025 / rstNom.RecordCount)
            
            strSQL = "DELETE FROM AsignaGasto In '" & gcPath & rstNom("ubica") & "inm.mdb' WHERE " & _
            "NDoc='NOMINA' and CodGasto='" & rstNom("CodGasto") & "' and Monto = " & Replace(rstNom("TOTALA"), ",", ".")
            
            cnnConexion.Execute strSQL, N 'SUELDOS
            
            If rstNom("DiasL") > 0 Then
                strSQL = "DELETE FROM AsignaGasto In '" & gcPath & rstNom("ubica") & "inm.mdb' WHERE " & _
                "NDoc='NOMINA' and CodGasto='" & codNom(0) & "' and Monto = " & Replace(rstNom("diasL"), ",", ".")
                
                cnnConexion.Execute strSQL, N 'dias libres laborados
            End If
            
            If rstNom("DiasF") > 0 Then
                strSQL = "DELETE FROM AsignaGasto In '" & gcPath & rstNom("ubica") & "inm.mdb' WHERE " & _
                "NDoc='NOMINA' and CodGasto='" & codNom(1) & "' and Monto = " & Replace(rstNom("diasF"), ",", ".")
                
                cnnConexion.Execute strSQL, N 'dias feriados
            End If
            If rstNom("OA") > 0 Then
                strSQL = "DELETE FROM AsignaGasto In '" & gcPath & rstNom("ubica") & "inm.mdb' WHERE " & _
                "NDoc='NOMINA' and CodGasto='" & codNom(2) & "' and Monto = " & Replace(rstNom("OA"), ",", ".")
                
                cnnConexion.Execute strSQL, N 'otras asignaciones
            End If
            
            rstNom.MoveNext
        Loop Until rstNom.EOF
        rstNom.Close
        
    End If
    'eliminamos el reintegro del seguro social
    strSQL = "SELECT Emp.CodInm, Sum(Nom_Detalle.SSO + Nom_Detalle.spf) AS sumSSO, Inmueble.Ubi" _
    & "ca FROM Inmueble INNER JOIN (Nom_Detalle INNER JOIN Emp ON Nom_Detalle.CodEmp = Emp.CodEm" _
    & "p) ON Inmueble.CodInm = Emp.CodInm WHERE (((Nom_Detalle.IDNom) =" & IDNomina & ")) GROUP BY Em" _
    & "p.CodInm, Inmueble.Ubica ORDER BY Emp.CodInm;"
    
    With rstNom
        '
        .Open strSQL, cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
        
        If Not (.EOF And .BOF) Then
            .MoveFirst
            Do
                If !sumSSO <> 0 Then
                    Call RtnProUtility("Eliminando Reintegro Seguro Social inm. " & !CodInm, _
                    .AbsolutePosition * 6015 / .RecordCount)
                    
                    strSQL = "DELETE FROM AsignaGasto in '" & gcPath & !Ubica & "inm.mdb' " & _
                    "WHERE NDoc='NOMINA' and CodGasto='192020' and Cargado=#02/01/2008# and " & _
                    "Monto =" & Replace(!sumSSO * -1, ",", ".")
                    cnnConexion.Execute strSQL, N
                    
                End If
                .MoveNext
            Loop Until .EOF
        End If
        .Close
    End With
    'elimina ahora el reintegro de politica habitacional
    strSQL = "SELECT Emp.CodInm, Sum(Nom_Detalle.LPH) AS sumLPH, Inmueble.Ubica FROM Inmueble I" _
    & "NNER JOIN (Nom_Detalle INNER JOIN Emp ON Nom_Detalle.CodEmp = Emp.CodEmp) ON Inmueble.Co" _
    & "dInm = Emp.CodInm WHERE (((Nom_Detalle.IDNom) =" & IDNomina & ")) GROUP BY Emp.CodInm, Inmuebl" _
    & "e.Ubica ORDER BY Emp.CodInm;"
    With rstNom
        .Open strSQL, cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
        If Not (.EOF And .BOF) Then
            .MoveFirst
            '
            Do
                If !sumLPH <> 0 Then
                    Call RtnProUtility("Eliminando Reintegro Ley Política Hab. inm. " & !CodInm, _
                    .AbsolutePosition * 6015 / .RecordCount)
                    strSQL = "DELETE FROM AsignaGasto in '" & gcPath & !Ubica & "inm.mdb' " & _
                    "WHERE NDoc='NOMINA' and CodGasto='192026' and Cargado=#02/01/2008# and " & _
                    "Monto =" & Replace(!sumLPH * -1, ",", ".")
                    cnnConexion.Execute strSQL, N
                End If
                .MoveNext
            Loop Until .EOF
        End If
        .Close
    End With
    'genera las cpp de los cheques
    rstNom.Open "qdfNomina_Bco", cnnConexion, adOpenKeyset, adLockOptimistic, adCmdTable
    rstNom.Filter = "Cuenta='' or Cuenta = Null"
    Call RtnProUtility("Eliminando pagos en cheques...", 0)
    
    If Not rstNom.EOF And Not rstNom.BOF Then
        rstNom.MoveFirst
        Dim sFact As String
        Dim rstCargado As ADODB.Recordset
        Dim NDoc As String
        sFact = Left(IDNomina, 3) & Right(IDNomina, 2)
        Set rstCargado = New ADODB.Recordset
        
        Do
            'cppDetalle = Descripcion(rstNom!codGasto, rstNom!CodInm)
            strSQL = "SELECT Ndoc from Cpp WHERE Tipo='NO' and Left(Fact,5)='" & sFact & "' and Monto=" & Replace(rstNom("neto"), ",", ".") & " and CodInm='" & rstNom("CodInm") & "' and Benef='" & rstNom("name") & "'"
            
            NDoc = cnnConexion.Execute(strSQL)("Ndoc")
            
            'strSql = "DELETE FROM Cpp WHERE Tipo='NO' and Left(Fact,5)='" & sFact & _
            '"' and Monto=" & Replace(rstNom("neto"), ",", ".") & " and CodInm='" & _
            'rstNom("CodInm") & "' and Benef='" & rstNom("name") & "'"
            strSQL = "DELETE FROM Cargado in '" & gcPath & "\" & rstNom("CodInm") & "\inm.mdb' WHERE Ndoc='" & NDoc & "'"
            cnnConexion.Execute strSQL, N
            strSQL = "DELETE FROM Cpp WHERE Ndoc='" & NDoc & "'"
            
            cnnConexion.Execute strSQL, N
            
            rstNom.MoveNext
        Loop Until rstNom.EOF
    End If
    rstNom.Close
    Set rstNom = Nothing
    
    strSQL = "DELETE FROM Nom_Inf WHERE IDNomina=" & IDNomina
    cnnConexion.Execute strSQL, N
    
    If Err <> 0 Then
        cnnConexion.RollbackTrans
        reversar_nomina = MsgBox("Ocurrieron errores durante el proceso." & vbCrLf _
        & "Consulte al administrador del sistema.")
    Else
        cnnConexion.CommitTrans
    End If
    
    End Function

    Public Sub rtnAjustasCpp()
    Dim strSQL As String
    Dim rstlocal As ADODB.Recordset
    Dim strUbica As String
    Dim curMonto As Currency
    Dim nReg As Long
    
    On Error GoTo salir
    
    strSQL = "SELECT * FROM Cpp WHERE Fact LIKE 'F%' AND (Frecep >=" & _
            "#05/01/2008# AND Frecep<=#05/31/08#) AND Estatus = " & _
            "'ASIGNADO' ORDER BY Frecep,CodInm"
    
    Set rstlocal = cnnConexion.Execute(strSQL)
    
    If Not (rstlocal.EOF And rstlocal.BOF) Then
        nReg = rstlocal.RecordCount
        cnnConexion.BeginTrans
        Do
            strUbica = "'c:\sac\datos\" & rstlocal("CodInm") & "\inm.mdb'"
            
            strSQL = "UPDATE Cargado in " & strUbica & _
                    " SET Monto = Clng(Monto * 100)/100 " & _
                    "WHERE Ndoc='" & rstlocal("Ndoc") & "'"
            
            cnnConexion.Execute strSQL
            
            strSQL = "SELECT Clng(Monto * 9)/100 FROM Cargado in " & strUbica & " WHERE CodGasto IN ('900018','600003') AND Ndoc='" & rstlocal("Ndoc") & "'"
            curMonto = cnnConexion.Execute(strSQL)(0)
            
            strSQL = "UPDATE Cargado in " & strUbica & " SET Monto = '" & curMonto & "' WHERE CodGasto='900025'"
            cnnConexion.Execute strSQL
            
            strSQL = "SELECT sum(Monto) from Cargado in " & strUbica & _
                        "WHERE Ndoc='" & rstlocal("Ndoc") & "'"
            curMonto = cnnConexion.Execute(strSQL)(0)
            
            strSQL = "UPDATE Cpp SET Monto ='" & curMonto & "', Total='" & curMonto _
                & "' WHERE Ndoc='" & rstlocal("Ndoc") & "'"
        
            cnnConexion.Execute (strSQL)
            
        rstlocal.MoveNext
        Loop Until rstlocal.EOF
'        strSQL = "Update Cpp SET Emitida = 0 WHERE Fact LIKE 'F%' AND Estatus=" & _
'            "'PAGADO'"
'        cnnConexion.Execute strSQL
        
salir:
        If Err.Number = 0 Then
            cnnConexion.CommitTrans
            MsgBox "Se actualizaron " & nReg & " registros.", vbInformation, App.ProductName
            
        Else
            cnnConexion.RollbackTrans
            MsgBox "Ocurrieron errores, no se actualizaron registros.", vbCritical, App.ProductName
        End If
    
    End If
    
    End Sub
    
    Private Sub verificar_actualizador()
    
    Const sServerFTP = "administradorasac.com"
    Const sUser = "admras"
    Const sPass = "dmn+str"
    Const sDir = "httpdocs\sacupdate"
    
    Dim hFind As Long
    Dim dLastDateModifLocal As Date
    Dim dLastDateModifRemote As Date
    Dim fso As New FileSystemObject
    Dim Fil As file
    
    'cargamos el formulario principal
    
    Set mFTP = New cFtp
    mFTP.SetModeActive
    mFTP.SetTransferBinary
    If Not mFTP.OpenConnection(sServerFTP, sUser, sPass) Then
    'si da algun error la conexion damos un mensaje al usuario
    'salimos de la rutina
        Exit Sub
    End If
    'establecemos el directorio de trabajo
    mFTP.SetFTPDirectory sDir
    'verificamos la fecha de la última actualización del archivo
    
    'ahora descargamos los archivos
    Dim lTimer As Long
    Dim strRemote As String
    Dim strLocal As String
    
    strRemote = "sacUpdate.exe"
    strLocal = App.Path & "\sacUpdate.exe"
    dLastDateModifRemote = mFTP.GetFTPFileLastDateModif(strRemote)
    
    If Dir(strLocal, vbArchive) = "" Then
        'MsgBox "No se encuentra el archivo sac.exe", vbCritical, App.ProductName
        Exit Sub
    End If
    
    Set Fil = fso.GetFile(strLocal)
    dLastDateModifLocal = Fil.DateLastModified
    'comparamos cual archivo es el mas actualizado
    If dLastDateModifRemote > dLastDateModifLocal Then
        'si el archivo remoto es más actual lo descargamos
        lTimer = Timer
        If strLocal <> "" Then
           If Not mFTP.FTPDownloadFile(strLocal, strRemote) Then
                MsgBox mFTP.GetLastErrorMessage, vbCritical, App.ProductName
           End If
        End If
    
    End If
    'cerramos la conexión
    mFTP.CloseConnection

    End Sub

    Public Function ejecutar_procedure(nombre_procedimiento As String, _
    ParamArray Parametros() As Variant) As ADODB.Recordset
    
    
    Dim cmdCP As New ADODB.Command
    '
    cmdCP.ActiveConnection = cnnConexion
    cmdCP.CommandText = nombre_procedimiento
    cmdCP.CommandType = adCmdTable
    For I = 0 To UBound(Parametros)
        cmdCP.Parameters(I) = Parametros(I)
    Next
    
    Set ejecutar_procedure = cmdCP.Execute
    
    End Function
    
    Public Sub insertar_registro(nombre_procedimiento As String, _
    ParamArray Parametros() As Variant)
    
    Dim cmdCP As New ADODB.Command
    '
    cmdCP.ActiveConnection = cnnConexion
    cmdCP.CommandText = nombre_procedimiento
    cmdCP.CommandType = adCmdStoredProc
'    For I = 0 To UBound(Parametros)
'        cmdCP.Parameters(0) = Parametros(I)
'    Next
'
    cmdCP.Execute N, Parametros()
    Call rtnBitacora(Parametros(4) & "  " & Parametros(5) & "  " & Parametros(6) & "  " & Parametros(7) & "  Bs..." & Parametros(8) & "  " & IIf(N = 0, "Falló", "OK"))
    End Sub
    
    Public Function listar_valores(strSQL As String, IncluirVacio As Boolean) As String
    
    Dim rst As ADODB.Recordset
    Set rst = New ADODB.Recordset
    
    rst.Open strSQL, cnnConexion, adOpenKeyset, adLockReadOnly, adCmdText
    If Not (rst.BOF And rst.EOF) Then
        Do
            If (listar_valores <> "") Then listar_valores = listar_valores & ","
            listar_valores = listar_valores + rst(0)
            rst.MoveNext
        Loop Until rst.EOF
        If IncluirVacio Then listar_valores = "," & listar_valores
    End If
    Set rst = Nothing
    End Function
    
    
    Public Sub mantenimiento_celda(Flex As MSFlexGrid, Columna As Integer, _
    codigo_tecla As Integer, Shift As Integer)
    If (codigo_tecla = 46 And Flex.Col = Columna) Then Flex.Text = ""
    If Shift = 2 Then
    '
        If codigo_tecla = 67 Then
            Clipboard.Clear
            Clipboard.SetText Flex.Text
        ElseIf codigo_tecla = 86 Then
            Flex.Text = Clipboard.GetText
        End If
        '
    End If
    End Sub

    Public Sub Printer_ReporteNomina(NombreReporte As String, OrigenDatos() As String, _
    Formulas() As String, Params(), NombreArchivo As String, TiuloReporte As String)
    
    Dim crReporte As ctlReport
    Dim I As Integer
    
    Set crReporte = New ctlReport
    crReporte.Reporte = gcReport + NombreReporte
    For I = 0 To UBound(OrigenDatos)
        crReporte.OrigenDatos(I) = gcPath + "\" + OrigenDatos(I)
    Next
    For I = 0 To UBound(Formulas)
        crReporte.Formulas(I) = Formulas(I)
    Next
    For I = 0 To UBound(Params)
        crReporte.Parametros(I) = Params(I)
    Next
    'guarda una copia local
'    crReporte.Salida = crArchivoDisco
'    crReporte.ArchivoSalida = NombreArchivo
'    crReporte.Imprimir
'    'envia a la impresora
'    crReporte.Salida = crImpresora
'    crReporte.Imprimir
    crReporte.Salida = crPantalla
    crReporte.TituloVentana = TiuloReporte
    crReporte.Imprimir
    
    Set crReporte = Nothing
    
    End Sub
    
    Public Sub printer_recibo_pago(IDNomina As Long, Guardar_Copia As Boolean, _
    Resp As VbMsgBoxStyle, Salida As crSalida)
    
    Dim crReporte As ctlReport
    Dim strTitulo As String
    Dim file As String
    Dim errLocal As Long
    
    Set crReporte = New ctlReport
    If Left(IDNomina, 1) = 1 Then
        strTitulo = "1era Quin."
    ElseIf Left(IDNomina, 1) = 2 Then
        strTitulo = "2da Quin. "
    Else
        If Left(IDNomina, 1) = 3 Then strTitulo = "Aguinaldos Año: " & Right(IDNomina, 4)
    End If
    strTitulo = strTitulo & MonthName(Mid(IDNomina, 2, 2)) & "-" & Right(IDNomina, 4)
    If Left(IDNomina, 1) = 3 Then
        strTitulo = "Aguinaldos Año: " & Right(IDNomina, 4)
    End If
    file = gcPath & "\nomina\" & "RP" & IDNomina & ".rpt"
    'si la copia de los recibos de pago no exise la generamos
    If Dir(file) = "" Then
        crReporte.Reporte = gcReport + "nom_rec_pago.rpt"
        crReporte.OrigenDatos(0) = gcPath & "\sac.mdb"
        crReporte.OrigenDatos(1) = gcPath & "\sac.mdb"
        crReporte.OrigenDatos(2) = gcPath & "\sac.mdb"
        crReporte.Parametros(0) = IDNomina
        
        crReporte.Formulas(0) = "subtitulo='" & strTitulo & "'"
        If Guardar_Copia Then
            crReporte.Salida = crArchivoDisco
            crReporte.ArchivoSalida = file
            errLocal = crReporte.Imprimir
            If errLocal <> 0 Then
                MsgBox "No se guardó la copia de este reporte. " & vbCrLf & _
                Err.Description, vbCritical, strTitulo
            End If
        End If
    Else
        'utilizamos la copia del archivo
        crReporte.Reporte = file
    End If
    If Resp = vbYes Then
        crReporte.TituloVentana = strTitulo
        crReporte.Salida = Salida
        errLocal = crReporte.Imprimir
        If errLocal <> 0 Then
            MsgBox "OCurrió un error durante la impresión. " & vbCrLf & _
                Err.Description, vbCritical, "Error " & Err.Number
        End If
    End If
    Set crReporte = Nothing
    '
    End Sub

    Function enviar_email(para As String, de As String, _
                        asunto As String, isHTML As Boolean, Mensaje As String, _
                        Optional archivo_adjunto As String) As Boolean
    
    Dim email As CDO.Message
    Dim adjunto() As String
    
    Set email = New CDO.Message
    
    'configuramos el objeto
    With email.Configuration
        .Fields(cdoSMTPServer) = SMTP_SERVER
        .Fields(cdoSendUsingMethod) = 2
        .Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = SMTP_SERVER_PORT
        .Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = Abs(1)
        .Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 30
        If SERVER_AUTH Then
            .Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = USER_NAME
            .Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = Password
            .Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = SSL
        End If
        .Fields.Update
    End With
    'estructura del email
    email.To = para
    email.BCC = "ynfantes@gmail.com"
    email.from = de
    email.Subject = asunto
    If isHTML Then
        email.HTMLBody = Mensaje
    Else
        email.TextBody = Mensaje
    End If
    'aqui colocamos los archivos adjuntos
    If Not IsNull(archivo_adjunto) Or archivo_adjunto <> "" Then
        If Dir(archivo_adjunto, vbArchive) <> "" Then
            email.AddAttachment (archivo_adjunto)
        End If
    End If
    On Error Resume Next
    email.Send
    enviar_email = Err = 0
    Set email = Nothing
    
    End Function
